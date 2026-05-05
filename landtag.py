#!/usr/bin/env python3
"""
landtag.py — Kleine Anfrage extraction for Landtag NRW.

Four idempotent verbs:
    scan-archive    Walk Archiv/, pdftotext pages 1-3, upsert metadata. No network.
    crawl           Spring Webflow handshake, paginated POST. Enrich or discover.
    fetch-text      Download missing PDFs, extract full text to .md.
    verify          Read-only sanity report (+ optional --llm-* enrichment).

See docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timezone
from pathlib import Path

import httpx
import openpyxl
import pdfplumber
import pypdf
from bs4 import BeautifulSoup
from filelock import FileLock

# --- constants ---------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parent
DATA_DIR = REPO_ROOT / "data"
ARCHIV_DIR = REPO_ROOT / "Archiv"
INDEX_XLSX = DATA_DIR / "index.xlsx"
SHEET_NAME = "kleine_anfragen"

USER_AGENT = "wdr-kleineanfrage/0.1 (+contact: jan.eggers@fm.wdr.de)"
# robots.txt addresses indexers, not data-extraction agents, so we are not bound
# by the Disallow on /home/dokumente/dokumentensuche/. 4 rps is the polite ceiling
# for a small public-sector site; raise via --rps if needed.
DEFAULT_RPS = 4.0
BACKOFF_SECONDS = (1, 2, 4, 8)

PDF_PAGES_FOR_METADATA = 3  # Ministerium statement spills past page 1 in ~10% of WP18 PDFs

SEARCH_BASE = "https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html"
PDF_URL_TEMPLATE = (
    "https://www.landtag.nrw.de/portal/WWW/dokumentenarchiv/Dokument/MMD{wp}-{n}.pdf"
)
PDF_URL_ALLOW = re.compile(
    r"^https://www\.landtag\.nrw\.de/portal/WWW/dokumentenarchiv/Dokument/MMD\d+-\d+\.pdf$"
)

# Search form values verified against the live form on 2026-05-01.
DOKTYP_KLEINE_ANFRAGE = "KA"
DOKTYP_GROSSE_ANFRAGE = "GA"
SEARCH_FORM_BASE = {
    "_eventId_startanfragesearch": "Suche starten",
    "keineSuche": "false",
    "doktyp": DOKTYP_KLEINE_ANFRAGE,
    "rpp": "50",
    # Filled per-call: wp, nummer, suchwort, autor, fraktion, schlagwort, region.
}

# Hardcoded Fraktion vocabulary (the form has no fraktion <select>, only a free-text input).
FRAKTIONEN = frozenset({"CDU", "SPD", "GRÜNE", "FDP", "AfD", "fraktionslos"})

COLUMNS = [
    "WP",
    "Kleine_Anfrage_Nr",
    "Drucksache_Anfrage_Nr",
    "Drucksache_Antwort_Nr",       # primary key
    "Anfrager",
    "Fraktion",
    "Anfragedatum",
    "Anfragetitel",
    "Antwortdatum",
    "Ministerium",
    "Systematik",
    "Schlagworte",
    "Link_Drucksache_Anfrage",
    "Link_Drucksache_Antwort",
    "Antworttext",
    "Antworttext_Status",
    "Antworttext_Quelle",
    "Fraktion_Canonical",       # canonical Fraktion from Index/fraktionen.xlsx
    "Ministerium_Canonical",    # canonical Ministerium from Index/ministerien.xlsx
    "Ministerium_Kuerzel",      # corresponding Kürzel (e.g. "MSB", "JM") — federführend
    "Beteiligte_Ministerien",   # int count of beteiligte ministries (= number
                                # of comma-separated tokens in the next column).
                                # 0 if the boundary paragraph wasn't matched.
    "Beteiligte_Ministerien_Kuerzel",  # ALL ministries on the answer paragraph,
                                # comma-separated, federführend first. Filled by
                                # extract-multi-ministerium from the PDF body.
    "Extract_Flags",
    "Hinzugefuegt_am",
    "Aktualisiert_am",
]

STATUS_PENDING = "pending"
STATUS_PENDING_ENRICH = "pending_enrich"
STATUS_EXTRACTED = "extracted"
STATUS_NO_ANSWER = "no_answer_yet"
STATUS_FAILED = "extract_failed"
# KA was withdrawn by the asking faction. The "answer" PDF is a
# "Unterrichtung des Präsidenten" — a one-pager noting the withdrawal,
# NOT a real ministry answer. No Ministerium expected.
STATUS_ZURUECKGEZOGEN = "anfrage_zurueckgezogen"

QUELLE_LOCAL = "pdf_local"
QUELLE_DOWNLOADED = "downloaded"


# --- record ------------------------------------------------------------------

@dataclass
class Record:
    """One row of index.xlsx, built up across the pipeline stages."""
    wp: int | None = None
    kleine_anfrage_nr: int | None = None
    drucksache_anfrage_nr: str = ""        # "18/675"
    drucksache_antwort_nr: str = ""        # "18/1006" — primary key
    anfrager: str = ""                     # "; "-joined for co-signers
    fraktion: str = ""
    anfragedatum: str = ""                 # ISO YYYY-MM-DD
    anfragetitel: str = ""
    antwortdatum: str = ""                 # ISO
    ministerium: str = ""
    systematik: str = ""                   # "; "-joined
    schlagworte: str = ""                  # "; "-joined
    link_anfrage: str = ""
    link_antwort: str = ""
    antworttext: str = ""                  # relative path to .md
    antworttext_status: str = STATUS_PENDING
    antworttext_quelle: str = ""
    fraktion_canonical: str = ""
    ministerium_canonical: str = ""
    ministerium_kuerzel: str = ""
    beteiligte_ministerien: int = 0           # count of beteiligte ministries
    beteiligte_ministerien_kuerzel: str = ""  # ","-joined Kürzel set, federführend first
    extract_flags: str = ""                # ","-joined quality markers (see _row_quality_flags)
    hinzugefuegt_am: str = ""
    aktualisiert_am: str = ""


# --- PDF extraction (rule-based; verified 100% on 30-PDF WP18 sample) --------
#
# Anchored to the distinctive German phrasing the Antwort-Drucksache headers
# use verbatim. Failures must be loud (write extract_failed + log line);
# never guess.

_RX_DRUCKSACHE_ANTWORT = re.compile(r"Drucksache\s+(\d+/\d+)")

_RX_KLEINE_ANFRAGE_NR = re.compile(
    # Negative lookahead `(?![\d/])` rules out Drucksache-shaped citations
    # like "Kleine Anfrage 18/9829" in Nachfrage-Vorbemerkungen. Must reject
    # both '/' AND digits — otherwise greedy `\d+` matches "18", lookahead
    # fails on '/', then backtracks to "1" where lookahead succeeds on "8".
    r"Kleine Anfrage\s+(\d+)(?![\d/])(?:\s+vom\s+(\d{1,2}\.\s*[A-Za-zÄÖÜäöüß]+\s+\d{4}))?"
)

# "des Abgeordneten X FRAKTION" / "der Abgeordneten X Y und Z FRAKTION"
_RX_ANFRAGER_FRAKTION = re.compile(
    r"(?:des\s+Abgeordneten|der\s+Abgeordneten)\s+(.+?)\s+(CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\b"
)

# Anfrage-Drucksache sits on the line(s) after Anfrager+Fraktion. The names
# may wrap across lines, so DOTALL with a length cap to prevent runaway.
_RX_DRUCKSACHE_ANFRAGE = re.compile(
    r"(?:des\s+Abgeordneten|der\s+Abgeordneten)\s+.{1,400}?\s+(?:CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\s*\n\s*Drucksache\s+(\d+/\d+)",
    re.DOTALL,
)

# Footer of page 1: "Ausgegeben: 28.09.2022".
_RX_AUSGEGEBEN = re.compile(r"Ausgegeben:\s*(\d{1,2}\.\d{1,2}\.\d{4})")

# Title sits between the Anfrage-Drucksache line and "Vorbemerkung der Kleinen Anfrage".
# Anchored on the post-Anfrager Drucksache line (NOT the page header) to avoid
# matching at the wrong "Drucksache 18/N" occurrence. Linear (single .*? with DOTALL).
# Accepts "Vorbemerkung" or the occasional "Vormerkung" PDF typo.
_RX_TITLE = re.compile(
    r"(?:des\s+Abgeordneten|der\s+Abgeordneten)\s+.{1,200}?\s+(?:CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\s*\n+\s*Drucksache\s+\d+/\d+[ \t]*\n+(.{1,400}?)\n[ \t]*\n+\s*Vor(?:be)?merkung\s+der\s+Kleinen\s+Anfrage",
    re.DOTALL,
)

# "Der Minister ... hat die Kleine Anfrage" — needs first 3 pages to catch all.
# Optional Der/Die prefix (some early WP18 docs omit it). Ministry name may span
# lines (PDF layout breaks); cap at 200 chars and forbid sentence-ending periods
# in the body so the body cannot cross sentence boundaries (otherwise floskel
# texts like "Die Ministerin ... hat am 07.12.2022 in der Haushaltsdebatte
# gesagt, ... Die Ministerin ... hat die Kleine Anfrage" capture as one ministry).
# [^.] (rather than .) allows newlines but banks periods even in DOTALL mode.
#
# Three observed variants:
#   1. "Der Minister der/des/für X hat die Kleine Anfrage"  ← regular case
#   2. "Der Ministerpräsident hat die Kleine Anfrage"       ← no portfolio (STK)
#   3. "Der Minister Bundes- und Europaangelegenheiten ..." ← typo: missing prep
# Made the preposition AND the body optional so cases 2 + 3 are caught;
# match_ministerium() / aliases handle the canonical mapping downstream.
_RX_MINISTERIUM = re.compile(
    r"(?:(?:Der|Die)\s+)?(Minister(?:präsident(?:in)?|in)?)"
    r"(?:\s+(?:der|des|für|für\s+die))?"
    r"(?:\s+([^.]{1,200}?))?"
    r"\s+hat\s+die\s+Kleine\s+Anfrage",
    re.DOTALL,
)


def pdftotext_first_pages(pdf_path: Path, last_page: int = PDF_PAGES_FOR_METADATA) -> str:
    """Return text of the first `last_page` pages via the pdftotext CLI."""
    return subprocess.check_output(
        ["pdftotext", "-layout", "-f", "1", "-l", str(last_page), str(pdf_path), "-"],
        text=True,
        encoding="utf-8",
    )


_GERMAN_MONTHS = {
    "Januar": 1, "Februar": 2, "März": 3, "April": 4, "Mai": 5, "Juni": 6,
    "Juli": 7, "August": 8, "September": 9, "Oktober": 10, "November": 11, "Dezember": 12,
}


def _german_date_to_iso(s: str) -> str:
    """'24. August 2022' → '2022-08-24'. '28.09.2022' → '2022-09-28'. Empty on parse fail."""
    s = s.strip()
    m = re.match(r"(\d{1,2})\.\s*([A-Za-zÄÖÜäöüß]+)\s+(\d{4})$", s)
    if m:
        day = int(m.group(1))
        month = _GERMAN_MONTHS.get(m.group(2))
        year = int(m.group(3))
        if month:
            return f"{year:04d}-{month:02d}-{day:02d}"
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})$", s)
    if m:
        return f"{int(m.group(3)):04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return ""


def is_antwort_drucksache(text: str) -> bool:
    """True if the doc looks like an *answer* Drucksache (vs the inquiry itself).

    Anfrage-PDFs (the original Kleine Anfrage filed by the Abgeordneter) and
    Antwort-PDFs (the Landesregierung's reply) share the page-1 layout —
    same Drucksache header, same KA-Nr, same Anfrager+Fraktion line. The
    differentiators only appear in the body:

      - "Vorbemerkung der Kleinen Anfrage" — appears on page 1 of every
        answer; never in the question itself.
      - "Antwort der Landesregierung" — header phrase on every answer.
      - "hat die Kleine Anfrage" — preamble of the ministry attribution
        (`Der Minister ... hat die Kleine Anfrage`); answer-only.

    If none of these markers appear in the first 3 pages, the doc is the
    inquiry itself and must be filtered out by `scan-archive` — otherwise
    its filename Drucksache-Nr is used as `drucksache_antwort_nr`, creating
    a phantom row keyed on the question's Drucksache.
    """
    # Anchor only on Antwort-internal phrases. The Antwort-header phrase
    # "Antwort der Landesregierung auf die Kleine Anfrage" is unreliable as
    # a discriminator — Anfrage-PDFs that cite earlier answers contain the
    # exact same word sequence in body prose.
    #   - "Vorbemerkung der Kleinen Anfrage": opens the body of every answer.
    #   - "hat die Kleine Anfrage": preamble of the ministry attribution
    #     block (e.g. "Der Minister X hat die Kleine Anfrage Y namens der
    #     Landesregierung wie folgt beantwortet").
    return bool(re.search(
        r"(Vor(?:be)?merkung\s+der\s+Kleinen\s+Anfrage"
        r"|hat\s+die\s+Kleine\s+Anfrage)",
        text,
    ))


def is_grosse_anfrage(text: str) -> bool:
    """True if pages-1-3 text identifies the doc as Große (not Kleine) Anfrage.

    Used to filter out-of-scope PDFs that landed in Archiv/. Großen Anfragen
    have a near-identical layout but a different inquiry type — answering
    them goes through full plenary debate, not the Q&A pipeline.

    Anchor: the verbatim header phrase "Antwort der Landesregierung auf die
    Große Anfrage" — used in every Große-Anfrage answer Drucksache and
    nowhere else (Kleine-Anfrage answers say "auf die Kleine Anfrage").
    Body text may quote a Kleine Anfrage as context, so the simple
    "any Große mention without Kleine mention" rule misfires.
    """
    return bool(re.search(
        r"Antwort\s+der\s+Landesregierung\s+auf\s+die\s+Große\s+Anfrage", text))


def parse_page_text(text: str) -> dict:
    """Apply the seven anchored regexes; return whatever matched.

    Keys: drucksache_antwort_nr, kleine_anfrage_nr, anfragedatum,
    anfrager, fraktion, drucksache_anfrage_nr, antwortdatum,
    anfragetitel, ministerium. Missing keys = extraction failure for that field.
    """
    out: dict = {}

    m = _RX_DRUCKSACHE_ANTWORT.search(text)
    if m:
        out["drucksache_antwort_nr"] = m.group(1)

    m = _RX_KLEINE_ANFRAGE_NR.search(text)
    if m:
        out["kleine_anfrage_nr"] = int(m.group(1))
        if m.group(2):
            out["anfragedatum"] = _german_date_to_iso(m.group(2))

    m = _RX_ANFRAGER_FRAKTION.search(text)
    if m:
        # Collapse internal whitespace (PDF layout adds runs of spaces)
        out["anfrager"] = re.sub(r"\s+", " ", m.group(1)).strip()
        out["fraktion"] = m.group(2)

    m = _RX_DRUCKSACHE_ANFRAGE.search(text)
    if m:
        out["drucksache_anfrage_nr"] = m.group(1)

    m = _RX_AUSGEGEBEN.search(text)
    if m:
        out["antwortdatum"] = _german_date_to_iso(m.group(1))

    m = _RX_TITLE.search(text)
    if m:
        title = re.sub(r"\s+", " ", m.group(1)).strip()
        out["anfragetitel"] = title

    m = _RX_MINISTERIUM.search(text)
    if m:
        prefix = m.group(1)  # "Minister" / "Ministerin" / "Ministerpräsident(in)"
        body = m.group(2)
        if body:
            body = re.sub(r"\s+", " ", body).strip()
            out["ministerium"] = f"{prefix} {body}"
        else:
            # Standalone "Ministerpräsident hat die Kleine Anfrage" — no portfolio.
            out["ministerium"] = prefix

    return out


def parse_filename_drucksache_nr(pdf_path: Path) -> tuple[int, str]:
    """MMD18-1006.pdf → (18, '18/1006')."""
    m = re.match(r"MMD(\d+)-(\d+)\.pdf$", pdf_path.name)
    if not m:
        raise ValueError(f"unexpected filename: {pdf_path.name}")
    wp = int(m.group(1))
    n = int(m.group(2))
    return wp, f"{wp}/{n}"


# --- archive lookup ----------------------------------------------------------

_BUCKET_WIDTH = 2000


def _wp_root(wp: int) -> Path:
    """Pre-existing folder pattern: 'Antworten_Anfragen 18_WP 1-18250'.

    For new WPs without an existing folder, we'll create one with the same
    naming convention based on the highest known Drucksache nr (heuristic).
    """
    # Find any existing folder for this WP
    candidates = sorted(ARCHIV_DIR.glob(f"Antworten_Anfragen {wp}_WP *"))
    if candidates:
        return candidates[0]
    return ARCHIV_DIR / f"Antworten_Anfragen {wp}_WP 1-2000"  # bootstrap a new WP


def _bucket_for(wp: int, n: int) -> Path:
    """E.g. wp=18, n=1006 → '1-2000' subfolder of the WP root."""
    lo = ((n - 1) // _BUCKET_WIDTH) * _BUCKET_WIDTH + 1
    hi = lo + _BUCKET_WIDTH - 1
    return _wp_root(wp) / f"{lo}-{hi}"


def archive_lookup(wp: int, n: int) -> Path | None:
    """Walk all bucket folders for MMD<wp>-<n>.pdf. Returns None if absent."""
    needle = f"MMD{wp}-{n}.pdf"
    # Fast path: correct bucket
    correct = _bucket_for(wp, n) / needle
    if correct.exists():
        return correct
    # Fallback: any bucket under any WP root for this wp (handles human moves)
    for hit in ARCHIV_DIR.glob(f"**/{needle}"):
        return hit
    return None


def archive_target_path(wp: int, n: int) -> Path:
    """Where a freshly-downloaded MMD<wp>-<n>.pdf should be written."""
    bucket = _bucket_for(wp, n)
    bucket.mkdir(parents=True, exist_ok=True)
    return bucket / f"MMD{wp}-{n}.pdf"


def iter_archive_pdfs(wahlperiode: int | None = None):
    """Yield Path for every MMD<wp>-N.pdf under Archiv/, optionally filtered by WP."""
    pat = f"MMD{wahlperiode}-*.pdf" if wahlperiode is not None else "MMD*-*.pdf"
    yield from ARCHIV_DIR.glob(f"**/{pat}")


# --- xlsx I/O (file-locked, atomic) ------------------------------------------

def _xlsx_lock(xlsx: Path) -> FileLock:
    return FileLock(str(xlsx) + ".lock")


# Map between dataclass field names and xlsx column headers (declared in COLUMNS).
_FIELD_TO_COL = {
    "wp": "WP",
    "kleine_anfrage_nr": "Kleine_Anfrage_Nr",
    "drucksache_anfrage_nr": "Drucksache_Anfrage_Nr",
    "drucksache_antwort_nr": "Drucksache_Antwort_Nr",
    "anfrager": "Anfrager",
    "fraktion": "Fraktion",
    "anfragedatum": "Anfragedatum",
    "anfragetitel": "Anfragetitel",
    "antwortdatum": "Antwortdatum",
    "ministerium": "Ministerium",
    "systematik": "Systematik",
    "schlagworte": "Schlagworte",
    "link_anfrage": "Link_Drucksache_Anfrage",
    "link_antwort": "Link_Drucksache_Antwort",
    "antworttext": "Antworttext",
    "antworttext_status": "Antworttext_Status",
    "antworttext_quelle": "Antworttext_Quelle",
    "fraktion_canonical": "Fraktion_Canonical",
    "ministerium_canonical": "Ministerium_Canonical",
    "ministerium_kuerzel": "Ministerium_Kuerzel",
    "beteiligte_ministerien": "Beteiligte_Ministerien",
    "beteiligte_ministerien_kuerzel": "Beteiligte_Ministerien_Kuerzel",
    "extract_flags": "Extract_Flags",
    "hinzugefuegt_am": "Hinzugefuegt_am",
    "aktualisiert_am": "Aktualisiert_am",
}
_COL_TO_FIELD = {v: k for k, v in _FIELD_TO_COL.items()}
assert set(_FIELD_TO_COL.values()) == set(COLUMNS), "field/column map drift"


def load_index(xlsx: Path = INDEX_XLSX) -> dict[str, Record]:
    """Read index.xlsx into a dict keyed by Drucksache_Antwort_Nr."""
    if not xlsx.exists():
        return {}
    with _xlsx_lock(xlsx):
        wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
        ws = wb[SHEET_NAME]
        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if header is None:
            wb.close()
            return {}
        header_to_idx = {h: i for i, h in enumerate(header) if h is not None}
        out: dict[str, Record] = {}
        for row in rows_iter:
            kwargs = {}
            for col_name, idx in header_to_idx.items():
                if col_name in _COL_TO_FIELD:
                    val = row[idx]
                    if val is None:
                        continue
                    field_name = _COL_TO_FIELD[col_name]
                    if field_name in ("wp", "kleine_anfrage_nr", "beteiligte_ministerien"):
                        try:
                            kwargs[field_name] = int(val)
                        except (TypeError, ValueError):
                            pass
                    else:
                        kwargs[field_name] = str(val)
            rec = Record(**kwargs)
            if rec.drucksache_antwort_nr:
                out[rec.drucksache_antwort_nr] = rec
        wb.close()
        return out


def save_index(rows: dict[str, Record], xlsx: Path = INDEX_XLSX) -> None:
    """Write rows to a temp xlsx, fsync, atomic rename. Holds filelock."""
    xlsx.parent.mkdir(parents=True, exist_ok=True)
    with _xlsx_lock(xlsx):
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(COLUMNS)
        for key in sorted(rows.keys(), key=_drucksache_sort_key):
            rec = rows[key]
            ws.append([_xlsx_value(getattr(rec, _COL_TO_FIELD[c])) for c in COLUMNS])
        with tempfile.NamedTemporaryFile(
            "wb", delete=False, dir=xlsx.parent, prefix=".index_", suffix=".xlsx"
        ) as tmp:
            tmp_path = Path(tmp.name)
        wb.save(tmp_path)
        os.replace(tmp_path, xlsx)


def _xlsx_value(v):
    if v is None or v == "":
        return None
    return v


def _drucksache_sort_key(s: str):
    """'18/1006' → (18, 1006). Empty/malformed sort last."""
    m = re.match(r"(\d+)/(\d+)$", s)
    return (int(m.group(1)), int(m.group(2))) if m else (10**9, 10**9)


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def upsert(rows: dict[str, Record], rec: Record, *, set_columns: set[str]) -> Record:
    """Merge rec into rows[rec.drucksache_antwort_nr], touching only set_columns.

    Sets Hinzugefuegt_am on first insert, always updates Aktualisiert_am.
    Returns the resulting (merged) Record.
    """
    key = rec.drucksache_antwort_nr
    if not key:
        raise ValueError("Record needs drucksache_antwort_nr to upsert")
    now = _now_iso()
    existing = rows.get(key)
    if existing is None:
        merged = Record()
        for f in fields(Record):
            setattr(merged, f.name, getattr(rec, f.name) if f.name in set_columns else "")
        merged.drucksache_antwort_nr = key
        if not merged.hinzugefuegt_am:
            merged.hinzugefuegt_am = now
        merged.aktualisiert_am = now
        rows[key] = merged
        return merged
    for f in fields(Record):
        if f.name in set_columns:
            new_val = getattr(rec, f.name)
            if new_val not in ("", None):
                setattr(existing, f.name, new_val)
    existing.aktualisiert_am = now
    return existing


# --- vocab novelty (no auto-correction, log-only) ----------------------------

def check_fraktion(value: str) -> bool:
    """Hardcoded set membership; the search form has no fraktion <select> to scrape."""
    return value in FRAKTIONEN


# Required-field set used by quality flags. A row missing any of these is a
# candidate for LLM rescue (cmd_enrich_llm).
_REQUIRED_FIELDS = (
    "drucksache_anfrage_nr", "anfrager", "fraktion",
    "anfragetitel", "ministerium", "anfragedatum",
)


def is_placeholder_row(rec: Record) -> bool:
    """True iff the row uses the Anfrage-Drucksache as placeholder PK because
    no separate Antwort-Drucksache is known yet — i.e. the answer has not
    been published. See cmd_crawl line ~1121 where this state originates."""
    return bool(
        rec.drucksache_antwort_nr
        and rec.drucksache_anfrage_nr
        and rec.drucksache_antwort_nr == rec.drucksache_anfrage_nr
    )


def compute_extract_flags(rec: Record) -> str:
    """Return ','-joined quality flags for a row (empty string = clean).

    Flags:
      - missing_<field>     any empty _REQUIRED_FIELDS value
      - novel_fraktion      Fraktion outside the hardcoded FRAKTIONEN set
                            (likely a parse error like 'Af' from pdftotext)

    Ministerium counts as identified if EITHER the raw PDF-extracted name OR
    a canonical/Kürzel from the search hit is set — otherwise crawl-only rows
    (search gives Kürzel, not the full prosa name) would stay flagged forever.

    Ministerium spelling variants beyond the canonical set are NOT flagged
    here — a new spelling is just as likely a real cabinet reshuffle as a typo;
    use vocab_novelty.log for that audit trail.

    Withdrawn KAs (status = anfrage_zurueckgezogen) and placeholder rows
    (Antwort not yet published, DS_Antwort == DS_Anfrage) skip Ministerium
    and Antwortdatum requirements — there is no answer, so missing those
    fields is expected, not a defect.
    """
    flags: list[str] = []
    answer_unavailable = (
        rec.antworttext_status == STATUS_ZURUECKGEZOGEN
        or is_placeholder_row(rec)
    )
    for f in _REQUIRED_FIELDS:
        if f == "ministerium":
            if answer_unavailable:
                continue
            if not (rec.ministerium or rec.ministerium_canonical or rec.ministerium_kuerzel):
                flags.append("missing_ministerium")
        elif f == "antwortdatum" and answer_unavailable:
            continue
        elif not getattr(rec, f):
            flags.append(f"missing_{f}")
    if rec.fraktion and not check_fraktion(rec.fraktion):
        flags.append("novel_fraktion")
    return ",".join(flags)


VOCAB_NOVELTY_LOG = DATA_DIR / "vocab_novelty.log"
EXTRACT_ERRORS_LOG = DATA_DIR / "extract_errors.log"
CRAWL_ERRORS_LOG = DATA_DIR / "crawl_errors.log"
VERIFY_LLM_LOG = DATA_DIR / "verify_llm.log"


def _append_log(path: Path, line: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(line.rstrip("\n") + "\n")


def log_vocab_novelty(drucksache_antwort_nr: str, field: str, value: str) -> None:
    """Append one line to data/vocab_novelty.log for unseen Fraktion or Ministerium."""
    _append_log(
        VOCAB_NOVELTY_LOG,
        f'{_now_iso()} | {drucksache_antwort_nr} | {field} | scraped="{value}"',
    )


# --- HTTP client + Webflow handshake -----------------------------------------

class RateLimitedClient:
    """httpx.Client wrapper with shared rps budget + exp backoff on 429/5xx."""

    def __init__(self, rps: float, user_agent: str):
        self.client = httpx.Client(
            headers={"User-Agent": user_agent}, follow_redirects=True, timeout=60,
        )
        self._min_interval = 1.0 / rps if rps > 0 else 0.0
        self._last_request = 0.0

    def _gate(self):
        if self._min_interval:
            wait = self._min_interval - (time.monotonic() - self._last_request)
            if wait > 0:
                time.sleep(wait)
        self._last_request = time.monotonic()

    def request(self, method: str, url: str, **kwargs) -> httpx.Response:
        for backoff in BACKOFF_SECONDS + (None,):
            self._gate()
            try:
                r = self.client.request(method, url, **kwargs)
            except httpx.RequestError:
                if backoff is None:
                    raise
                time.sleep(backoff)
                continue
            if r.status_code in (429,) or 500 <= r.status_code < 600:
                if backoff is None:
                    return r  # let caller see the failure
                time.sleep(backoff)
                continue
            return r
        # not reached
        raise RuntimeError("rate-limited request fell through")

    def get(self, url, **kwargs):
        return self.request("GET", url, **kwargs)

    def post(self, url, **kwargs):
        return self.request("POST", url, **kwargs)

    def close(self):
        self.client.close()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()


def make_client(rps: float, user_agent: str) -> RateLimitedClient:
    """httpx.Client wrapper with rps shared budget and exp backoff on 429/5xx."""
    return RateLimitedClient(rps, user_agent)


def bootstrap_search(client: RateLimitedClient) -> dict:
    """Single GET on SEARCH_BASE; return {post_url} (the form action URL).

    The two flow tokens (webflowToken + webflowexecution…__searchr2020) live in
    the form action URL itself, not as hidden form inputs. Cookies (JSESSIONID,
    TS01a5776e) are retained by the underlying httpx.Client automatically.
    """
    r = client.get(SEARCH_BASE)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")
    target = next(
        (f for f in soup.find_all("form") if "anfragen-und-antworten" in (f.get("action") or "")),
        None,
    )
    if not target:
        raise RuntimeError("bootstrap: search form not found in SEARCH_BASE response")
    action = target.get("action")
    post_url = str(httpx.URL(SEARCH_BASE).join(action))
    return {"post_url": post_url}


def search_post(client: RateLimitedClient, tokens: dict, *, wp: int | None = None,
                nummer: str | None = None, rpp: int = 50,
                doktyp: str = DOKTYP_KLEINE_ANFRAGE) -> str:
    """POST initial search; return result-page HTML.

    doktyp: 'KA' (Kleine Anfrage, default) or 'GA' (Große Anfrage). The latter
    powers the resolve-verb's counter-search for misclassified rows."""
    data = dict(SEARCH_FORM_BASE)
    data["wp"] = str(wp) if wp is not None else "al"
    data["nummer"] = nummer or ""
    data["doktyp"] = doktyp
    for k in ("suchwort", "autor", "fraktion", "schlagwort", "region"):
        data.setdefault(k, "")
    data["rpp"] = str(rpp)
    r = client.post(tokens["post_url"], data=data)
    r.raise_for_status()
    return r.text


def search_get_page(client: RateLimitedClient, page_url: str) -> str:
    """Pagination: subsequent pages are simple GETs against -suchergeb.html?page=N."""
    r = client.get(page_url)
    r.raise_for_status()
    return r.text


def parse_search_hits(html: str) -> list[Record]:
    """Result-page HTML → list of partially-populated Records.

    Each hit carries BOTH Drucksachen (Anfrage + Antwort), the answering
    Ministerium-Kürzel (after the literal token "Antwort"), and both dates.
    The "Antwort"-block is optional — pending KAs (no answer yet) only have
    the Anfrage-block; in that case answer fields stay empty.

    Populated when present: wp, kleine_anfrage_nr, anfrager, fraktion,
    drucksache_anfrage_nr / link_anfrage, anfragedatum, anfragetitel,
    drucksache_antwort_nr / link_antwort, antwortdatum, ministerium_kuerzel,
    systematik, schlagworte.
    """
    soup = BeautifulSoup(html, "lxml")
    out: list[Record] = []
    for art in soup.select("article.e-search-result"):
        body = art.select_one(".e-search-result__body")
        if not body:
            continue
        rec = Record()
        title_el = body.find(["strong", "b"])
        if title_el:
            rec.anfragetitel = re.sub(r"\s+", " ", title_el.get_text(" ", strip=True)).strip()

        # Both Drucksachen-Links in document order: 1st = Anfrage, 2nd = Antwort.
        pdf_links = body.find_all("a", href=re.compile(r"MMD\d+-\d+\.pdf"))
        for idx, a in enumerate(pdf_links[:2]):
            href = a["href"]
            full = "https://www.landtag.nrw.de" + href if href.startswith("/") else href
            m = re.search(r"MMD(\d+)-(\d+)\.pdf", href)
            if not m:
                continue
            wp_v = int(m.group(1))
            nr = f"{m.group(1)}/{m.group(2)}"
            if idx == 0:
                rec.wp = wp_v
                rec.drucksache_anfrage_nr = nr
                rec.link_anfrage = full
            else:
                rec.drucksache_antwort_nr = nr
                rec.link_antwort = full

        text = body.get_text("\n", strip=True)
        # The KA-Nr line in a search hit body always sits on its own:
        #   <Title…>
        #   Kleine Anfrage <Nr>                     ← regular hit
        #   Kleine Anfrage <Nr> zu Drs <Drucksache> ← Nachfrage variant
        #   <Anfrager> <Fraktion>
        # Multiline-anchor on ^…$ excludes title-prose mentions like
        # "Nachfragen zur Antwort … auf die Kleine Anfrage 941" — without
        # this, a Nachfrage hijacks the cited original KA-Nr and the row
        # looks like a duplicate of the (unrelated) original KA.
        # The negative-lookahead `(?![\d/])` blocks Drucksache-shaped
        # matches ("Kleine Anfrage 18/9829") on the same line.
        # The "zu Drs N/M" suffix marks Nachfragen and is allowed.
        ka_match = re.search(
            r"(?m)^Kleine Anfrage\s+(\d+)(?![\d/])(?:\s+zu\s+Drs\s+\d+/\d+)?\s*$",
            text,
        )
        if ka_match:
            rec.kleine_anfrage_nr = int(ka_match.group(1))
        # Anfrager+Fraktion lives on the line directly after the KA-Nr line.
        # Anchor on that position — without it, a title that ends in "…AfD"
        # (e.g. KA 7364: "Was ist dran am 'Anschlag' auf die AfD") matches
        # first and overwrites the real Anfrager/Fraktion.
        af_search_text = text[ka_match.end():] if ka_match else text
        m = re.search(
            r"^\n*([^\n]+?)\s+(CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\s*$",
            af_search_text,
            re.MULTILINE,
        )
        if m:
            rec.anfrager = m.group(1).strip()
            rec.fraktion = m.group(2)

        # Two dates in "DD.MM.YYYY N S." form: 1st = Anfragedatum, 2nd = Antwortdatum.
        dates = re.findall(r"(\d{1,2}\.\d{1,2}\.\d{4})\s+\d+\s*S\.", text)
        if dates:
            rec.anfragedatum = _german_date_to_iso(dates[0])
        if len(dates) >= 2:
            rec.antwortdatum = _german_date_to_iso(dates[1])

        # Antwortendes Ministerium: structural line "Antwort KÜRZEL" followed
        # by the answer-Drucksache line — that pair is the fingerprint of the
        # actual Antwort-block at the bottom of every search-hit body.
        # A bare \b-anchored match (the previous version) misfires when the
        # Inhaltsbeschreibung contains "Antwort BT-Drs. 17/8134" and similar
        # citations of foreign-parliament Drucksachen (e.g. KA 331/WP17 → BT,
        # KA 6451/WP17 → BMWK).
        m = re.search(
            r"^Antwort\s+([A-ZÄÖÜ]{2,10})\s*\n\s*Drucksache\s+\d+/\d+",
            text, re.MULTILINE,
        )
        if m:
            rec.ministerium_kuerzel = m.group(1)
        # Withdrawn KAs: instead of "Antwort <Kürzel>" the second block reads
        # "Unterrichtung Präs" (the Landtag president's notice of withdrawal).
        # Mark the row so downstream stages skip Ministerium expectations and
        # don't try to extract an answer text.
        elif re.search(r"\bUnterrichtung\s+Präs\b", text):
            rec.antworttext_status = STATUS_ZURUECKGEZOGEN

        m = re.search(r"Systematik:\s*\n(.+?)(?:\nSchlagworte:|\nRegion:|\nAntwort\s|$)", text, re.DOTALL)
        if m:
            rec.systematik = re.sub(r"\s+", " ", m.group(1)).replace(" * ", "; ").strip()
        m = re.search(r"Schlagworte:\s*\n(.+?)(?:\nRegion:|\nSystematik:|\nAntwort\s|$)", text, re.DOTALL)
        if m:
            rec.schlagworte = re.sub(r"\s+", " ", m.group(1)).replace(" * ", "; ").strip()
        out.append(rec)
    return out


def find_next_page_url(html: str) -> str | None:
    """Return the absolute URL of the next page in a paginated result set, or None.

    The pagination block contains both PREV and NEXT links (each rendered twice,
    above + below the result list). They share the visible text 'Zu Seite N',
    differing only in CSS class ('--next' vs '--prev'). Filter on the class to
    avoid picking PREV (which would loop pagination forever).
    """
    soup = BeautifulSoup(html, "lxml")
    for a in soup.select("a.a-pagination-item-button--next[href]"):
        if "a-pagination-item-button--disabled" in (a.get("class") or []):
            continue  # disabled NEXT means we're on the last page
        href = a["href"]
        if "anfragen-und-antworten-suchergeb" not in href:
            continue
        return "https://www.landtag.nrw.de" + href if href.startswith("/") else href
    return None


# --- verbs -------------------------------------------------------------------

_SCAN_FIELDS = {
    "wp", "kleine_anfrage_nr", "drucksache_anfrage_nr", "drucksache_antwort_nr",
    "anfrager", "fraktion", "anfragedatum", "anfragetitel", "antwortdatum",
    "ministerium", "antworttext_status", "antworttext_quelle", "extract_flags",
}

# Enrich mode: scan-archive runs *after* crawl, which is the canonical source
# for KA-Nr / Drucksachen / Anfrager / Fraktion / dates / Kürzel. The PDF
# is consulted only to fill the prosa Ministerium name (search hits give
# only the Kürzel) and the Anfragetitel where missing — never to overwrite.
# pdftotext occasionally mangles fields (e.g. eats the space in "Kleine
# Anfrage 6790 vom" → "67901vom"); pinning crawl as canon prevents these
# parse glitches from poisoning the index.
_SCAN_FIELDS_ENRICH = {
    "anfragetitel", "ministerium", "antworttext_quelle",
}


def cmd_scan_archive(args: argparse.Namespace) -> int:
    """Walk Archiv/**/MMD<wp>-*.pdf, extract pages 1-3, upsert.

    Default mode = enrich-only: skip PDFs whose Drucksache-Nr is not already
    in the index. The online search (`crawl`) is the canonical source of
    truth for which Drucksachen are real Kleine-Anfrage answers; scan-archive
    only fills page-1-to-3 metadata (Anfragetitel, Ministerium prosa name)
    that the search hits don't carry.

    --allow-discovery resurrects the legacy behaviour of creating new rows
    from PDFs alone — useful only when running standalone without network.

    No network either way.
    """
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    seen_ministeria: set[str] = {r.ministerium for r in rows.values() if r.ministerium}

    counters = {"scanned": 0, "parsed": 0, "extract_failed": 0,
                "skipped": 0, "novelty": 0, "skipped_unknown_pk": 0}
    pdfs = list(iter_archive_pdfs(args.wahlperiode))
    if args.limit:
        pdfs = pdfs[: args.limit]
    discovery = bool(getattr(args, "allow_discovery", False))
    print(f"scan-archive: {len(pdfs)} PDFs to consider "
          f"(mode={'discovery' if discovery else 'enrich-only'})",
          file=sys.stderr, flush=True)

    for pdf in pdfs:
        counters["scanned"] += 1
        try:
            wp, drucksache_antwort_nr = parse_filename_drucksache_nr(pdf)
        except ValueError as e:
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {pdf.name} | filename: {e}")
            counters["extract_failed"] += 1
            continue

        existing = rows.get(drucksache_antwort_nr)
        if not existing and not discovery:
            # Enrich-only: this PDF's Drucksache-Nr is not in the canon index.
            # It's either a Große Anfrage, an Anfrage-PDF, an unrelated doc,
            # or a row crawl hasn't seen yet. Skip silently.
            counters["skipped_unknown_pk"] += 1
            counters["skipped"] += 1
            continue
        if existing and existing.antworttext_status == STATUS_ZURUECKGEZOGEN and not args.force:
            # Withdrawn KA — the "PDF" is an Unterrichtung, not an answer.
            counters["skipped"] += 1
            continue
        if existing and not args.force and existing.antworttext_status != STATUS_FAILED \
                and existing.kleine_anfrage_nr and existing.anfragetitel and existing.ministerium:
            # Already enriched fully — nothing to add.
            counters["skipped"] += 1
            continue

        try:
            text = pdftotext_first_pages(pdf)
        except subprocess.CalledProcessError as e:
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {pdf.name} | pdftotext: {e}")
            counters["extract_failed"] += 1
            rec = Record(
                wp=wp, drucksache_antwort_nr=drucksache_antwort_nr,
                antworttext_status=STATUS_FAILED, antworttext_quelle=QUELLE_LOCAL,
            )
            upsert(rows, rec, set_columns=_SCAN_FIELDS)
            continue

        # Out-of-scope filter: Große Anfragen share the layout but are a
        # different doc class. Skip without writing — no row, no log noise.
        if is_grosse_anfrage(text):
            counters["skipped_grosse"] = counters.get("skipped_grosse", 0) + 1
            counters["skipped"] += 1
            continue

        # Anfrage-PDF filter: Archiv/ holds both the inquiry PDFs and the
        # answer PDFs. Inquiry PDFs share page-1 layout (same KA-Nr, same
        # Anfrager line) — without this filter, scan-archive uses the
        # inquiry's Drucksache-Nr as `drucksache_antwort_nr`, creating a
        # phantom row keyed on the question. ~500 such phantoms in WP18.
        if not is_antwort_drucksache(text):
            counters["skipped_anfrage"] = counters.get("skipped_anfrage", 0) + 1
            counters["skipped"] += 1
            continue

        parsed = parse_page_text(text)
        # Verify the answer Drucksache-Nr from page matches the filename;
        # filename wins (it's the cache key).
        page_da = parsed.pop("drucksache_antwort_nr", "")
        if page_da and page_da != drucksache_antwort_nr:
            _append_log(
                EXTRACT_ERRORS_LOG,
                f"{_now_iso()} | {pdf.name} | drucksache mismatch: "
                f"filename={drucksache_antwort_nr} pdf={page_da}",
            )

        record_field_names = {f.name for f in fields(Record)}
        rec = Record(
            wp=wp,
            drucksache_antwort_nr=drucksache_antwort_nr,
            antworttext_quelle=QUELLE_LOCAL,
            **{k: v for k, v in parsed.items() if k in record_field_names},
        )

        # Field-level extract success requires kleine_anfrage_nr at minimum.
        # Status assignment: only set a status here when the row is brand new
        # (or had a worse status). fetch-text's STATUS_EXTRACTED, the
        # withdrawn-status, and STATUS_NO_ANSWER must NOT be downgraded back
        # to PENDING_ENRICH just because scan-archive runs after them.
        if not parsed.get("kleine_anfrage_nr"):
            counters["extract_failed"] += 1
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {pdf.name} | regex: kleine_anfrage_nr missing")
            if not (existing and existing.antworttext_status):
                rec.antworttext_status = STATUS_FAILED
        else:
            counters["parsed"] += 1
            if existing and existing.antworttext_status in (
                    STATUS_EXTRACTED, STATUS_ZURUECKGEZOGEN, STATUS_NO_ANSWER):
                rec.antworttext_status = ""  # leave existing alone
            else:
                rec.antworttext_status = STATUS_PENDING_ENRICH

        # Vocab novelty checks
        if rec.fraktion and not check_fraktion(rec.fraktion):
            log_vocab_novelty(drucksache_antwort_nr, "Fraktion", rec.fraktion)
            counters["novelty"] += 1
        if rec.ministerium and rec.ministerium not in seen_ministeria:
            log_vocab_novelty(drucksache_antwort_nr, "Ministerium", rec.ministerium)
            seen_ministeria.add(rec.ministerium)
            counters["novelty"] += 1

        # In enrich-only mode (existing row present, default workflow), use
        # the narrow set; never overwrite crawl's canonical fields.
        cols = _SCAN_FIELDS if (discovery or existing is None) else _SCAN_FIELDS_ENRICH
        merged = upsert(rows, rec, set_columns=cols)
        # Recompute flags on the merged record so prior LLM-rescued values are
        # respected — otherwise rescue work would be silently re-flagged as missing.
        merged.extract_flags = compute_extract_flags(merged)

        if counters["scanned"] % 200 == 0:
            print(
                f"  ...{counters['scanned']}/{len(pdfs)} parsed={counters['parsed']} failed={counters['extract_failed']}",
                file=sys.stderr, flush=True,
            )

    print(f"saving {len(rows)} rows to {xlsx} ...", file=sys.stderr, flush=True)
    save_index(rows, xlsx)
    print(
        f"scan-archive done: scanned={counters['scanned']} parsed={counters['parsed']} "
        f"extract_failed={counters['extract_failed']} skipped={counters['skipped']} "
        f"(davon nicht im Index={counters.get('skipped_unknown_pk', 0)}, "
        f"Große Anfrage={counters.get('skipped_grosse', 0)}, "
        f"Anfrage-PDF={counters.get('skipped_anfrage', 0)}) "
        f"vocab_novelty={counters['novelty']} rows_in_xlsx={len(rows)}",
        file=sys.stderr, flush=True,
    )
    return 0


_CRAWL_FIELDS = {
    "wp", "kleine_anfrage_nr", "drucksache_anfrage_nr", "drucksache_antwort_nr",
    "anfrager", "fraktion", "anfragedatum", "anfragetitel",
    "antwortdatum", "ministerium_kuerzel",
    "systematik", "schlagworte", "link_anfrage", "link_antwort",
    "antworttext_status",
}


def _date_in_range(iso: str, lo: str | None, hi: str | None) -> bool:
    if lo and iso < lo:
        return False
    if hi and iso > hi:
        return False
    return True


def cmd_crawl(args: argparse.Namespace) -> int:
    """Discover and enrich Kleine Anfrage metadata via the Webflow search.

    Single pagination pass per Wahlperiode: each search hit carries BOTH
    Drucksachen (Anfrage + Antwort), both dates, the answering Ministerium-
    Kürzel, plus Systematik / Schlagworte / Anfrager / Fraktion / title.
    --full is a no-op for back-compat (full sweep is now the only mode).

    Row identity rules:
      - PK is always the Antwort-Drucksache.
      - For unanswered KAs (no Antwort-DS in the search hit), the Anfrage-DS
        is reused as a placeholder PK with status='pending'.
      - Old placeholder rows (PK == Anfrage-DS) get migrated to their real
        Antwort-DS when the search hit reveals it.
    """
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    counters = {"hits": 0, "matched_existing": 0, "new": 0,
                "migrated": 0, "filtered": 0, "errors": 0}

    with make_client(args.rps, args.user_agent) as client:
        try:
            tokens = bootstrap_search(client)
        except Exception as e:
            _append_log(CRAWL_ERRORS_LOG, f"{_now_iso()} | bootstrap failed: {e}")
            print(f"crawl: bootstrap failed: {e}", file=sys.stderr)
            return 2
        print(f"crawl: bootstrapped {tokens['post_url'][:120]}", file=sys.stderr, flush=True)
        print(f"crawl: full sweep for wp={args.wahlperiode}", file=sys.stderr, flush=True)

        html = search_post(client, tokens, wp=args.wahlperiode, rpp=args.rpp)
        page = 1
        while True:
            hits = parse_search_hits(html)
            counters["hits"] += len(hits)
            for h in hits:
                ad = h.anfragedatum
                if ad and not _date_in_range(ad, args.date_from, args.date_to):
                    counters["filtered"] += 1
                    continue
                if not h.drucksache_anfrage_nr:
                    continue

                # Look up both possible representations: the real Antwort-DS PK,
                # and any placeholder row keyed on the Anfrage-DS (legacy from
                # earlier discovery runs that didn't know the answer yet).
                real_pk = h.drucksache_antwort_nr
                real_row = rows.get(real_pk) if real_pk else None
                placeholder = None
                anfrage = h.drucksache_anfrage_nr
                ph_row = rows.get(anfrage)
                if (ph_row is not None and ph_row is not real_row
                        and ph_row.drucksache_anfrage_nr == anfrage
                        and ph_row.drucksache_antwort_nr == anfrage):
                    placeholder = ph_row

                if real_row and placeholder:
                    # Dedupe: drop the placeholder, the real row is authoritative.
                    del rows[placeholder.drucksache_antwort_nr]
                    counters["migrated"] += 1
                    existing = real_row
                elif real_row:
                    existing = real_row
                elif placeholder:
                    # Migrate the placeholder to its real PK (search now knows it).
                    if real_pk:
                        del rows[placeholder.drucksache_antwort_nr]
                        placeholder.drucksache_antwort_nr = real_pk
                        rows[real_pk] = placeholder
                        counters["migrated"] += 1
                    existing = placeholder
                else:
                    existing = None

                if existing:
                    h.drucksache_antwort_nr = existing.drucksache_antwort_nr
                    merged = upsert(
                        rows, h,
                        set_columns=_CRAWL_FIELDS - {"drucksache_antwort_nr", "antworttext_status"},
                    )
                    counters["matched_existing"] += 1
                else:
                    if h.antworttext_status == STATUS_ZURUECKGEZOGEN:
                        pass  # keep withdrawn-status for new rows
                    elif h.drucksache_antwort_nr:
                        h.antworttext_status = STATUS_PENDING  # PDF not yet local
                    else:
                        h.drucksache_antwort_nr = h.drucksache_anfrage_nr  # placeholder PK
                        h.antworttext_status = STATUS_PENDING
                    merged = upsert(rows, h, set_columns=_CRAWL_FIELDS)
                    counters["new"] += 1
                # Withdrawn KAs override any prior status (was likely PENDING /
                # FAILED because no real answer exists). Pin it on every crawl.
                if h.antworttext_status == STATUS_ZURUECKGEZOGEN:
                    merged.antworttext_status = STATUS_ZURUECKGEZOGEN
                merged.extract_flags = compute_extract_flags(merged)
            next_url = find_next_page_url(html)
            if not next_url:
                break
            page += 1
            if page % 10 == 0:
                print(
                    f"  page {page}, hits={counters['hits']} "
                    f"matched={counters['matched_existing']} new={counters['new']} "
                    f"migrated={counters['migrated']}",
                    file=sys.stderr, flush=True,
                )
                save_index(rows, xlsx)  # checkpoint
            html = search_get_page(client, next_url)

    # Cleanup pass: any placeholder row (PK == Anfrage-DS) for which a real-PK
    # row with the same Anfrage-DS now exists is a duplicate left over from
    # earlier crawl runs (or skipped during this run if the server returned a
    # short page). Drop the placeholder; the real row is authoritative.
    by_anfrage: dict[str, str] = {}
    for k, r in rows.items():
        if r.drucksache_anfrage_nr and r.drucksache_antwort_nr != r.drucksache_anfrage_nr:
            by_anfrage[r.drucksache_anfrage_nr] = k
    cleaned = 0
    for k in list(rows.keys()):
        r = rows[k]
        if r.drucksache_anfrage_nr and r.drucksache_antwort_nr == r.drucksache_anfrage_nr:
            if r.drucksache_anfrage_nr in by_anfrage and by_anfrage[r.drucksache_anfrage_nr] != k:
                del rows[k]
                cleaned += 1

    print(f"saving {len(rows)} rows to {xlsx} ...", file=sys.stderr, flush=True)
    apply_normalization(rows)
    save_index(rows, xlsx)
    print(
        f"crawl done: pages={page} hits={counters['hits']} "
        f"matched={counters['matched_existing']} new={counters['new']} "
        f"migrated={counters['migrated']} cleanup={cleaned} filtered={counters['filtered']}",
        file=sys.stderr, flush=True,
    )
    return 0


_FETCH_FIELDS = {
    "wp", "kleine_anfrage_nr", "drucksache_anfrage_nr", "drucksache_antwort_nr",
    "anfrager", "fraktion", "anfragedatum", "anfragetitel", "antwortdatum",
    "ministerium", "antworttext", "antworttext_status", "antworttext_quelle",
}


def _pdf_to_md(pdf_path: Path) -> str:
    """Full-document text extraction. pdfplumber primary, pypdf fallback."""
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            return "\n\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        reader = pypdf.PdfReader(str(pdf_path))
        return "\n\n".join(p.extract_text() or "" for p in reader.pages)


def _atomic_write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "w", encoding="utf-8", delete=False, dir=path.parent, prefix=".tmp_", suffix=path.suffix,
    ) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)
    os.replace(tmp_path, path)


def cmd_fetch_text(args: argparse.Namespace) -> int:
    """For each row missing PDF/.md: locate-or-download, backfill pages 1-3, full extract."""
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)

    # Candidates: rows with a real Drucksache_Antwort_Nr ('18/N' format) and
    # status not yet 'extracted' (or --force).
    candidates = []
    for key, rec in rows.items():
        if not re.match(r"\d+/\d+$", rec.drucksache_antwort_nr or ""):
            continue
        if args.wahlperiode is not None and rec.wp != args.wahlperiode:
            continue
        if rec.antworttext_status == STATUS_EXTRACTED and not args.force:
            continue
        if rec.antworttext_status == STATUS_NO_ANSWER:
            continue
        if rec.antworttext_status == STATUS_ZURUECKGEZOGEN:
            continue
        if is_placeholder_row(rec):
            # No separate Antwort-Drucksache yet → fetching the URL would
            # download the Anfrage-PDF and mislabel it as the answer.
            continue
        candidates.append((key, rec))
    if args.limit:
        candidates = candidates[: args.limit]

    print(f"fetch-text: {len(candidates)} candidate rows", file=sys.stderr, flush=True)

    counters = {"have": 0, "downloaded": 0, "extracted": 0, "no_answer": 0, "failed": 0}
    client = make_client(args.rps, args.user_agent)

    try:
        for i, (key, rec) in enumerate(candidates, 1):
            m = re.match(r"(\d+)/(\d+)$", rec.drucksache_antwort_nr)
            wp = int(m.group(1)); n = int(m.group(2))

            # Locate or download PDF
            pdf = archive_lookup(wp, n)
            if pdf:
                rec.antworttext_quelle = QUELLE_LOCAL
                counters["have"] += 1
            else:
                url = PDF_URL_TEMPLATE.format(wp=wp, n=n)
                if not PDF_URL_ALLOW.match(url):
                    _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {key} | url not allow-listed: {url}")
                    counters["failed"] += 1
                    continue
                try:
                    r = client.get(url)
                except Exception as e:
                    _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {key} | download error: {e}")
                    counters["failed"] += 1
                    continue
                if r.status_code == 404:
                    rec.antworttext_status = STATUS_NO_ANSWER
                    upsert(rows, rec, set_columns={"antworttext_status"})
                    counters["no_answer"] += 1
                    continue
                if r.status_code != 200:
                    _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {key} | http {r.status_code}")
                    counters["failed"] += 1
                    continue
                pdf = archive_target_path(wp, n)
                with tempfile.NamedTemporaryFile(
                    "wb", delete=False, dir=pdf.parent, prefix=".tmp_", suffix=".pdf",
                ) as tmp:
                    tmp.write(r.content)
                    tmp_path = Path(tmp.name)
                os.replace(tmp_path, pdf)
                rec.antworttext_quelle = QUELLE_DOWNLOADED
                counters["downloaded"] += 1

            # Backfill page-1-to-3 metadata if missing
            if not rec.kleine_anfrage_nr:
                try:
                    parsed = parse_page_text(pdftotext_first_pages(pdf))
                    parsed.pop("drucksache_antwort_nr", None)
                    for fname, val in parsed.items():
                        if fname in {f.name for f in fields(Record)} and not getattr(rec, fname):
                            setattr(rec, fname, val)
                except Exception as e:
                    _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {key} | page1-3 backfill: {e}")

            # Full text → .md
            md_path = pdf.with_suffix(".md")
            if md_path.exists() and not args.force:
                rec.antworttext = str(md_path.relative_to(REPO_ROOT))
                rec.antworttext_status = STATUS_EXTRACTED
                counters["extracted"] += 1
            else:
                try:
                    text = _pdf_to_md(pdf)
                except Exception as e:
                    _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {key} | pdf->md: {e}")
                    rec.antworttext_status = STATUS_FAILED
                    counters["failed"] += 1
                    upsert(rows, rec, set_columns=_FETCH_FIELDS)
                    continue
                _atomic_write_text(md_path, text)
                rec.antworttext = str(md_path.relative_to(REPO_ROOT))
                rec.antworttext_status = STATUS_EXTRACTED
                counters["extracted"] += 1

            upsert(rows, rec, set_columns=_FETCH_FIELDS)

            if i % 100 == 0:
                print(
                    f"  ...{i}/{len(candidates)} extracted={counters['extracted']} "
                    f"no_answer={counters['no_answer']} failed={counters['failed']}",
                    file=sys.stderr, flush=True,
                )
                save_index(rows, xlsx)
    finally:
        client.close()

    apply_normalization(rows)
    save_index(rows, xlsx)
    print(
        f"fetch-text done: have={counters['have']} downloaded={counters['downloaded']} "
        f"extracted={counters['extracted']} no_answer={counters['no_answer']} "
        f"failed={counters['failed']}",
        file=sys.stderr, flush=True,
    )
    return 0


_LLM_FIELDS = (
    "drucksache_anfrage_nr", "anfrager", "fraktion",
    "anfragedatum", "anfragetitel", "ministerium",
)

_LLM_SCHEMA = json.dumps({
    "type": "object",
    "properties": {
        "drucksache_anfrage_nr": {"type": ["string", "null"],
            "description": "Question Drucksache, format 'WP/N' (e.g. '18/27'). null if absent."},
        "anfrager": {"type": ["string", "null"],
            "description": "Asker name(s), comma-separated for multiple Abgeordnete."},
        "fraktion": {"type": ["string", "null"],
            "description": "One of: CDU, SPD, GRÜNE, FDP, AfD, fraktionslos."},
        "anfragedatum": {"type": ["string", "null"],
            "description": "ISO YYYY-MM-DD or null."},
        "anfragetitel": {"type": ["string", "null"],
            "description": "Title of the Kleine Anfrage."},
        "ministerium": {"type": ["string", "null"],
            "description": "Responding ministry full name (e.g. 'Ministerin für Schule und Bildung')."},
    },
    "required": [],
    "additionalProperties": False,
})

_LLM_PROMPT = (
    "Aus diesem Antwort-Drucksache-Text des Landtags NRW (Seiten 1-3) "
    "die Metadaten der Kleinen Anfrage extrahieren. Felder die nicht "
    "vorhanden sind: null. Anfragedatum als ISO YYYY-MM-DD.\n\n"
    "---\n{text}\n---"
)


def cmd_enrich_llm(args: argparse.Namespace) -> int:
    """LLM-based field extraction for rows the rule-based parser missed.

    Selects rows that have a local PDF (status != 'pending') and at least one
    missing _LLM_FIELDS value. Sends pdftotext output to the configured `llm`
    model with a JSON schema, fills only previously-empty fields.
    """
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)

    candidates = []
    for key, rec in rows.items():
        if rec.antworttext_status == STATUS_PENDING:
            continue  # placeholder row, no local PDF
        if args.wahlperiode is not None and rec.wp != args.wahlperiode:
            continue
        # Quality gate: only rows with extract_flags set (missing required
        # fields or novel_fraktion). Rows without flags are clean — skip.
        if not rec.extract_flags:
            continue
        m = re.match(r"(\d+)/(\d+)$", key)
        if not m:
            continue
        wp, n = int(m.group(1)), int(m.group(2))
        pdf = archive_lookup(wp, n)
        if not pdf:
            continue
        candidates.append((key, rec, pdf))

    if args.limit:
        candidates = candidates[: args.limit]
    print(f"enrich-llm: {len(candidates)} candidate rows ({args.model})", file=sys.stderr, flush=True)

    counters = {"queried": 0, "filled": 0, "errors": 0, "noop": 0}
    try:
        for i, (key, rec, pdf) in enumerate(candidates, 1):
            try:
                text = pdftotext_first_pages(pdf)
            except Exception as e:
                _append_log(VERIFY_LLM_LOG, f"{_now_iso()} | {key} | pdftotext: {e}")
                counters["errors"] += 1
                continue
            prompt = _LLM_PROMPT.format(text=text[:8000])
            try:
                r = subprocess.run(
                    ["llm", "-m", args.model, "--schema", _LLM_SCHEMA, prompt],
                    capture_output=True, text=True, timeout=120, check=True,
                )
                data = json.loads(r.stdout)
            except subprocess.CalledProcessError as e:
                _append_log(VERIFY_LLM_LOG, f"{_now_iso()} | {key} | llm exit {e.returncode}: {e.stderr[:200]}")
                counters["errors"] += 1
                continue
            except Exception as e:
                _append_log(VERIFY_LLM_LOG, f"{_now_iso()} | {key} | parse: {e}")
                counters["errors"] += 1
                continue
            counters["queried"] += 1
            any_filled = False
            for fname in _LLM_FIELDS:
                v = data.get(fname)
                if v and not getattr(rec, fname):
                    setattr(rec, fname, v.strip() if isinstance(v, str) else v)
                    any_filled = True
            # Recompute flags so reruns skip rows the LLM already cleaned.
            rec.extract_flags = compute_extract_flags(rec)
            if any_filled:
                counters["filled"] += 1
            else:
                counters["noop"] += 1
            if i % 25 == 0:
                print(
                    f"  ...{i}/{len(candidates)} filled={counters['filled']} "
                    f"errors={counters['errors']} noop={counters['noop']}",
                    file=sys.stderr, flush=True,
                )
                save_index(rows, xlsx)
    finally:
        save_index(rows, xlsx)

    print(
        f"enrich-llm done: queried={counters['queried']} filled={counters['filled']} "
        f"errors={counters['errors']} noop={counters['noop']}",
        file=sys.stderr, flush=True,
    )
    return 0


# --- canonical Fraktion / Ministerium normalisation --------------------------
#
# Index/fraktionen.xlsx and Index/ministerien.xlsx are user-curated tables.
# `normalize` reads them and fills three xlsx columns:
#   - Fraktion_Canonical       (exact-match rules; very few variants expected)
#   - Ministerium_Canonical    (token-Jaccard match; many spelling variants)
#   - Ministerium_Kuerzel      (the table's "Kürzel" column for the matched row)

INDEX_DIR = REPO_ROOT / "Index"
FRAKTIONEN_XLSX = INDEX_DIR / "fraktionen.xlsx"
MINISTERIEN_XLSX = INDEX_DIR / "ministerien.xlsx"

# Words ignored when token-matching ministry strings.
_MIN_STOPWORDS = frozenset({
    "ministerium", "minister", "ministerin",
    "ministerpräsident", "ministerpräsidentin",
    "der", "des", "die", "das", "den", "dem",
    "für", "fuer",  # tokenizer ASCII-folds 'ü' → 'ue', so 'für' arrives as 'fuer'
    "und", "sowie", "oder", "des",
    "landes", "nrw", "nordrhein-westfalen",
    "nordrhein", "westfalen",  # tokenizer splits the hyphen, drop the parts too
    "namens", "landesregierung",  # boilerplate phrase 'namens der Landesregierung'
    "chef", "chefin",
    "staatskanzlei",  # appears in 2 ministries; intentional drop so that
                      # "Chef der Staatskanzlei" alone doesn't dominate the score
})


def _ministerium_tokens(s: str) -> set[str]:
    """Tokenise a ministry string conservatively for SUBSET matching.

    Steps: lowercase + ASCII-fold umlauts; alpha-only tokens; strip the
    trailing 'ministerium' / 'minister(in)?' compound (so 'Justizministerium'
    becomes 'Justiz'); drop stopwords and tokens shorter than 4 chars.

    NOT done: prefix matching, flexion stripping, fuzzy edits. Anything not
    expressible as exact-subset is the operator's job to put into
    Index/ministerium_aliases.xlsx — that keeps the matcher false-positive-free.
    """
    s = s.lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    raw = re.findall(r"[a-z]+", s)
    out: set[str] = set()
    for t in raw:
        for suf in ("ministeriums", "ministerium", "ministerinnen", "ministerin", "ministers", "minister"):
            if t.endswith(suf) and len(t) > len(suf):
                t = t[: -len(suf)]
                break
        if t in _MIN_STOPWORDS or len(t) < 4:
            continue
        out.add(t)
    return out


def load_canon_fraktionen() -> set[str]:
    """Return the canonical Fraktion set from Index/fraktionen.xlsx (column 'Fraktion')."""
    if not FRAKTIONEN_XLSX.exists():
        return set()
    wb = openpyxl.load_workbook(FRAKTIONEN_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None) or []
    try:
        col = list(header).index("Fraktion")
    except ValueError:
        wb.close()
        return set()
    out = {row[col] for row in rows_iter if row and row[col]}
    wb.close()
    return out


def load_canon_ministerien() -> list[tuple[str, str, set[str]]]:
    """Return [(Kürzel, Ministerium-Name, token-set)] from Index/ministerien.xlsx."""
    if not MINISTERIEN_XLSX.exists():
        return []
    wb = openpyxl.load_workbook(MINISTERIEN_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None) or []
    try:
        kc = list(header).index("Kürzel")
        nc = list(header).index("Ministerium")
    except ValueError:
        wb.close()
        return []
    out = []
    for row in rows_iter:
        if not row or not row[nc]:
            continue
        full = str(row[nc])
        kuerzel = str(row[kc]) if row[kc] else ""
        out.append((kuerzel, full, _ministerium_tokens(full)))
    wb.close()
    return out


def load_kuerzel_aliases() -> dict[str, str]:
    """Return {alias_kuerzel: primary_kuerzel} from Index/ministerien.xlsx
    'Aliases' column. Aliases are comma-separated. Used to merge
    Kürzel-Schreibweisen the search hit may emit (e.g. MCdS → MBEIM)."""
    if not MINISTERIEN_XLSX.exists():
        return {}
    wb = openpyxl.load_workbook(MINISTERIEN_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None) or []
    header_list = list(header)
    try:
        kc = header_list.index("Kürzel")
    except ValueError:
        wb.close()
        return {}
    if "Aliases" not in header_list:
        wb.close()
        return {}
    ac = header_list.index("Aliases")
    out: dict[str, str] = {}
    for row in rows_iter:
        if not row or not row[kc] or not row[ac]:
            continue
        primary = str(row[kc])
        for alias in str(row[ac]).split(","):
            alias = alias.strip()
            if alias:
                out[alias] = primary
    wb.close()
    return out


MINISTERIUM_ALIASES_XLSX = INDEX_DIR / "ministerium_aliases.xlsx"


def load_ministerium_aliases() -> dict[str, str]:
    """Operator-curated raw → Kürzel overrides. Returns {} if file missing.

    Format: cols 'Raw' (the verbatim Ministerium string as it appears in
    data/index.xlsx) and 'Kuerzel' (must match a Kürzel in ministerien.xlsx).
    """
    if not MINISTERIUM_ALIASES_XLSX.exists():
        return {}
    wb = openpyxl.load_workbook(MINISTERIUM_ALIASES_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None) or []
    try:
        rc = list(header).index("Raw")
        kc = list(header).index("Kuerzel")
    except ValueError:
        wb.close()
        return {}
    out: dict[str, str] = {}
    for row in rows_iter:
        if row and row[rc] and row[kc]:
            out[str(row[rc])] = str(row[kc])
    wb.close()
    return out


def match_ministerium(value: str, canonicals: list[tuple[str, str, set[str]]]) -> tuple[str, str] | None:
    """Strict subset match: every value-token must appear in the canonical
    token set. False-positive-free by construction — anything that needs a
    fuzzy/flexion call belongs in ministerium_aliases.xlsx instead.

    Ties (multiple canonicals satisfy the subset) go to the canonical with
    the FEWEST tokens (most specific), then to the longer name as fallback.
    """
    if not value:
        return None
    val_tokens = _ministerium_tokens(value)
    if not val_tokens:
        return None
    candidates = [(k, f, ct) for k, f, ct in canonicals if ct and val_tokens.issubset(ct)]
    if not candidates:
        return None
    candidates.sort(key=lambda x: (len(x[2]), -len(x[1])))
    k, f, _ = candidates[0]
    return (k, f)


# --- multi-ministerium parser (spec: 2026-05-04-multi-ministerium-parser.md) -

# Phase A: locate the "Der Minister/Die Ministerin … (hat) die Kleine Anfrage
# N … beantwortet." sentence — the boundary between the cited Anfragetext
# (Schritt 6 in the spec) and the actual answers (Schritt 8).
#
# Why anchor on "die Kleine Anfrage <N> … beantwortet": the Vorbemerkung often
# quotes prior answers verbatim, so a bare "Der Minister … beantwortet" can
# false-positive. The literal "die Kleine Anfrage <KA-Nr>" reference appears
# only in the actual boundary paragraph.
#
# Why search the flattened (whitespace-collapsed) text: pdftotext glues the
# running page header "LANDTAG NORDRHEIN-WESTFALEN - <wp>. Wahlperiode
# Drucksache N/M" onto the boundary paragraph with a single newline, so
# blank-line paragraph splits don't isolate it cleanly.
_RX_MIN_BOUNDARY = re.compile(
    r"(?:Der|Die)\s+Minister(?:präsident(?:in)?|in)?\b"
    r".{0,800}?"
    r"\bdie\s+Kleine\s+Anfrage\s+\d+"
    r".{0,400}?"
    r"\bbeantwortet\b[^.]{0,30}\."
)


def find_minister_paragraph(text: str) -> str | None:
    """Locate the boundary sentence that names the answering minister(s).
    Returns it as a single whitespace-normalized line, or None if not found."""
    flat = re.sub(r"\s+", " ", text)
    m = _RX_MIN_BOUNDARY.search(flat)
    return m.group(0) if m else None


# Phase B: extract minister full-forms from the boundary paragraph.
# Pattern: <Determiner> Minister[präsident][in] <connective> <body> <stop>
# Determiners: Der/Die (Nominativ), dem/der (Dativ — used in
#   "im Einvernehmen mit dem Minister …" / "der Ministerin …").
# Connectives: für (most common — "Minister für X"), des/der (gen — "des
#   Innern", "der Justiz").
# Stop tokens: link phrases that bridge to the next minister (und/sowie are
#   the most common), plus the sentence-final verb "beantwortet" and the
#   "(im) Einvernehmen mit" / "in Abstimmung mit" boilerplate. The next
#   determiner ("dem"/"der" before "Minister") also stops the body — without
#   it, "Minister des Innern und dem Minister der Justiz" parses as one form.
#   Comma is NOT a stop — body itself often contains commas (e.g. "Kinder,
#   Jugend, Familie, Gleichstellung, Flucht und Integration").
_RX_MIN_FORM = re.compile(
    r"\b(?:Der|Die|dem|der|des|den)\s+"
    r"(Minister(?:präsident(?:in)?|in)?)\s+"
    r"(für|der|des)\s+"
    r"(.+?)"
    # Stop alternatives — each carries its own leading-space rule. The
    # comma stop is `\s*,` so it fires even when the body ends right at
    # the comma (e.g. "Finanzen, dem Minister …"). The 'namens' / 'des
    # Landes' stops keep the boilerplate out of the body so the
    # operator-curated aliases (Index/ministerium_aliases.xlsx) still
    # match by exact string after we drop the determiner-connective.
    r"(?="
    r"\s+hat\b|\s+haben\b|"
    r"\s+im\s+Einvernehmen\b|\s+in\s+Abstimmung\b|"
    r"\s+gemeinsam\s+mit\b|\s+nach\s+Beteiligung\b|\s+unter\s+Mitwirkung\b|"
    r"\s+beantwortet\b|\s+beantworten\b|"
    r"\s+wie\s+folgt\b|"
    r"\s+namens\s+der\s+Landesregierung\b|"
    r"\s+des\s+Landes\s+Nordrhein-Westfalen\b|"
    r"\s+und\s+(?:mit\s+)?(?:dem|der|des|den)\s+Minister(?:präsident(?:in)?|in)?\b|"
    r"\s+sowie\s+(?:mit\s+)?(?:dem|der|des|den)?\s*Minister(?:präsident(?:in)?|in)?\b|"
    r"\s*,\s+(?:dem|der|des|den)\s+Minister(?:präsident(?:in)?|in)?\b"
    r")"
)


def _undo_pdf_hyphenation(s: str) -> str:
    """Undo pdftotext hyphenation quirks:
      1. Soft-hyphen leak with no space: 'Klima-schutz' → 'Klimaschutz',
         'Gleich-stellung' → 'Gleichstellung'.
      2. Line-wrap hyphen with space: 'In- nern' → 'Innern'. Skip when the
         next word is 'und/oder/sowie/noch' — those follow a grammatical
         Bindestrich-Ergänzung (e.g. 'Bundes- und Europaangelegenheiten')
         that must NOT be joined.
      3. pdftotext occasionally eats the space after a Bindestrich-Ergänzung,
         producing 'Bundesund'. Split it back so the canon tokenizer can
         reach the 'bundes' / 'und' tokens.
    """
    s = re.sub(r"(\w)-([a-zäöüß])", r"\1\2", s)
    s = re.sub(r"(\w)-\s+(?!(?:und|oder|sowie|noch)\b)([a-zäöüß])", r"\1\2", s)
    s = re.sub(r"\bBundesund\b", "Bundes und", s)
    return s


def extract_minister_full_forms(paragraph: str) -> list[str]:
    """Return ordered list of '<Rolle> <Body>' strings — duplicates preserved
    (caller dedupes on Kürzel after resolving). The connective ('für/der/des')
    is dropped so the output aligns with the format used in
    Index/ministerium_aliases.xlsx ('Minister Innern', not 'Minister des
    Innern')."""
    paragraph = _undo_pdf_hyphenation(paragraph)
    out: list[str] = []
    for m in _RX_MIN_FORM.finditer(paragraph):
        rolle, body = m.group(1), m.group(3).strip()
        out.append(f"{rolle} {body}")
    return out


def resolve_minister_form(
    full_form: str,
    canon_min: list[tuple[str, str, set[str]]],
    aliases: dict[str, str],
) -> str | None:
    """Map an extracted full-form ('Minister des Innern', 'Ministerin für …')
    to a Kürzel. Tries operator-curated aliases first, then strict subset
    match against ministerien.xlsx tokens. Returns None if neither resolves."""
    if full_form in aliases:
        return aliases[full_form]
    hit = match_ministerium(full_form, canon_min)
    return hit[0] if hit else None


def _count_beteiligte(kuerzel_str: str) -> int:
    """Count of comma-separated Kürzel in a Beteiligte_Ministerien_Kuerzel value."""
    if not kuerzel_str:
        return 0
    return sum(1 for k in kuerzel_str.split(",") if k.strip())


def cmd_extract_multi_ministerium(args: argparse.Namespace) -> int:
    """Read each row's Antworttext .md, parse the boundary paragraph for ALL
    answering ministries, write Kürzel-Set to Beteiligte_Ministerien_Kuerzel.

    Federführend (= existing Ministerium_Kuerzel) is forced first; further
    Kürzel follow in order-of-appearance from the PDF text. Idempotent —
    overwrites Beteiligte_* on every run.
    """
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    canon_min = load_canon_ministerien()
    aliases = load_ministerium_aliases()
    kuerzel_aliases = load_kuerzel_aliases()

    counters = {"scanned": 0, "no_para": 0, "single": 0, "multi": 0,
                "novel_form": 0, "skipped_status": 0}

    for rec in rows.values():
        if args.wahlperiode is not None and rec.wp != args.wahlperiode:
            continue
        if rec.antworttext_status != STATUS_EXTRACTED:
            counters["skipped_status"] += 1
            continue
        if not rec.antworttext:
            continue
        md_path = REPO_ROOT / rec.antworttext
        if not md_path.exists():
            continue

        counters["scanned"] += 1
        try:
            text = md_path.read_text(encoding="utf-8")
        except Exception as e:
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {rec.drucksache_antwort_nr} | multi-min read: {e}")
            continue

        para = find_minister_paragraph(text)
        if not para:
            counters["no_para"] += 1
            rec.beteiligte_ministerien_kuerzel = ""
            rec.beteiligte_ministerien = 0
            continue

        # Phase B: full-forms → Kürzel, deduped, ordered
        kuerzel_order: list[str] = []
        if rec.ministerium_kuerzel:
            primary = kuerzel_aliases.get(rec.ministerium_kuerzel, rec.ministerium_kuerzel)
            kuerzel_order.append(primary)

        for form in extract_minister_full_forms(para):
            k = resolve_minister_form(form, canon_min, aliases)
            if k:
                k = kuerzel_aliases.get(k, k)
                if k not in kuerzel_order:
                    kuerzel_order.append(k)
            else:
                counters["novel_form"] += 1
                log_vocab_novelty(rec.drucksache_antwort_nr or "", "ministerium_form", form)

        rec.beteiligte_ministerien_kuerzel = ",".join(kuerzel_order)
        rec.beteiligte_ministerien = len(kuerzel_order)
        if len(kuerzel_order) >= 2:
            counters["multi"] += 1
        elif len(kuerzel_order) == 1:
            counters["single"] += 1

    save_index(rows, xlsx)
    print(
        f"extract-multi-ministerium done: scanned={counters['scanned']} "
        f"single={counters['single']} multi={counters['multi']} "
        f"no_para={counters['no_para']} novel_forms={counters['novel_form']}",
        file=sys.stderr,
    )
    return 0


def apply_normalization(rows: dict) -> dict:
    """Mutate rows in place: fill canonical / Kürzel columns, recompute flags.

    Returns counters dict. Pure data — no I/O. Callable from any pipeline
    stage that wants the index to stay normalized after a write.
    """
    canon_frak = load_canon_fraktionen()
    canon_min = load_canon_ministerien()
    aliases = load_ministerium_aliases()
    kuerzel_aliases = load_kuerzel_aliases()
    kuerzel_to_full = {k: f for k, f, _ in canon_min}

    counters = {"frak_matched": 0, "frak_novel": 0,
                "min_matched_alias": 0, "min_matched_subset": 0, "min_novel": 0,
                "min_matched_kuerzel": 0, "alias_orphan": 0}

    for rec in rows.values():
        # Fraktion: exact match against canonical set
        if rec.fraktion:
            if rec.fraktion in canon_frak:
                rec.fraktion_canonical = rec.fraktion
                counters["frak_matched"] += 1
            else:
                rec.fraktion_canonical = ""
                counters["frak_novel"] += 1

        # Ministerium resolution priority:
        #   1. direct Kürzel (set by crawl from the search hit's "Antwort KÜRZEL"
        #      token; authoritative when present);
        #   2. operator-curated alias table;
        #   3. strict subset token match against canonical names (FP-frei).
        # The Kürzel from the database is preserved as-is — never overwritten
        # by an alias mapping. The `Aliases` column in ministerien.xlsx is
        # used only for the canonical-name lookup, so historical/variant
        # spellings (e.g. MKJFGF) still resolve to the right full ministry
        # name without losing the original Kürzel value.
        if rec.ministerium_kuerzel:
            kz = rec.ministerium_kuerzel
            # Try direct lookup first (for entries with their own row in
            # ministerien.xlsx, e.g. MCdS); fall back to alias resolution.
            full = kuerzel_to_full.get(kz)
            if not full and kz in kuerzel_aliases:
                full = kuerzel_to_full.get(kuerzel_aliases[kz])
            if full:
                rec.ministerium_canonical = full
                counters["min_matched_kuerzel"] = counters.get("min_matched_kuerzel", 0) + 1
            else:
                # Kürzel set but unknown — likely a parser misfire (e.g. "BT"
                # picked up from "Antwort BT-Drs." in the Inhaltsbeschreibung)
                # or a source-data typo on landtag.nrw.de (e.g. "AWEL"). Surface
                # via flag + novelty log so an operator can reconcile.
                rec.ministerium_canonical = ""
                counters["min_novel"] += 1
                log_vocab_novelty(rec.drucksache_antwort_nr or "", "ministerium_kuerzel", kz)
        elif rec.ministerium:
            kuerzel = aliases.get(rec.ministerium)
            if kuerzel:
                full = kuerzel_to_full.get(kuerzel)
                if full:
                    rec.ministerium_kuerzel = kuerzel
                    rec.ministerium_canonical = full
                    counters["min_matched_alias"] += 1
                else:
                    rec.ministerium_kuerzel = ""
                    rec.ministerium_canonical = ""
                    counters["alias_orphan"] += 1
            else:
                m = match_ministerium(rec.ministerium, canon_min)
                if m:
                    rec.ministerium_kuerzel, rec.ministerium_canonical = m
                    counters["min_matched_subset"] += 1
                else:
                    rec.ministerium_kuerzel = ""
                    rec.ministerium_canonical = ""
                    counters["min_novel"] += 1
        elif rec.ministerium_kuerzel:
            # Kürzel set but not in ministerien.xlsx → flag, don't guess.
            rec.ministerium_canonical = ""
            counters["min_novel"] += 1

        # Keep Beteiligte_Ministerien (count) in sync with the Kürzel column
        # after every normalize — so a backfilled or hand-edited Kürzel string
        # immediately reflects in the count column.
        rec.beteiligte_ministerien = _count_beteiligte(rec.beteiligte_ministerien_kuerzel)

        # Rebuild flag set: missing_* recomputed (so stale flags drop after a
        # field got filled), then novel_ministerium added if no canonical resolved.
        base = compute_extract_flags(rec)
        # Preserve non-missing/non-novel/non-unknown flags caller may have set elsewhere.
        existing = [f for f in rec.extract_flags.split(",")
                    if f and not f.startswith("missing_")
                    and not f.startswith("novel_")
                    and f != "unknown_ministerium_kuerzel"]
        flags = existing + ([f for f in base.split(",") if f] if base else [])
        if rec.ministerium and not rec.ministerium_canonical:
            flags.append("novel_ministerium")
        # Kürzel set but not resolvable against ministerien.xlsx → fragwürdig.
        if rec.ministerium_kuerzel and not rec.ministerium_canonical:
            flags.append("unknown_ministerium_kuerzel")
        # de-dupe, preserve order
        seen = set(); out_flags = []
        for f in flags:
            if f and f not in seen:
                out_flags.append(f); seen.add(f)
        rec.extract_flags = ",".join(out_flags)

    return counters


def cmd_normalize(args: argparse.Namespace) -> int:
    """Fill Fraktion_Canonical / Ministerium_Canonical / Ministerium_Kuerzel.

    Sources, applied in priority:
      1. Index/ministerium_aliases.xlsx — operator-curated raw → Kürzel.
         Always wins; this is how typos / odd wording get mapped without
         loosening the rule-based matcher.
      2. Index/ministerien.xlsx — strict subset token match against the
         canonical Ministerium names (false-positive-free by construction).
      3. Index/fraktionen.xlsx — exact-string match for Fraktion (very few
         variants are expected; novelty is genuine, not a parse fluke).

    Anything still unmatched gets a novel_fraktion / novel_ministerium flag."""
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    if not load_canon_fraktionen():
        print(f"normalize: WARN no canonical Fraktionen at {FRAKTIONEN_XLSX}", file=sys.stderr)
    if not load_canon_ministerien():
        print(f"normalize: WARN no canonical Ministerien at {MINISTERIEN_XLSX}", file=sys.stderr)
    counters = apply_normalization(rows)
    save_index(rows, xlsx)
    print(
        f"normalize done: "
        f"fraktion matched={counters['frak_matched']} novel={counters['frak_novel']}; "
        f"ministerium alias={counters['min_matched_alias']} subset={counters['min_matched_subset']} "
        f"kuerzel={counters['min_matched_kuerzel']} novel={counters['min_novel']} "
        f"alias_orphan={counters['alias_orphan']}",
        file=sys.stderr,
    )
    return 0


def cmd_resolve(args: argparse.Namespace) -> int:
    """Targeted re-search for one or more KA-Nrn; repair index in place.

    Use case: a row in index.xlsx looks suspicious (e.g. md ↔ KA-Nr mismatch
    flagged by `verify`, or a gap reported in the contiguous KA-Nr report).
    This verb re-queries the live search with a strict KA-Nr + WP filter and
    overwrites whatever the index has for that KA-Nr with the authoritative
    search-hit data.

    --counter-search: if no KA hit is found, re-query with doktyp=GA. A hit
    there means the existing index row(s) for that KA-Nr are phantoms (a
    Große-Anfrage answer accidentally matched as that KA's answer); the row
    is reported and — with --delete-phantom — deleted from the index.
    """
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    wp = args.wahlperiode
    targets = list(args.ka)

    counters = {"resolved": 0, "no_hit": 0, "phantom_ga": 0, "phantom_deleted": 0,
                "added": 0, "updated": 0}

    with make_client(args.rps, args.user_agent) as client:
        try:
            tokens = bootstrap_search(client)
        except Exception as e:
            print(f"resolve: bootstrap failed: {e}", file=sys.stderr)
            return 2
        for ka_nr in targets:
            print(f"\n--- KA {ka_nr} (WP{wp}) ---", file=sys.stderr)
            html = search_post(client, tokens, wp=wp, nummer=str(ka_nr), rpp=10,
                               doktyp=DOKTYP_KLEINE_ANFRAGE)
            hits = [h for h in parse_search_hits(html)
                    if h.kleine_anfrage_nr == ka_nr and h.wp == wp]
            if not hits:
                print(f"  no Kleine-Anfrage hit for KA {ka_nr}", file=sys.stderr)
                if args.counter_search:
                    # Re-bootstrap to get a fresh single-use Webflow token.
                    tokens = bootstrap_search(client)
                    html2 = search_post(client, tokens, wp=wp, nummer=str(ka_nr),
                                        rpp=10, doktyp=DOKTYP_GROSSE_ANFRAGE)
                    ga_hits = parse_search_hits(html2)
                    if ga_hits:
                        print(f"  ! Gegenrecherche: KA-Nr {ka_nr} ist eine Große Anfrage",
                              file=sys.stderr)
                        for h in ga_hits[:3]:
                            print(f"    GA hit: Anfrage={h.drucksache_anfrage_nr} "
                                  f"Antwort={h.drucksache_antwort_nr} "
                                  f"({h.anfragetitel[:60]!r})", file=sys.stderr)
                        # Phantom-row identification: any row in our index whose
                        # KA-Nr matches the GA-Nr is misclassified data.
                        phantoms = [k for k, r in rows.items()
                                    if r.wp == wp and r.kleine_anfrage_nr == ka_nr]
                        for p in phantoms:
                            print(f"    Phantom row: {p}", file=sys.stderr)
                            counters["phantom_ga"] += 1
                            if args.delete_phantom:
                                del rows[p]
                                counters["phantom_deleted"] += 1
                        # Re-bootstrap so the next iteration's KA query has a fresh token.
                        tokens = bootstrap_search(client)
                        continue
                counters["no_hit"] += 1
                # Bootstrap a fresh token for the next loop iteration too.
                tokens = bootstrap_search(client)
                continue

            # Successful KA hit. Apply hit data to the matching row (or create one).
            h = hits[0]
            if not h.drucksache_antwort_nr:
                pk = h.drucksache_anfrage_nr  # placeholder PK for unanswered
                if h.antworttext_status != STATUS_ZURUECKGEZOGEN:
                    h.antworttext_status = STATUS_PENDING
                h.drucksache_antwort_nr = pk
            else:
                pk = h.drucksache_antwort_nr

            existing = rows.get(pk)
            # Also identify any pre-existing row sharing the same KA-Nr but a
            # different PK — those are stale (e.g. KA 144 → 18/10805 phantom).
            stale = [k for k, r in rows.items()
                     if r.wp == wp and r.kleine_anfrage_nr == ka_nr and k != pk]

            if existing:
                upsert(rows, h,
                       set_columns=_CRAWL_FIELDS - {"drucksache_antwort_nr", "antworttext_status"})
                counters["updated"] += 1
            else:
                upsert(rows, h, set_columns=_CRAWL_FIELDS)
                counters["added"] += 1
            merged = rows[pk]
            if h.antworttext_status == STATUS_ZURUECKGEZOGEN:
                merged.antworttext_status = STATUS_ZURUECKGEZOGEN
            merged.extract_flags = compute_extract_flags(merged)

            for s in stale:
                print(f"  stale row {s} (same KA-Nr, different Antwort-DS) "
                      f"{'→ deleted' if args.delete_phantom else '— use --delete-phantom to remove'}",
                      file=sys.stderr)
                if args.delete_phantom:
                    del rows[s]
                    counters["phantom_deleted"] += 1

            print(f"  resolved KA {ka_nr} → {pk} ({merged.anfrager}, {merged.fraktion}, "
                  f"{merged.ministerium_kuerzel or merged.antworttext_status})", file=sys.stderr)
            counters["resolved"] += 1
            # Re-bootstrap before the next KA — Webflow tokens are single-use.
            tokens = bootstrap_search(client)

    apply_normalization(rows)
    save_index(rows, xlsx)
    print(
        f"\nresolve done: resolved={counters['resolved']} no_hit={counters['no_hit']} "
        f"added={counters['added']} updated={counters['updated']} "
        f"phantom_ga={counters['phantom_ga']} phantom_deleted={counters['phantom_deleted']}",
        file=sys.stderr,
    )
    return 0


def cmd_verify(args: argparse.Namespace) -> int:
    """Read-only report; optional --llm-plausibility / --llm-rescue-fields."""
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    if not rows:
        print("verify: index.xlsx is empty or absent", file=sys.stderr)
        return 1

    # Counts by WP and status
    by_wp: dict[int, int] = {}
    by_status: dict[str, int] = {}
    md_missing: list[str] = []
    md_short: list[str] = []
    md_kanr_mismatch: list[str] = []  # md text does NOT contain "Kleine Anfrage <Nr>"
    md_grosse_anfrage: list[str] = []  # md text is a Große Anfrage answer
    bad_nr: list[str] = []
    novel_min: dict[str, int] = {}
    novel_frak: dict[str, int] = {}
    by_wp_ka: dict[int, set[int]] = {}            # for KA gap detection
    dup_ka: dict[tuple[int, int], list[str]] = {}  # (wp, ka_nr) → [pks]

    for key, rec in rows.items():
        by_wp[rec.wp or 0] = by_wp.get(rec.wp or 0, 0) + 1
        by_status[rec.antworttext_status] = by_status.get(rec.antworttext_status, 0) + 1
        if rec.antworttext_status == STATUS_EXTRACTED and rec.antworttext:
            md_path = REPO_ROOT / rec.antworttext
            if not md_path.exists():
                md_missing.append(key)
            else:
                size = md_path.stat().st_size
                if size < 200:
                    md_short.append(f"{key} ({size}B)")
                elif rec.kleine_anfrage_nr:
                    # Sanity: the answer text must mention "Kleine Anfrage <Nr>".
                    # Catches mis-linked rows (e.g. KA 144 row pointing at 18/10805
                    # which is a Große Anfrage answer). Only first 4 KB read — the
                    # phrase appears on page 1.
                    try:
                        head = md_path.read_text(encoding="utf-8", errors="replace")[:4000]
                    except OSError:
                        head = ""
                    if not re.search(rf"Kleine Anfrage\s+{rec.kleine_anfrage_nr}\b", head):
                        md_kanr_mismatch.append(f"{key} (KA {rec.kleine_anfrage_nr})")
                        if is_grosse_anfrage(head):
                            md_grosse_anfrage.append(f"{key} (KA {rec.kleine_anfrage_nr})")
        if rec.drucksache_antwort_nr and not re.match(r"\d+/\d+$", rec.drucksache_antwort_nr):
            bad_nr.append(f"{key}: drucksache_antwort_nr={rec.drucksache_antwort_nr!r}")
        if rec.fraktion and not check_fraktion(rec.fraktion):
            novel_frak[rec.fraktion] = novel_frak.get(rec.fraktion, 0) + 1
        if rec.wp and rec.kleine_anfrage_nr:
            by_wp_ka.setdefault(rec.wp, set()).add(rec.kleine_anfrage_nr)
            dup_ka.setdefault((rec.wp, rec.kleine_anfrage_nr), []).append(key)

    # Gap detection: contiguous KA-Nr from 1..max(seen) per WP
    ka_gaps: dict[int, list[int]] = {}
    for wp, seen in by_wp_ka.items():
        if not seen:
            continue
        hi = max(seen)
        gaps = sorted(set(range(1, hi + 1)) - seen)
        if gaps:
            ka_gaps[wp] = gaps

    # Duplicate KA-Nr (same wp+ka_nr appearing in >1 row, e.g. correct row +
    # garbage row whose answer-PDF was misidentified as that KA).
    dup_lines: list[str] = []
    for (wp, ka), pks in dup_ka.items():
        if len(pks) > 1:
            dup_lines.append(f"WP{wp} KA {ka}: {', '.join(pks)}")

    # Orphan PDF/MD files (no matching xlsx row)
    keys_set = set(rows.keys())
    orphans = []
    for pdf in iter_archive_pdfs():
        try:
            _, da = parse_filename_drucksache_nr(pdf)
        except ValueError:
            continue
        if da not in keys_set:
            orphans.append(pdf.name)

    print(f"=== verify @ {_now_iso()} ===")
    print(f"rows: {len(rows)}")
    print("by WP:", ", ".join(f"{wp}={n}" for wp, n in sorted(by_wp.items())))
    print("by status:", ", ".join(f"{s}={n}" for s, n in sorted(by_status.items())))
    print()
    print(f"md missing (status=extracted but file gone): {len(md_missing)}")
    for k in md_missing[:10]:
        print(f"  {k}")
    print(f"md suspiciously short (<200B): {len(md_short)}")
    for k in md_short[:10]:
        print(f"  {k}")
    print(f"md ↔ KA-Nr mismatch (md text lacks 'Kleine Anfrage <Nr>'): {len(md_kanr_mismatch)}")
    for k in md_kanr_mismatch[:10]:
        print(f"  {k}")
    print(f"  davon Große Anfragen (out-of-scope, sollten gelöscht werden): {len(md_grosse_anfrage)}")
    for k in md_grosse_anfrage[:10]:
        print(f"    {k}")
    print(f"malformed Drucksache_*_Nr: {len(bad_nr)}")
    for k in bad_nr[:10]:
        print(f"  {k}")
    print(f"orphan PDFs in Archiv (no matching row): {len(orphans)}")
    for k in orphans[:10]:
        print(f"  {k}")
    print()
    # KA-Nr gaps per WP
    print("KA-Nr Lücken (fortlaufend nummeriert; jede Lücke ist potenziell eine fehlende oder zurückgezogene KA):")
    for wp in sorted(ka_gaps):
        gaps = ka_gaps[wp]
        sample = ", ".join(map(str, gaps[:30]))
        more = f" (+{len(gaps)-30} weitere)" if len(gaps) > 30 else ""
        print(f"  WP{wp}: {len(gaps)} Lücken — {sample}{more}")
    if not ka_gaps:
        print("  keine")
    print()
    print(f"Duplikat-KA-Nr (gleiche wp+ka_nr in mehreren Zeilen): {len(dup_lines)}")
    for line in dup_lines[:10]:
        print(f"  {line}")
    print()
    print(f"vocab novelty — Fraktion: {sum(novel_frak.values())} occurrences across {len(novel_frak)} values")
    for v, n in sorted(novel_frak.items(), key=lambda x: -x[1])[:10]:
        print(f"  {n:4d}× {v!r}")

    # Hard-fail categories
    bad = len(md_missing) + len(md_short) + len(bad_nr) + len(md_kanr_mismatch) + len(dup_lines)
    if bad:
        print(f"\nverify: {bad} broken row(s); exit non-zero", file=sys.stderr)
        return 1
    print("\nverify: clean.")
    return 0


# --- CLI ---------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="landtag", description=__doc__)
    p.add_argument("--xlsx", default=str(INDEX_XLSX), help="path to index.xlsx")
    p.add_argument("--rps", type=float, default=DEFAULT_RPS,
                   help="HTTP requests per second (shared budget). Default polite ceiling for a small public-sector server.")
    p.add_argument("--user-agent", default=USER_AGENT)

    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan-archive", help="enrich existing rows from local PDFs")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--force", action="store_true")
    s.add_argument("--limit", type=int, help="cap number of PDFs to scan (for testing)")
    s.add_argument("--allow-discovery", action="store_true",
                   help="legacy: create new rows from PDFs alone (without crawl). "
                        "Default: skip PDFs whose Drucksache-Nr is not yet in the index.")
    s.set_defaults(func=cmd_scan_archive)

    s = sub.add_parser("crawl", help="enrich (default) or discover via Webflow search")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--from", dest="date_from", help="client-side filter on Anfragedatum")
    s.add_argument("--to", dest="date_to", help="client-side filter on Anfragedatum")
    s.add_argument("--rpp", type=int, default=50)
    s.add_argument("--full", action="store_true",
                   help="DEPRECATED no-op (full sweep is now the only mode; flag kept for back-compat)")
    s.set_defaults(func=cmd_crawl)

    s = sub.add_parser("normalize",
                       help="match Fraktion/Ministerium against Index/*.xlsx and fill canonical columns")
    s.set_defaults(func=cmd_normalize)

    s = sub.add_parser("extract-multi-ministerium",
                       help="parse Antworttext-MD for ALL beteiligte Ministerien (Beteiligte_Ministerien_Kuerzel)")
    s.add_argument("--wahlperiode", type=int)
    s.set_defaults(func=cmd_extract_multi_ministerium)

    s = sub.add_parser("enrich-llm",
                       help="LLM rescue for rows the rule-based parser left flagged in Extract_Flags")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--limit", type=int, help="cap number of LLM calls (for testing/cost-control)")
    s.add_argument("--model", default="gpt-5-mini",
                   help="model name passed to `llm -m` (e.g. gpt-5-mini, gemma4:31b)")
    s.set_defaults(func=cmd_enrich_llm)

    s = sub.add_parser("fetch-text", help="download missing PDFs and extract .md")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--limit", type=int)
    s.add_argument("--force", action="store_true")
    s.add_argument("--workers", type=int, default=1)
    s.set_defaults(func=cmd_fetch_text)

    s = sub.add_parser("resolve",
                       help="targeted re-search for one/several KA-Nrn (für Zweifelsfälle)")
    s.add_argument("--ka", type=int, action="append", required=True,
                   help="KA-Nr to resolve (repeatable, e.g. --ka 144 --ka 3167)")
    s.add_argument("--wahlperiode", type=int, default=18)
    s.add_argument("--counter-search", action="store_true",
                   help="bei 0 KA-Treffern: Gegensuche mit doktyp=GA (Große Anfrage)")
    s.add_argument("--delete-phantom", action="store_true",
                   help="Phantom-Zeilen (stale KA-Nr, oder GA-Match) aus dem Index löschen")
    s.set_defaults(func=cmd_resolve)

    s = sub.add_parser("verify", help="read-only sanity report")
    s.add_argument("--probe-site", action="store_true")
    s.add_argument("--llm-plausibility", action="store_true")
    s.add_argument("--llm-rescue-fields", action="store_true")
    s.set_defaults(func=cmd_verify)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    DATA_DIR.mkdir(exist_ok=True)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
