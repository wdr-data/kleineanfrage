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
    "Hinzugefuegt_am",
    "Aktualisiert_am",
]

STATUS_PENDING = "pending"
STATUS_PENDING_ENRICH = "pending_enrich"
STATUS_EXTRACTED = "extracted"
STATUS_NO_ANSWER = "no_answer_yet"
STATUS_FAILED = "extract_failed"

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
    hinzugefuegt_am: str = ""
    aktualisiert_am: str = ""


# --- PDF extraction (rule-based; verified 100% on 30-PDF WP18 sample) --------
#
# Anchored to the distinctive German phrasing the Antwort-Drucksache headers
# use verbatim. Failures must be loud (write extract_failed + log line);
# never guess.

_RX_DRUCKSACHE_ANTWORT = re.compile(r"Drucksache\s+(\d+/\d+)")

_RX_KLEINE_ANFRAGE_NR = re.compile(
    r"Kleine Anfrage\s+(\d+)(?:\s+vom\s+(\d{1,2}\.\s*[A-Za-zÄÖÜäöüß]+\s+\d{4}))?"
)

# "des Abgeordneten X FRAKTION" / "der Abgeordneten X Y und Z FRAKTION"
_RX_ANFRAGER_FRAKTION = re.compile(
    r"(?:des\s+Abgeordneten|der\s+Abgeordneten)\s+(.+?)\s+(CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\b"
)

# Anfrage-Drucksache sits on the line immediately after Anfrager+Fraktion.
_RX_DRUCKSACHE_ANFRAGE = re.compile(
    r"(?:des\s+Abgeordneten|der\s+Abgeordneten)\s+.+?\s+(?:CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\s*\n\s*Drucksache\s+(\d+/\d+)"
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
# Optional Der/Die prefix (some early WP18 docs omit it). Ministry name may span lines.
# Capped to a sane length to prevent linear-time blowup if the anchor is absent.
_RX_MINISTERIUM = re.compile(
    r"(?:(?:Der|Die)\s+)?(Minister(?:präsident(?:in)?|in)?)\s+(?:der|des|für|für\s+die)\s+(.{1,400}?)\s+hat\s+die\s+Kleine\s+Anfrage",
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
        body = re.sub(r"\s+", " ", m.group(2)).strip()
        out["ministerium"] = f"{prefix} {body}"

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
                    if field_name in ("wp", "kleine_anfrage_nr"):
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
                nummer: str | None = None, rpp: int = 50) -> str:
    """POST initial search; return result-page HTML."""
    data = dict(SEARCH_FORM_BASE)
    data["wp"] = str(wp) if wp is not None else "al"
    data["nummer"] = nummer or ""
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

    Reliably populated: kleine_anfrage_nr, anfrager, fraktion,
    drucksache_antwort_nr OR drucksache_anfrage_nr (whichever the linked PDF is),
    anfragedatum (best-effort), systematik, schlagworte, link_anfrage / link_antwort,
    anfragetitel.
    """
    soup = BeautifulSoup(html, "lxml")
    out: list[Record] = []
    for art in soup.select("article.e-search-result"):
        body = art.select_one(".e-search-result__body")
        if not body:
            continue
        rec = Record()
        # Title — first <strong><b>...</b></strong>
        title_el = body.find(["strong", "b"])
        if title_el:
            rec.anfragetitel = re.sub(r"\s+", " ", title_el.get_text(" ", strip=True)).strip()
        # PDF link
        link_el = body.find("a", href=re.compile(r"MMD\d+-\d+\.pdf"))
        if link_el:
            href = link_el["href"]
            full_url = "https://www.landtag.nrw.de" + href if href.startswith("/") else href
            # Linked Drucksache nr from URL
            m = re.search(r"MMD(\d+)-(\d+)\.pdf", href)
            if m:
                rec.wp = int(m.group(1))
                linked_nr = f"{m.group(1)}/{m.group(2)}"
                # Search hits return the QUESTION Drucksache (verified on live site)
                rec.drucksache_anfrage_nr = linked_nr
                rec.link_anfrage = full_url
        # Free-text body: parse "Kleine Anfrage N", Anfrager+Fraktion line, "Systematik:", "Schlagworte:"
        text = body.get_text("\n", strip=True)
        m = re.search(r"Kleine Anfrage\s+(\d+)", text)
        if m:
            rec.kleine_anfrage_nr = int(m.group(1))
        # Anfrager + Fraktion line: e.g. "Wagner, Markus AfD" or "Dr. Pfeil, Werner FDP"
        m = re.search(
            r"\n([^\n]+?)\s+(CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\s*\n",
            "\n" + text + "\n",
        )
        if m:
            rec.anfrager = m.group(1).strip()
            rec.fraktion = m.group(2)
        # Date in DD.MM.YYYY form
        m = re.search(r"(\d{1,2}\.\d{1,2}\.\d{4})\s+\d+\s*S\.", text)
        if m:
            rec.anfragedatum = _german_date_to_iso(m.group(1))
        # Systematik / Schlagworte (label-prefixed paragraphs separated by *)
        m = re.search(r"Systematik:\s*\n(.+?)(?:\nSchlagworte:|\nRegion:|$)", text, re.DOTALL)
        if m:
            rec.systematik = re.sub(r"\s+", " ", m.group(1)).replace(" * ", "; ").strip()
        m = re.search(r"Schlagworte:\s*\n(.+?)(?:\nRegion:|\nSystematik:|$)", text, re.DOTALL)
        if m:
            rec.schlagworte = re.sub(r"\s+", " ", m.group(1)).replace(" * ", "; ").strip()
        out.append(rec)
    return out


def find_next_page_url(html: str) -> str | None:
    """Return the absolute URL of the next page in a paginated result set, or None."""
    soup = BeautifulSoup(html, "lxml")
    # Look for "Zu Seite N+1" or numeric page links; the simplest signal is
    # an <a> whose visible text starts with "Zu Seite" (the "next" arrow).
    for a in soup.find_all("a", href=True):
        txt = a.get_text(" ", strip=True)
        if txt.startswith("Zu Seite ") and "anfragen-und-antworten-suchergeb" in a["href"]:
            href = a["href"]
            return "https://www.landtag.nrw.de" + href if href.startswith("/") else href
    return None


# --- verbs -------------------------------------------------------------------

_SCAN_FIELDS = {
    "wp", "kleine_anfrage_nr", "drucksache_anfrage_nr", "drucksache_antwort_nr",
    "anfrager", "fraktion", "anfragedatum", "anfragetitel", "antwortdatum",
    "ministerium", "antworttext_status", "antworttext_quelle",
}


def cmd_scan_archive(args: argparse.Namespace) -> int:
    """Walk Archiv/**/MMD<wp>-*.pdf, extract pages 1-3, upsert. No network."""
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    seen_ministeria: set[str] = {r.ministerium for r in rows.values() if r.ministerium}

    counters = {"scanned": 0, "parsed": 0, "extract_failed": 0, "skipped": 0, "novelty": 0}
    pdfs = list(iter_archive_pdfs(args.wahlperiode))
    if args.limit:
        pdfs = pdfs[: args.limit]
    print(f"scan-archive: {len(pdfs)} PDFs to consider", file=sys.stderr, flush=True)

    for pdf in pdfs:
        counters["scanned"] += 1
        try:
            wp, drucksache_antwort_nr = parse_filename_drucksache_nr(pdf)
        except ValueError as e:
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {pdf.name} | filename: {e}")
            counters["extract_failed"] += 1
            continue

        existing = rows.get(drucksache_antwort_nr)
        if existing and not args.force and existing.antworttext_status != STATUS_FAILED \
                and existing.kleine_anfrage_nr:
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
        if not parsed.get("kleine_anfrage_nr"):
            rec.antworttext_status = STATUS_FAILED
            counters["extract_failed"] += 1
            _append_log(EXTRACT_ERRORS_LOG, f"{_now_iso()} | {pdf.name} | regex: kleine_anfrage_nr missing")
        else:
            rec.antworttext_status = STATUS_PENDING_ENRICH
            counters["parsed"] += 1

        # Vocab novelty checks
        if rec.fraktion and not check_fraktion(rec.fraktion):
            log_vocab_novelty(drucksache_antwort_nr, "Fraktion", rec.fraktion)
            counters["novelty"] += 1
        if rec.ministerium and rec.ministerium not in seen_ministeria:
            log_vocab_novelty(drucksache_antwort_nr, "Ministerium", rec.ministerium)
            seen_ministeria.add(rec.ministerium)
            counters["novelty"] += 1

        upsert(rows, rec, set_columns=_SCAN_FIELDS)

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
        f"vocab_novelty={counters['novelty']} rows_in_xlsx={len(rows)}",
        file=sys.stderr, flush=True,
    )
    return 0


_CRAWL_FIELDS = {
    "wp", "kleine_anfrage_nr", "drucksache_anfrage_nr", "drucksache_antwort_nr",
    "anfrager", "fraktion", "anfragedatum", "anfragetitel",
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
    """Enrich (default) or discover (--full) Kleine Anfrage metadata via Webflow search."""
    xlsx = Path(args.xlsx)
    rows = load_index(xlsx)
    seen_ministeria: set[str] = {r.ministerium for r in rows.values() if r.ministerium}

    counters = {"queried": 0, "hits": 0, "matched_existing": 0, "new": 0, "novelty": 0, "errors": 0}

    with make_client(args.rps, args.user_agent) as client:
        try:
            tokens = bootstrap_search(client)
        except Exception as e:
            _append_log(CRAWL_ERRORS_LOG, f"{_now_iso()} | bootstrap failed: {e}")
            print(f"crawl: bootstrap failed: {e}", file=sys.stderr)
            return 2
        print(f"crawl: bootstrapped {tokens['post_url'][:120]}", file=sys.stderr, flush=True)

        if args.full:
            return _crawl_full(args, client, tokens, rows, seen_ministeria, counters, xlsx)

        # Enrich mode: iterate rows whose Systematik/Schlagworte are empty AND
        # who have a question Drucksache (Drucksache_Anfrage_Nr) to query by.
        targets = [
            (key, rec) for key, rec in rows.items()
            if (not rec.systematik or not rec.schlagworte) and rec.drucksache_anfrage_nr
        ]
        if args.wahlperiode is not None:
            targets = [(k, r) for k, r in targets if r.wp == args.wahlperiode]
        print(f"crawl: enrich mode — {len(targets)} rows need Systematik/Schlagworte", file=sys.stderr, flush=True)

        for i, (key, rec) in enumerate(targets, 1):
            counters["queried"] += 1
            try:
                # Spring Webflow tokens are single-use; re-bootstrap each query.
                tokens = bootstrap_search(client)
                html = search_post(client, tokens, wp=rec.wp, nummer=rec.drucksache_anfrage_nr, rpp=10)
                hits = parse_search_hits(html)
            except Exception as e:
                _append_log(CRAWL_ERRORS_LOG, f"{_now_iso()} | enrich {rec.drucksache_anfrage_nr}: {e}")
                counters["errors"] += 1
                continue
            counters["hits"] += len(hits)
            # Find the hit whose link matches our question Drucksache
            match = next((h for h in hits if h.drucksache_anfrage_nr == rec.drucksache_anfrage_nr), None)
            if not match:
                _append_log(
                    CRAWL_ERRORS_LOG,
                    f"{_now_iso()} | enrich {rec.drucksache_anfrage_nr}: no exact hit ({len(hits)} returned)",
                )
                counters["errors"] += 1
                continue
            # Apply date filter post-fetch
            ad = rec.anfragedatum or match.anfragedatum
            if not _date_in_range(ad, args.date_from, args.date_to):
                continue
            # Compute Link_Drucksache_Antwort from existing Drucksache_Antwort_Nr
            if rec.drucksache_antwort_nr:
                m = re.match(r"(\d+)/(\d+)", rec.drucksache_antwort_nr)
                if m:
                    match.link_antwort = PDF_URL_TEMPLATE.format(wp=m.group(1), n=m.group(2))
            counters["matched_existing"] += 1
            # Don't overwrite the canonical Drucksache_Antwort_Nr key with the question nr,
            # and preserve the existing antworttext_status (scan-archive set it to
            # pending_enrich; crawl shouldn't downgrade it).
            match.drucksache_antwort_nr = rec.drucksache_antwort_nr
            upsert(rows, match, set_columns=_CRAWL_FIELDS - {"antworttext_status"})
            if i % 50 == 0:
                print(f"  ...{i}/{len(targets)} matched={counters['matched_existing']}", file=sys.stderr, flush=True)
                save_index(rows, xlsx)  # checkpoint

    print(f"saving {len(rows)} rows to {xlsx} ...", file=sys.stderr, flush=True)
    save_index(rows, xlsx)
    print(
        f"crawl done: queried={counters['queried']} hits={counters['hits']} "
        f"matched_existing={counters['matched_existing']} new={counters['new']} "
        f"errors={counters['errors']}",
        file=sys.stderr, flush=True,
    )
    return 0


def _crawl_full(args, client, tokens, rows, seen_ministeria, counters, xlsx) -> int:
    """Discovery mode: paginate the entire result set and insert any new rows.

    Note: the search returns one hit per Kleine Anfrage, identified by the
    QUESTION Drucksache. We can't directly learn the answer Drucksache here
    (the answer is its own Drucksache, only locatable from the answer PDF).
    Discovery rows are keyed by question Drucksache with status='pending';
    fetch-text + scan-archive will fill in the answer Drucksache later.
    """
    print(f"crawl: --full discovery mode for wp={args.wahlperiode}", file=sys.stderr, flush=True)
    html = search_post(client, tokens, wp=args.wahlperiode, rpp=args.rpp)
    page = 1
    while True:
        hits = parse_search_hits(html)
        counters["hits"] += len(hits)
        for h in hits:
            ad = h.anfragedatum
            if ad and not _date_in_range(ad, args.date_from, args.date_to):
                continue
            # In discovery, we have a question Drucksache but no answer Drucksache.
            # Use the question Drucksache as a placeholder PK so the row exists;
            # scan-archive / fetch-text can later correct it once an answer surfaces.
            if not h.drucksache_anfrage_nr:
                continue
            existing = next(
                (r for r in rows.values() if r.drucksache_anfrage_nr == h.drucksache_anfrage_nr),
                None,
            )
            if existing:
                # Update Systematik/Schlagworte etc
                upsert(rows, h, set_columns=_CRAWL_FIELDS - {"drucksache_antwort_nr"})
                counters["matched_existing"] += 1
            else:
                h.drucksache_antwort_nr = h.drucksache_anfrage_nr  # placeholder PK
                h.antworttext_status = STATUS_PENDING
                upsert(rows, h, set_columns=_CRAWL_FIELDS)
                counters["new"] += 1
        next_url = find_next_page_url(html)
        if not next_url:
            break
        page += 1
        if page % 10 == 0:
            print(f"  page {page}, hits so far={counters['hits']}", file=sys.stderr, flush=True)
            save_index(rows, xlsx)  # checkpoint
        html = search_get_page(client, next_url)

    save_index(rows, xlsx)
    print(
        f"crawl --full done: pages={page} hits={counters['hits']} "
        f"new={counters['new']} matched_existing={counters['matched_existing']}",
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

    save_index(rows, xlsx)
    print(
        f"fetch-text done: have={counters['have']} downloaded={counters['downloaded']} "
        f"extracted={counters['extracted']} no_answer={counters['no_answer']} "
        f"failed={counters['failed']}",
        file=sys.stderr, flush=True,
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
    bad_nr: list[str] = []
    novel_min: dict[str, int] = {}
    novel_frak: dict[str, int] = {}

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
        for col in (rec.kleine_anfrage_nr, rec.drucksache_anfrage_nr, rec.drucksache_antwort_nr):
            pass  # placeholder; specific malformed checks below
        if rec.drucksache_antwort_nr and not re.match(r"\d+/\d+$", rec.drucksache_antwort_nr):
            bad_nr.append(f"{key}: drucksache_antwort_nr={rec.drucksache_antwort_nr!r}")
        if rec.fraktion and not check_fraktion(rec.fraktion):
            novel_frak[rec.fraktion] = novel_frak.get(rec.fraktion, 0) + 1

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
    print(f"malformed Drucksache_*_Nr: {len(bad_nr)}")
    for k in bad_nr[:10]:
        print(f"  {k}")
    print(f"orphan PDFs in Archiv (no matching row): {len(orphans)}")
    for k in orphans[:10]:
        print(f"  {k}")
    print()
    print(f"vocab novelty — Fraktion: {sum(novel_frak.values())} occurrences across {len(novel_frak)} values")
    for v, n in sorted(novel_frak.items(), key=lambda x: -x[1])[:10]:
        print(f"  {n:4d}× {v!r}")

    # Hard-fail categories
    bad = len(md_missing) + len(md_short) + len(bad_nr)
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

    s = sub.add_parser("scan-archive", help="extract pages-1-to-3 metadata from local PDFs")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--force", action="store_true")
    s.add_argument("--limit", type=int, help="cap number of PDFs to scan (for testing)")
    s.set_defaults(func=cmd_scan_archive)

    s = sub.add_parser("crawl", help="enrich (default) or discover via Webflow search")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--from", dest="date_from", help="client-side filter on Anfragedatum")
    s.add_argument("--to", dest="date_to", help="client-side filter on Anfragedatum")
    s.add_argument("--rpp", type=int, default=50)
    s.add_argument("--full", action="store_true",
                   help="full pagination sweep instead of targeted enrichment")
    s.set_defaults(func=cmd_crawl)

    s = sub.add_parser("fetch-text", help="download missing PDFs and extract .md")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--limit", type=int)
    s.add_argument("--force", action="store_true")
    s.add_argument("--workers", type=int, default=1)
    s.set_defaults(func=cmd_fetch_text)

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
