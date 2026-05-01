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
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

# --- third-party (declared in requirements.txt; not imported here in the
# skeleton so the file parses without them installed) -------------------------
# import httpx
# from bs4 import BeautifulSoup
# import openpyxl
# from filelock import FileLock
# import pdfplumber
# import pypdf

# --- constants ---------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parent
DATA_DIR = REPO_ROOT / "data"
ARCHIV_DIR = REPO_ROOT / "Archiv"
INDEX_XLSX = DATA_DIR / "index.xlsx"
SHEET_NAME = "kleine_anfragen"

USER_AGENT = "wdr-kleineanfrage/0.1 (+contact: jan.eggers@fm.wdr.de)"
DEFAULT_RPS = 1.0
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
    r"Kleine Anfrage\s+(\d+)\s+vom\s+(\d{1,2}\.\s*[A-Za-zÄÖÜäöüß]+\s+\d{4})"
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
_RX_TITLE = re.compile(
    r"Drucksache\s+\d+/\d+\s*\n+\s*((?:[^\n]+\n?)+?)\s*\n\s*\n\s*Vorbemerkung\s+der\s+Kleinen\s+Anfrage",
    re.DOTALL,
)

# "Der Minister ... hat die Kleine Anfrage" — needs first 3 pages to catch all.
# Optional Der/Die prefix (some early WP18 docs omit it). Ministry name may span lines.
_RX_MINISTERIUM = re.compile(
    r"(?:(?:Der|Die)\s+)?(Minister(?:präsident(?:in)?|in)?)\s+(?:der|des|für|für\s+die)\s+(.+?)\s+hat\s+die\s+Kleine\s+Anfrage",
    re.DOTALL,
)


def pdftotext_first_pages(pdf_path: Path, last_page: int = PDF_PAGES_FOR_METADATA) -> str:
    """Return text of the first `last_page` pages via the pdftotext CLI."""
    return subprocess.check_output(
        ["pdftotext", "-layout", "-f", "1", "-l", str(last_page), str(pdf_path), "-"],
        text=True,
        encoding="utf-8",
    )


def parse_page_text(text: str) -> dict:
    """Apply the seven anchored regexes; return a dict with whatever matched.

    Caller is responsible for treating missing keys as `extract_failed` and logging.
    """
    raise NotImplementedError("TODO: run each _RX_* and collect groups into a dict")


def parse_filename_drucksache_nr(pdf_path: Path) -> tuple[int, str]:
    """MMD18-1006.pdf → (18, '18/1006')."""
    m = re.match(r"MMD(\d+)-(\d+)\.pdf$", pdf_path.name)
    if not m:
        raise ValueError(f"unexpected filename: {pdf_path.name}")
    wp = int(m.group(1))
    n = int(m.group(2))
    return wp, f"{wp}/{n}"


# --- archive lookup ----------------------------------------------------------

def archive_lookup(wp: int, n: int) -> Path | None:
    """Walk all bucket folders for MMD<wp>-<n>.pdf. Returns None if absent."""
    raise NotImplementedError("TODO: glob across bucket folders under Archiv/")


def archive_target_path(wp: int, n: int) -> Path:
    """Where a freshly-downloaded MMD<wp>-<n>.pdf should be written."""
    raise NotImplementedError("TODO: compute correct 2000-wide bucket folder")


# --- xlsx I/O (file-locked, atomic) ------------------------------------------

def load_index() -> dict[str, Record]:
    """Read index.xlsx into a dict keyed by Drucksache_Antwort_Nr."""
    raise NotImplementedError("TODO: openpyxl read; return {} if file absent")


def save_index(rows: dict[str, Record]) -> None:
    """Write rows to a temp xlsx, fsync, atomic rename. Holds filelock."""
    raise NotImplementedError("TODO: openpyxl write + os.replace under FileLock")


def upsert(rows: dict[str, Record], rec: Record, *, set_columns: set[str]) -> None:
    """Merge rec into rows[rec.drucksache_antwort_nr], touching only set_columns.

    Sets Hinzugefuegt_am on first insert, always updates Aktualisiert_am.
    """
    raise NotImplementedError("TODO: column-scoped merge")


# --- vocab novelty (no auto-correction, log-only) ----------------------------

def check_fraktion(value: str) -> bool:
    """Hardcoded set membership; the search form has no fraktion <select> to scrape."""
    return value in FRAKTIONEN


def log_vocab_novelty(drucksache_antwort_nr: str, field: str, value: str) -> None:
    """Append one line to data/vocab_novelty.log for unseen Fraktion or Ministerium."""
    raise NotImplementedError("TODO: append-only log")


# --- HTTP client + Webflow handshake -----------------------------------------

def make_client(rps: float, user_agent: str):
    """httpx.Client with rps shared budget and exp backoff on 429/5xx."""
    raise NotImplementedError("TODO: httpx.Client + rate-limit hook + backoff")


def bootstrap_search(client) -> dict:
    """Single GET on SEARCH_BASE; return {webflow_token, flow_execution_param}.

    The two tokens live in the form action URL itself
    (`?webflowToken=<UUID>&webflowexecution<rand>__searchr2020=<value>`),
    not as hidden form inputs. Cookies (JSESSIONID, TS01a5776e) are retained
    by the httpx.Client automatically.
    """
    raise NotImplementedError("TODO: GET + parse form action attribute")


def search_post(client, tokens: dict, *, wp: int | None = None,
                nummer: str | None = None, page_token: str | None = None,
                rpp: int = 50):
    """POST to the search endpoint with SEARCH_FORM_BASE + harvested tokens.

    Enrich mode: pass `nummer="18/1006"` to look up one row.
    Discovery mode: pass `wp=18` and follow `page_token` for pagination.
    """
    raise NotImplementedError("TODO: form fields + paginate via flow token")


def parse_search_hits(html: str) -> list[Record]:
    """Result-page HTML → list of Records (cols 11-14 reliably; 1-10 best-effort)."""
    raise NotImplementedError("TODO: BS4 row extraction")


# --- verbs -------------------------------------------------------------------

def cmd_scan_archive(args: argparse.Namespace) -> int:
    """Walk Archiv/**/MMD<wp>-*.pdf, extract pages 1-3, upsert. No network."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_crawl(args: argparse.Namespace) -> int:
    """Webflow handshake, then enrich (default), discover (--full), or date-filter (--from/--to)."""
    raise NotImplementedError("TODO: see spec §5; date filter is client-side post-fetch")


def cmd_fetch_text(args: argparse.Namespace) -> int:
    """For each row missing PDF/.md: locate-or-download, backfill pages 1-3, full extract."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_verify(args: argparse.Namespace) -> int:
    """Read-only report; optional --llm-plausibility / --llm-rescue-fields."""
    raise NotImplementedError("TODO: see spec §5")


# --- CLI ---------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="landtag", description=__doc__)
    p.add_argument("--xlsx", default=str(INDEX_XLSX), help="path to index.xlsx")
    p.add_argument("--rps", type=float, default=DEFAULT_RPS,
                   help="HTTP requests per second (shared budget). Crawl: keep at 1.0 (search endpoint is robots-Disallow'd). Fetch-text: may raise; PDF dir is robots-allowed.")
    p.add_argument("--user-agent", default=USER_AGENT)

    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan-archive", help="extract pages-1-to-3 metadata from local PDFs")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--force", action="store_true")
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
