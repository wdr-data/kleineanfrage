#!/usr/bin/env python3
"""
landtag.py — Kleine Anfrage extraction for Landtag NRW.

Five idempotent verbs:
    scan-archive    Walk Archiv/, pdftotext page 1, upsert metadata. No network.
    crawl           Spring Webflow handshake, paginated POST. Enrich or discover.
    fetch-text      Download missing PDFs, extract full text to .md.
    refresh-vocab   Scrape Fraktion / Abgeordnete / Ministerium dropdowns.
    verify          Read-only sanity report (+ optional --llm-* enrichment).

See docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
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

SEARCH_BASE = "https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html"
PDF_URL_TEMPLATE = (
    "https://www.landtag.nrw.de/portal/WWW/dokumentenarchiv/Dokument/MMD{wp}-{n}.pdf"
)
PDF_URL_ALLOW = re.compile(
    r"^https://www\.landtag\.nrw\.de/portal/WWW/dokumentenarchiv/Dokument/MMD\d+-\d+\.pdf$"
)

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


# --- PDF page-1 extraction ---------------------------------------------------
#
# Anchor patterns derived from sample MMD18-1006.pdf. All anchored to the
# distinctive German phrasing the Antwort-Drucksache header uses verbatim.
# Failures must be loud (write extract_failed + log line); never guess.

_RX_DRUCKSACHE_ANTWORT = re.compile(r"Drucksache\s+(\d+/\d+)")
_RX_KLEINE_ANFRAGE_NR = re.compile(r"Kleine Anfrage\s+(\d+)\s+vom\s+(\d{1,2}\.\s*[A-Za-zÄÖÜäöüß]+\s+\d{4})")
_RX_ANFRAGER_FRAKTION = re.compile(
    r"des\s+(?:Abgeordneten|der\s+Abgeordneten)\s+(.+?)\s+(CDU|SPD|GRÜNE|FDP|AfD|fraktionslos)\b"
)
_RX_DRUCKSACHE_ANFRAGE = re.compile(r"Drucksache\s+(\d+/\d+)\s*$", re.MULTILINE)
_RX_AUSGEGEBEN = re.compile(r"Ausgegeben:\s*(\d{1,2}\.\d{1,2}\.\d{4})")
_RX_MINISTERIUM = re.compile(
    r"Der\s+(Minister(?:präsident)?|Ministerin)\s+(?:der|des|für|für die)\s+([^\.]+?)\s+hat\s+die\s+Kleine\s+Anfrage"
)


def pdftotext_page1(pdf_path: Path) -> str:
    """Return text of page 1 only via the pdftotext CLI."""
    return subprocess.check_output(
        ["pdftotext", "-layout", "-f", "1", "-l", "1", str(pdf_path), "-"],
        text=True,
        encoding="utf-8",
    )


def parse_page1(text: str) -> dict:
    """Apply the eight anchored regexes; return a dict with whatever matched."""
    raise NotImplementedError("TODO: regex extraction; see _RX_* above")


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


# --- vocab -------------------------------------------------------------------

def load_vocab(path: Path) -> set[str]:
    """Read one vocab xlsx into a set for O(1) membership testing."""
    raise NotImplementedError("TODO: openpyxl read")


def save_vocab(path: Path, values: list[str]) -> None:
    """Sorted, deduped, atomic write; preserve Aktualisiert_am for unchanged rows."""
    raise NotImplementedError("TODO: openpyxl upsert")


def scrape_vocab(client) -> dict[str, list[str]]:
    """Parse the three <select> blocks on the search page."""
    raise NotImplementedError("TODO: BeautifulSoup parse")


def levenshtein(a: str, b: str) -> int:
    raise NotImplementedError("TODO: standard DP, ~10 lines")


def log_vocab_mismatch(drucksache_antwort_nr: str, field: str, scraped: str, vocab: set[str]) -> None:
    """Append one line to data/vocab_mismatch.log with a 'nearest' suggestion."""
    raise NotImplementedError("TODO: append-only log")


# --- HTTP client + Webflow handshake -----------------------------------------

def make_client(rps: float, user_agent: str):
    """httpx.Client with 1-rps shared budget and exp backoff on 429/5xx."""
    raise NotImplementedError("TODO: httpx.Client + rate-limit hook + backoff")


def bootstrap_search(client) -> dict:
    """Single GET on SEARCH_BASE; return {webflowToken, flow_execution_param, jsessionid}."""
    raise NotImplementedError("TODO: GET + scrape form + capture cookie")


def search_post(client, tokens: dict, *, wp: int | None, date_from: str | None,
                date_to: str | None, rpp: int = 100, page_token: str | None = None):
    """POST to the search endpoint with dokytyp=KleineAnfrage and the harvested tokens."""
    raise NotImplementedError("TODO: form fields + paginate via flow token")


def parse_search_hits(html: str) -> list[Record]:
    """Result-page HTML → list of Records (cols 11-14 reliably; 1-10 best-effort)."""
    raise NotImplementedError("TODO: BS4 row extraction")


# --- verbs -------------------------------------------------------------------

def cmd_scan_archive(args: argparse.Namespace) -> int:
    """Walk Archiv/**/MMD<wp>-*.pdf, extract page 1, upsert. No network."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_crawl(args: argparse.Namespace) -> int:
    """Webflow handshake, then enrich (default) or discover (--full)."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_fetch_text(args: argparse.Namespace) -> int:
    """For each row missing PDF/.md: locate-or-download, backfill page 1, full extract."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_refresh_vocab(args: argparse.Namespace) -> int:
    """Scrape the three dropdowns, write three vocab xlsx."""
    raise NotImplementedError("TODO: see spec §5")


def cmd_verify(args: argparse.Namespace) -> int:
    """Read-only report; optional --llm-plausibility / --llm-rescue-fields."""
    raise NotImplementedError("TODO: see spec §5")


# --- CLI ---------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="landtag", description=__doc__)
    p.add_argument("--xlsx", default=str(INDEX_XLSX), help="path to index.xlsx")
    p.add_argument("--rps", type=float, default=DEFAULT_RPS, help="HTTP requests per second (shared budget)")
    p.add_argument("--user-agent", default=USER_AGENT)

    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan-archive", help="extract page-1 metadata from local PDFs")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--force", action="store_true")
    s.set_defaults(func=cmd_scan_archive)

    s = sub.add_parser("crawl", help="enrich (default) or discover via Webflow search")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--from", dest="date_from")
    s.add_argument("--to", dest="date_to")
    s.add_argument("--rpp", type=int, default=100)
    s.add_argument("--full", action="store_true", help="full pagination sweep instead of targeted enrichment")
    s.set_defaults(func=cmd_crawl)

    s = sub.add_parser("fetch-text", help="download missing PDFs and extract .md")
    s.add_argument("--wahlperiode", type=int)
    s.add_argument("--limit", type=int)
    s.add_argument("--force", action="store_true")
    s.add_argument("--workers", type=int, default=1)
    s.set_defaults(func=cmd_fetch_text)

    s = sub.add_parser("refresh-vocab", help="scrape Fraktion / Abgeordnete / Ministerium dropdowns")
    s.add_argument("--only", help="comma-separated subset: fraktionen,ministerien,abgeordnete")
    s.set_defaults(func=cmd_refresh_vocab)

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
