#!/usr/bin/env python3
"""Extract 'Datum des Originals' from page 1 of every Anfrage- and Antwort-PDF.

Treibt die Hypothese, dass die neue Referenz für Anfrage-/Antwortdatum
*nicht* die Landtag-DB, sondern der Datums-Footer auf Seite 1 jedes PDFs ist
(Pattern: "Datum des Originals: DD.MM.YYYY/Ausgegeben: DD.MM.YYYY").

Für jede Zeile in data/index.xlsx werden bis zu zwei PDFs geprüft:
  * Antwort-PDF — lokal in Archiv/<WP-Bucket>/MMDxx-N.pdf
  * Anfrage-PDF — meist nicht lokal; wird vom Landtag-CDN nachgeladen,
                  sofern --no-fetch nicht gesetzt ist.

Ausgabe: data/datum_original.xlsx
Spalten: WP, KA, Drucksache_Nr, Dokumentlink, Anfrage_oder_Antwort,
         Datum_Original (ISO yyyy-mm-dd), Datum_Ausgegeben (ISO, optional),
         Status, PDF_Pfad, Aktualisiert_am

Status:
  ok           — Selektor gefunden, Datum geparst.
  no_match     — Seite 1 enthält keinen "Datum des Originals:"-String.
  parse_failed — Selektor gefunden, aber das Datum dahinter ließ sich nicht parsen.
  pdf_missing  — PDF nicht lokal; --no-fetch oder Download lieferte kein PDF.
  fetch_failed — HTTP-/Netzwerkfehler beim Nachladen.

Resume-safe: rows mit Status=ok werden bei Re-Run übersprungen
(--force erzwingt Neu-Extraktion).
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

# Reuse landtag.py infrastructure (paths, archive layout, HTTP client).
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from landtag import (  # noqa: E402
    ARCHIV_DIR,
    DATA_DIR,
    DEFAULT_RPS,
    INDEX_XLSX,
    PDF_URL_ALLOW,
    USER_AGENT,
    archive_lookup,
    archive_target_path,
    make_client,
    pdftotext_first_pages,
)

OUT_XLSX = DATA_DIR / "datum_original.xlsx"

# "Datum des Originals: 21.09.2022/Ausgegeben: 28.09.2022"
# Toleriert variable Whitespaces, optionalen "Ausgegeben"-Teil und beide Trenner (/ oder ;).
_RX_DATUM_ORIGINAL = re.compile(
    r"Datum\s+des\s+Originals\s*:\s*"
    r"(?P<original>\d{1,2}\.\d{1,2}\.\d{2,4})"
    r"(?:\s*[/;]\s*Ausgegeben\s*:\s*(?P<ausgegeben>\d{1,2}\.\d{1,2}\.\d{2,4}))?"
)


def _to_iso(s: str) -> str:
    """'21.09.2022' → '2022-09-21'. Empty on parse fail."""
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", s.strip())
    if not m:
        return ""
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000  # bestätige 2-stellige Jahre als 20xx
    if not (1 <= mo <= 12 and 1 <= d <= 31 and 2000 <= y <= 2099):
        return ""
    return f"{y:04d}-{mo:02d}-{d:02d}"


def _drucksache_to_wp_n(drucksache: str) -> tuple[int, int] | None:
    """'18/1006' → (18, 1006). None bei Fehler / leerer Eingabe."""
    if not isinstance(drucksache, str):
        return None
    m = re.match(r"(\d+)\s*/\s*(\d+)$", drucksache.strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _scan_page_one(pdf_path: Path) -> tuple[str, str, str]:
    """Scanne Seite 1 → (status, datum_original_iso, datum_ausgegeben_iso)."""
    try:
        text = pdftotext_first_pages(pdf_path, last_page=1)
    except Exception as e:
        return f"parse_failed:{type(e).__name__}", "", ""
    m = _RX_DATUM_ORIGINAL.search(text)
    if not m:
        return "no_match", "", ""
    iso_orig = _to_iso(m.group("original"))
    iso_ausg = _to_iso(m.group("ausgegeben") or "")
    if not iso_orig:
        return "parse_failed", "", ""
    return "ok", iso_orig, iso_ausg


def _fetch_pdf(client, url: str, target: Path) -> tuple[bool, str]:
    """Download URL → target. Returns (ok, status_or_error_str)."""
    if not PDF_URL_ALLOW.match(url):
        return False, "url_not_allowed"
    try:
        r = client.get(url)
    except Exception as e:
        return False, f"fetch_failed:{type(e).__name__}"
    if r.status_code != 200:
        return False, f"fetch_failed:http_{r.status_code}"
    content = r.content
    # Sanity: leading %PDF magic
    if not content.startswith(b"%PDF"):
        return False, "fetch_failed:not_pdf"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(content)
    return True, "ok"


def _load_existing(xlsx: Path) -> pd.DataFrame:
    if not xlsx.exists():
        return pd.DataFrame(columns=[
            "WP", "KA", "Drucksache_Nr", "Dokumentlink",
            "Anfrage_oder_Antwort", "Datum_Original", "Datum_Ausgegeben",
            "Status", "PDF_Pfad", "Aktualisiert_am",
        ])
    return pd.read_excel(xlsx)


def _build_tasks(idx: pd.DataFrame) -> list[dict]:
    """Erzeuge eine Task pro (Zeile × Dokumentart). Skippt leere/ungültige Drucksachen."""
    tasks: list[dict] = []
    for _, row in idx.iterrows():
        wp = int(row["WP"])
        ka = row.get("Kleine_Anfrage_Nr")
        ka_int = int(ka) if pd.notna(ka) else None
        for kind, ds_col, link_col in (
            ("Anfrage", "Drucksache_Anfrage_Nr", "Link_Drucksache_Anfrage"),
            ("Antwort", "Drucksache_Antwort_Nr", "Link_Drucksache_Antwort"),
        ):
            ds = row.get(ds_col)
            link = row.get(link_col)
            if not isinstance(ds, str) or not isinstance(link, str):
                continue
            parsed = _drucksache_to_wp_n(ds)
            if parsed is None:
                continue
            pdf_wp, pdf_n = parsed
            if pdf_wp != wp:
                # Drucksache aus anderer WP — selten; trotzdem mitnehmen.
                pass
            tasks.append({
                "WP": pdf_wp,
                "KA": ka_int,
                "Drucksache_Nr": ds,
                "Dokumentlink": link,
                "Anfrage_oder_Antwort": kind,
                "pdf_n": pdf_n,
            })
    return tasks


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--index", default=str(INDEX_XLSX),
                   help=f"Quell-XLSX (default: {INDEX_XLSX})")
    p.add_argument("--out", default=str(OUT_XLSX),
                   help=f"Ziel-XLSX (default: {OUT_XLSX})")
    p.add_argument("--wahlperiode", "--wp", type=int, default=None,
                   help="auf eine WP einschränken")
    p.add_argument("--only", choices=("Anfrage", "Antwort"), default=None,
                   help="nur eine Dokumentart verarbeiten")
    p.add_argument("--no-fetch", action="store_true",
                   help="keine fehlenden PDFs nachladen (nur lokale Archiv-Treffer scannen)")
    p.add_argument("--force", action="store_true",
                   help="Status=ok-Zeilen erneut verarbeiten")
    p.add_argument("--limit", type=int, default=None, help="max. n Tasks (debug)")
    p.add_argument("--rps", type=float, default=DEFAULT_RPS, help="HTTP rps cap")
    p.add_argument("--user-agent", default=USER_AGENT)
    p.add_argument("--flush-every", type=int, default=200,
                   help="alle N Tasks Zwischenstand nach Ziel-XLSX schreiben")
    args = p.parse_args(argv)

    idx_path = Path(args.index)
    out_path = Path(args.out)
    idx = pd.read_excel(idx_path)
    if args.wahlperiode is not None:
        idx = idx[idx["WP"] == args.wahlperiode]

    tasks = _build_tasks(idx)
    if args.only:
        tasks = [t for t in tasks if t["Anfrage_oder_Antwort"] == args.only]

    existing = _load_existing(out_path)
    done_keys: set[tuple] = set()
    if not args.force and not existing.empty:
        for _, r in existing.iterrows():
            if r.get("Status") == "ok":
                done_keys.add((r["Drucksache_Nr"], r["Anfrage_oder_Antwort"]))

    pending = [t for t in tasks if (t["Drucksache_Nr"], t["Anfrage_oder_Antwort"]) not in done_keys]
    if args.limit:
        pending = pending[:args.limit]

    print(f"[extract_datum_original] {len(tasks)} Tasks gesamt, "
          f"{len(pending)} offen, {len(done_keys)} bereits ok.")

    results = existing.to_dict("records") if not existing.empty else []
    # Drop superseded entries (same key) wenn --force oder neuer Re-Versuch ansteht.
    pending_keys = {(t["Drucksache_Nr"], t["Anfrage_oder_Antwort"]) for t in pending}
    if pending_keys:
        results = [r for r in results
                   if (r.get("Drucksache_Nr"), r.get("Anfrage_oder_Antwort")) not in pending_keys]

    client = None
    if not args.no_fetch:
        client = make_client(args.rps, args.user_agent)

    def _flush():
        df = pd.DataFrame(results, columns=[
            "WP", "KA", "Drucksache_Nr", "Dokumentlink",
            "Anfrage_oder_Antwort", "Datum_Original", "Datum_Ausgegeben",
            "Status", "PDF_Pfad", "Aktualisiert_am",
        ])
        out_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(out_path, index=False)

    t0 = time.monotonic()
    try:
        for i, t in enumerate(pending, 1):
            wp, n = t["WP"], t["pdf_n"]
            pdf_path = archive_lookup(wp, n)
            status = ""
            datum_orig = ""
            datum_ausg = ""

            if pdf_path is None:
                if args.no_fetch or client is None:
                    status = "pdf_missing"
                else:
                    target = archive_target_path(wp, n)
                    ok, err = _fetch_pdf(client, t["Dokumentlink"], target)
                    if ok:
                        pdf_path = target
                    else:
                        status = err if err.startswith(("fetch_failed", "url")) else "pdf_missing"

            if pdf_path is not None:
                status, datum_orig, datum_ausg = _scan_page_one(pdf_path)

            results.append({
                "WP": wp,
                "KA": t["KA"],
                "Drucksache_Nr": t["Drucksache_Nr"],
                "Dokumentlink": t["Dokumentlink"],
                "Anfrage_oder_Antwort": t["Anfrage_oder_Antwort"],
                "Datum_Original": datum_orig,
                "Datum_Ausgegeben": datum_ausg,
                "Status": status,
                "PDF_Pfad": str(pdf_path.relative_to(ROOT)) if pdf_path else "",
                "Aktualisiert_am": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            })

            if i % args.flush_every == 0:
                _flush()
                rate = i / max(time.monotonic() - t0, 0.001)
                print(f"  [{i}/{len(pending)}] {rate:.1f} tasks/s — letzte: "
                      f"{t['Drucksache_Nr']} ({t['Anfrage_oder_Antwort']}) → {status}")
    finally:
        _flush()
        if client is not None:
            client.close()

    df_final = pd.DataFrame(results)
    counts = df_final.groupby(["Anfrage_oder_Antwort", "Status"]).size()
    print("\n=== Status-Übersicht ===")
    print(counts.to_string())
    print(f"\nGeschrieben: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
