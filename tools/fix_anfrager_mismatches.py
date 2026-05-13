#!/usr/bin/env python3
"""One-shot remediation for rows flagged in data/anfrager_novelty.log.

Walks every MISMATCH entry, re-reads the Antwort .md with more aggressive
PDF-quirk handling, and either fixes Anfrager_Alle/Anzahl_Abgeordnete (then
adds 'anfrager_manual' to Extract_Flags so future extract-all-anfrager runs
skip the row) or marks the row 'anfrager_unverified' for analyst attention.

Idempotent: rerunning skips rows that already carry 'anfrager_manual' or
'anfrager_unverified'. Run AFTER `landtag.py extract-all-anfrager`.
"""
from __future__ import annotations
import re
import sys
from pathlib import Path

# Reuse the project's already-loaded helpers.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import landtag

REPO = Path(__file__).resolve().parent.parent
INDEX_XLSX = REPO / "data" / "index.xlsx"
LOG = REPO / "data" / "anfrager_novelty.log"

# Cross-Fraktion zweiter Block ('und Frank Neppe FRAKTIONSLOS' nach dem ersten
# Fraktion-Marker — z.B. 17/971: Vogel/AfD + Neppe/fraktionslos).
RX_ZUSATZ = re.compile(
    r"\bund\s+(.{1,200}?)\s*"
    r"(?:BÜNDNIS\s*90\s*/\s*DIE\s*GRÜNEN?|CDU|SPD|GRÜNEN?|FDP|AfD|[Ff]raktionslos|FRAKTIONSLOS)\b",
    re.DOTALL,
)


def aggressive_whitespace(text: str) -> str:
    """All known pdftotext whitespace-loss patterns at once."""
    text = re.sub(r"\s+", " ", text).strip()
    # lower→upper, punct→upper
    text = re.sub(r"([a-zäöüß.,])([A-ZÄÖÜ])", r"\1 \2", text)
    # 'undXxx' / 'undXxx' (lowercase d→uppercase): already covered above.
    # 'Xxxund' (lowercase x→u of und) — rarer.
    text = re.sub(r"([a-zäöüß])(und)\s", r"\1 \2 ", text)
    return text


def collect_blocks(text: str) -> str:
    """Concatenate all 'der/des Abgeordneten ... FRAKTION' blocks PLUS any
    trailing 'und ... FRAKTION' continuation. Whitespace-aggressive."""
    head = text[:8000]
    out: list[str] = []
    for m in landtag._RX_ANFRAGERBLOCK.finditer(head):
        out.append(m.group(1))
        # Look for 'und ... FRAKTION' immediately after this block's end.
        tail = head[m.end(): m.end() + 400]
        m2 = RX_ZUSATZ.match(tail.lstrip())
        if m2:
            out.append(m2.group(1))
    return aggressive_whitespace(" ; ".join(out))


def match_names(block: str, candidates: list[tuple[str, str, str]]) -> list[tuple[str, str]]:
    """Greedy substring match — returns [(nach, vor)] in order of appearance."""
    if not block:
        return []
    masked = list(block)
    hits: list[tuple[int, str, str]] = []
    for pdf_form, nach, vor in candidates:  # already sorted longest-first
        cur = "".join(masked)
        idx = cur.find(pdf_form)
        if idx < 0:
            continue
        l_ok = idx == 0 or not cur[idx - 1].isalpha()
        r_end = idx + len(pdf_form)
        r_ok = r_end == len(cur) or not cur[r_end].isalpha()
        if not (l_ok and r_ok):
            continue
        hits.append((idx, nach, vor))
        for i in range(idx, r_end):
            masked[i] = "\x00"
    hits.sort(key=lambda t: t[0])
    return [(n, v) for _, n, v in hits]


def all_fraktion_candidates(abg: dict, wp: int) -> list[tuple[str, str, str]]:
    """All Index entries for `wp` regardless of Fraktion (cross-Fraktion case)."""
    out: list[tuple[str, str, str]] = []
    for (wp_k, _frak), lst in abg.items():
        if wp_k == wp:
            out.extend(lst)
    out.sort(key=lambda t: -len(t[0]))
    return out


def parse_mismatches(log: Path) -> dict[str, dict]:
    out: dict[str, dict] = {}
    if not log.exists():
        return out
    with log.open() as f:
        for line in f:
            if "MISMATCH" not in line:
                continue
            m = re.search(
                r"\| (\S+) \| (\S+) \| MISMATCH parsed=(\d+) db_min=(\d+) "
                r'\| db_anfrager="([^"]*)" \| block="([^"]*)"',
                line,
            )
            if not m:
                continue
            ds, frak, parsed, db, dba, blk = m.groups()
            out[ds] = {"frak": frak, "parsed": int(parsed), "db": int(db),
                       "dba": dba, "blk": blk}
    return out


def main() -> int:
    rows = landtag.load_index(INDEX_XLSX)
    abg = landtag.load_abgeordnete_index()
    mismatches = parse_mismatches(LOG)
    print(f"Mismatch entries to process: {len(mismatches)}", file=sys.stderr)

    counters = {"already_handled": 0, "fixed_by_zusatz": 0,
                "fixed_by_aggressive_ws": 0, "fixed_cross_fraktion": 0,
                "fixed_from_db": 0, "no_block_zurueckgezogen": 0,
                "marked_unverified": 0, "still_below_db_min": 0}

    for ds, info in mismatches.items():
        rec = rows.get(ds)
        if rec is None:
            continue
        flags = set(filter(None, (rec.extract_flags or "").split(",")))
        if "anfrager_manual" in flags or "anfrager_unverified" in flags:
            counters["already_handled"] += 1
            continue

        md = REPO / rec.antworttext if rec.antworttext else None
        if not md or not md.exists():
            flags.add("anfrager_unverified")
            counters["marked_unverified"] += 1
            rec.extract_flags = ",".join(sorted(flags))
            continue

        text = md.read_text(encoding="utf-8")
        block = collect_blocks(text)

        # Cross-Fraktion case — block may name people from different Fraktionen.
        # Try first with the row's Fraktion, then with all WP-wide candidates.
        own = abg.get((rec.wp, rec.fraktion), [])
        hits = match_names(block, own)
        # If still under DB-min, try cross-Fraktion to absorb co-signers from
        # OTHER Fraktionen (e.g. Vogel/AfD + Neppe/fraktionslos).
        cross_used = False
        if len(hits) < info["db"]:
            cross = all_fraktion_candidates(abg, rec.wp)
            cross_hits = match_names(block, cross)
            if len(cross_hits) > len(hits):
                hits = cross_hits
                cross_used = True

        # Decision tree
        if rec.antworttext_status == landtag.STATUS_ZURUECKGEZOGEN:
            # Withdrawn — Anfragerblock missing is expected.
            counters["no_block_zurueckgezogen"] += 1
            flags.add("anfrager_manual")
            rec.extract_flags = ",".join(sorted(flags))
            continue

        db_min = landtag.db_anfrager_min_count(rec.anfrager or "")
        if hits and len(hits) >= db_min:
            rec.anfrager_alle = "; ".join(f"{n}, {v}" for n, v in hits)
            rec.anzahl_abgeordnete = len(hits)
            flags.add("anfrager_manual")
            if cross_used:
                flags.add("anfrager_cross_fraktion")
                counters["fixed_cross_fraktion"] += 1
            else:
                counters["fixed_by_aggressive_ws" if info["blk"] != block else "fixed_by_zusatz"] += 1
            rec.extract_flags = ",".join(sorted(flags))
        else:
            # Fallback: trust the DB Anfrager column when it has no 'u.a.'
            # marker. The Landtag DB is authoritative for the first 2 names
            # — if the PDF Anfragerblock matched fewer (typo, layout glitch),
            # the DB row is still our best source. 'u.a.' rows can't fall back
            # because the DB itself doesn't know the full list.
            db_pairs = landtag.parse_db_anfrager(rec.anfrager or "")
            has_ua = " u.a." in (rec.anfrager or "") or (rec.anfrager or "").endswith("u.a.")
            if db_pairs and not has_ua and len(db_pairs) >= db_min:
                rec.anfrager_alle = "; ".join(f"{n}, {v}" for n, v in db_pairs)
                rec.anzahl_abgeordnete = len(db_pairs)
                flags.add("anfrager_manual")
                flags.add("anfrager_from_db")
                counters["fixed_from_db"] += 1
            else:
                if hits:
                    rec.anfrager_alle = "; ".join(f"{n}, {v}" for n, v in hits)
                    rec.anzahl_abgeordnete = len(hits)
                flags.add("anfrager_unverified")
                counters["marked_unverified"] += 1
                counters["still_below_db_min"] += 1
            rec.extract_flags = ",".join(sorted(flags))

    landtag.save_index(rows, INDEX_XLSX)
    for k, v in counters.items():
        print(f"  {k}: {v}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
