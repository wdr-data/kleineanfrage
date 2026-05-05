"""One-shot helper: aggregate ministerium_form entries from data/vocab_novelty.log
into data/ministerium_novelty.xlsx, ranked by occurrence count, with the
best-guess Kürzel from Index/ministerien.xlsx (Jaccard / containment over the
operator-curated tokenizer).

Run after `extract-multi-ministerium` to produce a triage list:
    python build_novelty_table.py
"""
import re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

import landtag


def best_match(form: str, canon_min):
    """Return (kuerzel, full, contain_pct, jaccard_pct) for the top canonical
    candidate. Containment = |val ∩ canon| / |val| (1.0 → strict subset, the
    matcher's threshold). Jaccard = |∩| / |∪|. Picks max containment, then
    max jaccard, then fewest canon-tokens (most specific)."""
    val = landtag._ministerium_tokens(form)
    if not val:
        return ("", "", 0.0, 0.0)
    scored = []
    for kz, full, ct in canon_min:
        if not ct:
            continue
        inter = val & ct
        if not inter:
            continue
        contain = len(inter) / len(val)
        jacc = len(inter) / len(val | ct)
        scored.append((contain, jacc, -len(ct), kz, full))
    if not scored:
        return ("", "", 0.0, 0.0)
    scored.sort(reverse=True)
    contain, jacc, _, kz, full = scored[0]
    return (kz, full, contain, jacc)


def main():
    log_path = Path("data/vocab_novelty.log")
    out_path = Path("data/ministerium_novelty.xlsx")
    canon_min = landtag.load_canon_ministerien()

    counts: Counter[str] = Counter()
    drs_by_form: defaultdict[str, list[str]] = defaultdict(list)
    line_rx = re.compile(
        r'^\S+\s\|\s(\S+)\s\|\sministerium_form\s\|\sscraped="(.+)"$'
    )
    for line in log_path.read_text(encoding="utf-8").splitlines():
        m = line_rx.match(line)
        if not m:
            continue
        drs, form = m.group(1), m.group(2)
        counts[form] += 1
        drs_by_form[form].append(drs)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "novelty"
    ws.append([
        "Häufigkeit", "Form (scraped)", "Vorschlag_Kuerzel",
        "Vorschlag_Ministerium", "Containment", "Jaccard",
        "Beispiel_Drucksachen",
    ])

    for form, n in counts.most_common():
        kz, full, contain, jacc = best_match(form, canon_min)
        examples = ", ".join(drs_by_form[form][:3])
        ws.append([
            n, form, kz, full,
            round(contain, 2), round(jacc, 2),
            examples,
        ])

    # Auto-width-ish column hints
    widths = [10, 80, 18, 70, 12, 10, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"wrote {out_path}: {len(counts)} unique forms, "
          f"{sum(counts.values())} total occurrences")


if __name__ == "__main__":
    main()
