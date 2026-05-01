---
name: landtag-nrw-extraction
description: Use when extracting or querying Kleine Anfrage data from Landtag NRW — metadata + answer-text harvesting into Excel + Markdown.
---

# Landtag NRW — Kleine Anfrage extraction

Five idempotent CLI verbs in `landtag.py`, all wrapping a single `data/index.xlsx` and a per-document `.md` cache alongside the source PDFs in `Archiv/`.

## When to use

- The user asks about who filed which Kleine Anfrage, or about ministry/topic distributions, or about answer turnaround → query `data/index.xlsx`.
- The user asks about the *content* of a specific answer → read `Archiv/.../MMD18-N.md`. Fall back to the PDF only if the .md is missing.
- The user asks for fresh data → run the refresh loop below.

## First-time setup

The `Archiv/` folder ships with ~18,000 WP18 PDFs. Build the index from them, no network:

```sh
python landtag.py scan-archive --wahlperiode 18
```

Fills 8 of 12 desired columns from PDF page 1 (Anfrager, Fraktion, both Drucksache numbers, both dates, title, ministry). Status of every row: `pending_enrich`.

## Refresh loop

Run in this order. Each verb is idempotent and re-runnable:

```sh
python landtag.py refresh-vocab                       # rare — once per Wahlperiode
python landtag.py crawl --wahlperiode 18              # enrich Systematik / Schlagworte / links
python landtag.py fetch-text --wahlperiode 18         # download missing PDFs, extract .md
python landtag.py verify                              # sanity report
```

Add `crawl --full` once per Wahlperiode to pick up newly published Drucksachen the local cache doesn't know about yet.

## Read access

```python
import pandas as pd
df = pd.read_excel("data/index.xlsx")
# common filters
df[df.Fraktion == "AfD"]                              # by Fraktion
df[df.Anfrager.str.contains("Wagner")]                # by Abgeordnete name
df[df.Schlagworte.str.contains("Polizei", na=False)] # by topic
```

For the answer text of one row:

```python
md_path = df.loc[df.Drucksache_Antwort_Nr == "18/1006", "Antworttext"].iloc[0]
print(open(md_path, encoding="utf-8").read())
```

CSV/JSONL export when an external tool needs it:

```sh
python -c "import pandas as pd; pd.read_excel('data/index.xlsx').to_csv('out.csv', index=False)"
```

## Don't

- Run two verbs concurrently against the same xlsx. The file lock will block, but it's still simpler not to.
- Hand-edit cols `Antworttext`, `Antworttext_Status`, `Antworttext_Quelle` — they're owned by `fetch-text`.
- Auto-correct values that landed in `data/vocab_mismatch.log`. A divergent Fraktion or Abgeordnete name might be a real new entry, not a typo. Inspect, then update the vocab xlsx if appropriate.
- Bypass the rate limiter. The 1-rps default is shared across all HTTP traffic.

## Common failures

- **`pdftotext: not found`** — install poppler (`brew install poppler` on macOS).
- **`crawl` returns 0 hits** — the Spring Webflow tokens drifted; check the bootstrap GET in `landtag.py`.
- **`fetch-text` reports 404 for many rows** — rows came from `crawl --full` discovery for Drucksachen that aren't yet answered. Status `no_answer_yet` is expected; re-run later.
- **Row count after `scan-archive` lower than PDF count** — failed regex matches; check `data/extract_errors.log`.

## Reference

Full design: `docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md`.
