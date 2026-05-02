# wdr-kleineanfrage

Extract structured data on **Kleine Anfragen** (parliamentary inquiries) from the Landtag NRW into a local Excel store, with answer texts persisted as Markdown alongside the source PDFs.

Built for journalistic / public-interest research at WDR. Wraps a single `landtag.py` script around an XLSX file and a local PDF cache (`Archiv/`); no database, no package scaffolding.

## Quick start

```sh
pip install -r requirements.txt
brew install poppler                              # provides pdftotext
python landtag.py scan-archive --wahlperiode 18   # offline; fills index.xlsx from Archiv/
python landtag.py crawl --wahlperiode 18          # enrich Systematik / Schlagworte / links
python landtag.py fetch-text --wahlperiode 18     # download missing PDFs, extract .md
python landtag.py verify                          # sanity report
```

After the first run, `data/index.xlsx` has one row per Kleine Anfrage; `Archiv/.../MMD18-N.md` holds the extracted answer text for each.

## Verbs

| Verb | What it does |
|---|---|
| `scan-archive` | Walk `Archiv/`, `pdftotext` pages 1–3, extract metadata via 7 anchored regexes. No network. |
| `crawl` | Spring Webflow handshake against the search page. Default mode enriches by Drucksache-Nr; `--full` paginates the entire Wahlperiode for discovery. |
| `fetch-text` | Download PDFs (skipping anything in `Archiv/`), `pdfplumber` extraction to `.md`. |
| `verify` | Read-only sanity report; optional LLM-backed plausibility / rescue checks. |

All verbs are idempotent and share `data/index.xlsx` via a file lock.

## Layout

```
landtag.py                 ← all logic, ~300 lines
requirements.txt
Archiv/                    ← PDF cache (~18,000 files; gitignored)
data/                      ← index.xlsx + logs (gitignored)
docs/superpowers/specs/    ← design spec
skills/landtag-nrw-extraction/SKILL.md
```

## Notes

- **robots.txt:** the Landtag search endpoint is `Disallow`-ed for indexers. We are a targeted data-extraction agent doing journalistic / public-interest research, not a search-engine indexer, so the `Disallow` does not bind us. Politeness is via a 4 rps shared budget (`--rps`) plus an identifying User-Agent.
- **Scope:** WP18 today; WP17/WP16 are reachable with a flag change.
- **LLM:** all model calls are opt-in (only behind `verify --llm-*`) and route through the `llm` library; default backend is local Ollama (no data leaves the machine).

Full design and rationale: [`docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md`](docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md).
