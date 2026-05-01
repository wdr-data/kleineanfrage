# Landtag NRW — Kleine Anfrage extraction: design

**Status:** draft for review
**Date:** 2026-05-01
**Owner:** Jan Eggers (jan.eggers@fm.wdr.de)

## 1. Goal and scope

Extract structured data on **Kleine Anfragen** from the public Landtag NRW parliamentary-document search at `https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html` into a local Excel store, with the full answer text persisted as Markdown files alongside the source PDFs.

In scope: document type **Kleine Anfrage** only; operating modes by `--wahlperiode N` and/or `--from / --to`, with idempotent incremental backfill on subsequent runs; LLM use limited to opt-in plausibility/rescue checks routed through the [`llm`](https://pypi.org/project/llm/) library (default backend: local Ollama).

Out of scope: Große and Mündliche Anfrage; GUI or database backend; cross-Wahlperiode merging beyond what the upsert key gives.

The design honours CLAUDE.md's "as simple and low-level as possible": one `landtag.py` script (no `src/` package, no `pyproject.toml`), thin wrapper around an XLSX file plus the existing local PDF cache.

## 2. Use cases

1. **XLSX-based data evaluation** — quantitative analysis across the dataset (most active Abgeordnete, answer-turnaround distributions, topic/ministry breakdowns).
2. **Thematic queries in an agent chat session** — a Claude Code agent uses the skill mid-conversation to refresh metadata, fetch matching answer texts, and quote from the resulting `.md` files.

The agent-driven use case is why the CLI splits into narrow, idempotent verbs.

## 3. Architecture

### 3.1 Flow

```
Archiv/**/MMD18-*.pdf
        │
        ▼
landtag scan-archive ──► index.xlsx     (cols: WP, Kleine_Anfrage_Nr,
   pdftotext page 1                       Drucksache_*_Nr, Anfrager, Fraktion,
   regex extraction                       Anfragedatum, Anfragetitel,
   no network                             Antwortdatum, Ministerium,
                                          status='pending_enrich')
        │
        ▼
landtag crawl ─────────► index.xlsx     (Webflow handshake → search hits;
   bootstrap GET                          fills Systematik, Schlagworte,
   paginated POST                         Link_Drucksache_*; discovers any
   parse search hits                      Drucksachen Archiv/ doesn't have)
        │
        ▼
landtag fetch-text ────► Archiv/.../MMD18-N.pdf  (download missing)
   single rps budget     Archiv/.../MMD18-N.md   (pdfplumber → text)
   targets rows whose    index.xlsx              (cols 15-17 + Aktualisiert_am)
   PDF or .md is missing
        │
        ▼
landtag verify         (read-only report; optional --llm-* enrichment)
```

### 3.2 Approach choices

- **PDF-page-1 extraction is rule-based.** `pdftotext -layout -f 1 -l 1` plus 8 anchored regexes. Failures → row marked `extract_failed`, never silent.
- **Spring Webflow handshake for crawl.** httpx + BeautifulSoup. The search form requires two session-scoped tokens scraped from the search-page HTML on a single bootstrap GET: `webflowToken` (a UUID) and a `webflowexecution…__searchr2020=eXsY` flow-execution parameter. The `JSESSIONID` cookie is held by an `httpx.Client` for the run. Pagination follows the next-page flow token returned in each result page. No browser dependency.
- **Targeted crawl, not bulk by default.** After `scan-archive` populates ~18,000 rows, `crawl` need only enrich the ones with `Systematik`/`Schlagworte` empty and discover new `Drucksache_Antwort_Nr` values not yet in the store. A `--full` flag re-paginates the entire result set when needed.

### 3.3 File layout

```
landtag.py                 ← all verbs, helpers inline (no src/ package)
test_landtag.py            ← only when a real bug demands it (initially empty)
data/
├── index.xlsx
├── fraktionen.xlsx
├── ministerien.xlsx
├── abgeordnete.xlsx
├── vocab_mismatch.log
├── crawl_errors.log
└── extract_errors.log
Archiv/
└── Antworten_Anfragen 18_WP 1-18250/
    ├── 1-2000/
    │   ├── MMD18-100.pdf      (existing)
    │   └── MMD18-100.md       (created by fetch-text or scan-archive)
    └── …
skills/landtag-nrw-extraction/SKILL.md
```

Buckets are 2000-wide partitions of the Drucksache number. `archive_lookup(N)` walks all buckets to find a PDF (handles human-moved files); newly downloaded PDFs go to the correct bucket, creating it if needed.

## 4. Data model

### 4.1 `data/index.xlsx`, sheet `kleine_anfragen`

| # | Column | Source | Notes |
|---|---|---|---|
| 1 | `WP` | Archiv path / search filter | Integer |
| 2 | `Kleine_Anfrage_Nr` | PDF p.1 (`Kleine Anfrage 366`) | **Sequential nr.** Empty if extraction failed. |
| 3 | `Drucksache_Anfrage_Nr` | PDF p.1 (`Drucksache 18/675`) | Question Drucksache |
| 4 | `Drucksache_Antwort_Nr` | filename `MMD18-1006.pdf` / search | **Unique key for upsert.** |
| 5 | `Anfrager` | PDF p.1 (`des Abgeordneten …`) | Multiple co-signers `; ` joined |
| 6 | `Fraktion` | PDF p.1 (trailing token after Anfrager) | Validated against vocab |
| 7 | `Anfragedatum` | PDF p.1 (`vom 24. August 2022`) | ISO `YYYY-MM-DD` |
| 8 | `Anfragetitel` | PDF p.1 (bold heading) | Useful for analytics scan |
| 9 | `Antwortdatum` | PDF p.1 header date | ISO |
| 10 | `Ministerium` | PDF p.1 (`Der Minister …`) | Mapped to Kürzel via vocab table |
| 11 | `Systematik` | search hit | `; ` joined. Empty until enriched. |
| 12 | `Schlagworte` | search hit | `; ` joined. Empty until enriched. |
| 13 | `Link_Drucksache_Anfrage` | search hit | Direct PDF URL |
| 14 | `Link_Drucksache_Antwort` | search hit | Direct PDF URL |
| 15 | `Antworttext` | fetch-text | Relative path to `.md` |
| 16 | `Antworttext_Status` | fetch-text / scan-archive | `pending` / `pending_enrich` / `extracted` / `no_answer_yet` / `extract_failed` |
| 17 | `Antworttext_Quelle` | fetch-text | `pdf_local` / `downloaded` |
| 18 | `Hinzugefuegt_am` | system | First-seen ISO timestamp |
| 19 | `Aktualisiert_am` | system | Last-write ISO timestamp |

### 4.2 Upsert rules

- **Primary key:** `Drucksache_Antwort_Nr` (col 4). Always present once a row exists; matches the cache filename `MMD<wp>-<n>.pdf`.
- `scan-archive` writes cols 1–10, 16, 17 (`pdf_local`), 18, 19. Sets status `pending_enrich`.
- `crawl` writes cols 11–14 plus 19. Inserts new rows for unseen Drucksachen with cols 1–10 to be filled when `fetch-text` later pulls and PDF-extracts.
- `fetch-text` writes cols 15–17 + 19. Triggers `scan-archive`-style page-1 extraction on freshly downloaded PDFs to backfill cols 1–10.
- `Hinzugefuegt_am` set once, never updated.

### 4.3 Vocab tables

Three companion xlsx files validate the controlled-vocabulary fields (`Fraktion`, `Anfrager`, `Ministerium`) by comparing scraped values against canonical lists harvested from the search form's own dropdowns. Mismatches are logged, never auto-corrected — a divergent value might be a real new MP, not a typo.

| File | Sheet | Columns |
|---|---|---|
| `data/fraktionen.xlsx` | `fraktionen` | `Wert`, `Aktualisiert_am` |
| `data/ministerien.xlsx` | `ministerien` | `Wert`, `Aktualisiert_am` |
| `data/abgeordnete.xlsx` | `abgeordnete` | `Wert`, `Aktualisiert_am` |

`Wert` is the option's visible text exactly as it appears in the dropdown — assumed canonical.

`data/vocab_mismatch.log` is append-only, one line per mismatch:
`<ISO-timestamp> | <Drucksache_Antwort_Nr> | <Feld> | scraped="…" | nearest="…" (distance=N)`

The "nearest" suggestion is computed by a Levenshtein helper inside `landtag.py`. Mismatches never block writes to `index.xlsx`.

## 5. CLI surface

Five verbs. All idempotent. Shared flags: `--rps` (default 1.0), `--user-agent` (default `wdr-kleineanfrage/0.1 (+contact: jan.eggers@fm.wdr.de)`), `--xlsx PATH` (default `data/index.xlsx`).

### `landtag scan-archive`

Walk `Archiv/**/MMD<wp>-*.pdf`, run page-1 extraction, upsert rows. **No network.**

```
landtag scan-archive                         # everything in Archiv/
landtag scan-archive --wahlperiode 18
landtag scan-archive --force                 # re-extract even if row exists
```

Per-PDF: `pdftotext -layout -f 1 -l 1` → 8 anchored regexes (one each for cols 2,3,5,6,7,8,9,10; col 4 from filename; col 1 from path). Anything that doesn't match → row gets `Antworttext_Status='extract_failed'` and one line in `extract_errors.log`. Never raises.

### `landtag crawl`

```
landtag crawl --wahlperiode 18
landtag crawl --from 2024-01-01 --to 2024-12-31
landtag crawl --wahlperiode 18 --rpp 100
landtag crawl --wahlperiode 18 --full
```

Spring Webflow handshake (see §3.2). Two modes:

- **Enrich mode (default):** for each row in `index.xlsx` with `Systematik`/`Schlagworte` empty, query the search by `Drucksache_Antwort_Nr` and parse the hit. Saves a full pagination sweep when `scan-archive` already populated rows.
- **Discovery mode (`--full`):** paginate the entire result set for the selector; insert any `Drucksache_Antwort_Nr` not already in the store with `Antworttext_Status='pending'`.

Validates `Fraktion`, `Anfrager`, `Ministerium` against vocab tables; mismatches → `vocab_mismatch.log` (never auto-corrected, never block the write). If a vocab file is missing or empty, one warning to stderr and validation is skipped for that field.

Step-by-step (discovery mode):

1. GET search page; parse `webflowToken` + flow-execution parameter from the form action; pick up `JSESSIONID`.
2. POST search with form fields `dokytyp=KleineAnfrage`, `wp=N`, optional date range, `rpp`, `_eventId_startanfragesearch=Suchen`, plus harvested tokens.
3. Parse records, upsert into xlsx (file-locked, atomic).
4. Follow next-page flow token until exhausted, with politeness (1 req/s baseline, exp backoff 1/2/4/8 s on 429/5xx, then fail).

### `landtag fetch-text`

```
landtag fetch-text                           # all missing PDFs and/or .md files
landtag fetch-text --wahlperiode 18 --limit 200
landtag fetch-text --force                   # re-extract even if .md exists
landtag fetch-text --workers 4               # parallel I/O, shared rps budget
```

For each candidate row:

1. **Locate PDF.** Walk buckets for `MMD<wp>-<n>.pdf`. If present → `Antworttext_Quelle='pdf_local'`. Else GET `https://www.landtag.nrw.de/portal/WWW/dokumentenarchiv/Dokument/MMD<wp>-<n>.pdf` → write to correct bucket → `'downloaded'`. URL allow-list check (host + path regex `^/portal/WWW/dokumentenarchiv/Dokument/MMD\d+-\d+\.pdf$`) before any GET. 404 → `Antworttext_Status='no_answer_yet'`, skip.
2. **Backfill page-1 metadata** if cols 2–10 are empty (i.e. row was created by `crawl --full` discovery). Same regexes as `scan-archive`.
3. **Extract full text.** `pdfplumber` page-by-page joined by `\n\n`; `pypdf` fallback on parse error; both fail → `extract_failed`, log, no `.md` written.
4. **Write `.md`** atomic temp+rename, update cols 15–17 + 19.

`--workers N` parallelises only network + extraction; xlsx writes are batched. The 1-rps budget is shared across all workers.

### `landtag refresh-vocab`

```
landtag refresh-vocab
landtag refresh-vocab --only fraktionen,ministerien
```

Run the bootstrap GET to obtain a primed `httpx.Client`, parse the three `<select>` blocks (Fraktion, Abgeordneter, Ministerium) by `name=` attribute, write each into its vocab xlsx. Sorted, deduped, atomic write with file-lock; preserves prior `Aktualisiert_am` for unchanged rows; never deletes rows that vanished from the dropdown (kept for historical lookup; operator may prune).

Console output: `Refreshed: F fraktionen (+a/-b), M ministerien (+a/-b), A abgeordnete (+a/-b)`. Exits non-zero only on HTTP/parse failure.

Not auto-run by `crawl`. Operator workflow: `refresh-vocab` once per Wahlperiode (or whenever drift is suspected), then `crawl`, then `verify`.

### `landtag verify`

Read-only sanity checks. No network unless `--probe-site`. Reports:

- Row count by Wahlperiode and status.
- Rows with `Drucksache_Antwort_Nr` set but `.md` missing.
- Rows marked `extracted` whose `.md` is suspiciously short relative to the PDF page count.
- Rows where `Kleine_Anfrage_Nr` or `Drucksache_*_Nr` look malformed.
- Orphan files in `Archiv/` (PDFs/MDs without a matching xlsx row).
- Vocab mismatch summary: total per field + top 10 unrecognised values per field.

Exits non-zero if any "broken" category is non-empty. The vocab mismatch summary is informational and never flips the exit code.

LLM-assisted modes (opt-in, off by default):

- `--llm-plausibility` — for short-text rows, ask the local model whether the extraction looks plausible. Output to `data/verify_llm.log`; never written back to xlsx.
- `--llm-rescue-fields` — for rows with empty cols 9/10/11/12 and an existing `.md`, ask the model to extract those fields from the markdown. Writes back only when confidence is high; logs everything.

CSV / JSONL export is intentionally not a separate verb — one line in the skill covers it: `python -c "import pandas as pd; pd.read_excel('data/index.xlsx').to_csv('out.csv', index=False)"`.

## 6. LLM access

All model calls go through the [`llm`](https://pypi.org/project/llm/) PyPI library. The default backend is **local Ollama** via the `llm-ollama` plugin; no data leaves the machine.

- Default model: `llama3.1:8b` (override via `LANDTAG_LLM_MODEL`).
- Switching to a cloud provider is a one-line config change because `llm` already normalises providers.
- LLM calls only happen behind explicit `verify --llm-*` flags. `crawl`, `fetch-text`, `scan-archive`, `refresh-vocab` never call any model.

## 7. Politeness, errors, recovery

- Default rate: 1 request per second across all HTTP traffic (search and PDF), as a single shared budget — `fetch-text --workers N` parallelises CPU-bound extraction but does not raise the network budget. Override with `--rps`.
- User-Agent: `wdr-kleineanfrage/0.1 (+contact: jan.eggers@fm.wdr.de)`.
- Backoff on 429/5xx: exponential at 1, 2, 4, 8 seconds, then fail.
- Atomic writes: every disk write that touches the xlsx, a `.md`, or a PDF goes through write-temp + `os.replace`.
- File lock: `filelock` around the xlsx. `crawl`, `scan-archive`, and `fetch-text` must not run concurrently against the same xlsx; the lock makes this safe even if attempted.
- Non-fatal errors logged to scoped logs under `data/` (`crawl_errors.log`, `extract_errors.log`, `verify_llm.log`).

## 8. Testing

Start with **none**. Add a `test_landtag.py` only when a parse regex actually breaks on a real PDF or a Webflow change actually trips us. Fixture-driven `pytest-httpx` mocks of the Webflow handshake are explicitly deferred — that scaffolding cost more than the design.

The operator-facing recovery for Webflow drift is `git diff` on `landtag.py`'s parse functions, not a green test suite.

## 9. Skill (`skills/landtag-nrw-extraction/SKILL.md`)

The skill points at the verbs, not at curl recipes (the Webflow handshake is too fiddly to expose raw):

- **When to use:** "Refreshing or querying Kleine Anfrage data from Landtag NRW."
- **First-time setup:** `landtag scan-archive` (offline, fills the bulk).
- **Refresh loop:** `landtag refresh-vocab` (rare) → `landtag crawl --wahlperiode 18` → `landtag fetch-text` → `landtag verify`.
- **Read access:** `data/index.xlsx` for filtering/aggregation; `Archiv/.../MMD18-N.md` for quoting.
- **Don't:** run two verbs concurrently against the same xlsx; hand-edit cols 15–17.

## 10. Repo layout

```
wdr-kleineanfrage/
├── CLAUDE.md
├── landtag.py                        ← all logic, ~300 lines
├── requirements.txt                  ← httpx, beautifulsoup4, openpyxl,
│                                       pdfplumber, pypdf, filelock, llm
├── Archiv/                           ← existing PDFs (not in git)
├── data/                             ← created on first run (not in git)
├── docs/superpowers/specs/
│   └── 2026-05-01-landtag-nrw-extraction-design.md
└── skills/landtag-nrw-extraction/SKILL.md
```

Dependencies live in `requirements.txt`, not `pyproject.toml`. No `src/` tree, no test infra, no lint/type toolchain pinned up front.

## 11. Open items deferred

- Front-matter in `.md` files (Anfragenummer, page markers).
- Große / Mündliche Anfrage.
- Cross-Wahlperiode aggregations as a verb.
- Per-MP detail-page enrichment for Abgeordnete.
- Test scaffolding (added if/when regexes break in practice).
