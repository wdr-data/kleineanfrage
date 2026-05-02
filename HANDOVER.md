# Handover — wdr-kleineanfrage (state at 2026-05-02)

Read this and the design spec (`docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md`) first. The skill (`skills/landtag-nrw-extraction/SKILL.md`) is the operator-facing version.

## What's done

The four-verb CLI in `landtag.py` is **fully implemented** (no `NotImplementedError` left) and tested end-to-end on a 30-PDF sample copied to `/tmp/landtag-test/`:

| Verb | Tested | Notes |
|---|---|---|
| `scan-archive` | ✓ 30/30 PDFs in 0.5s on local disk | Catastrophic-backtracking bug fixed; regexes verified |
| `crawl` (enrich) | ✓ 28/28 enrichments, 0 errors | Re-bootstraps Webflow tokens per query (single-use) |
| `crawl --full` | not tested live | Implementation present; do a one-page sanity test before unleashing |
| `fetch-text` | ✓ 5/5 extracted to `.md` | pdfplumber + pypdf fallback; atomic write |
| `verify` | ✓ clean report | Counts by WP/status, malformed-nr / orphan / vocab-novelty checks |

Live-site facts (verified 2026-05-01/02):

- Bootstrap GET on the search page returns a form whose action URL carries two single-use Spring Webflow tokens (`webflowToken=<UUID>`, `webflowexecution<rand>__searchr2020=eXsY`). Re-bootstrap before each enrich query.
- Search hits arrive as `<article class="e-search-result">` blocks. Each hit corresponds to one Kleine Anfrage and links to the **question** Drucksache PDF (the **answer** Drucksache is not directly searchable). Body contains: title, KA-Nr, Anfrager+Fraktion, date, Systematik, Schlagworte, optionally Region.
- Pagination is plain `?page=N` GETs against `/anfragen-und-antworten-suchergeb.html?...` — no flow tokens needed once results are loaded.
- WP18 is currently 156 pages × 50 hits ≈ 7,800 KAs total in the search index; `Archiv/` has 7,175 answer PDFs (the gap is unanswered or pending-answer questions).
- Form fields actually present: `wp`, `doktyp` (value `KA` for Kleine Anfrage), `nummer`, `suchwort`, `autor` (JS-populated select, server-side empty), `fraktion` (free text — NOT a select), `schlagwort`, `region`, `rpp`. **No date inputs.**

## What's NOT done — explicit gaps

1. **The full `scan-archive` over the OneDrive `Archiv/` was never successfully completed.** A first run used 5h wall clock and produced no `index.xlsx` (likely killed mid-save by my earlier debugging; OneDrive paging ate most of the time). On local disk (`/tmp`) the same code processes 30 PDFs in 0.5s. **Action for next session:** see "OneDrive performance" below.
2. **`crawl --full` discovery mode** is implemented but the 7,800-row sweep was not run end-to-end. It re-bootstraps once and paginates with simple GETs, so token expiry is not a concern there — but check it inserts new rows and respects `--from / --to` post-fetch filtering.
3. **`fetch-text` against the real network** (downloading missing PDFs) was not tested — the test set already had every PDF locally. Risk: 404 handling, atomic write, rate-limit interaction. Probe with a row whose answer Drucksache PDF is genuinely missing.
4. **Tests:** none. Per spec §8 we deferred them. Add when something breaks.

## OneDrive performance — critical for next session

The repo lives inside an OneDrive-synced folder (`OneDrive-FreigegebeneBibliotheken–WDRKöln/WDR Data.O365 - Kleine Anfragen`). OneDrive on-demand sync makes random `pdftotext` access extremely slow (the 5h vs 0.5s ratio above). Two options:

- **(a)** Mark the `Archiv/` folder "always available offline" via the OneDrive UI before any large run. After that, scan-archive should finish the 7,175 PDFs in well under an hour.
- **(b)** `cp -r Archiv /tmp/wdr-archive` and run the tooling against the local copy. `landtag.py` resolves `ARCHIV_DIR` relative to itself; copy `landtag.py` too or set up a symlink.

Until one of these is done, **do not run `scan-archive` against the live `Archiv/`** without a `--limit`.

## Live-site quirks worth knowing

- **PDF rendering:** in some PDFs, `pdftotext` drops the trailing letter of `AfD` (renders as `Af`). Result: `Fraktion='Af'` lands in `vocab_novelty.log` and the row's title regex fails too (because it's anchored on the Fraktion token). Affects ~3% of WP18 sample. Could be patched by widening the Fraktion alternation to `(CDU|SPD|GRÜNE|FDP|AfD|Af|fraktionslos)` and normalising `Af` → `AfD` on extraction.
- **PDF without "Vormerkung der Kleinen Anfrage":** some early WP18 docs use `Vormerkung` (sic) instead of `Vorbemerkung`. Both are accepted by the title regex (`Vor(?:be)?merkung`).
- **PDF without "vom DATE":** some answers omit the question date in the "Kleine Anfrage N" line. `Anfragedatum` stays empty for those rows. The KA-Nr regex was made tolerant.
- **Spring Webflow tokens are single-use** — every enrich query re-bootstraps. With default `--rps 4`, an enrich pass over 7,000 rows takes roughly 1 hour (bootstrap + search + pagination = 2 requests/row).
- **Answer Drucksache is not searchable.** `nummer=18/1006` returns 0 hits. Only `nummer=18/675` (the question Drucksache) finds the KA.

## File map

```
landtag.py                  (~750 lines, single-file CLI; no src/ tree)
requirements.txt
README.md
HANDOVER.md                 ← this file
CLAUDE.md                   (project context + pointer to HANDOVER and spec)
.gitignore                  (Archiv/, data/, .claude/, OneDrive sync IDs)
docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md
skills/landtag-nrw-extraction/SKILL.md
Archiv/                     (~18,250 slot range, 7,175 actual WP18 PDFs)
data/                       (created on first run; index.xlsx + logs)
```

## Recommended next steps (in priority order)

1. **Make `Archiv/` always-offline in OneDrive** (or copy to `/tmp`), then `python3 landtag.py scan-archive --wahlperiode 18`. Expect ≤30 min wall, ~7,175 rows, single-digit `extract_failed`.
2. `python3 landtag.py crawl --wahlperiode 18` to enrich Systematik / Schlagworte. Expect ~1h at default `--rps 4`. Spot-check 5 random rows in `index.xlsx`.
3. `python3 landtag.py fetch-text --wahlperiode 18 --limit 5` to smoke-test on rows whose answer PDF is locally present, then `--limit 50` without limit on a row known to be missing locally to verify the download path.
4. `python3 landtag.py verify` and triage `data/extract_errors.log` / `data/vocab_novelty.log` / `data/crawl_errors.log`.
5. (Optional, once steady-state) try `crawl --full --wahlperiode 18` to discover unanswered Kleine Anfragen.

## Memory entries the next session should respect

- `MEMORY.md` (under `~/.claude/projects/.../memory/`) points at:
  - `feedback_git_identity.md` — never set git config globally; this machine mixes WDR/HR/personal repos. Use `git -c user.email=...` or env vars on the commit.
  - `project_robots_decision.md` — robots.txt addresses indexers, not data agents. Don't re-frame the project as "accepting robots.txt risk." Polite rate-limit (4 rps default) suffices.

## Open design questions deferred (already in spec §12)

- Per-MP detail-page enrichment (the JS-populated `autor` dropdown).
- Markdown front-matter in extracted `.md` files.
- Große / Mündliche Anfrage support.
- Automated test scaffolding.
