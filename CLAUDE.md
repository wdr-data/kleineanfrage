# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Read these first

- **`HANDOVER.md`** — current state, what's tested vs untested, OneDrive performance gotcha, recommended next steps. Read this before doing anything else.
- **`docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md`** — full design spec.
- **`skills/landtag-nrw-extraction/SKILL.md`** — operator-facing recipe for the four CLI verbs.

## Project context

Objective is to get data from the public parliamentary database of Landtag Nordrhein-Westfalen via this URL: https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html

### Data in question

Documents of interest:

- "Kleine Anfrage" parliamentary inquiries, and the answers to them, containing
  - Anfrager
  - Fraktion des Anfragers
  - Anfragedatum
  - Anfragenummer
  - Link Drucksache Anfrage
  - Link Drucksache Antwort (wenn vorhanden)Fraktion,
  - Antwortdatum
  - Antworttext
  - Systematik
  - Schlagworte
  - Ministerium (Kürzel)

Optional: Also use

- "Große Anfrage"
- "Mündliche Anfrage" (as answers are given in the protocols of parliamentary sessions, extraction might be more complicated)

### Designs needed

- a skill, i.e. a high-level description of tool uses to extract data as desired
- code for automating data extraction

### Typical use cases

- **Data evaluation from the XLSX** — quantitative analysis across the dataset, e.g. which Abgeordnete file the most Kleine Anfragen, distribution of answer turnaround times (Antwortdatum − Anfragedatum), topic/Schlagworte and ministry breakdowns.
- **Individual thematic queries in an agent chat session** — an agent uses the skill to refresh and surface information relevant to a specific question. The skill must therefore be invokable mid-conversation: the agent can run `crawl` to update metadata, `fetch-text` to pull answer texts on demand, and read selected `.md` files to quote/cite.

## Design guideline

- Keep it as simple and low-level as possible: command-line, no database
- Construct as a thin wrapper around XLSX files, local folders as cache (for PDF doc as well as .md extractions of the question and answer content)
- Design around the rodney tool: https://github.com/simonw/showboat-demos/blob/main/rodney/README.md
- Whenever rule-based extraction is possible, prefer to LLM-based extraction. Keep LLM use strictly to problematic cases and high-level plausibility checks.
- If errors are found, or higher-level eval is needed, use the llm library (PyPi) as a wrapper around LLM calls (for things like error correction, or semantic interpretation)

