---
name: landtag-nrw-extraction
description: Use when extracting or querying Kleine Anfrage data from Landtag NRW — metadata + answer-text harvesting into Excel + Markdown.
---

# Landtag NRW — Kleine Anfrage extraction

Sechs idempotente CLI-Verben in `landtag.py`, alle wrappen ein gemeinsames `data/index.xlsx` und einen `.md`-Cache neben den Original-PDFs in `Archiv/`.

## When to use

- Frage nach **Wer/Was/Wann** einer Kleinen Anfrage (Anfrager, Fraktion, Datum, Ministerium, Schlagworte) → `data/index.xlsx` abfragen.
- Frage nach **Inhalt einer konkreten Antwort** → `Archiv/.../MMD18-N.md`. Fall back auf PDF nur wenn `.md` fehlt.
- **Quantitative Auswertung** über den Datenbestand → mit pandas auf der XLSX.
- Wunsch nach **frischen Daten** → Refresh-Loop (siehe unten).
- **Zweifel an einer einzelnen Zeile** → Nachrecherche-Strategie (siehe unten).

## First-time setup & full refresh

`crawl` ist Canon — die Online-Suche bestimmt, welche KAs der Wahlperiode existieren. Reihenfolge (jeder Verb idempotent):

```sh
python landtag.py crawl                       --wahlperiode 18  # discovery + enrich aus Search (~25 min)
python landtag.py fetch-text                  --wahlperiode 18  # lädt fehlende Antwort-PDFs, schreibt .md
python landtag.py scan-archive                --wahlperiode 18  # reichert Anfragetitel + Min-Prosaname an
python landtag.py extract-multi-ministerium                     # parst Antworttext-MD nach allen beteiligten Ressorts
python landtag.py verify                                        # sanity report
```

`crawl` zieht aus dem Search-Index für jede KA: beide Drucksachen (Anfrage+Antwort) mit Links + Anfragedatum, Antwortdatum, Anfrager, Fraktion, Titel, Systematik, Schlagworte UND das antwortende Ministerium-Kürzel direkt. `scan-archive` läuft im Default als **enrich-only** und überschreibt nie Crawl-Werte (`--allow-discovery` für legacy / standalone).

`normalize` (Fraktion/Ministerium → canonical) wird automatisch am Ende von `crawl` und `fetch-text` aufgerufen; `enrich-llm` ist optional für Rest-Lücken.

Mehrere Wahlperioden in derselben `index.xlsx` sind unterstützt — jede Zeile trägt ihre `WP`. Z. B. WP17 ergänzen:

```sh
python landtag.py crawl --wahlperiode 17    # legt WP17-Zeilen an
python landtag.py fetch-text --wahlperiode 17
python landtag.py scan-archive --wahlperiode 17
```

Date-bounded subset (filter is client-side, applied after fetch):

```sh
python landtag.py crawl --wahlperiode 18 --from 2024-01-01 --to 2024-12-31
```

## Nachrecherche-Strategie (für einzelne Zweifelsfälle)

`verify` listet Verdachtsfälle (md ↔ KA-Nr-Mismatch, Lücken, Duplikate). Reparatur in der Regel über das `resolve`-Verb:

```sh
# Eine oder mehrere KA-Nrn frisch von der Live-Suche holen, stale Phantom-Zeilen löschen:
python landtag.py resolve --ka 144 --ka 3167 --counter-search --delete-phantom
```

`--counter-search`: bei 0 KA-Treffern wird mit `doktyp=GA` (Große Anfrage) nachgeschaut — wenn dort gefunden, ist die fragliche Zeile ein Out-of-Scope-Phantom. `--delete-phantom`: löscht sowohl GA-Phantome als auch stale Zeilen (gleiche KA-Nr, aber andere Antwort-Drucksache als der Live-Treffer).

Manuelle Direktsuche im Browser (falls man's selbst sehen will):
```
https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten-suchergeb.html?nummer=<N>&doktyp=KA&wp=18
```

Das `verify`-Verb liefert in seinem Report:
- **md ↔ KA-Nr-Mismatch**: Antwort-PDF erwähnt eine andere KA-Nr als der Index — meist ein Crawl-Fehlmatch oder ein kaputtes PDF (pdftotext rendert manche Files mit Doppel-Buchstaben).
- **KA-Nr Lücken**: KAs sind fortlaufend nummeriert. Lücken sind entweder zurückgezogen (`anfrage_zurueckgezogen`) oder noch unveröffentlicht.
- **Duplikat-KA-Nr**: gleiche KA-Nr in mehreren Zeilen. Häufig **legitim** (Korrigenda / Nachgang); rarely a Parser-Bug.

### Bekannte Sonderfälle

- **`Antwort: Unterrichtung Präs`** im Search-Hit → Anfrage wurde zurückgezogen; das „Antwort"-Dokument ist eine Unterrichtung des Landtagspräsidenten. Wird automatisch mit Status `anfrage_zurueckgezogen` markiert (siehe `parse_search_hits`); `Antwortdatum` = Datum der Unterrichtung. Skipt in `fetch-text`/`scan-archive`.
- **`MCdS`** (Minister + Chef der Staatskanzlei) ist dasselbe Ministerium wie **`MBEIM`** (gleiche Person, andere Rolle). Bereits via `Aliases`-Spalte in `Index/ministerien.xlsx` zusammengeführt.
- **`MKJFGF` vs. `MKJFGFI`**: Familienministerium (Kinder, Jugend, Familie, Gleichstellung, Flucht und Integration). Korrektes Kürzel ist `MKJFGFI`; alte Schreibweise `MKJFGF` ist als Alias gelistet.
- **Mehrere beteiligte Ministerien:** Antwort-PDFs haben nach dem Anfragetext einen Absatz „Der Minister … / Die Ministerin …" der ggf. mehrere Ressorts nennt (z. B. KA 4338: Familienministerium hat Rücksprache mit Innenministerium gehalten). Dieser Absatz ist die autoritative Quelle für **alle** beteiligten Ministerien — der Search-Hit nennt nur das federführende. Extraktion via `landtag.py extract-multi-ministerium` → schreibt comma-separierte Kürzel-Sets nach `Beteiligte_Ministerien_Kuerzel` (federführend zuerst). WP18 abgedeckt; WP17 hat anderen Ministeriumszuschnitt und ist out of scope.
- **Anfrage-PDFs im `Archiv/`** sehen auf Seite 1 fast identisch aus wie die Antworten und werden seit dem Anfrage-PDF-Filter sauber ausgesondert.
- **Antworten auf Große Anfragen** liegen z.T. mit im `Archiv/` und werden über den GA-Filter ausgesondert.
- **Wahlperioden-Wechsel** ändern Ministeriumszuschnitt (WP17 = Schwarz-Gelb, WP18 = Schwarz-Grün). Neue Kürzel landen in `data/vocab_novelty.log` — vor automatischem Mergen prüfen.

## Read access

```python
import pandas as pd
df = pd.read_excel("data/index.xlsx")
df[df.Fraktion == "AfD"]                                # by Fraktion
df[df.Anfrager.str.contains("Wagner", na=False)]        # by Abgeordnete
df[df.Schlagworte.str.contains("Polizei", na=False)]   # by Schlagwort
df[df.Ministerium_Kuerzel == "IM"]                      # by Kürzel
df[df.Anfragedatum >= "2024-01-01"]                     # by Datum
```

Antworttext einer Zeile:

```python
md_path = df.loc[df.Drucksache_Antwort_Nr == "18/1006", "Antworttext"].iloc[0]
print(open(md_path, encoding="utf-8").read())
```

CSV/JSONL export:

```sh
python -c "import pandas as pd; pd.read_excel('data/index.xlsx').to_csv('out.csv', index=False)"
```

## Don't

- Zwei Verben **gleichzeitig** auf dieselbe XLSX. File-Lock blockt zwar, ist aber unsauber.
- Spalten `Antworttext`, `Antworttext_Status`, `Antworttext_Quelle` von Hand editieren — `fetch-text` Domain.
- Werte aus `data/vocab_novelty.log` automatisch korrigieren. Eine ungewohnte Fraktion/Ministerium kann ein echter neuer Eintrag sein. Erst prüfen, dann ggf. `Index/ministerium_aliases.xlsx` ergänzen oder `Index/ministerien.xlsx` anpassen.
- `--rps` über 4 ohne Grund. Polite-Default für ein kleines public-sector site.
- Bei `crawl`-0-Hits blind retryen — vermutlich `bootstrap_search` / Pagination-Parser-Anker hat sich verschoben (HTML-Struktur des Landtags geändert).

## Common failures

- **`pdftotext: not found`** → `brew install poppler` (macOS).
- **`pdftotext` schluckt das D in „AfD"** → Spelling landet als "Af" in `vocab_novelty.log`. Bekannter pdftotext-Quirk auf manchen PDFs.
- **`Vormerkung` vs `Vorbemerkung`** → frühe WP18-Antworten nutzen alte Schreibweise; Title-Regex akzeptiert beide.
- **`fetch-text` reportet 404** → Drucksache existiert online nicht (zurückgezogen oder noch unveröffentlicht). Status `no_answer_yet` ist erwartet.
- **`Extract_Flags` enthält `missing_ministerium`** auf einer Zeile mit gefülltem Kürzel → Bug-Indikator. `compute_extract_flags` checkt seit Fix `ministerium OR ministerium_canonical OR ministerium_kuerzel`. Wenn Lücke trotzdem markiert: nochmal `normalize` laufen lassen, dann manueller Re-Compute der Flags.

## Reference

Vollständige Spezifikation: `docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md`.
Operator-Beobachtungen: `auswertung_fehlende_daten.md`.
