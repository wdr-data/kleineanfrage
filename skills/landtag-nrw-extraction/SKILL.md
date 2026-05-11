---
name: landtag-nrw-extraction
description: Use when extracting or querying Kleine Anfrage data from Landtag NRW — metadata + answer-text harvesting into Excel + Markdown.
---

# Landtag NRW — Kleine Anfrage extraction

Idempotente CLI-Verben in `landtag.py` über drei Datenebenen:

- `data/db_index.xlsx` — **immutable DB-Snapshot** der Landtag-Suche, geschrieben nur von `crawl`. Verifikations-Referenz für die DB-Sicht. Nie von Hand editieren.
- `data/datum_original.xlsx` — **authoritative Datums-Quelle**: `Datum_Original` (Briefdatum, „Datum des Originals: DD.MM.YYYY") + `Datum_Ausgegeben` (Drucksachen-Veröffentlichung) pro PDF-Dokument (Anfrage & Antwort, eine Zeile je Drucksache). Geschrieben von `tools/extract_datum_original.py`. Seit 2026-05-11 maßgeblich für `Anfragedatum`/`Antwortdatum` in `index.xlsx`; DB-Datum dient nur noch der Korridor-Verifikation.
- `data/index.xlsx` — **Working-File** mit DB-Metadaten (Anfrager, Fraktion, Systematik, …) + PDF-Anreicherungen (Datum, Ministerium, Anfrager_Alle …) + Qualitäts-Spalten. Wird durch die Pipeline-Verben gepflegt.
- `Archiv/.../MMD<wp>-<nr>.{pdf,md}` — Original-PDFs + pdftotext-Cache.

Stammdaten-Tabellen liegen in `Index/` (Fraktionen, Ministerien, Abgeordnete) — siehe `vocabulary.md`.

`merge` materialisiert den DB↔PDF-Cross-Check pro Zeile in drei Spalten:
- `Mismatch_Flags` — `","`-getrennte Liste der Felder, in denen DB und PDF widersprechen (`anfragedatum`, `antwortdatum`, `drucksache_anfrage_nr`, `fraktion`, `md_kanr`).
- `Datenqualität` — `ok` / `korrigiert` (manueller Tag oder nicht-leere `Notizen`) / `ask_review` (offener Mismatch ohne Notiz).
- `Notizen` — Free-Text, Domain des `resolve`-Verbs. Sobald non-empty, bleibt die Zeile bei späteren `merge`-Läufen `korrigiert` (nicht-rückgängige Triage).

**Datums-Autorität & Korridor-Verifikation**: `Anfragedatum`/`Antwortdatum` in `index.xlsx` werden aus `datum_original.xlsx` (`Datum_Original`, also dem Briefdatum auf Seite 1 jedes PDFs) befüllt — das ist seit 2026-05-11 die Autoritäts-Quelle. Das DB-Datum aus `db_index.xlsx` wird nur noch zur Verifikation gegen das PDF-Datum gehalten. Korridor-Vergleich in `_date_mismatch`:

- **Innerhalb Korridor** (PDF früher 0–122d, beide Jahre 2015–2026): erwartbarer Verarbeitungs-Lag (Briefdatum vs Drucksachen-Veröff.). PDF-Datum wird stillschweigend in `index.xlsx` übernommen, weder Flag noch Notiz.
- **Außerhalb Korridor** (Diff > 122d, DB-früher > 14d, Jahres-Tippfehler ~365d, OCR-Glitch außerhalb [2015,2030]): `Mismatch_Flags` enthält `anfragedatum`/`antwortdatum`, `Datenqualität=ask_review`. `resolve --auto` setzt das maßgebliche Datum nach Heuristik-Tabelle (in der Regel PDF; bei klarem PDF-Parse-Artefakt ausnahmsweise DB) und schreibt das jeweils *nicht* übernommene Datum als „war: <Datum>" in `Notizen`.
- **PDF-Datum fehlt** (`Antworttext_Status=pdf_missing`, oder `no_match` in `datum_original.xlsx`): DB-Datum bleibt als Fallback in `index.xlsx`, mit `Extract_Flags=datum_pdf_missing`.

**pdftotext-Quirks** im Header-Parser (`_parse_pdf_header_fields`): Spacing-tolerant (`"Anfrage 5vom"`, `"Drucksache17/32"`) und Doubled-Character-Render (`"LLAANNDDTTAAGG …"` → automatischer Dedouble via `_maybe_dedouble`).

`Antworttext_Status` kennt zusätzlich `pdf_missing` (manuell im `resolve`-Schritt gesetzt) — Landtag-CDN liefert unter der DS-Nr ein anderes Dokument (typischerweise die zugehörige Anfrage statt der Antwort, oder ein Routing-Bug zwischen ähnlichen DS-Nrs). `fetch-text` überspringt diese Zeilen, bis `--force` gesetzt wird.

## When to use

- Frage nach **Wer/Was/Wann** einer Kleinen Anfrage → `data/index.xlsx` abfragen.
- Frage nach **allen Mit-Anfragenden** → Spalte `Anfrager_Alle` (DB-Spalte `Anfrager` ist auf 2 Namen + `u.a.` gekappt).
- Frage nach **allen beteiligten Ministerien** → Spalte `Beteiligte_Ministerien_Kuerzel` (Search-Hit nennt nur das federführende Ressort).
- Frage nach **Inhalt einer konkreten Antwort** → `Archiv/.../MMD<wp>-<nr>.md`.
- Frage nach **PDF-Briefdatum** (Datum des Originals) oder Verifikation Anfrage-/Antwortdatum gegen DB → `data/datum_original.xlsx`.
- **Quantitative Auswertung** → mit pandas auf der XLSX (siehe „Read access" unten).
- Wunsch nach **frischen Daten** → Refresh-Loop unten.
- **Zweifel an einer einzelnen Zeile** → `Datenqualität=ask_review` filtern, dann `resolve --auto` (Heuristik), gezielter `resolve --ka N`, oder im Restfall `resolve --interactive`.

## Pipeline

```sh
python landtag.py crawl                     --wahlperiode 18  # discovery + DB-Snapshot (~25 min)
python landtag.py fetch-text                --wahlperiode 18  # lädt fehlende Antwort-PDFs, schreibt .md
python landtag.py scan-archive              --wahlperiode 18  # reichert Anfragetitel + Min-Prosaname
python landtag.py extract-multi-ministerium                   # alle beteiligten Ressorts
python landtag.py build-abgeordnete-index                     # Index/abgeordnete.xlsx aus DB-Anfrager
python landtag.py extract-all-anfrager                        # Anfrager_Alle + Anzahl_Abgeordnete
python landtag.py build-abgeordnete-index                     # 2. Pass: absorbiert 3.+ Anfrager
python landtag.py extract-all-anfrager                        # 2. Pass: matcht erweiterten Namensindex
python tools/extract_datum_original.py                        # PDF-Briefdaten → data/datum_original.xlsx
python landtag.py merge                                       # PDF↔DB-Datums-Korridor + Mismatch_Flags + Datenqualität
python landtag.py resolve --auto                              # Heuristik-Auflösung in Richtung PDF-Datum
python landtag.py verify                                      # sanity report
```

`crawl` liefert pro KA: beide Drucksachen + Links, Anfragedatum, Antwortdatum, Anfrager (max 2 + `u.a.`), Fraktion, Titel, Systematik, Schlagworte, federführendes Ministerium-Kürzel. Alles weitere kommt aus dem Antwort-PDF.

`normalize` (Fraktion/Ministerium → kanonisch) wird automatisch am Ende von `crawl` und `fetch-text` aufgerufen. `enrich-llm` ist optional für Rest-Lücken.

Mehrere Wahlperioden in derselben `index.xlsx` möglich — jede Zeile trägt ihre `WP`. Date-bounded subset (client-side):

```sh
python landtag.py crawl --wahlperiode 18 --from 2024-01-01 --to 2024-12-31
```

## Resolve-Pipeline für `ask_review`-Zeilen

Nach `merge` zeigt `Datenqualität=ask_review`, welche Zeilen ungeklärte DB↔PDF-Mismatches haben. Reihenfolge:

```sh
python landtag.py resolve --auto                              # Stage 1
python landtag.py resolve --ka 144 --counter-search           # Stage 2 (gezielt)
python landtag.py enrich-llm                                  # Stage 3 (Rest-Lücken)
# Stage 4: Web-Recherche pro Zeile, Stage 5: resolve --interactive — siehe unten
```

1. **`resolve --auto`** — Heuristik-Pass über alle ask_review-Zeilen. Klassifiziert bekannte Mismatch-Muster (PDF-OCR-Glitch, Drucksachen-vs-Brief-Datum, Fraktionsaustritt, DS-Nr-Tippfehler, pdftotext-Spacing-Glitch …) und setzt `Datenqualität=korrigiert` mit erklärender Notiz. Idempotent; `--dry-run` für Trockenlauf. Schließt typischerweise 90 %+ der ask_review-Zeilen. Heuristiken siehe Tabelle unten.
2. **`resolve --ka N [--counter-search] [--delete-phantom]`** — Live-DB-Re-Fetch für eine konkrete KA-Nr. Use cases: Phantom-Zeile, md↔KA-Nr-Diff > Heuristik-Korridor, Korrigendum auf Landtag-Seite, oder Verdacht auf stale `db_index.xlsx`-Snapshot.
3. **`enrich-llm`** — LLM-Rescue für Rest-Lücken in `Extract_Flags` (nicht primärer Resolve-Pfad, aber hilft bei Anfrager-/Ministerium-Lücken).
4. **Web-Recherche** — Agent recherchiert via Web-Tools (Landtag-Suche live, offizielle MdL-Liste). Notizen-Konvention: `Quelle: <URL> — DB-Wert bestätigt / korrigiert auf <Wert>.` (filterbar via `Notizen.str.contains("Quelle:")`).
5. **`resolve --interactive`** — Human-Dialog für die letzten unklaren Zeilen. Pro Zeile: DB↔PDF anzeigen, Notiz + Verdict erfragen. Resume-safe. **Stdin-Pipe ist fragil** (bedingter Verdict-Prompt — leere Notiz = kein Verdict-Read; nicht-leere Notiz = zusätzlicher Read); für Batch-Prozessierung lieber `--auto` oder `openpyxl`-Direkt-Write nutzen.

### Heuristik-Tabelle (DB↔PDF, wer hat recht)

`resolve --auto` implementiert exakt diese Tabelle (`_classify_mismatch` in `landtag.py`).

| Mismatch | Heuristik | Verdict |
|---|---|---|
| `antwortdatum` Diff 0–122d, beide Jahre 2015–2026 | DB = Drucksachen-/Veröff.-Datum, PDF = Briefdatum „mit Schreiben vom …" | **PDF maßgeblich** (Korridor) — `index.xlsx`-Antwortdatum auf PDF, stille Übernahme ohne Notiz |
| `antwortdatum` Diff ~365d, gleiches MM/DD | Jahres-Tippfehler im PDF | **DB maßgeblich** (Ausnahme: PDF-Parse-Artefakt); Notiz „war (PDF): <PDF-Datum>" |
| `antwortdatum` Jahr außerhalb [2015,2026] | pdftotext-OCR-Glitch | **DB maßgeblich** (Ausnahme: PDF-Parse-Artefakt); Notiz „war (PDF): <PDF-Datum>" |
| `antwortdatum` Diff > 122d sonst | unklar, Korrigendum / Re-Datierung möglich | **PDF maßgeblich**, `ask_review` bleibt offen; Notiz „war (DB): <DB-Datum>, jetzt (PDF): <PDF-Datum>, Diff Nd" |
| `anfragedatum` Diff 0–122d (DB ≥ PDF) | DB = Drucksachen-Datum, PDF = Datum auf Anfrage-Schreiben | **PDF maßgeblich** (Korridor) — `index.xlsx`-Anfragedatum auf PDF, stille Übernahme ohne Notiz |
| `anfragedatum` Diff ~365d / Jahr außerhalb Bereich | Tippfehler / OCR-Glitch im PDF | **DB maßgeblich** (Ausnahme: PDF-Parse-Artefakt); Notiz „war (PDF): <PDF-Datum>" |
| `drucksache_anfrage_nr` PDF-Wert = KA-Nr | PDF-Parser fing falsche Drucksache-Zeile | DB maßgeblich |
| `drucksache_anfrage_nr` PDF = `<WP>/<WP>` (z. B. `17/17`) | False-Positive auf Wahlperiode-Header | DB maßgeblich |
| `drucksache_anfrage_nr` Diff Off-by-1/10/100/1000 oder zusätzliche/fehlende Ziffer | OCR-Digit-Tippfehler im PDF | DB maßgeblich |
| `drucksache_anfrage_nr` PDF-WP ≠ DB-WP | PDF-Parser fing falschen Header | DB maßgeblich |
| `fraktion` AfD ↔ fraktionslos | Fraktionsaustritt zwischen Anfrage und Antwort (Pretzell, Vogel/Neppe/Langguth Herbst 2017) | beide legitim |
| `fraktion` PDF zeigt andere Partei als Anfrager | Landtag-PDF-Header-Tippfehler oder Multi-Fraktions-Anfrage | DB = Fraktion der Anfrager:innen maßgeblich |
| `md_kanr`, PDF-Parser-Output komplett leer, .md enthält literal `Anfrage <KA>vom` | pdftotext-Spacing-Glitch | DB+PDF konsistent, Mismatch-Flag = Parser-Artefakt |
| `md_kanr`, .md enthält jedes Zeichen doppelt (`LLAANNDDTTAAGG`) | pdftotext-Render-Quirk bei bestimmten Schrift-Embeddings | Parser-Artefakt, DB maßgeblich |
| `md_kanr`, .md-KA-Nr passt zu *anderer* Drucksache | Landtag-PDF-Body nennt fremde KA-Nr (Doppelvergabe / Druckfehler) | DB-Snapshot maßgeblich |
| .md-Inhalt ist die Anfrage statt der Antwort | Landtag-CDN-Routing-Bug — Antwort-PDF effektiv nicht verfügbar | `Antworttext_Status=pdf_missing`, `Antworttext` leeren, Notiz |

Gezielte Direkt-Reparatur einzelner KAs (Phantom, md↔KA-Mismatch, Korrigendum):

```sh
python landtag.py resolve --ka 144 --ka 3167 --counter-search --delete-phantom
```

Edge-Case-Pattern und Tippfehler-Tabelle: siehe `edge-cases.md`. Schreibvarianten und Aliase: siehe `vocabulary.md`.

## Read access

```python
import pandas as pd
df = pd.read_excel("data/index.xlsx")
df[df.Fraktion == "AfD"]
df[df.Anfrager_Alle.str.contains("Wagner", na=False)]                 # FULL Anfrager-Liste
df[df.Schlagworte.str.contains("Polizei", na=False)]
df[df.Ministerium_Kuerzel == "IM"]                                    # federführend
df[df.Beteiligte_Ministerien_Kuerzel.fillna("").str.contains("IM")]   # auch beteiligt
df[df.Anfragedatum >= "2024-01-01"]
df[df.Anzahl_Abgeordnete >= 5]                                        # Sammelanfragen
df[df.Datenqualität == "ok"]                                          # nur saubere Zeilen
df[df.Antworttext_Status == "pdf_missing"]                            # Antwort-PDF auf Landtag-CDN nicht verfügbar
```

`Anfrager` (DB-Spalte) bleibt als Rohwert erhalten — max 2 Namen + `u.a.`. **Für Auswertungen `Anfrager_Alle` nutzen.**

Antworttext einer Zeile:

```python
md_path = df.loc[df.Drucksache_Antwort_Nr == "18/1006", "Antworttext"].iloc[0]
print(open(md_path, encoding="utf-8").read())
```

CSV-Export:

```sh
python -c "import pandas as pd; pd.read_excel('data/index.xlsx').to_csv('out.csv', index=False)"
```

## Don't

- `data/db_index.xlsx` von Hand editieren oder mit anderen Verben überschreiben — immutable Snapshot, ausschließlich Domain von `crawl`. Bleibt erhalten als Verifikations-Referenz, auch wenn die Datums-Autorität an `datum_original.xlsx` übergegangen ist.
- `data/datum_original.xlsx` von Hand editieren — Domain von `tools/extract_datum_original.py`. Bei Fehlerverdacht (no_match / parse_failed) mit `--force` neu extrahieren; vor manueller Korrektur prüfen, ob das PDF auf Seite 1 wirklich kein „Datum des Originals:" enthält (manche PDFs mit Doubled-Character-Render-Quirk brauchen Dedouble).
- Spalten `Anfragedatum`/`Antwortdatum` in `index.xlsx` von Hand auf DB-Wert zurücksetzen — werden bei `merge` aus `datum_original.xlsx` neu befüllt. Eingriffe gehören in `Notizen` + `Extract_Flags=datum_manual`.
- Spalten `Antworttext`, `Antworttext_Status`, `Antworttext_Quelle` von Hand editieren — `fetch-text`-Domain.
- Spalten `Anfrager_Alle`, `Anzahl_Abgeordnete`, `Beteiligte_Ministerien_Kuerzel` von Hand editieren — werden bei Re-Run überschrieben. Manuelle Korrekturen mit `anfrager_manual` (in `Extract_Flags`) markieren, dann werden sie geschützt.
- Zwei Verben **gleichzeitig** auf dieselbe XLSX. File-Lock blockt zwar, ist aber unsauber.
- Aus `Anfrager` (DB-Spalte) auf die echte Anzahl Mit-Anfragender schließen.
- Werte aus `data/vocab_novelty.log` automatisch absorbieren — siehe `vocabulary.md`.
- `--rps` über 4 ohne Grund.
- `Antworttext_Status=pdf_missing` rückgängig machen, ohne den Landtag-CDN-Bug zu verifizieren — `fetch-text` würde wieder denselben falschen PDF-Inhalt einlesen. Re-Try nur mit `fetch-text --force`, nachdem manuell geprüft wurde, dass der Landtag jetzt den richtigen PDF ausliefert.

## Reference

- `edge-cases.md` — Anfrager-Quirks, PDF-Tippfehler, Status-Edge-Cases, Crawl-Quirks.
- `vocabulary.md` — Stammdaten, Namens-/Format-Konventionen, Fraktion- und Min-Aliase.
- `docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md` — vollständige Architektur-Spec.
- `auswertung_fehlende_daten.md` — Operator-Beobachtungen.
