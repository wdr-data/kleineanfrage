---
name: landtag-nrw-extraction
description: Use when extracting or querying Kleine Anfrage data from Landtag NRW — metadata + answer-text harvesting into Excel + Markdown.
---

# Landtag NRW — Kleine Anfrage extraction

Idempotente CLI-Verben in `landtag.py`, alle wrappen ein gemeinsames `data/index.xlsx` und einen `.md`-Cache neben den Original-PDFs in `Archiv/`. Stammdaten-Tabellen liegen in `Index/` (Fraktionen, Ministerien, Abgeordnete).

`crawl` schreibt zusätzlich einen reinen DB-Snapshot nach `data/db_index.xlsx` — nur die Felder aus der Landtag-Suche, ohne PDF-Anreicherung. Diese Datei ist die immutable Single-Source-of-Truth für die DB-Sicht (nie von Hand editieren, nie von anderen Verben überschrieben).

## When to use

- Frage nach **Wer/Was/Wann** einer Kleinen Anfrage (Anfrager, Fraktion, Datum, Ministerium, Schlagworte) → `data/index.xlsx` abfragen.
- Frage nach **allen Mit-Anfragenden** (auch jenseits der ersten 2) → Spalte `Anfrager_Alle` (siehe „Anfrager-Vollständigkeit"). DB-Spalte `Anfrager` ist auf 2 Namen + `u.a.` gekappt.
- Frage nach **allen beteiligten Ministerien** → Spalte `Beteiligte_Ministerien_Kuerzel`. Search-Hit nennt nur das federführende Ressort.
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
python landtag.py build-abgeordnete-index                       # baut Index/abgeordnete.xlsx aus DB-Anfrager-Spalte
python landtag.py extract-all-anfrager                          # füllt Anfrager_Alle + Anzahl_Abgeordnete aus Antwort-PDF
python landtag.py build-abgeordnete-index                       # 2. Pass: absorbiert Namen, die nur in u.a.-PDFs auftauchten
python landtag.py extract-all-anfrager                          # 2. Pass: matcht jetzt auch diese Namen
python landtag.py verify                                        # sanity report
```

`crawl` zieht aus dem Search-Index für jede KA: beide Drucksachen (Anfrage+Antwort) mit Links + Anfragedatum, Antwortdatum, Anfrager, Fraktion, Titel, Systematik, Schlagworte UND das antwortende Ministerium-Kürzel direkt. **Vorsicht:** Der Search-Hit liefert für `Anfrager` nur die ersten **2** Mit-Anfragenden plus `u.a.`, wenn 3+ Abgeordnete eine KA gemeinsam einreichen. Die volle Liste steht im Antwort-PDF und wird über `build-abgeordnete-index` + `extract-all-anfrager` in die Spalten `Anfrager_Alle` / `Anzahl_Abgeordnete` geschrieben (siehe „Anfrager-Vollständigkeit"). Auch das im Search-Hit genannte Ministerium-Kürzel ist nur das **federführende**; zusätzlich beteiligte Ressorts werden über `extract-multi-ministerium` aus dem PDF-Boundary-Absatz nachgezogen (Spalte `Beteiligte_Ministerien_Kuerzel`). `scan-archive` läuft im Default als **enrich-only** und überschreibt nie Crawl-Werte (`--allow-discovery` für legacy / standalone).

`normalize` (Fraktion/Ministerium → canonical) wird automatisch am Ende von `crawl` und `fetch-text` aufgerufen; `enrich-llm` ist optional für Rest-Lücken.

Mehrere Wahlperioden in derselben `index.xlsx` sind unterstützt — jede Zeile trägt ihre `WP`. Z. B. WP17 ergänzen:

```sh
python landtag.py crawl --wahlperiode 17    # legt WP17-Zeilen an
python landtag.py fetch-text --wahlperiode 17
python landtag.py scan-archive --wahlperiode 17
python landtag.py build-abgeordnete-index   # Index/abgeordnete.xlsx hält WP17 + WP18 nebeneinander
python landtag.py extract-all-anfrager --wahlperiode 17
```

`Index/abgeordnete.xlsx` enthält jeweils eine Zeile pro (`WP`, `Fraktion`, `Person`) — Wahlperioden-übergreifend. `extract-all-anfrager --wahlperiode 17` matcht dabei nur gegen die WP17-Einträge (jede WP hat eigene Abgeordnete).

Date-bounded subset (filter is client-side, applied after fetch):

```sh
python landtag.py crawl --wahlperiode 18 --from 2024-01-01 --to 2024-12-31
```

## Anfrager-Vollständigkeit (Spalten `Anfrager_Alle`, `Anzahl_Abgeordnete`)

**Bug:** Die Suchergebnisseite des Landtags listet pro KA maximal 2 Anfragende, danach `u.a.` (~1.400 Zeilen sind betroffen). Die Antwort-PDF (Header) führt aber alle Mit-Anfragenden auf. Der Workflow gleicht das gegen einen lokalen Abgeordneten-Index ab.

```sh
python landtag.py build-abgeordnete-index   # → Index/abgeordnete.xlsx
python landtag.py extract-all-anfrager      # füllt Anfrager_Alle, Anzahl_Abgeordnete
python landtag.py build-abgeordnete-index   # Bootstrap-Pass 2 (siehe unten)
python landtag.py extract-all-anfrager
```

`build-abgeordnete-index` aggregiert alle (`WP`, `Fraktion`, `Nachname`, `Vorname`)-Tupel, die in der DB-Spalte `Anfrager` auftauchen. Da die Datenbank nur die ersten 2 Anfragenden je KA liefert, fehlen ggf. Personen, die ausschließlich als 3.+ Mit-Anfragende auftauchen — nach dem ersten `extract-all-anfrager`-Lauf stehen diese aber bereits in `Anfrager_Alle`, deshalb der **zweite Build+Extract-Pass**: er absorbiert diese Namen und matcht in der Folge auch KAs, die zuvor mit Residue-Eintrag im Log endeten.

**Offizielle Quelle ergänzen** — der Landtag NRW veröffentlicht eine vollständige Liste der aktuellen Abgeordneten unter <https://www.landtag.nrw.de/home/der-landtag/abgeordnete-und--fraktionen/die-abgeordneten/abgeordnetensuche/liste-aller-abgeordneten.html>. Die Bootstrap-Aggregation aus DB-Anfragern erfasst nur Personen, die mindestens einmal als 1./2. Anfrager (oder schon im 2. Pass als 3.+) auftauchen — Hinterbänkler*innen, die nur sehr selten oder gar nicht KAs mitzeichnen, fehlen. Vor dem ersten Lauf einer neuen WP daher idealerweise diese Liste scrapen und als zusätzliche Datenquelle in `Index/abgeordnete.xlsx` einspielen (separates One-shot-Skript, noch nicht implementiert; bis dahin bleibt der DB-only-Bootstrap das Default-Verhalten).

**Format-Konvention:**
- DB / `Anfrager` / `Anfrager_Alle`: `Nachname, Vorname` (mehrere `; `-getrennt). Titel (`Dr.`, `Prof. Dr.`) bleiben am Anfang des `Nachname`-Teils.
- Antwort-PDF-Header: `Vorname Nachname` (Komma-Liste, letzter mit `und`). Titel werden vor den Vornamen gestellt. Der Konverter `db_to_pdf_form` erzeugt aus DB-Form die PDF-Form; `db_to_pdf_form_aliases` zusätzlich vereinfachte Varianten (Mittelinitial weg, Hyphen-Mittelnamen weg, Titel weg) — nötig, weil PDFs oft kürzere Schreibweisen verwenden („Sven W. Tritschler" → „Sven Tritschler").

**Edge-Case Fraktionswechsel:** Verlässt eine Person ihre Fraktion mitten in der WP (z. B. Pretzell, Neppe, Müller-Witt), erscheint sie in der DB unter beiden Fraktionen. Der Index trägt dann je eine Zeile pro (WP, Fraktion, Person); die Spalte `Frueher_Fraktion` listet alle anderen Fraktionen, unter denen dieselbe Person in derselben WP geführt wird — als Doku und als Fallback-Hinweis bei Match-Misserfolgen.

**Residue-Log:** Was der Parser im PDF-Anfragerblock nicht zuordnen kann, landet in `data/anfrager_novelty.log` (Format: Drucksache | Fraktion | residue | block). Vor manueller Korrektur prüfen, ob ein dritter Bootstrap-Pass oder ein Eintrag aus der offiziellen MdL-Liste das Problem löst.

### Plausi-Check + Edge-Case-Nachrecherche durch den Agenten

`extract-all-anfrager` vergleicht am Ende jeder Zeile die **geparste Anzahl** mit der **Mindesterwartung aus der DB-Anfrager-Spalte** (1 oder 2 Namen, oder ≥3 wenn `u.a.` markiert). Liegt die Parser-Zahl darunter → Eintrag ins Mismatch-Log:

```
… | 17/172 | GRÜNE | MISMATCH parsed=0 db_min=1 | db_anfrager="Schäffer, Verena" | block="…"
```

Am Ende eines Laufs steht zusätzlich `plausi_mismatch=N` in der Status-Zeile. Diese Fälle **soll der Skript nicht selbst zu reparieren versuchen** — sie sind Edge-Cases, die ein Agent gezielt durchgeht. Workflow:

1. **Fälle holen:** `grep MISMATCH data/anfrager_novelty.log` (oder gefiltert nach `parsed=0`, einer KA-Nr usw.).
2. **Antwort-PDF lesen:** `df.loc[df.Drucksache_Antwort_Nr == "X/Y", "Antworttext"]` → `.md` öffnen, ersten Bildschirm anschauen — der Anfragerblock steht direkt unter „Antwort der Landesregierung auf die Kleine Anfrage … vom …".
3. **Fehlerursache identifizieren** (siehe Tabelle unten) und entweder: 
   - **direkt korrigieren**: `Anfrager_Alle` und `Anzahl_Abgeordnete` für die Zeile manuell setzen (z. B. via `openpyxl` oder kurzer Python-Patch). Originalspalte `Anfrager` unangetastet lassen.
   - oder **als ungeklärt markieren**: füge ein Token wie `anfrager_unverified` zur Spalte `Extract_Flags` der Zeile (komma-separiert), damit Auswertungen die Zeile ausschließen können.

Bekannte Edge-Case-Quellen (PDF-Header → Wirklichkeit):

| PDF-Symptom | Bedeutung | Schnell-Korrektur |
|---|---|---|
| `der AbgeordnetenVorname Nachname` (kein Leerzeichen) | pdftotext-Whitespace-Loss am Wortende | bereits regex-toleriert |
| `VornameNachname` / `undVorname` | Whitespace-Loss innerhalb / vor Name | bereits via Lower-Upper-Heuristik gefixt; Reste manuell |
| `Volker Baran` statt `Volkan Baran`, `Anrdt Klocke` statt `Arndt`, `Sarah Philip` statt `Philipp` | echter Buchstaben-Tippfehler im PDF | manuell den DB-Namen setzen |
| `Markus Pretzell` (PDF) vs. `Marcus Pretzell` (DB) | unterschiedliche Schreibweise | manuell gegen offizielle MdL-Liste prüfen, dominante Form übernehmen |
| 2 `der Abgeordneten`-Blöcke, einer pro Fraktion (z. B. `Nic Vogel AfD` + `und Frank Neppe FRAKTIONSLOS`) | Cross-Fraktion-Co-Signer (selten, aber legitim) | beide Namen manuell zusammenführen; Fraktion der Zeile = Fraktion der Anfrage-stellenden Mehrheit |
| Block leer / `Antwort: Unterrichtung Präs` | Anfrage zurückgezogen | erwartet — Status `anfrage_zurueckgezogen`; kein Mismatch-Eintrag nötig |
| DB-Anfrager enthält Komma + Fraktion-Token (`Lisa-Kristin SPD , Dr. Pfeil, Werner`) | Crawl-Fehler weiter oben | siehe `verify` / `resolve` — mit `--counter-search` gegenprüfen, bei Bedarf Zeile mit `resolve` neu ziehen |

**Markieren statt fixen:** Wenn Zweifel bestehen, lieber ein `Extract_Flags`-Token setzen (z. B. `anfrager_unverified`, `anfrager_pdf_typo`) als raten. Auswertungen können dann gezielt filtern. Die Mismatch-Logs sind das Inventar — Ziel ist nicht 100% automatischer Match, sondern transparent dokumentierter Restbestand.

**Re-Run-Schutz:** Nach einer manuellen Korrektur den Token `anfrager_manual` zu `Extract_Flags` hinzufügen. `extract-all-anfrager` skipt solche Zeilen — die Korrektur überlebt damit jeden Re-Run der Pipeline.

**Bulk-Remediation:** `tools/fix_anfrager_mismatches.py` arbeitet das aktuelle Mismatch-Log batch-mäßig ab — versucht (a) erweiterte Whitespace-Heuristik, (b) Cross-Fraktion-Matching und (c) DB-Fallback (DB-Werte übernehmen, wenn DB ohne `u.a.` ist). Setzt `anfrager_manual` (+ ggf. `anfrager_from_db`, `anfrager_cross_fraktion`) bei Erfolg, sonst `anfrager_unverified`. Idempotent — überspringt bereits behandelte Zeilen.

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
- **negative Antwortzeit**: Antwortdatum strikt vor Anfragedatum. Selten — entweder Tippfehler im PDF (z. B. Anfragedatum mit falschem Jahr), Crawl-Fehlmatch oder vertauschte Datumsspalten. 0-Tage-Antworten (gleicher Tag) sind nicht beanstandet, weil legitim möglich (z. B. dringliche Anfragen).

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
df[df.Fraktion == "AfD"]                                              # by Fraktion
df[df.Anfrager_Alle.str.contains("Wagner", na=False)]                 # by Abgeordnete (FULL list)
df[df.Schlagworte.str.contains("Polizei", na=False)]                  # by Schlagwort
df[df.Ministerium_Kuerzel == "IM"]                                    # by Kürzel (federführend)
df[df.Beteiligte_Ministerien_Kuerzel.fillna("").str.contains("IM")]   # by Kürzel (auch beteiligt)
df[df.Anfragedatum >= "2024-01-01"]                                   # by Datum
df[df.Anzahl_Abgeordnete >= 5]                                        # Sammelanfragen
```

`Anfrager` (DB-Spalte) bleibt für Rückwärtskompatibilität erhalten und enthält maximal 2 Namen + `u.a.`. **Für Auswertungen `Anfrager_Alle` nutzen** — das ist der vollständige Set inkl. aller Co-Anfrager.

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
- `data/db_index.xlsx` von Hand editieren oder mit anderen Verben überschreiben — das ist der immutable DB-Snapshot, ausschließlich Domain von `crawl`.
- Spalten `Antworttext`, `Antworttext_Status`, `Antworttext_Quelle` von Hand editieren — `fetch-text` Domain.
- Spalten `Anfrager_Alle`, `Anzahl_Abgeordnete`, `Beteiligte_Ministerien_Kuerzel` von Hand editieren — `extract-all-anfrager` / `extract-multi-ministerium` Domain. Beide sind idempotent und werden bei Re-Run überschrieben.
- Aus `Anfrager` (DB-Spalte) auf die tatsächliche Anzahl Mit-Anfragender schließen — die ist auf 2 gekappt + `u.a.`. Immer `Anzahl_Abgeordnete` / `Anfrager_Alle` nutzen.
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
