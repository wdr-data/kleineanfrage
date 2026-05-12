---
name: landtag-nrw-extraction
description: Use when extracting or querying Kleine Anfrage data from Landtag NRW — metadata + answer-text harvesting into Excel + Markdown.
---

# Landtag NRW — Kleine Anfrage extraction

Idempotente CLI-Verben in `landtag.py` über drei XLSX-Ebenen plus PDF-Archiv:

| Datei | Inhalt | Domain (nur ein Schreiber) |
|---|---|---|
| `data/db_index.xlsx` | immutable DB-Snapshot der Landtag-Suche (Anfrager, Fraktion, Titel, Systematik, Ministerium, DB-Daten). Verifikations-Referenz. | `crawl` |
| `data/datum_original.xlsx` | PDF-Datums-Quelle: `Datum_Original` (Seite-1-Footer „Datum des Originals: DD.MM.YYYY") + `Datum_Ausgegeben` pro Drucksache (Anfrage und Antwort). | `tools/extract_datum_original.py` |
| `data/index.xlsx` | Working-File: DB-Metadaten + PDF-Werte + Qualitäts-Spalten. Endprodukt für Auswertungen. | alle Pipeline-Verben |

PDFs + pdftotext-Cache in `Archiv/.../MMD<wp>-<nr>.{pdf,md}`. Stammdaten (Fraktionen, Ministerien, Abgeordnete) in `Index/` — siehe `vocabulary.md`.

## Datums-Modell

| Spalte in `index.xlsx` | Primärquelle | Fallback |
|---|---|---|
| `Anfragedatum` (canonical) | **Antwort-PDF-Body**, Zeile „auf die Kleine Anfrage X **vom `<Datum>`**" — parlamentarisch anerkannte Lesart | Anfrage-PDF-Footer, wenn Body Typo hat |
| `Antwortdatum` (canonical) | **Antwort-PDF-Footer**, „Datum des Originals: …" (Briefdatum Minister) | DB, wenn PDF-Glitch |
| `Anfragedatum_DB`, `Antwortdatum_DB` | DB-Drucksachen-Daten (preserved für Verifikation) | — |

`merge` rekonziliiert mit Korridor [-14d, +122d] (PDF früher = Brief-vs-Drucksachen-Lag, DB früher = Registrierungs-Lag). Außerhalb → `Mismatch_Flag`, `Datenqualität=ask_review`.

**Drei Quellen für Anfragedatum, automatische Plausibilitäts-Auflösung**: Body (Antwort-PDF), Anfrage-PDF-Footer („Datum des Originals" der Anfrage), DB. `resolve --auto` setzt:

- Wenn **Body matcht DB** im Korridor: Body bleibt, Footer-Tippfehler wird notiert.
- Wenn **Footer matcht DB** im Korridor und Body weicht ab: Body hat Tippfehler → `Anfragedatum` wird auf Footer-Wert gesetzt.
- Wenn **Body ≈ N×365d** von DB UND Footer entfernt (N=1–5): Jahr-Tippfehler im Body → rollback auf DB.
- Wenn **Body-Jahr außerhalb [2015,2030]**: OCR-Glitch → DB-Fallback (durch merge bereits gesetzt).
- Restliche Fälle bleiben `ask_review` oder `review` für manuelle Sichtung.

**Cross-Flag**: `anfragedatum_pdf_kreuz` (Body ↔ Footer diff > 14d) wird zusammen mit `anfragedatum`-Flag analysiert — siehe Heuristik-Tabelle unten.

**Manuelle Adoption**: Tag `anfragedatum_manual` / `antwortdatum_manual` in `Extract_Flags` sperrt die Reconcile für die markierte Spalte. Der adoptierte Wert bleibt bei künftigen merges erhalten.

## Datenqualität

`merge` setzt pro Zeile:

- `Mismatch_Flags` — `,`-getrennte Felder mit DB↔PDF-Konflikt (`anfragedatum`, `antwortdatum`, `anfragedatum_pdf_kreuz`, `drucksache_anfrage_nr`, `fraktion`, `md_kanr`).
- `Datenqualität` — eines von:
  - `ok` — alles glatt
  - `korrigiert` — manueller Tag oder existierende Notiz (sticky)
  - `ask_review` — offener Mismatch ohne Notiz (Domain von `resolve`)
  - `review` — sticky manuelle Priorität (z. B. nach Diff gegen externe Korrektur-Variante); überschreibt korrigiert/ask_review bis Mensch sie räumt
- `Notizen` — Free-Text, Domain von `resolve`.

`resolve --auto` schließt 100 % der auto-klassifizierbaren ask_review-Zeilen anhand bekannter Heuristiken:

| Mismatch-Klasse | Verdict |
|---|---|
| Korridor (PDF früher 0–122d ODER DB früher 0–14d) | silent PDF, kein Flag (greift in `merge` direkt) |
| **Body-Year-Glitch** (`Anfragedatum`-Body Jahr außerhalb [2015,2030]) | DB-Fallback (durch merge), Notiz dokumentiert OCR-Glitch |
| **Body-Year-Typo** (`Anfragedatum`-Body ≈365d von DB UND Anfrage-Footer entfernt) | rollback Body → DB |
| **Body-Multi-Year-Typo** (`Anfragedatum`-Body ≈ N×365d von DB UND Footer entfernt, N=2–5) | rollback Body → DB |
| **Body-Tippfehler** (Anfrage-PDF-Footer matcht DB im Korridor, Body weicht > 14d ab) | adoptiere Footer (rec.anfragedatum = Footer) |
| **Footer-Typo** (Body matcht DB, Anfrage-Footer >14d entfernt) | Body bleibt, Notiz dokumentiert Footer-Tippfehler |
| Year-Tippfehler oder OCR-Glitch im Antwort-Footer (`antwortdatum` Jahr ≈365d off oder außerhalb [2015,2030]) | DB maßgeblich, Notiz |
| Drucksachen-Nr-Mismatch (Off-by-N, KA-Nr-Konflikt, WP-falsch) | DB maßgeblich, Notiz |
| Fraktion AfD ↔ fraktionslos | Fraktionsaustritt zwischen Anfrage und Antwort — beide legitim |
| md_kanr-Parser-Artefakt (Spacing, Doubled-Render) | Parser-Quirk, DB maßgeblich |
| `.md`-Inhalt ist die Anfrage statt der Antwort | CDN-Routing-Bug → `Antworttext_Status=pdf_missing`, Notiz |

**Notiz-Format für Anfragedatum-Korrekturen**: jede Notiz listet alle drei Quellen — `Quellen: DB='…' | Antwort-Body='…' | Anfrage-PDF-Footer='…'` — plus erklärenden Verdict-Satz. Dadurch ist die Audit-Spur vollständig.

Edge-Case-Tabelle: `edge-cases.md`. Variantenschreibweisen: `vocabulary.md`.

`extract_datum_original` heilt drei pdftotext-Quirks vor dem Schreiben in `datum_original.xlsx`:

- Punkt-als-Leerzeichen-Render (`03 07.2025` → `03.07.2025`).
- 4-stelliges Year-Glitch (`0205` → `2025`, via Ausgegeben-Jahr).
- Year-Typo Original-vs-Ausgegeben (Original ≈1 Jahr vor Ausgegeben → +1 Jahr auf Original).

## When to use

- **Wer/Was/Wann** einer KA → `data/index.xlsx`.
- **Alle Mit-Anfragenden** → Spalte `Anfrager_Alle` (DB-Spalte `Anfrager` kappt bei 2 Namen + „u.a.").
- **Alle beteiligten Ministerien** → `Beteiligte_Ministerien_Kuerzel`.
- **Antworttext** → Pfad in Spalte `Antworttext` (→ `Archiv/.../MMD<wp>-<nr>.md`).
- **PDF-Datum-Verifikation** → `data/datum_original.xlsx`.
- **Quantitative Auswertung** → pandas, siehe „Read access" unten.
- **Zweifelhafte Zeile** → `Datenqualität ∈ {ask_review, review}` filtern, dann `resolve --auto` (Heuristik) oder gezielter `resolve --ka N`.

## Pipeline

```sh
python landtag.py crawl                     --wahlperiode 18  # DB-Snapshot (~25 min)
python landtag.py fetch-text                --wahlperiode 18  # Antwort-PDFs + .md
python landtag.py scan-archive              --wahlperiode 18  # Titel + Ministerium-Prosa
python landtag.py extract-multi-ministerium                   # beteiligte Ressorts
python landtag.py build-abgeordnete-index                     # MdL-Index aus DB-Anfrager
python landtag.py extract-all-anfrager                        # Anfrager_Alle + Anzahl_Abgeordnete
python landtag.py build-abgeordnete-index                     # 2. Pass: absorbiert 3.+ Anfrager
python landtag.py extract-all-anfrager                        # 2. Pass: matcht erweiterten Namensindex
python tools/extract_datum_original.py                        # PDF-Briefdaten (Anfrage & Antwort, lädt Anfrage-PDFs nach)
python landtag.py merge                                       # Mismatch_Flags + Datenqualität
python landtag.py resolve --auto                              # Heuristik-Auflösung
python landtag.py verify                                      # sanity report
```

`crawl` liefert pro KA: beide Drucksachen + Links, DB-Daten, Anfrager (max 2 + „u.a."), Fraktion, Titel, Systematik, Schlagworte, federführendes Ministerium-Kürzel. Der Rest kommt aus den PDFs.

`normalize` (Fraktion/Ministerium → kanonisch) läuft automatisch am Ende von `crawl` und `fetch-text`. `enrich-llm` ist optional für Rest-Lücken.

Mehrere Wahlperioden in derselben `index.xlsx` möglich — jede Zeile trägt ihre `WP`. Date-bounded subset:
```sh
python landtag.py crawl --wahlperiode 18 --from 2024-01-01 --to 2024-12-31
```

Gezielte Direkt-Reparatur einzelner KAs (Phantom, md↔KA-Mismatch, Korrigendum):
```sh
python landtag.py resolve --ka 144 --counter-search --delete-phantom
```

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
df[df.Datenqualität.isin(["ok", "korrigiert"])]                       # produktionsreife Zeilen
df[df.Datenqualität == "review"]                                      # offen für menschliche Sichtung
df[df.Antworttext_Status == "pdf_missing"]                            # Antwort-PDF nicht auf Landtag-CDN
```

`Anfrager` (DB-Spalte) max 2 Namen + „u.a." — **für Auswertungen `Anfrager_Alle` nutzen.**

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

- `data/db_index.xlsx` von Hand editieren — immutable Snapshot, ausschließlich Domain von `crawl`.
- `data/datum_original.xlsx` von Hand editieren — Domain von `tools/extract_datum_original.py`. Bei Fehlerverdacht (no_match / parse_failed) nur die betroffenen Zeilen löschen und re-extrahieren.
- `Anfragedatum` / `Antwortdatum` / `*_DB` in `index.xlsx` von Hand setzen — werden bei `merge` neu befüllt. Manuelle Adoption nur über Tag `anfragedatum_manual` / `antwortdatum_manual` in `Extract_Flags` (sperrt das Feld gegen Reconcile).
- `Antworttext`, `Antworttext_Status`, `Antworttext_Quelle` von Hand editieren — `fetch-text`-Domain.
- `Anfrager_Alle`, `Anzahl_Abgeordnete`, `Beteiligte_Ministerien_Kuerzel` von Hand editieren — werden bei Re-Run überschrieben. Korrektur mit `anfrager_manual`-Tag in `Extract_Flags` schützen.
- Zwei Verben gleichzeitig auf dieselbe XLSX (File-Lock blockt, ist aber unsauber).
- Aus DB-Spalte `Anfrager` auf echte Anzahl Mit-Anfragender schließen.
- Werte aus `data/vocab_novelty.log` automatisch absorbieren — siehe `vocabulary.md`.
- `--rps` über 4 ohne Grund.
- `Antworttext_Status=pdf_missing` rückgängig machen, ohne den CDN-Bug zu verifizieren — `fetch-text` würde dasselbe falsche PDF wieder laden. Re-Try nur mit `fetch-text --force` und manueller Prüfung.

## Reference

- `edge-cases.md` — Anfrager-Quirks, PDF-Tippfehler, Status-Edge-Cases, Crawl-Quirks.
- `vocabulary.md` — Stammdaten, Namens-/Format-Konventionen, Fraktion- und Min-Aliase.
- `docs/superpowers/specs/2026-05-01-landtag-nrw-extraction-design.md` — Architektur-Spec.
