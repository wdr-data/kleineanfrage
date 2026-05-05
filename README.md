# wdr-kleineanfrage

Strukturierte Daten zu **Kleinen Anfragen** des Landtags Nordrhein-Westfalen lokal extrahieren — Metadaten in einer Excel-Datei, Antworttexte als Markdown neben den Original-PDFs.

Wer hat wann was gefragt - und welches Ministerium hat wie schnell geantwortet? Die Daten aus der Datenbank - und die Verweise auf die Volltexte - landen in einer Excel-Tabelle (`index.xlsx`), die eine weitere Auswertung erlaubt.

## Auswertung in R

Im Ordner `R` findet sich das einfache Skript für die Auswertung der Tabelle.


## Der Python-Code

```
landtag.py                                   ← die gesamte Logik (eine Datei, ~750 Zeilen)
requirements.txt
Archiv/                                      ← PDF-Cache (~18.000 Dateien, gitignored)
data/index.xlsx                              ← eine Zeile pro Kleiner Anfrage
data/*.log                                   ← Extraktions-/Crawl-Fehler, Vokabel-Neulinge
docs/superpowers/specs/2026-05-01-…design.md ← vollständiger Entwurf
skills/landtag-nrw-extraction/SKILL.md       ← Skill-Definition für Agenten
CLAUDE.md                                    ← Agenten-Briefing
```

---

### Einmaliges Setup

```sh
pip install -r requirements.txt
brew install poppler          # liefert pdftotext (auf macOS)
```

## Der Agenten-Skill

### Die Verben

Alle Verben sind **idempotent** und teilen sich `data/index.xlsx` über einen File-Lock. Die Online-Suche (`crawl`) ist Canon — sie bestimmt, welche KAs der Wahlperiode existieren. `scan-archive` und `fetch-text` reichern den so aufgebauten Index an.

| # | Verb | Was es tut | Netz? |
|---|---|---|---|
| 1 | `crawl` | Spring-Webflow-Handshake gegen die Suchseite; paginiert die ganze Wahlperiode. Liefert Zeilen mit KA-Nr, Anfrager, Fraktion, beide Drucksachen + Daten, Ministeriums-Kürzel, Systematik, Schlagworte, Titel. **Canon für alle Felder, die der Such-Treffer enthält.** | ja |
| 2 | `fetch-text` | Lädt fehlende Antwort-PDFs herunter (überspringt alles im `Archiv/`) und extrahiert den Volltext mit `pdfplumber` nach `.md`. Skipt zurückgezogene und unbeantwortete KAs. | ja |
| 3 | `scan-archive` | Reichert vorhandene Index-Zeilen aus den lokalen Antwort-PDFs an (Anfragetitel, Ministerium-Prosaname). Filtert dabei Anfrage-PDFs und Große-Anfragen-Antworten heraus. **Default-Modus = enrich-only**: schreibt nie über Crawl-Werte. `--allow-discovery` resurrektiert das alte Verhalten (ohne Crawl). | nein |
| 4 | `normalize` | Matcht `Fraktion` / `Ministerium` gegen `Index/*.xlsx`, füllt `*_Canonical` und `Ministerium_Kuerzel`. Wird auch automatisch am Ende von `crawl` und `fetch-text` aufgerufen. | nein |
| 5 | `resolve` | Gezielte Re-Suche für eine oder mehrere KA-Nrn (z. B. nach Verdachtsfällen aus `verify`). `--counter-search` macht eine Gegensuche mit `doktyp=GA` (Große Anfrage), `--delete-phantom` löscht stale Zeilen mit anderer Antwort-Drucksache. | ja |
| 6 | `extract-multi-ministerium` | Parst pro Antwort-`.md` den Boundary-Absatz „Der Minister … hat die Kleine Anfrage … beantwortet." und schreibt das vollständige Set beteiligter Ressorts (federführend zuerst) nach `Beteiligte_Ministerien_Kuerzel`. Idempotent. WP18 abgedeckt; WP17 hat anderen Ministeriumszuschnitt → out of scope. | nein |
| 7 | `verify` | Read-only Sanity-Report: Status-Zählungen, KA-Nr-Lücken pro WP, md ↔ KA-Nr-Konsistenzcheck, Duplikat-KA-Nrn, Waisen-PDFs. Optional LLM-Plausibilität via `--llm-*`. | nein |

### Standardlauf

```sh
python landtag.py crawl                       --wahlperiode 18  # ~25 min bei --rps 4
python landtag.py fetch-text                  --wahlperiode 18  # nur fehlende PDFs (~5 min)
python landtag.py scan-archive                --wahlperiode 18  # ~10 min, Min/Titel anreichern
python landtag.py extract-multi-ministerium                     # Beteiligte Ressorts pro Antwort
python landtag.py verify
```

Ein zweites Wahlperiode-Set (z. B. WP17) wird mit denselben Verben in den gleichen Index gemerged; jede Zeile trägt ihre `WP`-Spalte.

### Datumsfilter (clientseitig nach Fetch)

```sh
python landtag.py crawl --wahlperiode 18 --from 2024-01-01 --to 2024-12-31
```

### Verdachtsfälle reparieren

```sh
# md ↔ KA-Nr-Mismatch oder ein Loch in der KA-Nummerierung — gezielt nachfragen:
python landtag.py resolve --ka 144 --ka 3167 --counter-search --delete-phantom
```

---

## Nutzung in Claude Code / anderen Agenten

### Wann dieses Tool greifen sollte

- Frage zu **Wer/Was/Wann** einer Kleinen Anfrage (Anfrager, Fraktion, Datum, Ministerium, Schlagworte) → **`data/index.xlsx`** abfragen.
- Frage zum **Inhalt einer konkreten Antwort** → **`Archiv/.../MMD18-N.md`** lesen. Nur auf die PDF zurückfallen, wenn die `.md` fehlt.
- **Quantitative Auswertung** über den Datenbestand (z. B. „welche Abgeordneten stellen die meisten KAs", „durchschnittliche Antwortdauer") → mit pandas auf der XLSX arbeiten.
- Wunsch nach **frischen Daten** → den Refresh-Loop oben anstoßen (vorher User informieren, dass `crawl` ~1 h läuft).

### Lesen mit pandas

```python
import pandas as pd
df = pd.read_excel("data/index.xlsx")

df[df.Fraktion == "AfD"]                                # nach Fraktion
df[df.Anfrager.str.contains("Wagner", na=False)]        # nach Abgeordnete:r
df[df.Schlagworte.str.contains("Polizei", na=False)]   # nach Schlagwort
df[df.Anfragedatum >= "2024-01-01"]                     # nach Datum
df[df.Ministerium_Kuerzel == "JM"]                      # nach federführendem Ressort
df[df.Beteiligte_Ministerien_Kuerzel.fillna("").str.contains("IM")]   # nach beteiligtem Ressort (auch nicht-federführend)
```

`Ministerium_Kuerzel` = federführendes Ressort aus dem Search-Hit.
`Beteiligte_Ministerien_Kuerzel` = comma-separierte Kürzel-Liste **aller** im Boundary-Absatz genannten Ressorts (federführend zuerst), gefüllt von `extract-multi-ministerium`.

Antworttext einer Zeile holen:

```python
md_path = df.loc[df.Drucksache_Antwort_Nr == "18/1006", "Antworttext"].iloc[0]
print(open(md_path, encoding="utf-8").read())
```

CSV-Export, falls ein externes Tool das braucht:

```sh
python -c "import pandas as pd; pd.read_excel('data/index.xlsx').to_csv('out.csv', index=False)"
```