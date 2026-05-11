# wdr-kleineanfrage

Strukturierte Daten zu **Kleinen Anfragen** des Landtags Nordrhein-Westfalen lokal extrahieren — Metadaten in einer Excel-Datei, Antworttexte als Markdown neben den Original-PDFs.

Wer hat wann was gefragt - und welches Ministerium hat wie schnell geantwortet? Die Daten aus der Datenbank - und die Verweise auf die Volltexte - landen in einer Excel-Tabelle (`data/index.xlsx`), die eine weitere Auswertung erlaubt.

In diesem Repository findet sich:

- Ein Tool-Skript, um Daten aus der Datenbank zu lesen und aus den PDFs zu extrahieren
- Ein "Skill", mit der ein KI-Agent dieses Tool für die Userin benutzen kann
- Ein R-Skript zur Auswertung der gefundenen Daten

## Was ein Skill ist

Ein "Skill" ist ein Paket für einen KI-Agenten wie Claude Code, mit dem er eine Aufgabe lösen können soll:

- **Werkzeuge**, die die KI einsetzen kann, um die Aufgabe zu lösen
- **prozedurales und Hintergrundwissen** - gewissermaßen die Betriebsanleitung zu den Tools.

Einen Skill kann man zu Beginn einer Claude-Code-Session wie ein Zusatzmodul laden - es ist im Prinzip eine Art langer System Prompt, der das Sprachmodell auf seine Aufgabe vorbereitet. In unserem Fall: saubere Daten über Kleine Anfragen in NRW gewinnen.

## Wie man den Skill benutzt

- Repository klonen, requirements installieren.
- Anstatt selbst `python landtag.py` aufzurufen: Claude Code im Projekt-Verzeichnis starten.
- "Lade Skill landtag-nrw-extraction" - der Agent liest `skills/landtag-nrw-extraction/SKILL.md` und kennt damit die Verben.
- Agenten anweisen, Daten und PDFs zu holen (~20 min je Wahlperiode), anschließend zu verifizieren und Widersprüche aufzulösen. Bei Unklarheiten fragt er nach.

Pipeline, Verb-Liste, pandas-Snippets und Resolve-Heuristik stehen in **`skills/landtag-nrw-extraction/SKILL.md`**.

## Datenqualität

Primärquelle ist die [Anfragen-Datenbank des Landtags](https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html); sie enthält die wesentlichen Informationen zum Wer, Was und Wann und verlinkt auf die parlamentarischen Dokumente. Die DB liefert den größeren Teil der Metadaten; die PDFs sind die Quelle für alles, was die DB kappt oder nicht enthält — sie liefern auch die maßgeblichen Daten zum Anfrage- und Antwortzeitpunkt (siehe „Datenmodell" unten).

Wie jede Datenbank enthält die NRW-Landtagsdatenbank kleine Fehler — Tipp- und Zuordnungsfehler, unklare Bezugspunkte. **Insgesamt ist die Datenqualität gut**; die DB wird im ersten Schritt als `data/db_index.xlsx` eingefroren und bleibt Referenz.

Diese typischen Lücken kennt der Skill und schließt sie aus den PDFs:

- Der Index nennt maximal zwei Abgeordnete pro Anfrage; tatsächlich können viel mehr beteiligt sein (bei der SPD einmal 69). Markiert mit „… u.a.", ergänzt in Spalte `Anfrager_Alle`.
- Nur das federführende Ministerium steht im Index; die weiteren beteiligten Ressorts kommen aus dem PDF-Antworttext (`Beteiligte_Ministerien_Kuerzel`).
- Anfrage- und Antwortdatum: die DB führt das Drucksachen-Veröffentlichungsdatum, das PDF das tatsächliche Briefdatum („Datum des Originals: …" auf Seite 1). Letzteres ist die maßgebliche Größe — die Pipeline schreibt es in `Anfragedatum`/`Antwortdatum` und hält das DB-Datum in `Anfragedatum_DB`/`Antwortdatum_DB` zur Verifikation fest.
- In mindestens zwei Fällen verlinkt die DB auf Dokumente, die doppelt unter unterschiedlichen Drucksachen-Nummern existieren.

All dies löst der Agent selbsttätig auf.

## Datenmodell — drei Quellen, eine Zieltabelle

| Datei | Inhalt | Domain |
|---|---|---|
| `data/db_index.xlsx` | eingefrorener Snapshot der Landtagsdatenbank-Suche (Anfrager, Fraktion, Titel, Systematik, Ministerium, Drucksachen-Daten). Verifikations-Referenz. | nur `landtag.py crawl` |
| `data/datum_original.xlsx` | PDF-Briefdaten („Datum des Originals: DD.MM.YYYY") + „Ausgegeben"-Datum pro Drucksache, sowohl Anfrage als auch Antwort. **Autoritative Datums-Quelle.** | nur `tools/extract_datum_original.py` |
| `data/index.xlsx` | Arbeits-Tabelle: DB-Metadaten + PDF-Datumswerte + PDF-Anreicherungen + Qualitäts-Spalten. Antwort an Auswertungs-Tools. | alle Pipeline-Verben |

### Wie das Datum geprüft wird

PDF-Briefdatum und DB-Drucksachen-Datum unterscheiden sich fast immer — die Drucksache wird typischerweise einige Tage nach dem Schreiben veröffentlicht. `landtag.py merge` joinst beide Werte und entscheidet:

- **Korridor** (0–122 Tage, PDF früher, beide Jahre 2015–2026): erwartbarer Verarbeitungs-Lag → PDF-Datum wird stillschweigend in `index.xlsx` übernommen, keine Flag.
- **Außerhalb Korridor** (große Differenz, Jahres-Tippfehler, OCR-Glitch im PDF): Zeile wird mit `Datenqualität=ask_review` markiert. `resolve --auto` arbeitet eine Heuristik-Tabelle ab — in der Regel gewinnt das PDF-Datum, bei klarem PDF-Parse-Fehler (Jahr außerhalb 2015–2030 oder Jahres-Tippfehler) fällt die Heuristik auf das DB-Datum zurück. Die jeweils ersetzten Werte landen als „war (DB): …" oder „war (PDF): …" in `Notizen`.
- **PDF fehlt** (`pdf_missing`, `no_match` in `datum_original.xlsx`): DB-Datum bleibt als Fallback in `index.xlsx`.

### Pipeline in einem Satz

`crawl` → `fetch-text` → `scan-archive` → `extract-multi-ministerium` → `build-abgeordnete-index` → `extract-all-anfrager` (Doppelpass) → `extract_datum_original.py` → `merge` → `resolve --auto` → `verify`. Volle Befehlsliste in `skills/landtag-nrw-extraction/SKILL.md`.

## Der eine Eingriff von Hand...

...betrifft einen AfD-Abgeordneten, der in der 18. Wahlperiode die AfD verlassen hat und als Fraktionsloser drei Anfragen stellte. Er trat der Fraktion danach wieder bei. In der Datei `index_fixed.xlsx` sind diese drei Anfragen auf die AfD umgeschlüsselt - ansonsten ist sie ein getreues Abbild der `index.xlsx`.


## Der Python-Code

...steckt praktisch vollständig in `landtag.py`, einem Kommandozeilen-Tool, mit dem man die Daten aus der Datenbank ziehen und aus PDFs extrahieren kann. Er automatisiert die Suche in den PDFs nach Datum, beteiligten Ministerien, und Volltext.

**Das Tool ist weniger für den Einsatz durch den/die Datenjourno als zum Einsatz für die KI.** Natürlich kann man das Tool auch auf der Kommandozeile aufrufen - die Befehle finden sich in SKILL.md - aber viel einfacher ist, den Agenten das Tool selbst nutzen zu lassen: im starren Python-Code werden immer wieder Ergebnisse produziert, die man sich anschauen muss. Die KI kann das erledigen - und die häufigen Fehler selbst abräumen: Tippfehler, Zahlendreher, Verwechslungen.

Am Ende einer Session steht dann eine vollständige Index-Tabelle - mit Links auf die Primärquellen zur weiteren Auswertung z.B. auf Anfrage- oder Antworttexte hin.

## Auswertung in R

Im Ordner `R` findet sich das Skript für die Auswertung der Tabelle. Es nimmt die `data/index.xlsx` und pivotiert/summiert auf - erzeugt für beide Wahlperioden 17 und 18 jeweils eine Reihe von Kreuztabellen, die Anzahl und Bearbeitungszeit der Anfragen aufgeschlüsselt nach Fraktionen, Ministerien, Abgeordneten. Außerdem werden die Themen-Tags aus der Rubrik "Systematik" ausgewertet - über den gesamten Zeitraum hinweg werden Tabellen mit den meist verwendeten Tags erzeugt.

Diese Tabellen können dann z.B. mit Datawrapper visualisiert werden.

Die Tabellen werden ins Verzeichnis `data/` geschrieben, in Unterverzeichnisse für die jeweilige Wahlperiode, soweit zutreffend.

## Wichtige Dateien

```
landtag.py                                   ← gesamte Logik (eine Datei)
tools/extract_datum_original.py              ← liest „Datum des Originals" aus allen PDFs
requirements.txt
Archiv/                                      ← PDF-Cache (~28.000 Dateien Anfrage+Antwort, gitignored)
data/db_index.xlsx                           ← immutable DB-Snapshot (nur von `crawl` geschrieben)
data/datum_original.xlsx                     ← PDF-Briefdaten je Drucksache (autoritative Datums-Quelle)
data/index.xlsx                              ← Working-File: DB-Metadaten + PDF-Anreicherung + Datenqualität
data/*.log                                   ← Extraktions-/Crawl-Fehler, Vokabel-Neulinge
skills/landtag-nrw-extraction/SKILL.md       ← Skill-Definition für Agenten (Kern)
skills/landtag-nrw-extraction/edge-cases.md  ← Quirks, PDF-Tippfehler, Sonderfälle
skills/landtag-nrw-extraction/vocabulary.md  ← Stammdaten, Format-Konventionen, Aliase
CLAUDE.md                                    ← Agenten-Briefing (Projektkontext, Designprinzipien)
```