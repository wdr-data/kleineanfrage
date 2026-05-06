# wdr-kleineanfrage

Strukturierte Daten zu **Kleinen Anfragen** des Landtags Nordrhein-Westfalen lokal extrahieren — Metadaten in einer Excel-Datei, Antworttexte als Markdown neben den Original-PDFs.

Wer hat wann was gefragt - und welches Ministerium hat wie schnell geantwortet? Die Daten aus der Datenbank - und die Verweise auf die Volltexte - landen in einer Excel-Tabelle (`data/index.xlsx`), die eine weitere Auswertung erlaubt.

## Setup

```sh
pip install -r requirements.txt
brew install poppler          # liefert pdftotext (auf macOS)
```

## Wie man den Skill benutzt

- Repository klonen, requirements installieren.
- Anstatt selbst `python landtag.py` aufzurufen: Claude Code im Projekt-Verzeichnis starten.
- "Lade Skill landtag-nrw-extraction" - der Agent liest `skills/landtag-nrw-extraction/SKILL.md` und kennt damit die Verben.
- Agenten anweisen, Daten und PDFs zu holen (~20 min je Wahlperiode), anschließend zu verifizieren und Widersprüche aufzulösen. Bei Unklarheiten fragt er nach.

Pipeline, Verb-Liste, pandas-Snippets und Resolve-Heuristik stehen in **`skills/landtag-nrw-extraction/SKILL.md`**.

## Datenqualität

Primärquelle ist die [Anfragen-Datenbank des Landtags](https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html); sie enthält die wesentlichen Informationen zum Wer, was und wann. In dieser Datenbank wird auf die parlamentarischen Dokumente verlinkt - PDFs, die kleinere Tippfehler und Widersprüche enthalten, die im Index weitgehend bereinigt sind.

Wie jede Datenbank enthält die NRW-Landtagsdatenbank kleine Fehler - resultierend aus Tipp- und Zuordnungsfehlern und unklaren Bezugspunkten. **Insgesamt ist die Datenqualität aber gut** - die Datenbank mit dem Dokument-Index verdient am meisten Vertrauen (und wird deshalb im ersten Schritt in eine Tabelle `data/db_index.xlsx` geklont).

Diese kleineren Probleme sind aufgefallen:

- Es ist nicht ganz klar, was die Datenbank als Antwortdatum auflistet - streng genommen ist das das Antwortdatum des jeweiligen Ministeriums ("...mit Schreiben vom x.y." aus den PDFs).
- Der Index enthält maximal zwei Abgeordnete hinter einer parlamentarischen Anfrage; tatsächlich können in einer solchen Anfrage aber deutlich mehr Abgeordnete genannt sein (bei der SPD einmal 69). Das wird in der Datenbank durch "... u.a." gekennzeichnet und muss aus den PDFs ergänzt werden (Spalte `Anfrager_Alle`).
- Es wird immer nur das federführende Ministerium genannt, keine weiteren, die an der Antwort beteiligt waren (Spalte `Beteiligte_Ministerien_Kuerzel` ergänzt das aus den PDFs).
- In mindestens zwei Fällen stehen in der Datenbank Links auf Dokumente, die es doppelt gibt - unter unterschiedlichen Dokumentnummern.

## Auswertung in R

Im Ordner `R` findet sich das Skript für die Auswertung der Tabelle. Es nimmt die `data/index.xlsx` und pivotiert/summiert auf - erzeugt für beide Wahlperioden 17 und 18 jeweils eine Reihe von Kreuztabellen, die Anzahl und Bearbeitungszeit der Anfragen aufgeschlüsselt nach Fraktionen, Ministerien, Abgeordneten. Außerdem werden die Themen-Tags aus der Rubrik "Systematik" ausgewertet - über den gesamten Zeitraum hinweg werden Tabellen mit den meist verwendeten Tags erzeugt.

## Wichtige Dateien

```
landtag.py                                   ← gesamte Logik (eine Datei)
requirements.txt
Archiv/                                      ← PDF-Cache (~18.000 Dateien, gitignored)
data/db_index.xlsx                           ← immutable DB-Snapshot (nur von `crawl` geschrieben)
data/index.xlsx                              ← Working-File: DB + PDF-Anreicherung + Datenqualität
data/*.log                                   ← Extraktions-/Crawl-Fehler, Vokabel-Neulinge
skills/landtag-nrw-extraction/SKILL.md       ← Skill-Definition für Agenten (Kern)
skills/landtag-nrw-extraction/edge-cases.md  ← Quirks, PDF-Tippfehler, Sonderfälle
skills/landtag-nrw-extraction/vocabulary.md  ← Stammdaten, Format-Konventionen, Aliase
CLAUDE.md                                    ← Agenten-Briefing (Projektkontext, Designprinzipien)
```
