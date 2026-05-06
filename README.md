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

Primärquelle ist die [Anfragen-Datenbank des Landtags](https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten.html); sie enthält die wesentlichen Informationen zum Wer, was und wann. In dieser Datenbank wird auf die parlamentarischen Dokumente verlinkt - PDFs, die kleinere Tippfehler und Widersprüche enthalten, die im Index weitgehend bereinigt sind.

Wie jede Datenbank enthält die NRW-Landtagsdatenbank kleine Fehler - resultierend aus Tipp- und Zuordnungsfehlern und unklaren Bezugspunkten. **Insgesamt ist die Datenqualität aber gut** - die Datenbank mit dem Dokument-Index verdient am meisten Vertrauen (und wird deshalb im ersten Schritt in eine Tabelle `data/db_index.xlsx` geklont).

Diese kleineren Probleme sind aufgefallen:

- Es ist nicht ganz klar, was die Datenbank als Antwortdatum auflistet - streng genommen ist das das Antwortdatum des jeweiligen Ministeriums ("...mit Schreiben vom x.y." aus den PDFs).
- Der Index enthält maximal zwei Abgeordnete hinter einer parlamentarischen Anfrage; tatsächlich können in einer solchen Anfrage aber deutlich mehr Abgeordnete genannt sein (bei der SPD einmal 69). Das wird in der Datenbank durch "... u.a." gekennzeichnet und muss aus den PDFs ergänzt werden (Spalte `Anfrager_Alle`).
- Es wird immer nur das federführende Ministerium genannt, keine weiteren, die an der Antwort beteiligt waren (Spalte `Beteiligte_Ministerien_Kuerzel` ergänzt das aus den PDFs).
- In mindestens zwei Fällen stehen in der Datenbank Links auf Dokumente, die es doppelt gibt - unter unterschiedlichen Dokumentnummern.

All dies konnte der Agent selbsttätig auflösen.

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