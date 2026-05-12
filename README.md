# wdr-kleineanfrage

Strukturierte Daten zu **Kleinen Anfragen** des Landtags Nordrhein-Westfalen lokal extrahieren — Metadaten in einer Excel-Datei, Antworttexte als Markdown neben den Original-PDFs.

**Wer hat wann was gefragt - und welches Ministerium hat wie schnell geantwortet?** Die Daten aus der Datenbank - und die Verweise auf die Volltexte - landen in einer Excel-Tabelle (`data/index.xlsx`), die eine weitere Auswertung erlaubt.

In diesem Repository findet sich:

- Ein Tool-Skript, um Daten aus der Datenbank zu lesen und aus den PDFs zu extrahieren
- Ein "Skill", mit der ein KI-Agent dieses Tool für die Userin benutzen kann
- Ein R-Skript zur Auswertung der gefundenen Daten

## Was ein Skill ist

Ein "Skill" ist ein Paket für einen KI-Agenten wie Claude Code, mit dem er eine Aufgabe lösen können soll:

- **Werkzeuge**, also Programme, die die KI einsetzen kann, um die Aufgabe zu lösen
- **prozedurales und Hintergrundwissen** - gewissermaßen die Betriebsanleitung zu den Tools.

Einen Skill kann man zu Beginn einer Claude-Code-Session wie ein Zusatzmodul laden - es ist im Prinzip eine Art langer System-Prompt, der das Sprachmodell auf seine Aufgabe vorbereitet. In unserem Fall: saubere Daten über Kleine Anfragen in NRW gewinnen.

## Wie man den Skill benutzt

- Repository klonen, requirements installieren.
- Anstatt selbst `python landtag.py` aufzurufen: Claude Code im Projekt-Verzeichnis starten.
- "Lade Skill landtag-nrw-extraction" - der Agent liest `skills/landtag-nrw-extraction/SKILL.md` und kennt damit die Verben.
- Agenten anweisen, Daten und PDFs zu holen (~20 min je Wahlperiode), anschließend zu verifizieren und Widersprüche aufzulösen. Bei Unklarheiten fragt er nach.

Pipeline, Verb-Liste, pandas-Snippets und Resolve-Heuristik stehen in **`skills/landtag-nrw-extraction/SKILL.md`**.

## Wie die Laufzeit einer Anfrage definiert wird

An sich ist es ganz einfach:

- Ein/e Abgeordnete/r stellt am Tag X eine Kleine Anfrage an die Landesregierung.
- Am Tag Y antwortet ein Ministerium, wofür es meist andere Ministerien eingebunden hat.
- Die Laufzeit der Kleinen Anfrage: Y-X <= 28 Tage. (So die gesetzliche Vorgabe.)

X und Y sind definiert als der Tag, an dem die Anfrage (bzw. die Antwort darauf) **auf der landeseigenen Datenaustausch-Plattform hochgeladen ist**.

Praktisch ist das Datum aber nicht so leicht herauszufinden: In der Datenbank und den PDFs finden sich unterschiedlichste Datumsangaben.

- Im Datenbank-Index stehen das Datum der Anfrage und das Datum der Antwort.
- Auf jedem PDF steht auf Seite 1 ein "Datum des Originals"
- Die Antworten der Ministerien beginnen immer mit einem Seitenkopf: "Antwort der Landesregierung auf die Kleine Anfrage <nr> vom <Datum>". Außerdem enthält der Seitenkopf ein weiteres Mal das Datum des Antwortdokuments.

Leider sind diese Datumsangaben nicht immer deckungsgleich - dazu im nächsten Abschnitt mehr - deswegen ist festgelegt:

- Datum der Anfrage ziehen wir aus dem Seitenkopf der Antwort
- Datum der Antwort ist das Dokumentendatum (also "Datum des Originals" unten auf Seite 1)

Die anderen Daten werden genutzt, um die Daten zu verifizieren und ggf. zu korrigieren - und das ist dringend nötig.

## Datenqualität

Wie jede Datenbank enthalten die NRW-Landtagsdatenbank und die PDFs kleine Fehler — Tipp- und Zuordnungsfehler, unklare Bezugspunkte.

**Insgesamt ist die Datenqualität gut**; die DB wird im ersten Schritt als `data/db_index.xlsx` eingefroren und bleibt Referenz.

Diese typischen Lücken kennt der Skill und schließt sie aus den PDFs:

- Wie erwähnt sind die Daten, die in der Datenbank stehen, nicht ganz richtig - sie scheinen die Veröffentlichung durch den Landtag zu markieren und sind nicht immer mit den oben genannten Anfrage- und Antwortdaten deckungsgleich.
- Der Index nennt maximal zwei Abgeordnete pro Anfrage; tatsächlich können viel mehr beteiligt sein (bei der SPD einmal 69). Markiert mit „… u.a.", ergänzt in Spalte `Anfrager_Alle`.
- Nur das federführende Ministerium steht im Index; die weiteren beteiligten Ressorts kommen aus dem PDF-Antworttext (`Beteiligte_Ministerien_Kuerzel`).
- In mindestens zwei Fällen verlinkt die DB auf Dokumente, die doppelt unter unterschiedlichen Drucksachen-Nummern existieren.

All dies löst der Agent selbsttätig auf.

## Datenmodell — drei Quellen, eine Zieltabelle

| Datei | Inhalt | Domain |
|---|---|---|
| `data/db_index.xlsx` | eingefrorener Snapshot der Landtags-DB-Suche (Anfrager, Fraktion, Titel, Systematik, Ministerium, Drucksachen-Daten). Verifikations-Referenz. | nur `landtag.py crawl` |
| `data/datum_original.xlsx` | PDF-Briefdaten („Datum des Originals: DD.MM.YYYY") + Ausgegeben-Datum je Drucksache, Anfrage und Antwort getrennt. Plausibilitäts-Referenz. | nur `tools/extract_datum_original.py` |
| `data/index.xlsx` | Arbeits-Tabelle: DB-Metadaten + canonical PDF-Daten + Anreicherungen + Qualitäts-Spalten. Eingang für R-Auswertung. | alle Pipeline-Verben |

### Wo welches Datum herkommt

Für jede Kleine Anfrage existieren bis zu drei Daten-Lesarten — die Pipeline führt die plausibelste in `index.xlsx` und hält die Verifikations-Quellen daneben:

| Spalte | Primärquelle | Bedeutung |
|---|---|---|
| `Anfragedatum` | **Antwort-PDF, Body**: „auf die Kleine Anfrage X **vom `<Datum>`**" | Der Tag, den die Landesregierung in ihrer eigenen Antwort als Anfragedatum nennt — die offiziell anerkannte Lesart |
| `Antwortdatum` | **Antwort-PDF, Footer Seite 1**: „Datum des Originals: …" | Briefdatum des Ministeriums (Unterschrift) |
| `Anfragedatum_DB`, `Antwortdatum_DB` | Drucksachen-Veröffentlichungs­datum aus DB | Verifikations-Referenz, bleibt unverändert |

Zusätzlich wird das **Anfrage-PDF-Footer-Datum** („Datum des Originals" auf dem Anfrage-Schreiben) aus `datum_original.xlsx` als dritte Quelle herangezogen — nicht in `index.xlsx` direkt, aber im Notizen-Audit jeder Datums-Korrektur.

### Wie das Datum geprüft wird

`landtag.py merge` rekonziliiert mit Korridor-Heuristik [-14d, +122d] gegen das DB-Datum: das deckt sowohl den Brief-vs-Drucksachen-Lag als auch den Registrierungs-Lag ab. Innerhalb Korridor → silent take-over, keine Flag. Außerhalb → `Datenqualität=ask_review`.

`resolve --auto` schließt den Rest mit drei Plausibilitäts-Mustern für `Anfragedatum`, die alle drei Quellen kreuzprüfen:

- **Body matcht DB** im Korridor → Body bleibt maßgeblich, Anfrage-Footer-Tippfehler wird notiert.
- **Footer matcht DB** im Korridor, Body weicht > 14 d ab → Body hat Tippfehler, `Anfragedatum` wird auf den Footer-Wert gesetzt (rollback).
- **Body ≈ N × 365 d** von DB und Footer entfernt (Jahr-Tippfehler im Body, N = 1–5) → rollback auf DB.
- **Body-Jahr außerhalb [2015, 2030]** (OCR-Glitch) → DB-Fallback (durch merge bereits gesetzt), Notiz dokumentiert.

Für `Antwortdatum` ist die Heuristik klassisch: Jahr-Tippfehler oder OCR-Glitch → DB gewinnt; sonst bleibt das PDF-Briefdatum maßgeblich.

Jede Datums-Korrektur-Notiz listet alle drei Quellen explizit: `Quellen: DB='…' | Antwort-Body='…' | Anfrage-PDF-Footer='…'`. Das macht den Audit-Trail vollständig nachvollziehbar.

- **PDF fehlt** (`pdf_missing`, `no_match` in `datum_original.xlsx`): DB-Datum bleibt als Fallback.
- **Manuelle Adoption**: Tags `anfragedatum_manual` / `antwortdatum_manual` in `Extract_Flags` sperren die Reconcile (für Fälle, die kein Pattern zuverlässig fängt — z. B. wenn Body und Footer beide das falsche Jahr nennen).
- **Datenqualität `review`** ist sticky: ein expliziter „bitte schauen"-Marker, der auch nach merge/resolve-Re-Runs erhalten bleibt, bis ein Mensch ihn räumt.

### Pipeline in einem Satz

`crawl` → `fetch-text` → `scan-archive` → `extract-multi-ministerium` → `build-abgeordnete-index` → `extract-all-anfrager` (Doppelpass) → `tools/extract_datum_original.py` → `merge` → `resolve --auto` → `verify`. Volle Befehlsliste in `skills/landtag-nrw-extraction/SKILL.md`.

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