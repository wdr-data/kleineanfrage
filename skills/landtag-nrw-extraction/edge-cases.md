# Landtag NRW — Edge-Cases & PDF-Quirks

Sammlung von Quirks, die die Pipeline ausbügeln muss. Nur lesen, wenn ein konkreter Fall davon zutrifft — die Kern-Pipeline (siehe `SKILL.md`) deckt den Standard ab.

## Anfrager (PDF-Header → DB-Form)

`extract-all-anfrager` parst den Anfrager-Block des Antwort-PDFs gegen `Index/abgeordnete.xlsx`. Die meisten Edge-Cases sind bereits regex-toleriert; die unten genannten brauchen aber Hand-Korrekturen.

| PDF-Symptom | Bedeutung | Schnell-Korrektur |
|---|---|---|
| `der AbgeordnetenVorname Nachname` (kein Leerzeichen) | pdftotext-Whitespace-Loss am Wortende | bereits regex-toleriert |
| `VornameNachname` / `undVorname` | Whitespace-Loss innerhalb / vor Name | bereits via Lower-Upper-Heuristik gefixt; Reste manuell |
| `Volker Baran` statt `Volkan Baran`, `Anrdt Klocke` statt `Arndt`, `Sarah Philip` statt `Philipp` | Buchstaben-Tippfehler im PDF | manuell den DB-Namen setzen |
| `Markus Pretzell` (PDF) vs. `Marcus Pretzell` (DB) | unterschiedliche Schreibweise | DB-Form ist kanonisch (siehe `vocabulary.md`) |
| 2 `der Abgeordneten`-Blöcke, einer pro Fraktion (z. B. `Nic Vogel AfD` + `und Frank Neppe FRAKTIONSLOS`) | Cross-Fraktion-Co-Signer (selten, aber legitim) | beide Namen manuell zusammenführen; Fraktion der Zeile = Mehrheits-Fraktion |
| Block leer / `Antwort: Unterrichtung Präs` | Anfrage zurückgezogen | erwartet — Status `anfrage_zurueckgezogen`; kein Mismatch-Eintrag nötig |
| DB-Anfrager enthält Komma + Fraktion-Token (`Lisa-Kristin SPD , Dr. Pfeil, Werner`) | Crawl-Fehler weiter oben | siehe `verify` / `resolve` — mit `--counter-search` gegenprüfen |

Beim Komma-Bug im DB-`Anfrager` wandert der Müll auch in `Index/abgeordnete.xlsx` (Bootstrap-Pass absorbiert verseuchte Vornamen). Filter im Lookup: Vornamen mit Komma / Fraktion-Token überspringen.

## Plausi-Mismatch-Workflow

`extract-all-anfrager` vergleicht parsed-Anzahl mit DB-Mindesterwartung (1 oder 2 Namen, ≥3 wenn `u.a.`). Bei Unterlauf:

```
… | 17/172 | GRÜNE | MISMATCH parsed=0 db_min=1 | db_anfrager="Schäffer, Verena" | block="…"
```

Workflow:
1. **Fälle holen:** `grep MISMATCH data/anfrager_novelty.log`
2. **Antwort-PDF lesen:** Anfrager-Block direkt unter „Antwort der Landesregierung auf die Kleine Anfrage … vom …"
3. **Korrigieren oder markieren:**
   - Direkt: `Anfrager_Alle` + `Anzahl_Abgeordnete` setzen, `anfrager_manual` zu `Extract_Flags` (sonst wird's beim Re-Run überschrieben)
   - Markieren: Token wie `anfrager_unverified`, `anfrager_pdf_typo` zu `Extract_Flags` (Auswertungen filtern dann)

**Bulk-Remediation:** `tools/fix_anfrager_mismatches.py` arbeitet das Mismatch-Log batchweise ab — versucht (a) erweiterte Whitespace-Heuristik, (b) Cross-Fraktion-Matching, (c) DB-Fallback. Setzt `anfrager_manual` (+ ggf. `anfrager_from_db`, `anfrager_cross_fraktion`) bei Erfolg, sonst `anfrager_unverified`.

## Status-Edge-Cases

- **`Antwort: Unterrichtung Präs`** im Search-Hit → Anfrage zurückgezogen; das „Antwort"-Dokument ist eine Unterrichtung des Landtagspräsidenten. Status `anfrage_zurueckgezogen`; `Antwortdatum` = Datum der Unterrichtung. Skipt in `fetch-text` / `scan-archive`.
- **md ↔ KA-Nr-Mismatch** (verify-Report): Antwort-PDF erwähnt eine andere KA-Nr als der Index. Meist Crawl-Fehlmatch oder kaputtes PDF (pdftotext rendert manchmal mit Doppel-Buchstaben). Mit `resolve --ka N --counter-search` gegenprüfen.
- **KA-Nr Lücken** (verify): KAs sind fortlaufend nummeriert. Lücken sind zurückgezogen (`anfrage_zurueckgezogen`) oder noch unveröffentlicht.
- **Duplikat-KA-Nr** (verify): gleiche KA-Nr in mehreren Zeilen. Häufig **legitim** (Korrigenda / Nachgang).
- **negative Antwortzeit** (verify): Antwortdatum strikt vor Anfragedatum. Selten — Tippfehler im PDF (z. B. „April 2028" statt 2018), Crawl-Fehlmatch oder vertauschte Datumsspalten. 0-Tage-Antworten (gleicher Tag) sind legitim und werden nicht beanstandet.
- **DB↔PDF-Mismatch** (`Mismatch_Flags`): re-extracted PDF-Werte weichen von DB ab. Datums-Drift bis ±14 Tage wird ignoriert (DB = Drucksachen-Ausgabedatum vs. PDF Schreiben-Datum). Größere Differenzen: Tippfehler oder Crawl-Fehlmatch.
- **`pdf_database_mismatch`-Tag**: Antwort-PDF zur Drucksache enthält nicht die Antwort, sondern eine andere Anfrage (Landtag-seitiger Upload-Fehler oder Drucksachen-Duplikat). Beispiel: KA 7140 Rasche FDP wurde unter sowohl 18/17220 als auch 18/17720 veröffentlicht; der 18/17220-Slot blockiert damit den Antwort-Upload für KA 6714 Baran SPD.

## Beteiligte Ministerien

- Antwort-PDFs haben nach dem Anfragetext einen Absatz „Der Minister … / Die Ministerin …", der ggf. mehrere Ressorts nennt. Das ist die autoritative Quelle für **alle** beteiligten Ministerien. Search-Hit nennt nur das federführende.
- `extract-multi-ministerium` parst diesen Absatz → `Beteiligte_Ministerien_Kuerzel` (federführend zuerst).
- **Floor**: Jede Zeile mit gefülltem `Ministerium_Kuerzel` hat mind. 1 Ministerium beteiligt — `extract-multi-ministerium` und `normalize` seedet `Beteiligte_Ministerien_Kuerzel` mit dem federführenden Kürzel, falls der Boundary-Absatz fehlt.
- **WP17 vs. WP18** haben unterschiedlichen Ministeriumszuschnitt (Schwarz-Gelb vs. Schwarz-Grün). WP18 ist von `extract-multi-ministerium` voll abgedeckt; WP17 wird oberflächlicher behandelt.
- **WP-Wechsel-Novelty:** Neue Kürzel landen in `data/vocab_novelty.log`. Vor automatischem Mergen prüfen (siehe `vocabulary.md`).

## Crawl- und PDF-Quirks

- **Anfrage-PDFs im Archiv** sehen auf Seite 1 fast identisch aus wie Antworten und werden über den Anfrage-PDF-Filter ausgesondert (`is_antwort_drucksache`).
- **Antworten auf Große Anfragen** liegen z.T. mit im `Archiv/` und werden über den GA-Filter aussortiert.
- **`pdftotext` schluckt das D in „AfD"** → Spelling landet als „Af" in `vocab_novelty.log`. Bekannter pdftotext-Quirk auf manchen PDFs.
- **`Vormerkung` vs `Vorbemerkung`** → frühe WP18-Antworten nutzen alte Schreibweise; Title-Regex akzeptiert beide.
- **Doppel-Buchstaben-Render** (z. B. PDF-Header zeigt `1188//11115599` statt `18/1159`): pdftotext renderingbug bei manchen Drucksachen. Im Parser kompensiert (jede zweite Char wegnehmen).
- **`fetch-text` reportet 404** → Drucksache existiert online nicht (zurückgezogen oder noch unveröffentlicht). Status `no_answer_yet` ist erwartet.
- **`Extract_Flags` enthält `missing_ministerium` trotz gefülltem Kürzel** → Bug-Indikator. `compute_extract_flags` checkt seit Fix `ministerium OR ministerium_canonical OR ministerium_kuerzel`. Wenn Lücke trotzdem markiert: `normalize` re-runnen, dann manueller Re-Compute.

## Resolve-Pipeline für Mismatches

`merge` taggt jede Zeile mit `Datenqualität` (ok / korrigiert / ask_review). Pipeline für die `ask_review`-Reste:

1. **Index-Lookup** (rule-based): `Index/abgeordnete.xlsx`, `Index/ministerien.xlsx`, `Index/fraktionen.xlsx` für Schreibvarianten / Aliase. Heute via `extract-all-anfrager` + `tools/fix_anfrager_mismatches.py`.
2. **LLM-Einzelfall** (`enrich-llm`): semantische Klärung von Lücken via `llm`-CLI.
3. **Web-Recherche**: macht der Agent selbst via Claude-Code Web-Tools (Live-Suche bei Landtag, offizielle MdL-Liste).
4. **Human-Dialog**: `python landtag.py resolve --interactive` — pro Zeile DB↔PDF zeigen, Notiz + Verdict (k=korrigiert / s=skip / q=quit). Resume-safe — Zeilen mit nicht-leeren `Notizen` werden übersprungen.
