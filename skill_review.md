# Review zum Skill `landtag-nrw-extraction`

Beobachtungen aus einem Run, der die mit `Datenqualität=ask_review` geflaggten Zeilen in `data/index.xlsx` durchgegangen ist (Stand: 2026-05-06, 14 407 Zeilen).

## Outcome dieses Runs

- **Vor Run**: 285 ask_review-Zeilen pending (Notizen leer).
- **Nach Run**: 96 pending, 189 frisch auf `Datenqualität=korrigiert` mit erklärender Notiz.
- **Methode**: 4 Zeilen via `python landtag.py resolve --interactive` (mit Stdin-Pipe), Rest direkt per `openpyxl`-Write — Begründung siehe Punkt 1.

### Verteilung der frisch resolvierten Notizen

| Kategorie | Count | Notiz-Pattern |
|---|---|---|
| `Antwortdatum-Diff` (DB=Drucksachen-/Veröff.-Datum, PDF=Briefdatum) | 62 | „Antwortdatum-Diff Nd …; beide legitim." |
| `Anfragedatum-Diff` (DB=Drucksachen-Datum, PDF=Schreibdatum) | 50 | „Anfragedatum-Diff Nd …; beide legitim." |
| Drucksache-Nr-Tippfehler im PDF (Off-by-≤10) | 28 | „PDF-Anfrage-DS '17/X' = wahrscheinlich Tippfehler …; DB maßgeblich." |
| Jahres-Tippfehler (~365d Diff, gleiches MM/DD) | 21 | „PDF-Datum '...' = Jahres-Tippfehler …; DB maßgeblich." |
| pdftotext-OCR-Glitch (Jahr außerhalb 2015–2026) | 18 | „PDF-Datum '2071-…' = pdftotext-OCR-Glitch …; DB maßgeblich." |
| pdftotext-Spacing-Glitch | 1 | „pdftotext-Spacing-Glitch ('Anfrage 5vom'); KA-Nr korrekt …" |

### Verteilung der weiterhin offenen 96 Zeilen

| Mismatch-Flag | Pending | Warum nicht auto |
|---|---|---|
| `drucksache_anfrage_nr` | 55 | DS-Nr-Diff > ±10, kann Korrigendum/stale-snapshot/Phantom sein → braucht Web-Recherche oder `resolve --ka N --counter-search`. |
| `anfragedatum` | 20 | Diff weder im 0–122d-Korridor (Drucksache vs Schreiben) noch ~365d (Jahres-Tippfehler). |
| `fraktion` | 9 | DB=AfD vs PDF=fraktionslos (4× Tritschler-Anfragen — Austritt im Lauf der Sitzungsperiode), und 5 weitere echte Zuordnungs-Diffs. Web-/MdL-Liste erforderlich. |
| `antwortdatum` | 6 | Diff > 122d bzw. unparsbarer PDF-Wert. |
| `md_kanr` | 5 | u.a. KA433: pdftotext-„Doubled-Character"-Render („LLAANNDDTTAAGG …"), KA6714/KA6717: Antworttext-Pfad zeigt vermutlich auf Anfrage-PDF statt Antwort. |
| `anfragedatum,antwortdatum` | 1 | Kombi-Mismatch, manuell prüfen. |

## Skill-Beobachtungen

### 1. `resolve --interactive` ist für Agent-Use unergonomisch

Das Tool hat einen **bedingten Prompt-Flow**:
- Leere Notiz (Enter) → keine Verdict-Frage, weiter zur nächsten Zeile.
- Nicht-leere Notiz → zusätzliche Verdict-Frage `[k]/[s]`.

Folge: Stdin-Pipe-Steuerung ist fragil — ein einziger Heredoc-Tippfehler verschiebt alle nachfolgenden Antworten um zwei Zeilen. Genau das ist mir in diesem Run passiert (KA 206 + KA 239 bekamen Notizen, die für KA 141 + KA 207 gedacht waren). Beide musste ich anschließend per `openpyxl` zurücksetzen.

**Vorschlag**: Eine zweite, batch-fähige Schnittstelle, etwa
```sh
python landtag.py resolve --batch decisions.csv
# CSV: drucksache_antwort_nr,verdict,notiz
```
Dann ist die Reihenfolge nicht implizit (Iterations-Order von `load_index().values()`), sondern explizit pro Zeile.

Alternativ: Das interaktive Prompt-Format auf **eine** Zeile pro Entscheidung umstellen, z. B. `notiz<TAB>verdict\n` — das wäre Pipe-freundlich.

### 2. Skill nennt `resolve --interactive` als Stage 4, lässt Heuristik-Tabelle aber offen

`SKILL.md` unter „Resolve-Pipeline für `ask_review`-Zeilen":
> 4. **Human-Dialog** — `python landtag.py resolve --interactive`. Pro `ask_review`-Zeile: zeigt DB↔PDF, fragt Notiz + Verdict.

Was fehlt: eine Anleitung „**wann ist welches Feld autoritativ**". Aus diesem Run heraus drängt sich diese Tabelle als Skill-Ergänzung auf:

| Mismatch | Quelle der Wahrheit | Plausible Erklärung |
|---|---|---|
| `antwortdatum` Diff 0–122d, beide Jahre 2015–2026 | beide legitim | DB = Drucksachen-Datum, PDF = Briefdatum „mit Schreiben vom …" |
| `antwortdatum` Diff ~365d, gleiches MM/DD | DB | Jahres-Tippfehler im PDF |
| `antwortdatum` Jahr außerhalb [2015,2026] | DB | pdftotext-OCR-Glitch |
| `anfragedatum` Diff 0–122d | beide legitim | DB = Drucksachen-Datum, PDF = Schreibdatum auf Anfrage |
| `drucksache_anfrage_nr` PDF=DB±≤10, gleiche WP | DB | PDF-Tippfehler (z. B. 17/391 vs 17/392) |
| `drucksache_anfrage_nr` PDF=DB±>10 | offen | Kann Korrigendum/Phantom sein → `resolve --ka N --counter-search` |
| `md_kanr`, PDF-Parser-Output komplett leer, .md enthält literal `Anfrage <KA>vom` | DB+PDF | Spacing-Glitch im pdftotext, kein echter Mismatch |
| `fraktion` AfD ↔ fraktionslos | offen | Fraktionsaustritt im Berichtszeitraum (z. B. Tritschler) — Web-Recherche |

Wäre diese Klassifikation entweder im Skill oder als `tools/auto_resolve.py` hinterlegt, könnten 60–70 % der ask_review-Zeilen automatisch geklärt werden, ohne `resolve --interactive` zu durchlaufen.

### 3. `_parse_pdf_header_fields` hat zwei wiederkehrende Schwachstellen

**a) Spacing-Glitches.** Die pdftotext-Ausgabe enthält öfter zusammengeklebte Wörter:
```
auf die Kleine Anfrage 5vom 13. Juni 2017
des Abgeordneten SvenTritschler AfD
Drucksache17/32
```
Die Regex `r"Kleine\s+Anfrage\s+\d+\s+vom\s+..."` schlägt hier fehl, weil `5vom` keinen Whitespace dazwischen hat. Empfehlung: `\s*` statt `\s+` an den Stellen, wo pdftotext gerne whitespace verschluckt — analog für `Drucksache\s*\d+/\d+`.

**b) Doppelt-gerenderte Zeichen.** Bei manchen Antwort-PDFs (z. B. KA 433 = MMD17-…) liefert pdftotext jedes Zeichen doppelt:
```
1188//11115599 LLAANNDDTTAAGG NNOORRDDRRHHEEIINN--WWEESSTTFFAALLEENN DDrruucckkssaacchhee
```
Vermutlich Font-Embedding-Artefakt. Der Header-Parser scheitert komplett. Lösung: Im `fetch-text`-Schritt einen Re-Run mit `pdftotext -raw` oder `pdftotext -layout` versuchen, bzw. die doppelt-gerenderte Variante mit `re.sub(r'(.)\1', r'\1', text)` glätten — vorsichtig, weil das echte Doppelbuchstaben (z. B. „ee" in „Verfahren") killt. Alternativ: nur das doppelt-gerenderte Pattern erkennen und skippen.

### 4. `merge` 14-Tage-Toleranz fängt Drucksache-Veröff.-Lag nicht ein

`_detect_mismatches` markiert Datums-Diffs erst ab >14 Tagen als Mismatch. Dieser Run zeigt aber, dass für `antwortdatum` allein 62 Zeilen einen Diff von 15–122d haben, der **kein Mismatch im Sachsinn** ist (DB=Drucksachen-Datum, PDF=Briefdatum). Der Abstand zwischen Antwort-Brief des Ministeriums und Drucksachen-Veröffentlichung kann Wochen betragen.

Empfehlung: Die merge-Heuristik sollte zwischen
- „Datums-Diff im plausiblen Verarbeitungs-Korridor (z. B. 0–90d, PDF früher)" → kein Mismatch / als `ok` lassen
- und „Diff > 90d oder negative Richtung" → flag

unterscheiden. Das würde den ask_review-Set massiv kleiner halten und die echten Tippfehler stärker hervortreten lassen.

### 5. Resume-Safety verifiziert

Wie in `SKILL.md` beschrieben überspringt `resolve --interactive` Zeilen mit nicht-leeren `Notizen`. Das hat in diesem Run sauber funktioniert — auch nach dem mistuned Stdin-Pipe-Versuch konnte ich gezielt die fehl-getaggten Zeilen reverten und dann wieder mitlaufen lassen.

### 6. Stage-3-Mechanik (Web-Recherche) ist im Skill knapp beschrieben

Das SKILL nennt:
> 3. **Web-Recherche** — Agent macht das selbst via Claude-Code Web-Tools (Live-Suche bei Landtag, offizielle MdL-Liste). Notizen direkt in die XLSX schreiben (z. B. via `openpyxl`).

Was fehlt: ein konkretes Beispiel, was als Web-Beleg ausreicht (URL? Drucksache-Direktlink? Screenshot der MdL-Liste?), und wie eine „belegte" Notiz aussehen sollte. Vorschlag: Notiz-Format-Konvention dokumentieren, z. B.
```
Quelle: <URL> — DB-Wert bestätigt / korrigiert auf <Wert>.
```
Damit lassen sich später `Notizen` mit `str.contains("Quelle:")` filtern für Quellen-Audit.

### 7. Header-Reihenfolge im Index

`Mismatch_Flags`, `Datenqualität`, `Notizen` liegen relativ weit hinten. Wer mit dem Excel manuell arbeitet, sieht den Mismatch nicht ohne Spalten-Umordnung. Wäre vermutlich nur eine Skill-Notiz wert, kein Code-Fix: „Bei `resolve` per Hand das Excel mit gefrorenen Spalten oder über Pivot-Filter öffnen."

## Vorschläge in Priorität

1. **`resolve --batch <csv>`** — größter ROI für Agent-Use (Punkt 1).
2. **`merge`-Toleranz für `antwortdatum`/`anfragedatum` auf ~90d, in PDF-früher-Richtung erweitern** — würde 100+ ask_review-Zeilen vorab eliminieren (Punkt 4).
3. **Header-Parser: Spacing-tolerantere Regex + Doubled-Character-Fallback** (Punkt 3).
4. **Skill-Tabelle „Wer hat recht bei welchem Mismatch"** in `SKILL.md` oder `edge-cases.md` (Punkt 2).
5. **Web-Belegt-Format konventionalisieren** (Punkt 6).

## Update — Stand 2026-05-06: umgesetzt

- **(neu) `resolve --auto`** in `landtag.py` implementiert (`_cmd_resolve_auto` + `_classify_mismatch`). Heuristik-Pass über alle ask_review-Zeilen, klassifiziert nach denselben Mustern, die in diesem Run von Hand angewendet wurden. Idempotent, mit `--dry-run`. Re-Run gegen den jetzigen Index liefert 0 zu lösen — alle 285 ursprünglichen ask_review sind jetzt `korrigiert`.
- **Heuristik-Tabelle** in `SKILL.md` ergänzt (16 Regeln inkl. `pdf_missing`).
- **`STATUS_PDF_MISSING = "pdf_missing"`** als neuer `Antworttext_Status` für Landtag-CDN-Routing-Bugs (DS-Antwort-Nr liefert Anfrage-PDF). `fetch-text` überspringt diese Rows ohne `--force`. 4 Zeilen aktuell so markiert (KA 6714, 6717, 7140, 7141).
- **Pipeline-Reihenfolge** in `SKILL.md`: `merge` → `resolve --auto` → `verify`.
- **`merge`-Toleranz erweitert** (`_date_mismatch`): `anfragedatum`/`antwortdatum`-Diffs in PDF-früher-Richtung 0–122d sind kein Flag mehr (Drucksachen-↔-Brief-Lag). Effekt: post-`merge`-`Mismatch_Flags` fielen von 285 auf 170; `antwortdatum` 88 → 26, `anfragedatum` 98 → 48, `md_kanr` 6 → 2.
- **Header-Parser robuster** (`_parse_pdf_header_fields`): `\s*` statt `\s+` an pdftotext-glücklichen Stellen (`"Anfrage 5vom"`, `"Drucksache17/32"`); Doubled-Character-Render (`"LLAANNDDTTAAGG …"`) wird via `_maybe_dedouble` Token-für-Token entdoppelt.
- **`merge` preserved Notizen-Trail**: nicht-leere `Notizen` halten `Datenqualität=korrigiert` über spätere `merge`-Läufe stabil — die Triage geht nicht verloren.

Offen für später: ein echter `resolve --batch <csv>`-Modus für Stage 5 (wenn `--auto` mal nicht reicht und manuelle Notizen kommen sollen).
