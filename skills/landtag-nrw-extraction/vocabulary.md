# Landtag NRW — Vocabulary & Format-Konventionen

Stammdaten-Tabellen, Schreib-Konventionen, Aliase. Nur lesen, wenn ein konkreter Begriff aufschlägt — die Kern-Pipeline (siehe `SKILL.md`) deckt den Standard ab.

## Stammdaten-Dateien

- `Index/fraktionen.xlsx` — kanonische Fraktion-Namen + Aliase
- `Index/ministerien.xlsx` — kanonische Ministeriums-Namen, Kürzel, Aliase, WP
- `Index/ministerium_aliases.xlsx` — alte/abgekürzte Schreibweisen → Kürzel
- `Index/abgeordnete.xlsx` — `(WP, Fraktion, Nachname, Vorname, PDF_Form, Is_Alias, Frueher_Fraktion, n_treffer)` — eine Zeile pro (WP, Fraktion, Person), wahlperioden-übergreifend

## Namens-Format

- **DB / `Anfrager` / `Anfrager_Alle`**: `Nachname, Vorname` (mehrere mit `; ` getrennt). Titel (`Dr.`, `Prof. Dr.`) bleiben am Anfang des `Nachname`-Teils — also `Dr. Vincentz, Martin`, nicht `Vincentz, Dr. Martin`.
- **Antwort-PDF-Header**: `Vorname Nachname` (Komma-Liste, letzter mit `und`). Titel werden vor den Vornamen gestellt: `der Abgeordneten Berivan Aymaz, Josefine Paul und Dr. Sigrid Beer`.
- **Konverter:** `db_to_pdf_form` erzeugt aus DB-Form die PDF-Form; `db_to_pdf_form_aliases` ergänzt vereinfachte Varianten (Mittelinitial weg, Hyphen-Mittelnamen weg, Titel weg) — nötig, weil PDFs oft kürzere Schreibweisen verwenden („Sven W. Tritschler" → „Sven Tritschler").

## Kanonische Schreibweisen (häufige Konflikte)

| DB-Form (kanonisch) | PDF-Variante(n) | Notiz |
|---|---|---|
| `Pretzell, Marcus` | `Markus Pretzell` | DB-Form folgt der offiziellen MdL-Liste |
| `Tritschler, Sven W.` | `Sven Werner Tritschler`, `Sven Tritschler`, `Sven W. Trischler` | Mittelinitial steht in DB, in PDFs oft weg oder als Vollname |
| `Klocke, Arndt` | `Anrdt Klocke` | PDF-Tippfehler; DB ist korrekt |
| `Baran, Volkan` | `Volker Baran` | PDF-Tippfehler |
| `Loose, Christian` | `Christian Losse` | PDF-Tippfehler |
| `Strotebeck, Herbert` | `Hebert Strotebeck` | PDF-Tippfehler |
| `Müller-Witt, Elisabeth` | `Elisabeth Müller Witt` | Bindestrich-Verlust beim pdftotext |
| `Engin, Dilek` | `Engin Dilek` (im PDF Vorname/Nachname-Reihenfolge unklar) | DB-Reihenfolge: Nachname=Engin, Vorname=Dilek |
| `dos Santos Herrmann, Susana` | `Susana dos Santos Herrmann` | mehrteiliger Nachname — Lookup via PDF_Form, nicht via Last-Token |

## Fraktion-Mapping (DB ↔ PDF-Token)

| DB-Wert | PDF-Token-Varianten |
|---|---|
| `GRÜNE` | `BÜNDNIS 90/DIE GRÜNEN`, `BÜNDNIS 90/ DIE GRÜNEN`, `BÜNDNIS 90 / DIE GRÜNEN`, `BÜNDNIS90/DIEGRÜNEN` |
| `fraktionslos` | `FRAKTIONSLOS` |
| `AfD` | `AfD`, `AFD` (allcaps) |
| `SPD`, `CDU`, `FDP` | identisch |

## Edge-Case Fraktionswechsel

Verlässt eine Person ihre Fraktion mitten in der Wahlperiode (z. B. Pretzell, Neppe, Müller-Witt), erscheint sie in der DB unter beiden Fraktionen. `Index/abgeordnete.xlsx` trägt dann je eine Zeile pro `(WP, Fraktion, Person)`; die Spalte `Frueher_Fraktion` listet die anderen Fraktionen, unter denen dieselbe Person in derselben WP geführt wird.

## Ministeriums-Aliase

- **`MCdS`** (Minister + Chef der Staatskanzlei) = **`MBEIM`** (gleiche Person, andere Rolle). Via `Aliases`-Spalte in `Index/ministerien.xlsx` zusammengeführt.
- **`MKJFGF` vs. `MKJFGFI`**: Familienministerium (Kinder, Jugend, Familie, Gleichstellung, Flucht und Integration). Korrektes Kürzel ist `MKJFGFI`; alte Schreibweise `MKJFGF` ist als Alias gelistet.
- Alte Schreibweisen / Tippfehler / Genitiv-Formen: alle in `Index/ministerium_aliases.xlsx`.

## Wahlperioden-Wechsel

WP-Wechsel ändern den Ministeriumszuschnitt:
- **WP17** (Schwarz-Gelb, CDU+FDP)
- **WP18** (Schwarz-Grün, CDU+GRÜNE)

Neue Kürzel/Schreibweisen landen in `data/vocab_novelty.log`. Vor automatischem Mergen prüfen, ob es echt neue Ressorts sind oder nur Schreibvarianten — dann ggf. `Index/ministerium_aliases.xlsx` ergänzen oder `Index/ministerien.xlsx` anpassen. **Nicht blind absorbieren.**

## Anfrager-Vollständigkeit

Die DB-Spalte `Anfrager` liefert max. 2 Namen + `u.a.` (~1.400 Zeilen sind durch dieses Cap betroffen). Die volle Liste steht im Antwort-PDF-Header und wird über den Bootstrap-Loop in `Anfrager_Alle` / `Anzahl_Abgeordnete` geschrieben:

```sh
python landtag.py build-abgeordnete-index   # → Index/abgeordnete.xlsx aus DB-Anfrager-Spalte
python landtag.py extract-all-anfrager      # füllt Anfrager_Alle, Anzahl_Abgeordnete
python landtag.py build-abgeordnete-index   # 2. Pass: absorbiert Namen, die nur als 3.+ auftauchten
python landtag.py extract-all-anfrager      # 2. Pass: matcht jetzt auch diese
```

**Bootstrap-Limit:** Die Aggregation aus DB-Anfragern erfasst nur Personen, die mindestens einmal als 1./2. Anfrager (oder im 2. Pass als 3.+) auftauchen. Hinterbänkler*innen, die nie KAs (mit-)zeichnen, fehlen. Vor dem ersten Lauf einer neuen WP idealerweise die offizielle MdL-Liste scrapen:
<https://www.landtag.nrw.de/home/der-landtag/abgeordnete-und--fraktionen/die-abgeordneten/abgeordnetensuche/liste-aller-abgeordneten.html>

(One-shot-Skript noch nicht implementiert; bis dahin DB-only-Bootstrap.)

**Residue-Log:** Was der Parser im PDF-Anfragerblock nicht zuordnen kann, landet in `data/anfrager_novelty.log` (Format: `Drucksache | Fraktion | residue | block`). Vor manueller Korrektur prüfen, ob ein dritter Bootstrap-Pass das Problem löst.
