# Multi-Ministerium-Parser — Spezifikation

> **Historisches Dokument (Stand 2026-05-04).** Diese Spec war die Vorlage für den Multi-Ministerium-Parser, der inzwischen in `landtag.py` umgesetzt ist (Spalten `Beteiligte_Ministerien` + `Beteiligte_Ministerien_Kuerzel`). Aktuelle Pipeline-Sicht in [`../../../skills/landtag-nrw-extraction/SKILL.md`](../../../skills/landtag-nrw-extraction/SKILL.md).

Stand 2026-05-04. Ergänzt das bestehende `landtag-nrw-extraction`-Skill um Erkennung **mehrerer beteiligter Ministerien** je Antwort-PDF.

## Problem

Der Search-Hit nennt nur das **federführende** Ministerium (Spalte `Ministerium_Kuerzel`). Antworten der Landesregierung werden in der Praxis aber häufig **mit weiteren Ressorts abgestimmt** — diese Ressorts erscheinen nur im Antwort-PDF, nicht im Search-Index.

Beispiel: KA 4338 — federführend MKJFGFI, aber der Antworttext beginnt mit
> „Die Ministerin für Kinder, Jugend, Familie, Gleichstellung, Flucht und Integration hat im Einvernehmen mit dem Minister des Innern die Kleine Anfrage … beantwortet."

→ Beteiligt: MKJFGFI **und** IM. Aktuell nur MKJFGFI in der xlsx.

## Wo der Hinweis im PDF steht

Stabiles Layout aller Antwort-Drucksachen:
1. Header (Wahlperiode, Drucksache-Nr, Datum)
2. „Antwort der Landesregierung auf die Kleine Anfrage N"
3. Anfrager + Fraktion
4. Drucksache-Verweis auf Frage-DS
5. Titel der KA
6. Vorbemerkung der Kleinen Anfrage (= zitierter Anfragetext)
7. **Erster Absatz nach dem Anfragetext** = "Der Minister … / Die Ministerin … hat (im Einvernehmen mit …) die Kleine Anfrage … beantwortet."
8. Antworten auf Einzelfragen

**Schritt 7 ist die autoritative Quelle.** Vor Schritt 7 nicht parsen — der Anfragetext (Schritt 6) zitiert oft selber Ministerien als Recherche-Aufhänger; das wären False Positives.

## Vorgehen

### Phase A — Boundary finden

Das Ende des Anfrage-Absatzes (Schritt 6) ist marker-fähig. Mögliche Anker:

1. Das wörtliche Ende des Anfragetexts ist nicht eindeutig markiert. Stattdessen:
2. Erster Absatz, der mit `Der Minister`, `Die Ministerin`, `Der Minister für`, `Die Ministerin für`, `Der Minister präsident`, `Die Ministerpräsidentin` oder ähnlich beginnt UND das Wort `beantwortet` (oder Variante) im selben Absatz enthält.

Regex-Skizze:
```
(?:^|\n)\s*((?:Der|Die)\s+(?:Minister(?:präsident(?:in)?|in)?)\s+[^\n]+(?:\n[^\n]+)*?\s+beantwortet[^.]*\.)
```

Erstes Match nach Schritt 6 = der gesuchte Absatz. Cap auf z. B. 800 Zeichen Absatzlänge.

### Phase B — Ministerien extrahieren

Aus dem so isolierten Absatz mehrere „Minister*"-Erwähnungen ziehen. Pro Erwähnung Kürzel via Index-Tabelle (`Index/ministerien.xlsx` + `Index/ministerium_aliases.xlsx`) auflösen:

- `match_ministerium(value, canon_min)` (existiert) für jede Vollform
- Aliases-Lookup, falls Vollform 1:1 in Tabelle

Phrasen, die immer mehrere Ministerien koppeln:
- „im Einvernehmen mit", „in Abstimmung mit", „gemeinsam mit", „nach Beteiligung von", „unter Mitwirkung von"

Regex-Idee zur Extraktion aus dem Boundary-Absatz:
```
(?:Der|Die)\s+(Minister(?:präsident(?:in)?|in)?)\s+(?:der|des|für|für\s+die)?\s*([^,.\n]{1,150}?)(?=\s*(?:hat|im\s+Einvernehmen|in\s+Abstimmung|,|\.|\n))
```

Multi-finditer → Liste roher Vollformen → Alias-/Subset-Match → Set unique Kürzel.

### Phase C — Persistenz

Neue Spalte in `index.xlsx`:
- `Beteiligte_Ministerien_Kuerzel` — comma-separated, primary Kürzel zuerst (= bisheriges `Ministerium_Kuerzel`)
- Optional: `Beteiligte_Ministerien_Anzahl` als int für quick-filter

`Ministerium_Kuerzel` (Singular) bleibt das **federführende** Ressort, wird nicht überschrieben — sonst zerstören wir die Daten aus dem Search-Hit.

## Edge Cases

- **Nur ein Ministerium**: `Beteiligte_*` enthält genau das eine Kürzel — gleich `Ministerium_Kuerzel`.
- **Anfrage zurückgezogen** (Status `anfrage_zurueckgezogen`): Phase A findet keinen passenden Absatz → `Beteiligte_*` bleibt leer.
- **PDF-Parser-Fehler** (status `extract_failed`): überspringen.
- **Subjekt nicht als Vollform** sondern als Akronym im PDF: kommt selten vor; falls `match_ministerium` nichts findet, Vollform unverändert in `vocab_novelty.log` schreiben (zur Pflege von `ministerium_aliases.xlsx`).

## Trigger-Workflow

1. Implementierung als neuer Verb `extract-multi-ministerium` ODER als optionaler Pass in `scan-archive` (gesteuert per Flag, Default an).
2. Liest `data/index.xlsx`, iteriert beantwortete Rows mit lokalem PDF, parst Phase A+B, schreibt Spalten.
3. Idempotent — re-runnable; überschreibt `Beteiligte_*` bei jedem Lauf.
4. Nach erfolgreichem Pass `normalize` ausführen, damit ggf. neue Aliases aus `vocab_novelty.log` gepflegt werden können.

## Validierung

Stichprobe der bekannten Multi-Ministerium-Fälle:
- KA 4338 → erwartet `MKJFGFI, IM`
- Weitere bekannte zu sammeln, sobald 1. Pass durchläuft (User reviewt vocab_novelty.log)

## Out of scope (vorerst)

- Reihenfolge der Beteiligten als „Hierarchie" auswerten (federführend vs. mitwirkend) — schon implizit via `Ministerium_Kuerzel` (federführend) vs `Beteiligte_*` (alle).
- LLM-Fallback wenn Regex-Pfad scheitert. Erst nachschalten, wenn Stichprobe Lücken zeigt.
- Cross-Wahlperioden — derzeit nur WP18.

## Aufwand

~150 LOC Python (Phase A regex + Phase B extraction loop + Spalten-Migration). Schätzung 2-3 h inkl. Validierung an 10 Stichproben.
