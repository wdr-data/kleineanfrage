# Beobachtungen fehlende Daten

- KA 144: Anfrage-Dokument und Anfrager fehlen; über die Suche https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten-suchergeb.html?nummer=144&doktyp=KA&wp=18 leicht zu finden. Datenreihe enthält völlig falsche Daten für Antwort und antwortendes Ministerium. Gelistetes Antwortdokument ist Antwort auf eine Große Anfrage
- KA 256: als missing_ministerium markiert, aber vollständig und korrekt
- KA 448: Neues Kürzel "Unterrichtung Präs", Unterrichtung durch den Landtagspräsidenten: Fraktion hat die KA zurückgezogen
- KA 1266: Neues Kürzel "MCdS" für "Der Minister für Bundes- und Europaangelegenheiten, Internationales sowie Medien des Landes Nordrhein-Westfalen und Chef der Staatskanzlei", entspricht also MBEIM, allerdings in anderer Rolle.
- KA 32: Korrekt markiert MKJFGF; Suche ergibt MKJFGFI als Ministeriums-Kürzel.
- KA 4338: MKJFGFI musste erst Rücksprache halten mit Innenministerium; haben keine klare Policy dafür
- Drucksache 18/13129: Antwort auf eine Große Anfrage - out of scope; vermutlich schon in den Dokumenten.

# Vorgeschlagene Konsequenzen

- Nachrecherche über KA-Nummer bei Zweifelsfällen, mit Filter "Kleine Anfrage"
- Möglicher Bug: Wenn Daten über Suche ergänzt werden, werden missing_ Tags nicht rausgenommen
- KA sind fortlaufend nummeriert; regelbasiert Lücken suchen.
- Kleiner Sicherheitcheck auf String "Kleine Anfrage <nr>"
- Filter "Unterrichtung Präs": im Datensatz lassen, aber mit anfrage_zurueckgezogen markieren; Antwortdatum ist Datum der Unterrichtung.
- MBEIM und MCdS zusammenlegen
- Kürzel für Familienministerium ändern: MKJFGFI
- Mögliche Misch-Strategie LLM und regelbasiert zur Extraktion beteiligter Ministerien (vgl. KA4338). Die Antworten sind immer gleich aufgebaut: Im Anschluss an den Text der Anfrage kommt ein Textabsatz "Der Minister..." bzw. "Die Ministerin...", der die beteiligten Ministerien nennt. Vorgehen: Abgleich mit Anfragetext; dann den ersten Absatz auswerten. Zum Scan nach Ministerien in Schreibweise eventuell die ministerien_aliases Datei ergänzen und vom Human korrigieren lassen.