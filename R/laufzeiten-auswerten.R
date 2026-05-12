#' Datenbank Kleine Anfrage auswerten
#' 
#' Nutzt die Daten der Landtags-Datenbank NRW für Kleine Anfragen: 
#' https://www.landtag.nrw.de/home/dokumente/dokumentensuche/anfragen-und-antworten-suchergeb.html
#' 
#' Die Daten aus der Datenbank enthalten die wesentlichen Datenpunkte: 
#' Wer hat wann angefragt, wann kam von wem die Antwort. Außerdem liefert
#' sie "Sachgebiete" (eine kuratierte Schlagwortliste) und Schlagworte, 
#' die unspezifischer sind und deshalb nicht zur Klassifizierung taugen. 
#' 
#' Aus den PDFs der Kleinen Anfragen wurde ergänzt, welche Ministerien an
#' der Beantwortung beteiligt waren; in der Datenbank ist nur das federführende
#' Ministerium genannt. 
#' 
#' Mai 2026, CC-BY-SA Jan Eggers jan.eggers@fm.wdr.de

library(pacman)
p_load(dplyr)
p_load(tidyr)
p_load(xml2)
p_load(lubridate)
p_load(stringr)
p_load(openxlsx)

# Quelldatei
fname <- "data/index_fixed.xlsx"
# 17. und 18. Wahlperiode. Wenn mehr Daten genutzt werden sollen muss der Index 
# der Ministerien ergänzt werden, weil die sich von WP zu WP ändern. 

WAHLPERIODEN <- c(17,18)
# Für die Zeitreihen: 
GRANULARITY <- "month" # monatsweise summieren; Gruppierung nach Woche zu verrauscht
CUTOFF <- "2026-03-24"
FRIST_TAGE <- 28
TOP_N_THEMEN <- 20 # Für die Gesamtauswertung: Was waren die Topthemen?
TOP_N_THEMEN2 <- 20 # Auswertung nach Fraktionen: Cut nach 20 Topthemen
N_TOP <- 5

# Hilfsfunktion: HTML-Entities dekodieren (vektorisiert, NA-sicher)
decode_html <- function(x) {
  if (!is.character(x)) return(x)
  out <- vapply(x, function(s) {
    if (is.na(s) || !nzchar(s)) return(s)
    xml2::xml_text(xml2::read_html(paste0("<x>", s, "</x>")))
  }, character(1), USE.NAMES = FALSE)
  out
}

# Hilfsfunktion: Fraktion normalisieren
# Nimmt alle Einträge als Vektor. 
# wenn immer durchgängig bei einer Fraktion, die zurückgeben,
# wenn nicht, durch andere Fraktion ergänzen
n_fraktion <- function(f_v) {
  if (length(unique(f_v)) > 1)
    return (paste0(unique(f_v),collapse="/"))
  return (first(f_v))
}

# Hilfsfunktion: Sachgebiete-Tabelle
# "Sachgebiete" ist eine kuratierte Schlagwort-Liste, die zwar einige Überschneidungen
# enthält (z.B. "Arbeit und Beschäftigung" vs. "Arbeitsbedingungen"), aber einen
# ganz guten Überblick gibt über Themenkarrieren. 
# 
# Wird im Späteren nach Fraktion und gesamt ausgewertet, deshalb brauchen wir 
# die Funktionen. 

sachgebiete <- function(df) {
  sachgebiete_df <- df %>%
    separate_rows(Systematik, sep = "\\s*;\\s*") %>%
    mutate(Sachgebiet = str_squish(Systematik)) %>%
    filter(Sachgebiet != "") %>% 
    # Gruppieren nach Zeitraum und Sachgebieten - Zählung der Anfragen
    # je Sachgebiet nach Zeiteinheit
    group_by(periode,Sachgebiet) %>% 
    summarize(WP = first(WP),
              Anzahl = n(),
              Zeit =mean(Antwortzeit, na.rm=T),
              Verspätet = sum(Antwortzeit > FRIST_TAGE,na.rm=T),
    ) %>% 
    mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100))
  return (sachgebiete_df)
  # Daraus kann man dann die Gesamt-Anfragen für einen Zeitraum berechnen
  # (und die durchschnittliche Laufzeit, wenn man die Zahl der Anfragen
  # mit dem Durchschnittswert malnimmt, um die unterschiedlichen Anzahlen
  # an Anfragen pro Woche zu kompensieren)
}

# Ausgangsdatei laden

anfragen_alle_df <- read.xlsx(fname) %>% 
  # Daten aufbereiten
  mutate(Anfragedatum = as.Date(Anfragedatum),
         Antwortdatum = as.Date(Antwortdatum)) %>%
  mutate(periode = floor_date(Anfragedatum, unit=GRANULARITY)) %>% 
  mutate(Antwortzeit = as.numeric(Antwortdatum - Anfragedatum, units = "days")) %>% 
  mutate(across(where(is.character), decode_html)) %>% 
  # Rausziehen, was noch läuft oder was zurückgezogen wurde
  filter(!str_detect(Antworttext_Status,"pending")) %>% 
  filter(!str_detect(Antworttext_Status,"anfrage_zurueckgezogen")) %>% 
  filter(!str_detect(Antworttext_Status,"abbruch")) %>% 
  # Keine Anfragen nach dem CUTOFF-Datum berücksichtigen. 
  filter(Anfragedatum <= CUTOFF) 

# Der Haupt-Loop: 
# Wahlperioden werden einzeln aufgeschlüsselt.  
# (Zeitreihen werden für die gesamte Dauer der Daten berechnet)
for (wp in WAHLPERIODEN) {
  anfragen_df <- anfragen_alle_df %>% 
    # Wahlperiode filtern
    filter(WP==wp) 
  
    # Verspätung nach Ministerium
    
    ministerien_df <- anfragen_df %>% 
      group_by(Ministerium_Kuerzel) %>% 
      summarize(
        Ministerium = first(Ministerium_Canonical),
        Anzahl = n(),
        Zeit =mean(Antwortzeit, na.rm=T),
        Verspätet = sum(Antwortzeit > FRIST_TAGE,na.rm=T)
      ) %>% 
      arrange(desc(Anzahl)) %>% 
      # Gesamtwert ergänzen
      bind_rows(
        anfragen_df %>% 
          summarize(
            Ministerium = "LANDESREGIERUNG GESAMT",
            Anzahl = n(),
            Zeit = mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > FRIST_TAGE, na.rm=T),
            Ministerium_Kuerzel ="LREG"
          )
      ) %>% 
      # Wieviel Prozent der Anfragen wurden rechtzeitig beantwortet, 
      # also innerhalb der vorgeschriebenen FRIST_TAGE Tage?
      mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) 
    
    write.xlsx(ministerien_df,
               paste0("data/WP",wp,"/WP",wp,"_auswertung_ministerien.xlsx"),
               overwrite=T)
    
    # Verspätung nach Fraktion
    
    parteien_df <- anfragen_df %>% 
      group_by(Fraktion) %>% 
      summarize(Anzahl = n(),
                Zeit =mean(Antwortzeit, na.rm=T),
                Verspätet = sum(Antwortzeit > FRIST_TAGE,na.rm =T)
      ) %>% 
      mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
      arrange(desc(Anzahl))
    
    write.xlsx(parteien_df,
               paste0("data/WP",wp,"/WP",wp,"_auswertung_parteien.xlsx"),
               overwrite=T)
    
    # Verspätung nach Abgeordneten
    
    abgeordnete_df <- anfragen_df %>% 
      # Kompatibilität: Inzwischen haben wir entdeckt, dass
      # in der Datenbank nur die ersten beiden Abgeordneten auf 
      # einem Antrag stehen. Kein Problem, der Skill kratzt die
      # anderen aus den PDFs und gleicht sie mit der Abgeordneten-Referenz ab, 
      # das Ergebnis steht semikolon-separiert in "Anfrager_Alle".
      # Rüberkopieren, die Original_Anfrager brauchen wir nicht
      mutate(Anfrager = Anfrager_Alle) %>% 
      
      # Inkorrekte Trennzeichen fixen
      mutate(Anfrager = str_replace(Anfrager," \\, ","; ")) %>% 
      # Reingerutsche Parteikürzel extrahieren
      mutate(Anfrager = str_remove(Anfrager," SPD")) %>% 
      # Lange Liste aus Anfrager-Spalte
      separate_rows(Anfrager, sep = "\\s*;\\s*") %>%   # an "; " trennen, Whitespace-tolerant
      mutate(Anfrager = str_squish(Anfrager)) %>%      # mehrfache/Rand-Leerzeichen weg
      filter(Anfrager != "") %>% 
      # ("u.a." mergen)
      mutate(Anfrager = str_remove(Anfrager," u\\.a\\.")) %>% 
      # Weiter
      group_by(Anfrager) %>% 
      summarize(Fraktion = n_fraktion(Fraktion),
                Anzahl = n(),
                Zeit =mean(Antwortzeit, na.rm=T),
                Verspätet = sum(Antwortzeit > FRIST_TAGE,na.rm =T)
      ) %>% 
      mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
      arrange(desc(Anzahl))
    
    write.xlsx(abgeordnete_df,
               paste0("data/WP",wp,"/WP",wp,"_auswertung_abgeordnete.xlsx"),
               overwrite=T)
    
    # Sachgebiete in dieser Legislaturperiode
    sachgebiete_df <- sachgebiete(anfragen_df) %>% 
      # Gesamtwerte ergänzen und aufsummieren (ist ja noch per periode)
      ungroup() %>% 
      group_by(Sachgebiet) %>% 
      mutate(sum_zeit = Zeit * Anzahl) %>%  # Zum Aufsummieren
      summarize(
        WP = first(WP),
        Anzahl = sum(Anzahl,na.rm=T),
        Zeit = sum(sum_zeit, na.rm=T ) / sum(Anzahl, na.rm=T),
        Verspätet = sum(Verspätet)
      ) %>% 
      mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
      arrange(desc(Anzahl))
    
    write.xlsx(sachgebiete_df,
               paste0("data/WP",wp,"/WP",wp,"_auswertung_sachgebiete.xlsx"),
               overwrite=T)
    
    # Sachgebiete nach Fraktion in dieser WP
    fraktionen <- anfragen_df %>% pull(Fraktion) %>% unique()
    
    for (f in fraktionen) {
      fraktion_themen_zeitreihe_df <- sachgebiete(anfragen_df %>% 
                                          filter(Fraktion == f)) %>% 
        group_by(periode) %>%
        # Absteigend nach Anzahl sortieren
        arrange(desc(Anzahl)) %>% 
        # drei topthemen 
        mutate(rank = row_number()) %>% 
        slice_head(n = 3) %>% 
        pivot_wider(
          id_cols      = periode,
          names_from   = rank,
          values_from  = c(Sachgebiet, Anzahl),
          names_glue   = "{.value}_{rank}"
        ) %>% 
        ungroup() %>% 
        arrange(periode) %>% 
        relocate(1,2,5,3,6,4,7) 
        
      # Themen der Fraktion als Zeitreihe für diese Wahlperiode
      write.xlsx(fraktion_themen_zeitreihe_df,
                 paste0("data/WP",wp,"/WP",wp,"_sachgebiete_",f,"_zeitreihe.xlsx"),
                 overwrite=T)
      
      # Jetzt aufsummieren
      fraktion_themen_df <- sachgebiete(anfragen_df %>% 
                                          filter(Fraktion == f)) %>% 
        ungroup() %>% 
        group_by(Sachgebiet) %>% 
        mutate(sum_zeit = Zeit * Anzahl) %>%  # Zum Aufsummieren
        summarize(
          WP = first(WP),
          Anzahl = sum(Anzahl,na.rm=T),
          Zeit = sum(sum_zeit, na.rm=T ) / sum(Anzahl, na.rm=T),
          Verspätet = sum(Verspätet)
        ) %>% 
        arrange(desc(Anzahl)) %>% 
        # Top 20 reichen
        slice_head(n=TOP_N_THEMEN2)
      
      # Schreiben
      write.xlsx(fraktion_themen_df,
                 paste0("data/WP",wp,"/WP",wp,"_sachgebiete_",f,".xlsx"),
                 overwrite=T)
      
    } 
    # Viele Mitwirkende, viel Verspätung?
    mitwirkende_df <- anfragen_df %>% 
      # Bug in der Erzeugung der Ministerien-Anzahl fixen:
      # bei einigen gibt der Parser "0" an (wenn keine aus dem Text gefunden wurden)
      mutate(Beteiligte_Ministerien = ifelse(Beteiligte_Ministerien > 0, 
                                             Beteiligte_Ministerien,
                                             1)) %>% 
      group_by(Beteiligte_Ministerien) %>% 
      summarize(Anzahl = n(),
                Zeit = mean(Antwortzeit, na.rm=T),
                Verspätet = sum(Antwortzeit > FRIST_TAGE)) %>% 
      mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100))
    
    write.xlsx(mitwirkende_df,
               paste0("data/WP",wp,"/WP",wp,"_beteiligte_ministerien.xlsx"),
               overwrite=T)
}  

#' Zeitverläufe:
#'  
#'  Wann wurden besonders viel klein angefragt, und von wem?
#'  Welche Themen waren besonders beliebt? 


zeitreihe_df <- anfragen_alle_df %>%
  ungroup() %>%     
  separate_rows(Systematik, sep = "\\s*;\\s*") %>%
  mutate(Sachgebiet = str_squish(Systematik)) %>%
  filter(Sachgebiet != "") 

# Die n meistgenannten Themen nach Zeit
topthemen <- zeitreihe_df %>%
  count(Sachgebiet, sort = TRUE) %>%
  slice_max(n, n = TOP_N_THEMEN) %>%
  pull(Sachgebiet)

themen_df <- zeitreihe_df %>% 
  # Jetzt haben wir für jedes Sachgebiet eine Zeile. Zählen.
  group_by(periode,Sachgebiet) %>% 
  summarize(Anzahl = n()) %>% 
  ungroup() %>% 
  group_by(periode) %>%
  # Absteigend nach Anzahl sortieren
  arrange(desc(Anzahl)) %>% 
  # drei topthemen 
  mutate(rank = row_number()) %>% 
  slice_head(n = 5) %>% 
  pivot_wider(
    id_cols      = periode,
    names_from   = rank,
    values_from  = c(Sachgebiet, Anzahl),
    names_glue   = "{.value}_{rank}"
  ) %>% 
  ungroup() %>% 
  arrange(periode) %>% 
  relocate(1,2,7,3,8,4,9,5,10,6,11) 


write.xlsx(themen_df,
           paste0("data/ALLE_topthemen_zeitreihe.xlsx"),
           overwrite=T)

# Summenverlauf alle Anfragen, alle Zeiten nach Monaten

anfragen_zeitreihe_df <- anfragen_alle_df %>% 
  group_by(periode) %>% 
  summarize(
    WP = first(WP),
    Anzahl = n(),
    Zeit = mean(Antwortzeit, na.rm=T),
    Verspätet = sum(Antwortzeit > FRIST_TAGE, na.rm=T)
  ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(periode) %>% 
  rename(Datum=periode)

write.xlsx(anfragen_zeitreihe_df,
           "data/ALLE_zeitreihe.xlsx",
           overwrite=T)

# N_TOP längste, N_TOP schnellste
extreme_df <- anfragen_alle_df %>% 
  filter(WP=="18") %>% 
  arrange(desc(Antwortzeit)) %>% 
  slice_head(n=N_TOP) %>% 
  bind_rows(anfragen_alle_df %>% 
              filter(WP=="18") %>% 
              filter(!is.na(Antwortzeit)) %>% 
              arrange(desc(Antwortzeit)) %>% 
              slice_tail(n=N_TOP)) %>% 
  select(WP,Kleine_Anfrage_Nr,
         Anfrager,
         Fraktion_Canonical,
         Anfragedatum,Anfragetitel,
         Antwortdatum,
         Ministerium_Kuerzel,Beteiligte_Ministerien_Kuerzel,
         Beteiligte_Ministerien,Antwortzeit,
         Link_Drucksache_Anfrage,
         Link_Drucksache_Antwort)

write.xlsx(extreme_df,"data/WP18/WP18_extremwerte.xlsx")

übersicht_df <- anfragen_alle_df %>% 
  summarize(Anzahl = as.integer(n()),
            Zeit_Tage =mean(Antwortzeit, na.rm=T),
            Verspätet = as.integer(sum(Antwortzeit > FRIST_TAGE,na.rm =T))
  ) %>% 
  mutate(Pünktlich_Prozent = 100-(Verspätet/Anzahl*100)) %>% 
  pivot_longer(cols=everything(),names_to="Name", values_to="Wert")

wp18_übersicht_df <- anfragen_alle_df %>% 
  filter(WP==18) %>% 
  summarize(Anzahl = as.integer(n()),
            Zeit_Tage =mean(Antwortzeit, na.rm=T),
            Verspätet = as.integer(sum(Antwortzeit > FRIST_TAGE,na.rm =T))
  ) %>% 
  mutate(Pünktlich_Prozent = 100-(Verspätet/Anzahl*100)) %>% 
  pivot_longer(cols=everything(),names_to="Name", values_to="Wert")


write.xlsx(wp18_übersicht_df,"data/WP18/wp18_uebersicht.xlsx")

wp17_übersicht_df <- anfragen_alle_df %>% 
  filter(WP==17) %>% 
  summarize(Anzahl = as.integer(n()),
            Zeit_Tage =mean(Antwortzeit, na.rm=T),
            Verspätet = as.integer(sum(Antwortzeit > FRIST_TAGE,na.rm =T))
  ) %>% 
  mutate(Pünktlich_Prozent = 100-(Verspätet/Anzahl*100)) %>% 
  pivot_longer(cols=everything(),names_to="Name", values_to="Wert")


write.xlsx(wp17_übersicht_df,"data/WP17/wp17_uebersicht.xlsx")


