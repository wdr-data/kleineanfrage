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
fname <- "data/index.xlsx"
# 17. und 18. Wahlperiode. Wenn mehr Daten genutzt werden sollen muss der Index 
# der Ministerien ergänzt werden, weil die sich von WP zu WP ändern. 

WAHLPERIODEN <- c(17,18)
# Für die Zeitreihen: 
GRANULARITY <- "month" # monatsweise summieren; Gruppierung nach Woche zu verrauscht
CUTOFF <- "2026-03-24"
FRIST_TAGE <- 28
TOP_N_THEMEN <- 30

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
  filter(!str_detect(Antworttext_Status,"abbruch"))

# Der Haupt-Loop: 
# Wahlperioden werden einzeln aufgeschlüsselt.  
# (Zeitreihen werden für die gesamte Dauer der Daten berechnet)
for (wp in WAHLPERIODEN) {
  anfragen_df <- anfragen_alle_df %>% 
    # Wahlperiode filtern
    filter(WP==wp) %>% 
    filter(Anfragedatum <= CUTOFF) 
  
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
               paste0("data/WP",wp,"_auswertung_ministerien.xlsx"),
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
               paste0("data/WP",wp,"_auswertung_parteien.xlsx"),
               overwrite=T)
    
    # Verspätung nach Abgeordneten
    
    abgeordnete_df <- anfragen_df %>% 
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
               paste0("data/WP",wp,"_auswertung_abgeordnete.xlsx"),
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
               paste0("data/WP",wp,"_auswertung_sachgebiete.xlsx"),
               overwrite=T)
    
    # Sachgebiete nach Fraktion in dieser WP
    fraktionen <- anfragen_df %>% pull(Fraktion) %>% unique()
    
    for (f in fraktionen) {
      fraktion_themen_zeitreihe_df <- sachgebiete(anfragen_df %>% 
                                          filter(Fraktion == f))
      
      # Themen der Fraktion als Zeitreihe für diese Wahlperiode
      write.xlsx(fraktion_themen_zeitreihe_df,
                 paste0("data/WP",wp,"_sachgebiete_",f,"_zeitreihe.xlsx"),
                 overwrite=T)
      
      # Jetzt aufsummieren
      fraktion_themen_df <- fraktion_themen_zeitreihe_df %>% 
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
      
      # Schreiben
      write.xlsx(fraktion_themen_df,
                 paste0("data/WP",wp,"_sachgebiete_",f,".xlsx"),
                 overwrite=T)
      
    } 
    # Viele Mitwirkende, viel Verspätung?
    mitwirkende_df <- anfragen_alle_df %>% 
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
               paste0("data/WP",wp,"_beteiligte_ministerien.xlsx"),
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
  filter(Sachgebiet %in% topthemen) %>% 
  select(periode,Sachgebiet) %>% 
  count(periode,Sachgebiet) %>% 
  pivot_wider(names_from = Sachgebiet, values_from = n, values_fill = 0) %>%
  arrange(periode) 

write.xlsx(themen_df,
           paste0("data/ALLE_sachgebiete_zeitreihe.xlsx"),
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

    
