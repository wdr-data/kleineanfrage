library(pacman)
p_load(dplyr)
p_load(tidyr)
p_load(xml2)
p_load(lubridate)
p_load(stringr)
p_load(openxlsx)


fname <- "data/index.xlsx"
wp <- 18

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
  if (length(unique(f_v) > 1))
    return (paste0(unique(f_v),collapse="/"))
  return (first(f_v))
}


# Ausgangsdatei laden

anfragen_df <- read.xlsx(fname) %>% 
  mutate(Anfragedatum = as.Date(Anfragedatum),
         Antwortdatum = as.Date(Antwortdatum)) %>% 
  mutate(Antwortzeit = (Antwortdatum - Anfragedatum)) %>% 
  mutate(across(where(is.character), decode_html)) %>% 
  # Wahlperiode filtern
  filter(WP==wp) %>% 
  filter(Anfragedatum <= "2026-03-24") %>%   
  # Rausziehen, was noch läuft oder was zurückgezogen wurde
  filter(!str_detect(Antworttext_Status,"pending")) %>% 
  filter(!str_detect(Antworttext_Status,"anfrage_zurueckgezogen")) %>% 
  filter(!str_detect(Antworttext_Status,"abbruch"))
  

  
  

# Verspätung nach Ministerium

ministerien_df <- anfragen_df %>% 
  group_by(Ministerium_Kuerzel) %>% 
  summarize(
            Ministerium = first(Ministerium_Canonical),
            Anzahl = n(),
            Zeit =mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > 28,na.rm=T)
            ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(desc(Anzahl))

write.xlsx(ministerien_df,
           paste0("data/auswertung_ministerien_",wp,".xlsx"),
           overwrite=T)

# Verspätung nach Fraktion

parteien_df <- anfragen_df %>% 
  group_by(Fraktion) %>% 
  summarize(Anzahl = n(),
            Zeit =mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > 28,na.rm =T)
  ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(desc(Anzahl))

write.xlsx(parteien_df,
           paste0("data/auswertung_parteien_",wp,".xlsx"),
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
            Verspätet = sum(Antwortzeit > 28,na.rm =T)
  ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(desc(Anzahl))

write.xlsx(abgeordnete_df,
           paste0("data/auswertung_abgeordnete_",wp,".xlsx"),
           overwrite=T)

# Zeitverlauf
# gruppiert nach Woche der Einreichung, Filter auf Ministerium

startdatum <- anfragen_df %>% pull(Anfragedatum) %>% first()

zeitreihe_df <- anfragen_df %>%
  arrange(Anfragedatum) %>% 
  mutate(woche = as.integer(
    (Anfragedatum-startdatum)/7)) %>% 
  group_by(woche) %>% 
  summarize(
    Anzahl = n(),
    Zeit =mean(Antwortzeit, na.rm=T),
    Verspätet = sum(Antwortzeit > 28,na.rm=T)
  ) %>% 
  mutate(datum = startdatum+7*woche)

write.xlsx(zeitreihe_df,
           paste0("data/auswertung_zeitreihe_",wp,".xlsx"),
           overwrite=T)

sachgebiete_long_df <- anfragen_df %>%
  separate_rows(Systematik, sep = "\\s*;\\s*") %>%
  mutate(Sachgebiet = str_squish(Systematik)) %>%
  filter(Sachgebiet != "") %>% group_by(Fraktion, Sachgebiet) %>% 
  summarize(Anzahl = n(),
            Zeit =mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > 28,na.rm=T),
            ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100))