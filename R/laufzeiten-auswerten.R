library(pacman)
p_load(dplyr)
p_load(tidyr)
p_load(xml2)
p_load(lubridate)
p_load(stringr)
p_load(openxlsx)


fname <- "data/index_work1.xlsx"

# Hilfsfunktion: HTML-Entities dekodieren (vektorisiert, NA-sicher)
decode_html <- function(x) {
  if (!is.character(x)) return(x)
  out <- vapply(x, function(s) {
    if (is.na(s) || !nzchar(s)) return(s)
    xml2::xml_text(xml2::read_html(paste0("<x>", s, "</x>")))
  }, character(1), USE.NAMES = FALSE)
  out
}

# Ausgangsdatei laden

anfragen_df <- read.xlsx(fname) %>% 
  mutate(Anfragedatum = as.Date(Anfragedatum),
         Antwortdatum = as.Date(Antwortdatum)) %>% 
  mutate(Antwortzeit = (Antwortdatum - Anfragedatum)) %>% 
  mutate(across(where(is.character), decode_html)) %>% 
  filter(WP==18) %>% 
  filter(Anfragedatum <= "2026-03-24")
  

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

write.xlsx(ministerien_df,"data/auswertung_ministerien.xlsx", overwrite=T)

# Verspätung nach Fraktion

parteien_df <- anfragen_df %>% 
  group_by(Fraktion) %>% 
  summarize(Anzahl = n(),
            Zeit =mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > 28,na.rm =T)
  ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(desc(Anzahl))

write.xlsx(parteien_df,"data/auswertung_parteien.xlsx", overwrite=T)

# Verspätung nach Abgeordneten

abgeordnete_df <- anfragen_df %>% 
  # Zwei inkorrekte Trennzeichen fixen
  mutate(Anfrager = str_replace(Anfrager," \\, ","; ")) %>% 
  # Drei reingerutsche Parteikürzel extrahieren
  mutate(Anfrager = str_remove(Anfrager," SPD")) %>% 
  # Lange Liste aus Anfrager-Spalte
  separate_rows(Anfrager, sep = "\\s*;\\s*") %>%   # an "; " trennen, Whitespace-tolerant
  mutate(Anfrager = str_squish(Anfrager)) %>%      # mehrfache/Rand-Leerzeichen weg
  filter(Anfrager != "") %>% 
  # ("u.a." mergen)
  mutate(Anfrager = str_remove(Anfrager," u\\.a\\.")) %>% 
  # Weiter
  group_by(Anfrager) %>% 
  summarize(Anzahl = n(),
            Zeit =mean(Antwortzeit, na.rm=T),
            Verspätet = sum(Antwortzeit > 28,na.rm =T)
  ) %>% 
  mutate(Pünktlichkeitsquote = 100-(Verspätet/Anzahl*100)) %>% 
  arrange(desc(Anzahl))

write.xlsx(abgeordnete_df,"data/auswertung_abgeordnete.xlsx", overwrite=T)


blanks_df <- anfragen_df %>% 
  filter(is.na(Ministerium_Kuerzel)) %>% 
  filter(!str_detect(Antworttext_Status,"anfrage_zurueckgezogen")) %>% 
  filter(!str_detect(Antworttext_Status,"pending"))

