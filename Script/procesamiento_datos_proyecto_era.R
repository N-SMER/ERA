
rm(list = ls())
gc()
getwd()
setwd("C:/Users/USUARIO/Downloads/Proyecto ERA Emergencias/Data procesada")

# install.packages(pacman)
require(pacman)
p_load(foreign, tidyverse, rio, here, dplyr, viridis, readxl, stringr, RColorBrewer, ggcorrplot,  
       flextable, officer, classInt, foreign, stargazer, sf, mapview, leaflet, writexl, lmtest,
       tseries, car, haven, officer, xlsx, openxlsx, httr)



################################################################################
### PARTE 1: Construcción de la base de datos de IC del segundo proceso de muestreo
################################################################################


# 1. Leer archivo remoto desde GitHub

# --- Archivo 1: REPORTE_ADM_UF_COMPET_ACTIVIDADES_X_SUBSECTOR.xlsx
url1 <- "https://github.com/N-SMER/Proyecto-ERA-Emergencias/raw/refs/heads/main/Rawdata/REPORTE_ADM_UF_COMPET_ACTIVIDADES_X_SUBSECTOR.xlsx"
tmp1 <- tempfile(fileext = ".xlsx")
GET(url1, write_disk(tmp1, overwrite = TRUE))

# --- Archivo 2: Administrados_DS.xlsx
url2 <- "https://github.com/N-SMER/Proyecto-ERA-Emergencias/raw/refs/heads/main/Rawdata/Administrados_DS.xlsx"
tmp2 <- tempfile(fileext = ".xlsx")
GET(url2, write_disk(tmp2, overwrite = TRUE))

# --- Archivo 3: Administrados_DS.xlsx
url3 <- "https://github.com/N-SMER/Proyecto-ERA-Emergencias/raw/refs/heads/main/Rawdata/export14102025.xlsx"
tmp3 <- tempfile(fileext = ".xlsx")
GET(url3, write_disk(tmp3, overwrite = TRUE))

# 2. Leer hojas específicas de cada archivo

# Del archivo de REPORTE_ADM_UF_COMPET_ACTIVIDADES_X_SUBSECTOR.xlsx 
uni_adm <- read_excel(tmp1, sheet = "Sheet 1")

names(uni_adm)[names(uni_adm) == "ID_ADM"] <- "ID"
names(uni_adm)[names(uni_adm) == "NUM_DOC"] <- "RUC"
names(uni_adm)[names(uni_adm) == "ADMINISTRADO"] <- "NOMBRE O RAZÓN SOCIAL"

uni_adm <- uni_adm %>%
  select(ID, RUC, `NOMBRE O RAZÓN SOCIAL`, SECTOR,SUBSECTOR, COMPETENCIA, ACTIVIDAD)

uni_adm_collap <- uni_adm %>%
  group_by(RUC) %>%
  summarise(
    NOMBRE = first(`NOMBRE O RAZÓN SOCIAL`),
    SECTOR = first(SECTOR),
    SUBSECTOR = first(SUBSECTOR),
    COMPETENCIA = first(COMPETENCIA),
    ACTIVIDAD = first(ACTIVIDAD)
  )

rm(tmp1, url1)

# Del archivo de Administrados_DS.xlsx
datos <- read_excel(tmp2, sheet = "BD")

names(datos)[names(datos) == "R.U.C."] <- "RUC"

datos <- datos %>%
  select(RUC, CORREO_1 , CORREO_2, NÚMERO,DS )

datos <- datos %>%
  mutate(RUC = as.character(RUC)) %>%
  group_by(RUC) %>%
  summarise(
    CORREO_1 = paste(unique(na.omit(CORREO_1)), collapse = "; "),
    CORREO_2 = paste(unique(na.omit(CORREO_2)), collapse = "; "),
    NÚMERO   = paste(unique(na.omit(NÚMERO)), collapse = "; "),
    DS       = paste(unique(na.omit(DS)), collapse = "; "),
    .groups = "drop"
  )

# Si alguna columna quedó vacía tras el collapse, convertir "" en NA
datos <- datos %>%
  mutate(across(c(CORREO_1, CORREO_2, NÚMERO, DS), ~na_if(.x, "")))

rm(tmp2, url2)

# Del archivo de export14102025.xlsx = ERA EMERGENCIAS
ERA <- read_excel(tmp3, sheet = "Exportar Hoja de Trabajo")

names(ERA)[names(ERA) == "ID_ADM"] <- "ID"
names(ERA)[names(ERA) == "NUM_DOC"] <- "RUC"

ERA <- ERA %>%
  select(ID, RUC)

ERA <- ERA %>%
  select(ID, RUC) %>%
  distinct() %>%
  mutate(USA_ERA = "SI")

rm(tmp3, url3)

# 3. UNION 
uni_adm_collap <- uni_adm_collap %>%
  mutate(RUC = as.character(RUC))

datos <- datos %>%
  mutate(RUC = as.character(RUC))

ERA <- ERA %>%
  mutate(RUC = as.character(RUC))

uni_adm_collap <- uni_adm_collap %>%
  left_join(
    datos %>% select(RUC, CORREO_1, CORREO_2, NÚMERO),
    by = "RUC"
  )

uni_adm_collap <- uni_adm_collap %>%
  left_join(
    ERA %>% select(RUC, USA_ERA),
    by = "RUC"
  )


glimpse(uni_adm_collap)

sum(!is.na(uni_adm_collap$CORREO_1) | 
      !is.na(uni_adm_collap$CORREO_2) |
      !is.na(uni_adm_collap$NÚMERO))


n_distinct(uni_adm$RUC)
n_distinct(datos$RUC)


###################################
###### EXPORTANDO LAS BASES #######
###################################

# Exportamos las dos hojas a un solo archivo Excel
library(writexl)

write_xlsx(
  list(
    "Informes" = uni_adm_collap
   
  ),
  "basefinal.xlsx"
)






