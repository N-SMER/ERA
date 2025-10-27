
rm(list = ls())
gc()
getwd()
setwd("C:/Users/USUARIO/Downloads/ERA/Data")

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
url1 <- "https://github.com/N-SMER/ERA/raw/refs/heads/main/Rawdata/REPORTE_ADM_UF_COMPET_ACTIVIDADES_X_SUBSECTOR.xlsx"
tmp1 <- tempfile(fileext = ".xlsx")
GET(url1, write_disk(tmp1, overwrite = TRUE))

# --- Archivo 5: ERA
url5 <- "https://github.com/N-SMER/ERA/raw/refs/heads/main/Rawdata/export14102025.xlsx"
tmp5 <- tempfile(fileext = ".xlsx")
GET(url5, write_disk(tmp5, overwrite = TRUE))

# --- Archivo 2: DSAP
url2 <- "https://github.com/N-SMER/ERA/raw/refs/heads/main/Rawdata/Directorio%20DSAP%202025_Actualizado.Enero2025.xlsx"
tmp2 <- tempfile(fileext = ".xlsx")
GET(url2, write_disk(tmp2, overwrite = TRUE))

# --- Archivo 3: DSIS
url3 <- "https://github.com/N-SMER/ERA/raw/refs/heads/main/Rawdata/administrados%20DSIS%2029.1.2025.xlsx"
tmp3 <- tempfile(fileext = ".xlsx")
GET(url3, write_disk(tmp3, overwrite = TRUE))

# --- Archivo 4: DEAM
url4 <- "https://github.com/N-SMER/ERA/raw/refs/heads/main/Rawdata/Reporte-Administrados%2030-01-2025.xlsx"
tmp4 <- tempfile(fileext = ".xlsx")
GET(url4, write_disk(tmp4, overwrite = TRUE))

# 2. Leer hojas específicas de cada archivo

# Del archivo de DSAP

CAGR <- read_excel(tmp2, sheet = "CAGR")
CAGR <- CAGR %>% distinct() %>% mutate(Direccion = "DSAP-CAGR")

CIND <- read_excel(tmp2, sheet = "CIND")
CIND <- CIND %>% distinct() %>% mutate(Direccion = "DSAP-CIND")

CPES <- read_excel(tmp2, sheet = "CPES")
CPES <- CPES %>% distinct() %>% mutate(Direccion = "DSAP-CPES")

rm(url2, tmp2)

DSAP <- rbind(CAGR, CIND, CPES)
rm (CAGR, CIND , CPES)

names(DSAP)[names(DSAP) == "RUC/DNI"] <- "RUC"
names(DSAP)[names(DSAP) == "CORREO ELECTRONICO"] <- "CORREO_1"


DSAP <- DSAP %>%
  select(RUC, `DOMICILIO DEL ADMINISTRADO`, `NOMBRE DEL CONTACTO`, CORREO_1, TELEFONO, Direccion) %>%
  filter(
    !is.na(`CORREO_1`),
    !is.na(TELEFONO)
  ) %>%
  distinct()

DSAP <- DSAP %>%
  select(RUC, `DOMICILIO DEL ADMINISTRADO`, `NOMBRE DEL CONTACTO`, CORREO_1, TELEFONO, Direccion) %>%
  mutate(
    # Creamos una variable de prioridad
    prioridad = case_when(
      !is.na(`NOMBRE DEL CONTACTO`) & !is.na(CORREO_1) & !is.na(TELEFONO) ~ 3,
      is.na(`NOMBRE DEL CONTACTO`) & !is.na(CORREO_1) & !is.na(TELEFONO) ~ 2,
      TRUE ~ 1
    )
  ) %>%
  arrange(RUC, desc(prioridad)) %>%     # Ordenamos por RUC y prioridad (mayor primero)
  distinct(RUC, .keep_all = TRUE) %>%   # Nos quedamos con la mejor fila por RUC
  select(-prioridad) 

unique(DSAP$RUC[duplicated(DSAP$RUC)])

# Del archivo de DSIS
DSIS <- read_excel(tmp3, sheet = "ADM_DSIS")
DSIS <- DSIS %>% distinct() %>% mutate(Direccion = "DSIS")
rm(url3, tmp3)

names(DSIS)[names(DSIS) == "NÚMERO DOCUMENTO"] <- "RUC"
names(DSIS)[names(DSIS) == "NÚMERO"] <- "TELEFONO"
DSIS <- DSIS %>% mutate(TELEFONO = as.character(TELEFONO))

DSIS <- DSIS %>%
  select(RUC, CORREO_1, CORREO_2, TELEFONO, Direccion) %>%
  filter(
    !is.na(CORREO_1),
    !is.na(TELEFONO)
  ) %>%
  distinct()

unique(DSIS$RUC[duplicated(DSIS$RUC)])

# Del archivo de DEAM
DEAM <- read_excel(tmp4, sheet = "Sheet1")
DEAM <- DEAM %>% distinct() %>% mutate(Direccion = "DEAM")
rm(url4, tmp4)

names(DEAM)[names(DEAM) == "NUMERO_IDENTIFICACION"] <- "RUC"
names(DEAM)[names(DEAM) == "CORREO ELECTRÓNICO"] <- "CORREO_1"
names(DEAM)[names(DEAM) == "CORREO ELECTRÓNICO A"] <- "CORREO_2"
names(DEAM)[names(DEAM) == "NRO. CELULAR"] <- "TELEFONO"
DEAM <- DEAM %>% mutate(TELEFONO = as.character(TELEFONO))

DEAM <- DEAM %>%
  select(RUC, CORREO_1, CORREO_2, TELEFONO, Direccion) %>%
  filter(
    !is.na(CORREO_1),
    !is.na(TELEFONO)
  ) %>%
  distinct()

unique(DEAM$RUC[duplicated(DEAM$RUC)])

# Unimos la data de las tres direcciones
DSAP <- DSAP %>% mutate(RUC = as.character(RUC))
DSIS <- DSIS %>% mutate(RUC = as.character(RUC))
DEAM <- DEAM %>% mutate(RUC = as.character(RUC))

datos <- bind_rows(DSAP, DSIS, DEAM)
rm (DSAP, DSIS, DEAM)

datos <- datos %>%
  select(RUC, `DOMICILIO DEL ADMINISTRADO`, `NOMBRE DEL CONTACTO`, CORREO_1, CORREO_2, TELEFONO, Direccion) %>%
  mutate(
    # Creamos una variable de prioridad
    prioridad = case_when(
      !is.na(`NOMBRE DEL CONTACTO`) & !is.na(CORREO_1) & !is.na(TELEFONO) ~ 3,
      is.na(`NOMBRE DEL CONTACTO`) & !is.na(CORREO_1) & !is.na(TELEFONO) ~ 2,
      TRUE ~ 1
    )
  ) %>%
  arrange(RUC, desc(prioridad)) %>%     # Ordenamos por RUC y prioridad (mayor primero)
  distinct(RUC, .keep_all = TRUE) %>%   # Nos quedamos con la mejor fila por RUC
  select(-prioridad) 

unique(datos$RUC[duplicated(datos$RUC)])

# Del archivo de REPORTE_ADM_UF_COMPET_ACTIVIDADES_X_SUBSECTOR.xlsx 
uni_adm <- read_excel(tmp1, sheet = "Sheet 1")
names(uni_adm)[names(uni_adm) == "NUM_DOC"] <- "RUC"
uni_adm <- uni_adm %>% mutate(RUC = as.character(RUC))
names(uni_adm)[names(uni_adm) == "ADMINISTRADO"] <- "NOMBRE O RAZÓN SOCIAL"

rm(url1, tmp1)

# Del archivo de export14102025.xlsx = ERA EMERGENCIAS
ERA <- read_excel(tmp5, sheet = "Exportar Hoja de Trabajo")
rm(tmp5, url5)

names(ERA)[names(ERA) == "ID_ADM"] <- "ID"
names(ERA)[names(ERA) == "NUM_DOC"] <- "RUC"
ERA <- ERA %>% mutate(RUC = as.character(RUC))

ERA <- ERA %>% select(RUC)  %>% distinct()
ERA <- ERA %>% mutate(USO_ERA = "SI")

# 3. UNION 
uni_adm <- uni_adm %>% left_join(datos,by = "RUC")

uni_adm <- uni_adm %>% left_join(ERA, by = "RUC")

glimpse(uni_adm)

sum(!is.na(uni_adm$CORREO_1) | 
      !is.na(uni_adm$CORREO_2) |
      !is.na(uni_adm$TELEFONO))

n_distinct(uni_adm$RUC)
n_distinct(uni_adm$CORREO_1)
n_distinct(uni_adm$TELEFONO)
n_distinct(datos$RUC)


###################################
###### EXPORTANDO LAS BASES #######
###################################

# Exportamos las dos hojas a un solo archivo Excel
library(writexl)

write_xlsx(
  list(
    "Informes" = uni_adm
   
  ),
  "basefinal.xlsx"
)






