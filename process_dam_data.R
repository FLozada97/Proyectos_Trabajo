# --- 1. Load Libraries ---
library(readxl)
library(dplyr)
library(openxlsx)
library(here)
library(janitor)
library(zoo)
library(tidyr)

# --- 2. Configuration & File Paths ---
# Using here() builds file paths relative to the project's root folder.
# This makes the script work on any computer without changing the paths manually.
# It assumes your project has the following structure:
# your_project_folder/
# ├── data/
# │   ├── base_presas_completa.xlsx
# │   └── dam_name_corrections.csv
# ├── output/
# └── process_dam_data.R

# Input files are in the 'data' folder
data_dir <- here("data")
# Output files will be saved in the 'output' folder
output_dir <- here("output")

# --- 3. Load and Prepare the Main Dataset ---

#In this part, the user chooses whether to run option 1 or option 2.
#Afterwards, they can run all the code without any problems.

##Option 1
# Load the main Excel file
# The clean_names() function automatically converts column names to a standard,
# easy-to-use format (e.g., "Nombre común" becomes "nombre_comun").
tryCatch({
  data <- read_excel(file.path(data_dir, "base_presas_completa.xlsx")) %>%
    clean_names()
}, error = function(e) {
  stop("Failed to load 'base_presas_completa.xlsx'. Make sure the file is in the 'data' directory. Error: ", e$message)
})

##Option 2
## En caso de actualización de meses por separado, se cargan los meses por separado y después se consolida
#Cargar datos a utilizar--> se cargan las bases divididas en años o en meses que se obtuvieron previamente de la base grande
data_01 <- read_excel("data/data_presas_01.xlsx")
data_02 <- read_excel("data/data_presas_02.xlsx")
data_03 <- read_excel("data/data_presas_03.xlsx")
data_04 <- read_excel("data/data_presas_04.xlsx")
data_05 <- read_excel("data/data_presas_05.xlsx")
data_06 <- read_excel("data/data_presas_06.xlsx")
data_07 <- read_excel("data/data_presas_07.xlsx")
data_08 <- read_excel("data/data_presas_08.xlsx")
data_09 <- read_excel("data/data_presas_09.xlsx")
data_10 <- read_excel("data/data_presas_10.xlsx")
data_11 <- read_excel("data/data_presas_11.xlsx")
data_12 <- read_excel("data/data_presas_12.xlsx")
data <- bind_rows(data_01, data_02, data_03, data_04, data_05, data_06, data_07, data_08, data_09, data_10, data_11, data_12)

# --- 4. Apply Data Corrections from the Dictionary ---
# This section replaces the long series of 'ifelse' statements from the original script.
# All corrections are now managed in the 'dam_name_corrections.csv' file.
# To make a new correction, just add a new row to that CSV file.

print("Limpieza inicia")
corrections_df <- read.csv(file.path(data_dir, "dam_name_corrections.csv"))

# Loop through each rule in the corrections dataframe and apply it
for (i in 1:nrow(corrections_df)) {
  match_col <- as.character(corrections_df$match_column[i])
  match_val <- corrections_df$match_value[i]
  target_col <- as.character(corrections_df$target_column[i])
  corrected_val <- corrections_df$corrected_value[i]

  # Clean the column names from the CSV to match the cleaned names in the main dataframe
  match_col_clean <- make_clean_names(match_col)
  target_col_clean <- make_clean_names(target_col)

  # Find the rows that match the condition
  if (match_col_clean %in% names(data)) {
    rows_to_update <- data[[match_col_clean]] == match_val & !is.na(data[[match_col_clean]])

    # Update the target column with the corrected value for those rows
    if (target_col_clean %in% names(data)) {
      data[rows_to_update, target_col_clean] <- corrected_val
    }
  }
}

## Consolidar base con variables auxiliares----
data_total <- data
data_total <- data_total[order(data_total$fecha, decreasing = FALSE), ]
data_total$id_num <- seq_len(nrow(data_total))
data_total <- data_total[, c("id_num", setdiff(names(data_total), "id_num"))]
data_total_o <- data.frame(data_total)
data_total$id_error_max <- if_else(data_total$percent_de_llenado_actual >= 115, 1, 0)

# --- 5. Generate Catalogs and Lists ---
# Create a list of unique dam names from the corrected data
nom_com_corr <- c(unique(data_total$nombre_comun))
# Match with other relevant data to create a catalog
lista_de_presas <- data_total %>%
  distinct(nombre_comun, .keep_all = TRUE) %>%
  select(
    nombre_de_presa,
    nombre_comun,
    organismo_de_cuenca,
    entidad_federativa,
    region,
    name_elevacion_msnm
  ) %>%
  arrange(organismo_de_cuenca)

# The logic for 'filtro_data_or' and 'filtro_data_corr' is complex and seems
# intended for manual verification. It is preserved here for that purpose.
nom_com_or<- c(unique(data_total_o$nombre_comun))
filtro_id_or<- data_total_o$id_num[match(nom_com_or, data_total_o$nombre_comun)]
nom_com_or<-data_total_o$nombre_comun[match(filtro_id_or, data_total_o$id_num)]
nom_pre_or<-data_total_o$nombre_de_presa[match(filtro_id_or, data_total_o$id_num)]
nom_com_corr_filtered<-data_total$nombre_comun[match(filtro_id_or, data_total$id_num)]
nom_pre_corr_filtered<-data_total$nombre_de_presa[match(filtro_id_or, data_total$id_num)]

filtro_data_or<- data.frame(filtro_id_or,nom_com_or,nom_com_corr_filtered,nom_pre_or,nom_pre_corr_filtered)
filtro_data_or <- filtro_data_or[order(filtro_data_or$nom_com_or, decreasing = FALSE), ]


# --- 6. Auxiliaries for 'agro' and 'quincena' ---
# This part depends on a file named 'id_data.xlsx'
tryCatch({
  id_agro <- read_excel(file.path(data_dir, "id_data.xlsx"), sheet = "agro") %>% clean_names()
  id_quincena <- read_excel(file.path(data_dir, "id_data.xlsx"), sheet = "quincena") %>% clean_names()

  data_total$aux_agro <- id_agro$agro[match(data_total$nombre_comun, id_agro$nom_com_corr)]
  data_total$aux_quincena <- id_quincena$id_quincena[match(data_total$dia, id_quincena$id_dia)]
  data_total$aux_quincena[data_total$dia == 15 & data_total$mes == 2] <- 2
}, warning = function(w) {
  message("Warning: 'id_data.xlsx' not found. Skipping the 'agro' and 'quincena' auxiliary steps.")
})

print("Termina Limpieza")

# --- 7. Original Script Logic (with adaptations for new column names) ---

## División y agrupación de bases----
### Agrupación en meses bases limpias
meses <- 1:12
for (mes in meses) {
  data_filtrada <- data_total %>% filter(mes == !!mes)
  nombre_archivo <- paste0("data_presas_", sprintf("%02d", mes),".xlsx")
  # Save to the 'output' directory
  write.xlsx(data_filtrada, file = file.path(output_dir, nombre_archivo))
}

### Agrupación en años bases limpias
anos <- 2014:2025
for (ano_val in anos) {
  data_filtrada <- data_total %>% filter(ano == !!ano_val)
  nombre_archivo <- paste0("data_presas_", ano_val, ".xlsx")
  # Save to the 'output' directory
  write.xlsx(data_filtrada, file = file.path(output_dir, nombre_archivo))
}

# --- 8. Save Main Output Files ---
# Save the fully cleaned and consolidated database
wb <- createWorkbook()
addWorksheet(wb, "data")
writeData(wb, sheet = "data", data_total)
saveWorkbook(wb, file = file.path(output_dir, "base_presas_completa.xlsx"), overwrite = TRUE)

# Save the catalogs
cat_wb <- createWorkbook()
addWorksheet(cat_wb, "catalogo_1_original")
addWorksheet(cat_wb, "Lista_de_presas")
writeData(cat_wb, sheet = "catalogo_1_original", filtro_data_or)
writeData(cat_wb, sheet = "Lista_de_presas", lista_de_presas)
saveWorkbook(cat_wb, file = file.path(output_dir, "catalogo_base_presas_completa.xlsx"), overwrite = TRUE)


# --- 9. Aggregation and Analysis Function ---
# This is the powerful function from the end of the original script,
# adapted to use the new cleaned column names.

procesar_y_exportar_agregado <- function(data_in, group_col, file_out) {
  # The function expects the column name as a string
  group_col_sym <- sym(group_col)

  df_agregado <- data_in %>%
    group_by(mes, ano, {{ group_col_sym }}) %>%
    summarise(
      prom_namo_alm = sum(namo_almacenamiento_hm3, na.rm = TRUE),
      prom_alm_act = sum(almacenamiento_actual_hm3, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      alm_prom = (prom_alm_act / prom_namo_alm) * 100,
      fecha = as.yearmon(paste(ano, mes), "%Y %m")
    ) %>%
    arrange({{ group_col_sym }}, fecha)

  # Pivot to wide format for Excel output
  df_pivot <- df_agregado %>%
    pivot_wider(
      id_cols = fecha,
      names_from = {{ group_col_sym }},
      values_from = c(prom_namo_alm, prom_alm_act, alm_prom)
    ) %>%
    arrange(fecha)

  # Separate into the 3 dataframes for Excel sheets
  df_prom_namo <- df_pivot %>% select(fecha, starts_with("prom_namo_alm_"))
  df_prom_act <- df_pivot %>% select(fecha, starts_with("prom_alm_act_"))
  df_alm_prom <- df_pivot %>% select(fecha, starts_with("alm_prom_"))

  # Save to Excel
  wb <- createWorkbook()
  addWorksheet(wb, "data_long_format")
  addWorksheet(wb, "NAMO_Alm")
  addWorksheet(wb, "Alm_Act")
  addWorksheet(wb, "Alm_Prom")

  writeData(wb, sheet = "data_long_format", df_agregado)
  writeData(wb, sheet = "NAMO_Alm", df_prom_namo)
  writeData(wb, sheet = "Alm_Act", df_prom_act)
  writeData(wb, sheet = "Alm_Prom", df_alm_prom)

  saveWorkbook(wb, file = file.path(output_dir, file_out), overwrite = TRUE)
  print(paste("Aggregation saved to:", file.path(output_dir, file_out)))
}


# --- 10. Execute the Aggregation ---
# Process by "Entidad Federativa"
procesar_y_exportar_agregado(
  data_in = data_total,
  group_col = "entidad_federativa",
  file_out = "base_presas_mes_entidad.xlsx"
)

# Process by "Región"
procesar_y_exportar_agregado(
  data_in = data_total,
  group_col = "region",
  file_out = "base_presas_mes_region.xlsx"
)


# --- 11. Final Combined Entity and Region Export ---
alm_ent_reg <- data_total %>%
  group_by(mes, ano, entidad_federativa, region) %>%
  summarise(
    sum_namo_alm = sum(namo_almacenamiento_hm3, na.rm = TRUE),
    sum_alm_act = sum(almacenamiento_actual_hm3, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    alm_prom = (sum_alm_act / sum_namo_alm) * 100,
    fecha = as.Date(paste(ano, mes, "01"), format = "%Y %m %d")
  )

# Save this final file
wb_ent_reg <- createWorkbook()
addWorksheet(wb_ent_reg, "data")
writeData(wb_ent_reg, sheet = "data", alm_ent_reg)
saveWorkbook(wb_ent_reg, file = file.path(output_dir, "base_presas_mes_ent_cesf.xlsx"), overwrite = TRUE)

print("Script refactoring complete. All output files have been saved to the 'output' directory.")
