# --- Load Required Packages ---
library(dplyr)
library(readxl)
library(tidyr)
library(purrr)
library(tibble)


generate_calibration_sheet <- function(results_df, analytes_df, user_input) {
  
  message("Generating calibration tracking sheet...")
  
  # --- 1. Define Constants and File Paths ---
  
  # IMPORTANT: This file path may need to be updated in the future.
  calibration_log_path <- "O:/ED/Private/Water Quality/Data/FY2025/2025 Field/Calibration Log/Calibration Logs.xlsx"
  
  # Define the parameters and their Storet Numbers that require calibration
  cal_params <- c("00094", "00400", "00090", "00299", "00301", "32241", "32244", "00076")
  
  # --- 2. Read All Calibration Log Sheets ---
  message("-> Reading calibration log file...")
  
  tryCatch({
    cal_do <- read_excel(calibration_log_path, sheet = "DO", skip = 5, col_names = c("Analyst", "Date", "Time", "Barom_Pressure_mmHg", "Temp", "Initial_DO_pct", "Initial_DO_mgL", "Calibrated_DO_pct", "Calibrated_DO_mgL"), range = cell_cols("A:I"))
    cal_weekly <- read_excel(calibration_log_path, sheet = "Weekly", skip = 5, col_names = c("Analyst", "Date", "Time", "SC_Temp", "Initial_SC", "Standard_SC", "Calibrated_SC", "pH_Temp", "pH7_Initial", "pH7_Standard", "pH7_Calibrated", "pH4_Initial", "pH4_Standard", "pH4_Calibrated", "pH10_Initial", "pH10_Standard", "pH10_Calibrated"))
    cal_2month_torp <- read_excel(calibration_log_path, sheet = "Every 2 Months - Turb & ORP", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "ORP_Temp", "ORP_Initial", "ORP_Standard", "ORP_Calibrated", "Turb_Temp", "Turb_0_Initial", "Turb_0_Calibrated", "Turb_124_Initial", "Turb_124_Calibrated"))
    cal_2month_chl <- read_excel(calibration_log_path, sheet = "Every 2 Months - Chl", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "PC_RFU_Temp", "PC_RFU_0_Init", "PC_RFU_0_Post", "PC_RFU_Initial", "PC_RFU_Standard", "PC_RFU_Calibrated", "Chl_RFU_0_Init", "Chl_RFU_0_Post", "Chl_RFU_Initial", "Chl_RFU_Standard", "Chl_RFU_Calibrated", "Chl_ugL_0_Init", "Chl_ugL_0_Post", "Chl_ugL_Initial", "Chl_ugL_Standard", "Chl_ugL_Calibrated"))
  }, error = function(e) {
    stop(paste("Failed to read calibration log file. Check path and sheet names.", e$message), call. = FALSE)
  })
  
  # Clean and prepare calibration dates
  do_dates <- cal_do %>% 
    filter(!is.na(Date)) %>% 
    transmute(DO_Cal_Date = as.Date(Date, format = "%d/%m/%Y"))
  
  weekly_dates <- cal_weekly %>% 
    filter(!is.na(Date)) %>% 
    transmute(Weekly_Cal_Date = as.Date(Date, format = "%d/%m/%Y"))
  
  orp_turb_dates <- cal_2month_torp %>% 
    filter(!is.na(Cal_Date)) %>% 
    transmute(ORP_Turb_Cal_Date = as.Date(Cal_Date, format = "%d/%m/%Y"))
  
  chl_pc_dates <- cal_2month_chl %>% 
    filter(!is.na(Cal_Date)) %>% 
    transmute(Chl_PC_Cal_Date = as.Date(Cal_Date, format = "%d/%m/%Y"))

  
# --- 3. Identify Unique Sampling Trips (by Date) and Parameters Collected ---
  message("-> Identifying unique field sampling days and collected parameters...")
  
  # Create a small lookup table for parameter names
  param_names_lookup <- analytes_df %>%
    select(Storet_Num, anl_name) %>%
    distinct()
  

  sampling_days <- results_df %>%

    filter(
      Lab_ID == "FIELD",
      Storet_Num %in% cal_params
    ) %>%
    mutate(Sample_Date = as.Date(substr(Assoc_Samp, 1, 8), format = "%Y%m%d")) %>%
    filter(Sample_Date >= user_input$start_date & Sample_Date <= user_input$end_date) %>%
    left_join(param_names_lookup, by = "Storet_Num") %>%
    mutate(anl_name = str_trim(anl_name)) %>%
    group_by(Sample_Date) %>%
    summarise(Collected_Parameters = paste(sort(unique(anl_name)), collapse = ", "), .groups = "drop") %>%
    arrange(Sample_Date)
  
  # The 'sampling_days' dataframe now contains one row for each day a sample
  # was taken, and a complete list of what was measured on that day.
  # You can print the first few rows to check it:
  # print(head(sampling_days))