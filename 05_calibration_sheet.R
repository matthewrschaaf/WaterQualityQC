# --- Load Required Packages ---
library(dplyr)
library(readxl)
library(tidyr)
library(purrr)
library(tibble)
library(stringr)

generate_calibration_sheet <- function(results_df, analytes_df, user_input) {
  
  message("Generating calibration tracking sheet...")
  
  # --- 1. Define Constants and File Paths ---
  
  calibration_log_path <- "O:/ED/Private/Water Quality/Data/FY2025/2025 Field/Calibration Log/Calibration Logs.xlsx"
  cal_params <- c("00094", "00400", "00090", "00299", "00301", "32241", "32244", "00076")
  
  # --- 2. Read All Calibration Log Sheets ---
  message("-> Reading calibration log file...")
  tryCatch({  
    cal_do <- read_excel(calibration_log_path, sheet = "DO", range = "A6:I1000",col_names = c("Analyst", "Date", "Time", "Barom_Pressure_mmHg", "Temp", "Initial_DO_pct", "Initial_DO_mgL", "Calibrated_DO_pct", "Calibrated_DO_mgL"))
    cal_weekly <- read_excel(calibration_log_path, sheet = "Weekly", skip = 5, col_names = c("Analyst", "Date", "Time", "SC_Temp", "Initial_SC", "Standard_SC", "Calibrated_SC", "pH_Temp", "pH7_Initial", "pH7_Standard", "pH7_Calibrated", "pH4_Initial", "pH4_Standard", "pH4_Calibrated", "pH10_Initial", "pH10_Standard", "pH10_Calibrated"))
    cal_2month_torp <- read_excel(calibration_log_path, sheet = "Every 2 Months - Turb & ORP", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "ORP_Temp", "ORP_Initial", "ORP_Standard", "ORP_Calibrated", "Turb_Temp", "Turb_0_Initial", "Turb_0_Calibrated", "Turb_124_Initial", "Turb_124_Calibrated"))
    cal_2month_chl <- read_excel(calibration_log_path, sheet = "Every 2 Months - Chl", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "PC_RFU_Temp", "PC_RFU_0_Init", "PC_RFU_0_Post", "PC_RFU_Initial", "PC_RFU_Standard", "PC_RFU_Calibrated", "Chl_RFU_0_Init", "Chl_RFU_0_Post", "Chl_RFU_Initial", "Chl_RFU_Standard", "Chl_RFU_Calibrated", "Chl_ugL_0_Init", "Chl_ugL_0_Post", "Chl_ugL_Initial", "Chl_ugL_Standard", "Chl_ugL_Calibrated"))
  }, error = function(e) {
    stop(paste("Failed to read calibration log file. Check path and sheet names.", e$message), call. = FALSE)
  })
  
  # Clean and prepare calibration dates
  do_dates <- cal_do %>% filter(!is.na(Date)) %>% transmute(DO_Cal_Date = as.Date(Date, format = "%d/%m/%Y"))
  weekly_dates <- cal_weekly %>% filter(!is.na(Date)) %>% transmute(Weekly_Cal_Date = as.Date(Date, format = "%d/%m/%Y"))
  orp_turb_dates <- cal_2month_torp %>% filter(!is.na(Cal_Date)) %>% transmute(ORP_Turb_Cal_Date = as.Date(Cal_Date, format = "%d/%m/%Y"))
  chl_pc_dates <- cal_2month_chl %>% filter(!is.na(Cal_Date)) %>% transmute(Chl_PC_Cal_Date = as.Date(Cal_Date, format = "%d/%m/%Y"))
  
  
  # --- 3. Identify Unique Sampling Trips (by Date) and Parameters Collected ---
  message("-> Identifying unique field sampling days and collected parameters...")
  
  param_names_lookup <- analytes_df %>%
    select(Storet_Num, Parameter) %>%
    distinct()
  
  sampling_days <- results_df %>%
    filter(
      Field_Lab_ID == "FIELD",
      Storet_Num %in% cal_params
    ) %>%
    mutate(Sample_Date = as.Date(substr(Assoc_Samp, 1, 8), format = "%Y%m%d")) %>%
    filter(Sample_Date >= user_input$start_date & Sample_Date <= user_input$end_date) %>%
    left_join(param_names_lookup, by = "Storet_Num") %>%
    mutate(Parameter = str_trim(Parameter)) %>%
    group_by(Sample_Date) %>%
    summarise(
      Collected_Parameters = paste(sort(unique(Parameter)), collapse = ", "), 
      Storet_Collected = paste(sort(unique(Storet_Num)), collapse = ", "),
      .groups = "drop"
    ) %>%
    arrange(Sample_Date)
  
  
  # --- 4. Match Sampling Trips to Latest Valid Calibration Dates ---
  message("-> Matching trips to the latest valid calibration dates...")
  
  calibration_sheet <- purrr::map_df(1:nrow(sampling_days), function(i) {
    trip <- sampling_days[i, ]
    collected_storets <- trip$Storet_Collected
    
  
  # --- Find the latest calibration date for each parameter IF it was collected ---
  
  # Specific Conductance (00094) - Weekly
  sc_cal_date <- if (grepl("00094", collected_storets)) {
    # Find all weekly cals within 7 days prior to the trip
    valid_dates <- weekly_dates %>%
      filter(Weekly_Cal_Date <= trip$Sample_Date & Weekly_Cal_Date >= (trip$Sample_Date - 7))
    # Return the single most recent date
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, Weekly_Cal_Date)$Weekly_Cal_Date else NA
  } else {
    NA
  }
  
  # pH (00400) - Weekly
  ph_cal_date <- if (grepl("00400", collected_storets)) {
    valid_dates <- weekly_dates %>%
      filter(Weekly_Cal_Date <= trip$Sample_Date & Weekly_Cal_Date >= (trip$Sample_Date - 7))
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, Weekly_Cal_Date)$Weekly_Cal_Date else NA
  } else {
    NA
  }
  
  # DO/Oxygen Saturation (00299, 00301) - Same Day
  do_cal_date <- if (grepl("00299", collected_storets) || grepl("00301", collected_storets)) {
    valid_dates <- do_dates %>%
      filter(DO_Cal_Date == trip$Sample_Date)
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, DO_Cal_Date)$DO_Cal_Date else NA
  } else {
    NA
  }
  
  # ORP (00090) - 60 Days
  orp_cal_date <- if (grepl("00090", collected_storets)) {
    valid_dates <- orp_turb_dates %>%
      filter(ORP_Turb_Cal_Date <= trip$Sample_Date & ORP_Turb_Cal_Date >= (trip$Sample_Date - 60))
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, ORP_Turb_Cal_Date)$ORP_Turb_Cal_Date else NA
  } else {
    NA
  }
  
  # Chlorophyll (32241, 32244) - 60 Days
  chl_cal_date <- if (grepl("32241", collected_storets) || grepl("32244", collected_storets)) {
    valid_dates <- chl_pc_dates %>%
      filter(Chl_PC_Cal_Date <= trip$Sample_Date & Chl_PC_Cal_Date >= (trip$Sample_Date - 60))
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, Chl_PC_Cal_Date)$Chl_PC_Cal_Date else NA
  } else {
    NA
  }
  
  # Phycocyanin (32242) - 60 Days
  phyco_cal_date <- if (grepl("32242", collected_storets)) {
    valid_dates <- chl_pc_dates %>%
      filter(Chl_PC_Cal_Date <= trip$Sample_Date & Chl_PC_Cal_Date >= (trip$Sample_Date - 60))
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, Chl_PC_Cal_Date)$Chl_PC_Cal_Date else NA
  } else {
    NA
  }
  
  # Turbidity (00076) - 60 Days
  turb_cal_date <- if (grepl("00076", collected_storets)) {
    valid_dates <- orp_turb_dates %>%
      filter(ORP_Turb_Cal_Date <= trip$Sample_Date & ORP_Turb_Cal_Date >= (trip$Sample_Date - 60))
    if (nrow(valid_dates) > 0) top_n(valid_dates, 1, ORP_Turb_Cal_Date)$ORP_Turb_Cal_Date else NA
  } else {
    NA
  }
  
  # --- Assemble the Row for the Final Table ---
  tibble::tibble(
    Sample_Date = trip$Sample_Date,
    Collected_Parameters = trip$Collected_Parameters,
    Storet_Collected = trip$Storet_Collected,
    `Spec_Cond_Cal_Date` = sc_cal_date,
    `pH_Cal_Date` = ph_cal_date,
    `DO_Cal_Date` = do_cal_date,
    `ORP_Cal_Date` = orp_cal_date,
    `Chl_Cal_Date` = chl_cal_date,
    `Phyco_Cal_Date` = phyco_cal_date,
    `Turb_Cal_Date` = turb_cal_date
  )
})

message("-> Final calibration sheet populated successfully.")

# --- Bundle all dataframes into a single list for return ---

raw_cal_logs <- list(
  DO_Log = cal_do,
  Weekly_Log = cal_weekly,
  Turb_ORP_Log = cal_2month_torp,
  Chl_Log = cal_2month_chl
)

cleaned_cal_dates <- list(
  DO_Dates = do_dates,
  Weekly_Dates = weekly_dates,
  ORP_Turb_Dates = orp_turb_dates,
  Chl_PC_Dates = chl_pc_dates
)

final_cal_list <- list(
  Raw_Calibration_Logs = raw_cal_logs,
  Cleaned_Calibration_Dates = cleaned_cal_dates,
  Identified_Sampling_Days = sampling_days,
  Final_Calibration_Sheet = calibration_sheet
)

return(final_cal_list)
}