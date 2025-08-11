# --- Load Required Packages ---
library(dplyr)
library(readxl)
library(tidyr)


generate_calibration_sheet <- function(results_df, analytes_df, user_input) {
  
  message("Generating calibration tracking sheet...")
  
  # --- 1. Define Constants and File Paths ---
  
  # IMPORTANT: This file path may need to be updated in the future.
  calibration_log_path <- "O:/ED/Private/Water Quality/Data/FY2025/2025 Field/Calibration Log.xlsx"
  
  # Define the parameters and their Storet Numbers that require calibration
  cal_params <- c("00094", "00400", "00090", "00299", "00301", "32241", "32244", "00076")
  
  # --- 2. Read All Calibration Log Sheets ---
  message("-> Reading calibration log file...")
  
  # Use tryCatch to handle potential errors if a sheet is missing or named incorrectly
  tryCatch({
    cal_do <- read_excel(calibration_log_path, sheet = "DO", skip = 5, col_names = c("Analyst", "Date", "Time", "Barom_Pressure_mmHg", "Temp", "Initial_DO_pct", "Initial_DO_mgL", "Calibrated_DO_pct", "Calibrated_DO_mgL"))
    cal_weekly <- read_excel(calibration_log_path, sheet = "Weekly", skip = 5, col_names = c("Analyst", "Date", "Time", "SC_Temp", "Initial_SC", "Standard_SC", "Calibrated_SC", "pH_Temp", "pH7_Initial", "pH7_Standard", "pH7_Calibrated", "pH4_Initial", "pH4_Standard", "pH4_Calibrated", "pH10_Initial", "pH10_Standard", "pH10_Calibrated"))
    cal_2month_torp <- read_excel(calibration_log_path, sheet = "Every 2 Months - Turb & ORP", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "ORP_Temp", "ORP_Initial", "ORP_Standard", "ORP_Calibrated", "Turb_Temp", "Turb_0_Initial", "Turb_0_Calibrated", "Turb_124_Initial", "Turb_124_Calibrated"))
    cal_2month_chl <- read_excel(calibration_log_path, sheet = "Every 2 Months - Chl", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "PC_RFU_Temp", "PC_RFU_0_Init", "PC_RFU_0_Post", "PC_RFU_Initial", "PC_RFU_Standard", "PC_RFU_Calibrated", "Chl_RFU_0_Init", "Chl_RFU_0_Post", "Chl_RFU_Initial", "Chl_RFU_Standard", "Chl_RFU_Calibrated", "Chl_ugL_0_Init", "Chl_ugL_0_Post", "Chl_ugL_Initial", "Chl_ugL_Standard", "Chl_ugL_Calibrated"))
  }, error = function(e) {
    stop(paste("Failed to read calibration log file. Check path and sheet names.", e$message), call. = FALSE)
  })
  
  # Clean and prepare calibration dates
  do_dates <- cal_do %>% filter(!is.na(Date)) %>% transmute(DO_Cal_Date = as.Date(Date))
  weekly_dates <- cal_weekly %>% filter(!is.na(Date)) %>% transmute(Weekly_Cal_Date = as.Date(Date))
  orp_turb_dates <- cal_2month_torp %>% filter(!is.na(Cal_Date)) %>% transmute(ORP_Turb_Cal_Date = as.Date(Cal_Date))
  chl_pc_dates <- cal_2month_chl %>% filter(!is.na(Cal_Date)) %>% transmute(Chl_PC_Cal_Date = as.Date(Cal_Date))
  
  # --- 3. Identify Unique Sampling Trips ---
  message("-> Identifying unique field sampling trips...")
  
  sampling_trips <- results_df %>%
    filter(
      LAB_ID == "FIELD",
      Storet_Num %in% cal_params
    ) %>%
    mutate(Sample_Date = as.Date(substr(Assoc_Samp, 1, 8), format = "%Y%m%d")) %>%
    filter(Sample_Date >= user_input$start_date & Sample_Date <= user_input$end_date) %>%
    distinct(Loc_ID, Assoc_Samp, Sample_Date) %>%
    arrange(Sample_Date, Loc_ID)
  
  # --- 4. Match Sampling Trips to Latest Valid Calibration Dates ---
  message("-> Matching trips to the latest valid calibration dates...")
  
  # This is a complex operation. We use map_df to iterate through each sampling trip.
  calibration_sheet <- purrr::map_df(1:nrow(sampling_trips), function(i) {
    trip <- sampling_trips[i, ]
    
    # Find the latest valid calibration date for each parameter type
    latest_weekly <- weekly_dates %>% filter(Weekly_Cal_Date <= trip$Sample_Date & Weekly_Cal_Date >= (trip$Sample_Date - 7)) %>% top_n(1, Weekly_Cal_Date)
    latest_do <- do_dates %>% filter(DO_Cal_Date == trip$Sample_Date) %>% top_n(1, DO_Cal_Date)
    latest_orp_turb <- orp_turb_dates %>% filter(ORP_Turb_Cal_Date <= trip$Sample_Date & ORP_Turb_Cal_Date >= (trip$Sample_Date - 60)) %>% top_n(1, ORP_Turb_Cal_Date)
    latest_chl_pc <- chl_pc_dates %>% filter(Chl_PC_Cal_Date <= trip$Sample_Date & Chl_PC_Cal_Date >= (trip$Sample_Date - 60)) %>% top_n(1, Chl_PC_Cal_Date)
    
    # Build a single row for the final table
    tibble(
      Loc_ID = trip$Loc_ID,
      Sample_Num = trip$Assoc_Samp,
      Sample_Date = trip$Sample_Date,
      `Specific_Conductivity_Cal_Date` = if(nrow(latest_weekly) > 0) latest_weekly$Weekly_Cal_Date else NA_Date_,
      `pH_Cal_Date` = if(nrow(latest_weekly) > 0) latest_weekly$Weekly_Cal_Date else NA_Date_,
      `DO_Cal_Date` = if(nrow(latest_do) > 0) latest_do$DO_Cal_Date else NA_Date_,
      `ORP_Cal_Date` = if(nrow(latest_orp_turb) > 0) latest_orp_turb$ORP_Turb_Cal_Date else NA_Date_,
      `Chlorophyll_Cal_Date` = if(nrow(latest_chl_pc) > 0) latest_chl_pc$Chl_PC_Cal_Date else NA_Date_,
      `Phycocyanin_Cal_Date` = if(nrow(latest_chl_pc) > 0) latest_chl_pc$Chl_PC_Cal_Date else NA_Date_,
      `Turbidity_Cal_Date` = if(nrow(latest_orp_turb) > 0) latest_orp_turb$ORP_Turb_Cal_Date else NA_Date_
    )
  }) %>%
    # Add a placeholder "Collected?" column
    mutate(Collected = "Yes", .before = `Specific_Conductivity_Cal_Date`)
  
  message("-> Calibration tracking sheet created successfully with ", nrow(calibration_sheet), " rows.")
  
  return(calibration_sheet)
}