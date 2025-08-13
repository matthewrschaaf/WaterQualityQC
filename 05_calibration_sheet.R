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
  cal_params <- c("00094", "00400", "00090", "00299", "00301", "32241", "32244", "32242", "00076")
  
  # --- 2. Read All Calibration Log Sheets ---
  message("-> Reading calibration log file...")
  tryCatch({
    cal_DO <- read_excel(calibration_log_path, sheet = "DO", skip = 5, col_names = c("Analyst", "Date", "Time", "Barom_Pressure_mmHg", "Temp", "Initial_DO_mgL", "Initial_DO_pct", "Calibrated_DO_mgL", "Calibrated_DO_pct"))
    cal_SpCond <- read_excel(calibration_log_path, sheet = "Specific Conductance", skip = 5, col_names = c("Analyst", "Date", "Time", "SC_Temp", "Initial_SC", "Standard_SC", "Calibrated_SC"))
    cal_pH <- read_excel(calibration_log_path, sheet = "pH", skip = 5, col_names = c("Analyst", "Date", "Time", "pH_Temp", "pH7_Initial", "pH7_Standard", "pH7_Calibrated", "pH4_Initial", "pH4_Standard", "pH4_Calibrated", "pH10_Initial", "pH10_Standard", "pH10_Calibrated"))
    cal_ORP <- read_excel(calibration_log_path, sheet = "ORP", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "ORP_Temp", "ORP_Initial", "ORP_Standard", "ORP_Calibrated"))
    cal_Turbidity <- read_excel(calibration_log_path, sheet = "Turbidity", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "Turb_Temp", "Turb_0_Initial", "Turb_0_Calibrated", "Turb_124_Initial", "Turb_124_Calibrated"))
    cal_Phycocyanin <- read_excel(calibration_log_path, sheet = "Phycocyanin", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "PC_RFU_Temp", "PC_RFU_0_Init", "PC_RFU_0_Post", "PC_RFU_Initial", "PC_RFU_Standard", "PC_RFU_Calibrated"))
    cal_Chlorophyll <- read_excel(calibration_log_path, sheet = "Chlorophyll", skip = 5, col_names = c("Analyst", "Cal_Date", "Time", "Next_Cal_Date", "Chl_Temp", "Chl_RFU_0_Init", "Chl_RFU_0_Post", "Chl_RFU_Initial", "Chl_RFU_Standard", "Chl_RFU_Calibrated", "Chl_ugL_0_Init", "Chl_ugL_0_Post", "Chl_ugL_Initial", "Chl_ugL_Standard", "Chl_ugL_Calibrated"))
    Chlorophyll_Standards <- read_excel(calibration_log_path, sheet = "Chlorophyll Standards Table")
    Standards <- read_excel(calibration_log_path, sheet = "Standards")
  }, error = function(e) {
    stop(paste("Failed to read calibration log file. Check path and sheet names.", e$message), call. = FALSE)
  })
  
  # Clean and prepare calibration dates by removing the 'format' argument ***
  DO_dates <- cal_DO %>% filter(!is.na(Date)) %>% transmute(do_date = as.Date(Date))
  SpCond_dates <- cal_SpCond %>% filter(!is.na(Date)) %>% transmute(sc_date = as.Date(Date))
  pH_dates <- cal_pH %>% filter(!is.na(Date)) %>% transmute(ph_date = as.Date(Date))
  ORP_dates <- cal_ORP %>% filter(!is.na(Cal_Date)) %>% transmute(orp_date = as.Date(Cal_Date))
  Turbidity_dates <- cal_Turbidity %>% filter(!is.na(Cal_Date)) %>% transmute(turb_date = as.Date(Cal_Date))
  Phycocyanin_dates <- cal_Phycocyanin %>% filter(!is.na(Cal_Date)) %>% transmute(phyco_date = as.Date(Cal_Date))
  Chlorophyll_dates <- cal_Chlorophyll %>% filter(!is.na(Cal_Date)) %>% transmute(chl_date = as.Date(Cal_Date))
  
  
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
  
  full_calibration_sheet <- purrr::map_df(1:nrow(sampling_days), function(i) {
    trip <- sampling_days[i, ]
    collected_storets <- trip$Storet_Collected
    
    # --- Find the latest calibration date for each parameter IF it was collected ---
    
    sc_cal_date <- if (grepl("00094", collected_storets)) {
      valid_dates <- SpCond_dates %>% filter(sc_date <= trip$Sample_Date & sc_date >= (trip$Sample_Date - 7))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, sc_date)$sc_date else NA
    } else { NA }
    
    ph_cal_date <- if (grepl("00400", collected_storets)) {
      valid_dates <- pH_dates %>% filter(ph_date <= trip$Sample_Date & ph_date >= (trip$Sample_Date - 7))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, ph_date)$ph_date else NA
    } else { NA }
    
    do_cal_date <- if (grepl("00299", collected_storets) || grepl("00301", collected_storets)) {
      valid_dates <- DO_dates %>% filter(do_date == trip$Sample_Date)
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, do_date)$do_date else NA
    } else { NA }
    
    orp_cal_date <- if (grepl("00090", collected_storets)) {
      valid_dates <- ORP_dates %>% filter(orp_date <= trip$Sample_Date & orp_date >= (trip$Sample_Date - 60))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, orp_date)$orp_date else NA
    } else { NA }
    
    chl_cal_date <- if (grepl("32241", collected_storets) || grepl("32244", collected_storets)) {
      valid_dates <- Chlorophyll_dates %>% filter(chl_date <= trip$Sample_Date & chl_date >= (trip$Sample_Date - 60))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, chl_date)$chl_date else NA
    } else { NA }
    
    phyco_cal_date <- if (grepl("32242", collected_storets)) {
      valid_dates <- Phycocyanin_dates %>% filter(phyco_date <= trip$Sample_Date & phyco_date >= (trip$Sample_Date - 60))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, phyco_date)$phyco_date else NA
    } else { NA }
    
    turb_cal_date <- if (grepl("00076", collected_storets)) {
      valid_dates <- Turbidity_dates %>% filter(turb_date <= trip$Sample_Date & turb_date >= (trip$Sample_Date - 60))
      if (nrow(valid_dates) > 0) top_n(valid_dates, 1, turb_date)$turb_date else NA
    } else { NA }
    
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
  
  message("-> Full calibration sheet populated successfully.")
  
  # --- 5. Split the sheet into two separate dataframes ---
  message("-> Separating Lake Profile data from main calibration sheet...")
  profile_storets <- c("00299", "00301", "00299, 00301")
  
  lake_profiles_and_misc <- full_calibration_sheet %>%
    filter(Storet_Collected %in% profile_storets)
  
  calibration_sheet <- full_calibration_sheet %>%
    filter(!Storet_Collected %in% profile_storets)
  
  
  # --- Bundle all dataframes into a single list for return ---
  raw_cal_logs <- list(
    DO_Log = cal_DO,
    pH_Log = cal_pH,
    Specific_Conducitivity_Log = cal_SpCond,
    ORP_Log = cal_ORP,
    Turbidity_Log = cal_Turbidity,
    Chlorophyll_Log = cal_Chlorophyll,
    Phycocyanin_Log = cal_Phycocyanin
  )
  
  cleaned_cal_dates <- list(
    DO_dates = DO_dates,
    pH_dates = pH_dates,
    SpCond_dates = SpCond_dates,
    ORP_dates = ORP_dates,
    Turbidity_dates = Turbidity_dates,
    Chlorophyll_dates = Chlorophyll_dates,
    Phycocyanin_dates = Phycocyanin_dates
  )
  
  final_cal_list <- list(
    Raw_Calibration_Logs = raw_cal_logs,
    Cleaned_Calibration_Dates = cleaned_cal_dates,
    Identified_Sampling_Days = sampling_days,
    Final_Calibration_Sheet = calibration_sheet,
    Lake_Profiles_Sheet = lake_profiles_and_misc,
    Chlorophyll_Stnadards = Chlorophyll_Standards,
    Standards = Standards
  )
  
  return(final_cal_list)
  
}