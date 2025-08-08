# --- Load Required Packages ---
library(dplyr)


# --- Calculate Summary Statistics from Screened Chemical QC Data ---

calculate_summary_statistics <- function(screened_dataframe) {
  
  message("Calculating summary statistics for chemical data...")
  
  if (!is.data.frame(screened_dataframe) || nrow(screened_dataframe) == 0) {
    message("Input dataframe is empty or invalid. Skipping statistics calculation.")
    return(NULL)
  }
  
  # --- 1. Overall Pass/Fail Rates ---
  overall_pass_fail <- screened_dataframe %>%
    group_by(QC_Failed_Overall) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  # --- 2. Pass/Fail Rates by different groups ---
  pass_fail_by_location <- screened_dataframe %>%
    group_by(Loc_ID, QC_Failed_Overall) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    group_by(Loc_ID) %>%
    mutate(Percentage = (Count / sum(Count)) * 100) %>%
    ungroup()
  
  pass_fail_by_type <- screened_dataframe %>%
    group_by(QC_Type, QC_Failed_Overall) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    group_by(QC_Type) %>%
    mutate(Percentage = (Count / sum(Count)) * 100) %>%
    ungroup()
  
  pass_fail_by_parameter <- screened_dataframe %>%
    group_by(Parameter, QC_Failed_Overall) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    group_by(Parameter) %>%
    mutate(Percentage = (Count / sum(Count)) * 100) %>%
    ungroup()
  
  # --- 3. All Samples Table ---
  all_QC_samples_table <- screened_dataframe 
  message("-> Created table of all chemical QC samples.")
  
  # --- 4. Mean and Median of RPD ---
  rpd_stats <- screened_dataframe %>%
    filter(QC_Type %in% c("DDL", "DUP", "SPL")) %>%
    summarise(
      Mean_RPD = mean(RPD, na.rm = TRUE),
      Median_RPD = median(RPD, na.rm = TRUE)
    )
  
  # --- Compile all results into a single list and return ---
  summary_list <- list(
    Overall_Pass_Fail_Rates = overall_pass_fail,
    Pass_Fail_By_Location = pass_fail_by_location,
    Pass_Fail_By_QC_Type = pass_fail_by_type,
    Pass_Fail_By_Parameter = pass_fail_by_parameter,
    All_QC_Samples_Rows = all_QC_samples_table,
    RPD_Statistics = rpd_stats
  )
  
  message("Chemical summary statistics compiled successfully.")
  return(summary_list)
}


# --- Calculate Summary Statistics from Bray-Curtis Results ---

calculate_bio_summary_statistics <- function(bray_curtis_df) {
  
  message("\nCalculating summary statistics for biological data...")
  
  if (!is.data.frame(bray_curtis_df) || nrow(bray_curtis_df) == 0) {
    message("Bray-Curtis results dataframe is empty or invalid. Skipping bio-statistics calculation.")
    return(NULL)
  }
  
  # --- 1. Overall Pass/Fail Rates for Bray-Curtis ---
  bio_pass_fail <- bray_curtis_df %>%
    group_by(BC_Result) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  # --- 2. Mean and Median of the Bray-Curtis Index ---
  bc_index_stats <- bray_curtis_df %>%
    summarise(
      Mean_Bray_Curtis_Index = mean(Bray_Curtis_Index, na.rm = TRUE),
      Median_Bray_Curtis_Index = median(Bray_Curtis_Index, na.rm = TRUE)
    )
  
  # --- Compile all results into a single list and return ---
  bio_summary_list <- list(
    Bio_Overall_Pass_Fail = bio_pass_fail,
    BC_Index_Statistics = bc_index_stats
  )
  
  message("Biological summary statistics compiled successfully.")
  return(bio_summary_list)
}
