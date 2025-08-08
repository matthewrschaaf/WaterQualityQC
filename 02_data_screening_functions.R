# --- Load Required Packages ---
library(dplyr)
library(stringr)

# --- Screening Data Function Chemistry ---
screen_qc_data <- function(merged_dataframe) {
  
  message("Screening data against performance criteria...")
  
  if (!is.data.frame(merged_dataframe) || nrow(merged_dataframe) == 0) {
    message("Input is not a valid dataframe or is empty. Skipping screening.")
    return(NULL)
  }
  
  screened_df <- merged_dataframe %>%
    mutate(Parameter = str_trim(Parameter)) %>%
    mutate(
      # RPD calculation (independent of LOQ)
      RPD = if_else(
        (abs(QC_Value) + abs(Field_Value)) == 0, 
        0, 
        abs(abs(QC_Value) - abs(Field_Value)) / ((abs(QC_Value) + abs(Field_Value)) / 2) * 100
      ),
      
      # --- Step 1: Calculate Criteria 1 for all samples ---
      QC_Failed_Criteria1 = case_when(
        QC_Type %in% c("FBLK", "RNS") & is.na(LOQ) ~ NA_character_,
        QC_Type %in% c("FBLK", "RNS") & QC_Value >= LOQ ~ "Fail",
        QC_Type %in% c("DDL", "DUP", "SPL") & RPD > 20 ~ "Fail",
        Parameter == "ORP" & RPD > 20 ~ "Fail",
        TRUE ~ "Pass"
      ),
      
      # --- Step 2: Calculate Criteria 2 ONLY if necessary ---
      # only necessary for dups/splits that failed Criteria 1 and have a valid LOQ.
      QC_Failed_Criteria2 = case_when(
        QC_Failed_Criteria1 == "Pass" ~ NA_character_, # If C1 passes, C2 automatically passes.
        is.na(LOQ) ~ "Can Not Be Calculated due to NA LOQ", # If LOQ is NA, C2 cannot be calculated.
        !QC_Type %in% c("DDL", "DUP", "SPL") ~ QC_Failed_Criteria1, # For non-dups, C2 is the same as C1.
        
        # For dups that failed C1 and have a valid LOQ, apply the specific rules:
        QC_Value == 0 & Field_Value == 0 ~ "Pass",
        abs(QC_Value) >= 5 * LOQ & abs(Field_Value) >= 5 * LOQ & RPD > 20 ~ "Fail",
        abs(QC_Value) < 5 * LOQ | abs(Field_Value) < 5 * LOQ & abs(abs(QC_Value) - abs(Field_Value)) > 2 * LOQ ~ "Fail",
        TRUE ~ "Pass"
      )
    ) %>%
    # --- Step 3: Determine Final Overall Status ---
    mutate(
      QC_Failed_Overall = case_when(
        # If C1 passes, the overall is Pass.
        QC_Failed_Criteria1 == "Pass" ~ "Pass",
        
        # If C2 could not be calculated (due to NA LOQ), the overall defaults to C1's result.
        QC_Failed_Criteria2 == "Can Not Be Calculated due to NA LOQ" ~ QC_Failed_Criteria1,
        
        # Otherwise, the overall status is the same as the C2 result.
        TRUE ~ QC_Failed_Criteria2
      )
    )

  # --- Special Flagging for Total Nitrogen ---
  
  # Define the component parameters
  tn_components <- c("Kjeldahl, Tot", "NO2+NO3, Tot")
  
  # Group by each original sample and apply the new TN logic
  screened_df <- screened_df %>%
    group_by(Loc_ID, QC_Sample) %>%
    mutate(
      # Check if a "Kjeldahl, Tot" sample exists and passed in this group
      k_pass = any(Parameter == "Kjeldahl, Tot" & QC_Failed_Overall == "Pass"),
      # Check if a "NO2+NO3, Tot" sample exists and passed in this group
      n_pass = any(Parameter == "NO2+NO3, Tot" & QC_Failed_Overall == "Pass"),
      
      # Determine the final status for "Nitrogen, Tot"
      # It only passes if BOTH components passed.
      tn_status = if_else(k_pass & n_pass, "Pass", "Fail")
    ) %>%
    # Apply the new rules specifically to the "Nitrogen, Tot" rows
    mutate(
      QC_Failed_Criteria1 = if_else(Parameter == "Nitrogen, Tot", NA_character_, QC_Failed_Criteria1),
      QC_Failed_Criteria2 = if_else(Parameter == "Nitrogen, Tot", NA_character_, QC_Failed_Criteria2),
      QC_Failed_Overall = if_else(Parameter == "Nitrogen, Tot", tn_status, QC_Failed_Overall)
    ) %>%
    # Remove the temporary helper columns and ungroup
    ungroup() %>%
    select(-k_pass, -n_pass, -tn_status)

message("-> Screening complete. Initial screened dataframe has ", ncol(screened_df), " columns and ", nrow(screened_df), " rows.")
  
  # --- Create a separate dataframe for sediment samples AND remove them from the main df ---
  message("\nSeparating sediment (S) samples from water column samples...")
  
  # 1. Create the new dataframe containing only sediment samples
  screened_sedimentdf <- screened_df %>%
    filter(Sample_Type == "S")
  
  # 2. Update the original dataframe to remove the sediment rows
  screened_df <- screened_df %>%
    filter(Sample_Type != "S")
  
  message("-> Created sediment dataframe with ", nrow(screened_sedimentdf), " rows.")
  message("-> Removed sediment samples from screened dataframe which now has ", nrow(screened_df), " rows.")
  
  # Split the screened dataframe by QC_Type
  split_screened_data <- split(screened_df, screened_df$QC_Type)
  
  message("\n-> Screened dataframe split into ", length(split_screened_data), " separate dataframes by QC_Type.", names(split_screened_data))
  
  if ("DUP" %in% names(split_screened_data)) {

    original_dup_df <- split_screened_data$DUP
    
    # Create the new dataframe for physical duplicates (Lab_ID is "Field")
    phys_dup_df <- original_dup_df %>%
      filter(Field_Lab_ID == "FIELD")
    
    # Create the new dataframe for chemical duplicates (Lab_ID is not "Field")
    chem_dup_df <- original_dup_df %>%
      filter(Field_Lab_ID != "FIELD")
    
    # Add the two new dataframes to the list
    split_screened_data$Phys_DUP <- phys_dup_df
    split_screened_data$Chem_DUP <- chem_dup_df
    
    # Remove the original, combined DUP dataframe from the list
    split_screened_data$DUP <- NULL
    
    message("\n-> 'DUP' dataframe has been split into 'Phys_DUP' and 'Chem_DUP'.")
  }
  
  # Create a list to hold all results
  results_list <- list(
    All_Screened_Data = screened_df,
    Split_by_QC_Type = split_screened_data,
    Screened_Sediment_Data = screened_sedimentdf
  )
  
  # Return the new list containing all objects
  return(results_list)
}


# --- Calculate Bray-Curtis Dissimilarity for Biological Samples ---
calculate_bray_curtis <- function(final_bio_df) {
  
  message("\nCalculating Bray-Curtis Dissimilarity Index for each sample pair...")
  
  if (!is.data.frame(final_bio_df) || nrow(final_bio_df) == 0) {
    message("Input dataframe is empty or invalid. Skipping calculation.")
    return(NULL)
  }
  
  bray_curtis_results <- final_bio_df %>%
    # Step 1: Filter out any rows where the key identifiers are missing
    filter(!is.na(Loc_ID) & !is.na(OG_Sample)) %>%
    
    # Step 2: Group the data by each unique sample pair
    group_by(Loc_ID, OG_Sample, QC_Type, Sample_Type) %>%
    
    # Step 3: Summarise the data for each group to get A, B, and C
    summarise(
      # A = sum of all abundances in the original sample
      Sum_A = sum(Density_OG, na.rm = TRUE),
      
      # B = sum of all abundances in the QC sample
      Sum_B = sum(Density_QC, na.rm = TRUE),
      
      # C = sum of the minimum abundance for each shared species
      Sum_C = sum(pmin(Density_OG, Density_QC), na.rm = TRUE),
      
      .groups = 'drop'
    ) %>%
    
    mutate(
      Bray_Curtis_Index = if_else(
        (Sum_A + Sum_B) == 0,
        0,
        1 - (2 * Sum_C) / (Sum_A + Sum_B)
      ),
      
      BC_Result = if_else(Bray_Curtis_Index <= 0.25, "Pass", "Fail")
    )
  
  message("-> Bray-Curtis calculation complete.")
  
  return(bray_curtis_results)
}
