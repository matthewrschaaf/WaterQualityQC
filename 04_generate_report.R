# --- Load Required Packages ---
library(openxlsx)
library(dplyr)


# --- Generate Excel Report Function ---
generate_excel_report <- function(user_input, 
                                  chem_stats, 
                                  bio_stats, 
                                  bray_curtis, 
                                  bio_data, 
                                  sediment_data, 
                                  chem_split_data, 
                                  chem_data, 
                                  locations, 
                                  output_filepath) {
  
  # --- 1. Setup Workbook and Styles ---
  wb <- createWorkbook()
  bold_style <- createStyle(textDecoration = "Bold")
  wrap_style <- createStyle(wrapText = TRUE)
  
  # --- 2. Build Sheet 2: Locations Sampled ---
  addWorksheet(wb, "Locations Sampled")
  writeData(wb, "Locations Sampled", locations, startCol = 1, startRow = 1)
  addStyle(wb, "Locations Sampled", style = bold_style, rows = 1, cols = 1:ncol(locations))
  addFilter(wb, "Locations Sampled", rows = 1, cols = 1:ncol(locations))
  setColWidths(wb, "Locations Sampled", cols = 1:ncol(locations), widths = "auto")
  
  
  # --- 3. Placeholder for Sheet 3
  addWorksheet(wb, "Samples Collected vs Projected")
  writeData(wb, "Samples Collected vs Projected", "Placeholder: See information about tracking sampling (PSP?).")
  
  
  # --- 4. Build Sheet 4-7: All Chem QC Results Tables ---
  
  # Define the detailed criteria text once to reuse it
  criteria_text <- paste(
    "The data was screened using a two-tiered system. A sample receives an 'Overall Fail' status only if it fails both Criteria 1 and 2.",
    "",
    "Criteria 1:",
    "• For blanks: Considered a 'fail' if the measured 'QC_Value' is >= LOQ.",
    "• For duplicates/splits: Considered a 'fail' if RPD between 'QC_Value' and 'Field_Value' is >20%.",
    "• Otherwise: Considered a 'pass'.",
    "",
    "Criteria 2:",
    "• For blanks/other types: Inherits the result from Criteria 1.",
    "• For duplicates/splits:",
    "  – High concentration samples (both values >= 5x LOQ): Fails if RPD >20%.",
    "  – Low concentration samples (either value < 5x LOQ): Fails if absolute difference > 2x LOQ.",
    "  – If both values are 0, it is a 'pass'.",
    "• Otherwise: Considered a 'pass'.",
    "",
    "Overall QC Status:",
    "• A sample is marked 'fail' only if both Criteria 1 and Criteria 2 are 'fail'.",
    "• If either (or both) are 'pass', the overall status is 'pass'.",
    "",
    "Nitrogen, Total (Tot):",
    "• A sample is marked 'fail' if either 'Kjeldahl, Tot' or 'NO2+NO3, Tot' is 'fail' for the same associated sample",
    sep = "\n"
  )
  
  
  # --- Helper function to add stats and criteria text to a sheet ---
  add_summary_stats_to_sheet <- function(workbook, sheet_name, qc_data, rpd_stats = FALSE) {
    
    # A small helper to ensure both "Pass" and "Fail" columns exist after pivoting
    ensure_pass_fail_columns <- function(df) {
      if (!"Pass" %in% names(df)) {
        df$Pass <- 0
      }
      if (!"Fail" %in% names(df)) {
        df$Fail <- 0
      }
      return(df)
    }
    
    # --- Calculate all required statistics for the specific QC type ---
    
    # By Location: Calculate fail ratio and sort
    stats_by_loc <- qc_data %>%
      group_by(Loc_ID, QC_Failed_Overall) %>%
      summarise(Count = n(), .groups = 'drop') %>%
      tidyr::pivot_wider(names_from = QC_Failed_Overall, values_from = Count, values_fill = 0) %>%
      ensure_pass_fail_columns() %>%
      mutate(Total = Pass + Fail, Fail_Ratio = Fail / Total) %>%
      arrange(desc(Fail_Ratio))
    
    # By Parameter: Calculate fail ratio and sort
    stats_by_param <- qc_data %>%
      group_by(Parameter, QC_Failed_Overall) %>%
      summarise(Count = n(), .groups = 'drop') %>%
      tidyr::pivot_wider(names_from = QC_Failed_Overall, values_from = Count, values_fill = 0) %>%
      ensure_pass_fail_columns() %>%
      mutate(Total = Pass + Fail, Fail_Ratio = Fail / Total) %>%
      arrange(desc(Fail_Ratio))
    
    # Calculate Overall Pass/Fail Rate for this sheet
    overall_stats_sheet <- qc_data %>%
      group_by(QC_Failed_Overall) %>%
      summarise(Count = n(), .groups = 'drop') %>%
      mutate(Percentage = (Count / sum(Count)) * 100)
    
    # Calculate Pass/Fail Rate by Year for this sheet
    stats_by_year <- qc_data %>%
      mutate(Year = substr(QC_Sample, 1, 4)) %>%
      group_by(Year, QC_Failed_Overall) %>%
      summarise(Count = n(), .groups = 'drop') %>%
      tidyr::pivot_wider(names_from = QC_Failed_Overall, values_from = Count, values_fill = 0) %>%
      ensure_pass_fail_columns() %>%
      mutate(Total = Pass + Fail, Fail_Ratio = Fail / Total) %>%
      arrange(Year)
    
    # --- Write all tables to the sheet ---
    
    # Define where to start writing the first set of summary tables
    start_col_set1 <- ncol(qc_data) + 2
    
    # Write stats by location
    writeData(workbook, sheet_name, "Pass/Fail by Location (Highest Fail Rate First)", startCol = start_col_set1, startRow = 1)
    addStyle(workbook, sheet_name, style = bold_style, rows = 1, cols = start_col_set1)
    writeData(workbook, sheet_name, stats_by_loc, startCol = start_col_set1, startRow = 2)
    
    # Write stats by parameter, leaving space below the location stats
    start_row_param <- nrow(stats_by_loc) + 4
    writeData(workbook, sheet_name, "Pass/Fail by Parameter (Highest Fail Rate First)", startCol = start_col_set1, startRow = start_row_param)
    addStyle(workbook, sheet_name, style = bold_style, rows = start_row_param, cols = start_col_set1)
    writeData(workbook, sheet_name, stats_by_param, startCol = start_col_set1, startRow = start_row_param + 1)
    
    # Define where to start writing the second set of summary tables
    start_col_set2 <- start_col_set1 + ncol(stats_by_loc) + 2
    
    # Write the overall pass/fail summary for the sheet
    writeData(workbook, sheet_name, "Overall Pass/Fail Rate (This Sheet)", startCol = start_col_set2, startRow = 1)
    addStyle(workbook, sheet_name, style = bold_style, rows = 1, cols = start_col_set2)
    writeData(workbook, sheet_name, overall_stats_sheet, startCol = start_col_set2, startRow = 2)
    
    # Write the pass/fail summary by year
    start_row_year <- nrow(overall_stats_sheet) + 4
    writeData(workbook, sheet_name, "Pass/Fail Rate by Year", startCol = start_col_set2, startRow = start_row_year)
    addStyle(workbook, sheet_name, style = bold_style, rows = start_row_year, cols = start_col_set2)
    writeData(workbook, sheet_name, stats_by_year, startCol = start_col_set2, startRow = start_row_year + 1)
    
    # Conditionally write RPD stats (if applicable)
    if (rpd_stats) {
      rpd_summary <- qc_data %>% summarise(Mean_RPD = mean(RPD, na.rm = TRUE), Median_RPD = median(RPD, na.rm = TRUE))
      start_row_rpd <- start_row_year + nrow(stats_by_year) + 3
      writeData(workbook, sheet_name, "RPD Statistics", startCol = start_col_set2, startRow = start_row_rpd)
      addStyle(workbook, sheet_name, style = bold_style, rows = start_row_rpd, cols = start_col_set2)
      writeData(workbook, sheet_name, rpd_summary, startCol = start_col_set2, startRow = start_row_rpd + 1)
    }
    
    # --- Add the criteria text block to the sheet ---
    start_col_criteria <- start_col_set2 + ncol(overall_stats_sheet) + 3
    writeData(workbook, sheet_name, "Enlarge Comment Box to View QC Criteria", startCol = start_col_criteria, startRow = 1)
    addStyle(workbook, sheet_name, style = bold_style, rows = 1, cols = start_col_criteria)
    criteria_comment <- createComment(comment = criteria_text)
    writeComment(workbook, sheet_name, col = start_col_criteria, row = 1, comment = criteria_comment)
  }
  
  
  # Adding Failure Formatting to Color Code Fails
  add_fail_formatting <- function(workbook, sheet_name, data_df) {
    
    fail_col_index <- which(names(data_df) == "QC_Failed_Overall")
    
    if (length(fail_col_index) > 0) {
      
      fail_fill_style <- createStyle(bgFill = "#FFC7CE")
      
      fail_font_style <- createStyle(fontColour = "#9C0006")
      
      conditionalFormatting(
        workbook,
        sheet = sheet_name,
        cols = fail_col_index,
        rows = 2:(nrow(data_df) + 1),
        rule = '=="Fail"',
        style = fail_fill_style
      )
      
      conditionalFormatting(
        workbook,
        sheet = sheet_name,
        cols = fail_col_index,
        rows = 2:(nrow(data_df) + 1),
        rule = '=="Fail"',
        style = fail_font_style
      )
    }
  }
  
  
  # --- Build Chem QC Sheets ---
  
  # --- Create Sediment QC Sheet ---
  if (!is.null(sediment_data) && nrow(sediment_data) > 0) {
    
    # Add the blank "Notes" and "Don't Use" columns
    sediment_data <- sediment_data %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "Sediment QC Results")
    writeData(wb, "Sediment QC Results", sediment_data, headerStyle = bold_style)
    addFilter(wb, "Sediment QC Results", rows = 1, cols = 1:ncol(sediment_data))
    add_summary_stats_to_sheet(wb, "Sediment QC Results", sediment_data, rpd_stats = TRUE)
    setColWidths(wb, "Sediment QC Results", cols = 1:(ncol(sediment_data) + 12), widths = "auto")
  }
  
  # --- Create DUP (Duplicate) Sheets ---
  if ("Phys_DUP" %in% names(chem_split_data)) {
    phys_DUP_data <- chem_split_data$Phys_DUP %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "DUP Phys QC Results")
    writeData(wb, "DUP Phys QC Results", phys_DUP_data, headerStyle = bold_style)
    addFilter(wb, "DUP Phys QC Results", rows = 1, cols = 1:ncol(phys_DUP_data))
    add_summary_stats_to_sheet(wb, "DUP Phys QC Results", phys_DUP_data, rpd_stats = TRUE)
    add_fail_formatting(wb, "DUP Phys QC Results", phys_DUP_data)
    setColWidths(wb, "DUP Phys QC Results", cols = 1:(ncol(phys_DUP_data) + 12), widths = "auto")
  }
  
  if ("Chem_DUP" %in% names(chem_split_data)) {
    chem_DUP_data <- chem_split_data$Chem_DUP %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "DUP Chem QC Results")
    writeData(wb, "DUP Chem QC Results", chem_DUP_data, headerStyle = bold_style)
    addFilter(wb, "DUP Chem QC Results", rows = 1, cols = 1:ncol(chem_DUP_data))
    add_summary_stats_to_sheet(wb, "DUP Chem QC Results", chem_DUP_data, rpd_stats = TRUE)
    add_fail_formatting(wb, "DUP Chem QC Results", chem_DUP_data)
    setColWidths(wb, "DUP Chem QC Results", cols = 1:(ncol(chem_DUP_data) + 12), widths = "auto")
  }
  
  # --- Create SPL (Split) Sheet ---
  if ("SPL" %in% names(chem_split_data)) {
    spl_data <- chem_split_data$SPL %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "SPL Chem QC Results")
    writeData(wb, "SPL Chem QC Results", spl_data, headerStyle = bold_style)
    addFilter(wb, "SPL Chem QC Results", rows = 1, cols = 1:ncol(spl_data))
    add_summary_stats_to_sheet(wb, "SPL Chem QC Results", spl_data, rpd_stats = TRUE)
    add_fail_formatting(wb, "SPL Chem QC Results", spl_data)
    setColWidths(wb, "SPL Chem QC Results", cols = 1:(ncol(spl_data) + 12), widths = "auto")
  }
  
  # --- Create RNS (Rinse) Sheet ---
  if ("RNS" %in% names(chem_split_data)) {
    rns_data <- chem_split_data$RNS %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "RNS Chem QC Results")
    writeData(wb, "RNS Chem QC Results", rns_data, headerStyle = bold_style)
    addFilter(wb, "RNS Chem QC Results", rows = 1, cols = 1:ncol(rns_data))
    add_summary_stats_to_sheet(wb, "RNS Chem QC Results", rns_data, rpd_stats = FALSE)
    add_fail_formatting(wb, "RNS Chem QC Results", rns_data)
    setColWidths(wb, "RNS Chem QC Results", cols = 1:(ncol(rns_data) + 12), widths = "auto")
  }
  
  # --- Create FBLK Sheet (if necessary)
  if ("FBLK" %in% names(chem_split_data)) {
    fblk_data <- chem_split_data$FBLK %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "FBLK Chem QC Results")
    writeData(wb, "FBLK Chem QC Results", fblk_data, headerStyle = bold_style)
    addFilter(wb, "FBLK Chem QC Results", rows = 1, cols = 1:ncol(fblk_data))
    add_summary_stats_to_sheet(wb, "FBLK Chem QC Results", fblk_data, rpd_stats = FALSE)
    add_fail_formatting(wb, "FBLK Chem QC Results", fblk_data)
    setColWidths(wb, "FBLK Chem QC Results", cols = 1:(ncol(fblk_data) + 12), widths = "auto")
  }
  
  # --- Create TBLK Sheet (if necessary) ---
  if ("TBLK" %in% names(chem_split_data)) {
    tblk_data <- chem_split_data$TBLK %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "TBLK Chem QC Results")
    writeData(wb, "TBLK Chem QC Results", tblk_data, headerStyle = bold_style)
    addFilter(wb, "TBLK Chem QC Results", rows = 1, cols = 1:ncol(tblk_data))
    add_summary_stats_to_sheet(wb, "TBLK Chem QC Results", tblk_data, rpd_stats = FALSE)
    add_fail_formatting(wb, "TBLK Chem QC Results", tblk_data)
    setColWidths(wb, "TBLK Chem QC Results", cols = 1:(ncol(tblk_data) + 12), widths = "auto")
  }
  
  # --- Create DDL Sheet (if necessary) ---
  if ("DDL" %in% names(chem_split_data)) {
    ddl_data <- chem_split_data$DDL %>%
      mutate(Notes = NA, `Don't Use` = NA)
    
    addWorksheet(wb, "DDL Chem QC Results")
    writeData(wb, "DDL Chem QC Results", ddl_data, headerStyle = bold_style)
    addFilter(wb, "DDL Chem QC Results", rows = 1, cols = 1:ncol(ddl_data))
    add_summary_stats_to_sheet(wb, "DDL Chem QC Results", ddl_data, rpd_stats = FALSE)
    add_fail_formatting(wb, "DDL Chem QC Results", ddl_data)
    setColWidths(wb, "DDL Chem QC Results", cols = 1:(ncol(ddl_data) + 12), widths = "auto")
  }
  
  
  # --- 5. Biological QC Sheets ---
  
  # --- 5a. All Bio QC Results Sheet ---
  addWorksheet(wb, "All Bio QC Results")
  writeData(wb, "All Bio QC Results", bio_data)
  addStyle(wb, "All Bio QC Results", style = bold_style, rows = 1, cols = 1:ncol(bio_data))
  addFilter(wb, "All Bio QC Results", rows = 1, cols = 1:ncol(bio_data))
  setColWidths(wb, "All Bio QC Results", cols = 1:ncol(bio_data), widths = "auto")
  
  # --- 5b. Biological Summary Sheet ---
  addWorksheet(wb, "Bio Summary Statistics")
  
  bc_criteria_text <- paste(
    "Precision and comparability for phytoplankton and zooplankton samples were assessed using the Bray-Curtis Dissimilarity index (B-C), with a value of <= 0.25 considered a successful pass.",
    "",
    "The index quantifies precision by comparing species composition using the formula: BC = 1 - (2*C)/(A+B), where:",
    "  • A = the sum of abundances for all species in the first subsample.",
    "  • B = the sum of abundances for all species in the second subsample.",
    "  • C = the sum of the lesser abundance for each species shared between the two.",
    "",
    "The resulting value ranges from 0 to 1. A score of 0 signifies the samples are identical in composition, while a score of 1 signifies they share no species.",
    sep = "\n"
  )

  # Write the Index statistics table, leaving some space
  writeData(wb, "Bio Summary Statistics", "Bray-Curtis Index Statistics", startCol = 1, startRow = 7)
  addStyle(wb, "Bio Summary Statistics", style = bold_style, rows = 7, cols = 1)
  writeData(wb, "Bio Summary Statistics", bio_stats$BC_Index_Statistics, startCol = 1, startRow = 8)
  setColWidths(wb, "Bio Summary Statistics", cols = 1:3, widths = "auto")
  
  # Write the Bray Curtis Results table to the right of the other stats
  writeData(wb, "Bio Summary Statistics", "Bray-Curtis Results", startCol = 6, startRow = 1)
  addStyle(wb, "Bio Summary Statistics", style = bold_style, rows = 1, cols = 6)
  writeData(wb, "Bio Summary Statistics", bray_curtis, startCol = 6, startRow = 2)
  addFilter(wb, "Bio Summary Statistics", rows = 2, cols = 6:(6 + ncol(bray_curtis) - 1))
  setColWidths(wb, "Bio Summary Statistics", cols = 6:(6 + ncol(bray_curtis) - 1), widths = "auto")
  
  start_col_criteria <- 18
  writeData(wb, "Bio Summary Statistics", "Enlarge Comment Box to View QC Criteria", startCol = start_col_criteria, startRow = 1)
  addStyle(wb, "Bio Summary Statistics", style = bold_style, rows = 1, cols = start_col_criteria)
  criteria_comment <- createComment(comment = bc_criteria_text)
  writeComment(wb, "Bio Summary Statistics", col = start_col_criteria, row = 1, comment = criteria_comment)
  
  fail_style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
  
  conditionalFormatting(
    wb,
    sheet = "Bio Summary Statistics",
    cols = 14,
    rows = 1:1048576,
    rule = '=="Fail"',
    style = fail_style
  )
  
  
  # --- 6. Placeholder for Sheet 6 ---
  addWorksheet(wb, "SOP Deviations")
  writeData(wb, "SOP Deviations", "Placeholder: Use this sheet to document any deviations from SOPs.")
  
  
  # --- 7. Save the Workbook ---
  saveWorkbook(wb, file = output_filepath, overwrite = TRUE)
  message(paste("Excel report generated successfully at:", output_filepath))
}
