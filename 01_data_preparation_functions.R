# --- Load Required Packages ---
library(DBI)
library(odbc)
library(dplyr)
library(tidyr)


fetch_and_prep_tables <- function(db_path, start_date, end_date) {
  
  connection_string <- "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=O:/ED/Private/Water Quality/DASLER/Louisville_DASLER 81.mdb"
  
  con <- NULL
  
  tryCatch({
    message("Attempting to connect to the database...")
    con <- dbConnect(odbc::odbc(), .connection_string = connection_string)
    message("Connection successful.")
    
    # --- Chemical Data ---
    
    # --- 1. Fetch and Process [QC Results] ---
    message("\nFetching and processing 'QC Results'...")
    QC_Results <- dbGetQuery(con, "SELECT LOC_ID, QC_SAMPLE, LAB_ID, STORET_NUM, UNITS,
                                     VALUE, TEXT_VALUE, PQUANT, QC_TYPE, LAB_SNUM
                                  FROM [QC Results]") %>%
      dplyr::rename(
        Loc_ID = LOC_ID,
        QC_Sample = QC_SAMPLE,
        QC_Lab_ID = LAB_ID,
        Storet_Num = STORET_NUM,
        QC_Value = VALUE,
        Text_QC_Value = TEXT_VALUE,
        QC_Type = QC_TYPE,
        UnitsQC = UNITS,
        QC_LOQ = PQUANT,
        QC_Lab_Sample_Num = LAB_SNUM
      ) %>%
      mutate(QC_Sample_Date = as.Date(substr(QC_Sample, 1, 8), format = "%Y%m%d")) %>%
      filter(QC_Sample_Date >= start_date & QC_Sample_Date <= end_date) %>%
      select(-QC_Sample_Date)
    
    message("-> Processed ", nrow(QC_Results), " records from [QC Results].")
    
    
    # --- 2. Fetch and Process [QC Samples] ---
    message("Fetching and processing 'QC Samples'...")
    QC_Samples <- dbGetQuery(con, "SELECT LOC_ID, QC_SAMPLE, QC_TYPE, ASSOC_SAMP, SAMPLE_TYPE
                                     FROM [QC Samples]") %>%
      dplyr::rename(
        Loc_ID = LOC_ID,
        QC_Sample = QC_SAMPLE,
        QC_Type = QC_TYPE,
        Assoc_Samp = ASSOC_SAMP,
        Sample_Type = SAMPLE_TYPE
      )
    message("-> Fetched ", nrow(QC_Samples), " records from [QC Samples].")
    
    
    # --- 3. Fetch and Process [Results] ---
    message("Fetching and processing 'Results'...")
    Results <- dbGetQuery(con, "SELECT LOC_ID, SAMPLE_NUM, VALUE, TEXT_VALUE, STORET_NUM, UNITS, LAB_SNUM, LAB_ID, PQUANT
                                   FROM [Results]") %>%
      dplyr::rename(
        Loc_ID = LOC_ID,
        Assoc_Samp = SAMPLE_NUM,
        Field_Value = VALUE,
        Text_Field_Value = TEXT_VALUE,
        Storet_Num = STORET_NUM,
        Field_LOQ = PQUANT,
        UnitsField = UNITS,
        Field_Lab_ID = LAB_ID,
        Field_Lab_Sample_Num = LAB_SNUM,
      ) %>% 
      mutate(Assoc_Samp_Date = as.Date(substr(Assoc_Samp, 1, 8), format = "%Y%m%d")) %>%
      filter(Assoc_Samp_Date >= start_date & Assoc_Samp_Date <= end_date) %>% 
      select(-Assoc_Samp_Date)
    
    message("-> Fetched ", nrow(Results), " records from [Results].")
    
    message("Deduplicating the Results_Raw dataframe...")
    
    # --- Step 1: Isolate the rows that are NOT true duplicates ---
    discrepancies_and_uniques_res <- Results %>%
      group_by(Loc_ID, Assoc_Samp, Storet_Num) %>%
      filter(n_distinct(Text_Field_Value) > 1 | n() == 1) %>%
      ungroup()
    
    # --- Step 2: Isolate and deduplicate the TRUE duplicates ---
    true_duplicates_res <- Results %>%
      group_by(Loc_ID, Assoc_Samp, Storet_Num) %>%
      filter(n_distinct(Text_Field_Value) == 1 & n() > 1) %>%
      slice(1) %>%
      ungroup()
    
    # --- Step 3: Combine the two sets back into the final dataframe ---
    Results <- bind_rows(
      discrepancies_and_uniques_res, 
      true_duplicates_res
    )
    
    message("-> Deduplication of Results_Raw complete.")
    
    
    # --- 4. Fetch and Process [Analytes] ---
    message("Fetching and processing 'Analytes'...")
    Analytes <- dbGetQuery(con, "SELECT STORET_NUM, ANL_SHORT
                                     FROM [Analytes]") %>%
      dplyr::rename(
        Storet_Num = STORET_NUM,
        Parameter = ANL_SHORT
      )
    message("-> Fetched ", nrow(Analytes), " records from [Analytes].")
    
    
    # --- Biological Data ---
    
    # --- 5. Fetch and Process [QC Biosamples]
    message("\nFetching and processing 'QC Biosamples'...")
    QC_Biosamples <- dbGetQuery(con, "SELECT Loc_ID, QC_Sample, Sample_Type, QC_Type, Assoc_Samp FROM [QC Biosamples]") %>% 
      mutate(Sample_Date = as.Date(substr(QC_Sample, 1, 8), format = "%Y%m%d")) %>%
      filter(Sample_Date >= start_date & Sample_Date <= end_date) %>% 
      filter(Sample_Type %in% c("PQN", "PQL", "ZQL", "ZQN")) %>% 
      select(-Sample_Date)
    message("-> Fetched and processed ", nrow(QC_Biosamples), " records from [QC Biosamples].")
    
    # --- 6. Fetch and Process [QC Bioresults] ---
    message("Fetching and processing 'QC Bioresults'...")
    QC_samp_list <- unique(QC_Biosamples$QC_Sample)
    QC_samp_list_for_sql <- paste0("'", paste(QC_samp_list, collapse = "','"), "'")
    qc_bioresults_query <- paste0(
      "SELECT LOC_ID, QC_SAMPLE, LAB_ID, SAMPLE_TYPE, TAXACODE, COUNT, BIOVOLUME, DENSITY, QC_TYPE, DENSITY_UNITS, BV_UNITS, ANAL_METHOD FROM [QC Bioresults] ",
      "WHERE QC_SAMPLE IN (", QC_samp_list_for_sql, ")"
    )
    
    QC_Bioresults <- dbGetQuery(con, qc_bioresults_query) %>%
      dplyr::rename(
        Loc_ID = LOC_ID,
        QC_Sample = QC_SAMPLE,
        Sample_Type = SAMPLE_TYPE,
        QC_Lab_ID = LAB_ID,
        Taxacode = TAXACODE,
        Count_QC = COUNT,
        Biovolume_QC = BIOVOLUME,
        Density_QC = DENSITY,
        QC_Type = QC_TYPE,
        Density_Units_QC = DENSITY_UNITS,
        BV_Units_QC = BV_UNITS,
        Analysis_Method = ANAL_METHOD
      ) %>%
      mutate(Sample_Date = as.Date(substr(QC_Sample, 1, 8), format = "%Y%m%d")) %>%
      filter(Sample_Date >= start_date & Sample_Date <= end_date) %>% 
      filter(Sample_Type %in% c("PQN", "PQL", "ZQL", "ZQN"))
    
    message("-> Fetched and processed ", nrow(QC_Bioresults), " records from [QC Bioresults].")
    
    # --- 7. Fetch and Process corresponding [BioResults] ---
    message("Fetching and processing 'BioResults'...")
    assoc_samp_list <- unique(QC_Biosamples$Assoc_Samp)
    samp_list_for_sql <- paste0("'", paste(assoc_samp_list, collapse = "','"), "'")
    bioresults_query <- paste0(
      "SELECT sample_num, loc_id, taxacode, lab_id, sample_type, count, biovolume, density, anal_method, count_units, density_units, bv_units FROM [BioResults] ",
      "WHERE sample_num IN (", samp_list_for_sql, ")"
    )
    
    # Execute the targeted query to fetch only the relevant rows
    BioResults <- dbGetQuery(con, bioresults_query) %>%
      dplyr::rename(
        Loc_ID = loc_id,
        OG_Sample = sample_num,
        Sample_Type = sample_type,
        Field_Lab_ID = lab_id,
        Taxacode = taxacode,
        Count_OG = count,
        Biovolume_OG = biovolume,
        Density_OG = density,
        Density_Units_OG = density_units,
        BV_Units_OG = bv_units,
        Analysis_Method = anal_method
      ) %>% 
      mutate(Sample_Date = as.Date(substr(OG_Sample, 1, 8), format = "%Y%m%d"))
    
    message("-> Fetched ", nrow(BioResults), " records from [BioResults].")
    
    # --- 8. Fetch and Process [Sample Type Codes]
    message("Fetching and processing 'Sample Type Codes'...")
    Sample_Type_Codes <- dbGetQuery(con, "SELECT SAMPLE_TYPE, TYPE_NAME FROM [Sample Type Codes]") %>% 
      dplyr::rename(
        Sample_Type = SAMPLE_TYPE,
        Type_Name = TYPE_NAME)
    
    message("-> Fetched ", nrow(Sample_Type_Codes), " records from [Sample Type Codes].")
    
    # --- 9. Fetch and Process [Taxa]
    message("Fetching and processing 'Taxa'...")
    Taxa <- dbGetQuery(con, "SELECT taxacode, sci_name, com_name, itis_tsn FROM [Taxa]") %>% 
      dplyr::rename(
        Taxacode = taxacode,
        Sci_Name = sci_name,
        Common_Name = com_name,
        ITIS_TSN = itis_tsn
      )
    
    message("-> Fetched ", nrow(Taxa), " records from [Taxa].")
    
    # --- 10. Fetch and Process [Locations] ---
    message("Fetching and processing 'Locations'...")
    locations <- dbGetQuery(con, "SELECT LOC_ID, LAT_DEG, LAT_MIN, LAT_SEC, LONG_DEG, LONG_MIN, LONG_SEC, LOC_DESC FROM [Locations]") %>%
      mutate(
        # Convert DMS to Decimal Degrees using the formula: DD = D + (M/60) + (S/3600)
        Latitude_DD = LAT_DEG + (LAT_MIN / 60) + (LAT_SEC / 3600),
        
        # For longitude in the Western Hemisphere, the final value must be negative.
        Longitude_DD = (LONG_DEG + (LONG_MIN / 60) + (LONG_SEC / 3600)) * -1
      ) %>% 
      
      dplyr::rename(
        Loc_ID = LOC_ID,
        Lat_Deg = LAT_DEG,
        Lat_Min = LAT_MIN,
        Lat_Sec = LAT_SEC,
        Long_Deg = LONG_DEG,
        Long_Min = LONG_MIN,
        Long_Sec = LONG_SEC,
        Loc_Desc = LOC_DESC
      )
    
    message("-> Fetched and processed ", nrow(locations), " records from [Locations].")
    
    
    # --- 11. Merge all dataframes ---
    message("\nMerging dataframes...")
    
    
    # --- Chemical Data ---
    
    # Merge 1: QC_Results with QC_Samples
    merged_df <- left_join(QC_Results, QC_Samples, 
                           by = c("Loc_ID", "QC_Sample", "QC_Type"))
    
    # Merge 2: The result with the main Results table
    merged_df <- left_join(merged_df, Results, 
                           by = c("Loc_ID", "Assoc_Samp", "Storet_Num"))
    
    # --- Merge 3: The result with the Analytes lookup table ---
    final_merged_df <- left_join(merged_df, Analytes, by = "Storet_Num") %>%
      mutate(
        # --- Handle NA replacement based on new rules ---
        
        # Rule for QC_Value: Replace NA with 0.5 * QC_LOQ only if QC_LOQ is available.
        QC_Value = if_else(is.na(QC_Value) & !is.na(QC_LOQ), 0.5 * QC_LOQ, QC_Value),
        
        # Rule for Field_Value: Replace NA with 0.5 * Field_LOQ only if Field_LOQ is available.
        Field_Value = if_else(is.na(Field_Value) & !is.na(Field_LOQ), 0.5 * Field_LOQ, Field_Value),
        
        # --- Unit conversion logic (applied after NA replacement) ---
        QC_Value = ifelse(UnitsQC == "ug/L" & UnitsField == "mg/L", QC_Value / 1000, QC_Value),
        UnitsQC = ifelse(UnitsQC == "ug/L" & UnitsField == "mg/L", "mg/L", UnitsQC),
        
        Field_Value = ifelse(UnitsField == "ug/L" & UnitsQC == "mg/L", Field_Value / 1000, Field_Value),
        UnitsField = ifelse(UnitsField == "ug/L" & UnitsQC == "mg/L", "mg/L", UnitsField)
      )
    
    final_merged_df <- final_merged_df %>%
      select(
        Loc_ID,
        QC_Sample,
        QC_Lab_ID,
        Storet_Num,
        Parameter,
        Sample_Type,
        QC_Value,
        Text_QC_Value,
        UnitsQC,
        QC_LOQ,
        QC_Type,
        QC_Lab_Sample_Num,
        Field_Lab_ID,
        Assoc_Samp,
        Field_Value,
        Text_Field_Value,
        UnitsField,
        Field_LOQ,
        Field_Lab_Sample_Num
      )
    
    
    message("\nDeduplicating the Final_Merged_Data dataframe...")
    
    # --- Step 1: Isolate rows that are NOT true duplicates ---
    discrepancies_and_uniques_chem <- final_merged_df %>%
      group_by(Loc_ID, Assoc_Samp, Storet_Num, QC_Type) %>%
      filter(n_distinct(Text_QC_Value) > 1 | n_distinct(Text_Field_Value) > 1 | n() == 1) %>%
      ungroup()
    
    # --- Step 2: Isolate and deduplicate the TRUE duplicates ---
    true_duplicates_chem <- final_merged_df %>%
      group_by(Loc_ID, Assoc_Samp, Storet_Num, QC_Type) %>%
      filter(n_distinct(Text_QC_Value) == 1 & n_distinct(Text_Field_Value) == 1 & n() > 1) %>%
      slice(1) %>%
      ungroup()
    
    # --- Step 3: Combine the two sets back into the final dataframe ---
    final_merged_df <- bind_rows(
      discrepancies_and_uniques_chem, 
      true_duplicates_chem
    )
    
    message("-> Deduplication of Final_Merged_Data complete.")
    
    
    # --- Biological Data ---
        
    message("\nMerging biological data ...")
        
    # Merge 4: QC_Bioresults with QC_Biosamples
    merged_bio_df <- left_join(QC_Bioresults, QC_Biosamples,
                               by = c("Loc_ID", "QC_Sample", "QC_Type", "Sample_Type"))
    
    # --- Merge 5: merged bio_df with applicable Assoc_Samp rows from BioResults
    final_bio_comparison <- full_join(
      merged_bio_df,
      BioResults,
      by = c("Loc_ID", "Assoc_Samp" = "OG_Sample", "Taxacode")
    ) %>%
      
      filter(QC_Sample %in% QC_samp_list) %>%
      
      # --- Create the definitive Sample_Type column ---
      mutate(
        Sample_Type = coalesce(Sample_Type.x, Sample_Type.y),
        Analysis_Method = coalesce(Analysis_Method.x, Analysis_Method.y)
      ) %>%
      
      # --- Join with the lookup tables LAST ---
      left_join(Taxa, by = "Taxacode") %>%
      left_join(Sample_Type_Codes, by = "Sample_Type") %>%
      
      # --- Rename all columns at once ---
      rename(
        OG_Sample = Assoc_Samp,
      ) %>%
      
      # --- Handle NA values for counts, densities, biovolume ---
      mutate(
        Count_QC = replace_na(Count_QC, 0),
        Count_OG = replace_na(Count_OG, 0),
        Density_QC = replace_na(Density_QC, 0),
        Density_OG = replace_na(Density_OG, 0),
        Biovolume_QC = replace_na(Biovolume_QC, 0),
        Biovolume_OG = replace_na(Biovolume_OG, 0)
      ) %>%
      
      # --- Select and reorder the final columns ---
      select(
        # Primary Sample Identifiers
        Loc_ID,
        OG_Sample,
        QC_Sample,
        
        # Sample & QC Type Information
        Sample_Type,
        Type_Name,
        QC_Type,
        Field_Lab_ID,
        QC_Lab_ID,
        
        # Taxon Information
        Sci_Name,
        Common_Name,
        ITIS_TSN,
        Taxacode,
        
        # Original Sample Measurements
        Count_OG,
        Density_OG,
        Density_Units_OG,
        Biovolume_OG,
        BV_Units_OG,
        
        # QC Sample Measurements
        Count_QC,
        Density_QC,
        Density_Units_QC,
        Biovolume_QC,
        BV_Units_QC,
        
        # Methodology
        Analysis_Method
      )
    
    message("-> Biological data successfully joined and cleaned.")
    
    
    # --- 12. Create Locations Table ---
    # Get a list of unique Loc_IDs that are present in the Results table.
    unique_locs_in_results <- Results %>%
      distinct(Loc_ID)
    
    # Use a semi_join to filter the main Locations table.
    locations <- semi_join(
      locations,
      unique_locs_in_results,
      by = "Loc_ID"
    ) %>%
      # Select only the final columns you need.
      select(Loc_ID,  Loc_Desc, Latitude_DD, Longitude_DD,)
    
    
    # --- 13. Return a named list of all created dataframes ---
    message("\nData preparation and merging complete.")
    return(
      list(
        # Chemical Data
        QC_Results_Filtered = QC_Results,
        QC_Samples_Raw = QC_Samples,
        Results_Raw = Results,
        Analytes_Raw = Analytes,
        Final_Merged_Data = final_merged_df,
        
        
        # Biological Data
        QC_Bioresults_Filtered = QC_Bioresults,
        QC_Biosamples_Raw = QC_Biosamples,
        BioResults_Raw = BioResults,
        Sample_Type_Codes_Raw = Sample_Type_Codes,
        Taxa_Raw = Taxa,
        Final_Merged_BioData = final_bio_comparison,
        
        # Metadata
        Locations_Data = locations
      )
    )
    
  }, error = function(e) {
    message("An error occurred during the database operation:")
    message(e$message)
    return(NULL)
    
  }, finally = {
    if (!is.null(con) && dbIsValid(con)) {
      dbDisconnect(con)
      message("\nDatabase connection has been closed.")
    }
  })
}