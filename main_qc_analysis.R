# --- Checking and Loading Necessary Packages ---
message("Checking for required packages...")
required_packages <- c("shiny", "dplyr", "DBI", "odbc", "DT", "knitr", "rmarkdown", "openxlsx", "purrr")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message(paste("Package not found:", pkg, "- Installing..."))
    install.packages(pkg)
  }
  # Load the package
  library(pkg, character.only = TRUE)
}
message("All required packages are loaded.")


# --- 1. Define the GUI function ---
get_user_input <- function() {
  ui <- fluidPage(
    titlePanel("QC and Program Activity Closeout Report"),
    sidebarLayout(
      sidebarPanel(
        helpText("Select the date range for the data quality screening and program activity closeout report generation."),
        
        dateRangeInput("date_range", 
                       label = "1. Select Date Range:",
                       start = Sys.Date() - 365,
                       end = Sys.Date()),
        
        actionButton("run_analysis", "Run Analysis", class = "btn-primary", icon = icon("play"))
      ),
      
      mainPanel(
        h3("Ready to Go!"),
        p("Once you have made your selection in the panel on the left, click the 'Run Analysis' button to begin."),
        br(),
        p("This window will close automatically, and the analysis will start in your RStudio console.")
      )
    )
  )
  
  # Server Logic
  server <- function(input, output, session) {
    observeEvent(input$run_analysis, {
      user_selections <- list(
        start_date = as.Date(input$date_range[1]),
        end_date = as.Date(input$date_range[2])
      )
      stopApp(user_selections)
    })
  }
  
  # Run the App
  runApp(shinyApp(ui, server), launch.browser = .rs.invokeShinyPaneViewer)
}



# --- 2. Define the main analysis workflow ---
main <- function() {
  
  # --- Step 1: Call the GUI and get user input ---
  message("\nLaunching user input GUI...")
  user_input <- get_user_input()
  
  # --- Step 2: Check for valid input and print summary ---
  if (is.null(user_input)) {
    stop("Analysis cancelled: The GUI window was closed before submitting.", call. = FALSE)
  }
  
  message("\n--- Analysis Parameters Captured ---")
  message("Start Date: ", user_input$start_date)
  message("End Date: ", user_input$end_date)
  message("------------------------------------\n")
  
  
  # --- 3. Source External Function Scripts --- 
  message("Sourcing data preparation functions...")
  source("01_data_preparation_functions.R")
  
  
  # --- 4. Execute the Workflow ---
  message("Starting data preparation...\n")
  data_list <- fetch_and_prep_tables(
    start_date = user_input$start_date,
    end_date   = user_input$end_date
  )
  
# --- Step 5: Generate Calibration Tracking Sheet ---
  if (!is.null(data_list$Results_Raw) && !is.null(data_list$Analytes_Raw)) {
    
    message("\nSourcing calibration sheet generator...")
    source("05_calibration_sheet.R")
    
    
    calibration_results <- generate_calibration_sheet(
      results_df = data_list$Results_Raw,
      analytes_df = data_list$Analytes_Raw,
      user_input = user_input
    )
    
    data_list$Calibration_Data <- calibration_results
    
    
  } else {
    message("Results or Analytes data not available. Skipping calibration sheet generation.")
  }
  
  
  # --- 6a. Source and Execute Data Screening ---
  if (!is.null(data_list$Final_Merged_Data) && nrow(data_list$Final_Merged_Data) > 0) {
    
    message("\nSourcing data screening functions...")
    source("02_data_screening_functions.R")
  
    screening_results <- screen_qc_data(
      merged_dataframe = data_list$Final_Merged_Data
    )
    
    data_list$Screened_Data <- screening_results$All_Screened_Data
    data_list$Split_Screened_Data <- screening_results$Split_by_QC_Type
    data_list$Screened_Sediment_Data <- screening_results$Screened_Sediment_Data
    
  } else {
    message("Data preparation failed or returned no data. Skipping screening.")
  }
  
  
  # --- 6b. Screen Biological Data ---
  if (!is.null(data_list$Final_Merged_BioData) && nrow(data_list$Final_Merged_BioData) > 0) {
    
    bray_curtis_results <- calculate_bray_curtis(
      final_bio_df = data_list$Final_Merged_BioData
    )
    
    data_list$Bray_Curtis_Results <- bray_curtis_results
    
    message("\nWorkflow complete")
  }
  

  # --- 7. Generate Final Excel Report ---
  
  if (!is.null(data_list$Summary_Statistics) || !is.null(data_list$Bio_Summary_Statistics)) {
    
    message("\nSourcing Excel report generator...")
    source("04_generate_report.R") 
    
    # Define the folder path and filename
    output_folder <- "O:/ED/Private/Water Quality/Co-op & Others Oversight/Matthew Schaaf/QualityControl_and_ProgramActivityCloseoutReport/WaterQualityQC/reports"
    start_date_str <- format(user_input$start_date, "%Y-%m-%d")
    end_date_str   <- format(user_input$end_date, "%Y-%m-%d")
    report_filename <- paste0("QC_Report_", start_date_str, "_to_", end_date_str, ".xlsx")
    full_output_path <- file.path(output_folder, report_filename)
    
    generate_excel_report(
      user_input      = user_input,
      bray_curtis     = data_list$Bray_Curtis_Results,
      bio_data        = data_list$Final_Merged_BioData,
      sediment_data   = data_list$Screened_Sediment_Data, 
      chem_split_data = data_list$Split_Screened_Data,
      chem_data       = data_list$Screened_Data,
      locations       = data_list$Locations_Data,
      output_filepath = full_output_path,
      calibration_data = data_list$Calibration_Data$Final_Calibration_Sheet
    )
    
  } else {
    message("No statistics available. Skipping report generation.")
  }
  
  message("\nWorkflow complete. Reorganizing final results list...")
  
  # --- Reorganize data_list into a cleaner structure ---
  
  chem_list <- list(
    QC_Results_Filtered = data_list$QC_Results_Filtered,
    QC_Samples_Raw = data_list$QC_Samples_Raw,
    Results_Raw = data_list$Results_Raw,
    Analytes_Raw = data_list$Analytes_Raw,
    Final_Merged_Data = data_list$Final_Merged_Data
  )
  
  bio_list <- list(
    QC_Bioresults_Filtered = data_list$QC_Bioresults_Filtered,
    QC_Biosamples_Raw = data_list$QC_Biosamples_Raw,
    Bioresults_Raw = data_list$BioResults_Raw,
    Sample_Type_Codes = data_list$Sample_Type_Codes_Raw,
    Taxa_Raw = data_list$Taxa_Raw,
    Final_Merged_BioData = data_list$Final_Merged_BioData
  )
  
  summary_list <- list(
    Chem_Summary_Statistics = data_list$Summary_Statistics,
    Bio_Summary_Statistics = data_list$Bio_Summary_Statistics
  )
  
  screened_list <- list(
    Screened_Chem_Data = data_list$Screened_Data,
    Split_Screened_Chem_Data = data_list$Split_Screened_Data,
    Screened_Sediment_Data = data_list$Screened_Sediment_Data
  )
  
  final_organized_list <- list(
    Chem = chem_list,
    Bio = bio_list,
    Summary = summary_list,
    Screened = screened_list,
    Bray_Curtis_Results = data_list$Bray_Curtis_Results,
    Locations_Data = data_list$Locations_Data,
    Calibration_Data = data_list$Calibration_Data
  )
  
  return(final_organized_list)
}

# --- 8. Execute the script ---
analysis_results <- main()
