#GSO work order script 
#this is going to knit a markdown document 
#and have a download csv option 
#Author: Nicholas Manna
#date: 12-10-2020 (Dec 10)

#0.0 load libraries -------
#shiny
library(shiny)
#for rendering the report
library(markdown)
#shiny themes for colors
library(shinythemes)
#knitting report
library(knitr)
#shinyjs for java script functions
library(shinyjs)
#reactable for interactive table
library(reactable)
#reading excel document
library(readxl)
#connect to cityworks
library(odbc)
#render breakdown table
library(kableExtra)
#data manipulation
library(tidyverse)
# package version
library(renv)

#1.0 user interface (UI) ------
ui <- fluidPage(theme = shinytheme("cerulean"),
  titlePanel("Public GSI-Related Issue Tracking Work Order Query"),
  h5("The Public GSI-Related Issue Tracking Work Order Query intakes a spreadsheet with column headers \"SMP ID\", \"Component ID\", and \"Number\" (which represent Work Number). Common Findings is the default sheet name, and skipping the first 5 rows is the default for finding where the data starts. This app uses this spreadsheet to query Cityworks for updated work order statuses and child work orders. The user can download the markdown report and .csv separately."), 
    sidebarPanel(
      #select the excel file to use
           fileInput("excel_input", "Select Excel file", accept = ".xlsx"),
           #write in the sheet name
           selectInput("sheet_name", "Sheet Name", choices = c("Issues - GSI Only", "Issues - GSI & Nearby Infrast", ""), selected = NULL),
           #how many rows to skip before headers?
           numericInput("skip_rows", "Skip Rows", value = 5),
           #Download Report button
           downloadButton("report", "Download Report"), 
           #Download .csv button
           downloadButton("csv", "Download .csv")
  ), 
  #need to include this for shinyjs to work
  useShinyjs()
  )

#2.0 server ------
server <- function(input, output, session){
  
  #2.0.1 set up --------
  rv <- reactiveValues()
  
  #2.0.2 toggle states--------
  #enable/disable buttons 
  observe(toggleState(id = "report", condition = !is.null(input$excel_input)))
  observe(toggleState(id = "csv", condition = !is.null(input$excel_input)))
  
  #2.1 data manipulation ----
  #input a spreadsheet, receive updated WO data
  #i moved all the data manipulation from the markdown document to here. then it is passed to markdown and the csv separately. they are generated with different buttons here
  data_return <- reactive({
    #set the excel input to a variable
    excel_file <- input$excel_input
    #if the input is null, stop here
    if(is.null(excel_file)){return(NULL)}
    
    #get work order
    work_order <- read_xlsx(path = excel_file$datapath, sheet = input$sheet_name, skip = input$skip_rows) %>% 
      dplyr::select( "SMP ID", "Component ID", "Number") %>% 
      dplyr::rename("work_order" = "Number", "component_id" = "Component ID",  "smp_id" = "SMP ID")
    #split work orders that are in the same row
    work_order_list <- strsplit(work_order$work_order, split = ",")
    
    #create a dataframe that has a separate row for each work order
    work_order_split <- data.frame(smp_id = rep(work_order$smp_id, sapply(work_order_list, length)),
                                   component_id = rep(work_order$component_id, sapply(work_order_list, length)), 
                                   work_order = as.numeric(unlist(work_order_list)))
    
    #create a dataframe that does not have NAs
    work_order_no_NAs <- work_order_split[complete.cases(work_order_split), ]
    
    #concatenate work orders for search within Azteca.WORKORDER
    WO_concat <- paste(work_order_no_NAs$work_order, collapse = ", ")
    
    WO_concat_cast <- paste(work_order_no_NAs$work_order, collapse = " AS nvarchar), CAST(")
    
    WO_cast <- paste0("CAST( ", WO_concat_cast, " AS nvarchar)")
    
    #connect to cityworks
    # cw <- dbConnect(odbc(), dsn = "Cityworks", uid = Sys.getenv("cw_uid"), pwd = Sys.getenv("cw_pwd"))
    cw <- dbConnect(odbc(),
                    Driver = "ODBC Driver 17 for SQL Server",
                    Server = "pwdcwsqlP.pwd.phila.local",
                    Database = "PWD_Cityworks",
                    uid = Sys.getenv("cw_uid"),
                    pwd = Sys.getenv("cw_pwd"))
    
    #query work orders
    # browser()
    work_order_query <- paste0("SELECT WORKORDERID, DESCRIPTION, INITIATEDATE, ACTUALSTARTDATE, ACTUALFINISHDATE, SOURCEWOID, STATUS FROM Azteca.WORKORDER where WORKORDERID IN (", WO_concat, ") OR SOURCEWOID IN (", WO_cast, ")")
    
    result <- dbGetQuery(cw, work_order_query) %>% 
      mutate_at(c("WORKORDERID", "SOURCEWOID"),  as.numeric) %>%
      mutate_at(c("INITIATEDATE", "ACTUALSTARTDATE", "ACTUALFINISHDATE"), as.Date) 
    
    #concatenate work orders for search within Azteca.WOCOMMENT
    WO_concat_quotes <- paste(result$WORKORDERID, collapse = "', '")
    
    #query and group comments
    comment_query <- paste0("SELECT WORKORDERID, COMMENTS FROM Azteca.WOCOMMENT where WORKORDERID IN ('", WO_concat_quotes, "')") 
    
    comments <- dbGetQuery(cw, comment_query) %>% 
      mutate_at("WORKORDERID", as.numeric)
    
    comments_grouped <- comments %>% 
      group_by(WORKORDERID) %>% 
      summarise(COMMENTS = paste(COMMENTS, collapse = ", "))
    
    #order results to match the order within the spreadsheet
    ordered_result_step_one <- work_order_no_NAs %>% 
      dplyr::left_join(result, by = c("work_order" = "WORKORDERID")) %>% 
      #dplyr::left_join(result, by = c("work_order" = "SOURCEWOID")) %>%
      dplyr::select(-component_id)
    
    ordered_result_step_two <- work_order_no_NAs %>% 
      dplyr::left_join(result, by = c("work_order" = "SOURCEWOID")) %>%
      dplyr::select(-component_id) %>% 
      dplyr::rename("SOURCEWOID" = "work_order", "work_order" = "WORKORDERID") %>% 
      dplyr::filter(!is.na(work_order))
    
    ordered_result <- bind_rows(ordered_result_step_one, ordered_result_step_two)
    
    #attach the query results to the original list from the spreadsheet
    results_with_NAs <- dplyr::left_join(work_order_split, result, by = c("work_order" = "WORKORDERID"))
    
    #changed from ordered_resut_step_2 to ordered_result, 10/19/2022
    results_check <- bind_rows(results_with_NAs, ordered_result)
    
    #group results by the status, then count
    breakdown <- results_with_NAs %>% group_by(STATUS) %>% tally()
    
    #prepare for writing to csv
    #arrange table to match spreadsheet, so it can be copy + pasted from a csv into it
    copying <- results_check %>% left_join(comments_grouped, by = c("work_order" = "WORKORDERID")) %>% 
      group_by(smp_id, component_id) %>% 
      summarise(work_order = paste(work_order, collapse = ", "), 
                type = paste(DESCRIPTION, collapse = ", "),
                status = paste(STATUS, collapse = ", "), 
                source_wo = paste(SOURCEWOID, collapse = ", "),
                comments = paste(COMMENTS, collapse = ", "))
    
    #remove leading and trailing whitespace from comments
    copying$comments <- trimws(copying$comments)
    
    #left join to original list to arrange
    copying_ordered <- dplyr::left_join(work_order, copying, by = c("smp_id", "component_id", "work_order"))
    
    copying_extra <- dplyr::anti_join(copying, copying_ordered, by = c("work_order"))
    
    copying <- bind_rows(copying_ordered, copying_extra)
    
    rl <- list(work_order = work_order, result = result, breakdown = breakdown, ordered_result = ordered_result,  copying = copying, comments = comments)
    
    return(rl)
  })
  
  #2.2 markdown -------
  #render and download markdown document
  output$report <- downloadHandler(
    filename = function(){
      paste0(Sys.Date(), "_WO_Coordination.html")
    },
    
    
    content = function(file){
      
      # Copy the report file to a temporary directory before processing it, in
      # case we don't have write permissions to the current working dir (which
      # can happen when deployed).
      
      tempReport <- file.path(tempdir(), "WO_coordination.rmd")
      file.copy("WO_coordination.rmd", tempReport, overwrite = TRUE)
      
      #set up parameters to pass to .rmd document
      params <- list(data_return = data_return(), 
                     datapath = input$excel_input$name, 
                     sheet = input$sheet_name)
      
      # Knit the document, passing in the `params` list, and eval it in a
      # child of the global environment (this isolates the code in the document
      # from the code in this app).
      
      #  report_location <- "O:///Watershed Sciences//GSI Monitoring//07 Databases and Tracking Spreadsheets//12 Coordination Master Record//GSO Coordination//GSO Work Order Script//GSO_coordination.rmd"
      rmarkdown::render(tempReport,
                        output_file = file,
                        params = params,
                        envir=new.env(parent=globalenv()))
    }
  )
  
  #2.3 download csv --------
  output$csv <- downloadHandler(
    filename = function(){
      paste0(Sys.Date(), "_WO_Coord.csv")
    },
    
    content = function(file){
      write.csv(data_return()[["copying"]], file)
    }
  )
}

#3.0 runApp -------
#Run this function to run the app!
shinyApp(ui, server)