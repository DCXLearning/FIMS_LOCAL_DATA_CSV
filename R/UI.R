
library(shiny)
library(shinyjs)
library(DT)

ui <- fluidPage(
  shinyjs::useShinyjs(),
  
  # ========================= CSS =========================
  tags$head(
    tags$style(HTML("
      body {
        font-family: 'Arial', sans-serif;
        overflow-x: hidden;
        margin: 0;
        padding: 0;
      }

      /* ----- Fixed header (top bar + centered tabs row) ----- */
      .fixed-header {
        position: fixed;
        top: 0; left: 0;
        width: 100%;
        background-color: #e44d26;
        color: white;
        padding: 10px 20px;
        z-index: 1050;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        display: block;
        box-sizing: border-box;
      }
      .header-top-row {
        display: flex;
        align-items: center;
        width: 100%;
        margin-bottom: 6px;
      }
      .header-title { font-size: 20px; font-weight: bold; margin-right: auto; }
      #openSidebarBtn, #closeSidebarBtn {
        background-color: transparent; color: white; border: none;
        font-size: 28px; padding: 0; cursor: pointer; transition: color .3s;
        outline: none !important; box-shadow: none !important;
      }
      #openSidebarBtn { margin-left: 10px; }
      #openSidebarBtn:hover, #closeSidebarBtn:hover { color: #18bc9c; }

      /* ----- Center the tabs ----- */
      .tabs-centered .nav-tabs {
        display: flex !important;
        justify-content: center;
        align-items: center;
        flex-wrap: wrap;
        border-bottom: none;
        margin-bottom: 0;
        background-color: #c0392b;
      }
      .tabs-centered .nav-tabs > li { float: none; }
      .tabs-centered .nav-tabs > li > a {
        color: white; border-radius: 0; border: none; padding: 8px 15px; margin-right: 0;
      }
      .tabs-centered .nav-tabs > li.active > a,
      .tabs-centered .nav-tabs > li.active > a:hover,
      .tabs-centered .nav-tabs > li.active > a:focus {
        color: #fff; background-color: #a02012; border: none; border-bottom: 3px solid #18bc9c;
      }
      .tabs-centered .nav-tabs > li > a:hover,
      .tabs-centered .nav-tabs > li > a:focus {
        background-color: #d14836; border-color: transparent;
      }

      /* If you ever put content inside tabPanels, remove this next line */
      .tab-content { display: none; }

      /* ----- Sidebar ----- */
      .sidebar {
        position: fixed; width: 260px; height: 100%;
        background-color: #2c3e50; padding: 20px; color: white;
        box-shadow: 2px 0 5px rgba(0,0,0,0.1);
        overflow-y: auto; top: 0; left: 0; transition: transform .3s; z-index: 2000;
        box-sizing: border-box; padding-top: 60px;
      }
      .sidebar.hidden { transform: translateX(-260px); z-index: 1000; }
      .sidebar h3 { color: #ecf0f1; margin-bottom: 25px; }
      .sidebar h4 { color: #18bc9c; margin-top: 30px; border-bottom: 1px solid #4a627d; padding-bottom: 5px; }
      .sidebar label { color: #bdc3c7; font-weight: bold; margin-bottom: 5px; display: block; }
      .sidebar select, .sidebar input[type='checkbox'] {
        width: 100%; margin-bottom: 15px; padding: 8px; border-radius: 5px;
        background-color: #34495e; color: white; border: 1px solid #4a627d; box-sizing: border-box;
      }
      .sidebar .selectize-input, .sidebar .selectize-dropdown {
        background-color: #34495e !important; color: white !important; border-color: #4a627d !important;
      }
      .sidebar .radio-inline { color: #bdc3c7; }

      /* ----- Main content (leave room for header) ----- */
      .main-content {
        margin-left: 260px; padding: 20px; background-color: #f8f9fa;
        transition: margin-left .3s; min-height: 100vh; box-sizing: border-box;
        padding-top: 140px;  /* header height + tabs row */
      }
      .main-content.full-width { margin-left: 0; }

      /* Buttons */
      .btn-success { background-color: #18bc9c; border-color: #18bc9c; color: white; font-weight: bold; }
      .btn-success:hover { background-color: #15a88c; border-color: #15a88c; }
      .btn-primary { background-color: #3498db; border-color: #3498db; color: white; font-weight: bold; }
      .btn-primary:hover { background-color: #2980b9; border-color: #2980b9; }

      /* ----- Toolbar inside the National tab ----- */
      .tab-toolbar {
        display: flex; justify-content: space-between; align-items: center;
        gap: 12px; margin-bottom: 10px;
      }
      .tab-toolbar .controls {
        display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
      }
      .tab-toolbar .controls .form-group { margin-bottom: 0; }  /* tighten selectInput spacing */
      .tab-toolbar .controls .selectize-control { min-width: 260px; }
    "))
  ),
  
  # ========================= HEADER (centered tabs; no downloads here) =========================
  div(
    class = "fixed-header",
    div(
      class = "header-top-row",
      actionButton("openSidebarBtn", "", icon = icon("bars"), class = "btn-main-open"),
      span(" ", class = "header-title")
    ),
    div(
      class = "tabs-centered",
      tabsetPanel(
        id = "mainTabs",
        tabPanel("DOCS"),
        tabPanel("EXCEL")
      )
    )
  ),
  
  # ========================= SIDEBAR (no download buttons here) =========================
  div(
    id = "sidebarPanel", class = "sidebar",
    actionButton("closeSidebarBtn", "", icon = icon("times"), class = "btn-sidebar-close"),
    h3("Filters"),
    
    selectInput("Province_Group", "Select Province Group",
                choices = province_group_choices, selected = "All"),
    selectInput("province", "Select Province",
                choices = get_province_choices_for_group("All"), selected = "All"),
    selectInput("year", "Select Year",
                choices = year_choices, selected = "2025"),
    selectInput("quarter", "Select Quarter", choices = quarter_choices),
    selectInput("month",   "Select Month",   choices = month_choices),
    selectInput("area_en", "Select Area", choices = area_choices),
    checkboxInput("show_all", "Show All Rows", value = FALSE),
    
    tags$hr(),
    h4("Data"),
    actionButton("syncBtn", "Sync data", icon = icon("sync"), class = "btn-primary")
    # (no download UI in sidebar)
  ),
  
  # ========================= MAIN CONTENT =========================
  div(
    id = "mainContent",
    class = "main-content",
    
    # -------- TAB 1: National Summary (toolbar + tables) --------
    conditionalPanel(
      condition = "input.mainTabs == ''DOCS",
      
      # Toolbar that lives INSIDE this tab only
      div(
        class = "tab-toolbar",
        h4("សរុបជាតិ (National Summary)"),
        div(
          class = "controls",
          # Put the items selector here so it's not in the sidebar
          selectInput(
            "dl_choices",
            NULL,
            choices = c(
              "Report Status Table"                           = "reportStatusTable",
              "Fish Catch by Quarter"                         = "natPivotTable",
              "Fisheries Capture in Rice-Fields by Species"   = "riceFieldTable",
              "Freshwater Capture in Rice-Fields by Province" = "riceFieldProvinceTable",
              "Inland Production by Province"                 = "provinceProductionTable",
              "Aquaculture Production by Province"            = "aquaProvinceProductionTable",
              "Processed Fish Production by Quarter"          = "processingTable",
              "Freshwater, Marine and Aquaculture Chart"      = "summaryProductionChart"
              # "Aquaculture & Fingerling Production Chart"     = "aquaProductionChart"
            ),
            multiple = TRUE,
            selected = NULL
          ),
          downloadButton("download_report", "Download Report", class = "btn btn-success"),
          downloadButton("download_excel_report", "Excel Report", class = "btn btn-primary")
        )
      ),
      
      h4("Total number of statistical reports submitted by month"),
      DTOutput("reportStatusTable"),
      br(), tags$hr(), br(),
      
      h4("Fish Catch by Quarter"),
      DTOutput("natPivotTable"),
      br(), tags$hr(), br(),
      
      h4("Fisheries capture in rice-fields by species"),
      DTOutput("riceFieldTable"),
      br(), tags$hr(), br(),
      
      h4("Freshwater capture in rice-fields by province"),
      DTOutput("riceFieldProvinceTable"),
      br(), tags$hr(), br(),
      
      h4("Inland Production by Province"),
      DTOutput("provinceProductionTable"),
      br(), tags$hr(), br(),
      
      h4("Aquaculture Production by Province"),
      DTOutput("aquaProvinceProductionTable"),
      br(), tags$hr(), br(),
      
      h4("Processed Fish Production by Quarter"),
      radioButtons("processing_source_filter", "Select Source:",
                   choices = source_choices, inline = TRUE),
      DTOutput("processingTable")
    ),
    
    # -------- TAB 2: Detailed Tables (no download toolbar here) --------
    conditionalPanel(
      condition = "input.mainTabs == 'EXCEL'",
      
      h4("Marine Fish Catch Summary (តារាងត្រីសមុទ្រ)"),
      DTOutput("catch_table"),
      br(), tags$hr(), br(),
      
      h4("Aquaculture Ponds & Cages (ស្រះ បែ ស៊ង)"),
      DTOutput("aquaculture_pond_table"),
      br(), tags$hr(), br(),
      
      h4("Aquaculture Products (ផលវារីវប្បកម្ម)"),
      DTOutput("aquaculture_table"),
      br(), tags$hr(), br(),
      
      h4("Marine Fishing Products (ផលនេសាទសមុទ្រ)"),
      DTOutput("fishing_products_table"),
      br(), tags$hr(), br(),
      
      h4("Rice Field Catch (ផលក្នុងវាលស្រែ)"),
      DTOutput("fishing_products_rice_field_table"),
      br(), tags$hr(), br(),
      
      h4("Fishery Domain Catch (ផលក្នុងដែននេសាទ)"),
      DTOutput("fishing_products_fishery_domain_table"),
      br(), tags$hr(), br(),
      
      h4("Dai Fishing Catch (ផលឧបករណ៍ដាយ)"),
      DTOutput("fishing_products_dai_table"),
      br(), tags$hr(), br(),
      
      h4("Law Enforcement Incidents (បទល្មើស)"),
      DTOutput("law_enforcement_table")
    )
  )
)
