# UI
ui <- fluidPage(
  shinyjs::useShinyjs(),
  
  tags$head(
    tags$style(HTML("
      body {
        font-family: 'Arial', sans-serif;
        overflow-x: hidden; /* Prevent horizontal scrollbar on toggle */
        margin: 0; /* Reset default body margin */
        padding: 0; /* Reset default body padding */
      }

      /* Fixed Header/Navbar Styling */
      .fixed-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #e44d26; /* Example red color, similar to your reference image */
        color: white;
        padding: 10px 20px;
        z-index: 1050; /* Higher than sidebar (1000) when sidebar is HIDDEN, but sidebar will go higher when OPEN */
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        display: flex; /* Use flexbox for horizontal alignment of items */
        align-items: center; /* Vertically center items */
        box-sizing: border-box; /* Include padding in width */
        height: 100px; /* Increased height to accommodate tabs below logo/title */
        flex-wrap: wrap; /* Allow items to wrap to the next line if needed */
      }

      /* Style for the logo/title container within the fixed header */
      .header-top-row {
          display: flex;
          align-items: center;
          width: 100%; /* Take full width of the header */
          margin-bottom: 5px; /* Space between top row and tabs */
      }

      /* Style for the logo/title within the fixed header */
      .header-logo {
        height: 30px; /* Adjust logo size */
        margin-right: 15px;
      }

      .header-title {
        font-size: 20px;
        font-weight: bold;
        margin-right: auto; /* Pushes other items (like open button) to the right */
      }

      /* Initial state for sidebar */
      .sidebar {
        position: fixed; /* Stays in place when scrolling */
        width: 260px;
        height: 100%; /* --- IMPORTANT FIX: Full height --- */
        background-color: #2c3e50;
        padding: 20px;
        color: white;
        box-shadow: 2px 0 5px rgba(0,0,0,0.1);
        overflow-y: auto; /* Allows scrolling if content is too long */
        top: 0; /* --- IMPORTANT FIX: Starts from top of viewport --- */
        left: 0;
        transition: transform 0.3s ease-in-out; /* Smooth transition for hiding/showing */
        z-index: 2000; /* --- IMPORTANT FIX: Higher z-index to cover fixed header --- */
        box-sizing: border-box; /* Include padding in width/height calculation */
        padding-top: 60px; /* Space for the close button at the top of the sidebar */
      }

      /* State when sidebar is hidden */
      .sidebar.hidden {
        transform: translateX(-260px); /* Move completely off-screen to the left */
        z-index: 1000; /* Lower z-index when hidden, so it doesn't block clicks if anything overlaps */
      }

      /* Initial state for main content */
      .main-content {
        margin-left: 260px; /* Equal to sidebar width to make space */
        padding: 20px; /* Base padding for content */
        background-color: #f8f9fa;
        transition: margin-left 0.3s ease-in-out; /* Smooth transition for margin adjustment */
        min-height: 100vh; /* Ensure main content takes full viewport height */
        box-sizing: border-box; /* Include padding in width/height calculation */
        padding-top: calc(100px + 20px); /* Header height + desired padding for content below header */
      }

      /* State when sidebar is hidden - main content takes full width */
      .main-content.full-width {
        margin-left: 0; /* Remove left margin when sidebar is hidden */
      }

      /* Close button styling (inside sidebar) */
      #closeSidebarBtn {
        position: absolute; /* Positioned relative to .sidebar */
        top: 15px;
        right: 15px;
        z-index: 2001; /* Ensure button is above sidebar content, and above the fixed header */
        background-color: transparent;
        color: white; /* White icon for visibility on dark sidebar */
        border: none;
        font-size: 28px;
        padding: 0;
        cursor: pointer;
        transition: color 0.3s ease-in-out;
        outline: none !important; /* Remove blue outline on focus */
        box-shadow: none !important; /* Remove box shadow on focus/active */
      }
      #closeSidebarBtn:hover {
        color: #18bc9c;
      }

      /* Open button styling (INSIDE FIXED HEADER) */
      #openSidebarBtn {
        background-color: transparent;
        color: white; /* Changed to white to be visible on the red header */
        border: none;
        font-size: 28px;
        padding: 0;
        cursor: pointer;
        transition: color 0.3s ease-in-out;
        outline: none !important;
        box-shadow: none !important;
        margin-left: 10px; /* Space between logo/title and button */
      }
      #openSidebarBtn:hover {
        color: #18bc9c;
      }

      /* Styling for the tabsetPanel when it's inside the fixed header */
      .fixed-header .tabbable { /* Targets the tabsetPanel container */
        width: 100%; /* Make tabs take full width of the header */
        background-color: transparent; /* Ensure background matches header */
      }

      .fixed-header .nav-tabs {
        border-bottom: none; /* Remove default tab border */
        background-color: #c0392b; /* Slightly darker red for tab background */
      }

      .fixed-header .nav-tabs > li > a {
        color: white; /* White text for tabs */
        border-radius: 0; /* No rounded corners */
        border: none; /* No border */
        padding: 8px 15px; /* Adjust padding */
        margin-right: 0; /* Remove space between tabs */
      }

      .fixed-header .nav-tabs > li.active > a,
      .fixed-header .nav-tabs > li.active > a:hover,
      .fixed-header .nav-tabs > li.active > a:focus {
        color: #fff; /* Active tab text color */
        background-color: #a02012; /* Darker red for active tab */
        border: none;
        border-bottom: 3px solid #18bc9c; /* Highlight active tab with green border */
      }

      .fixed-header .nav-tabs > li > a:hover,
      .fixed-header .nav-tabs > li > a:focus {
        background-color: #d14836; /* Hover color for tabs */
        border-color: transparent;
      }

      /* Hide the default content area for tabs, as content is in main-content */
      .tab-content {
        display: none; /* Hide the default tab content area */
      }

      /* General styles (remain largely unchanged) */
      .sidebar h3 { color: #ecf0f1; margin-bottom: 25px; }
      .sidebar h4 { color: #18bc9c; margin-top: 30px; border-bottom: 1px solid #4a627d; padding-bottom: 5px;}
      .sidebar label { color: #bdc3c7; font-weight: bold; margin-bottom: 5px; display: block; }
      .sidebar select, .sidebar input[type='checkbox'] {
        width: 100%; margin-bottom: 15px; padding: 8px; border-radius: 5px;
        background-color: #34495e; color: white; border: 1px solid #4a627d;
        box-sizing: border-box;
      }
      .sidebar input[type='checkbox'] { width: auto; margin-left: 0; }
      .sidebar .selectize-input, .sidebar .selectize-dropdown {
        background-color: #34495e !important; color: white !important; border-color: #4a627d !important;
      }
      .sidebar .radio-inline { color: #bdc3c7; }

      h4 { color: #34495e; margin-top: 20px; margin-bottom: 15px; }
      .dataTables_wrapper .dataTables_filter input {
        border: 1px solid #ced4da; border-radius: .25rem; padding: .375rem .75rem; width: auto;
      }
      .btn-success { background-color: #18bc9c; border-color: #18bc9c; color: white; font-weight: bold; }
      .btn-success:hover { background-color: #15a88c; border-color: #15a88c; }

      .action-button:focus,
      .action-button:active {
        outline: none !important;
        box-shadow: none !important;
      }
    "))
  ),
  # --- Fixed Header/Navbar at the very top ---
  div(class = "fixed-header",
      div(class = "header-top-row",
          actionButton("openSidebarBtn", "", icon = icon("bars"), class = "btn-main-open"),
          # tags$img(src = "Picture1.jpg", class = "header-logo"), # Assuming Picture1.jpg is in your 'www' folder
          #  span("Ministry of Agriculture Forestry and Fisheries - Fisheries Administration", class = "header-title")
      ),
      tabsetPanel(
        id = "mainTabs",
        tabPanel("Catch & Aquaculture Tables"),
        tabPanel("Processed Fish Summary"),
        tabPanel("Charts & Summaries")
      )
  ),
  
  div(id = "sidebarPanel", class = "sidebar",
      actionButton("closeSidebarBtn", "", icon = icon("times"), class = "btn-sidebar-close"),
      h3("Filters"),
      selectInput("Province_Group", "Select Province Group",
                  choices = province_group_choices, selected = "All"),
      selectInput("province", "Select Province",
                  choices = get_province_choices_for_group("All"), selected = "All"),
      selectInput(
        "year", "Select Year",
        choices  = year_choices,
        selected = "2025"   # default
      ),
      selectInput("quarter", "Select Quarter", choices = quarter_choices),
      selectInput("month",   "Select Month",   choices = month_choices),
      # selectInput("province", "Select Province", choices = province_choices),
      selectInput("area_en", "Select Area", choices = area_choices),
      checkboxInput("show_all", "Show All Rows", value = FALSE),
      
      tags$hr(),
      h4("Download Report"),
      
      selectInput("dl_choices",
                  "Select Items for Report",
                  choices = c(
                    "Report Status Table" = "reportStatusTable",
                    "Fish Catch by Quarter" = "natPivotTable",
                    "Fisheries Capture in Rice-Fields by Species" = "riceFieldTable",
                    "Freshwater Capture in Rice-Fields by Province" = "riceFieldProvinceTable",
                    "Inland Production by Province" = "provinceProductionTable",
                    "Aquaculture Production by Province" = "aquaProvinceProductionTable",
                    "Processed Fish Production by Quarter" = "processingTable",
                    "Freshwater, Marine and Aquaculture Production Chart" = "summaryProductionChart",
                    "Aquaculture and Fingerling Production Chart" = "aquaProductionChart"
                  ),
                  multiple = TRUE,
                  selected = c(" ")
      ),
      downloadButton("download_report", "Download Report", class = "btn-success"),
      br(), br(),
      downloadButton("download_excel_report", "Download Excel Report", class = "btn-primary")
  ),
  div(id = "mainContent", class = "main-content",
      # Content for "Catch & Aquaculture Tables" tab
      conditionalPanel(
        condition = "input.mainTabs == 'Catch & Aquaculture Tables'",
        h4("Total number of statistical reports submitted by month"),
        DTOutput("reportStatusTable"),
        h4("Fish Catch by Quarter"),
        DTOutput("natPivotTable"),
        br(), br(),
        h4("Fisheries capture in rice-fields by species"),
        DTOutput("riceFieldTable"),
        br(), br(),
        h4("Freshwater capture in rice-fields by province"),
        DTOutput("riceFieldProvinceTable"),
        br(), br(),
        h4("Inland Production by Province"),
        DTOutput("provinceProductionTable"),
        br(), br(),
        h4("Aquaculture Production by Province"),
        DTOutput("aquaProvinceProductionTable")
      ),
      # Content for "Processed Fish Summary" tab
      conditionalPanel(
        condition = "input.mainTabs == 'Processed Fish Summary'",
        h4("Processed Fish Production by Quarter"),
        radioButtons("processing_source_filter", "Select Source:",
                     choices = source_choices, inline = TRUE),
        DTOutput("processingTable")
      ),
      # Content for "Charts & Summaries" tab
      conditionalPanel(
        condition = "input.mainTabs == 'Charts & Summaries'",
        h4("Freshwater, Marine and Aquaculture Production"),
        plotOutput("summaryProductionChart", height = "500px")
        # hr(),
        # fluidRow(
        #   column(width = 6,
        #          h4("Aquaculture and Fingerling Production"),
        #          plotOutput("aquaProductionChart", height = "500px")
        #   )
        # )
      )
  )
)