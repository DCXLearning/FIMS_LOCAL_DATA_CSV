# 
# 
# # install.packages(c("DBI","RMariaDB","readr"))
# library(DBI)
# library(RMariaDB)  # better for MariaDB/MySQL than RMySQL
# library(readr)
# 
# # ---- DB CONNECTION (UTF-8) ----
# con <- dbConnect(
#   RMySQL::MySQL(),
#   dbname   = "fishstat_database",
#   host     = "104.248.155.82",
#   port     = 3306,
#   user     = "root",
#   password = "Mymariadb123",
#   encoding = "UTF-8"         # ensure UTF-8 connection
# )
# dbExecute(con, "SET NAMES utf8mb4;")  # enforce full Unicode (Khmer, emoji, etc.)
# 
# # ---- TABLES TO EXPORT ----
# tables <- c(
#   "KOBO_CHOICES_FISHERIES_MART",
#   "KOBO_FAI_FISHCATCHING_MAIN_MART",
#   "KOBO_FAI_NAT_FISHCATCHING_MAIN_MART",
#   "KOBO_FIA_AQU_FISHCATCH_FISHCATCHING_MART",
#   "KOBO_PATROL_FIA_FISHCATCH_MART",
#   "KOBO_PROCESSING_FIA_FISHCATCH_MART"
# )
# 
# # ---- OUTPUT FOLDER ----
# out_dir <- "D:\\DCX\\FIMS\\25_08_2025_Anual_Report_10_08_am\\Data"
# 
# # ---- EXPORT LOOP (writes Excel-friendly UTF-8 CSV with BOM) ----
# for (tbl in tables) {
#   df <- dbGetQuery(con, paste0("SELECT * FROM ", tbl))
#   
#   # (Optional) ensure all character columns are marked as UTF-8 in R
#   chr_cols <- vapply(df, is.character, logical(1))
#   df[chr_cols] <- lapply(df[chr_cols], enc2utf8)
#   
#   out_path <- file.path(out_dir, paste0(tbl, ".csv"))
#   
#   # readr::write_excel_csv writes UTF-8 **with BOM** so Excel shows Khmer correctly
#   write_excel_csv(df, out_path)
#   message("Exported: ", out_path)
# }
# 
# dbDisconnect(con)
library(shiny)
library(shinyjs)
library(DBI)
library(RMySQL)
library(readr)

tables <- c(
  "KOBO_CHOICES_FISHERIES_MART",
  "KOBO_FAI_FISHCATCHING_MAIN_MART",
  "KOBO_FAI_NAT_FISHCATCHING_MAIN_MART",
  "KOBO_FIA_AQU_FISHCATCH_FISHCATCHING_MART",
  "KOBO_PATROL_FIA_FISHCATCH_MART",
  "KOBO_PROCESSING_FIA_FISHCATCH_MART"
)

# Use forward slashes on Windows for safety
output_dir <- "Data"

refreshData <- function() {
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
  }
  if (file.access(output_dir, 2) != 0) {
    stop(sprintf("No write permission to folder: %s", output_dir))
  }
  
  con <- dbConnect(
    RMySQL::MySQL(),
    dbname   = "fishstat_database",
    host     = "104.248.155.82",
    port     = 3306,
    user     = "root",
    password = "Mymariadb123",
    encoding = "UTF-8"
  )
  on.exit(dbDisconnect(con), add = TRUE)
  try(dbExecute(con, "SET NAMES utf8mb4;"), silent = TRUE)
  
  results <- data.frame(
    table = character(), status = character(), path = character(),
    stringsAsFactors = FALSE
  )
  
  ts <- format(Sys.time(), "%Y%m%d_%H%M%S")
  locked_dir <- file.path(output_dir, "locked")
  if (!dir.exists(locked_dir)) dir.create(locked_dir, showWarnings = FALSE)
  
  for (tbl in tables) {
    out_path <- file.path(output_dir, paste0(tbl, ".csv"))
    tmp_path <- paste0(out_path, ".tmp")
    
    status <- "ok"; saved_path <- out_path
    
    # query
    df <- dbGetQuery(con, paste0("SELECT * FROM ", tbl))
    chr <- vapply(df, is.character, logical(1))
    if (any(chr)) df[chr] <- lapply(df[chr], enc2utf8)
    
    # write temp
    ok_tmp <- TRUE
    tryCatch(
      { readr::write_excel_csv(df, tmp_path) },
      error = function(e){ ok_tmp <<- FALSE; status <<- paste0("write_tmp_failed: ", e$message) }
    )
    if (!ok_tmp) {
      results <- rbind(results, data.frame(table=tbl, status=status, path="", stringsAsFactors=FALSE))
      next
    }
    
    # try replace target
    locked <- FALSE
    if (file.exists(out_path)) {
      ok_rm <- suppressWarnings(file.remove(out_path))
      if (!ok_rm && file.exists(out_path)) locked <- TRUE
    }
    
    if (!locked) {
      ok_mv <- suppressWarnings(file.rename(tmp_path, out_path))
      if (!ok_mv) {
        # fallback to timestamp if rename fails for any reason
        alt <- file.path(output_dir, paste0(tbl, "_", ts, ".csv"))
        ok_mv2 <- suppressWarnings(file.rename(tmp_path, alt))
        if (ok_mv2) {
          status <- "renamed_fallback"
          saved_path <- alt
        } else {
          unlink(tmp_path)
          status <- "rename_failed"
          saved_path <- ""
        }
      } else {
        status <- "ok"
        saved_path <- out_path
      }
    } else {
      # file locked by Excel → save under locked/<tbl>_YYYYmmdd_HHMMSS.csv
      alt <- file.path(locked_dir, paste0(tbl, "_", ts, ".csv"))
      ok_mv3 <- suppressWarnings(file.rename(tmp_path, alt))
      if (ok_mv3) {
        status <- "locked_saved_as_alt"
        saved_path <- alt
      } else {
        unlink(tmp_path)
        status <- "locked_and_alt_failed"
        saved_path <- ""
      }
    }
    
    results <- rbind(results, data.frame(table=tbl, status=status, path=saved_path, stringsAsFactors=FALSE))
  }
  
  return(results)
}

server <- function(input, output, session) {
  useShinyjs()
  
  observeEvent(input$syncBtn, {
    shinyjs::disable("syncBtn")
    withProgress(message = "Syncing data…", value = 0, {
      tryCatch({
        incProgress(0.3, detail = "Exporting tables")
        res <- refreshData()
        incProgress(1, detail = "Done")
        
        # Build a compact summary
        ok_n     <- sum(res$status == "ok")
        locked_n <- sum(res$status == "locked_saved_as_alt")
        other_n  <- nrow(res) - ok_n - locked_n
        
        msg <- sprintf("✅ %d ok, 🔒 %d locked (saved as /locked/*.csv), ⚠️ %d other.",
                       ok_n, locked_n, other_n)
        showNotification(msg, type = if (other_n==0) "message" else "warning", duration = 6)
        
        # Print details to R console (helpful for debugging)
        print(res)
        
        # Optional: show a modal with details for the user
        # showModal(modalDialog(
        #   title = "Sync summary",
        #   renderTable(res, bordered = TRUE, striped = TRUE, hover = TRUE),
        #   easyClose = TRUE, size = "l"
        # ))
        
      }, error = function(e) {
        showNotification(paste("❌ Sync failed:", e$message), type = "error", duration = 10)
      }, finally = {
        shinyjs::enable("syncBtn")
      })
    })
  })
}
