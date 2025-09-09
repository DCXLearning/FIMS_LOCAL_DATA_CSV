# ---- CSV-backed clean_all_kobo() ----
library(readr)
library(dplyr)
library(tibble)

clean_all_kobo <- function(
    data_dir = "Data",
    strict = TRUE   # if FALSE, missing files return empty tibbles instead of stop()
) {
  csv_path <- function(name) file.path(data_dir, paste0(name, ".csv"))
  read_tbl <- function(name) {
    path <- csv_path(name)
    if (!file.exists(path)) {
      msg <- paste("File not found:", path)
      if (strict) stop(msg) else { message(msg); return(tibble()) }
    }
    readr::read_csv(
      path,
      locale = readr::locale(encoding = "UTF-8"),
      na = c("", "NA", "NaN", "null", "NULL"),
      guess_max = 200000,
      show_col_types = FALSE
    )
  }
  
  # One table = one file
  main_df       <- read_tbl("KOBO_FAI_FISHCATCHING_MAIN_MART")
  nat_df        <- read_tbl("KOBO_FAI_NAT_FISHCATCHING_MAIN_MART")
  aqu_df        <- read_tbl("KOBO_FIA_AQU_FISHCATCH_FISHCATCHING_MART")
  processing_df <- read_tbl("KOBO_PROCESSING_FIA_FISHCATCH_MART")
  patrol_df     <- read_tbl("KOBO_PATROL_FIA_FISHCATCH_MART")
  choices_df    <- read_tbl("KOBO_CHOICES_FISHERIES_MART")  # <- included as requested
  
  message(
    sprintf("Loaded rows → main:%s nat:%s aqu:%s processing:%s patrol:%s choices:%s",
            nrow(main_df), nrow(nat_df), nrow(aqu_df), nrow(processing_df), nrow(patrol_df), nrow(choices_df))
  )
  
  # Return in the same shape you already use
  list(
    main       = main_df,
    nat        = nat_df,
    aqu        = aqu_df,
    processing = processing_df,
    patrol     = patrol_df,
    choices    = choices_df
  )
}
# ---- end CSV-backed clean_all_kobo() ----
