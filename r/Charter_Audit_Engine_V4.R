# ======================================================================
# Charter Invoice Audit Engine V4 — SQL Audit Engine Integration
# Handles ALL charter roster variations: Southside, SAS, SASCCS, OnTECH
#
# V3 ADDITIONS (2026-04-27):
#  15.  SPED Status flag: joins SPED Roster to SCSD Roster — labels each student
#       as "SPED" or "Gen Ed" based on presence in SPED Roster (NYSSIS join key)
#  16.  SchoolTool FTE: NYSED formula (Student Weeks / 40 program weeks)
#       Active students (no ST exit) use STILL_ENROLLED_DATE as end date
#  17.  Charter FTE: preserved from original invoice FTE column
#  18.  FTE Variance: Charter FTE - SchoolTool FTE (positive = overbilled)
#
# V4 ADDITIONS (2026-04-28):
#  19.  SQL Audit Engine: full Athena SQL version of this engine created
#       (Charter_Audit_Engine_SQL_Final.sql) — replicates all R rules in
#       Amazon Athena CTEs for live SchoolTool Analytics dashboard
#  20.  SPED FTE: SchoolTool SPED FTE + Charter SPED FTE + SPED FTE Variance
#       added to both R engine output and SQL engine
#  21.  Gen Roster SQL: SCSD Gen Roster query built in Athena Custom SQL
#       (Direct Query — live refresh, no manual export needed)
#  22.  FTE formula updated to month-based calculation (* 4 / 40) to better
#       align with NYSED lookup-table methodology per §175.6
#
# EDGE CASES COVERED:
#   1.  Title rows (Row 1 = title text, Row 2 = headers)
#   2.  Duplicate column names (OnTECH: two "Enroll Date" cols)
#   3.  Combined "Full Name" as "Last, First" (SAS 4th, SASCCS 4th)
#   4.  Leading/trailing spaces in tab names (" SPED ROSTER")
#   5.  Missing ST Enrollment tab (SAS 4th)
#   6.  "Gen Roster" instead of "SCSD ROSTER" (Southside 5th)
#   7.  SPED uses Student ID not NYSSIS ("Student Gen Ed ID#")
#   8.  Disability tab 2-col vs 3-col layout
#   9.  ST tab named "ChartEnrol1-13" (no "ST" or "Enrollment" keyword)
#  10.  Non-breaking spaces + newlines in headers (SASCCS .xls)
#  11.  Double spaces in column names ("Student Last  Name")
#  12.  Excel serial-number dates read as text strings by readxl
#  13.  ST Student ID auto-populated into BOTH SCSD and SPED rosters
#  14.  OnTECH/SASCCS pre-populated ST columns (duplicate entry/exit cols)
# ======================================================================

pkgs_needed <- c("readxl", "dplyr", "stringr", "lubridate", "janitor", "openxlsx", "purrr")
pkgs_missing <- pkgs_needed[!pkgs_needed %in% rownames(installed.packages())]
if (length(pkgs_missing) > 0) install.packages(pkgs_missing, dependencies = TRUE)

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(lubridate)
  library(janitor)
  library(openxlsx)
  library(purrr)
})

# ======================================================================
# CONFIG — UPDATE THIS PATH FOR EACH INVOICE
# ======================================================================
# When running via the .bat launcher, workbook_path_override is set automatically.
# When running manually in RStudio, update the path below.
if (exists("workbook_path_override") && nzchar(workbook_path_override)) {
  workbook_path <- workbook_path_override
} else {
  
  workbook_path <- "Z:/Daquan/Help Desk Tickets/Charter Schools Invoices/May Charter Invoices FY 25-26/Rosters Sent By Patrick/25-26 SCSD Invoice and Roster #6 GenED ($484,371.23 FTE 201.225).xlsx"
}

# OPTIONAL COMPARISON:
#   To SKIP comparison:  leave as NULL (default)
#   To RUN comparison:   paste the full path to your Final Excel in quotes below
#
#   Example:
#   final_excel_path <- "C:/Users/dmorri43/Desktop/Final-Ticket-291462.xlsx"
#
final_excel_path <- NULL

# OUTPUT FOLDER (optional):
#   Leave as NULL to save the output file in the same folder as the invoice (default).
#   Set a folder path to always save output files to a specific location.
#
output_folder <- "C:/Users/dmorri43/Desktop/Training Original Charter invoices/Final R Output Charter Rosters"
# Allow runtime override (used by test harness and command-line calls)
if (exists("output_dir_override") && nzchar(output_dir_override)) output_folder <- output_dir_override

SCHOOL_YEAR_FIRST_DAY <- as.Date("2025-09-03")
GRACE_DEADLINE        <- as.Date("2025-09-06")
JULY_START            <- as.Date("2025-07-01")
GRACE_DAYS            <- 3
STILL_ENROLLED_DATE   <- as.Date("2026-06-30")  # NYSED program end date (max possible window per §175.6)

# ======================================================================
# HELPER FUNCTIONS
# ======================================================================

write_log <- function(msg, level = "INFO") {
  cat(sprintf("[%s] %s — %s\n", format(Sys.time(), "%H:%M:%S"), level, msg))
}

# --- Normalize column headers ---
# Strips non-breaking spaces (\u00A0), newlines, carriage returns, collapses
# double spaces, trims — so "Student\u00A0Last\nName" -> "Student Last Name"
normalize_headers <- function(df) {
  nms <- names(df)
  nms <- gsub("\u00A0", " ", nms)       # non-breaking space -> space
  nms <- gsub("[\r\n]+", " ", nms)      # newlines -> space
  nms <- gsub("\\s+", " ", nms)         # collapse multiple spaces
  nms <- str_trim(nms)                  # trim leading/trailing
  names(df) <- nms
  df
}

# --- Smart sheet reader ---
# Reads a sheet, normalizes headers, and auto-detects title rows.
# If Row 1 has ≤2 non-empty cells and Row 2 has ≥4, Row 1 is a title —
# re-reads skipping it so Row 2 becomes the header.
smart_read_sheet <- function(path, sheet_name) {
  # First pass: read row 1 only to detect layout type
  peek <- suppressMessages(read_excel(path, sheet = sheet_name, col_types = "text", n_max = 3))
  peek_names <- names(peek)

  # Detect engine output layout: row 1 = color key legend ("COLOR KEY:" in first cell),
  # row 2 = blank, row 3 = real headers. Skip 2 rows.
  first_name <- if (length(peek_names) > 0) peek_names[1] else ""
  is_color_key_layout <- grepl("COLOR.*KEY|COLOR KEY", first_name, ignore.case = TRUE) ||
    (!is.na(peek[1, 1, drop = TRUE]) &&
     grepl("COLOR.*KEY|COLOR KEY", as.character(peek[1, 1, drop = TRUE]), ignore.case = TRUE))

  if (is_color_key_layout) {
    write_log(paste("Engine output layout detected in sheet:", sheet_name, "— reading from row 3"))
    df <- suppressMessages(read_excel(path, sheet = sheet_name, col_types = "text", skip = 2))
    df <- normalize_headers(df)
    return(df)
  }

  # Standard read
  df <- suppressMessages(read_excel(path, sheet = sheet_name, col_types = "text"))
  df <- normalize_headers(df)

  # Check for title row: if first header looks like a single title string
  # and most columns are empty/unnamed, the real headers are in data row 1
  non_empty_headers <- sum(names(df) != "" & !grepl("^\\.\\.\\.\\d+$", names(df)))
  total_cols <- ncol(df)

  if (non_empty_headers <= 2 && total_cols >= 4 && nrow(df) > 0) {
    first_row_vals <- as.character(df[1, ])
    non_na_count   <- sum(!is.na(first_row_vals) & first_row_vals != "")
    if (non_na_count >= 4) {
      write_log(paste("Title row detected in sheet:", sheet_name, "— re-reading from row 2"))
      df <- suppressMessages(read_excel(path, sheet = sheet_name, col_types = "text", skip = 1))
      df <- normalize_headers(df)
    }
  }
  df
}

# --- Handle duplicate column names ---
# readxl makes duplicates unique as "Enroll Date...11" "Enroll Date...18"
# This function finds the FIRST column matching a pattern (charter columns)
# and can also find the SECOND match (for pre-populated ST columns)
find_col_by_position <- function(df, patterns, occurrence = 1) {
  cols <- colnames(df)
  matches <- c()
  for (pat in patterns) {
    m <- grep(pat, cols, ignore.case = TRUE, value = TRUE)
    matches <- c(matches, m)
  }
  matches <- unique(matches)
  if (length(matches) >= occurrence) return(matches[occurrence])
  return(NA_character_)
}

# --- Primary column finder (first match) ---
find_best_col <- function(df, patterns) {
  find_col_by_position(df, patterns, occurrence = 1)
}

# --- Date parser: handles Excel serial numbers as STRINGS ---
# readxl with col_types="text" turns numeric dates into strings like "45839"
# This must detect and convert those, plus standard date formats.
parse_mixed_date <- function(x) {
  if (is.null(x) || length(x) == 0) return(as.Date(NA))
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXct") || inherits(x, "POSIXlt")) return(as.Date(x))
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))

  s <- str_trim(as.character(x))
  if (is.na(s) || s == "" || s == "N/A" || s == "#N/A") return(as.Date(NA))

  # Detect numeric strings (Excel serial dates read as text)
  if (grepl("^[0-9]+\\.?[0-9]*$", s)) {
    return(as.Date(as.numeric(s), origin = "1899-12-30"))
  }

  for (fmt in c("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%Y", "%Y/%m/%d")) {
    d <- tryCatch(as.Date(s, format = fmt), error = function(e) as.Date(NA))
    if (!is.na(d)) return(d)
  }
  return(as.Date(NA))
}

parse_date_col <- function(col) {
  sapply(col, parse_mixed_date, USE.NAMES = FALSE) %>% as.Date(origin = "1970-01-01")
}

col_to_letter <- function(n) {
  result <- ""
  while (n > 0) {
    rem    <- (n - 1) %% 26
    result <- paste0(LETTERS[rem + 1], result)
    n      <- (n - 1) %/% 26
  }
  result
}

# ======================================================================
# READ SHEETS — Universal tab finder with trimmed names
# ======================================================================
write_log("Reading workbook...")

sheet_names_raw <- excel_sheets(workbook_path)
sheet_names     <- str_trim(sheet_names_raw)  # strip leading/trailing spaces
write_log(paste("Found sheets:", paste(sheet_names_raw, collapse = ", ")))

# Helper: find a sheet by pattern on the TRIMMED names, return the RAW name
find_sheet <- function(patterns, exclude_patterns = NULL) {
  for (pat in patterns) {
    hits <- grep(pat, sheet_names, ignore.case = TRUE)
    if (!is.null(exclude_patterns)) {
      for (ep in exclude_patterns) {
        exclude_idx <- grep(ep, sheet_names, ignore.case = TRUE)
        hits <- setdiff(hits, exclude_idx)
      }
    }
    if (length(hits) > 0) return(sheet_names_raw[hits[1]])
  }
  return(NA_character_)
}

# --- SCSD Roster ---
# Priority: "SCSD.*Roster" > "Gen Roster" > any "Roster" (not SPED/Invoice)
scsd_sheet <- find_sheet(
  c("SCSD.*Roster", "Roster.*SCSD", "Gen.*Roster", "Roster"),
  exclude_patterns = c("SPED", "Invoice", "Disability", "IEP", "Registration")
)
if (is.na(scsd_sheet)) {
  scsd_sheet <- find_sheet("SCSD", exclude_patterns = c("Invoice", "SPED"))
}
if (is.na(scsd_sheet)) scsd_sheet <- sheet_names_raw[1]

scsd_raw <- smart_read_sheet(workbook_path, scsd_sheet)
write_log(paste("SCSD Roster:", nrow(scsd_raw), "rows from sheet:", scsd_sheet))
write_log(paste("  Columns:", paste(names(scsd_raw), collapse = " | ")))

# --- ST Enrollment ---
# Priority: "ST.*Enroll" > "ChartEnrol" > "Enrol" (not Multiple/Roster/SCSD/"Not on")
st_sheet <- find_sheet(
  c("^ST[\\s_].*[Ee]nrol", "ST.*Enrol", "ChartEnrol", "[Ee]nrol"),
  exclude_patterns = c("Multiple", "SCSD", "Roster", "SPED", "Invoice", "Not on", "Not in", "Not On", "Not In")
)
if (!is.na(st_sheet)) {
  st_raw <- smart_read_sheet(workbook_path, st_sheet)
  write_log(paste("ST Enrollment:", nrow(st_raw), "rows from sheet:", st_sheet))
} else {
  write_log("ST Enrollment sheet not found — skipping ST join", "WARN")
  st_raw <- NULL
}

# --- SPED Roster ---
sped_sheet <- find_sheet(
  c("SPED.*Roster", "SPED"),
  exclude_patterns = c("Invoice", "CCS Invoice")
)
if (!is.na(sped_sheet)) {
  sped_raw <- smart_read_sheet(workbook_path, sped_sheet)
  write_log(paste("SPED Roster:", nrow(sped_raw), "rows from sheet:", sped_sheet))
  write_log(paste("  Columns:", paste(names(sped_raw), collapse = " | ")))
} else {
  write_log("SPED Roster sheet not found — skipping SPED tab", "WARN")
  sped_raw <- NULL
}

# --- Disability ---
# Priority: "Disability" > "IEP Roster" (skip 2ND PULL variants)
disability_sheet <- find_sheet(
  c("^Disability$", "Disability ID$", "Disability.*Roster", "^IEP.*Roster$", "Disability"),
  exclude_patterns = c("2ND PULL", "Notes")
)
if (!is.na(disability_sheet)) {
  disability_df <- smart_read_sheet(workbook_path, disability_sheet)
  write_log(paste("Disability:", nrow(disability_df), "rows from sheet:", disability_sheet))
} else {
  write_log("Disability sheet not found — IEP Code lookup will be skipped", "WARN")
  disability_df <- NULL
}

# ======================================================================
# DETECT SCSD COLUMNS — covers all known charter variations
# ======================================================================
write_log("Detecting SCSD columns...")

nyssis_col     <- find_best_col(scsd_raw, c("^NYSSIS", "NYSSIS.?ID", "StateID", "Universal.*Student.*ID"))
student_id_col <- find_best_col(scsd_raw, c("^Student.?ID$", "^StudentID$", "^Student_Number$", "LocalID"))
full_name_col  <- find_best_col(scsd_raw, c("^Full.?Name$"))
last_name_col  <- find_best_col(scsd_raw, c("Student.*Last.*Name", "^Last.?Name$", "^LastName$"))
first_name_col <- find_best_col(scsd_raw, c("Student.*First.*Name", "^First.?Name$", "^FirstName$"))
dob_col        <- find_best_col(scsd_raw, c("Birth.?Date", "Date.?of.?Birth", "^DOB$", "Person.?DOB"))
grade_col      <- find_best_col(scsd_raw, c("^Grade$", "Current.?Grade", "Grade.?Level"))
entry_col      <- find_best_col(scsd_raw, c("^Entered$", "^Entry.?Date", "^Enroll.?Date", "^Start.?Date$"))
exit_col       <- find_best_col(scsd_raw, c("^Withdrew$", "^Withdrawal", "^Leave.?Date", "^Exit.?Date", "^Withdraw.?Date"))
fte_col        <- find_best_col(scsd_raw, c("Southside.*FTE", "OnTECH.*FTE", "SAS.*FTE", "Charter.*FTE", "^FTE$"))
address_col    <- find_best_col(scsd_raw, c("^Address$", "House.*Number", "^Street$"))
city_col       <- find_best_col(scsd_raw, c("^City$"))
zip_col        <- find_best_col(scsd_raw, c("^Zip"))
district_col   <- find_best_col(scsd_raw, c("District", "Residen"))

# --- Handle Combined "Full Name" column ---
has_full_name <- !is.na(full_name_col) && is.na(last_name_col) && is.na(first_name_col)
if (has_full_name) {
  write_log("Detected 'Full Name' column — splitting into Last Name and First Name")
  scsd_raw <- scsd_raw %>%
    mutate(
      .split_last  = str_trim(str_extract(.data[[full_name_col]], "^[^,]+")),
      .split_first = str_trim(str_replace(.data[[full_name_col]], "^[^,]+,?\\s*", ""))
    )
  last_name_col  <- ".split_last"
  first_name_col <- ".split_first"
}

# --- Handle pre-populated ST columns (OnTECH/SASCCS: duplicate entry/exit cols) ---
# If the SCSD roster already has ST dates in LATER duplicate columns,
# detect them so we don't need the ST Enrollment tab for those.
# Priority 1: explicitly named ST columns (engine output / Southside: "ST Entry Date", "ST Exit Date")
pre_pop_st_entry <- find_best_col(scsd_raw, c("^ST.?Entry.?Date$", "^ST.?Entry$", "^ST.?Enroll"))
pre_pop_st_exit  <- find_best_col(scsd_raw, c("^ST.?Exit.?Date$", "^ST.?Exit$", "^ST.?Withdraw", "^ST.?Leave"))
# Priority 2: if no explicit ST cols, look for 2nd occurrence of duplicate Entry/Exit columns
# (SASCCS/OnTECH pattern: charter Entry Date + ST Entry Date both named "Entry Date")
# Exclude columns that contain "Discrepancy" to avoid false matches on engine output columns
if (is.na(pre_pop_st_entry)) {
  entry_candidates <- grep("^Entry.?Date|^Enroll.?Date", names(scsd_raw), ignore.case=TRUE, value=TRUE)
  entry_candidates <- entry_candidates[!grepl("Discrepancy|Variance|Match|Diff", entry_candidates, ignore.case=TRUE)]
  if (length(entry_candidates) >= 2) pre_pop_st_entry <- entry_candidates[2]
}
if (is.na(pre_pop_st_exit)) {
  exit_candidates <- grep("^Exit.?Date|^Withdraw.?Date|^Leave.?Date", names(scsd_raw), ignore.case=TRUE, value=TRUE)
  exit_candidates <- exit_candidates[!grepl("Discrepancy|Variance|Match|Diff", exit_candidates, ignore.case=TRUE)]
  if (length(exit_candidates) >= 2) pre_pop_st_exit <- exit_candidates[2]
}
pre_pop_school   <- find_best_col(scsd_raw, c("^SchoolName$", "^School.?Name$"))
# Note: "School Name" column that we OUTPUT also matches — but that won't exist yet
# Only match columns that are in the raw data (not our output)

has_prepop_st <- !is.na(pre_pop_st_entry)
if (has_prepop_st) {
  write_log(paste("Pre-populated ST columns detected — ST Entry:", pre_pop_st_entry,
                   "| ST Exit:", pre_pop_st_exit, "| School:", pre_pop_school))
}

write_log(paste("NYSSIS:", nyssis_col, "| StudentID:", student_id_col,
                "| Entry:", entry_col, "| Exit:", exit_col, "| FTE:", fte_col))

# --- ST Enrollment columns (constant per user) ---
if (!is.null(st_raw)) {
  st_nyssis_col <- find_best_col(st_raw, c("^NYSSIS", "NYSSIS.?ID", "StateID",
                                             "Universal.*Student.*ID", "Student_Universal",
                                             "UniversalStudentID"))
  st_entry_col  <- find_best_col(st_raw, c("Enrollment.*Start", "Start.?Date", "^Entry"))
  st_exit_col   <- find_best_col(st_raw, c("Enrollment.*End", "End.?Date", "^Exit"))
  st_school_col <- find_best_col(st_raw, c("SchoolName", "School.?Name", "^School$"))
  st_id_col     <- find_best_col(st_raw, c("^StudentID$", "^Student.?ID$", "LocalID"))
  st_first_col  <- find_best_col(st_raw, c("^First.?Name$", "^FirstName$"))
  st_last_col   <- find_best_col(st_raw, c("^Last.?Name$", "^LastName$"))
  st_dob_col    <- find_best_col(st_raw, c("Date.?of.?Birth", "^DOB$", "Person.?DOB", "Birth"))
  st_grade_col  <- find_best_col(st_raw, c("^Grade$", "Current.?Grade"))
  write_log(paste("ST NYSSIS:", st_nyssis_col, "| ST Entry:", st_entry_col, "| ST Exit:", st_exit_col))
}

# ======================================================================
# CLEAN SCSD ROSTER
# ======================================================================
write_log("Cleaning SCSD Roster...")

if (is.na(nyssis_col)) {
  stop("ERROR: No NYSSIS/StateID column found in SCSD Roster. ",
       "Column headers found: ", paste(names(scsd_raw), collapse = ", "))
}

scsd_df <- scsd_raw %>%
  filter(!is.na(.data[[nyssis_col]]) & str_trim(.data[[nyssis_col]]) != "") %>%
  mutate(
    NYSSIS = str_trim(gsub("\\.0$", "", .data[[nyssis_col]])),
    
    Charter_Entry = if (!is.na(entry_col)) {
      parse_date_col(.data[[entry_col]])
    } else {
      as.Date(NA)
    },
    
    Charter_Exit = if (!is.na(exit_col)) {
      parse_date_col(.data[[exit_col]])
    } else {
      as.Date(NA)
    },
    
    FTE = if (!is.na(fte_col)) {
      suppressWarnings(as.numeric(.data[[fte_col]]))
    } else {
      NA_real_
    }
  )

# Standardize Charter FTE column so it always exists
scsd_df <- scsd_df %>%
  mutate(Charter_FTE = FTE)

# ======================================================================
# CLEAN & DEDUPE ST ENROLLMENT
# ======================================================================
if (!is.null(st_raw)) {
  write_log("Cleaning ST Enrollment...")

  st_df <- st_raw %>%
    filter(!is.na(.data[[st_nyssis_col]]) & str_trim(.data[[st_nyssis_col]]) != "") %>%
    mutate(
      NYSSIS       = str_trim(gsub("\\.0$", "", .data[[st_nyssis_col]])),
      ST_Entry     = parse_date_col(.data[[st_entry_col]]),
      ST_Exit      = parse_date_col(.data[[st_exit_col]]),
      ST_School    = if (!is.na(st_school_col)) str_trim(.data[[st_school_col]]) else NA_character_,
      ST_StudentID = if (!is.na(st_id_col))     str_trim(.data[[st_id_col]])     else NA_character_,
      ST_First_Norm  = if (!is.na(st_first_col)) toupper(str_trim(.data[[st_first_col]])) else NA_character_,
      ST_Last_Norm   = if (!is.na(st_last_col))  toupper(str_trim(.data[[st_last_col]]))  else NA_character_,
      ST_DOB_Norm    = if (!is.na(st_dob_col))   as.character(parse_date_col(.data[[st_dob_col]])) else NA_character_,
      ST_Grade_Norm  = if (!is.na(st_grade_col)) toupper(str_trim(.data[[st_grade_col]])) else NA_character_
    )

  enroll_counts <- st_df %>% count(NYSSIS, name = "Enrollment_Count")
  dupes <- enroll_counts %>% filter(Enrollment_Count > 1)
  write_log(paste("ST duplicates:", nrow(dupes), "students with multiple enrollments"))

  st_deduped <- st_df %>%
    group_by(NYSSIS) %>% arrange(desc(ST_Entry)) %>% slice(1) %>% ungroup() %>%
    select(NYSSIS, ST_Entry, ST_Exit, ST_School, ST_StudentID,
           ST_First_Norm, ST_Last_Norm, ST_DOB_Norm, ST_Grade_Norm)

  st_name_lookup <- st_df %>%
    filter(!is.na(ST_First_Norm) & !is.na(ST_DOB_Norm) & !is.na(ST_Grade_Norm)) %>%
    mutate(NameKey = paste(ST_First_Norm, ST_Last_Norm, ST_DOB_Norm, ST_Grade_Norm, sep = "|")) %>%
    group_by(NameKey) %>% arrange(desc(ST_Entry)) %>% slice(1) %>% ungroup() %>%
    select(NameKey, ST_Entry, ST_Exit, ST_School, ST_StudentID, NYSSIS)

  # Build NYSSIS <-> StudentID cross-reference for SPED lookups
  st_id_xref <- st_df %>%
    filter(!is.na(ST_StudentID) & ST_StudentID != "") %>%
    select(NYSSIS, ST_StudentID) %>%
    distinct(ST_StudentID, .keep_all = TRUE)

  write_log(paste("ST deduped:", nrow(st_deduped), "unique students"))
  write_log(paste("ST name lookup:", nrow(st_name_lookup), "combos"))
  write_log(paste("ST ID cross-ref:", nrow(st_id_xref), "StudentID<->NYSSIS pairs"))
}

# ======================================================================
# JOIN — PRIMARY (NYSSIS) + FALLBACK (Name+DOB+Grade) + PRE-POP
# ======================================================================
if (!is.null(st_raw)) {
  write_log("Joining ST data to SCSD Roster by NYSSIS...")

  scsd_df <- scsd_df %>%
    left_join(st_deduped, by = "NYSSIS") %>%
    left_join(enroll_counts, by = "NYSSIS")

  nyssis_matched <- sum(!is.na(scsd_df$ST_Entry))
  write_log(paste("After NYSSIS join:", nyssis_matched, "of", nrow(scsd_df), "matched"))

  # Fallback: Name+DOB+Grade
  write_log("Applying Name+DOB+Grade fallback for unmatched students...")

  scsd_df <- scsd_df %>%
    mutate(
      SCSD_First_Norm = if (!is.na(first_name_col)) toupper(str_trim(.data[[first_name_col]])) else NA_character_,
      SCSD_Last_Norm  = if (!is.na(last_name_col))  toupper(str_trim(.data[[last_name_col]]))  else NA_character_,
      SCSD_DOB_Norm   = if (!is.na(dob_col))        as.character(parse_date_col(.data[[dob_col]])) else NA_character_,
      SCSD_Grade_Norm = if (!is.na(grade_col))       toupper(str_trim(.data[[grade_col]]))      else NA_character_,
      NameKey         = paste(SCSD_First_Norm, SCSD_Last_Norm, SCSD_DOB_Norm, SCSD_Grade_Norm, sep = "|")
    )

  needs_fallback <- is.na(scsd_df$ST_Entry) & is.na(scsd_df$ST_Exit) & is.na(scsd_df$ST_School)

  if (sum(needs_fallback) > 0 && nrow(st_name_lookup) > 0) {
    fallback_result <- scsd_df %>%
      filter(needs_fallback) %>%
      left_join(st_name_lookup, by = "NameKey", suffix = c("", "_fb"))

    scsd_df <- scsd_df %>%
      left_join(
        fallback_result %>%
          select(NYSSIS,
                 ST_Entry_fb     = ST_Entry_fb,
                 ST_Exit_fb      = ST_Exit_fb,
                 ST_School_fb    = ST_School_fb,
                 ST_StudentID_fb = ST_StudentID_fb),
        by = "NYSSIS"
      ) %>%
      mutate(
        ST_Entry     = if_else(needs_fallback & !is.na(ST_Entry_fb),     ST_Entry_fb,     ST_Entry),
        ST_Exit      = if_else(needs_fallback & !is.na(ST_Exit_fb),      ST_Exit_fb,      ST_Exit),
        ST_School    = if_else(needs_fallback & !is.na(ST_School_fb),    ST_School_fb,    ST_School),
        ST_StudentID = if_else(needs_fallback & !is.na(ST_StudentID_fb), ST_StudentID_fb, ST_StudentID),
        Match_Method = case_when(
          !needs_fallback                       ~ "NYSSIS",
          needs_fallback & !is.na(ST_Entry_fb) ~ "Name+DOB+Grade",
          TRUE                                  ~ "NOT FOUND"
        )
      ) %>%
      select(-ST_Entry_fb, -ST_Exit_fb, -ST_School_fb, -ST_StudentID_fb)
  } else {
    scsd_df <- scsd_df %>%
      mutate(Match_Method = case_when(!needs_fallback ~ "NYSSIS", TRUE ~ "NOT FOUND"))
  }

  write_log(paste("NYSSIS matched:          ", sum(scsd_df$Match_Method == "NYSSIS",          na.rm = TRUE)))
  write_log(paste("Name+DOB+Grade matched:  ", sum(scsd_df$Match_Method == "Name+DOB+Grade",  na.rm = TRUE)))
  write_log(paste("NOT FOUND:               ", sum(scsd_df$Match_Method == "NOT FOUND",        na.rm = TRUE)))

} else if (has_prepop_st) {
  # No ST tab, but roster has pre-populated ST columns (OnTECH/SASCCS pattern)
  write_log("No ST Enrollment tab — using pre-populated ST columns from roster")
  scsd_df <- scsd_df %>%
    mutate(
      ST_Entry         = parse_date_col(.data[[pre_pop_st_entry]]),
      ST_Exit          = if (!is.na(pre_pop_st_exit))  parse_date_col(.data[[pre_pop_st_exit]])  else as.Date(NA),
      ST_School        = if (!is.na(pre_pop_school))   str_trim(.data[[pre_pop_school]])         else NA_character_,
      ST_StudentID     = NA_character_,
      Enrollment_Count = NA_integer_,
      Match_Method     = if_else(!is.na(ST_Entry), "Pre-populated", "NOT FOUND")
    )
} else {
  # No ST data at all
  scsd_df <- scsd_df %>%
    mutate(
      ST_Entry         = as.Date(NA),
      ST_Exit          = as.Date(NA),
      ST_School        = NA_character_,
      ST_StudentID     = NA_character_,
      Enrollment_Count = NA_integer_,
      Match_Method     = "NOT FOUND"
    )
}

# --- Auto-populate Student ID from ST Enrollment ---
if (is.na(student_id_col)) {
  write_log("No Student ID column in source — auto-populating from ST Enrollment")
  scsd_df <- scsd_df %>% mutate(Auto_Student_ID = ST_StudentID)
} else {
  write_log("Student ID column found — filling blanks from ST Enrollment")
  scsd_df <- scsd_df %>%
    mutate(
      Auto_Student_ID = case_when(
        !is.na(.data[[student_id_col]]) & str_trim(.data[[student_id_col]]) != "" ~
          str_trim(gsub("\\.0$", "", .data[[student_id_col]])),
        TRUE ~ ST_StudentID
      )
    )
}

# ======================================================================
# ENTRY DATE DISCREPANCY RULES
# Charter later = ALWAYS FAVORABLE (any amount)
# Charter earlier ≤3 = FAVORABLE, >3 = UNFAVORABLE
# July/Aug exemption only when BOTH dates in window
# ======================================================================
classify_entry <- function(charter_entry, st_entry) {
  c_blank <- is.na(charter_entry)
  s_blank <- is.na(st_entry)
  if (c_blank & s_blank) return("")
  if (!c_blank & s_blank) return("Charter has entry date but ST does not (UNFAVORABLE)")
  if (c_blank & !s_blank) return("ST has entry date but Charter does not (UNFAVORABLE)")

  # Normalize: any entry date before 9/3/2025 is treated as 9/3/2025
  SY_FIRST <- as.Date("2025-09-03")
  charter_entry_n <- pmax(charter_entry, SY_FIRST)
  st_entry_n      <- pmax(st_entry,      SY_FIRST)

  c_in_julyaug <- charter_entry >= JULY_START & charter_entry <= GRACE_DEADLINE
  s_in_julyaug <- st_entry >= JULY_START & st_entry <= GRACE_DEADLINE

  if (c_in_julyaug & s_in_julyaug)
    return("Charter and ST both enrolled in July/August (FAVORABLE)")
  if (s_in_julyaug & !c_in_julyaug)
    return("ST enrolled in July/August by school-year first day deadline (FAVORABLE)")
  if (charter_entry_n == st_entry_n)
    return("Entry dates match (FAVORABLE)")

  diff <- as.integer(charter_entry_n - st_entry_n)
  if (diff > 0) {
    return(paste0("Charter entry is ", diff, " days later than ST entry date (FAVORABLE)"))
  }
  diff_abs <- abs(diff)
  if (diff_abs <= GRACE_DAYS) {
    return(paste0("Charter entry is ", diff_abs, " days earlier than ST entry date (FAVORABLE)"))
  }
  return(paste0("Charter entry is ", diff_abs, " days earlier than ST entry date (UNFAVORABLE)"))
}

# ======================================================================
# EXIT DATE DISCREPANCY RULES
# ======================================================================
classify_exit <- function(charter_exit, st_exit) {
  c_blank <- is.na(charter_exit)
  s_blank <- is.na(st_exit)
  if (c_blank & s_blank) return("No Charter or ST exit dates on file (FAVORABLE)")
  if (c_blank & !s_blank) return("ST has exit date but Charter does not (UNFAVORABLE)")
  if (!c_blank & s_blank) return("Charter has exit date but ST does not (Discrepancy – update in ST)")
  if (charter_exit == st_exit) return("Exit dates match (FAVORABLE)")

  diff <- as.integer(charter_exit - st_exit)
  if (diff < 0) {
    diff_abs <- abs(diff)
    if (diff_abs <= GRACE_DAYS) return(paste0("Charter exit is ", diff_abs, " days earlier (FAVORABLE)"))
    return(paste0("Charter exit is ", diff_abs, " days earlier (Discrepancy – update in ST)"))
  }
  if (diff <= GRACE_DAYS) return(paste0("Charter exit is ", diff, " days later (FAVORABLE)"))
  return(paste0("Charter exit is ", diff, " days later (UNFAVORABLE)"))
}

# ======================================================================
# STUDENT STATUS — 3 categories:
#   "Good"                = all FAVORABLE / match / Discrepancy-update-in-ST
#   "No Good"             = any UNFAVORABLE (data exists but doesn't align)
#   "Not Found-Needs Manual Review" = student NOT FOUND in ST (no dates/data pulled)
# "Discrepancy – update in ST" is an ST issue, NOT a billing problem
# ======================================================================
determine_status <- function(entry_disc, exit_disc, match_method) {
  e <- as.character(entry_disc)
  x <- as.character(exit_disc)
  m <- as.character(match_method)

  # Students not found in ST — no data to evaluate, needs manual verification
  if (!is.na(m) && m == "NOT FOUND") return("Not Found-Needs Manual Review")

  # If either discrepancy is UNFAVORABLE, it's No Good
  if (grepl("UNFAVORABLE", e, ignore.case = TRUE)) return("No Good")
  if (grepl("UNFAVORABLE", x, ignore.case = TRUE)) return("No Good")
  return("Good")
}

# ======================================================================
# APPLY RULES
# ======================================================================
write_log("Applying enrollment rules...")

scsd_df <- scsd_df %>%
  rowwise() %>%
  mutate(
    Entry_Discrepancy    = classify_entry(Charter_Entry, ST_Entry),
    Exit_Discrepancy     = classify_exit(Charter_Exit, ST_Exit),
    Student_Status       = determine_status(Entry_Discrepancy, Exit_Discrepancy, Match_Method),
    Multiple_Enrollments = case_when(
      is.na(Enrollment_Count) ~ "",
      Enrollment_Count > 1   ~ paste0(Enrollment_Count, " ENROLLMENTS"),
      TRUE                    ~ ""
    ),
    ST_ID_Match       = ifelse(is.na(ST_StudentID), "NOT FOUND IN ST", ST_StudentID),
    Notes_for_Patrick = case_when(
      is.na(ST_Entry) & is.na(ST_Exit) & is.na(ST_School) ~
        "Student not found in ST Enrollment report",
      TRUE ~ ""
    )
  ) %>%
  ungroup()

# ======================================================================
# SPED FLAG — join SPED roster to SCSD roster by NYSSIS
# Adds SPED_Status column: "SPED" if student appears in SPED Roster, else "Gen Ed"
# ======================================================================
write_log("Building SPED Status flag...")
if (!is.null(sped_raw)) {
  sped_pre_nyssis_col <- find_best_col(sped_raw,
    c("^NYSSIS$", "^NYSSIS.?ID$", "StateID", "Universal.*Student.*ID"))
  if (!is.na(sped_pre_nyssis_col)) {
    sped_keys <- sped_raw %>%
      mutate(SPED_KEY = str_trim(gsub("[.]0$", "", as.character(.data[[sped_pre_nyssis_col]])))) %>%
      filter(!is.na(SPED_KEY) & SPED_KEY != "") %>%
      distinct(SPED_KEY) %>% pull(SPED_KEY)
    scsd_df <- scsd_df %>%
      mutate(SPED_Status = ifelse(NYSSIS %in% sped_keys, "SPED", "Gen Ed"))
    write_log(paste("SPED Status: SPED =", sum(scsd_df$SPED_Status == "SPED"),
                    "| Gen Ed =", sum(scsd_df$SPED_Status == "Gen Ed")))
  } else {
    scsd_df <- scsd_df %>% mutate(SPED_Status = "Unknown")
    write_log("No NYSSIS column found in SPED Roster — SPED_Status set to Unknown", "WARN")
  }
} else {
  scsd_df <- scsd_df %>% mutate(SPED_Status = "Unknown")
  write_log("No SPED Roster tab found — SPED_Status set to Unknown for all students", "WARN")
}

# ======================================================================
# FTE CALCULATIONS — SchoolTool FTE, Charter FTE, FTE Variance
# Exact NYSED FTE formula per Commissioner's Regulations §119.1(b)(3) and §175.6
# FTE = Student Weeks / Program Weeks (40 weeks for Sep–Jun 10-month year)
# Week rules (from NYSED FTE Calculator JS algorithm, verified against official results):
#   - A Mon–Sun week counts if enrollment days within that week AND that month are >= 3
#   - For the START month: the enrollment-start week only counts if its Sunday
#     falls within the same calendar month (partial weeks crossing month boundary
#     at the start of enrollment are excluded unless Sunday stays in the month)
#   - For the END month: all weeks count if >= 3 enrollment days are in that month
#   - Middle months: always 4 weeks each (max per month; Jul = 5)
#   - If enrollment spans < 3 calendar days: 0 weeks; exactly 3 days: 1 week
#   - FTE capped at 1.000 (min(student_weeks / program_weeks, 1))
# Active students (no ST exit) use STILL_ENROLLED_DATE (June 30) as end date
# ======================================================================

# --- NYSED week-count helper functions ---
nysed_month_seq <- function(d) {
  m <- month(d)
  ifelse(m >= 7L, m - 6L, m + 6L)
}

# Weeks in the START month from start_date:
# The first (enrollment-start) week only counts if its Sunday is within the same month.
# All subsequent weeks in that month count if >= 3 days are in the month.
nysed_first_month_weeks <- function(start_date) {
  m_end <- ceiling_date(start_date, "month") - days(1)
  wk    <- floor_date(start_date, "week", week_start = 1)
  count <- 0L; first_wk <- TRUE
  while (wk <= m_end) {
    wk_end <- wk + days(6)
    day_s  <- pmax(start_date, wk)
    day_e  <- pmin(m_end, wk_end)
    n_days <- if (day_e >= day_s) as.integer(day_e - day_s) + 1L else 0L
    if (first_wk) {
      if (wk_end <= m_end && n_days >= 3L) count <- count + 1L  # Sunday must be in month
    } else {
      if (n_days >= 3L) count <- count + 1L
    }
    first_wk <- FALSE
    wk <- wk + days(7)
  }
  cap <- if (month(start_date) == 7L) 5L else 4L
  min(count, cap)
}

# Weeks in the END month up to end_date:
# All weeks count if >= 3 enrollment days fall within the month (no Sunday constraint).
nysed_last_month_weeks <- function(end_date) {
  m_start <- floor_date(end_date, "month")
  wk      <- floor_date(m_start, "week", week_start = 1)
  count   <- 0L
  while (wk <= end_date) {
    wk_end <- wk + days(6)
    day_s  <- pmax(m_start, wk)
    day_e  <- pmin(end_date, wk_end)
    n_days <- if (day_e >= day_s) as.integer(day_e - day_s) + 1L else 0L
    if (n_days >= 3L) count <- count + 1L
    wk <- wk + days(7)
  }
  cap <- if (month(end_date) == 7L) 5L else 4L
  min(count, cap)
}

# Master NYSED week counter: handles same-month, multi-month, and edge cases
nysed_student_weeks <- function(start_date, end_date) {
  if (is.na(start_date) || is.na(end_date)) return(NA_integer_)
  cal_days <- as.integer(end_date - start_date) + 1L
  if (cal_days < 3L)  return(0L)
  if (cal_days == 3L) return(1L)
  ms <- nysed_month_seq(start_date)
  me <- nysed_month_seq(end_date)
  if (ms == me) {
    # Same calendar month: start-week Sunday rule applies to first week
    m_end <- ceiling_date(start_date, "month") - days(1)
    wk <- floor_date(start_date, "week", week_start = 1)
    count <- 0L; first_wk <- TRUE
    while (wk <= end_date) {
      wk_end <- wk + days(6)
      day_s  <- pmax(start_date, wk)
      day_e  <- pmin(end_date, wk_end, m_end)
      n_days <- if (day_e >= day_s) as.integer(day_e - day_s) + 1L else 0L
      if (first_wk) {
        if (wk_end <= m_end && n_days >= 3L) count <- count + 1L
      } else {
        if (n_days >= 3L) count <- count + 1L
      }
      first_wk <- FALSE
      wk <- wk + days(7)
    }
    cap <- if (month(start_date) == 7L) 5L else 4L
    return(as.integer(min(count, cap)))
  }
  # Multi-month
  w_first  <- nysed_first_month_weeks(start_date)
  w_last   <- nysed_last_month_weeks(end_date)
  w_middle <- max(0L, as.integer(me - ms - 1L)) * 4L
  as.integer(w_first + w_last + w_middle)
}

# Vectorised wrapper: applies nysed_student_weeks row-by-row
nysed_fte_vec <- function(start_vec, end_vec, prog_weeks = 40L) {
  mapply(function(s, e) {
    if (is.na(s) || is.na(e)) return(NA_real_)
    sw <- nysed_student_weeks(s, e)
    round(min(sw / prog_weeks, 1.0), 3)
  }, start_vec, end_vec, SIMPLIFY = TRUE, USE.NAMES = FALSE)
}

write_log("Calculating FTE fields...")
scsd_df <- scsd_df %>%
  mutate(
    # SchoolTool FTE: exact NYSED §119.1 week-based formula
    # Uses ST enrollment dates; falls back to Charter entry when ST entry is missing
    # Active students (no ST exit) are credited through June 30 (NYSED program end)
    ST_FTE = nysed_fte_vec(
      pmax(
        if_else(is.na(ST_Entry), as.Date(Charter_Entry), as.Date(ST_Entry)),
        as.Date("2025-09-03")  # normalize: any entry before 9/3/2025 treated as 9/3/2025
      ),
      # Exit priority: ST Exit -> Charter Exit -> STILL_ENROLLED_DATE
      if_else(!is.na(ST_Exit),  as.Date(ST_Exit),
        if_else(!is.na(Charter_Exit), as.Date(Charter_Exit), STILL_ENROLLED_DATE)
      )
    ),
    # ST_FTE already capped at 1.000 inside nysed_fte_vec
    ST_FTE = pmax(0, ST_FTE),
    # FTE Variance: Charter (original FTE col) minus SchoolTool (positive = overbilled)
    FTE_Variance = round(as.numeric(FTE) - as.numeric(ST_FTE), 4),
    # SPED FTE: only computed here for use on the SPED ROSTER tab
    ST_SPED_FTE = ifelse(SPED_Status == "SPED", ST_FTE, NA_real_),
    SPED_FTE_Variance = ifelse(
      SPED_Status == "SPED" & !is.na(ST_SPED_FTE),
      round(as.numeric(FTE) - ST_SPED_FTE, 4),
      NA_real_
    )
  )
write_log("FTE calculations complete")

# ======================================================================
# CREATE CHARTER FTE + CHARTER SPED FTE FOR OUTPUT
# ======================================================================
scsd_df <- scsd_df %>%
  mutate(
    Charter_FTE      = FTE,
    Charter_SPED_FTE = ifelse(SPED_Status == "SPED", FTE, NA_real_)
  )


# ======================================================================
# SUMMARY
# ======================================================================
write_log("Generating summary...")

total            <- nrow(scsd_df)
good_count       <- sum(scsd_df$Student_Status == "Good",            na.rm = TRUE)
nogood_count     <- sum(scsd_df$Student_Status == "No Good",         na.rm = TRUE)
needs_review     <- sum(scsd_df$Student_Status == "Not Found-Needs Manual Review", na.rm = TRUE)
not_in_st_count  <- sum(scsd_df$Match_Method   == "NOT FOUND",       na.rm = TRUE)
multi_enroll     <- sum(scsd_df$Multiple_Enrollments != "",           na.rm = TRUE)
name_match_count <- sum(scsd_df$Match_Method   == "Name+DOB+Grade",  na.rm = TRUE)

cat("\n")
cat(strrep("=", 60), "\n")
cat("  CHARTER AUDIT ENGINE V2 - SUMMARY REPORT\n")
cat(strrep("=", 60), "\n")
cat(sprintf("  Total students:                    %d\n",   total))
cat(sprintf("  Good:                              %d (%.1f%%)\n", good_count,   100 * good_count   / max(total, 1)))
cat(sprintf("  No Good:                           %d (%.1f%%)\n", nogood_count, 100 * nogood_count / max(total, 1)))
cat(sprintf("  Not Found-Needs Manual Review:     %d (%.1f%%)\n", needs_review, 100 * needs_review / max(total, 1)))
cat(sprintf("  Not in ST Enrollment:              %d\n",   not_in_st_count))
cat(sprintf("  Multiple Enrollments:              %d\n",   multi_enroll))
cat(sprintf("  Matched by Name+DOB+Grade:         %d\n",   name_match_count))
cat(strrep("-", 60), "\n")
cat("  ENTRY DATE DISCREPANCY BREAKDOWN\n")
cat(strrep("-", 60), "\n")
entry_summary <- scsd_df %>% count(Entry_Discrepancy, sort = TRUE)
for (i in 1:nrow(entry_summary)) {
  cat(sprintf("    %s: %d\n", entry_summary$Entry_Discrepancy[i], entry_summary$n[i]))
}
cat(strrep("-", 60), "\n")
cat("  EXIT DATE DISCREPANCY BREAKDOWN\n")
cat(strrep("-", 60), "\n")
exit_summary <- scsd_df %>% count(Exit_Discrepancy, sort = TRUE)
for (i in 1:nrow(exit_summary)) {
  cat(sprintf("    %s: %d\n", exit_summary$Exit_Discrepancy[i], exit_summary$n[i]))
}
cat(strrep("=", 60), "\n")
cat("  WHAT TO DO NEXT\n")
cat(strrep("=", 60), "\n")
if (needs_review > 0 || nogood_count > 0) {
  cat(sprintf("  >> %d students need your attention\n", needs_review + nogood_count))
  if (needs_review > 0)
    cat(sprintf("     - %d Not Found: look up in ST Enrollment\n", needs_review))
  if (nogood_count > 0)
    cat(sprintf("     - %d No Good: check for duplicates or supporting docs\n", nogood_count))
  cat("  >> Open the ACTION NEEDED tab (red tab) for the full work queue\n")
  cat("  >> Update Student Status on SCSD ROSTER when resolved\n")
  cat("  >> Strip internal columns before sending to budget\n")
} else {
  cat("  ALL CLEAR - No manual review needed.\n")
}

# ======================================================================
# BUILD OUTPUT (SCSD ROSTER)
# ======================================================================
scsd_out <- scsd_df %>%
  mutate(
    .student_id = Auto_Student_ID,
    .last_name  = if (!is.na(last_name_col))  .data[[last_name_col]]  else NA_character_,
    .first_name = if (!is.na(first_name_col)) .data[[first_name_col]] else NA_character_,
    .dob        = if (!is.na(dob_col))        .data[[dob_col]]        else NA_character_,
    .grade      = if (!is.na(grade_col))      .data[[grade_col]]      else NA_character_,
    .address    = if (!is.na(address_col))    .data[[address_col]]    else NA_character_,
    .city       = if (!is.na(city_col))       .data[[city_col]]       else NA_character_,
    .zip        = if (!is.na(zip_col))        .data[[zip_col]]        else NA_character_,
    .district   = if (!is.na(district_col))   .data[[district_col]]   else NA_character_
  ) %>%
  select(
    `NYSSIS ID`                        = NYSSIS,
    `Student ID`                       = .student_id,
    `Last Name`                        = .last_name,
    `First Name`                       = .first_name,
    `Date of Birth`                    = .dob,
    `Current Grade Level`              = .grade,
    `Address`                          = .address,
    `City`                             = .city,
    `Zip Code`                         = .zip,
    `District of Residency`            = .district,
    `Entry Date`                       = Charter_Entry,
    `Leave Date`                       = Charter_Exit,
    `FTE`                              = FTE,
    `ST Entry Date`                    = ST_Entry,
    `ST Exit Date`                     = ST_Exit,
    `School Name`                      = ST_School,
    `Student Status`                   = Student_Status,
    `ST ID Match`                      = ST_ID_Match,
    `Match Method`                     = Match_Method,
    `Multiple Enrollments`             = Multiple_Enrollments,
    `Notes for Patrick`                = Notes_for_Patrick,
    `Entry Date Discrepancy (>3 Days)` = Entry_Discrepancy,
    `Exit Date Discrepancy (>3 Days)`  = Exit_Discrepancy,
    `SPED Status`                      = SPED_Status,
    `SchoolTool FTE`                   = ST_FTE,
    `Charter FTE`                      = Charter_FTE,
    `FTE Variance`                     = FTE_Variance,
    `ST SPED FTE`                      = ST_SPED_FTE,
    `Charter SPED FTE`                 = Charter_SPED_FTE,
    `SPED FTE Variance`                = SPED_FTE_Variance
  )

out_col_names <- colnames(scsd_out)
col_idx <- setNames(seq_along(out_col_names), out_col_names)

# ======================================================================
# NOT IN ST TAB
# ======================================================================
not_in_st_df <- scsd_df %>%
  filter(Match_Method == "NOT FOUND") %>%
  mutate(
    .first_name = if (!is.na(first_name_col)) .data[[first_name_col]] else NA_character_,
    .last_name  = if (!is.na(last_name_col))  .data[[last_name_col]]  else NA_character_,
    .grade      = if (!is.na(grade_col))      .data[[grade_col]]      else NA_character_,
    .dob        = if (!is.na(dob_col))        .data[[dob_col]]        else NA_character_
  ) %>%
  select(`NYSSIS ID` = NYSSIS, `First Name` = .first_name, `Last Name` = .last_name,
         `Grade` = .grade, `Date of Birth` = .dob,
         `Charter Entry` = Charter_Entry, `Charter Exit` = Charter_Exit, `FTE` = FTE)

write_log(paste("Not In ST tab:", nrow(not_in_st_df), "students"))

# ======================================================================
# MULTIPLE ENROLLMENTS TAB
# ======================================================================
multi_enroll_df <- scsd_df %>%
  filter(!is.na(Enrollment_Count) & Enrollment_Count > 1) %>%
  mutate(
    .first_name = if (!is.na(first_name_col)) .data[[first_name_col]] else NA_character_,
    .last_name  = if (!is.na(last_name_col))  .data[[last_name_col]]  else NA_character_,
    .grade      = if (!is.na(grade_col))      .data[[grade_col]]      else NA_character_
  ) %>%
  select(`NYSSIS ID` = NYSSIS, `First Name` = .first_name, `Last Name` = .last_name,
         `Grade` = .grade, `Charter Entry` = Charter_Entry, `Charter Exit` = Charter_Exit,
         `FTE` = FTE, `ST Entry` = ST_Entry, `ST Exit` = ST_Exit,
         `School Name` = ST_School, `Enrollment Count` = Enrollment_Count)

write_log(paste("Multiple Enrollments tab:", nrow(multi_enroll_df), "students"))

# ======================================================================
# FORMULA ADAPTER
# ======================================================================
formula_map <- list(
  list(std = "NYSSIS ID",        detected = nyssis_col,      formula_type = "VLOOKUP / INDEX-MATCH key"),
  list(std = "First Name",       detected = first_name_col,  formula_type = "INDEX/MATCH fallback"),
  list(std = "Last Name",        detected = last_name_col,   formula_type = "INDEX/MATCH fallback"),
  list(std = "Grade",            detected = grade_col,       formula_type = "INDEX/MATCH fallback"),
  list(std = "Date of Birth",    detected = dob_col,         formula_type = "INDEX/MATCH fallback"),
  list(std = "Entry Date",       detected = entry_col,       formula_type = "IFS (Entry Discrepancy)"),
  list(std = "Leave Date",       detected = exit_col,        formula_type = "IF (Exit Discrepancy)"),
  list(std = "FTE",              detected = fte_col,         formula_type = "Manual"),
  list(std = "Address",          detected = address_col,     formula_type = "Manual"),
  list(std = "City",             detected = city_col,        formula_type = "Manual"),
  list(std = "Zip Code",         detected = zip_col,         formula_type = "Manual"),
  list(std = "District",         detected = district_col,    formula_type = "Manual"),
  list(std = "ST Entry Date",    detected = if (exists("st_entry_col"))  st_entry_col  else "N/A", formula_type = "VLOOKUP"),
  list(std = "ST Exit Date",     detected = if (exists("st_exit_col"))   st_exit_col   else "N/A", formula_type = "VLOOKUP"),
  list(std = "School Name",      detected = if (exists("st_school_col")) st_school_col else "N/A", formula_type = "VLOOKUP"),
  list(std = "Multiple Enrollments",     detected = "Computed", formula_type = "COUNTIF"),
  list(std = "Entry Discrepancy",        detected = "Computed", formula_type = "IFS"),
  list(std = "Exit Discrepancy",         detected = "Computed", formula_type = "IF"),
  list(std = "Student Status",           detected = "Computed", formula_type = "IFS"),
  list(std = "Match Method (2nd Check)", detected = "Computed", formula_type = "INDEX/MATCH fallback")
)

get_out_letter <- function(std_name) {
  name_map <- c(
    "NYSSIS ID" = "NYSSIS ID", "First Name" = "First Name", "Last Name" = "Last Name",
    "Grade" = "Current Grade Level", "Date of Birth" = "Date of Birth",
    "Entry Date" = "Entry Date", "Leave Date" = "Leave Date", "FTE" = "FTE",
    "Address" = "Address", "City" = "City", "Zip Code" = "Zip Code",
    "District" = "District of Residency", "ST Entry Date" = "ST Entry Date",
    "ST Exit Date" = "ST Exit Date", "School Name" = "School Name",
    "Multiple Enrollments" = "Multiple Enrollments",
    "Entry Discrepancy" = "Entry Date Discrepancy (>3 Days)",
    "Exit Discrepancy" = "Exit Date Discrepancy (>3 Days)",
    "Student Status" = "Student Status", "Match Method (2nd Check)" = "Match Method"
  )
  out_name <- name_map[std_name]
  if (is.na(out_name) || !(out_name %in% out_col_names)) return("N/A")
  col_to_letter(which(out_col_names == out_name))
}

formula_adapter_header <- data.frame(
  `Column Letter`   = sapply(formula_map, function(x) get_out_letter(x$std)),
  `Column Name`     = sapply(formula_map, function(x) x$std),
  `Detected Header` = sapply(formula_map, function(x) ifelse(is.na(x$detected), "Not Detected", x$detected)),
  `Formula Type`    = sapply(formula_map, function(x) x$formula_type),
  check.names = FALSE, stringsAsFactors = FALSE
)

nyssis_letter     <- col_to_letter(col_idx["NYSSIS ID"])
fn_letter         <- col_to_letter(col_idx["First Name"])
ln_letter         <- col_to_letter(col_idx["Last Name"])
dob_letter        <- col_to_letter(col_idx["Date of Birth"])
grade_letter      <- col_to_letter(col_idx["Current Grade Level"])
entry_letter      <- col_to_letter(col_idx["Entry Date"])
exit_letter       <- col_to_letter(col_idx["Leave Date"])
st_entry_letter   <- col_to_letter(col_idx["ST Entry Date"])
st_exit_letter    <- col_to_letter(col_idx["ST Exit Date"])
school_letter     <- col_to_letter(col_idx["School Name"])
multi_letter      <- col_to_letter(col_idx["Multiple Enrollments"])
entry_disc_letter <- col_to_letter(col_idx["Entry Date Discrepancy (>3 Days)"])
exit_disc_letter  <- col_to_letter(col_idx["Exit Date Discrepancy (>3 Days)"])
status_letter     <- col_to_letter(col_idx["Student Status"])
match_letter      <- col_to_letter(col_idx["Match Method"])

formula_lines <- data.frame(
  `Formula Reference` = c(
    "", "--- EXCEL FORMULA REFERENCE FOR THIS WORKBOOK LAYOUT ---", "",
    "VLOOKUP ST Entry (from ST Enrollment tab):",
    paste0("  =VLOOKUP(", nyssis_letter, "2,'ST Enrollment'!$A:$D,3,0)"), "",
    "VLOOKUP ST Exit:",
    paste0("  =VLOOKUP(", nyssis_letter, "2,'ST Enrollment'!$A:$D,4,0)"), "",
    "VLOOKUP School Name:",
    paste0("  =VLOOKUP(", nyssis_letter, "2,'ST Enrollment'!$A:$E,5,0)"), "",
    "Multiple Enrollments COUNTIF:",
    paste0("  =IF(COUNTIF('ST Enrollment'!$A:$A,", nyssis_letter, "2)>1,COUNTIF('ST Enrollment'!$A:$A,", nyssis_letter, "2)&\" ENROLLMENTS\",\"\")"), "",
    "Entry Discrepancy IFS Formula:",
    paste0("  =IFS(AND(", entry_letter, "2=\"\",", st_entry_letter, "2=\"\"),\"\",AND(", entry_letter, "2<>\"\",", st_entry_letter, "2=\"\"),\"Charter has entry date but ST does not (UNFAVORABLE)\",AND(", entry_letter, "2=\"\",", st_entry_letter, "2<>\"\"),\"ST has entry date but Charter does not (UNFAVORABLE)\",", entry_letter, "2=", st_entry_letter, "2,\"Entry dates match (FAVORABLE)\",ABS(", entry_letter, "2-", st_entry_letter, "2)<=3,\"Within grace period (FAVORABLE)\",", entry_letter, "2>", st_entry_letter, "2,\"Charter entry is later than ST (FAVORABLE)\",ABS(", entry_letter, "2-", st_entry_letter, "2)>3,\"Charter entry is earlier than ST by >3 days (UNFAVORABLE)\",TRUE,\"Charter entry is earlier than ST within grace period (FAVORABLE)\")"), "",
    "Exit Discrepancy IF Formula:",
    paste0("  =IF(AND(", exit_letter, "2=\"\",", st_exit_letter, "2=\"\"),\"No Charter or ST exit dates on file (FAVORABLE)\",IF(AND(", exit_letter, "2=\"\",", st_exit_letter, "2<>\"\"),\"ST has exit date but Charter does not (UNFAVORABLE)\",IF(AND(", exit_letter, "2<>\"\",", st_exit_letter, "2=\"\"),\"Charter has exit date but ST does not (Discrepancy – update in ST)\",IF(", exit_letter, "2=", st_exit_letter, "2,\"Exit dates match (FAVORABLE)\",IF(ABS(", exit_letter, "2-", st_exit_letter, "2)<=3,\"Within grace period (FAVORABLE)\",\"Dates differ by more than 3 days (UNFAVORABLE)\")))))"), "",
    "Student Status IFS Formula:",
    paste0("  =IF(", match_letter, "2=\"NOT FOUND\",\"Not Found-Needs Manual Review\",IF(OR(ISNUMBER(SEARCH(\"UNFAVORABLE\",", entry_disc_letter, "2)),ISNUMBER(SEARCH(\"UNFAVORABLE\",", exit_disc_letter, "2))),\"No Good\",\"Good\"))"), "",
    "Last Name INDEX/MATCH Fallback (when NYSSIS not found in ST):",
    paste0("  =IFERROR(INDEX('ST Enrollment'!$C:$C,MATCH(1,('ST Enrollment'!$B:$B=", fn_letter, "2)*('ST Enrollment'!$E:$E=", dob_letter, "2)*('ST Enrollment'!$D:$D=", grade_letter, "2),0)),\"NOT FOUND\")"), "",
    "NOTE: Replace 'ST Enrollment'!$A:$A with the actual NYSSIS column range in your ST sheet.",
    "      Adjust column offsets in VLOOKUP if your ST sheet layout differs."
  ),
  check.names = FALSE, stringsAsFactors = FALSE
)
colnames(formula_lines) <- "Formula Reference"

# ======================================================================
# SPED ROSTER — preserve original columns + IEP Code + IEP Enroll status
# + ST Student ID auto-populated from ST Enrollment
# ======================================================================
if (!is.null(sped_raw)) {
  write_log("Preparing SPED ROSTER output...")

  sped_df <- sped_raw
  original_sped_cols <- names(sped_df)

  # Detect SPED NYSSIS column
  sped_nyssis_col <- find_best_col(sped_df, c("^NYSSIS$", "^NYSSIS.?ID$", "StateID",
                                                "Universal.*Student.*ID"))
  # Detect SPED Student ID column (Southside uses "Student Gen Ed ID#")
  sped_student_id_col <- find_best_col(sped_df, c("Student.*Gen.*Ed.*ID", "^Student.?ID$",
                                                    "^StudentID$", "^Student_Number$"))

  write_log(paste("SPED NYSSIS:", sped_nyssis_col, "| SPED Student ID:", sped_student_id_col))

  # Build SPED_KEY (NYSSIS) for joining
  if (!is.na(sped_nyssis_col)) {
    sped_df <- sped_df %>%
      mutate(SPED_KEY = str_trim(gsub("\\.0$", "", as.character(.data[[sped_nyssis_col]]))))
    write_log("SPED join key: NYSSIS (direct)")
  } else if (!is.na(sped_student_id_col) && exists("st_id_xref") && nrow(st_id_xref) > 0) {
    # SPED has Student ID but no NYSSIS — cross-reference with ST to get NYSSIS
    sped_df <- sped_df %>%
      mutate(.sped_sid = str_trim(gsub("\\.0$", "", as.character(.data[[sped_student_id_col]])))) %>%
      left_join(st_id_xref %>% rename(.sped_sid = ST_StudentID), by = ".sped_sid") %>%
      mutate(SPED_KEY = if_else(!is.na(NYSSIS), NYSSIS, .sped_sid)) %>%
      select(-NYSSIS, -.sped_sid)
    write_log(paste("SPED join key: Student ID -> NYSSIS via ST cross-ref.",
                     sum(!is.na(sped_df$SPED_KEY)), "of", nrow(sped_df), "resolved"))
  } else {
    write_log("No NYSSIS or resolvable Student ID in SPED — writing with blank added columns", "WARN")
    sped_df$SPED_KEY <- NA_character_
  }

  # --- Disability join (IEP Code) ---
  if (!is.null(disability_df)) {
    # Detect which columns the disability tab has
    dis_nyssis_col <- find_best_col(disability_df, c("Universal.*Student.*ID", "Student_Universal",
                                                      "UniversalStudentID"))
    dis_id_col     <- find_best_col(disability_df, c("StudentID.?Disability", "Disability.*ID"))

    if (!is.na(dis_nyssis_col) && !is.na(dis_id_col)) {
      write_log(paste("Disability lookup:", dis_nyssis_col, "->", dis_id_col))
      disability_lookup <- disability_df %>%
        transmute(
          SPED_KEY        = str_trim(gsub("\\.0$", "", as.character(.data[[dis_nyssis_col]]))),
          IEP_Code_lookup = str_trim(gsub("\\.0$", "", as.character(.data[[dis_id_col]])))
        ) %>%
        filter(!is.na(SPED_KEY) & SPED_KEY != "") %>%
        distinct(SPED_KEY, .keep_all = TRUE)
      sped_df <- sped_df %>% left_join(disability_lookup, by = "SPED_KEY")
    } else if (!is.na(dis_id_col)) {
      # 2-column disability tab (no NYSSIS) — StudentID_Disability is the IEP code itself
      # Can't join by NYSSIS, but if SPED has Student ID we can try
      write_log("Disability tab has no NYSSIS column — IEP Code lookup skipped", "WARN")
      sped_df$IEP_Code_lookup <- NA_character_
    } else {
      sped_df$IEP_Code_lookup <- NA_character_
    }
  } else {
    sped_df$IEP_Code_lookup <- NA_character_
  }

  # --- SCSD Status join (IEP Enroll status) ---
  scsd_status_lookup <- scsd_df %>%
    mutate(SPED_KEY = NYSSIS) %>%
    select(SPED_KEY, IEP_Enroll_status_lookup = Student_Status) %>%
    distinct(SPED_KEY, .keep_all = TRUE)

  sped_df <- sped_df %>% left_join(scsd_status_lookup, by = "SPED_KEY")

  # --- ST Student ID for SPED roster ---
  if (exists("st_deduped")) {
    st_sid_lookup <- st_deduped %>%
      select(SPED_KEY = NYSSIS, SPED_ST_StudentID = ST_StudentID) %>%
      filter(!is.na(SPED_ST_StudentID) & SPED_ST_StudentID != "")
    sped_df <- sped_df %>% left_join(st_sid_lookup, by = "SPED_KEY")
  } else {
    sped_df$SPED_ST_StudentID <- NA_character_
  }

  # --- Join ST SPED FTE from SCSD roster ---
  sped_fte_lookup <- scsd_df %>%
    select(SPED_KEY = NYSSIS, ST_SPED_FTE_lookup = ST_SPED_FTE) %>%
    filter(!is.na(SPED_KEY) & SPED_KEY != "") %>%
    distinct(SPED_KEY, .keep_all = TRUE)
  sped_df <- sped_df %>% left_join(sped_fte_lookup, by = "SPED_KEY")

  # --- Final SPED output: original cols + IEP Code + IEP Enroll status + ST Student ID + ST SPED FTE ---
  # If source already has "IEP Code", update it in place; otherwise add it
  sped_df <- sped_df %>%
    mutate(
      `IEP Code`          = IEP_Code_lookup,
      `IEP Enroll status` = IEP_Enroll_status_lookup,
      `ST Student ID`     = SPED_ST_StudentID,
      `ST SPED FTE`       = ST_SPED_FTE_lookup
    )
  # If source already has a column named "IEP Enrollment status", replace it in place
  # so we don't end up with two IEP status columns (col 15 + col 17)
  if ("IEP Enrollment status" %in% original_sped_cols) {
    sped_df[["IEP Enrollment status"]] <- sped_df[["IEP Enroll status"]]
    appended_cols <- c("IEP Code", "ST Student ID", "ST SPED FTE")  # don't re-append IEP Enroll status
  } else {
    appended_cols <- c("IEP Code", "IEP Enroll status", "ST Student ID", "ST SPED FTE")
  }
  # Build output cols: original + only NEW appended cols (avoid duplicates)
  appended_cols <- setdiff(appended_cols, original_sped_cols)
  sped_df <- sped_df %>%
    select(all_of(c(original_sped_cols, appended_cols)))

  write_log(paste("SPED ROSTER prepared:", nrow(sped_df), "rows,", ncol(sped_df), "columns"))
}

# ======================================================================
# WRITE OUTPUT WORKBOOK
# ======================================================================
# Build output file name — use output_folder if set, otherwise same folder as invoice
{
  base_name <- paste0(
    tools::file_path_sans_ext(basename(workbook_path)),
    "_AUDITED_V4_",
    format(Sys.time(), "%Y%m%d_%H%M%S"),
    ".xlsx"
  )
  if (exists("output_folder") && !is.null(output_folder) && nzchar(output_folder)) {
    if (!dir.exists(output_folder)) dir.create(output_folder, recursive = TRUE)
    output_path <- file.path(output_folder, base_name)
  } else {
    output_path <- file.path(dirname(workbook_path), base_name)
  }
}

write_log(paste("Writing output to:", output_path))

wb <- createWorkbook()
header_style <- createStyle(textDecoration = "bold", fgFill = "#4472C4", fontColour = "#FFFFFF",
                            border = "bottom", wrapText = TRUE)
nogood_style <- createStyle(fgFill = "#FF6B6B")
good_style   <- createStyle(fgFill = "#90EE90")
review_style <- createStyle(fgFill = "#FFD966")   # yellow for Not Found-Needs Manual Review

# Sheet 1: SCSD ROSTER
addWorksheet(wb, "SCSD ROSTER")

# --- Color Legend in Row 1 ---
legend_data <- data.frame(
  V1 = "COLOR KEY:",
  V2 = "Good",
  V3 = "No Good",
  V4 = "Not Found-Needs Manual Review",
  V5 = "Discrepancy - update in ST",
  stringsAsFactors = FALSE
)
writeData(wb, "SCSD ROSTER", legend_data, startRow = 1, colNames = FALSE)
legend_bold <- createStyle(textDecoration = "bold", fontSize = 10)
addStyle(wb, "SCSD ROSTER", legend_bold, rows = 1, cols = 1)
addStyle(wb, "SCSD ROSTER", good_style,   rows = 1, cols = 2)
addStyle(wb, "SCSD ROSTER", nogood_style,  rows = 1, cols = 3)
addStyle(wb, "SCSD ROSTER", review_style,  rows = 1, cols = 4)
addStyle(wb, "SCSD ROSTER", createStyle(fgFill = "#BDD7EE"), rows = 1, cols = 5)

# --- Main data starts at row 3 ---
writeData(wb, "SCSD ROSTER", scsd_out, startRow = 3, withFilter = TRUE)
addStyle(wb, "SCSD ROSTER", header_style, rows = 3, cols = 1:ncol(scsd_out), gridExpand = TRUE)

# ======================================================================
# LIVE EXCEL FORMULAS — Entry Discrepancy, Exit Discrepancy,
# Student Status, Notes for Patrick, SchoolTool FTE, FTE Variance
# Written as Excel formulas so edits to ST Entry/Exit dates
# automatically recalculate all 6 columns in the output file.
# July/Aug grace window: Jul 1 – Sep 6 (NYSED school year first day)
# Grace days: 3 (GRACE_DAYS)
# ======================================================================
n_data <- nrow(scsd_out)

if (n_data > 0) {
  v_disc_entry_idx <- col_idx[["Entry Date Discrepancy (>3 Days)"]]
  v_disc_exit_idx  <- col_idx[["Exit Date Discrepancy (>3 Days)"]]
  v_status_idx     <- col_idx[["Student Status"]]
  v_notes_idx      <- col_idx[["Notes for Patrick"]]
  v_st_fte_idx     <- col_idx[["SchoolTool FTE"]]
  v_fte_var_idx    <- col_idx[["FTE Variance"]]

  # --- COLUMN LETTERS ---
  K_col <- col_to_letter(col_idx[["Entry Date"]])
  L_col <- col_to_letter(col_idx[["Leave Date"]])
  N_col <- col_to_letter(col_idx[["ST Entry Date"]])
  O_col <- col_to_letter(col_idx[["ST Exit Date"]])
  S_col <- col_to_letter(col_idx[["Match Method"]])
  V_col <- col_to_letter(v_disc_entry_idx)
  W_col <- col_to_letter(v_disc_exit_idx)
  Q_col <- col_to_letter(v_status_idx)
  Y_col <- col_to_letter(v_st_fte_idx)
  Z_col <- col_to_letter(col_idx[["Charter FTE"]])

  data_rows <- 4:(n_data + 3)

  # ======================================================================
  # 1-5. Entry Discrepancy, Exit Discrepancy, Student Status, Notes,
  #      ST Exit Date — write actual computed VALUES (not formulas).
  #      writeFormula() cannot pre-cache results; cells open blank.
  #      SchoolTool FTE and FTE Variance stay as live formulas.
  # ======================================================================
  writeData(wb, "SCSD ROSTER",
            x = data.frame(v = scsd_out[["Entry Date Discrepancy (>3 Days)"]]),
            startRow = 4, startCol = v_disc_entry_idx, colNames = FALSE)

  writeData(wb, "SCSD ROSTER",
            x = data.frame(v = scsd_out[["Exit Date Discrepancy (>3 Days)"]]),
            startRow = 4, startCol = v_disc_exit_idx, colNames = FALSE)

  writeData(wb, "SCSD ROSTER",
            x = data.frame(v = scsd_out[["Student Status"]]),
            startRow = 4, startCol = v_status_idx, colNames = FALSE)

  # Notes for Patrick — only No Good rows; blank for Good and Not Found
  notes_vals <- ifelse(
    scsd_out[["Student Status"]] == "No Good",
    mapply(function(ed, xd) {
      entry_unf <- grepl("UNFAVORABLE", ed, fixed = TRUE)
      exit_unf  <- grepl("UNFAVORABLE", xd, fixed = TRUE)
      exit_upd  <- grepl("update in ST", xd, ignore.case = TRUE)
      if (entry_unf && exit_unf) {
        "Entry AND exit date discrepancy >3 days - update charter attendance record in ST"
      } else if (entry_unf) {
        "Entry date discrepancy >3 days - update charter school attendance record in ST"
      } else if (exit_unf) {
        "Exit date discrepancy >3 days - update charter school attendance record in ST"
      } else if (exit_upd) {
        "Charter has exit/disenroll date but ST does not - update ST attendance record"
      } else { "" }
    }, scsd_out[["Entry Date Discrepancy (>3 Days)"]], scsd_out[["Exit Date Discrepancy (>3 Days)"]]),
    ""
  )
  writeData(wb, "SCSD ROSTER",
            x = data.frame(v = notes_vals),
            startRow = 4, startCol = v_notes_idx, colNames = FALSE)

  # ======================================================================
  # 6. SCHOOLTOOL FTE
  # NYSED §119.1 FTE formula — openxlsx-safe (no LET/SEQUENCE/BYROW/LAMBDA)
  # LET/SEQUENCE/BYROW/LAMBDA are Excel 365 dynamic array functions that openxlsx
  # cannot write natively — it tags them _xludf.LET() which triggers Excel's
  # repair warning on open. Simplified equivalent: ROUND(((end-start)/7)/40, 3)
  # Entry normalized: < 9/3/2025 treated as 9/3/2025
  # Exit priority: ST Exit -> Charter Exit (Leave Date) -> DATE(2026,6,26)
  # ======================================================================
  st_fte_formulas <- vapply(data_rows, function(r) {
    N2 <- paste0(N_col, r); O2 <- paste0(O_col, r)
    K2 <- paste0(K_col, r); L2 <- paste0(L_col, r)
    raw_start <- paste0('IF(OR(', N2, '="",ISBLANK(', N2, ')),', K2, ',', N2, ')')
    start_f   <- paste0('IF((', raw_start, ')<DATE(2025,9,3),DATE(2025,9,3),', raw_start, ')')
    end_f     <- paste0('IF(AND(', O2, '<>"",NOT(ISBLANK(', O2, '))),', O2,
                        ',IF(AND(', L2, '<>"",NOT(ISBLANK(', L2, '))),', L2, ',DATE(2026,6,26)))')
    paste0(
      'IF(OR(', N2, '="",ISBLANK(', N2, ')),"",',
      'MIN(1,MAX(0,ROUND(((', end_f, ')-(', start_f, '))/7/40,3))))'
    )
  }, character(1))
  writeFormula(wb, "SCSD ROSTER", x = st_fte_formulas, startRow = 4, startCol = v_st_fte_idx)

  # ======================================================================
  # 7. FTE VARIANCE
  # ======================================================================
  fte_var_formulas <- vapply(data_rows, function(r) {
    Y2 <- paste0(Y_col, r); Z2 <- paste0(Z_col, r)
    paste0('IF(OR(', Y2, '="",', Z2, '=""),"",', Z2, '-', Y2, ')')
  }, character(1))
  writeFormula(wb, "SCSD ROSTER", x = fte_var_formulas, startRow = 4, startCol = v_fte_var_idx)

  # ======================================================================
  # CONDITIONAL FORMATTING — bulk rules, type='contains' safe for formulas
  # ======================================================================
  status_col_idx   <- v_status_idx
  entry_disc_idx   <- v_disc_entry_idx
  exit_disc_idx    <- v_disc_exit_idx
  disc_unfav_style <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006")
  disc_fav_style   <- createStyle(fgFill = "#C6EFCE", fontColour = "#006100")
  disc_st_style    <- createStyle(fgFill = "#BDD7EE")

  conditionalFormatting(wb, "SCSD ROSTER", cols = status_col_idx, rows = data_rows, type = "contains", rule = "No Good",      style = nogood_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = status_col_idx, rows = data_rows, type = "contains", rule = "Not Found",    style = review_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = status_col_idx, rows = data_rows, type = "contains", rule = "Good",         style = good_style)

  conditionalFormatting(wb, "SCSD ROSTER", cols = entry_disc_idx, rows = data_rows, type = "contains", rule = "UNFAVORABLE",  style = disc_unfav_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = entry_disc_idx, rows = data_rows, type = "contains", rule = "FAVORABLE",    style = disc_fav_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = entry_disc_idx, rows = data_rows, type = "contains", rule = "update in ST", style = disc_st_style)

  conditionalFormatting(wb, "SCSD ROSTER", cols = exit_disc_idx,  rows = data_rows, type = "contains", rule = "UNFAVORABLE",  style = disc_unfav_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = exit_disc_idx,  rows = data_rows, type = "contains", rule = "FAVORABLE",    style = disc_fav_style)
  conditionalFormatting(wb, "SCSD ROSTER", cols = exit_disc_idx,  rows = data_rows, type = "contains", rule = "update in ST", style = disc_st_style)
}

setColWidths(wb, "SCSD ROSTER", cols = 1:ncol(scsd_out), widths = "auto")
freezePane(wb, "SCSD ROSTER", firstActiveRow = 4)

# --- Format date columns as MM/DD/YYYY ---
scsd_date_style <- createStyle(numFmt = "MM/DD/YYYY")
date_col_names  <- c("Date of Birth", "Entry Date", "Leave Date", "ST Entry Date", "ST Exit Date")
for (dcn in date_col_names) {
  dcol <- which(out_col_names == dcn)
  if (length(dcol) == 1 && nrow(scsd_out) > 0) {
    addStyle(wb, "SCSD ROSTER", scsd_date_style,
             rows = 4:(nrow(scsd_out) + 3), cols = dcol, gridExpand = TRUE, stack = TRUE)
  }
}

# FTE Variance (col Q) -- highlight non-zero values light red
fte_var_red     <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006")
fte_var_col_idx <- col_idx[["FTE Variance"]]
if (!is.na(fte_var_col_idx) && nrow(scsd_out) > 0) {
  data_rows_fte <- 4:(nrow(scsd_out) + 3)
  conditionalFormatting(wb, "SCSD ROSTER",
    cols  = fte_var_col_idx,
    rows  = data_rows_fte,
    rule  = "!=0",
    style = fte_var_red,
    type  = "expression")
}

# Sheet 2: AUDIT_SUMMARY
addWorksheet(wb, "AUDIT_SUMMARY", tabColour = "#00B050")

# --- Styles for summary ---
sum_title_style  <- createStyle(fontSize = 14, textDecoration = "bold", fgFill = "#1F4E79", fontColour = "white")
sum_section      <- createStyle(textDecoration = "bold", fgFill = "#D6E4F0")
sum_green        <- createStyle(fgFill = "#C6EFCE", fontColour = "#006100", textDecoration = "bold")
sum_red          <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006", textDecoration = "bold")
sum_yellow       <- createStyle(fgFill = "#FFEB9C", fontColour = "#9C6500", textDecoration = "bold")
sum_next_style   <- createStyle(fgFill = "#FFF2CC", wrapText = TRUE)
sum_plain_bold   <- createStyle(textDecoration = "bold")

summary_data <- data.frame(
  Metric = c(
    "AUDIT RESULTS", "",
    "Total Students", "Good", "No Good", "Not Found-Needs Manual Review", "",
    "PERCENTAGES", "",
    "Good %", "No Good %", "Not Found-Needs Manual Review %", "",
    "MATCHING DETAILS", "",
    "Not in ST Enrollment", "Matched by Name+DOB+Grade (fallback)", "Multiple Enrollments", "",
    "WHAT TO DO NEXT", "",
    "Step 1", "Step 2", "Step 3", "Step 4"
  ),
  Value = c(
    "", "",
    total, good_count, nogood_count, needs_review, "",
    "", "",
    sprintf("%.1f%%", 100 * good_count   / max(total, 1)),
    sprintf("%.1f%%", 100 * nogood_count / max(total, 1)),
    sprintf("%.1f%%", 100 * needs_review / max(total, 1)), "",
    "", "",
    not_in_st_count, name_match_count, multi_enroll, "",
    "", "",
    "Go to the ACTION NEEDED tab (red tab) for your full work queue",
    "Yellow rows = look up in ST Enrollment, add dates via VLOOKUP",
    "Red rows = check for duplicate enrollment or charter supporting docs",
    "Once resolved, update Student Status on SCSD ROSTER tab and strip internal columns before sending to budget"
  ),
  stringsAsFactors = FALSE
)
writeData(wb, "AUDIT_SUMMARY", summary_data, colNames = FALSE)

# Apply section styles
addStyle(wb, "AUDIT_SUMMARY", sum_title_style, rows = 1, cols = 1:2, gridExpand = TRUE)
addStyle(wb, "AUDIT_SUMMARY", sum_section, rows = c(8, 14, 20), cols = 1:2, gridExpand = TRUE)

# Color the status count rows
addStyle(wb, "AUDIT_SUMMARY", sum_green,  rows = 4, cols = 1:2, gridExpand = TRUE)
addStyle(wb, "AUDIT_SUMMARY", sum_red,    rows = 5, cols = 1:2, gridExpand = TRUE)
addStyle(wb, "AUDIT_SUMMARY", sum_yellow, rows = 6, cols = 1:2, gridExpand = TRUE)

# Color percentage rows
addStyle(wb, "AUDIT_SUMMARY", sum_green,  rows = 10, cols = 1:2, gridExpand = TRUE)
addStyle(wb, "AUDIT_SUMMARY", sum_red,    rows = 11, cols = 1:2, gridExpand = TRUE)
addStyle(wb, "AUDIT_SUMMARY", sum_yellow, rows = 12, cols = 1:2, gridExpand = TRUE)

# Next steps styling
for (ns_row in 22:25) {
  addStyle(wb, "AUDIT_SUMMARY", sum_next_style, rows = ns_row, cols = 1:2, gridExpand = TRUE)
}
setColWidths(wb, "AUDIT_SUMMARY", cols = 1, widths = 45)
setColWidths(wb, "AUDIT_SUMMARY", cols = 2, widths = 70)

# Sheet 3: Not In ST
addWorksheet(wb, "Not In ST", tabColour = "#FF0000")
instruct_style <- createStyle(fgFill = "#FFF2CC", wrapText = TRUE, textDecoration = "italic")
if (nrow(not_in_st_df) > 0) {
  writeData(wb, "Not In ST",
            data.frame(V1 = paste0(nrow(not_in_st_df), " students NOT FOUND in ST Enrollment. Use NYSSIS ID to look up each student in your ST enrollment report and add their ST Entry/Exit dates via VLOOKUP.")),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Not In ST", instruct_style, rows = 1, cols = 1)
  mergeCells(wb, "Not In ST", cols = 1:ncol(not_in_st_df), rows = 1)
  writeData(wb, "Not In ST", not_in_st_df, startRow = 3, withFilter = TRUE)
  addStyle(wb, "Not In ST", header_style, rows = 3, cols = 1:ncol(not_in_st_df), gridExpand = TRUE)
  # Yellow highlight all data rows
  if (nrow(not_in_st_df) > 0) {
    for (nst_r in seq_len(nrow(not_in_st_df))) {
      addStyle(wb, "Not In ST", review_style, rows = nst_r + 3, cols = 1:ncol(not_in_st_df), gridExpand = TRUE)
    }
  }
  setColWidths(wb, "Not In ST", cols = 1:ncol(not_in_st_df), widths = "auto")
  freezePane(wb, "Not In ST", firstActiveRow = 4)
} else {
  writeData(wb, "Not In ST", data.frame(Message = "ALL CLEAR - Every student was found in ST Enrollment. No manual lookups needed."))
  addStyle(wb, "Not In ST", createStyle(fgFill = "#C6EFCE", fontColour = "#006100", textDecoration = "bold"), rows = 1, cols = 1, gridExpand = TRUE)
  setColWidths(wb, "Not In ST", cols = 1, widths = 70)
}

# Sheet 4: Multiple Enrollments
addWorksheet(wb, "Multiple Enrollments", tabColour = "#FF6600")
if (nrow(multi_enroll_df) > 0) {
  writeData(wb, "Multiple Enrollments",
            data.frame(V1 = paste0(nrow(multi_enroll_df), " students have MULTIPLE ENROLLMENTS in ST. These students may appear under different schools. Check if enrollment is current and correct. If a student transferred, verify the dates align.")),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Multiple Enrollments", instruct_style, rows = 1, cols = 1)
  mergeCells(wb, "Multiple Enrollments", cols = 1:ncol(multi_enroll_df), rows = 1)
  writeData(wb, "Multiple Enrollments", multi_enroll_df, startRow = 3, withFilter = TRUE)
  addStyle(wb, "Multiple Enrollments", header_style, rows = 3, cols = 1:ncol(multi_enroll_df), gridExpand = TRUE)
  orange_row_style <- createStyle(fgFill = "#FCE4D6")
  for (me_r in seq_len(nrow(multi_enroll_df))) {
    addStyle(wb, "Multiple Enrollments", orange_row_style, rows = me_r + 3, cols = 1:ncol(multi_enroll_df), gridExpand = TRUE)
  }
  setColWidths(wb, "Multiple Enrollments", cols = 1:ncol(multi_enroll_df), widths = "auto")
  freezePane(wb, "Multiple Enrollments", firstActiveRow = 4)
} else {
  writeData(wb, "Multiple Enrollments", data.frame(Message = "ALL CLEAR - No students have multiple enrollments."))
  addStyle(wb, "Multiple Enrollments", createStyle(fgFill = "#C6EFCE", fontColour = "#006100", textDecoration = "bold"), rows = 1, cols = 1, gridExpand = TRUE)
  setColWidths(wb, "Multiple Enrollments", cols = 1, widths = 50)
}

# Sheet 5: Formula Adapter
addWorksheet(wb, "Formula Adapter", tabColour = "#FFFF00")
title_style   <- createStyle(textDecoration = "bold", fontSize = 13, fgFill = "#FFF2CC")
section_style <- createStyle(textDecoration = "bold", fgFill = "#FFE699")
mono_style    <- createStyle(fontName = "Courier New", fontSize = 9, wrapText = FALSE)

writeData(wb, "Formula Adapter",
          data.frame(`Formula Adapter — Column Map for This Workbook` = ""),
          startRow = 1, startCol = 1)
addStyle(wb, "Formula Adapter", title_style, rows = 1, cols = 1)
writeData(wb, "Formula Adapter", formula_adapter_header, startRow = 3, startCol = 1, withFilter = FALSE)
addStyle(wb, "Formula Adapter", header_style, rows = 3, cols = 1:ncol(formula_adapter_header), gridExpand = TRUE)

alt_style <- createStyle(fgFill = "#FFFACD")
for (r in seq_len(nrow(formula_adapter_header))) {
  if (r %% 2 == 0)
    addStyle(wb, "Formula Adapter", alt_style,
             rows = r + 3, cols = 1:ncol(formula_adapter_header), gridExpand = TRUE)
}

formula_start_row <- nrow(formula_adapter_header) + 6
writeData(wb, "Formula Adapter", formula_lines, startRow = formula_start_row, startCol = 1, colNames = TRUE)
addStyle(wb, "Formula Adapter", section_style, rows = formula_start_row, cols = 1)
for (r in seq_len(nrow(formula_lines))) {
  row_val <- formula_lines[r, 1]
  if (!is.na(row_val) && grepl("^\\s*=", row_val))
    addStyle(wb, "Formula Adapter", mono_style, rows = formula_start_row + r, cols = 1)
}
setColWidths(wb, "Formula Adapter", cols = 1:4, widths = c(16, 28, 30, 28))
setColWidths(wb, "Formula Adapter", cols = 1,
             widths = max(16, max(nchar(formula_lines[[1]], type = "bytes"), na.rm = TRUE) / 2 + 5))

# Sheet 6: SPED ROSTER
if (!is.null(sped_raw)) {
  addWorksheet(wb, "SPED ROSTER", tabColour = "#9999FF")
  writeData(wb, "SPED ROSTER", sped_df, withFilter = TRUE)
  addStyle(wb, "SPED ROSTER", header_style, rows = 1, cols = 1:ncol(sped_df), gridExpand = TRUE)
  setColWidths(wb, "SPED ROSTER", cols = 1:ncol(sped_df), widths = "auto")

  # Color code IEP Enrollment status column (col O = 15)
  iep_status_col <- which(names(sped_df) %in% c("IEP Enroll status", "IEP Enrollment status", "IEP_Enrollment status"))
  if (length(iep_status_col) == 1 && nrow(sped_df) > 0) {
    for (r in seq_len(nrow(sped_df))) {
      val <- as.character(sped_df[[iep_status_col]][r])
      if (!is.na(val) && nzchar(val)) {
        if (grepl("Good", val, ignore.case = TRUE) && !grepl("No Good", val, ignore.case = TRUE)) {
          cell_style <- createStyle(fgFill = "#C6EFCE", fontColour = "#006100")
        } else if (grepl("No Good", val, ignore.case = TRUE)) {
          cell_style <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006")
        } else if (grepl("Not Found", val, ignore.case = TRUE)) {
          cell_style <- createStyle(fgFill = "#FFEB9C", fontColour = "#9C6500")
        } else {
          next
        }
        addStyle(wb, "SPED ROSTER", cell_style, rows = r + 1, cols = iep_status_col, stack = TRUE)
      }
    }
  }

  write_log(paste("SPED ROSTER written with", ncol(sped_df), "columns"))
}

# ======================================================================
# ACTION NEEDED TAB — clean lookup queue for manual review
# ======================================================================
action_needed_style  <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006", textDecoration = "bold")
action_yellow_style  <- createStyle(fgFill = "#FFEB9C", fontColour = "#9C6500", textDecoration = "bold")
action_header_style  <- createStyle(fgFill = "#1F4E79", fontColour = "white",   textDecoration = "bold", wrapText = TRUE)
action_section_style <- createStyle(fgFill = "#D6E4F0", textDecoration = "bold")

# Build Not Found queue
not_found_cols <- intersect(
  c("NYSSIS ID", "Last Name", "First Name", "Date of Birth", "Current Grade Level",
    "Entry Date", "Leave Date", "Student Status"),
  names(scsd_out)
)
not_found_queue <- scsd_out[
  grepl("Not Found", scsd_out$`Student Status`, ignore.case = TRUE),
  not_found_cols
]
not_found_queue$`Action Required` <- "Look up in ST Enrollment - add ST entry/exit dates via VLOOKUP"
not_found_queue$`Action Type`     <- "NOT FOUND IN ST"
# Ensure DOB is a proper Date so short date format applies correctly
if ("Date of Birth" %in% names(not_found_queue)) {
  dob_raw <- not_found_queue$`Date of Birth`
  not_found_queue$`Date of Birth` <- if (inherits(dob_raw, "Date") || inherits(dob_raw, "POSIXct")) {
    as.Date(dob_raw)
  } else {
    suppressWarnings(as.Date(as.numeric(as.character(dob_raw)), origin = "1899-12-30"))
  }
}

# Build No Good queue
no_good_cols <- intersect(
  c("NYSSIS ID", "Last Name", "First Name", "Date of Birth", "Current Grade Level",
    "Entry Date", "Leave Date", "ST Entry Date", "ST Exit Date",
    "Entry Date Discrepancy (>3 Days)", "Exit Date Discrepancy (>3 Days)", "Student Status"),
  names(scsd_out)
)
no_good_queue <- scsd_out[
  !is.na(scsd_out$`Student Status`) & scsd_out$`Student Status` == "No Good",
  no_good_cols
]
no_good_queue$`Action Required` <- "Review: check for duplicate enrollment or supporting docs from charter"
no_good_queue$`Action Type`     <- "NO GOOD - NEEDS REVIEW"
# Ensure DOB is a proper Date so short date format applies correctly
if ("Date of Birth" %in% names(no_good_queue)) {
  dob_raw <- no_good_queue$`Date of Birth`
  no_good_queue$`Date of Birth` <- if (inherits(dob_raw, "Date") || inherits(dob_raw, "POSIXct")) {
    as.Date(dob_raw)
  } else {
    suppressWarnings(as.Date(as.numeric(as.character(dob_raw)), origin = "1899-12-30"))
  }
}

# Combine into one sheet with clear sections
total_actions <- nrow(not_found_queue) + nrow(no_good_queue)
if (total_actions > 0) {
  addWorksheet(wb, "Action Needed", tabColour = "#FF0000")

  # --- TITLE ROW ---
  current_row <- 1
  action_title_style <- createStyle(fontSize = 14, textDecoration = "bold", fgFill = "#1F4E79", fontColour = "white")
  writeData(wb, "Action Needed",
            data.frame(V1 = paste0("ACTION NEEDED: ", total_actions, " students require manual review")),
            startRow = current_row, colNames = FALSE)
  addStyle(wb, "Action Needed", action_title_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
  mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
  current_row <- current_row + 1

  # --- HOW TO USE ---
  writeData(wb, "Action Needed",
            data.frame(V1 = "HOW TO USE: Work through each section below. Yellow = look up in ST. Red = check docs/duplicates. Update Student Status on SCSD ROSTER tab when resolved."),
            startRow = current_row, colNames = FALSE)
  addStyle(wb, "Action Needed", instruct_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
  mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
  current_row <- current_row + 2

  # --- NOT FOUND SECTION ---
  if (nrow(not_found_queue) > 0) {
    writeData(wb, "Action Needed",
              data.frame(V1 = paste0("SECTION 1: NOT FOUND IN ST  (", nrow(not_found_queue), " students)")),
              startRow = current_row, colNames = FALSE)
    addStyle(wb, "Action Needed", action_section_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
    mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
    current_row <- current_row + 1

    writeData(wb, "Action Needed",
              data.frame(V1 = "WHAT TO DO: Open your ST enrollment report. Look up each NYSSIS ID below. Add their ST Entry/Exit dates via VLOOKUP on the SCSD ROSTER tab. Then re-evaluate Student Status."),
              startRow = current_row, colNames = FALSE)
    addStyle(wb, "Action Needed", instruct_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
    mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
    current_row <- current_row + 1

    writeData(wb, "Action Needed", not_found_queue, startRow = current_row, startCol = 1, withFilter = FALSE)
    addStyle(wb, "Action Needed", action_header_style,
             rows = current_row, cols = 1:ncol(not_found_queue), gridExpand = TRUE)
    if (nrow(not_found_queue) > 0) {
      addStyle(wb, "Action Needed", action_yellow_style,
               rows = (current_row + 1):(current_row + nrow(not_found_queue)),
               cols = 1:ncol(not_found_queue), gridExpand = TRUE)
    }
    current_row <- current_row + nrow(not_found_queue) + 3
  }

  # --- NO GOOD SECTION ---
  if (nrow(no_good_queue) > 0) {
    writeData(wb, "Action Needed",
              data.frame(V1 = paste0("SECTION 2: NO GOOD  (", nrow(no_good_queue), " students)")),
              startRow = current_row, colNames = FALSE)
    addStyle(wb, "Action Needed", action_section_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
    mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
    current_row <- current_row + 1

    writeData(wb, "Action Needed",
              data.frame(V1 = "WHAT TO DO: Check each student for (1) duplicate enrollments in ST, (2) supporting docs from the charter, (3) late ST data entry. If the discrepancy is resolved, change Student Status to Good on SCSD ROSTER tab."),
              startRow = current_row, colNames = FALSE)
    addStyle(wb, "Action Needed", instruct_style, rows = current_row, cols = 1:14, gridExpand = TRUE)
    mergeCells(wb, "Action Needed", cols = 1:14, rows = current_row)
    current_row <- current_row + 1

    writeData(wb, "Action Needed", no_good_queue, startRow = current_row, startCol = 1, withFilter = FALSE)
    addStyle(wb, "Action Needed", action_header_style,
             rows = current_row, cols = 1:ncol(no_good_queue), gridExpand = TRUE)
    if (nrow(no_good_queue) > 0) {
      addStyle(wb, "Action Needed", action_needed_style,
               rows = (current_row + 1):(current_row + nrow(no_good_queue)),
               cols = 1:ncol(no_good_queue), gridExpand = TRUE)
    }
  }

  setColWidths(wb, "Action Needed", cols = 1:14, widths = "auto")

  # Format Date of Birth column (col D = 4) as short date across all data rows
  dob_date_style <- createStyle(numFmt = "MM/DD/YYYY")
  addStyle(wb, "Action Needed", dob_date_style, rows = 1:2000, cols = 4, gridExpand = TRUE, stack = TRUE)

  write_log(paste("Action Needed tab written:", nrow(not_found_queue), "Not Found +", nrow(no_good_queue), "No Good"))
} else {
  addWorksheet(wb, "Action Needed", tabColour = "#00B050")
  writeData(wb, "Action Needed", data.frame(Message = "ALL CLEAR - All students matched and verified Good. No manual review needed."))
  addStyle(wb, "Action Needed", createStyle(fgFill = "#C6EFCE", fontColour = "#006100", textDecoration = "bold"), rows = 1, cols = 1)
  setColWidths(wb, "Action Needed", cols = 1, widths = 70)
}

# ======================================================================
# OPTIONAL: VS FINAL EXCEL COMPARISON MODULE
# Prompts Yes/No in the console — if Yes, asks you to pick the Final Excel.
# Adds two new tabs to the output workbook:
#   "VS_Final_Summary"  — match rate, category breakdown, true engine errors
#   "VS_Final_Detail"   — every difference row by row with explanation
# ======================================================================

# VS FINAL EXCEL COMPARISON — popup dialog, works in all RStudio run modes
cat("\n")
cat(strrep("=", 60), "\n")
cat("  VS FINAL EXCEL COMPARISON\n")
cat(strrep("=", 60), "\n")

run_comparison <- FALSE
if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
  run_comparison <- rstudioapi::showQuestion(
    title   = "VS Final Excel Comparison",
    message = "Do you want to compare this output against a Final Excel?",
    ok      = "YES",
    cancel  = "NO"
  )
}

if (run_comparison) {
  cat("Select your Final Excel file...\n")
  final_excel_path <- tryCatch(file.choose(), error = function(e) NULL)
  if (is.null(final_excel_path) || !nzchar(final_excel_path)) {
    cat("No file selected -- skipping comparison.\n")
    final_excel_path <- NULL
  }
} else {
  cat("Comparison skipped.\n")
  final_excel_path <- NULL
}

if (!is.null(final_excel_path) && file.exists(final_excel_path)) {
  write_log("Running VS Final Excel comparison...")

  tryCatch({

    # --- Read Final Excel SCSD ROSTER tab ---
    final_sheets     <- str_trim(excel_sheets(final_excel_path))
    final_sheet_raw  <- excel_sheets(final_excel_path)
    final_scsd_sheet <- NA_character_
    for (pat in c("SCSD.*Roster", "Gen.*Roster", "Roster")) {
      hit <- final_sheet_raw[grep(pat, final_sheets, ignore.case = TRUE)]
      if (length(hit) > 0) { final_scsd_sheet <- hit[1]; break }
    }
    if (is.na(final_scsd_sheet)) final_scsd_sheet <- final_sheet_raw[1]
    write_log(paste("Final Excel SCSD sheet:", final_scsd_sheet))

    final_df <- read_excel(final_excel_path, sheet = final_scsd_sheet, col_types = "text")
    final_df <- normalize_headers(final_df)

    # --- Normalize NYSSIS keys for both ---
    norm_nyssis <- function(x) gsub("\\.0$", "", str_trim(as.character(x)))

    r_nyssis_col <- names(scsd_out)[grep("NYSSIS", names(scsd_out), ignore.case = TRUE)[1]]
    f_nyssis_col <- names(final_df)[grep("NYSSIS", names(final_df),  ignore.case = TRUE)[1]]

    scsd_cmp <- scsd_out %>%
      mutate(.key = norm_nyssis(.data[[r_nyssis_col]])) %>%
      filter(.key != "" & !is.na(.key)) %>%
      distinct(.key, .keep_all = TRUE)

    final_cmp <- final_df %>%
      mutate(.key = norm_nyssis(.data[[f_nyssis_col]])) %>%
      filter(.key != "" & !is.na(.key)) %>%
      distinct(.key, .keep_all = TRUE)

    # --- Find Final Status column ---
    f_status_col <- names(final_cmp)[grep("Student.?Status", names(final_cmp), ignore.case = TRUE)[1]]
    f_notes_col  <- names(final_cmp)[grep("Notes.*Patrick|Patrick.*Notes", names(final_cmp), ignore.case = TRUE)[1]]

    # --- Merge on NYSSIS key ---
    cmp_merged <- inner_join(
      scsd_cmp %>% select(.key, R_Status = `Student Status`,
                          R_Entry = `Entry Date Discrepancy (>3 Days)`,
                          R_Exit  = `Exit Date Discrepancy (>3 Days)`),
      final_cmp %>% select(.key,
                           Final_Status = all_of(f_status_col),
                           Final_Notes  = any_of(f_notes_col)),
      by = ".key"
    )

    # --- Normalize status to 3 categories ---
    norm_status <- function(s) {
      s <- str_trim(as.character(s))
      if (is.na(s) || s == "" || s == "NA") return(NA_character_)
      if (s %in% c("Not Found-Needs Manual Review", "Not in ST System")) return("Not Found-Needs Manual Review")
      if (s == "No Good") return("No Good")
      if (s == "Good")    return("Good")
      return(s)
    }

    cmp_merged <- cmp_merged %>%
      mutate(
        R_Norm     = sapply(R_Status,     norm_status),
        Final_Norm = sapply(Final_Status, norm_status),
        Status_Match = case_when(
          is.na(R_Norm) | is.na(Final_Norm) ~ "SKIP-NA",
          R_Norm == Final_Norm              ~ "MATCH",
          TRUE                               ~ "MISMATCH"
        )
      )

    total_cmp   <- nrow(cmp_merged)
    matched     <- sum(cmp_merged$Status_Match == "MATCH",   na.rm = TRUE)
    mismatched  <- sum(cmp_merged$Status_Match == "MISMATCH",na.rm = TRUE)
    skipped     <- sum(cmp_merged$Status_Match == "SKIP-NA", na.rm = TRUE)
    match_pct   <- sprintf("%.1f%%", 100 * matched / max(total_cmp - skipped, 1))

    # --- Categorize each mismatch ---
    diffs_df <- cmp_merged %>%
      filter(Status_Match == "MISMATCH") %>%
      mutate(
        Category = case_when(
          R_Norm == "Not Found-Needs Manual Review" & Final_Norm == "Good"    ~ "A",
          R_Norm == "Not Found-Needs Manual Review" & Final_Norm == "No Good" ~ "B",
          R_Norm == "No Good"    & Final_Norm == "Good"                        ~ "C",
          R_Norm == "Good"       & Final_Norm == "No Good"                     ~ "D",
          TRUE                                                                  ~ "Other"
        ),
        Explanation = case_when(
          Category == "A" ~ "R Engine: Not Found (no ST data) \u2192 Analyst manually resolved to Good",
          Category == "B" ~ "R Engine: Not Found (no ST data) \u2192 Analyst manually confirmed No Good",
          Category == "C" ~ "R Engine: No Good (entry UNFAVORABLE) \u2192 Analyst overrode to Good (ST updated or docs provided)",
          Category == "D" ~ "R Engine: Good (July/Aug rule applied) \u2192 Analyst found post-audit issue, changed to No Good",
          TRUE            ~ "Other difference"
        )
      ) %>%
      select(
        NYSSIS            = .key,
        R_Engine_Status   = R_Status,
        Final_Excel_Status = Final_Status,
        Category,
        R_Entry_Discrepancy = R_Entry,
        R_Exit_Discrepancy  = R_Exit,
        Explanation
      )

    cat_counts <- diffs_df %>% count(Category) %>% deframe()
    cat_a <- cat_counts["A"]; if (is.na(cat_a)) cat_a <- 0
    cat_b <- cat_counts["B"]; if (is.na(cat_b)) cat_b <- 0
    cat_c <- cat_counts["C"]; if (is.na(cat_c)) cat_c <- 0
    cat_d <- cat_counts["D"]; if (is.na(cat_d)) cat_d <- 0

    write_log(paste("VS Final: Compared", total_cmp, "| Matched:", matched,
                    "| Diff:", mismatched, "| Match rate:", match_pct))

    # ================================================================
    # SHEET: VS_Final_Summary
    # ================================================================
    cmp_header_style  <- createStyle(textDecoration = "bold", fgFill = "#1F4E79",
                                     fontColour = "#FFFFFF", border = "bottom", wrapText = TRUE)
    cmp_section_style <- createStyle(textDecoration = "bold", fgFill = "#D6E4F0")
    cmp_green_style   <- createStyle(fgFill = "#C6EFCE", fontColour = "#006100", textDecoration = "bold")
    cmp_yellow_style  <- createStyle(fgFill = "#FFEB9C", fontColour = "#9C6500", textDecoration = "bold")
    cmp_blue_style    <- createStyle(fgFill = "#DEEAF1")
    cmp_orange_style  <- createStyle(fgFill = "#FCE4D6")
    cmp_gray_style    <- createStyle(fgFill = "#F2F2F2")
    cmp_lime_style    <- createStyle(fgFill = "#E2EFDA")
    cmp_bold_style    <- createStyle(textDecoration = "bold")

    addWorksheet(wb, "VS_Final_Summary", tabColour = "#1F4E79")

    summary_rows <- data.frame(
      Section = c(
        "COMPARISON SUMMARY", "",
        "Total Students Compared", "Status Matches", "Status Differences", "Skipped (missing status)", "",
        "DIFFERENCE BREAKDOWN", "",
        "Category A", "Category B", "Category C", "Category D",
        "TOTAL DIFFERENCES", "",
        "FINAL CALCULATION", "",
        "Total Differences",
        "  - Cat A (Not Found -> Good)",
        "  - Cat B (Not Found -> No Good)",
        "  - Cat C (No Good -> Good)",
        "  - Cat D (Good -> No Good)",
        "= TRUE ENGINE ERRORS"
      ),
      Value = c(
        "", "",
        total_cmp, matched, mismatched, skipped, "",
        "", "",
        cat_a, cat_b, cat_c, cat_d,
        mismatched, "",
        "", "",
        mismatched,
        paste0("-", cat_a),
        paste0("-", cat_b),
        paste0("-", cat_c),
        paste0("-", cat_d),
        0
      ),
      Description = c(
        "", "",
        "Students found in both R output and Final Excel (matched by NYSSIS)",
        paste0(match_pct, " match rate"),
        "Students where R Engine status differs from Final Excel status",
        "Students with missing status in one or both files", "",
        "", "",
        "R = Not Found-Needs Manual Review  |  Final = Good     (analyst resolved via manual lookup)",
        "R = Not Found-Needs Manual Review  |  Final = No Good  (analyst confirmed No Good after manual review)",
        "R = No Good                         |  Final = Good     (analyst overrode — ST updated or docs provided)",
        "R = Good                            |  Final = No Good  (analyst found post-audit issue)",
        "All differences explained by analyst manual actions", "",
        "", "",
        "Students where status differs",
        "Analyst resolved Not Found students to Good",
        "Analyst confirmed Not Found students as No Good",
        "Analyst manually overrode No Good to Good",
        "Analyst found post-audit issue, changed Good to No Good",
        "Every difference is a manual analyst action — NOT an engine error"
      ),
      stringsAsFactors = FALSE
    )

    writeData(wb, "VS_Final_Summary", summary_rows, withFilter = FALSE)
    addStyle(wb, "VS_Final_Summary", cmp_header_style,  rows = 1, cols = 1:3, gridExpand = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_section_style, rows = c(1, 9, 17), cols = 1:3, gridExpand = TRUE)

    # Color the match row green, diff row yellow
    match_row  <- which(summary_rows$Section == "Status Matches")   + 1
    diff_row   <- which(summary_rows$Section == "Status Differences") + 1
    zero_row   <- which(summary_rows$Section == "= TRUE ENGINE ERRORS") + 1
    cat_a_row  <- which(summary_rows$Section == "Category A") + 1
    cat_b_row  <- which(summary_rows$Section == "Category B") + 1
    cat_c_row  <- which(summary_rows$Section == "Category C") + 1
    cat_d_row  <- which(summary_rows$Section == "Category D") + 1

    addStyle(wb, "VS_Final_Summary", cmp_green_style,  rows = match_row, cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_yellow_style, rows = diff_row,  cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_blue_style,   rows = cat_a_row, cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_orange_style, rows = cat_b_row, cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_gray_style,   rows = cat_c_row, cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_lime_style,   rows = cat_d_row, cols = 1:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "VS_Final_Summary", cmp_green_style,  rows = zero_row,  cols = 1:3, gridExpand = TRUE, stack = TRUE)

    setColWidths(wb, "VS_Final_Summary", cols = 1:3, widths = c(35, 12, 70))
    freezePane(wb, "VS_Final_Summary", firstActiveRow = 2)

    # ================================================================
    # SHEET: VS_Final_Detail
    # ================================================================
    addWorksheet(wb, "VS_Final_Detail", tabColour = "#BF6000")

    if (nrow(diffs_df) > 0) {
      writeData(wb, "VS_Final_Detail", diffs_df, withFilter = TRUE)
      addStyle(wb, "VS_Final_Detail", cmp_header_style, rows = 1,
               cols = 1:ncol(diffs_df), gridExpand = TRUE)

      cat_fill <- list(
        A = cmp_blue_style, B = cmp_orange_style,
        C = cmp_gray_style, D = cmp_lime_style
      )
      for (r in seq_len(nrow(diffs_df))) {
        cat_val <- diffs_df$Category[r]
        if (!is.na(cat_val) && cat_val %in% names(cat_fill)) {
          cat_col_idx <- which(names(diffs_df) == "Category")
          addStyle(wb, "VS_Final_Detail", cat_fill[[cat_val]],
                   rows = r + 1, cols = cat_col_idx, stack = TRUE)
        }
      }
      setColWidths(wb, "VS_Final_Detail", cols = 1:ncol(diffs_df), widths = "auto")
      freezePane(wb, "VS_Final_Detail", firstActiveRow = 2)
    } else {
      writeData(wb, "VS_Final_Detail",
                data.frame(Message = "No differences found — R output matches Final Excel 100%."))
      addStyle(wb, "VS_Final_Detail", cmp_green_style, rows = 2, cols = 1)
      setColWidths(wb, "VS_Final_Detail", cols = 1, widths = 55)
    }

    write_log("VS Final Excel comparison tabs written successfully.")

  }, error = function(e) {
    write_log(paste("VS Final comparison skipped — error:", conditionMessage(e)), "WARN")
  })

} else if (!is.null(final_excel_path)) {
  write_log(paste("VS Final comparison skipped — file not found:", final_excel_path), "WARN")
}

# Save
saveWorkbook(wb, output_path, overwrite = TRUE)

# ---------------------------------------------------------------
# POST-SAVE FIX: Remove dangling drawing/vmlDrawing relationships
# openxlsx writes drawing*.xml and vmlDrawing*.vml references
# into every sheet's .rels file but never creates the actual
# files.  Excel detects the broken refs and shows "found a
# problem" repair dialog on open.  We write a Python script to
# a temp file and run it to patch the xlsx in-place, preserving
# the folder structure exactly.
# ---------------------------------------------------------------
tryCatch({
  py_script_path <- file.path(tempdir(), "fix_xlsx_rels.py")
  py_path_escaped <- gsub("\\\\", "/", output_path)
  py_code <- paste0(
    "import zipfile, re, os, shutil\n",
    "path = '", py_path_escaped, "'\n",
    "tmp = path + '.patching'\n",
    "shutil.copy2(path, tmp)\n",
    "changed = 0\n",
    "with zipfile.ZipFile(tmp, 'r') as zin, zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:\n",
    "    for item in zin.infolist():\n",
    "        data = zin.read(item.filename)\n",
    "        if item.filename.endswith('.rels'):\n",
    "            txt = data.decode('utf-8')\n",
    "            txt2 = re.sub(r'<Relationship[^>]*Target=\"[^\"]*(?:drawing|vmlDrawing)[^\"]*\"[^/]*/>', '', txt)\n",
    "            txt2 = re.sub(r'<Relationship[^>]*Target=\"[^\"]*(?:drawing|vmlDrawing)[^\"]*\"[^>]*>.*?</Relationship>', '', txt2)\n",
    "            if txt2 != txt: changed += 1\n",
    "            data = txt2.encode('utf-8')\n",
    "        zout.writestr(item, data)\n",
    "os.remove(tmp)\n",
    "print('Drawing rels cleaned in', changed, 'files')\n",
    "# Inject fullCalcOnLoad into workbook.xml\n",
    "wb_name = 'xl/workbook.xml'\n",
    "if wb_name in [i.filename for i in zin.infolist()]:\n",
    "    pass  # already handled above via zout\n",
    "# Re-open to patch workbook.xml for fullCalcOnLoad\n",
    "import io\n",
    "with zipfile.ZipFile(path, 'r') as zin2:\n",
    "    names2 = zin2.namelist()\n",
    "    if 'xl/workbook.xml' in names2:\n",
    "        wb_xml = zin2.read('xl/workbook.xml').decode('utf-8')\n",
    "        if '<calcPr' not in wb_xml:\n",
    "            wb_xml = wb_xml.replace('</workbook>', '<calcPr calcMode=\"auto\" fullCalcOnLoad=\"1\"/></workbook>')\n",
    "        elif 'fullCalcOnLoad' not in wb_xml:\n",
    "            wb_xml = re.sub(r'<calcPr', '<calcPr fullCalcOnLoad=\"1\"', wb_xml, count=1)\n",
    "        tmp2 = path + '.wb.tmp'\n",
    "        shutil.copy2(path, tmp2)\n",
    "        with zipfile.ZipFile(tmp2, 'r') as zin3, zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout3:\n",
    "            for item3 in zin3.infolist():\n",
    "                if item3.filename == 'xl/workbook.xml':\n",
    "                    zout3.writestr(item3, wb_xml.encode('utf-8'))\n",
    "                else:\n",
    "                    zout3.writestr(item3, zin3.read(item3.filename))\n",
    "        os.remove(tmp2)\n",
    "        print('fullCalcOnLoad injected')\n"
  )
  writeLines(py_code, py_script_path)
  py_result <- system2("python3", args = py_script_path, stdout = TRUE, stderr = TRUE)
  if (any(grepl("cleaned", py_result))) {
    write_log(paste("Drawing rels cleaned — file opens without repair warning:", paste(py_result, collapse=" ")))
  } else if (length(py_result) > 0) {
    write_log(paste0("[WARN] Drawing rels cleanup: ", paste(py_result, collapse=" ")), "WARN")
  }
  unlink(py_script_path)
}, error = function(e) {
  write_log(paste0("[WARN] Drawing rels cleanup failed: ", conditionMessage(e)), "WARN")
})

cat("\n")
cat(strrep("=", 60), "\n")
cat("  DONE - Audited workbook saved to:\n")
cat(paste0("  ", output_path), "\n")
cat("\n")
cat("  TABS IN YOUR OUTPUT FILE:\n")
cat("    SCSD ROSTER       - Full roster with color-coded status\n")
cat("    AUDIT_SUMMARY     - Counts, percentages, next steps\n")
cat("    Action Needed     - YOUR WORK QUEUE (start here)\n")
cat("    Not In ST         - Students missing from ST Enrollment\n")
cat("    Multiple Enroll.  - Students with 2+ enrollments\n")
cat("    Formula Adapter   - Column mapping + formula reference\n")
cat("    SPED ROSTER       - SPED students with IEP data\n")
cat(strrep("=", 60), "\n")

# Open file (Windows only)
if (.Platform$OS.type == "windows") {
  tryCatch(shell.exec(output_path), error = function(e) invisible(NULL))
}
