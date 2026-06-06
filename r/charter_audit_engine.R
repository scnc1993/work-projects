# ============================================================
# CHARTER WORKSHEET DATA VALIDATION ENGINE  (R)
# SPED / SCSD Roster vs SchoolTool Audit Engine
# ------------------------------------------------------------
# Layout-agnostic: auto sheet detection, auto header-row
# detection, fuzzy column matching, flexible date parsing.
# Produces the 13 calculated audit fields and writes a
# multi-tab .xlsx report with severity highlighting.
#
# Mirrors sql/charter_audit_engine.sql and the pandas version
# in python/charter_audit_engine.py.
#
# Join model: rosters drive the audit (roster LEFT JOIN
# SchoolTool on NYSSIS, both cast to text). StudentID columns
# are validation inputs, NOT join keys.
#
# Author: scnc1993 / Thrivora Holdings LLC
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(tibble)
  library(openxlsx)
})

.charter_engine_version <- "Charter Worksheet Data Validation Engine (R) v2026.06.06"
cat("\n", .charter_engine_version, "loaded\n")

# ============================================================
# 1. AUTO SHEET DETECTION (SPED + SCSD + GEN ROSTER)
# ============================================================
detect_roster_sheet <- function(path) {
  sheets <- excel_sheets(path)
  patterns <- c(
    "^SPED ROSTER$", "SPED ROSTER", "SCSD SPED ROSTER", "SPED",
    "^SCSD ROSTER$", "SCSD ROSTER", "SCSD GEN ED", "SCSD",
    "^GEN ROSTER$", "GEN ROSTER", "GENERAL ED ROSTER", "GEN ED ROSTER",
    "ROSTER", "ENROLL", "STUDENT", "INVOICE"
  )
  for (p in patterns) {
    hit <- sheets[str_detect(sheets, regex(p, ignore_case = TRUE))]
    if (length(hit) > 0) return(hit[1])
  }
  sheets[1]
}

# ============================================================
# 2. AUTO HEADER-ROW DETECTION
# ============================================================
detect_header_row <- function(df_raw) {
  for (i in 1:min(10, nrow(df_raw))) {
    if (any(!is.na(df_raw[i, ]))) return(i)
  }
  1
}

# ============================================================
# 3. FUZZY COLUMN DETECTOR
# ============================================================
find_col <- function(df, patterns) {
  cols <- names(df)
  for (p in patterns) {
    hit <- cols[str_detect(cols, regex(p, ignore_case = TRUE))]
    if (length(hit) > 0) return(hit[1])
  }
  NA_character_
}

# ============================================================
# 4. NAME NORMALIZER
# ============================================================
normalize_name <- function(x) {
  x %>% toupper() %>% gsub("[^A-Z ]", "", .) %>% str_squish()
}

# ============================================================
# 5. DOB NORMALIZER (Excel serial, YYYYMMDD, SAS 14FEB2026, many formats)
# ============================================================
normalize_dob <- function(x) {
  if (is.null(x) || length(x) == 0) return(NA_character_)
  x <- toupper(trimws(as.character(x)))
  if (x == "" || x %in% c("NA", "N/A", "NULL")) return(NA_character_)

  suppressWarnings({
    num <- as.numeric(x)
    if (!is.na(num) && num > 20000 && num < 60000) {
      d <- as.Date(num, origin = "1899-12-30")
      if (!is.na(d)) return(as.character(d))
    }
  })
  if (grepl("^\\d{8}$", x)) {
    d <- suppressWarnings(as.Date(x, "%Y%m%d")); if (!is.na(d)) return(as.character(d))
  }
  if (grepl("^\\d{1,2}[A-Z]{3}\\d{4}$", x)) {
    d <- suppressWarnings(as.Date(x, "%d%b%Y")); if (!is.na(d)) return(as.character(d))
  }
  formats <- c("%m/%d/%Y","%Y-%m-%d","%m-%d-%Y","%m/%d/%y","%m-%d-%y",
               "%d-%b-%Y","%d-%b-%y","%d/%m/%Y","%Y/%m/%d","%Y.%m.%d")
  for (f in formats) {
    d <- suppressWarnings(as.Date(x, format = f)); if (!is.na(d)) return(as.character(d))
  }
  NA_character_
}

# ============================================================
# 6. MATERIALITY / VALUE NORMALIZER
#    Trim, uppercase, collapse spaces, strip trailing .0,
#    keep leading zeros, whitelist [^0-9A-Z/: -], parse dates.
# ============================================================
normalize_value <- function(x) {
  if (is.null(x) || length(x) == 0) return(NA_character_)
  x <- toupper(trimws(as.character(x)))
  if (x == "" || x %in% c("NA", "N/A", "NULL")) return(NA_character_)
  x <- gsub("\\s+", " ", x)
  x <- gsub("\\.0+$", "", x)
  x <- gsub("[^0-9A-Z/: \\-]", "", x)
  NA_or <- function(v) if (length(v) == 0 || all(is.na(v))) NA else v

  dt <- suppressWarnings(tryCatch(
    as.POSIXct(x, tz = "UTC", tryFormats = c(
      "%m/%d/%Y %H:%M:%S","%m/%d/%Y %H:%M","%Y-%m-%d %H:%M:%S","%Y-%m-%d %H:%M")),
    error = function(e) NA))
  if (!all(is.na(dt))) return(as.character(as.Date(dt)))

  d <- suppressWarnings(as.Date(x, tryFormats = c(
    "%m/%d/%Y","%Y-%m-%d","%m-%d-%Y","%m/%d/%y","%m-%d-%y")))
  if (!all(is.na(d))) return(as.character(d))

  if (grepl("^[0-9]+$", x)) {
    num <- suppressWarnings(as.numeric(x))
    if (!is.na(num) && num > 20000 && num < 60000) {
      d2 <- suppressWarnings(as.Date(num, origin = "1899-12-30"))
      if (!is.na(d2)) return(as.character(d2))
    }
    return(as.character(num))
  }
  x
}

# vectorized helpers
v_norm <- function(col) vapply(col, normalize_value, character(1), USE.NAMES = FALSE)
v_dob  <- function(col) vapply(col, normalize_dob,   character(1), USE.NAMES = FALSE)

# safe date diff in days (roster - schooltool); NA if either missing
diff_days <- function(a, b) {
  da <- suppressWarnings(as.Date(a)); db <- suppressWarnings(as.Date(b))
  as.integer(da - db)
}

# ============================================================
# 7. READ A ROSTER (auto sheet + header, all text)
# ============================================================
read_roster <- function(path, sheet = NULL) {
  if (is.null(sheet)) sheet <- detect_roster_sheet(path)
  raw <- read_excel(path, sheet = sheet, col_names = FALSE)
  hdr <- detect_header_row(raw)
  read_excel(path, sheet = sheet, skip = hdr - 1, col_types = "text")
}

# ============================================================
# 8. CORE AUDIT — produces the 13 calculated fields
# ------------------------------------------------------------
# roster_df : the charter/SPED roster (drives the audit)
# st_df     : SchoolTool export
# school_label : e.g. "Southside", "SASCCS", "SAS", "OnTech"
# ============================================================
charter_audit <- function(roster_df, st_df, school_label = "Roster",
                          tolerance_days = 5) {

  # ---- detect columns in roster ----
  r_ny   <- find_col(roster_df, c("NYSSIS","NYS ?ID","STATE ?ID","UNIVERSAL"))
  r_sid  <- find_col(roster_df, c("STUDENT.?ID","LOCAL.?ID","STUDENT NUMBER"))
  r_dis  <- find_col(roster_df, c("DISABILITY","CLASSIFICATION","IEP","SPED"))
  r_ent  <- find_col(roster_df, c("ENTRY","ENROLL","START"))
  r_exit <- find_col(roster_df, c("EXIT","WITHDRAW","DISENROLL","END"))

  # ---- detect columns in SchoolTool ----
  s_ny   <- find_col(st_df, c("NYSSIS","NYS ?ID","STATE ?ID","UNIVERSAL"))
  s_sid  <- find_col(st_df, c("STUDENT.?ID","LOCAL.?ID","STUDENT NUMBER"))
  s_uid  <- find_col(st_df, c("UNIVERSAL.?STUDENT","STUDENT_?UNIVERSAL"))
  s_dis  <- find_col(st_df, c("DISABILITY","CLASSIFICATION","IEP","SPED"))
  s_ent  <- find_col(st_df, c("ENTRY","ENROLL","START"))
  s_exit <- find_col(st_df, c("EXIT","WITHDRAW","DISENROLL","END"))

  get <- function(df, col) if (!is.na(col)) df[[col]] else rep(NA_character_, nrow(df))

  # ---- build text join keys on NYSSIS (both sides) ----
  st <- st_df %>% mutate(
    .st_ny   = v_norm(get(st_df, s_ny)),
    .st_sid  = v_norm(get(st_df, s_sid)),
    .st_uid  = v_norm(get(st_df, s_uid)),
    .st_dis  = get(st_df, s_dis),
    .st_ent  = v_norm(get(st_df, s_ent)),
    .st_exit = v_norm(get(st_df, s_exit))
  ) %>% select(starts_with(".st_"))

  r <- roster_df %>% mutate(
    .r_ny   = v_norm(get(roster_df, r_ny)),
    .r_sid  = v_norm(get(roster_df, r_sid)),
    .r_dis  = get(roster_df, r_dis),
    .r_ent  = v_norm(get(roster_df, r_ent)),
    .r_exit = v_norm(get(roster_df, r_exit))
  )

  # roster-driven LEFT JOIN on NYSSIS (text)
  j <- r %>% left_join(st, by = c(".r_ny" = ".st_ny"), keep = TRUE)

  out <- j %>% mutate(
    # 1. resolved_studentid — COALESCE waterfall
    resolved_studentid = dplyr::coalesce(
      dplyr::na_if(.st_sid, ""),
      dplyr::na_if(.st_uid, ""),
      dplyr::na_if(.r_sid, ""),
      ""
    ),

    # 3. id_source
    id_source = dplyr::case_when(
      !is.na(.st_sid) & .st_sid != "" ~ "SchoolTool StudentID",
      !is.na(.st_uid) & .st_uid != "" ~ "Universal StudentID",
      !is.na(.r_sid)  & .r_sid  != "" ~ "NYSSID Mapping",
      TRUE ~ "No Match"
    ),

    # 2. id_match_status
    id_match_status = dplyr::case_when(
      is.na(.st_ny) ~ "No Roster Match",
      resolved_studentid == "" ~ "Missing Student ID Match",
      TRUE ~ paste0("Matched - ", school_label)
    ),

    # 4. is_iep_flag
    is_iep_flag = ifelse(is.na(.r_dis) | trimws(.r_dis) == "", "Gen Ed", "IEP"),
    .st_iep = ifelse(is.na(.st_dis) | trimws(.st_dis) == "", "Gen Ed", "IEP"),

    # 5. roster_vs_disability_check
    roster_vs_disability_check = dplyr::case_when(
      is.na(.st_ny) ~ "No SchoolTool Record",
      is_iep_flag == "IEP"   & .st_iep == "IEP"    ~ "Consistent SPED",
      is_iep_flag == "Gen Ed"& .st_iep == "Gen Ed" ~ "Consistent Gen Ed",
      is_iep_flag == "IEP"   & .st_iep == "Gen Ed" ~ "Mismatch - Roster SPED / ST Gen Ed",
      is_iep_flag == "Gen Ed"& .st_iep == "IEP"    ~ "Mismatch - Roster Gen Ed / ST SPED",
      TRUE ~ "Unknown"
    ),

    # 6. entry_diff_days  (roster entry - ST entry)
    entry_diff_days = diff_days(.r_ent, .st_ent),

    # 7. entry_status_category
    entry_status_category = dplyr::case_when(
      is.na(.r_ent) | .r_ent == "" | is.na(.st_ent) | .st_ent == "" ~ "Missing Date",
      entry_diff_days == 0 ~ "Exact Match",
      abs(entry_diff_days) <= tolerance_days ~ "Within Tolerance",
      TRUE ~ "Finance Risk"
    ),

    # 8. exit_diff_days  (roster exit - ST exit)
    exit_diff_days = diff_days(.r_exit, .st_exit),

    # 9. exit_status_category
    exit_status_category = dplyr::case_when(
      is.na(.r_exit) | .r_exit == "" | is.na(.st_exit) | .st_exit == "" ~ "Missing Date",
      exit_diff_days == 0 ~ "Exact Match",
      abs(exit_diff_days) <= tolerance_days ~ "Within Tolerance",
      TRUE ~ "Finance Risk"
    ),

    # 10. exit_issue_type
    exit_issue_type = dplyr::case_when(
      (is.na(.r_exit)|.r_exit=="") & (is.na(.st_exit)|.st_exit=="") ~ "Both Dates Missing",
      (is.na(.r_exit)|.r_exit=="") ~ "Missing Roster Exit",
      (is.na(.st_exit)|.st_exit=="") ~ "Missing ST Exit",
      !is.na(exit_diff_days) & exit_diff_days >  tolerance_days ~ "Roster Exit Late (>5)",
      !is.na(exit_diff_days) & exit_diff_days < -tolerance_days ~ "Roster Exit Early (<-5)",
      !is.na(exit_diff_days) & abs(exit_diff_days) <= tolerance_days ~ "Within Tolerance",
      TRUE ~ "No Issue"
    ),

    # 11. final_validation_status
    final_validation_status = dplyr::case_when(
      id_match_status == "No Roster Match" | resolved_studentid == "" ~ "Invalid - No ID Match",
      entry_status_category == "Missing Date" | exit_status_category == "Missing Date" ~ "Incomplete - Missing Dates",
      entry_status_category == "Finance Risk" | exit_status_category == "Finance Risk" ~ "Invalid - Date Mismatch",
      roster_vs_disability_check %in% c("Mismatch - Roster SPED / ST Gen Ed",
                                        "Mismatch - Roster Gen Ed / ST SPED") ~ "Valid with Warning",
      entry_status_category == "Within Tolerance" | exit_status_category == "Within Tolerance" ~ "Valid with Warning",
      TRUE ~ "Valid"
    ),

    # 12. severity
    severity = dplyr::case_when(
      final_validation_status %in% c("Invalid - No ID Match", "Invalid - Date Mismatch") ~ "Critical",
      final_validation_status == "Incomplete - Missing Dates" ~ "High",
      final_validation_status == "Valid with Warning" ~ "Low",
      TRUE ~ "None"
    ),

    # 13. recommended_action
    recommended_action = dplyr::case_when(
      id_match_status == "No Roster Match" ~ "Locate student in SchoolTool; confirm NYSSIS on roster.",
      resolved_studentid == "" ~ "Resolve Student ID via NYSSIS mapping or Universal ID.",
      exit_issue_type == "Missing Roster Exit" ~ "Add exit/withdrawal date to roster.",
      exit_issue_type == "Missing ST Exit" ~ "Enter exit date in SchoolTool.",
      exit_issue_type == "Both Dates Missing" ~ "Verify enrollment status; both exit dates absent.",
      exit_issue_type == "Roster Exit Late (>5)" ~ "Reconcile roster exit date — billed beyond ST exit.",
      exit_issue_type == "Roster Exit Early (<-5)" ~ "Reconcile roster exit date — earlier than ST exit.",
      entry_status_category == "Finance Risk" ~ "Reconcile entry dates — FTE/billing risk.",
      entry_status_category == "Missing Date" ~ "Add missing entry/enrollment date.",
      roster_vs_disability_check == "Mismatch - Roster SPED / ST Gen Ed" ~ "Confirm IEP classification; roster says SPED.",
      roster_vs_disability_check == "Mismatch - Roster Gen Ed / ST SPED" ~ "Confirm IEP classification; ST says SPED.",
      TRUE ~ "No action — record valid."
    ),

    school = school_label
  )

  # keep original roster cols + the 13 audit fields
  audit_cols <- c("school","resolved_studentid","id_match_status","id_source",
                  "is_iep_flag","roster_vs_disability_check",
                  "entry_diff_days","entry_status_category",
                  "exit_diff_days","exit_status_category","exit_issue_type",
                  "final_validation_status","severity","recommended_action")
  orig_cols <- setdiff(names(roster_df), names(out))
  out %>% select(any_of(c(names(roster_df), audit_cols)))
}

# ============================================================
# 9. EXPORT — multi-tab workbook with severity highlighting
# ============================================================
export_audit <- function(audit_df, out_file = "CHARTER_AUDIT_REPORT.xlsx") {
  wb <- createWorkbook()

  # --- Full Audit ---
  addWorksheet(wb, "Full_Audit")
  writeData(wb, "Full_Audit", audit_df)
  freezePane(wb, "Full_Audit", firstActiveRow = 2, firstActiveCol = 1)
  setColWidths(wb, "Full_Audit", cols = 1:ncol(audit_df), widths = "auto")

  # severity color styles
  if ("severity" %in% names(audit_df)) {
    sev_col <- which(names(audit_df) == "severity")
    styles <- list(
      Critical = createStyle(bgFill = "#F4CCCC"),
      High     = createStyle(bgFill = "#FCE5CD"),
      Low      = createStyle(bgFill = "#FFF2CC")
    )
    for (lvl in names(styles)) {
      rows <- which(audit_df$severity == lvl)
      if (length(rows) > 0) {
        addStyle(wb, "Full_Audit", styles[[lvl]],
                 rows = rows + 1, cols = 1:ncol(audit_df),
                 gridExpand = TRUE, stack = TRUE)
      }
    }
  }

  # --- Exceptions only ---
  exc <- audit_df %>% filter(severity %in% c("Critical","High","Low"))
  addWorksheet(wb, "Exceptions")
  writeData(wb, "Exceptions", exc)
  freezePane(wb, "Exceptions", firstActiveRow = 2, firstActiveCol = 1)

  # --- Severity Summary ---
  summ <- audit_df %>% count(severity, final_validation_status, name = "count") %>%
    arrange(desc(count))
  addWorksheet(wb, "Severity_Summary")
  writeData(wb, "Severity_Summary", summ)

  saveWorkbook(wb, out_file, overwrite = TRUE)
  cat("Audit report saved to:", out_file, "\n")
}

# ============================================================
# 10. MAIN WRAPPER
# ============================================================
run_charter_audit <- function(roster_file = NULL,
                              schooltool_file = NULL,
                              school_label = "Roster",
                              out_file = "CHARTER_AUDIT_REPORT.xlsx") {
  if (is.null(roster_file))     roster_file     <- file.choose()
  if (is.null(schooltool_file)) schooltool_file <- file.choose()

  roster_df <- read_roster(roster_file)
  st_df     <- read_roster(schooltool_file)

  audit_df <- charter_audit(roster_df, st_df, school_label = school_label)
  export_audit(audit_df, out_file)

  cat("\n===== CHARTER AUDIT COMPLETE =====\n")
  cat("Rows audited:        ", nrow(audit_df), "\n")
  cat("Critical:            ", sum(audit_df$severity == "Critical"), "\n")
  cat("High:                ", sum(audit_df$severity == "High"), "\n")
  cat("Low:                 ", sum(audit_df$severity == "Low"), "\n")
  cat("Valid (None):        ", sum(audit_df$severity == "None"), "\n")
  invisible(audit_df)
}

# ------------------------------------------------------------
# To run:
#   source("charter_audit_engine.R")
#   run_charter_audit("roster.xlsx", "schooltool.xlsx",
#                     school_label = "Southside")
# ------------------------------------------------------------
