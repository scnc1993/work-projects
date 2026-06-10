# ============================================================================
# Build Unified Charter Roster — Multi-Tab Workbook (25-26)
# Daquan Morrison · Data Analyst · Syracuse City School District
# ----------------------------------------------------------------------------
# SOURCE FILES (place all in the same folder as this script):
#   southside_291462_unified-roster.xlsx
#   sas_291463_unified-roster.xlsx
#   sasccs_291464_unified-roster.xlsx
#   ontech_final_unified-roster.xlsx
#
# HOW TO RUN ON YOUR DESKTOP:
#   1. Put this script + all 4 xlsx files in the same folder
#   2. Open RStudio, set working directory to that folder
#      (Session > Set Working Directory > To Source File Location)
#   3. Run: Rscript build_unified_charter_roster_tabs.R
#
# OUTPUT: all_charters_unified_roster_25-26_REBUILT.xlsx (5 tabs)
#   Tab 1 — All Charters Unified     : master combined data
#   Tab 2 — Formula Version          : NYSSIS lookup formulas vs Update Log
#   Tab 3 — Copy & Paste             : plain values, paste into FY 25-26 master
#   Tab 4 — Not Found in ST          : students with no NYSSIS / not in ST
#   Tab 5 — Update Log               : drop resolved NYSSIS records here (R/Python)
# ============================================================================

pkgs_needed  <- c("readxl","dplyr","stringr","lubridate","openxlsx","tibble")
pkgs_missing <- pkgs_needed[!pkgs_needed %in% rownames(installed.packages())]
if (length(pkgs_missing) > 0) install.packages(pkgs_missing, dependencies = TRUE)
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr)
  library(lubridate); library(openxlsx); library(tibble)
})

# ============================================================================
# SETTINGS — only change data_dir if files are in a different folder
# ============================================================================
# Set to "." to use the same folder as the script (default for desktop use)
data_dir   <- "."
output_path <- file.path(data_dir, "all_charters_unified_roster_25-26_REBUILT.xlsx")
INVOICE_MONTH <- as.Date("2026-06-01")

# Exact source file names — do not rename the xlsx files
FILE_SOUTHSIDE <- file.path(data_dir, "southside_291462_unified-roster.xlsx")
FILE_SAS       <- file.path(data_dir, "sas_291463_unified-roster.xlsx")
FILE_SASCCS    <- file.path(data_dir, "sasccs_291464_unified-roster.xlsx")
FILE_ONTECH    <- file.path(data_dir, "ontech_final_unified-roster.xlsx")

# Verify all files exist before proceeding
for (f in c(FILE_SOUTHSIDE, FILE_SAS, FILE_SASCCS, FILE_ONTECH)) {
  if (!file.exists(f)) stop(paste("ERROR: File not found ->", f,
    "\nMake sure all 4 xlsx files are in the same folder as this script."))
}

# ============================================================================
# HELPERS
# ============================================================================
parse_mixed_date <- function(x) {
  if (inherits(x, "Date"))                      return(as.Date(x))
  if (inherits(x, c("POSIXct","POSIXlt")))      return(as.Date(x))
  x_chr <- str_trim(as.character(x))
  x_chr[x_chr %in% c("","NA","N/A","null","NULL","0")] <- NA_character_
  suppressWarnings(x_num <- as.numeric(x_chr))
  out <- rep(as.Date(NA), length(x_chr))
  # Excel serial: valid 2020-2030 serials are 43831-49006 — use > 40000 threshold
  idx_serial <- !is.na(x_num) & x_num > 40000 & x_num < 80000
  if (any(idx_serial)) out[idx_serial] <- as.Date(x_num[idx_serial], origin = "1899-12-30")
  idx_text <- !idx_serial & !is.na(x_chr)
  if (any(idx_text)) {
    d1 <- suppressWarnings(mdy(x_chr[idx_text]))
    d2 <- suppressWarnings(ymd(x_chr[idx_text]))
    d3 <- suppressWarnings(dmy(x_chr[idx_text]))
    out[idx_text] <- coalesce(d1, d2, d3)
  }
  out
}

clean_id <- function(x) {
  s <- str_trim(as.character(x))
  s <- str_replace(s, "\\.0+$", "")
  s[s %in% c("","NA","NaN","NULL")] <- NA_character_
  s
}

clean_fte <- function(x) {
  s <- str_trim(as.character(x))
  s <- str_replace_all(s, "%", "")
  suppressWarnings(as.numeric(s))
}

split_full_name <- function(x) {
  s <- str_trim(as.character(x))
  last  <- str_trim(str_extract(s, "^[^,]+"))
  first <- str_trim(str_replace(s, "^[^,]+,?\\s*", ""))
  tibble(last = last, first = first)
}

# ============================================================================
# BUILD EACH CHARTER FROM UNIFIED GEN-SPED ROSTER TAB
# ============================================================================

# ---------- SOUTHSIDE 291462 ------------------------------------------------
build_southside <- function() {
  message("  Reading southside_291462_unified-roster.xlsx ...")
  df <- read_excel(FILE_SOUTHSIDE, sheet = "Unified Gen-SPED Roster",
                   col_names = TRUE, .name_repair = "minimal")
  df <- df[!is.na(df[["NYSSIS ID"]]) | !is.na(df[["Student Last Name"]]), ]
  tibble(
    `Charter_Source Ticket #` = "Southside 291462",
    `Charter Month`           = INVOICE_MONTH,
    `Source_File`             = "southside_291462_unified-roster.xlsx",
    `NYSSIS ID`               = clean_id(df[["NYSSIS ID"]]),
    `Student ID`              = clean_id(df[["Student ID"]]),
    `Last Name`               = str_trim(as.character(df[["Student Last Name"]])),
    `First Name`              = str_trim(as.character(df[["Student First Name"]])),
    `Grade Level`             = str_trim(as.character(df[["Grade"]])),
    `DOB`                     = parse_mixed_date(df[["Birth Date"]]),
    `School Name`             = str_trim(as.character(df[["School Name"]])),
    `Enroll Date`             = parse_mixed_date(df[["Entered"]]),
    `Withdraw Date`           = parse_mixed_date(df[["Withdrew"]]),
    `SCSD FTE`                = clean_fte(df[["Southside FTE"]]),
    `Student Status`          = str_trim(as.character(df[["Student Status"]])),
    `Notes for Patrick`       = str_trim(as.character(df[["Notes for Patrick"]])),
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ---------- SAS 291463 -------------------------------------------------------
build_sas <- function() {
  message("  Reading sas_291463_unified-roster.xlsx ...")
  df <- read_excel(FILE_SAS, sheet = "Unified Gen-SPED Roster",
                   col_names = TRUE, .name_repair = "minimal")
  df <- df[!is.na(df[["NYSSIS ID"]]) | !is.na(df[["Last Name"]]), ]
  tibble(
    `Charter_Source Ticket #` = "SAS 291463",
    `Charter Month`           = INVOICE_MONTH,
    `Source_File`             = "sas_291463_unified-roster.xlsx",
    `NYSSIS ID`               = clean_id(df[["NYSSIS ID"]]),
    `Student ID`              = clean_id(df[["Student ID"]]),
    `Last Name`               = str_trim(as.character(df[["Last Name"]])),
    `First Name`              = str_trim(as.character(df[["First Name"]])),
    `Grade Level`             = str_trim(as.character(df[["Current Grade Level"]])),
    `DOB`                     = parse_mixed_date(df[["Date of Birth"]]),
    `School Name`             = str_trim(as.character(df[["School Name"]])),
    `Enroll Date`             = parse_mixed_date(df[["Entry Date"]]),
    `Withdraw Date`           = parse_mixed_date(df[["Leave Date"]]),
    `SCSD FTE`                = clean_fte(df[["FTE"]]),
    `Student Status`          = str_trim(as.character(df[["Student Status"]])),
    `Notes for Patrick`       = NA_character_,
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ---------- SASCCS 291464 ----------------------------------------------------
build_sasccs <- function() {
  message("  Reading sasccs_291464_unified-roster.xlsx ...")
  df <- read_excel(FILE_SASCCS, sheet = "Unified Gen-SPED Roster",
                   col_names = TRUE, .name_repair = "minimal")
  df <- df[!is.na(df[[1]]) | !is.na(df[[3]]), ]
  nm <- split_full_name(df[["Full Name"]])
  tibble(
    `Charter_Source Ticket #` = "SASCCS 291464",
    `Charter Month`           = INVOICE_MONTH,
    `Source_File`             = "sasccs_291464_unified-roster.xlsx",
    `NYSSIS ID`               = clean_id(df[["NYSSIS ID"]]),
    `Student ID`              = clean_id(df[["Student ID"]]),
    `Last Name`               = nm$last,
    `First Name`              = nm$first,
    `Grade Level`             = str_trim(as.character(df[["Current Grade Level"]])),
    `DOB`                     = parse_mixed_date(df[["Date of Birth"]]),
    `School Name`             = str_trim(as.character(df[["School Name"]])),
    `Enroll Date`             = parse_mixed_date(df[["Entry Date"]]),
    `Withdraw Date`           = parse_mixed_date(df[["Leave Date"]]),
    `SCSD FTE`                = clean_fte(df[["Status"]]),
    `Student Status`          = str_trim(as.character(df[["Student Status"]])),
    `Notes for Patrick`       = NA_character_,
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ---------- ONTECH 291465 ----------------------------------------------------
build_ontech <- function() {
  message("  Reading ontech_final_unified-roster.xlsx ...")
  df <- read_excel(FILE_ONTECH, sheet = "Unified Gen-SPED Roster",
                   col_names = TRUE, .name_repair = "minimal")
  df <- df[!is.na(df[["NYSSIS ID"]]) | !is.na(df[["Last Name"]]), ]
  sid_col <- if ("Studemt ID" %in% names(df)) "Studemt ID" else "Student ID"
  tibble(
    `Charter_Source Ticket #` = "OnTech 291465",
    `Charter Month`           = INVOICE_MONTH,
    `Source_File`             = "ontech_final_unified-roster.xlsx",
    `NYSSIS ID`               = clean_id(df[["NYSSIS ID"]]),
    `Student ID`              = clean_id(df[[sid_col]]),
    `Last Name`               = str_trim(as.character(df[["Last Name"]])),
    `First Name`              = str_trim(as.character(df[["First Name"]])),
    `Grade Level`             = str_trim(as.character(df[["Grade Level"]])),
    `DOB`                     = parse_mixed_date(df[["DOB"]]),
    `School Name`             = "OnTECH Charter High School",
    `Enroll Date`             = parse_mixed_date(df[["Enroll Date"]]),
    `Withdraw Date`           = parse_mixed_date(df[["Withdraw Date"]]),
    `SCSD FTE`                = clean_fte(df[["OnTECH FTE"]]),
    `Student Status`          = NA_character_,
    `Notes for Patrick`       = NA_character_,
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ============================================================================
# COMBINE ALL 4 CHARTERS
# ============================================================================
message("\nBuilding unified roster...")
unified <- bind_rows(
  build_southside(),
  build_sas(),
  build_sasccs(),
  build_ontech()
)

template_cols <- c(
  "Charter_Source Ticket #", "Charter Month", "Source_File", "NYSSIS ID",
  "Student ID", "Last Name", "First Name", "Grade Level", "DOB",
  "School Name", "Enroll Date", "Withdraw Date", "SCSD FTE",
  "Student Status", "Notes for Patrick", "SPED_Flag"
)
unified <- unified %>%
  select(all_of(template_cols)) %>%
  mutate(across(where(is.character),
    ~ifelse(.x %in% c("NA","N/A","NULL","null",""), NA_character_, .x)))

# ============================================================================
# STEP: PATCH STUDENT IDs AND NYSSIS IDs FROM ST ENROLL REPORTS
# Pass 1 — patch missing Student IDs  (NYSSIS/name/DOB → Student ID)
# Pass 2 — patch missing NYSSIS IDs   (Student ID/name/DOB → NYSSIS)
# This ensures SPED-only rows that have a Student ID but no NYSSIS get resolved.
# ============================================================================
message("\nPatching Student IDs and NYSSIS IDs from ST Enroll reports...")

mk_st <- function(df, src) {
  data.frame(
    nyssis = sub("\\.0+$","", str_trim(as.character(df[["Student_UniversalStudentID"]]))),
    sid    = sub("\\.0+$","", str_trim(as.character(df[["StudentID"]]))),
    fname  = str_trim(toupper(as.character(df[["FirstName"]]))),
    lname  = str_trim(toupper(as.character(df[["LastName"]]))),
    dob    = substr(as.character(df[["Person_DOB"]]),1,10),
    source = src, stringsAsFactors=FALSE
  )
}

st_all <- bind_rows(
  mk_st(read_excel(FILE_SOUTHSIDE, sheet="ST Enroll Export", .name_repair="minimal"), "Southside ST"),
  mk_st(read_excel(FILE_SAS,       sheet="ST Enroll Export", .name_repair="minimal"), "SAS ST"),
  mk_st(read_excel(FILE_SASCCS,    sheet="ST Enroll Export", .name_repair="minimal"), "SASCCS ST"),
  mk_st(read_excel(FILE_ONTECH,    sheet="ST Enroll Export", .name_repair="minimal"), "OnTech ST")
)
st_all$nyssis[st_all$nyssis %in% c("NA","NULL","","0")] <- NA
st_all$sid   [st_all$sid    %in% c("NA","NULL","","0")] <- NA
st_all$fname [st_all$fname  %in% c("NA","NULL","")]     <- NA
st_all$lname [st_all$lname  %in% c("NA","NULL","")]     <- NA

safe_nchar <- function(x) ifelse(is.na(x), 0L, nchar(x))

patch_count <- 0L
for (i in seq_len(nrow(unified))) {
  # Only patch rows that still have no Student ID
  cur_sid <- unified$`Student ID`[i]
  if (!is.na(cur_sid) && cur_sid != "0") next

  m_nyssis <- unified$`NYSSIS ID`[i]
  m_fname  <- str_trim(toupper(as.character(unified$`First Name`[i])))
  m_lname  <- str_trim(toupper(as.character(unified$`Last Name`[i])))
  m_dob    <- substr(as.character(unified$DOB[i]),1,10)
  if (m_fname %in% c("NA","NULL","")) m_fname <- NA_character_
  if (m_lname %in% c("NA","NULL","")) m_lname <- NA_character_
  if (m_dob   %in% c("NA","NULL","")) m_dob   <- NA_character_

  found_sid <- NA_character_

  # Try 1: NYSSIS match
  if (!is.na(m_nyssis)) {
    idx <- which(!is.na(st_all$nyssis) & st_all$nyssis == m_nyssis & !is.na(st_all$sid))
    if (length(idx)>0) found_sid <- st_all$sid[idx[1]]
  }

  # Try 2: First + Last name
  if (is.na(found_sid) && !is.na(m_fname) && !is.na(m_lname) &&
      safe_nchar(m_fname)>1 && safe_nchar(m_lname)>1) {
    idx <- which(!is.na(st_all$fname) & !is.na(st_all$lname) &
                 st_all$fname==m_fname & st_all$lname==m_lname & !is.na(st_all$sid))
    if (length(idx)>0) found_sid <- st_all$sid[idx[1]]
  }

  # Try 3: Last name + DOB
  if (is.na(found_sid) && !is.na(m_lname) && !is.na(m_dob) && safe_nchar(m_lname)>1) {
    idx <- which(!is.na(st_all$lname) & !is.na(st_all$dob) &
                 st_all$lname==m_lname & st_all$dob==m_dob & !is.na(st_all$sid))
    if (length(idx)>0) found_sid <- st_all$sid[idx[1]]
  }

  # Try 4: First name + DOB (unique match only)
  if (is.na(found_sid) && !is.na(m_fname) && !is.na(m_dob) && safe_nchar(m_fname)>1) {
    idx <- which(!is.na(st_all$fname) & !is.na(st_all$dob) &
                 st_all$fname==m_fname & st_all$dob==m_dob & !is.na(st_all$sid))
    if (length(idx)==1) found_sid <- st_all$sid[idx[1]]
  }

  if (!is.na(found_sid)) {
    unified$`Student ID`[i] <- found_sid
    patch_count <- patch_count + 1L
  }
}

cat(sprintf("  Pass 1 — Student IDs patched from ST: %d\n", patch_count))
cat(sprintf("  Still missing Student IDs:            %d\n",
  sum(is.na(unified$`Student ID`) | unified$`Student ID`=="0")))

# ============================================================================
# PASS 2 — Patch missing NYSSIS IDs from ST using Student ID / name / DOB
# ============================================================================
nyssis_patch_count <- 0L
for (i in seq_len(nrow(unified))) {
  # Only patch rows that still have no NYSSIS
  cur_nyssis <- unified$`NYSSIS ID`[i]
  if (!is.na(cur_nyssis) && cur_nyssis != "0") next

  m_sid   <- unified$`Student ID`[i]
  m_fname <- str_trim(toupper(as.character(unified$`First Name`[i])))
  m_lname <- str_trim(toupper(as.character(unified$`Last Name`[i])))
  m_dob   <- substr(as.character(unified$DOB[i]),1,10)
  if (m_fname %in% c("NA","NULL","")) m_fname <- NA_character_
  if (m_lname %in% c("NA","NULL","")) m_lname <- NA_character_
  if (m_dob   %in% c("NA","NULL","")) m_dob   <- NA_character_

  found_nyssis <- NA_character_

  # Try 1: Student ID → NYSSIS
  if (!is.na(m_sid) && m_sid != "0") {
    idx <- which(!is.na(st_all$sid) & st_all$sid == m_sid & !is.na(st_all$nyssis))
    if (length(idx)>0) found_nyssis <- st_all$nyssis[idx[1]]
  }

  # Try 2: First + Last name → NYSSIS
  if (is.na(found_nyssis) && !is.na(m_fname) && !is.na(m_lname) &&
      safe_nchar(m_fname)>1 && safe_nchar(m_lname)>1) {
    idx <- which(!is.na(st_all$fname) & !is.na(st_all$lname) &
                 st_all$fname==m_fname & st_all$lname==m_lname & !is.na(st_all$nyssis))
    if (length(idx)>0) found_nyssis <- st_all$nyssis[idx[1]]
  }

  # Try 3: Last name + DOB → NYSSIS
  if (is.na(found_nyssis) && !is.na(m_lname) && !is.na(m_dob) && safe_nchar(m_lname)>1) {
    idx <- which(!is.na(st_all$lname) & !is.na(st_all$dob) &
                 st_all$lname==m_lname & st_all$dob==m_dob & !is.na(st_all$nyssis))
    if (length(idx)>0) found_nyssis <- st_all$nyssis[idx[1]]
  }

  # Try 4: First name + DOB → NYSSIS (unique match only)
  if (is.na(found_nyssis) && !is.na(m_fname) && !is.na(m_dob) && safe_nchar(m_fname)>1) {
    idx <- which(!is.na(st_all$fname) & !is.na(st_all$dob) &
                 st_all$fname==m_fname & st_all$dob==m_dob & !is.na(st_all$nyssis))
    if (length(idx)==1) found_nyssis <- st_all$nyssis[idx[1]]
  }

  if (!is.na(found_nyssis)) {
    unified$`NYSSIS ID`[i] <- found_nyssis
    nyssis_patch_count <- nyssis_patch_count + 1L
  }
}

cat(sprintf("  Pass 2 — NYSSIS IDs patched from ST:   %d\n", nyssis_patch_count))
cat(sprintf("  Still missing NYSSIS IDs:              %d\n",
  sum(is.na(unified$`NYSSIS ID`) | unified$`NYSSIS ID`=="0")))

n <- nrow(unified)

# Not Found in ST — missing NYSSIS or "Not in ST System"
not_found <- unified %>%
  filter(is.na(`NYSSIS ID`) | `NYSSIS ID` == "0" |
         (!is.na(`Student Status`) & `Student Status` == "Not in ST System"))

# ============================================================================
# VALIDATION SUMMARY (console)
# ============================================================================
cat("\n================ UNIFIED ROSTER SUMMARY ================\n")
cat(sprintf("Total students: %s\n\n", format(n, big.mark=",")))
cat("Rows per charter:\n")
print(unified %>% count(`Charter_Source Ticket #`, name="rows"), n=Inf)
cat("\nSPED breakdown:\n")
print(unified %>% count(`Charter_Source Ticket #`, SPED_Flag, name="rows"), n=Inf)
cat("\nFTE by charter:\n")
print(unified %>% group_by(`Charter_Source Ticket #`) %>%
  summarise(FTE = round(sum(`SCSD FTE`, na.rm=TRUE), 3)), n=Inf)
cat(sprintf("\nTotal FTE: %.3f\n", sum(unified$`SCSD FTE`, na.rm=TRUE)))
cat(sprintf("Missing NYSSIS (after ST patch):      %d\n", sum(is.na(unified$`NYSSIS ID`) | unified$`NYSSIS ID`=="0", na.rm=TRUE)))
cat(sprintf("Missing Student IDs (after ST patch): %d\n", sum(is.na(unified$`Student ID`) | unified$`Student ID`=="0", na.rm=TRUE)))
cat(sprintf("Not Found in ST tab: %d rows\n", nrow(not_found)))

# ============================================================================
# STYLES
# ============================================================================
hdr_dark   <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#1F3864", fontColour="#FFFFFF")
hdr_teal   <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#17375E", fontColour="#FFFFFF")
hdr_orange <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#C55A11", fontColour="#FFFFFF")
hdr_purple <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#4B2D83", fontColour="#FFFFFF")
hdr_green  <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#375623", fontColour="#FFFFFF")
date_style   <- createStyle(numFmt="mm/dd/yyyy")
sped_style   <- createStyle(fgFill="#FFF2CC")
gened_style  <- createStyle(fgFill="#E2EFDA")
good_style   <- createStyle(fgFill="#C6EFCE", fontColour="#276221")
nogood_style <- createStyle(fgFill="#FFC7CE", fontColour="#9C0006")
flag_style   <- createStyle(fgFill="#FFEB9C", fontColour="#9C5700", textDecoration="bold")

col_widths <- c(18,13,34,13,12,14,12,10,12,32,12,12,9,16,28,12)

apply_dates <- function(wb, sheet, start_row, nrows) {
  for (dc in c(2,9,11,12)) {
    addStyle(wb, sheet, date_style, rows=start_row:(start_row+nrows-1),
             cols=dc, gridExpand=TRUE)
  }
}

apply_cf <- function(wb, sheet, start_row, nrows) {
  r <- start_row:(start_row+nrows-1)
  conditionalFormatting(wb, sheet, cols=16, rows=r,
    type="contains", rule="SPED",    style=sped_style)
  conditionalFormatting(wb, sheet, cols=16, rows=r,
    type="contains", rule="Gen Ed",  style=gened_style)
  conditionalFormatting(wb, sheet, cols=14, rows=r,
    type="contains", rule="Good",    style=good_style)
  conditionalFormatting(wb, sheet, cols=14, rows=r,
    type="contains", rule="No Good", style=nogood_style)
}

# ============================================================================
# CREATE WORKBOOK
# ============================================================================
wb <- createWorkbook()

# ============================================================================
# TAB 1 — All Charters Unified
# ============================================================================
addWorksheet(wb, "All Charters Unified")
writeData(wb, "All Charters Unified", unified, withFilter=TRUE)
addStyle(wb, "All Charters Unified", hdr_dark,
         rows=1, cols=1:16, gridExpand=TRUE)
apply_dates(wb, "All Charters Unified", 2, n)
apply_cf(wb, "All Charters Unified", 2, n)
freezePane(wb, "All Charters Unified", firstRow=TRUE)
setColWidths(wb, "All Charters Unified", cols=1:16, widths=col_widths)

# ============================================================================
# TAB 2 — Formula Version
# ============================================================================
addWorksheet(wb, "Formula Version")
fv_hdr <- c(template_cols, "Resolved_NYSSIS", "Match_Source")
writeData(wb, "Formula Version",
          as.data.frame(matrix(fv_hdr, nrow=1)), colNames=FALSE, startRow=1)
addStyle(wb, "Formula Version", hdr_teal, rows=1, cols=1:18, gridExpand=TRUE)

writeData(wb, "Formula Version", unified, startRow=2, colNames=FALSE)
apply_dates(wb, "Formula Version", 2, n)
apply_cf(wb, "Formula Version", 2, n)

nyssis_f <- sapply(2:(n+1), function(i) sprintf(
  "IF(AND(E%d<>\"\",E%d<>\"Not Found\",NOT(ISBLANK(E%d))),IFERROR(INDEX('Update Log'!$A:$A,MATCH(E%d,'Update Log'!$B:$B,0)),D%d),D%d)",
  i,i,i,i,i,i))
writeFormula(wb, "Formula Version", x=nyssis_f, startCol=17, startRow=2)

source_f <- sapply(2:(n+1), function(i) sprintf(
  "IF(AND(E%d<>\"\",E%d<>\"Not Found\",NOT(ISBLANK(E%d))),IF(ISNUMBER(MATCH(E%d,'Update Log'!$B:$B,0)),\"Update Log\",\"Master Roster\"),\"No Student ID\")",
  i,i,i,i))
writeFormula(wb, "Formula Version", x=source_f, startCol=18, startRow=2)

conditionalFormatting(wb, "Formula Version", cols=18, rows=2:(n+1),
  type="contains", rule="Update Log", style=flag_style)
freezePane(wb, "Formula Version", firstRow=TRUE)
setColWidths(wb, "Formula Version", cols=1:16, widths=col_widths)
setColWidths(wb, "Formula Version", cols=17:18, widths=c(16,14))

# ============================================================================
# TAB 3 — Copy & Paste (Values Only)
# ============================================================================
addWorksheet(wb, "Copy & Paste")

instr_txt <- paste0(
  "COPY & PASTE READY  |  Select A3 through P", n+2,
  ", copy, then paste as VALUES ONLY into your FY 25-26 master running tab.")
writeData(wb, "Copy & Paste",
          data.frame(x=instr_txt), colNames=FALSE, startRow=1)
addStyle(wb, "Copy & Paste",
  createStyle(fgFill="#FCE4D6", textDecoration="bold",
              fontColour="#843C0C", halign="left"),
  rows=1, cols=1:16, gridExpand=TRUE)
mergeCells(wb, "Copy & Paste", cols=1:16, rows=1)

writeData(wb, "Copy & Paste",
          as.data.frame(matrix(template_cols, nrow=1)),
          colNames=FALSE, startRow=2)
addStyle(wb, "Copy & Paste", hdr_orange, rows=2, cols=1:16, gridExpand=TRUE)

writeData(wb, "Copy & Paste", unified, startRow=3, colNames=FALSE)
apply_dates(wb, "Copy & Paste", 3, n)
apply_cf(wb, "Copy & Paste", 3, n)
freezePane(wb, "Copy & Paste", firstActiveRow=3)
setColWidths(wb, "Copy & Paste", cols=1:16, widths=col_widths)

# ============================================================================
# TAB 4 — Not Found in ST
# ============================================================================
addWorksheet(wb, "Not Found in ST")
nf_cols <- c(template_cols, "Action Needed")

nf_instr <- paste0("NOT FOUND IN ST  |  ", nrow(not_found),
  " students have no NYSSIS ID or are not in the ST enrollment system. ",
  "Look up manually and add resolved NYSSIS to the Update Log tab.")
writeData(wb, "Not Found in ST",
          data.frame(x=nf_instr), colNames=FALSE, startRow=1)
addStyle(wb, "Not Found in ST",
  createStyle(fgFill="#FFC7CE", textDecoration="bold",
              fontColour="#9C0006", halign="left"),
  rows=1, cols=1:17, gridExpand=TRUE)
mergeCells(wb, "Not Found in ST", cols=1:17, rows=1)

writeData(wb, "Not Found in ST",
          as.data.frame(matrix(nf_cols, nrow=1)),
          colNames=FALSE, startRow=2)
addStyle(wb, "Not Found in ST", hdr_purple, rows=2, cols=1:17, gridExpand=TRUE)

nf_out <- not_found %>%
  mutate(`Action Needed` = case_when(
    is.na(`NYSSIS ID`) | `NYSSIS ID` == "0" ~
      "Look up NYSSIS in SchoolTool — enter resolved NYSSIS in Update Log tab",
    `Student Status` == "Not in ST System" ~
      "Student not found in ST enrollment — verify enrollment status with Patrick",
    TRUE ~ "Review required"
  )) %>%
  select(all_of(nf_cols))

writeData(wb, "Not Found in ST", nf_out, startRow=3, colNames=FALSE)
nf_n <- nrow(nf_out)
apply_dates(wb, "Not Found in ST", 3, nf_n)
addStyle(wb, "Not Found in ST",
  createStyle(fgFill="#FFE4E1"),
  rows=3:(nf_n+2), cols=1:17, gridExpand=TRUE)
freezePane(wb, "Not Found in ST", firstActiveRow=3)
setColWidths(wb, "Not Found in ST", cols=1:16, widths=col_widths)
setColWidths(wb, "Not Found in ST", cols=17, widths=44)

# ============================================================================
# TAB 5 — Update Log
# ============================================================================
addWorksheet(wb, "Update Log")
ul_cols <- c("NYSSIS ID","Student ID","Last Name","First Name","DOB",
             "Grade Level","School Name","Charter_Source Ticket #",
             "Resolved_By","Resolved_Date","Notes")

ul_instr <- paste0(
  "UPDATE LOG  |  Paste resolved NYSSIS records here from R or Python output. ",
  "Col A = NYSSIS ID | Col B = Student ID (must match master). ",
  "Formula Version tab will auto-pull resolved NYSSIS for any matching Student ID.")
writeData(wb, "Update Log",
          data.frame(x=ul_instr), colNames=FALSE, startRow=1)
addStyle(wb, "Update Log",
  createStyle(fgFill="#E2EFDA", textDecoration="bold",
              fontColour="#375623", halign="left"),
  rows=1, cols=1:length(ul_cols), gridExpand=TRUE)
mergeCells(wb, "Update Log", cols=1:length(ul_cols), rows=1)

writeData(wb, "Update Log",
          as.data.frame(matrix(ul_cols, nrow=1)),
          colNames=FALSE, startRow=2)
addStyle(wb, "Update Log", hdr_green, rows=2,
         cols=1:length(ul_cols), gridExpand=TRUE)

# Seed with students still missing NYSSIS after both ST lookup passes
ul_seed <- unified %>%
  filter(is.na(`NYSSIS ID`) | `NYSSIS ID` == "0" | `NYSSIS ID` == "Not Found") %>%
  mutate(
    Resolved_By   = "Pending",
    Resolved_Date = as.Date(NA),
    Notes = case_when(
      str_detect(coalesce(`First Name`,""), "(?i)^Anessa$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Bobby$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Donald$") ~
        "No last name in source — search ST by DOB 12/07/2020.",
      str_detect(coalesce(`First Name`,""), "(?i)^Edward$") ~
        "No last name in source — search ST by DOB 03/20/2020.",
      str_detect(coalesce(`First Name`,""), "(?i)^Kaiden$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Liliana$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Rahmier$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Xavier$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Johnny$") ~
        "Not found in any ST export (Gaston LLL, Johnny) — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Atlantis$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Lenia$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Eric$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Scarlett$") ~
        "Not found in any ST export — verify enrollment with Patrick.",
      str_detect(coalesce(`First Name`,""), "(?i)^Mackenzie$") ~
        "Not found in any ST export (Aldrich, Mackenzie) — verify enrollment with Patrick.",
      is.na(`First Name`) | `First Name` %in% c("NA","") ~
        "No name in source file — manual lookup required.",
      TRUE ~ "Not found in any ST export — verify enrollment with Patrick."
    )
  ) %>%
  select(`NYSSIS ID`, `Student ID`, `Last Name`, `First Name`, `DOB`,
         `Grade Level`, `School Name`, `Charter_Source Ticket #`,
         Resolved_By, Resolved_Date, Notes)

writeData(wb, "Update Log", ul_seed, startRow=3, colNames=FALSE)
ul_n <- nrow(ul_seed)
if (ul_n > 0) {
  addStyle(wb, "Update Log", date_style, rows=3:(ul_n+2), cols=5, gridExpand=TRUE)
  addStyle(wb, "Update Log", date_style, rows=3:(ul_n+2), cols=10, gridExpand=TRUE)
  addStyle(wb, "Update Log",
    createStyle(fgFill="#FFFACD"),
    rows=3:(ul_n+2), cols=1:length(ul_cols), gridExpand=TRUE)
}
freezePane(wb, "Update Log", firstActiveRow=3)
setColWidths(wb, "Update Log", cols=1:length(ul_cols),
             widths=c(13,12,14,12,12,10,30,20,14,14,52))

# ============================================================================
# SAVE
# ============================================================================
saveWorkbook(wb, output_path, overwrite=TRUE)
cat("\n==== DONE ====\n")
cat("Output file:", output_path, "\n")
cat(sprintf("Tab 1 — All Charters Unified : %s students\n", format(n, big.mark=",")))
cat(sprintf("Tab 2 — Formula Version       : %s rows + NYSSIS lookup formulas\n", format(n, big.mark=",")))
cat(sprintf("Tab 3 — Copy & Paste          : %s rows (values only)\n", format(n, big.mark=",")))
cat(sprintf("Tab 4 — Not Found in ST       : %d students flagged\n", nf_n))
cat(sprintf("Tab 5 — Update Log            : %d pending resolution\n", ul_n))
