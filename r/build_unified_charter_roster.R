# ============================================================================
# Build Unified Charter Roster (25-26)
# Daquan Morrison · Data Analyst · Syracuse City School District
# ----------------------------------------------------------------------------
# Reads the "Unified Gen-SPED Roster" tab from each charter's pre-built file
# and combines them into one standardized output matching the unified template.
#
# Source files (Unified Gen-SPED Roster tab):
#   southside_291462_unified-roster.xlsx
#   sas_291463_unified-roster.xlsx
#   sasccs_291464_unified-roster.xlsx
#   ontech_final_unified-roster.xlsx
#
# Output columns (exact order):
#   Charter_Source Ticket # | Charter Month | Source_File |
#   NYSSIS ID | Student ID | Last Name | First Name | Grade Level | DOB |
#   School Name | Enroll Date | Withdraw Date | SCSD FTE |
#   Student Status | Notes for Patrick | SPED_Flag
# ============================================================================

pkgs_needed  <- c("readxl","dplyr","stringr","lubridate","purrr","tibble","openxlsx")
pkgs_missing <- pkgs_needed[!pkgs_needed %in% rownames(installed.packages())]
if (length(pkgs_missing) > 0) install.packages(pkgs_missing, dependencies = TRUE)
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr)
  library(lubridate); library(purrr); library(tibble); library(openxlsx)
})

# ============================================================================
# SETTINGS
# ============================================================================
data_dir      <- "/home/user/workspace/uploaded_attachments/8dc6b308dcc94470ad3ac69527bb318b"
INVOICE_MONTH <- as.Date("2026-06-01")
output_path   <- "/home/user/workspace/all_charters_unified_roster_25-26_REBUILT.xlsx"

# ============================================================================
# HELPERS
# ============================================================================
parse_mixed_date <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (inherits(x, c("POSIXct","POSIXlt"))) return(as.Date(x))
  x_chr <- str_trim(as.character(x))
  x_chr[x_chr %in% c("","NA","N/A","null","NULL","0")] <- NA_character_
  suppressWarnings(x_num <- as.numeric(x_chr))
  out <- rep(as.Date(NA), length(x_chr))
  idx_serial <- !is.na(x_num) & x_num > 20000
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
  # Expects "Last, First" format
  s <- str_trim(as.character(x))
  last  <- str_trim(str_extract(s, "^[^,]+"))
  first <- str_trim(str_replace(s, "^[^,]+,?\\s*", ""))
  tibble(last = last, first = first)
}

# ============================================================================
# BUILD EACH CHARTER
# All read from "Unified Gen-SPED Roster" tab — SPED_Flag already present
# ============================================================================

# ---------- SOUTHSIDE 291462 ------------------------------------------------
# Cols: NYSSIS ID | Student ID | Student First Name | Student Last Name |
#       Grade | Birth Date | Entered | Withdrew | Southside FTE | SCSD FTE |
#       School Name | Student Status | Notes for Patrick | SPED_Flag
build_southside <- function() {
  message("Building Southside 291462...")
  df <- read_excel(
    file.path(data_dir, "southside_291462_unified-roster.xlsx"),
    sheet = "Unified Gen-SPED Roster", col_names = TRUE, .name_repair = "minimal"
  )
  # Drop rows with no NYSSIS and no name (true blank rows)
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
# Cols: NYSSIS ID | Student ID | Last Name | First Name | Date of Birth |
#       Current Grade Level | Entry Date | Leave Date | FTE |
#       School Name | Student Status | SPED_Flag
build_sas <- function() {
  message("Building SAS 291463...")
  df <- read_excel(
    file.path(data_dir, "sas_291463_unified-roster.xlsx"),
    sheet = "Unified Gen-SPED Roster", col_names = TRUE, .name_repair = "minimal"
  )
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
    `Notes for Patrick`       = NA_character_,    # SAS has no Notes for Patrick column
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ---------- SASCCS 291464 ----------------------------------------------------
# Cols: NYSSIS ID | Student ID | Full Name | Date of Birth |
#       Current Grade Level | Entry Date | Leave Date | Status (=FTE) |
#       School Name | Student Status | SPED_Flag
build_sasccs <- function() {
  message("Building SASCCS 291464...")
  df <- read_excel(
    file.path(data_dir, "sasccs_291464_unified-roster.xlsx"),
    sheet = "Unified Gen-SPED Roster", col_names = TRUE, .name_repair = "minimal"
  )
  # Filter by col position — NYSSIS ID = col 1, Full Name = col 3
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
    `SCSD FTE`                = clean_fte(df[["Status"]]),    # Status = billed FTE for SASCCS
    `Student Status`          = str_trim(as.character(df[["Student Status"]])),
    `Notes for Patrick`       = NA_character_,    # SASCCS has no Notes for Patrick in unified tab
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ---------- ONTECH 291465 ----------------------------------------------------
# Cols: NYSSIS ID | Studemt ID (typo) | Last Name | First Name |
#       Grade Level | DOB | Enroll Date | Withdraw Date | OnTECH FTE | SPED_Flag
build_ontech <- function() {
  message("Building OnTech 291465...")
  df <- read_excel(
    file.path(data_dir, "ontech_final_unified-roster.xlsx"),
    sheet = "Unified Gen-SPED Roster", col_names = TRUE, .name_repair = "minimal"
  )
  df <- df[!is.na(df[["NYSSIS ID"]]) | !is.na(df[["Last Name"]]), ]

  # Resolve Student ID col — typo "Studemt ID" in source
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
    `Student Status`          = NA_character_,    # OnTech unified tab has no Student Status
    `Notes for Patrick`       = NA_character_,
    `SPED_Flag`               = str_trim(as.character(df[["SPED_Flag"]]))
  )
}

# ============================================================================
# COMBINE
# ============================================================================
unified <- bind_rows(
  build_southside(),
  build_sas(),
  build_sasccs(),
  build_ontech()
)

template_cols <- c(
  "Charter_Source Ticket #","Charter Month","Source_File","NYSSIS ID",
  "Student ID","Last Name","First Name","Grade Level","DOB",
  "School Name","Enroll Date","Withdraw Date","SCSD FTE",
  "Student Status","Notes for Patrick","SPED_Flag"
)
unified <- unified %>% select(all_of(template_cols))

# Normalize NA text to actual NA
unified <- unified %>%
  mutate(across(where(is.character), ~ifelse(.x %in% c("NA","N/A","NULL","null",""), NA_character_, .x)))

# ============================================================================
# VALIDATION SUMMARY
# ============================================================================
cat("\n================ UNIFIED ROSTER SUMMARY ================\n")
cat(sprintf("Total students combined: %s\n\n", format(nrow(unified), big.mark=",")))

cat("Rows per charter:\n")
print(unified %>% count(`Charter_Source Ticket #`, name="rows"), n=Inf)

cat("\nSPED_Flag breakdown:\n")
print(unified %>% count(`Charter_Source Ticket #`, `SPED_Flag`, name="rows"), n=Inf)

cat("\nStudent Status breakdown:\n")
print(unified %>% count(`Charter_Source Ticket #`, `Student Status`, name="rows"), n=Inf)

cat(sprintf("\nTotal billed SCSD FTE: %.3f\n", round(sum(unified$`SCSD FTE`, na.rm=TRUE), 3)))

cat("\nFTE by charter:\n")
print(unified %>% group_by(`Charter_Source Ticket #`) %>%
  summarise(FTE = round(sum(`SCSD FTE`, na.rm=TRUE), 3)), n=Inf)

dup_ids <- unified %>%
  filter(!is.na(`NYSSIS ID`)) %>%
  count(`NYSSIS ID`, name="n") %>% filter(n > 1)
cat(sprintf("\nDuplicate NYSSIS IDs across all charters: %s\n", nrow(dup_ids)))
if (nrow(dup_ids) > 0) {
  dup_detail <- unified %>%
    filter(`NYSSIS ID` %in% dup_ids$`NYSSIS ID`) %>%
    select(`NYSSIS ID`, `Charter_Source Ticket #`, `Last Name`, `First Name`) %>%
    arrange(`NYSSIS ID`)
  cat("Duplicate NYSSIS details:\n")
  print(dup_detail, n=Inf)
}

missing_nyssis <- sum(is.na(unified$`NYSSIS ID`))
missing_fte    <- sum(is.na(unified$`SCSD FTE`))
cat(sprintf("\nMissing NYSSIS ID: %s\n", missing_nyssis))
cat(sprintf("Missing SCSD FTE:  %s\n", missing_fte))

# ============================================================================
# WRITE OUTPUT
# ============================================================================
wb <- createWorkbook()
addWorksheet(wb, "All Charters Unified")
writeData(wb, "All Charters Unified", unified, withFilter=TRUE)

hdr_style  <- createStyle(textDecoration="bold", halign="left",
                           fgFill="#1F3864", fontColour="#FFFFFF")
date_style <- createStyle(numFmt="mm/dd/yyyy")
addStyle(wb, 1, hdr_style, rows=1, cols=seq_along(template_cols), gridExpand=TRUE)

date_cols <- match(c("DOB","Enroll Date","Withdraw Date","Charter Month"), template_cols)
for (dc in date_cols) {
  addStyle(wb, 1, date_style, rows=2:(nrow(unified)+1), cols=dc, gridExpand=TRUE)
}
freezePane(wb, 1, firstRow=TRUE)
setColWidths(wb, 1, cols=seq_along(template_cols), widths="auto")

# Conditional formatting: SPED_Flag
sped_style   <- createStyle(fgFill="#FFF2CC")   # yellow
gened_style  <- createStyle(fgFill="#E2EFDA")   # green
flag_col     <- which(template_cols == "SPED_Flag")
conditionalFormatting(wb, 1, cols=flag_col, rows=2:(nrow(unified)+1),
  type="contains", rule="SPED", style=sped_style)
conditionalFormatting(wb, 1, cols=flag_col, rows=2:(nrow(unified)+1),
  type="contains", rule="Gen Ed", style=gened_style)

# Conditional formatting: Student Status
status_col   <- which(template_cols == "Student Status")
good_style   <- createStyle(fgFill="#C6EFCE", fontColour="#276221")
nogood_style <- createStyle(fgFill="#FFC7CE", fontColour="#9C0006")
conditionalFormatting(wb, 1, cols=status_col, rows=2:(nrow(unified)+1),
  type="contains", rule="Good", style=good_style)
conditionalFormatting(wb, 1, cols=status_col, rows=2:(nrow(unified)+1),
  type="contains", rule="No Good", style=nogood_style)

saveWorkbook(wb, output_path, overwrite=TRUE)
cat("\n==== DONE ====\n")
cat("Output written to:", output_path, "\n")
cat(sprintf("Rows: %s  |  Cols: %s\n", format(nrow(unified), big.mark=","), ncol(unified)))
