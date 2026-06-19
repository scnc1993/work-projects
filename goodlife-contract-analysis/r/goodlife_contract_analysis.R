# =============================================================================
# Good Life, LLC — Contract Evaluation Analysis (R)
# -----------------------------------------------------------------------------
# Rebuilds the GoodLife contract-evaluation summaries from the corrected,
# cleaned long-format dataset (one row per student per school year).
#
# The corrected dataset EXCLUDES non-SCSD / unknown placeholder records
# (e.g. the 16 "Pine Grove Middle School" rows that were a field-trip-location
# artifact from "OCM SKATE PINE GROVE MS" in the raw SchoolTool export and are
# NOT SCSD enrollments). Grand total of students served = 403.
#
# Outputs (written to ../output/):
#   - students_served_by_school_year.csv   (matches workbook "Summary" table 1)
#   - suspensions_by_school_year.csv        (OSS / ISS counts & days)
#   - grades_by_school_year.csv             (avg year score, % passing all MPs)
#   - attendance_by_school_year.csv         (avg attendance rate, % >= 90%)
#   - contract_kpis.csv                     (district-wide KPI rollup)
#
# Usage:
#   Rscript r/goodlife_contract_analysis.R
#
# Author: Daquan Morrison  |  SCSD Education Data Analysis  |  Thrivora Holdings
# =============================================================================

suppressWarnings(suppressMessages({
  library(dplyr)
  library(tidyr)
  library(readr)
  library(stringr)
}))

# ---- Paths ------------------------------------------------------------------
this_dir   <- tryCatch(dirname(sub("--file=", "",
                  commandArgs(trailingOnly = FALSE)[grep("--file=",
                  commandArgs(trailingOnly = FALSE))])), error = function(e) ".")
if (length(this_dir) == 0 || this_dir == "") this_dir <- "."
proj_dir   <- normalizePath(file.path(this_dir, ".."), mustWork = FALSE)
data_path  <- file.path(proj_dir, "data", "goodlife_long_data.csv")
out_dir    <- file.path(proj_dir, "output")
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

# ---- Load -------------------------------------------------------------------
raw <- read_csv(data_path, show_col_types = FALSE)

# Standardize column names to snake_case for safe handling
df <- raw %>%
  rename(
    year             = `Year`,
    student_id       = `Student ID`,
    name             = `Name`,
    school           = `School`,
    mp1              = `MP1`,
    mp2              = `MP2`,
    mp3              = `MP3`,
    mp4              = `MP4`,
    year_avg         = `Year Avg`,
    passing_all_mps  = `Passing All MPs`,
    oss              = `OSS`,
    iss              = `ISS`,
    total_susp       = `Total Susp`,
    oss_days         = `OSS Days`,
    iss_days         = `ISS Days`,
    total_susp_days  = `Total Susp Days`,
    susp_1plus       = `1+ Suspensions`,
    susp_2plus       = `2+ Suspensions`,
    days_5plus       = `5+ Days Suspended`,
    attendance_rate  = `Attendance Rate`
  )

# The final source column header carries a unicode ">=" symbol. Rename it
# positionally so the script is robust across locales and text editors.
names(df)[ncol(df)] <- "attendance_90"

# ---- Data-quality guard: drop any non-SCSD / placeholder rows ---------------
# The Pine Grove artifact rows (a field-trip-location string mistaken for a
# school) and any explicit Non-SCSD / Unknown bucket are excluded here so they
# can never re-enter the contract reporting.
# NOTE: real Student IDs are de-identified to surrogate keys (STU0001 ...) in
# the public dataset; exclusion is therefore driven by the school name.
df <- df %>%
  filter(school != "Non-SCSD / Unknown",
         !str_detect(school, regex("pine grove", ignore_case = TRUE)))

cat(sprintf("Loaded %d student-year records across %d schools, %d years.\n",
            nrow(df), n_distinct(df$school), n_distinct(df$year)))

# ---- 1) Students served by school by year -----------------------------------
served <- df %>%
  count(school, year, name = "students") %>%
  pivot_wider(names_from = year, values_from = students, values_fill = 0) %>%
  arrange(school)

served <- served %>%
  mutate(Total = rowSums(across(where(is.numeric)))) %>%
  bind_rows(
    served %>%
      summarise(across(where(is.numeric), sum)) %>%
      mutate(school = "GRAND TOTAL", Total = rowSums(across(where(is.numeric))))
  )

write_csv(served, file.path(out_dir, "students_served_by_school_year.csv"))

# ---- 2) Suspensions by school by year ---------------------------------------
susp <- df %>%
  group_by(school, year) %>%
  summarise(
    students        = n(),
    oss             = sum(oss, na.rm = TRUE),
    iss             = sum(iss, na.rm = TRUE),
    oss_days        = sum(oss_days, na.rm = TRUE),
    iss_days        = sum(iss_days, na.rm = TRUE),
    total_susp_days = sum(total_susp_days, na.rm = TRUE),
    students_1plus  = sum(susp_1plus, na.rm = TRUE),
    students_2plus  = sum(susp_2plus, na.rm = TRUE),
    students_5plus_days = sum(days_5plus, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  arrange(year, school)

write_csv(susp, file.path(out_dir, "suspensions_by_school_year.csv"))

# ---- 3) Grades / academic performance by school by year ---------------------
grades <- df %>%
  group_by(school, year) %>%
  summarise(
    students          = n(),
    avg_year_score    = round(mean(year_avg, na.rm = TRUE), 2),
    n_passing_all_mps = sum(passing_all_mps, na.rm = TRUE),
    pct_passing_all_mps = round(100 * sum(passing_all_mps, na.rm = TRUE) /
                                  sum(!is.na(year_avg)), 1),
    .groups = "drop"
  ) %>%
  arrange(year, school)

write_csv(grades, file.path(out_dir, "grades_by_school_year.csv"))

# ---- 4) Attendance by school by year ----------------------------------------
attend <- df %>%
  group_by(school, year) %>%
  summarise(
    students          = n(),
    avg_attendance_rate = round(mean(attendance_rate, na.rm = TRUE), 3),
    n_at_or_above_90  = sum(attendance_90, na.rm = TRUE),
    pct_at_or_above_90 = round(100 * sum(attendance_90, na.rm = TRUE) /
                                 sum(!is.na(attendance_rate)), 1),
    .groups = "drop"
  ) %>%
  arrange(year, school)

write_csv(attend, file.path(out_dir, "attendance_by_school_year.csv"))

# ---- 5) District-wide contract KPIs -----------------------------------------
kpis <- df %>%
  group_by(year) %>%
  summarise(
    students_served      = n(),
    schools_served       = n_distinct(school),
    avg_year_score       = round(mean(year_avg, na.rm = TRUE), 2),
    pct_passing_all_mps  = round(100 * sum(passing_all_mps, na.rm = TRUE) /
                                   sum(!is.na(year_avg)), 1),
    avg_attendance_rate  = round(mean(attendance_rate, na.rm = TRUE), 3),
    pct_attendance_90    = round(100 * sum(attendance_90, na.rm = TRUE) /
                                   sum(!is.na(attendance_rate)), 1),
    students_1plus_susp  = sum(susp_1plus, na.rm = TRUE),
    total_susp_days      = sum(total_susp_days, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  arrange(year)

write_csv(kpis, file.path(out_dir, "contract_kpis.csv"))

# ---- Console summary --------------------------------------------------------
cat("\n=== Students served by school by year ===\n")
print(as.data.frame(served), row.names = FALSE)
cat("\n=== District-wide KPIs by year ===\n")
print(as.data.frame(kpis), row.names = FALSE)
cat(sprintf("\nGRAND TOTAL students served (all years): %d\n", nrow(df)))
cat(sprintf("Outputs written to: %s\n", out_dir))
