# Good Life, LLC — Contract Evaluation Analysis

End-to-end analysis of a multi-year K–12 vendor contract for the **Syracuse City
School District (SCSD)**, evaluating student outcomes (grades, suspensions,
attendance) delivered by the contractor **Good Life, LLC** across school years
**SY 2023-24, SY 2024-25, and SY 2025-26**.

This project reproduces the workbook's reporting summaries from a single cleaned,
long-format dataset using **R** and **portable SQL**, so the results can be
audited and regenerated outside of Excel.

## Headline numbers

| Metric | SY 2023-24 | SY 2024-25 | SY 2025-26 |
| --- | ---: | ---: | ---: |
| Students served | 176 | 92 | 135 |
| Schools served | 13 | 9 | 6 |
| Avg. year score | 69.87 | 69.79 | 72.38 |
| % passing all marking periods | 53.5% | 44.4% | 63.7% |
| Avg. attendance rate | 85.1% | 86.1% | 88.5% |
| % at/above 90% attendance | 44.4% | 55.6% | 60.8% |

**Grand total students served (all years): 403**, across 14 SCSD schools.

## Data-quality note (why the corrected file exists)

An earlier version of the workbook listed **"Pine Grove Middle School" (16
students, SY 2025-26)** as if it were an SCSD school. It is not — Pine Grove
Middle School belongs to the **East Syracuse Minoa Central School District**.
Those 16 records had short 5-digit IDs (not the SCSD `200…`/`210…` format) and
no name, grades, behavior, or attendance data. They originated from a raw
SchoolTool export field that read **"OCM SKATE PINE GROVE MS"** — a field-trip /
activity location, not an enrollment.

The corrected dataset **excludes** these placeholder rows, bringing the grand
total from 418 to **403**. Both the R script and the SQL silver layer
additionally guard against any such rows re-entering the analysis.

## Repository layout

```
goodlife-contract-analysis/
├── FY-23-25-GoodLife-Contract-Analysis-CORRECTED.xlsx  # source workbook
├── data/
│   └── goodlife_long_data.csv          # cleaned long-format extract (1 row / student / year)
├── r/
│   └── goodlife_contract_analysis.R    # R pipeline -> CSV summaries
├── sql/
│   └── goodlife_contract_analysis.sql  # ANSI-SQL medallion pipeline (bronze/silver/gold)
└── output/                             # generated summary CSVs
```

## Dataset grain

`data/goodlife_long_data.csv` — one row per student per school year, 21 columns:
`Year, Student ID, Name, School, MP1–MP4, Year Avg, Passing All MPs, OSS, ISS,
Total Susp, OSS Days, ISS Days, Total Susp Days, 1+ Suspensions, 2+ Suspensions,
5+ Days Suspended, Attendance Rate, Attendance ≥90%`.

## Running the R analysis

```bash
# requires: dplyr, tidyr, readr, stringr
Rscript r/goodlife_contract_analysis.R
```

Writes five summary CSVs to `output/`: students served, suspensions, grades,
attendance (each by school × year), plus a district-wide KPI rollup.

## Running the SQL pipeline

Portable ANSI SQL; tested on SQLite (also runs on Azure SQL / SQL Server,
Postgres, DuckDB with minimal change):

```bash
sqlite3 goodlife.db < sql/goodlife_contract_analysis.sql
sqlite3 goodlife.db ".mode csv" ".import --skip 1 data/goodlife_long_data.csv bronze_goodlife_long"
sqlite3 goodlife.db "SELECT * FROM gold_contract_kpis;"
```

The medallion layers are:

- **bronze_goodlife_long** — raw landing table matching the CSV
- **silver_student_year** — typed, cleaned, non-SCSD/placeholder rows removed
- **gold_*** — reporting views: students served, suspensions, grades,
  attendance, and district KPIs

Both engines produce identical figures (verified: grand total = 403, per-school
counts and KPIs match the workbook's `Summary` tab).

---

*Author: Daquan Morrison — SCSD Education Data Analysis / Thrivora Holdings.
Student-level data is aggregated; no personally identifying information is
included in this repository.*
