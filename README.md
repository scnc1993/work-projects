# Work Projects — Data Engineering & Analytics Portfolio

A portfolio of production data-engineering work by **Daquan Morrison** — Education Data Analyst / Analytics Engineer (Thrivora Holdings LLC). The flagship project is the **Charter Worksheet Data Validation Engine**, a layout-agnostic audit system that reconciles charter-school enrollment rosters against the SchoolTool student-information system and flags FTE/billing and compliance risks.

> Built with **SQL · Python (pandas) · R**, designed around a **medallion (bronze → silver → gold)** architecture, and engineered to run identically across Microsoft Fabric, Snowflake, PostgreSQL, and local files.

---

## 🔎 Flagship: Charter Audit Engine

A single audit pipeline implemented **three ways** (SQL, Python, R) from one shared specification, so the same business logic runs in a warehouse, a notebook, or a local script.

**What it does:** joins each charter roster to SchoolTool on the NYSSIS state ID, resolves Student IDs through a fallback waterfall, and computes **13 calculated validation fields** — match status, IEP/SPED consistency, entry/exit date variances, severity, and a recommended action per record.

| Field | Purpose |
|---|---|
| `resolved_studentid` | COALESCE waterfall: SchoolTool ID → Universal ID → NYSSIS mapping |
| `id_match_status` / `id_source` | How (and whether) each roster row matched SchoolTool |
| `is_iep_flag` / `roster_vs_disability_check` | Gen Ed vs. SPED consistency between systems |
| `entry_diff_days` / `entry_status_category` | Enrollment-date variance (Exact / ±5 tolerance / Finance Risk) |
| `exit_diff_days` / `exit_status_category` / `exit_issue_type` | Withdrawal-date variance + specific issue classification |
| `final_validation_status` | Valid / Valid with Warning / Incomplete / Invalid |
| `severity` | Critical / High / Low / None |
| `recommended_action` | Plain-language remediation step |

**Engineering highlights**
- **Layout-agnostic ingestion** — auto-detects the roster sheet, header row, and columns via fuzzy matching, so it works across inconsistent yearly file formats without rework.
- **Robust normalization** — handles Excel serial dates, `YYYYMMDD`, SAS-style `14FEB2026`, leading-zero IDs, and trailing-`.0` artifacts.
- **Correct join semantics** — roster-driven LEFT JOIN on NYSSIS cast to text on both sides, eliminating a String-vs-Integer all-null defect; Student IDs are validation inputs, not join keys.
- **Severity-coded Excel output** with full-audit, exceptions-only, and summary tabs.

---

## 📁 Repository layout

```
sql/        charter_audit_engine.sql        4-layer CTE pipeline (bronze→silver→gold), all 13 fields
python/     charter_audit_engine.py         pandas implementation + CLI + openpyxl export
r/          charter_audit_engine.R          dplyr/tidyr implementation + openxlsx export
notebooks/  charter_audit_engine.ipynb      runnable %%sql notebook + severity summary
            nyssis_studentid_matching.ipynb NYSSIS/StudentID resolution (Excel-array ↔ SQL ↔ code)
docs/       charter_audit_engine_README.md  field-by-field spec, provenance, join rules
```

---

## 🛠️ Tech stack

**Languages:** SQL (T-SQL / Fabric, Snowflake, Postgres dialects), Python (pandas, numpy, openpyxl), R (dplyr, tidyr, readxl, openxlsx)
**Platforms:** Microsoft Fabric · Power BI · Azure SQL · SchoolTool / NYSSIS
**Patterns:** Medallion architecture, idempotent pipelines, data-quality validation, fuzzy schema detection

---

## ▶️ Running the engine

**Python**
```bash
pip install pandas numpy openpyxl
python python/charter_audit_engine.py roster.xlsx schooltool.xlsx --school "Southside" --out audit.xlsx
```

**R**
```r
source("r/charter_audit_engine.R")
run_charter_audit("roster.xlsx", "schooltool.xlsx", school_label = "Southside")
```

**SQL** — run `sql/charter_audit_engine.sql` against your warehouse (adjust the `DATEDIFF` dialect note at the top of the file).

---

## 🔒 Data & privacy

This repository contains **engine logic and schema only** — no student records, no real identifiers, no credentials. All examples use synthetic data. Real datasets are excluded via `.gitignore`.

---

*Part of the Thrivora Holdings analytics platform — a vendor-independent, Git-versioned, durable data environment (Azure + Fabric + Power BI).*
