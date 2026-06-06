# Charter School Audit Engine — Documentation

## What it does
Resolves student identity across **SchoolTool (ST) enrollment** and four monthly
charter rosters (**Southside, SASCCS, SAS, OnTech**), then validates enrollment
**entry/exit dates** and **SPED status** against ST to flag billing/finance
discrepancies for charter invoice review (NYS 3-day enrollment rule context).

## Provenance
Originally built as **13 Mindex calculated fields** on a SchoolTool Advanced
Analytics visual join model, with Excel array-formula preprocessing for NYSSIS /
Student ID matching. This repo holds the faithful, vendor-independent versions so
the logic survives outside any single tool.

## Files
| File | What it is |
|------|------------|
| `sql/charter_audit_engine.sql` | The full engine as a 4-layer CTE pipeline (bronze join → silver resolve → silver audit → gold validation) |
| `notebooks/charter_audit_engine.ipynb` | Runnable Fabric/SQL notebook of the engine + a severity summary query |
| `notebooks/nyssis_studentid_matching.ipynb` | NYSSIS/Student ID matching — original Excel formulas + SQL equivalent |

## The 13 calculated fields (engine logic)
1. **resolved_studentid** — ST waterfall: `studentid` → `student_universalstudentid` → `studentid[aa_nyssis]` → `''`
2. **id_match_status** — compares resolved ID against each roster's Student ID (Southside/SASCCS/SAS/OnTech)
3. **id_source** — which ST field produced the ID
4. **is_iep_flag** — Gen Ed vs IEP from `schoolyear_disability`
5. **roster_vs_disability_check** — roster IEP code vs ST disability record consistency
6. **entry_diff_days** — roster entry − ST entry (days)
7. **entry_status_category** — Exact / Within Tolerance (±5) / Finance Risk (>5) / Missing
8. **exit_diff_days** — roster exit − ST exit (days)
9. **exit_status_category** — same tolerance scheme as entry
10. **exit_issue_type** — Missing Roster Exit / Missing ST Exit / Both Missing / Late / Early / No Issue
11. **final_validation_status** — Valid / Valid with Warning / Incomplete / Invalid (No ID / Date Mismatch)
12. **severity** — Critical / High / Low / None
13. **recommended_action** — concrete next step per discrepancy type

## Join model (hard-won from production)
- Join rosters to ST on **NYSSIS only**, **cast to text on both sides** to avoid the
  String-vs-Integer mismatch that silently produced all-null rows.
- **LEFT JOIN** (ST as spine) = keep all ST students. **RIGHT JOIN** = roster-driven.
- Student ID columns are **validation inputs**, not join keys — never add a second
  `studentid = Student ID` join clause.

## Dialect note
`DATEDIFF(DAY, start, end)` is T-SQL / Fabric / Synapse. For Snowflake use
`DATEDIFF('day', start, end)`; for Postgres use `(end - start)`.
