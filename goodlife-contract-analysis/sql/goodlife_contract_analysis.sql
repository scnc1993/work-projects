-- =============================================================================
-- Good Life, LLC -- Contract Evaluation Analysis (SQL)
-- -----------------------------------------------------------------------------
-- Portable ANSI-SQL pipeline that rebuilds the GoodLife contract-evaluation
-- summaries from the corrected long-format dataset (one row per student per
-- school year). Tested against SQLite; the analytical SELECTs are standard SQL
-- and run on SQL Server / Azure SQL, Postgres, and DuckDB with minimal changes.
--
-- Medallion layout:
--   bronze_goodlife_long  -- raw load of goodlife_long_data.csv (as-is)
--   silver_student_year   -- typed + cleaned, non-SCSD/placeholder rows removed
--   gold_*                -- contract reporting views (served / grades / susp /
--                            attendance / KPIs)
--
-- IMPORTANT (data quality): the corrected source excludes the 16 placeholder
-- "Pine Grove Middle School" records. Pine Grove is NOT an SCSD school -- it
-- belongs to East Syracuse Minoa CSD and entered the raw roster only as a
-- field-trip location string ("OCM SKATE PINE GROVE MS") with no real student
-- data. The silver layer guards (by school name) against any such rows
-- re-entering. Grand total served = 403.
--
-- PRIVACY: real Student IDs and names are de-identified in the published
-- dataset (surrogate keys STU0001 ...; names removed) -- no FERPA PII.
--
-- Author: Daquan Morrison | SCSD Education Data Analysis | Thrivora Holdings
-- =============================================================================

-- ---------------------------------------------------------------------------
-- BRONZE: raw landing table (load goodlife_long_data.csv into this shape)
-- ---------------------------------------------------------------------------
-- SQLite load example (run in the sqlite3 shell from the project root):
--   .mode csv
--   .import --skip 1 data/goodlife_long_data.csv bronze_goodlife_long
--
DROP TABLE IF EXISTS bronze_goodlife_long;
CREATE TABLE bronze_goodlife_long (
    year                TEXT,    -- 'SY 2023-24' | 'SY 2024-25' | 'SY 2025-26'
    student_id          TEXT,    -- de-identified surrogate key (STU0001 ...)
    name                TEXT,
    school              TEXT,
    mp1                 REAL,
    mp2                 REAL,
    mp3                 REAL,
    mp4                 REAL,
    year_avg            REAL,
    passing_all_mps     INTEGER, -- 1 if passed all marking periods (>=65), else 0
    oss                 INTEGER, -- out-of-school suspension incidents
    iss                 INTEGER, -- in-school suspension incidents
    total_susp          INTEGER,
    oss_days            REAL,
    iss_days            REAL,
    total_susp_days     REAL,
    susp_1plus          INTEGER, -- 1 if student had 1+ suspensions
    susp_2plus          INTEGER, -- 1 if student had 2+ suspensions
    days_5plus          INTEGER, -- 1 if 5+ suspended days
    attendance_rate     REAL,    -- 0.0 - 1.0
    attendance_90       INTEGER  -- 1 if attendance >= 90%
);

-- ---------------------------------------------------------------------------
-- SILVER: typed + cleaned student-year grain
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS silver_student_year;
CREATE VIEW silver_student_year AS
SELECT
    CAST(year AS TEXT)                       AS school_year,
    CAST(student_id AS TEXT)                 AS student_id,
    name,
    school,
    CAST(NULLIF(mp1, '') AS REAL)            AS mp1,
    CAST(NULLIF(mp2, '') AS REAL)            AS mp2,
    CAST(NULLIF(mp3, '') AS REAL)            AS mp3,
    CAST(NULLIF(mp4, '') AS REAL)            AS mp4,
    CAST(NULLIF(year_avg, '') AS REAL)       AS year_avg,
    CAST(COALESCE(passing_all_mps, 0) AS INTEGER) AS passing_all_mps,
    CAST(COALESCE(oss, 0) AS INTEGER)        AS oss,
    CAST(COALESCE(iss, 0) AS INTEGER)        AS iss,
    CAST(COALESCE(total_susp, 0) AS INTEGER) AS total_susp,
    CAST(COALESCE(oss_days, 0) AS REAL)      AS oss_days,
    CAST(COALESCE(iss_days, 0) AS REAL)      AS iss_days,
    CAST(COALESCE(total_susp_days, 0) AS REAL) AS total_susp_days,
    CAST(COALESCE(susp_1plus, 0) AS INTEGER) AS susp_1plus,
    CAST(COALESCE(susp_2plus, 0) AS INTEGER) AS susp_2plus,
    CAST(COALESCE(days_5plus, 0) AS INTEGER) AS days_5plus,
    CAST(NULLIF(attendance_rate, '') AS REAL) AS attendance_rate,
    CAST(COALESCE(attendance_90, 0) AS INTEGER) AS attendance_90
FROM bronze_goodlife_long
WHERE school IS NOT NULL
  AND school <> 'Non-SCSD / Unknown'
  AND school NOT LIKE '%Pine Grove%';   -- guard against placeholder artifact

-- ---------------------------------------------------------------------------
-- GOLD 1: Students served by school by year (matches workbook Summary table 1)
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS gold_students_served;
CREATE VIEW gold_students_served AS
SELECT
    school,
    SUM(CASE WHEN school_year = 'SY 2023-24' THEN 1 ELSE 0 END) AS sy_2023_24,
    SUM(CASE WHEN school_year = 'SY 2024-25' THEN 1 ELSE 0 END) AS sy_2024_25,
    SUM(CASE WHEN school_year = 'SY 2025-26' THEN 1 ELSE 0 END) AS sy_2025_26,
    COUNT(*)                                                    AS total
FROM silver_student_year
GROUP BY school
ORDER BY school;

-- Grand total check (should return 403):
-- SELECT COUNT(*) AS grand_total FROM silver_student_year;

-- ---------------------------------------------------------------------------
-- GOLD 2: Suspensions by school by year
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS gold_suspensions;
CREATE VIEW gold_suspensions AS
SELECT
    school,
    school_year,
    COUNT(*)                 AS students,
    SUM(oss)                 AS oss,
    SUM(iss)                 AS iss,
    SUM(oss_days)            AS oss_days,
    SUM(iss_days)            AS iss_days,
    SUM(total_susp_days)     AS total_susp_days,
    SUM(susp_1plus)          AS students_1plus,
    SUM(susp_2plus)          AS students_2plus,
    SUM(days_5plus)          AS students_5plus_days
FROM silver_student_year
GROUP BY school, school_year
ORDER BY school_year, school;

-- ---------------------------------------------------------------------------
-- GOLD 3: Grades / academic performance by school by year
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS gold_grades;
CREATE VIEW gold_grades AS
SELECT
    school,
    school_year,
    COUNT(*)                                          AS students,
    ROUND(AVG(year_avg), 2)                           AS avg_year_score,
    SUM(passing_all_mps)                              AS n_passing_all_mps,
    ROUND(100.0 * SUM(passing_all_mps)
          / SUM(CASE WHEN year_avg IS NOT NULL THEN 1 ELSE 0 END), 1)
                                                      AS pct_passing_all_mps
FROM silver_student_year
GROUP BY school, school_year
ORDER BY school_year, school;

-- ---------------------------------------------------------------------------
-- GOLD 4: Attendance by school by year
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS gold_attendance;
CREATE VIEW gold_attendance AS
SELECT
    school,
    school_year,
    COUNT(*)                                          AS students,
    ROUND(AVG(attendance_rate), 3)                    AS avg_attendance_rate,
    SUM(attendance_90)                                AS n_at_or_above_90,
    ROUND(100.0 * SUM(attendance_90)
          / SUM(CASE WHEN attendance_rate IS NOT NULL THEN 1 ELSE 0 END), 1)
                                                      AS pct_at_or_above_90
FROM silver_student_year
GROUP BY school, school_year
ORDER BY school_year, school;

-- ---------------------------------------------------------------------------
-- GOLD 5: District-wide contract KPIs by year
-- ---------------------------------------------------------------------------
DROP VIEW IF EXISTS gold_contract_kpis;
CREATE VIEW gold_contract_kpis AS
SELECT
    school_year,
    COUNT(*)                                          AS students_served,
    COUNT(DISTINCT school)                            AS schools_served,
    ROUND(AVG(year_avg), 2)                           AS avg_year_score,
    ROUND(100.0 * SUM(passing_all_mps)
          / SUM(CASE WHEN year_avg IS NOT NULL THEN 1 ELSE 0 END), 1)
                                                      AS pct_passing_all_mps,
    ROUND(AVG(attendance_rate), 3)                    AS avg_attendance_rate,
    ROUND(100.0 * SUM(attendance_90)
          / SUM(CASE WHEN attendance_rate IS NOT NULL THEN 1 ELSE 0 END), 1)
                                                      AS pct_attendance_90,
    SUM(susp_1plus)                                   AS students_1plus_susp,
    SUM(total_susp_days)                              AS total_susp_days
FROM silver_student_year
GROUP BY school_year
ORDER BY school_year;

-- =============================================================================
-- Quick verification queries
-- =============================================================================
-- SELECT * FROM gold_students_served;
-- SELECT * FROM gold_contract_kpis;
-- SELECT COUNT(*) AS grand_total FROM silver_student_year;   -- expect 403
