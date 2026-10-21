-- ============================================================================
-- Educational Analyst Interview Prep — advanced SQL patterns
-- Dialect: T-SQL / Microsoft Fabric / Azure SQL.
-- Uses the source tables and columns documented in charter_audit_engine.sql.
-- Run each query separately. These are query examples, not production reports:
-- confirm the reporting period, student grain, and local business rules first.
-- No student-level data is included in this file.
-- ============================================================================

-- 1. Reconcile each charter roster to SchoolTool.
-- Skills: UNION ALL, roster-driven LEFT JOIN, conditional distinct counts,
-- and safe key normalization. Keep Student IDs as comparison inputs, not keys.
WITH roster_rows AS (
    SELECT 'Southside' AS charter, nyssis_id, student_id
    FROM monthly_southside_roster
    UNION ALL
    SELECT 'SASCCS', nyssis_id, student_id
    FROM monthly_sasccs_roster
    UNION ALL
    SELECT 'SAS', nyssis_id, student_id
    FROM monthly_sas_roster
    UNION ALL
    SELECT 'OnTech', nyssis_id, sped_student_id
    FROM monthly_ontech_roster
),
schooltool_keys AS (
    SELECT DISTINCT CAST(st_nyssis_id AS VARCHAR(50)) AS normalized_nyssis_id
    FROM st_enrollment
    WHERE st_nyssis_id IS NOT NULL
),
roster_matches AS (
    SELECT
        r.charter,
        r.nyssis_id,
        r.student_id AS roster_student_id,
        st.normalized_nyssis_id AS schooltool_nyssis_id
    FROM roster_rows r
    LEFT JOIN schooltool_keys st
        ON CAST(r.nyssis_id AS VARCHAR(50)) = st.normalized_nyssis_id
)
SELECT
    charter,
    COUNT(*) AS roster_rows,
    COUNT(DISTINCT nyssis_id) AS distinct_rostered_students,
    COUNT(DISTINCT CASE WHEN schooltool_nyssis_id IS NOT NULL
                        THEN nyssis_id END) AS matched_students,
    COUNT(DISTINCT CASE WHEN schooltool_nyssis_id IS NULL
                        THEN nyssis_id END) AS unmatched_students,
    CAST(
        100.0 * COUNT(DISTINCT CASE WHEN schooltool_nyssis_id IS NOT NULL
                                    THEN nyssis_id END)
        / NULLIF(COUNT(DISTINCT nyssis_id), 0)
        AS DECIMAL(5, 2)
    ) AS match_rate_pct
FROM roster_matches
GROUP BY charter
ORDER BY charter;

-- 2. Find gaps and overlaps between consecutive SchoolTool enrollment records.
-- Skills: LAG, window partitions, date arithmetic, and anomaly classification.
-- A null prior end date is not classified because it may represent an active
-- enrollment; clarify whether same-day transitions count as gaps locally.
WITH enrollment_sequence AS (
    SELECT
        student_universalstudentid,
        studentid,
        schoolname,
        st_studentenrollment_startdate,
        st_studentenrollment_enddate,
        LAG(st_studentenrollment_enddate) OVER (
            PARTITION BY student_universalstudentid
            ORDER BY st_studentenrollment_startdate, st_studentenrollment_enddate
        ) AS prior_enddate
    FROM st_enrollment
    WHERE student_universalstudentid IS NOT NULL
      AND st_studentenrollment_startdate IS NOT NULL
),
classified AS (
    SELECT
        *,
        DATEDIFF(DAY, prior_enddate, st_studentenrollment_startdate) AS days_between,
        CASE
            WHEN st_studentenrollment_startdate < prior_enddate THEN 'Overlap'
            WHEN st_studentenrollment_startdate > prior_enddate THEN 'Gap'
            ELSE 'Same-day transition'
        END AS transition_type
    FROM enrollment_sequence
    WHERE prior_enddate IS NOT NULL
)
SELECT
    student_universalstudentid,
    studentid,
    schoolname,
    prior_enddate,
    st_studentenrollment_startdate,
    days_between,
    transition_type
FROM classified
WHERE transition_type IN ('Gap', 'Overlap')
ORDER BY student_universalstudentid, st_studentenrollment_startdate;

-- 3. Summarize absolute Southside roster-vs-SchoolTool date variance.
-- Skills: CTEs, aggregates, PERCENTILE_CONT, and windowed median.
-- Median is returned in days alongside average; missing dates are excluded.
WITH date_variance AS (
    SELECT
        st.schoolname,
        DATEDIFF(
            DAY,
            st.st_studentenrollment_startdate,
            roster.st_entry_date
        ) AS entry_diff_days
    FROM st_enrollment st
    INNER JOIN monthly_southside_roster roster
        ON CAST(st.st_nyssis_id AS VARCHAR(50)) =
           CAST(roster.nyssis_id AS VARCHAR(50))
    WHERE st.st_studentenrollment_startdate IS NOT NULL
      AND roster.st_entry_date IS NOT NULL
),
variance_stats AS (
    SELECT
        schoolname,
        COUNT(*) AS compared_rows,
        AVG(CAST(ABS(entry_diff_days) AS DECIMAL(10, 2))) AS avg_abs_variance_days,
        SUM(CASE WHEN ABS(entry_diff_days) > 5 THEN 1 ELSE 0 END) AS over_5_day_rows
    FROM date_variance
    GROUP BY schoolname
),
variance_medians AS (
    SELECT DISTINCT
        schoolname,
        PERCENTILE_CONT(0.5) WITHIN GROUP (
            ORDER BY ABS(entry_diff_days)
        ) OVER (PARTITION BY schoolname) AS median_abs_variance_days
    FROM date_variance
)
SELECT
    stats.schoolname,
    stats.compared_rows,
    stats.avg_abs_variance_days,
    medians.median_abs_variance_days,
    stats.over_5_day_rows,
    CAST(100.0 * stats.over_5_day_rows / NULLIF(stats.compared_rows, 0)
         AS DECIMAL(5, 2)) AS over_5_day_pct
FROM variance_stats stats
INNER JOIN variance_medians medians
    ON stats.schoolname = medians.schoolname
ORDER BY over_5_day_pct DESC;

-- 4. Detect duplicate or conflicting SchoolTool NYSSIS mappings.
-- Skills: data-quality profiling, COUNT DISTINCT, and HAVING.
-- Review duplicates before joining: they can multiply roster rows.
SELECT
    CAST(st_nyssis_id AS VARCHAR(50)) AS normalized_nyssis_id,
    COUNT(*) AS enrollment_rows,
    COUNT(DISTINCT studentid) AS distinct_student_ids,
    COUNT(DISTINCT student_universalstudentid) AS distinct_universal_ids
FROM st_enrollment
WHERE st_nyssis_id IS NOT NULL
GROUP BY CAST(st_nyssis_id AS VARCHAR(50))
HAVING COUNT(*) > 1
    OR COUNT(DISTINCT studentid) > 1
    OR COUNT(DISTINCT student_universalstudentid) > 1
ORDER BY enrollment_rows DESC, normalized_nyssis_id;

-- 5. Compare disability-record representation by school and grade.
-- Skills: conditional aggregation, distinct counts, NULLIF, and rate calculation.
-- The source column indicates a disability record; it does not establish service
-- eligibility or identify a particular program.
SELECT
    schoolname,
    grade,
    COUNT(DISTINCT studentid) AS students,
    COUNT(DISTINCT CASE WHEN schoolyear_disability IS NOT NULL
                        THEN studentid END) AS students_with_disability_record,
    CAST(
        100.0 * COUNT(DISTINCT CASE WHEN schoolyear_disability IS NOT NULL
                                    THEN studentid END)
        / NULLIF(COUNT(DISTINCT studentid), 0)
        AS DECIMAL(5, 2)
    ) AS disability_record_pct
FROM st_enrollment
WHERE studentid IS NOT NULL
GROUP BY schoolname, grade
ORDER BY schoolname, grade;
