-- ============================================================================
-- Charter School Audit Engine — SQL translation
-- Source: SCSD Charter Invoice / SchoolTool (ST) enrollment vs. monthly charter
--         rosters (Southside, SASCCS, SAS, OnTech) + SPED disability records.
--
-- Originally built as 13 Mindex calculated fields on a visual join model.
-- This is the faithful, vendor-independent SQL version (runs on Fabric SQL
-- endpoint, Azure SQL, Snowflake, Postgres — adjust DATEDIFF dialect as noted).
--
-- Purpose: resolve student identity across systems, then validate enrollment
-- entry/exit dates and SPED status against SchoolTool to flag billing/finance
-- discrepancies for charter invoice review (NYS 3-day rule context).
--
-- JOIN MODEL (from production notes):
--   * ST enrollment (Charter Invoice Enrollment) is the spine.
--   * Each monthly roster LEFT JOINs to ST on NYSSIS, cast to text on both sides
--     (toString) to avoid String-vs-Integer key mismatch that produced all-null
--     rows. Use RIGHT JOIN (roster-driven) when you want only rostered students.
--   * Removed the secondary (studentid = Student ID) join clause — match on
--     NYSSIS only; StudentID columns are validation INPUTS, not join keys.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- Layer 1: BRONZE — raw inputs joined on NYSSIS (cast to text both sides)
-- ----------------------------------------------------------------------------
WITH joined AS (
    SELECT
        st.student_universalstudentid,
        st.studentid,
        st.firstname,
        st.lastname,
        st.person_dob,
        st.grade,
        st.schoolname,
        st.st_studentenrollment_startdate,
        st.st_studentenrollment_enddate,
        st.schoolyear_disability,
        st.studentid_nyssis_export        AS studentid_aa_nyssis,  -- studentid[st_custom_export_aa_studentnyssis]
        st.st_nyssis_id,

        -- Roster Student IDs (validation inputs, one per charter school roster)
        ss.student_id                     AS southside_student_id,
        ss.st_entry_date                  AS southside_st_entry_date,
        ss.st_exit_date                   AS southside_st_exit_date,
        ss.iep_code                       AS southside_iep_code,
        sasccs.student_id                 AS sasccs_student_id,
        sas.student_id                    AS sas_student_id,
        ot.sped_student_id                AS ontech_sped_student_id
    FROM st_enrollment st
    LEFT JOIN monthly_southside_roster ss
           ON CAST(st.st_nyssis_id AS VARCHAR(50)) = CAST(ss.nyssis_id AS VARCHAR(50))
    LEFT JOIN monthly_sasccs_roster sasccs
           ON CAST(st.st_nyssis_id AS VARCHAR(50)) = CAST(sasccs.nyssis_id AS VARCHAR(50))
    LEFT JOIN monthly_sas_roster sas
           ON CAST(st.st_nyssis_id AS VARCHAR(50)) = CAST(sas.nyssis_id AS VARCHAR(50))
    LEFT JOIN monthly_ontech_roster ot
           ON CAST(st.st_nyssis_id AS VARCHAR(50)) = CAST(ot.nyssis_id AS VARCHAR(50))
),

-- ----------------------------------------------------------------------------
-- Layer 2: SILVER — identity resolution + base flags (fields 1-5)
-- ----------------------------------------------------------------------------
resolved AS (
    SELECT
        joined.*,

        -- (1) resolved_studentid: ST waterfall ONLY (no roster IDs in resolution)
        COALESCE(
            NULLIF(studentid, ''),
            NULLIF(student_universalstudentid, ''),
            NULLIF(studentid_aa_nyssis, ''),
            ''
        ) AS resolved_studentid,

        -- (3) id_source
        CASE
            WHEN studentid              IS NOT NULL THEN 'SchoolTool StudentID'
            WHEN student_universalstudentid IS NOT NULL THEN 'Universal StudentID'
            WHEN studentid_aa_nyssis    IS NOT NULL THEN 'NYSSID Mapping'
            ELSE 'No Match'
        END AS id_source,

        -- (4) is_iep_flag
        CASE WHEN schoolyear_disability IS NULL THEN 'Gen Ed' ELSE 'IEP' END AS is_iep_flag,

        -- (5) roster_vs_disability_check (Southside roster vs ST disability record)
        CASE
            WHEN southside_iep_code IS NULL     AND schoolyear_disability IS NULL     THEN 'Consistent - Gen Ed'
            WHEN southside_iep_code IS NOT NULL AND schoolyear_disability IS NOT NULL THEN 'Consistent - SPED'
            WHEN southside_iep_code IS NULL     AND schoolyear_disability IS NOT NULL THEN 'Mismatch - Roster Gen Ed but Disability Record IEP'
            WHEN southside_iep_code IS NOT NULL AND schoolyear_disability IS NULL     THEN 'Mismatch - Roster SPED but No Disability Record'
            ELSE 'Unknown'
        END AS roster_vs_disability_check
    FROM joined
),

-- ----------------------------------------------------------------------------
-- Layer 3: SILVER — id_match_status + date diffs (fields 2, 6-10)
--   DATEDIFF NOTE: T-SQL/Fabric/Synapse = DATEDIFF(DAY, start, end).
--   Snowflake     = DATEDIFF('day', start, end).  Postgres = (end - start).
--   Sign convention matches Mindex dateDiff(ST, roster, "DD") = roster - ST.
-- ----------------------------------------------------------------------------
audited AS (
    SELECT
        resolved.*,

        -- (2) id_match_status: compare resolved ID against each roster's Student ID
        CASE
            WHEN resolved_studentid = '' THEN 'Missing Student ID Match'
            WHEN southside_student_id IS NOT NULL AND southside_student_id = resolved_studentid THEN 'Matched - Southside'
            WHEN sasccs_student_id    IS NOT NULL AND sasccs_student_id    = resolved_studentid THEN 'Matched - SASCCS'
            WHEN sas_student_id       IS NOT NULL AND sas_student_id       = resolved_studentid THEN 'Matched - SAS'
            WHEN ontech_sped_student_id IS NOT NULL AND ontech_sped_student_id = resolved_studentid THEN 'Matched - OnTech'
            ELSE 'No Roster Match'
        END AS id_match_status,

        -- (6) entry_diff_days  (roster entry minus ST entry, in days)
        CASE
            WHEN southside_st_entry_date IS NULL OR st_studentenrollment_startdate IS NULL THEN NULL
            ELSE DATEDIFF(DAY, st_studentenrollment_startdate, southside_st_entry_date)
        END AS entry_diff_days,

        -- (7) entry_status_category (tolerance +/- 5 days; > 5 = finance risk)
        CASE
            WHEN southside_st_entry_date IS NULL OR st_studentenrollment_startdate IS NULL THEN 'Missing Date'
            WHEN DATEDIFF(DAY, st_studentenrollment_startdate, southside_st_entry_date) = 0 THEN 'Exact Match'
            WHEN DATEDIFF(DAY, st_studentenrollment_startdate, southside_st_entry_date) BETWEEN -5 AND 5 THEN 'Within Tolerance'
            WHEN DATEDIFF(DAY, st_studentenrollment_startdate, southside_st_entry_date) > 5 THEN 'Finance Risk'
            ELSE 'Finance Risk'
        END AS entry_status_category,

        -- (8) exit_diff_days
        CASE
            WHEN southside_st_exit_date IS NULL OR st_studentenrollment_enddate IS NULL THEN NULL
            ELSE DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date)
        END AS exit_diff_days,

        -- (9) exit_status_category
        CASE
            WHEN southside_st_exit_date IS NULL OR st_studentenrollment_enddate IS NULL THEN 'Missing Date'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) = 0 THEN 'Exact Match'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) BETWEEN -5 AND 5 THEN 'Within Tolerance'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) > 5 THEN 'Finance Risk'
            ELSE 'Finance Risk'
        END AS exit_status_category,

        -- (10) exit_issue_type
        CASE
            WHEN southside_st_exit_date IS NULL     AND st_studentenrollment_enddate IS NOT NULL THEN 'Missing Roster Exit'
            WHEN southside_st_exit_date IS NOT NULL AND st_studentenrollment_enddate IS NULL     THEN 'Missing ST Exit'
            WHEN southside_st_exit_date IS NULL     AND st_studentenrollment_enddate IS NULL     THEN 'Both Dates Missing'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) = 0 THEN 'No Issue'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) > 5 THEN 'Roster Exit Late'
            WHEN DATEDIFF(DAY, st_studentenrollment_enddate, southside_st_exit_date) < -5 THEN 'Roster Exit Early'
            ELSE 'Within Tolerance'
        END AS exit_issue_type
    FROM resolved
)

-- ----------------------------------------------------------------------------
-- Layer 4: GOLD — final validation, severity, recommended action (fields 11-13)
-- ----------------------------------------------------------------------------
SELECT
    resolved_studentid,
    id_match_status,
    id_source,
    is_iep_flag,
    roster_vs_disability_check,
    entry_diff_days,
    entry_status_category,
    exit_diff_days,
    exit_status_category,
    exit_issue_type,

    -- (11) final_validation_status
    CASE
        WHEN id_match_status = 'Missing Student ID Match' THEN 'Invalid - No ID Match'
        WHEN entry_status_category = 'Missing Date' OR exit_status_category = 'Missing Date' THEN 'Incomplete - Missing Dates'
        WHEN entry_status_category = 'Exact Match' AND exit_status_category = 'Exact Match' THEN 'Valid'
        WHEN entry_status_category = 'Within Tolerance' OR exit_status_category = 'Within Tolerance' THEN 'Valid with Warning'
        ELSE 'Invalid - Date Mismatch'
    END AS final_validation_status,

    -- (12) severity
    CASE
        WHEN CASE
                WHEN id_match_status = 'Missing Student ID Match' THEN 'Invalid - No ID Match'
                WHEN entry_status_category = 'Missing Date' OR exit_status_category = 'Missing Date' THEN 'Incomplete - Missing Dates'
                WHEN entry_status_category = 'Exact Match' AND exit_status_category = 'Exact Match' THEN 'Valid'
                WHEN entry_status_category = 'Within Tolerance' OR exit_status_category = 'Within Tolerance' THEN 'Valid with Warning'
                ELSE 'Invalid - Date Mismatch'
             END = 'Invalid - No ID Match'      THEN 'Critical'
        WHEN entry_status_category = 'Missing Date' OR exit_status_category = 'Missing Date' THEN 'High'
        WHEN entry_status_category = 'Within Tolerance' OR exit_status_category = 'Within Tolerance' THEN 'Low'
        WHEN entry_status_category = 'Exact Match' AND exit_status_category = 'Exact Match' THEN 'None'
        ELSE 'High'
    END AS severity,

    -- (13) recommended_action
    CASE
        WHEN exit_issue_type = 'Missing Roster Exit' THEN 'Add exit date to Southside roster'
        WHEN exit_issue_type = 'Missing ST Exit'     THEN 'Verify exit date in StudentTracker'
        WHEN exit_issue_type = 'Both Dates Missing'  THEN 'Investigate enrollment record - both exit dates absent'
        WHEN exit_issue_type = 'Roster Exit Late'    THEN 'Review late exit discrepancy with billing team'
        WHEN exit_issue_type = 'Roster Exit Early'   THEN 'Review early exit discrepancy with billing team'
        WHEN entry_status_category = 'Finance Risk'  THEN 'Do not approve billing prior to ST entry'
        WHEN entry_status_category = 'Missing Date'  THEN 'Add entry date to enrollment record'
        ELSE 'No Action Required'
    END AS recommended_action,

    -- carry through identity / enrollment context for the dashboard
    student_universalstudentid,
    studentid,
    firstname,
    lastname,
    person_dob,
    grade,
    schoolname,
    st_studentenrollment_startdate,
    st_studentenrollment_enddate,
    schoolyear_disability,
    st_nyssis_id
FROM audited;
