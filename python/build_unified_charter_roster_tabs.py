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
#   2. Open terminal / command prompt in that folder
#   3. Run: python build_unified_charter_roster_tabs.py
#
# REQUIREMENTS (install once):
#   pip install pandas openpyxl xlsxwriter
#
# OUTPUT: all_charters_unified_roster_25-26_REBUILT.xlsx (5 tabs)
#   Tab 1 — All Charters Unified     : master combined data (Student IDs patched from ST)
#   Tab 2 — Formula Version          : NYSSIS lookup formulas vs Update Log
#   Tab 3 — Copy & Paste             : plain values, paste into FY 25-26 master
#   Tab 4 — Not Found in ST          : students with no NYSSIS / not in ST
#   Tab 5 — Update Log               : 15 still-unresolved students pending manual lookup
# ============================================================================

import os
import sys
import re
import warnings
import pandas as pd
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarnings("ignore")

# ============================================================================
# SETTINGS
# ============================================================================
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
DATA_DIR      = SCRIPT_DIR
OUTPUT_PATH   = os.path.join(DATA_DIR, "all_charters_unified_roster_25-26_REBUILT.xlsx")
INVOICE_MONTH = "2026-06-01"

FILE_SOUTHSIDE = os.path.join(DATA_DIR, "southside_291462_unified-roster.xlsx")
FILE_SAS       = os.path.join(DATA_DIR, "sas_291463_unified-roster.xlsx")
FILE_SASCCS    = os.path.join(DATA_DIR, "sasccs_291464_unified-roster.xlsx")
FILE_ONTECH    = os.path.join(DATA_DIR, "ontech_final_unified-roster.xlsx")

TEMPLATE_COLS = [
    "Charter_Source Ticket #", "Charter Month", "Source_File", "NYSSIS ID",
    "Student ID", "Last Name", "First Name", "Grade Level", "DOB",
    "School Name", "Enroll Date", "Withdraw Date", "SCSD FTE",
    "Student Status", "Notes for Patrick", "SPED_Flag"
]

# ============================================================================
# VERIFY FILES EXIST
# ============================================================================
for f in [FILE_SOUTHSIDE, FILE_SAS, FILE_SASCCS, FILE_ONTECH]:
    if not os.path.exists(f):
        print(f"ERROR: File not found -> {f}")
        print("Make sure all 4 xlsx files are in the same folder as this script.")
        sys.exit(1)

# ============================================================================
# HELPERS
# ============================================================================
def clean_id(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(r'\.0+$', '', regex=True)
    s = s.replace(['nan','NaN','NULL','null','','None'], pd.NA)
    return s

def clean_fte(series):
    s = series.astype(str).str.strip().str.replace('%', '', regex=False)
    return pd.to_numeric(s, errors='coerce')

def parse_date(series):
    out = []
    for v in series:
        if pd.isna(v):
            out.append(pd.NaT)
            continue
        if isinstance(v, (datetime, date)):
            out.append(pd.Timestamp(v))
            continue
        s = str(v).strip()
        if s in ('', 'nan', 'NaT', '0', 'NA', 'N/A', 'NULL'):
            out.append(pd.NaT)
            continue
        # Excel serial number — valid 2020-2030 serials ~43831-49006
        try:
            num = float(s)
            if 40000 < num < 80000:
                out.append(pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(num)))
                continue
        except:
            pass
        parsed = pd.NaT
        for fmt in ('%m/%d/%Y','%Y-%m-%d','%d/%m/%Y','%m-%d-%Y'):
            try:
                parsed = pd.Timestamp(datetime.strptime(s, fmt))
                break
            except:
                continue
        out.append(parsed)
    return pd.Series(out, index=series.index)

def split_full_name(series):
    last  = series.str.extract(r'^([^,]+)')[0].str.strip()
    first = series.str.replace(r'^[^,]+,?\s*', '', regex=True).str.strip()
    return last, first

def normalize_str(series):
    return series.astype(str).str.strip().replace(
        ['nan','NaN','NULL','null','None',''], pd.NA)

# ============================================================================
# BUILD EACH CHARTER
# ============================================================================
def build_southside():
    print("  Reading southside_291462_unified-roster.xlsx ...")
    df = pd.read_excel(FILE_SOUTHSIDE, sheet_name="Unified Gen-SPED Roster", header=0, dtype=str)
    df = df[df["NYSSIS ID"].notna() | df["Student Last Name"].notna()].copy()
    return pd.DataFrame({
        "Charter_Source Ticket #": "Southside 291462",
        "Charter Month":           INVOICE_MONTH,
        "Source_File":             "southside_291462_unified-roster.xlsx",
        "NYSSIS ID":               clean_id(df["NYSSIS ID"]),
        "Student ID":              clean_id(df["Student ID"]),
        "Last Name":               normalize_str(df["Student Last Name"]),
        "First Name":              normalize_str(df["Student First Name"]),
        "Grade Level":             normalize_str(df["Grade"]),
        "DOB":                     parse_date(df["Birth Date"]),
        "School Name":             normalize_str(df["School Name"]),
        "Enroll Date":             parse_date(df["Entered"]),
        "Withdraw Date":           parse_date(df["Withdrew"]),
        "SCSD FTE":                clean_fte(df["Southside FTE"]),
        "Student Status":          normalize_str(df["Student Status"]),
        "Notes for Patrick":       normalize_str(df["Notes for Patrick"]),
        "SPED_Flag":               normalize_str(df["SPED_Flag"]),
    })

def build_sas():
    print("  Reading sas_291463_unified-roster.xlsx ...")
    df = pd.read_excel(FILE_SAS, sheet_name="Unified Gen-SPED Roster", header=0, dtype=str)
    df = df[df["NYSSIS ID"].notna() | df["Last Name"].notna()].copy()
    return pd.DataFrame({
        "Charter_Source Ticket #": "SAS 291463",
        "Charter Month":           INVOICE_MONTH,
        "Source_File":             "sas_291463_unified-roster.xlsx",
        "NYSSIS ID":               clean_id(df["NYSSIS ID"]),
        "Student ID":              clean_id(df["Student ID"]),
        "Last Name":               normalize_str(df["Last Name"]),
        "First Name":              normalize_str(df["First Name"]),
        "Grade Level":             normalize_str(df["Current Grade Level"]),
        "DOB":                     parse_date(df["Date of Birth"]),
        "School Name":             normalize_str(df["School Name"]),
        "Enroll Date":             parse_date(df["Entry Date"]),
        "Withdraw Date":           parse_date(df["Leave Date"]),
        "SCSD FTE":                clean_fte(df["FTE"]),
        "Student Status":          normalize_str(df["Student Status"]),
        "Notes for Patrick":       pd.NA,
        "SPED_Flag":               normalize_str(df["SPED_Flag"]),
    })

def build_sasccs():
    print("  Reading sasccs_291464_unified-roster.xlsx ...")
    df = pd.read_excel(FILE_SASCCS, sheet_name="Unified Gen-SPED Roster", header=0, dtype=str)
    nyssis_col = df.columns[0]
    full_name_col = df.columns[2]
    df = df[df[nyssis_col].notna() | df[full_name_col].notna()].copy()
    last, first = split_full_name(df["Full Name"])
    return pd.DataFrame({
        "Charter_Source Ticket #": "SASCCS 291464",
        "Charter Month":           INVOICE_MONTH,
        "Source_File":             "sasccs_291464_unified-roster.xlsx",
        "NYSSIS ID":               clean_id(df["NYSSIS ID"]),
        "Student ID":              clean_id(df["Student ID"]),
        "Last Name":               last,
        "First Name":              first,
        "Grade Level":             normalize_str(df["Current Grade Level"]),
        "DOB":                     parse_date(df["Date of Birth"]),
        "School Name":             normalize_str(df["School Name"]),
        "Enroll Date":             parse_date(df["Entry Date"]),
        "Withdraw Date":           parse_date(df["Leave Date"]),
        "SCSD FTE":                clean_fte(df["Status"]),
        "Student Status":          normalize_str(df["Student Status"]),
        "Notes for Patrick":       pd.NA,
        "SPED_Flag":               normalize_str(df["SPED_Flag"]),
    })

def build_ontech():
    print("  Reading ontech_final_unified-roster.xlsx ...")
    df = pd.read_excel(FILE_ONTECH, sheet_name="Unified Gen-SPED Roster", header=0, dtype=str)
    df = df[df["NYSSIS ID"].notna() | df["Last Name"].notna()].copy()
    sid_col = "Studemt ID" if "Studemt ID" in df.columns else "Student ID"
    return pd.DataFrame({
        "Charter_Source Ticket #": "OnTech 291465",
        "Charter Month":           INVOICE_MONTH,
        "Source_File":             "ontech_final_unified-roster.xlsx",
        "NYSSIS ID":               clean_id(df["NYSSIS ID"]),
        "Student ID":              clean_id(df[sid_col]),
        "Last Name":               normalize_str(df["Last Name"]),
        "First Name":              normalize_str(df["First Name"]),
        "Grade Level":             normalize_str(df["Grade Level"]),
        "DOB":                     parse_date(df["DOB"]),
        "School Name":             "OnTECH Charter High School",
        "Enroll Date":             parse_date(df["Enroll Date"]),
        "Withdraw Date":           parse_date(df["Withdraw Date"]),
        "SCSD FTE":                clean_fte(df["OnTECH FTE"]),
        "Student Status":          pd.NA,
        "Notes for Patrick":       pd.NA,
        "SPED_Flag":               normalize_str(df["SPED_Flag"]),
    })

# ============================================================================
# COMBINE
# ============================================================================
print("\nBuilding unified roster...")
unified = pd.concat([
    build_southside(), build_sas(), build_sasccs(), build_ontech()
], ignore_index=True)[TEMPLATE_COLS]

for col in unified.select_dtypes(include="object").columns:
    unified[col] = unified[col].replace(['nan','NaN','NULL','null','None',''], pd.NA)

# ============================================================================
# STEP: PATCH STUDENT IDs AND NYSSIS IDs FROM ST ENROLL REPORTS
# Pass 1 — patch missing Student IDs  (NYSSIS/name/DOB → Student ID)
# Pass 2 — patch missing NYSSIS IDs   (Student ID/name/DOB → NYSSIS)
# ============================================================================
print("\nPatching Student IDs and NYSSIS IDs from ST Enroll reports...")

def load_st(filepath, sheet, source):
    df = pd.read_excel(filepath, sheet_name=sheet, header=0, dtype=str)
    return pd.DataFrame({
        "nyssis": df["Student_UniversalStudentID"].astype(str).str.strip().str.replace(r'\.0+$','',regex=True).replace(['nan','NULL','','0'], pd.NA),
        "sid":    df["StudentID"].astype(str).str.strip().str.replace(r'\.0+$','',regex=True).replace(['nan','NULL',''], pd.NA),
        "fname":  df["FirstName"].astype(str).str.strip().str.upper().replace(['nan','NULL',''], pd.NA),
        "lname":  df["LastName"].astype(str).str.strip().str.upper().replace(['nan','NULL',''], pd.NA),
        "dob":    df["Person_DOB"].astype(str).str[:10].replace(['nan','NULL','NaT'], pd.NA),
        "source": source
    })

st_all = pd.concat([
    load_st(FILE_SOUTHSIDE, "ST Enroll Export", "Southside ST"),
    load_st(FILE_SAS,       "ST Enroll Export", "SAS ST"),
    load_st(FILE_SASCCS,    "ST Enroll Export", "SASCCS ST"),
    load_st(FILE_ONTECH,    "ST Enroll Export", "OnTech ST"),
], ignore_index=True)

# Build NYSSIS→SID and Name→SID lookup dicts
nyssis_to_sid = {}
for _, row in st_all[st_all["nyssis"].notna() & st_all["sid"].notna()].iterrows():
    if row["nyssis"] not in nyssis_to_sid:
        nyssis_to_sid[row["nyssis"]] = row["sid"]

name_to_sid = {}  # key = (FNAME, LNAME)
for _, row in st_all[st_all["fname"].notna() & st_all["lname"].notna() & st_all["sid"].notna()].iterrows():
    key = (row["fname"], row["lname"])
    if key not in name_to_sid:
        name_to_sid[key] = row["sid"]

lastdob_to_sid = {}  # key = (LNAME, DOB)
for _, row in st_all[st_all["lname"].notna() & st_all["dob"].notna() & st_all["sid"].notna()].iterrows():
    key = (row["lname"], row["dob"])
    if key not in lastdob_to_sid:
        lastdob_to_sid[key] = row["sid"]

firstdob_counts = {}  # key = (FNAME, DOB) → count of unique SIDs
for _, row in st_all[st_all["fname"].notna() & st_all["dob"].notna() & st_all["sid"].notna()].iterrows():
    key = (row["fname"], row["dob"])
    firstdob_counts.setdefault(key, set()).add(row["sid"])
firstdob_to_sid = {k: list(v)[0] for k, v in firstdob_counts.items() if len(v) == 1}

patch_count = 0
for idx, row in unified.iterrows():
    cur_sid = row["Student ID"]
    if pd.notna(cur_sid) and cur_sid != "0":
        continue

    m_nyssis = str(row["NYSSIS ID"]).strip() if pd.notna(row["NYSSIS ID"]) else None
    m_fname  = str(row["First Name"]).strip().upper() if pd.notna(row["First Name"]) and str(row["First Name"]) not in ("NA","nan","0") else None
    m_lname  = str(row["Last Name"]).strip().upper()  if pd.notna(row["Last Name"])  and str(row["Last Name"])  not in ("NA","nan","0") else None
    m_dob    = str(row["DOB"])[:10] if pd.notna(row["DOB"]) else None
    if m_dob in ("NaT","nan","None",""): m_dob = None

    found_sid = None

    # Try 1: NYSSIS
    if m_nyssis and m_nyssis not in ("nan","0","NA"):
        found_sid = nyssis_to_sid.get(m_nyssis)

    # Try 2: First + Last
    if not found_sid and m_fname and m_lname and len(m_fname) > 1 and len(m_lname) > 1:
        found_sid = name_to_sid.get((m_fname, m_lname))

    # Try 3: Last + DOB
    if not found_sid and m_lname and m_dob and len(m_lname) > 1:
        found_sid = lastdob_to_sid.get((m_lname, m_dob))

    # Try 4: First + DOB (unique match only)
    if not found_sid and m_fname and m_dob and len(m_fname) > 1:
        found_sid = firstdob_to_sid.get((m_fname, m_dob))

    if found_sid:
        unified.at[idx, "Student ID"] = found_sid
        patch_count += 1

still_missing_sid = int((unified["Student ID"].isna() | (unified["Student ID"] == "0")).sum())
print(f"  Pass 1 — Student IDs patched from ST: {patch_count}")
print(f"  Still missing Student IDs:            {still_missing_sid}")

# ============================================================================
# PASS 2 — Patch missing NYSSIS IDs from ST using Student ID / name / DOB
# ============================================================================
# Build SID→NYSSIS lookup from ST
sid_to_nyssis = {}
for _, row in st_all[st_all["sid"].notna() & st_all["nyssis"].notna()].iterrows():
    if row["sid"] not in sid_to_nyssis:
        sid_to_nyssis[row["sid"]] = row["nyssis"]

name_to_nyssis = {}
for _, row in st_all[st_all["fname"].notna() & st_all["lname"].notna() & st_all["nyssis"].notna()].iterrows():
    key = (row["fname"], row["lname"])
    if key not in name_to_nyssis:
        name_to_nyssis[key] = row["nyssis"]

lastdob_to_nyssis = {}
for _, row in st_all[st_all["lname"].notna() & st_all["dob"].notna() & st_all["nyssis"].notna()].iterrows():
    key = (row["lname"], row["dob"])
    if key not in lastdob_to_nyssis:
        lastdob_to_nyssis[key] = row["nyssis"]

firstdob_nyssis_counts = {}
for _, row in st_all[st_all["fname"].notna() & st_all["dob"].notna() & st_all["nyssis"].notna()].iterrows():
    key = (row["fname"], row["dob"])
    firstdob_nyssis_counts.setdefault(key, set()).add(row["nyssis"])
firstdob_to_nyssis = {k: list(v)[0] for k, v in firstdob_nyssis_counts.items() if len(v) == 1}

nyssis_patch_count = 0
for idx, row in unified.iterrows():
    cur_nyssis = row["NYSSIS ID"]
    if pd.notna(cur_nyssis) and cur_nyssis != "0":
        continue

    m_sid   = str(row["Student ID"]).strip() if pd.notna(row["Student ID"]) and str(row["Student ID"]) not in ("nan","0","NA") else None
    m_fname = str(row["First Name"]).strip().upper() if pd.notna(row["First Name"]) and str(row["First Name"]) not in ("NA","nan","0") else None
    m_lname = str(row["Last Name"]).strip().upper()  if pd.notna(row["Last Name"])  and str(row["Last Name"])  not in ("NA","nan","0") else None
    m_dob   = str(row["DOB"])[:10] if pd.notna(row["DOB"]) else None
    if m_dob in ("NaT","nan","None",""): m_dob = None

    found_nyssis = None

    # Try 1: Student ID → NYSSIS
    if m_sid:
        found_nyssis = sid_to_nyssis.get(m_sid)

    # Try 2: First + Last → NYSSIS
    if not found_nyssis and m_fname and m_lname and len(m_fname)>1 and len(m_lname)>1:
        found_nyssis = name_to_nyssis.get((m_fname, m_lname))

    # Try 3: Last + DOB → NYSSIS
    if not found_nyssis and m_lname and m_dob and len(m_lname)>1:
        found_nyssis = lastdob_to_nyssis.get((m_lname, m_dob))

    # Try 4: First + DOB → NYSSIS (unique only)
    if not found_nyssis and m_fname and m_dob and len(m_fname)>1:
        found_nyssis = firstdob_to_nyssis.get((m_fname, m_dob))

    if found_nyssis:
        unified.at[idx, "NYSSIS ID"] = found_nyssis
        nyssis_patch_count += 1

still_missing_nyssis = int((unified["NYSSIS ID"].isna() | (unified["NYSSIS ID"] == "0")).sum())
print(f"  Pass 2 — NYSSIS IDs patched from ST:   {nyssis_patch_count}")
print(f"  Still missing NYSSIS IDs:              {still_missing_nyssis}")

still_missing = still_missing_sid
n = len(unified)

# Not Found in ST
not_found = unified[
    unified["NYSSIS ID"].isna() |
    (unified["NYSSIS ID"] == "0") |
    (unified["Student Status"] == "Not in ST System")
].copy()

# ============================================================================
# VALIDATION SUMMARY
# ============================================================================
print(f"\n{'='*52}")
print(f"UNIFIED ROSTER SUMMARY")
print(f"{'='*52}")
print(f"Total students: {n:,}")
print("\nRows per charter:")
print(unified.groupby("Charter_Source Ticket #").size().reset_index(name="rows").to_string(index=False))
print("\nFTE by charter:")
fte_summary = unified.groupby("Charter_Source Ticket #")["SCSD FTE"].sum().round(3).reset_index()
print(fte_summary.to_string(index=False))
print(f"\nTotal FTE: {unified['SCSD FTE'].sum():.3f}")
missing_nyssis = int((unified["NYSSIS ID"].isna() | (unified["NYSSIS ID"] == "0")).sum())
print(f"Missing NYSSIS: {missing_nyssis}")
print(f"Missing Student IDs: {still_missing}")
print(f"Not Found in ST tab: {len(not_found)} rows")

# ============================================================================
# WRITE EXCEL WITH OPENPYXL
# ============================================================================
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule

wb = Workbook()
wb.remove(wb.active)

CLR_DARK_BLUE  = "1F3864"
CLR_TEAL       = "17375E"
CLR_ORANGE     = "C55A11"
CLR_PURPLE     = "4B2D83"
CLR_GREEN      = "375623"
CLR_SPED_BG    = "FFF2CC"
CLR_GENED_BG   = "E2EFDA"
CLR_GOOD_BG    = "C6EFCE"
CLR_GOOD_FG    = "276221"
CLR_NOGOOD_BG  = "FFC7CE"
CLR_NOGOOD_FG  = "9C0006"
CLR_FLAG_BG    = "FFEB9C"
CLR_FLAG_FG    = "9C5700"

COL_WIDTHS = [18,13,34,13,12,14,12,10,12,32,12,12,9,16,28,12]

def make_hdr_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_hdr_font(bold=True):
    return Font(bold=bold, color="FFFFFF")

def style_header_row(ws, row_num, num_cols, fill_hex):
    fill = make_hdr_fill(fill_hex)
    font = make_hdr_font()
    for c in range(1, num_cols+1):
        cell = ws.cell(row=row_num, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="left", wrap_text=False)

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_df_to_sheet(ws, df, start_row=1, include_header=True):
    if include_header:
        for c_idx, col_name in enumerate(df.columns, 1):
            ws.cell(row=start_row, column=c_idx, value=col_name)
        start_row += 1
    for r_idx, row in enumerate(df.itertuples(index=False), start_row):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if not isinstance(val, str) and pd.isna(val):
                cell.value = None
            elif isinstance(val, pd.Timestamp):
                cell.value = val.to_pydatetime() if not pd.isna(val) else None
                cell.number_format = "MM/DD/YYYY"
            else:
                cell.value = val
    return start_row + len(df) - 1

def apply_sped_status_cf(ws, data_rows, sped_col=16, status_col=14):
    sc = get_column_letter(sped_col)
    stc = get_column_letter(status_col)
    sped_range   = f"{sc}2:{sc}{data_rows+1}"
    status_range = f"{stc}2:{stc}{data_rows+1}"
    ws.conditional_formatting.add(sped_range,
        FormulaRule(formula=[f'NOT(ISERROR(SEARCH("SPED",{sc}2)))'],
                    fill=PatternFill("solid", fgColor=CLR_SPED_BG)))
    ws.conditional_formatting.add(sped_range,
        FormulaRule(formula=[f'NOT(ISERROR(SEARCH("Gen Ed",{sc}2)))'],
                    fill=PatternFill("solid", fgColor=CLR_GENED_BG)))
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=[f'NOT(ISERROR(SEARCH("No Good",{stc}2)))'],
                    fill=PatternFill("solid", fgColor=CLR_NOGOOD_BG),
                    font=Font(color=CLR_NOGOOD_FG)))
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=[f'=AND(NOT(ISERROR(SEARCH("Good",{stc}2))),ISERROR(SEARCH("No Good",{stc}2)))'],
                    fill=PatternFill("solid", fgColor=CLR_GOOD_BG),
                    font=Font(color=CLR_GOOD_FG)))

# ============================================================================
# TAB 1 — All Charters Unified
# ============================================================================
ws1 = wb.create_sheet("All Charters Unified")
write_df_to_sheet(ws1, unified, start_row=1, include_header=True)
style_header_row(ws1, 1, len(TEMPLATE_COLS), CLR_DARK_BLUE)
ws1.freeze_panes = "A2"
set_col_widths(ws1, COL_WIDTHS)
apply_sped_status_cf(ws1, n)

# ============================================================================
# TAB 2 — Formula Version
# ============================================================================
ws2 = wb.create_sheet("Formula Version")
fv_cols = TEMPLATE_COLS + ["Resolved_NYSSIS", "Match_Source"]
for c_idx, col_name in enumerate(fv_cols, 1):
    ws2.cell(row=1, column=c_idx, value=col_name)
style_header_row(ws2, 1, len(fv_cols), CLR_TEAL)
write_df_to_sheet(ws2, unified, start_row=2, include_header=False)
for i in range(2, n+2):
    ws2.cell(row=i, column=17).value = (
        f'=IF(AND(E{i}<>"",E{i}<>"Not Found",NOT(ISBLANK(E{i}))),'
        f'IFERROR(INDEX(\'Update Log\'!$A:$A,MATCH(E{i},\'Update Log\'!$B:$B,0)),D{i}),D{i})'
    )
    ws2.cell(row=i, column=18).value = (
        f'=IF(AND(E{i}<>"",E{i}<>"Not Found",NOT(ISBLANK(E{i}))),'
        f'IF(ISNUMBER(MATCH(E{i},\'Update Log\'!$B:$B,0)),"Update Log","Master Roster"),"No Student ID")'
    )
ws2.freeze_panes = "A2"
set_col_widths(ws2, COL_WIDTHS + [16, 14])
apply_sped_status_cf(ws2, n)
rng18 = f"R2:R{n+1}"
ws2.conditional_formatting.add(rng18,
    FormulaRule(formula=['NOT(ISERROR(SEARCH("Update Log",R2)))'],
                fill=PatternFill("solid", fgColor=CLR_FLAG_BG),
                font=Font(color=CLR_FLAG_FG, bold=True)))

# ============================================================================
# TAB 3 — Copy & Paste (Values Only)
# ============================================================================
ws3 = wb.create_sheet("Copy & Paste")
instr = (f"COPY & PASTE READY  |  Select A3 through P{n+2}, "
         f"copy, then paste as VALUES ONLY into your FY 25-26 master running tab.")
ws3.cell(row=1, column=1, value=instr)
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
ws3.cell(row=1, column=1).fill = PatternFill("solid", fgColor="FCE4D6")
ws3.cell(row=1, column=1).font = Font(bold=True, color="843C0C")
for c_idx, col_name in enumerate(TEMPLATE_COLS, 1):
    ws3.cell(row=2, column=c_idx, value=col_name)
style_header_row(ws3, 2, len(TEMPLATE_COLS), CLR_ORANGE)
write_df_to_sheet(ws3, unified, start_row=3, include_header=False)
ws3.freeze_panes = "A3"
set_col_widths(ws3, COL_WIDTHS)
apply_sped_status_cf(ws3, n)

# ============================================================================
# TAB 4 — Not Found in ST
# ============================================================================
ws4 = wb.create_sheet("Not Found in ST")
nf_cols = TEMPLATE_COLS + ["Action Needed"]
not_found2 = not_found.copy()
not_found2["Action Needed"] = not_found2.apply(lambda r: (
    "Look up NYSSIS in SchoolTool — enter resolved NYSSIS in Update Log tab"
    if pd.isna(r["NYSSIS ID"]) or r["NYSSIS ID"] == "0"
    else "Student not found in ST enrollment — verify with Patrick"
), axis=1)
instr4 = (f"NOT FOUND IN ST  |  {len(not_found2)} students have no NYSSIS or are "
          f"not in ST enrollment. Look up manually and add to Update Log tab.")
ws4.cell(row=1, column=1, value=instr4)
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=17)
ws4.cell(row=1, column=1).fill = PatternFill("solid", fgColor=CLR_NOGOOD_BG)
ws4.cell(row=1, column=1).font = Font(bold=True, color=CLR_NOGOOD_FG)
for c_idx, col_name in enumerate(nf_cols, 1):
    ws4.cell(row=2, column=c_idx, value=col_name)
style_header_row(ws4, 2, len(nf_cols), CLR_PURPLE)
write_df_to_sheet(ws4, not_found2[nf_cols], start_row=3, include_header=False)
row_fill = PatternFill("solid", fgColor="FFE4E1")
for r in range(3, len(not_found2)+3):
    for c in range(1, len(nf_cols)+1):
        ws4.cell(row=r, column=c).fill = row_fill
ws4.freeze_panes = "A3"
set_col_widths(ws4, COL_WIDTHS + [44])

# ============================================================================
# TAB 5 — Update Log
# ============================================================================
ws5 = wb.create_sheet("Update Log")
ul_cols = ["NYSSIS ID","Student ID","Last Name","First Name","DOB",
           "Grade Level","School Name","Charter_Source Ticket #",
           "Resolved_By","Resolved_Date","Notes"]
instr5 = ("UPDATE LOG  |  Paste resolved NYSSIS records here from R or Python output. "
          "Col A = NYSSIS ID | Col B = Student ID (must match master). "
          "Formula Version tab will auto-pull resolved NYSSIS for any matching Student ID.")
ws5.cell(row=1, column=1, value=instr5)
ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
ws5.cell(row=1, column=1).fill = PatternFill("solid", fgColor="E2EFDA")
ws5.cell(row=1, column=1).font = Font(bold=True, color=CLR_GREEN)
for c_idx, col_name in enumerate(ul_cols, 1):
    ws5.cell(row=2, column=c_idx, value=col_name)
style_header_row(ws5, 2, len(ul_cols), CLR_GREEN)

# Seed with students still missing NYSSIS after both ST lookup passes
ul_seed = unified[
    unified["NYSSIS ID"].isna() |
    (unified["NYSSIS ID"].astype(str).str.strip().isin(["0","Not Found","nan",""]))
].copy()
ul_seed["Resolved_By"]   = "Pending"
ul_seed["Resolved_Date"] = pd.NaT

def get_note(row):
    fn = str(row.get("First Name","")) if pd.notna(row.get("First Name")) else ""
    ln = str(row.get("Last Name",""))  if pd.notna(row.get("Last Name"))  else ""
    notes_map = {
        "ANESSA":    "Not found in any ST export — verify enrollment with Patrick.",
        "BOBBY":     "Not found in any ST export — verify enrollment with Patrick.",
        "DONALD":    "No last name in source — search ST by DOB 12/07/2020.",
        "EDWARD":    "No last name in source — search ST by DOB 03/20/2020.",
        "KAIDEN":    "Not found in any ST export — verify enrollment with Patrick.",
        "LILIANA":   "Not found in any ST export — verify enrollment with Patrick.",
        "RAHMIER":   "Not found in any ST export — verify enrollment with Patrick.",
        "XAVIER":    "Not found in any ST export — verify enrollment with Patrick.",
        "JOHNNY":    "Not found in any ST export (Gaston LLL, Johnny) — verify enrollment with Patrick.",
        "ATLANTIS":  "Not found in any ST export — verify enrollment with Patrick.",
        "LENIA":     "Not found in any ST export — verify enrollment with Patrick.",
        "ERIC":      "Not found in any ST export — verify enrollment with Patrick.",
        "SCARLETT":  "Not found in any ST export — verify enrollment with Patrick.",
        "MACKENZIE": "Not found in any ST export (Aldrich, Mackenzie) — verify enrollment with Patrick.",
    }
    return notes_map.get(fn.upper(), "Not found in any ST export — verify enrollment with Patrick.")

ul_seed["Notes"] = ul_seed.apply(get_note, axis=1)
ul_out = ul_seed[["NYSSIS ID","Student ID","Last Name","First Name","DOB",
                   "Grade Level","School Name","Charter_Source Ticket #",
                   "Resolved_By","Resolved_Date","Notes"]]
write_df_to_sheet(ws5, ul_out, start_row=3, include_header=False)
seed_fill = PatternFill("solid", fgColor="FFFACD")
for r in range(3, len(ul_out)+3):
    for c in range(1, len(ul_cols)+1):
        ws5.cell(row=r, column=c).fill = seed_fill
ws5.freeze_panes = "A3"
set_col_widths(ws5, [13,12,14,12,12,10,30,20,14,14,52])

# ============================================================================
# SAVE
# ============================================================================
wb.save(OUTPUT_PATH)
print(f"\n{'='*52}")
print(f"DONE")
print(f"Output: {OUTPUT_PATH}")
print(f"Tab 1 — All Charters Unified : {n:,} students")
print(f"Tab 2 — Formula Version       : {n:,} rows + NYSSIS lookup formulas")
print(f"Tab 3 — Copy & Paste          : {n:,} rows (values only)")
print(f"Tab 4 — Not Found in ST       : {len(not_found2)} students flagged")
print(f"Tab 5 — Update Log            : {len(ul_out)} pending resolution")
