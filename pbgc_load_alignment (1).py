#!/usr/bin/env python3
"""
PBGC Load File Alignment Analysis
=================================
Compares the MANUAL ("PBGC") load file against the AUTOMATION load file for one
or more plans and writes a 9-tab workbook laid out like PBGC_Load_Alignment_MOCKUP.xlsx:

    README | Summary_by_Plan | Population_Diff | Element_Diff | Heatmap_9A |
    Heatmap_Payee | Value_Diff_9A | Value_Diff_Payee | Participant_Summary

How to use
----------
1. Edit the CONFIGURATION block below (file paths, plan name, output path).
   There are no command-line arguments by design.
2. Run:  python pbgc_load_alignment.py
3. Open the output workbook. Formulas (match %, totals, heatmap averages,
   Numeric_Diff) are live; everything else is written as data.

Terminology
-----------
"PBGC"        = the manual file.
"Automation"  = the automation file.
Participant_ID = the manual file's "Unique ID" column. The automation file has no
                 such column, so it is DERIVED (see build_uid) as:
                     LAST NAME + FIRST NAME + last 4 digits of PRT_ID + first letter of CUST_TYPE
                 e.g. SMITHJANE9999P. Verify this against your real data - the script
                 prints how many derived IDs were found in the manual file.
"""

import difflib
import numbers
import os
import re
import shutil
import tempfile
import warnings
import zipfile
from xml.sax.saxutils import escape as xml_escape
from collections import Counter, OrderedDict, defaultdict
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

# =============================================================================
# CONFIGURATION  (edit this block)
# =============================================================================

# One entry per plan. Leave a load as None if that file pair is not available
# (the matching 9A / Payee tabs are then created empty).
PLANS = [
    {
        "plan_name": "Dalton",
        "payee": {
            "manual": "Manual_Sample.xlsx",
            "automation": "Automation_Sample.xlsx",
        },
        "9a": None,  # e.g. {"manual": "Dalton_9A_Manual.xlsx", "automation": "Dalton_9A_Auto.xlsx"}
    },
]

OUTPUT_PATH = "PBGC_Load_Alignment_Dalton.xlsx"
SHEET_NAME = 0  # sheet to read in every input file (0 = first sheet, or a sheet name)

# ---- Key / context columns --------------------------------------------------
MANUAL_KEY_COLUMN = "Unique ID"          # manual column holding the Participant_ID
# Automation columns used to DERIVE the Unique ID:
AUTO_LAST_NAME, AUTO_FIRST_NAME = "NAME_LAST", "NAME_FIRST"
AUTO_PRT_ID, AUTO_ROLE = "PRT_ID", "CUST_TYPE"
UID_STRIP_NON_ALNUM = True               # drop spaces / punctuation when building & comparing IDs
UID_PRT_DIGITS = 4                       # number of trailing PRT_ID digits in the ID
UID_ROLE_CHARS = 1                       # number of leading CUST_TYPE characters in the ID
PRT_ID_WIDTH = 9                         # numeric PRT_IDs shorter than this are left-padded with zeros
                                         # (Excel drops leading zeros); set to 0 to disable

# (manual column, automation column) used for the context columns on Population_Diff
MANUAL_PRT_ID, MANUAL_ROLE = "PRT_ID", "CUST_TYPE"
CONTEXT_STATUS = ("Eligibility Status", "ELIGIBILITY_STATUS")
CONTEXT_DOB = ("Date of Birth", "DOB")
CONTEXT_CASE = ("CASE_ID_NUMBER", "CASE_ID_NMBER")

# ---- Column mapping: AUTOMATION header -> MANUAL header ---------------------
# Columns with identical names (ignoring case / surrounding spaces) are paired
# automatically, so only the differently-named ones are listed here.
# Anything not paired is reported as "Missing from Automation" / "Extra in Automation".
COLUMN_MAP = {
    "NAME_LAST": "Last Name",
    "NAME_FIRST": "First Name",
    "NAME_MIDDLE": "Middle Name",
    "SEX": "Sex",
    "DOB": "Date of Birth",
    "DOD": "Date of Death",
    "STREET_1": "Address Line 1",
    "STREET_2": "Address Line 2",
    "CITY": "City",
    "STATE_ABBRV": "State",
    "POSTAL_CODE": "Postal Code",
    "COUNTRY_CODE": "Country",
    "CASE_ID_NMBER": "CASE_ID_NUMBER",
    "TAX_ID_TYPE": "Tax ID Type",
    "ELIGIBILITY_STATUS": "Eligibility Status",
    "FIRST_PAYMENT_DATE": "FIRST_PAYMENT",
    "LAST_PAYMENT_DATE": "LAST_PAYMENT",
    "LINKING_PRT_ID": "Linking",
    # PAYMENT_DATE <-> FIRST_PAYMENT intentionally NOT mapped (treated as unrelated for now).
}
# Optional per-load-type additions/overrides, e.g. {"9A": {"SOME_AUTO_COL": "Some Manual Col"}}
COLUMN_MAP_OVERRIDES = {"9A": {}, "Payee": {}}

# ---- Comparison rules -------------------------------------------------------
# Element names below use the AUTOMATION header where a pair exists.
DATE_ELEMENTS = {
    "DOB", "DOD", "ACTUAL_RETIREMENT_DATE", "BENEFIT_START_DATE", "BENEFIT_END_DATE",
    "FIRST_PAYMENT_DATE", "LAST_PAYMENT_DATE", "LS_EST_DATE", "PAYMENT_DATE",
}
NUMERIC_ELEMENTS = {
    "PAY_AMOUNT", "PE_WITHHELD_DOLLAR", "PP_POPUP_AMT", "JS_ELECTION_PERCENT",
    "DRO_PERCENT", "LS_EST_AMT", "CERTAIN_PERIOD",
}
NUMERIC_TOLERANCE = 0.01   # |automation - pbgc| <= this is a match
ROUNDING_MAX = 0.10        # mismatches up to this size are tagged "Rounding"
# Elements where a numeric 0 is treated the same as blank (automation writes 0, manual leaves blank).
ZERO_AS_BLANK_ELEMENTS = {"PE_WITHHELD_DOLLAR", "PP_POPUP_AMT"}
# Should "Format Only" differences count as differences in Cell_Match_% / Participant_Summary?
FORMAT_ONLY_COUNTS_AS_DIFF = False

# Value equivalences applied to BOTH sides before text comparison.
# Keys are upper-case element names; values are {upper-case value: canonical value}.
# Values not listed pass through unchanged. Add to these as you learn the code tables.
VALUE_MAPS = {
    "SEX": {"M": "M", "MALE": "M", "F": "F", "FEMALE": "F", "U": "U", "UNKNOWN": "U"},
    "COUNTRY_CODE": {
        "USA": "UNITED STATES OF AMERICA", "US": "UNITED STATES OF AMERICA",
        "UNITED STATES": "UNITED STATES OF AMERICA",
        "CAN": "CANADA", "CA": "CANADA",
    },
    "ACH_ACCOUNT_TYPE": {"CHECKING": "CH", "SAVINGS": "SV"},
    # TODO - fill in once the code tables are known (left empty => values compared as-is):
    "APP_ELECTION": {},   # e.g. {"NO WITHHOLDING": "007", ...}
    "CUST_TYPE": {},      # e.g. {"BE": "SP"} if the two files use different role codes
}
# Elements compared on the code before " - " only (e.g. "E - ELIGIBLE" -> "E").
CODE_BEFORE_DASH_ELEMENTS = {"ELIGIBILITY_STATUS"}
# Elements where full state / province names are converted to their abbreviation.
STATE_ELEMENTS = {"STATE_ABBRV"}

# =============================================================================
# END OF CONFIGURATION
# =============================================================================

US_STATES_CA_PROVINCES = (
    "ALABAMA AL|ALASKA AK|ARIZONA AZ|ARKANSAS AR|CALIFORNIA CA|COLORADO CO|CONNECTICUT CT|"
    "DELAWARE DE|DISTRICT OF COLUMBIA DC|FLORIDA FL|GEORGIA GA|HAWAII HI|IDAHO ID|ILLINOIS IL|"
    "INDIANA IN|IOWA IA|KANSAS KS|KENTUCKY KY|LOUISIANA LA|MAINE ME|MARYLAND MD|MASSACHUSETTS MA|"
    "MICHIGAN MI|MINNESOTA MN|MISSISSIPPI MS|MISSOURI MO|MONTANA MT|NEBRASKA NE|NEVADA NV|"
    "NEW HAMPSHIRE NH|NEW JERSEY NJ|NEW MEXICO NM|NEW YORK NY|NORTH CAROLINA NC|NORTH DAKOTA ND|"
    "OHIO OH|OKLAHOMA OK|OREGON OR|PENNSYLVANIA PA|RHODE ISLAND RI|SOUTH CAROLINA SC|"
    "SOUTH DAKOTA SD|TENNESSEE TN|TEXAS TX|UTAH UT|VERMONT VT|VIRGINIA VA|WASHINGTON WA|"
    "WEST VIRGINIA WV|WISCONSIN WI|WYOMING WY|PUERTO RICO PR|"
    "ALBERTA AB|BRITISH COLUMBIA BC|MANITOBA MB|NEW BRUNSWICK NB|NEWFOUNDLAND AND LABRADOR NL|"
    "NOVA SCOTIA NS|ONTARIO ON|PRINCE EDWARD ISLAND PE|QUEBEC QC|SASKATCHEWAN SK|"
    "NORTHWEST TERRITORIES NT|NUNAVUT NU|YUKON YT"
)
STATE_NAME_TO_ABBR = {p.rsplit(" ", 1)[0]: p.rsplit(" ", 1)[1] for p in US_STATES_CA_PROVINCES.split("|")}

STATUS_ORDER = ["Only in PBGC", "Only in Automation", "Duplicate in PBGC",
                "Duplicate in Automation", "Role Mismatch"]
LOAD_TYPES = ("9A", "Payee")

warnings.filterwarnings("ignore")


# =============================================================================
# Value helpers
# =============================================================================
def is_blank(v):
    if v is None or v is pd.NaT:
        return True
    if isinstance(v, float) and v != v:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def canon_text(v):
    """Upper-case, whitespace-collapsed text; whole floats lose the '.0'."""
    if isinstance(v, (datetime, date)):
        v = v.strftime("%Y-%m-%d")
    elif isinstance(v, bool):
        v = str(v)
    elif isinstance(v, numbers.Integral):
        v = str(int(v))
    elif isinstance(v, numbers.Real):
        v = str(int(v)) if float(v) == int(v) else format(float(v), ".10g")
    return re.sub(r"\s+", " ", str(v)).strip().upper()


def to_number(v):
    if is_blank(v) or isinstance(v, (bool, datetime, date)):
        return None
    if isinstance(v, numbers.Real):
        return float(v)
    s = re.sub(r"[,$%\s]", "", str(v))
    try:
        return float(s)
    except ValueError:
        return None


def to_date(v):
    if is_blank(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        d = pd.to_datetime(v.strip(), errors="coerce")
        return None if pd.isna(d) else d.date()
    return None


def date_text(v):
    """Text form of a date cell, used only to decide 'Format Only'."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def out_val(v):
    """Value as written to the workbook (dates as ISO text, whole floats as ints)."""
    if is_blank(v):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, numbers.Integral):
        return int(v)
    if isinstance(v, numbers.Real):
        return int(v) if float(v) == int(v) else float(v)
    return str(v)


def infer_type(values):
    kinds = set()
    for v in values:
        if is_blank(v):
            continue
        if isinstance(v, (datetime, date)):
            kinds.add("Date")
        elif isinstance(v, numbers.Real) and not isinstance(v, bool):
            kinds.add("Number")
        else:
            kinds.add("Text")
    if not kinds:
        return "Empty"
    return kinds.pop() if len(kinds) == 1 else "Mixed"


def norm_header(h):
    return re.sub(r"\s+", " ", str(h)).strip().upper()


def clean_id(v):
    return "" if is_blank(v) else canon_text(v)


# =============================================================================
# Unique ID
# =============================================================================
def _uid_clean(v):
    s = "" if is_blank(v) else canon_text(v)
    return re.sub(r"[^A-Z0-9]", "", s) if UID_STRIP_NON_ALNUM else s


def build_uid(last, first, prt_id, role):
    """Derive the manual-style Unique ID for an automation row."""
    digits = re.sub(r"\D", "", clean_id(prt_id))
    tail = digits.zfill(max(PRT_ID_WIDTH, UID_PRT_DIGITS))[-UID_PRT_DIGITS:] if digits else ""
    r = "" if is_blank(role) else canon_text(role)[:UID_ROLE_CHARS]
    return _uid_clean(last) + _uid_clean(first) + tail + r


# =============================================================================
# Comparison engine
# =============================================================================
def norm_value(element, v):
    s = canon_text(v)
    if element in CODE_BEFORE_DASH_ELEMENTS:
        s = re.split(r"\s+-\s+", s, maxsplit=1)[0]
    if element in STATE_ELEMENTS:
        s = STATE_NAME_TO_ABBR.get(s, s)
    return VALUE_MAPS.get(element, {}).get(s, s)


def date_note(da, db):
    months = (db.year - da.year) * 12 + db.month - da.month
    if da.day == db.day and abs(months) == 1:
        return "Off by one month"
    days = abs((db - da).days)
    return "Off by one day" if days == 1 else f"Dates differ by {days} days"


def compare_values(el, a, b):
    """a = PBGC (manual) value, b = automation value.
    Returns (status, note); status in Match | Both Blank | Format Only | Value Mismatch |
    Missing in Automation | Missing in PBGC."""
    name = el["key"]
    if name in ZERO_AS_BLANK_ELEMENTS:
        if to_number(a) == 0:
            a = None
        if to_number(b) == 0:
            b = None
    ba, bb = is_blank(a), is_blank(b)
    if ba and bb:
        return "Both Blank", ""
    if ba:
        return "Missing in PBGC", ""
    if bb:
        return "Missing in Automation", ""

    if el["kind"] == "date":
        da, db = to_date(a), to_date(b)
        if da and db:
            if da == db:
                if date_text(a) == date_text(b):
                    return "Match", ""
                return "Format Only", "Date format"
            return "Value Mismatch", date_note(da, db)
    elif el["kind"] == "numeric":
        na, nb = to_number(a), to_number(b)
        if na is not None and nb is not None:
            d = nb - na
            if abs(d) <= NUMERIC_TOLERANCE:
                return "Match", ""
            return "Value Mismatch", ("Rounding" if abs(d) <= ROUNDING_MAX else "")

    ta, tb = norm_value(name, a), norm_value(name, b)
    if ta == tb:
        return "Match", ""
    if ta.isdigit() and tb.isdigit() and ta.lstrip("0") == tb.lstrip("0"):
        who = "PBGC" if len(ta) < len(tb) else "Automation"   # the shorter value lost its zeros
        return "Format Only", f"Leading zeros dropped ({who})"
    alnum_a, alnum_b = re.sub(r"[^A-Z0-9]", "", ta), re.sub(r"[^A-Z0-9]", "", tb)
    if alnum_a and alnum_a == alnum_b:
        return "Format Only", "Punctuation/spacing differs"
    na, nb = to_number(a), to_number(b)
    if na is not None and nb is not None and abs(na - nb) < 1e-9:
        return "Match", ""
    return "Value Mismatch", ""


def counts_as_diff(status):
    if status in ("Value Mismatch", "Missing in Automation", "Missing in PBGC"):
        return True
    return status == "Format Only" and FORMAT_ONLY_COUNTS_AS_DIFF


# =============================================================================
# Loading & column pairing
# =============================================================================
def read_table(path):
    df = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=object,
                       keep_default_na=False, na_values=[""])
    df.columns = [str(c).strip() for c in df.columns]
    return df


def require(df, col, label):
    if col not in df.columns:
        raise KeyError(f"{label}: expected column '{col}' not found. Columns present: {list(df.columns)}")


def pair_columns(man_cols, aut_cols, load_type):
    cmap = {norm_header(k): norm_header(v) for k, v in COLUMN_MAP.items()}
    cmap.update({norm_header(k): norm_header(v) for k, v in COLUMN_MAP_OVERRIDES.get(load_type, {}).items()})
    man_by_norm = {norm_header(c): c for c in man_cols}
    key_norm = norm_header(MANUAL_KEY_COLUMN)

    used, both, auto_only = set(), [], []
    for ac in aut_cols:
        target = cmap.get(norm_header(ac), norm_header(ac))
        mc = man_by_norm.get(target)
        if mc is not None and norm_header(mc) != key_norm and mc not in used:
            used.add(mc)
            both.append({"name": ac, "man_col": mc, "aut_col": ac})
        else:
            auto_only.append({"name": ac, "man_col": None, "aut_col": ac})
    man_only = [{"name": mc, "man_col": mc, "aut_col": None}
                for mc in man_cols if mc not in used]
    return both, man_only, auto_only


# =============================================================================
# Per-load analysis
# =============================================================================
class LoadResult:
    pass


def analyze_load(plan_name, load_type, manual_path, auto_path):
    print(f"\n[{plan_name} / {load_type}] reading files ...")
    man, aut = read_table(manual_path), read_table(auto_path)

    for col in (MANUAL_KEY_COLUMN, MANUAL_PRT_ID, MANUAL_ROLE):
        require(man, col, f"Manual file ({manual_path})")
    for col in (AUTO_LAST_NAME, AUTO_FIRST_NAME, AUTO_PRT_ID, AUTO_ROLE):
        require(aut, col, f"Automation file ({auto_path})")

    man_rows, aut_rows = man.to_dict("records"), aut.to_dict("records")
    both, man_only, aut_only = pair_columns(list(man.columns), list(aut.columns), load_type)

    # ---- element definitions -------------------------------------------------
    for el in both + man_only + aut_only:
        # "key" = upper-case element name used for every rule lookup in the CONFIGURATION block
        # (for paired columns the name is the automation header).
        el["key"] = el["name"].upper()
        has_real_dates = any(
            isinstance(r[col], (datetime, date))
            for col, rows in ((el["man_col"], man_rows), (el["aut_col"], aut_rows)) if col
            for r in rows)
        if el["key"] in DATE_ELEMENTS or has_real_dates:
            el["kind"] = "date"
        elif el["key"] in NUMERIC_ELEMENTS:
            el["kind"] = "numeric"
        else:
            el["kind"] = "text"
    n_both = len(both)

    # ---- keys -----------------------------------------------------------------
    man_keys = [_uid_clean(r[MANUAL_KEY_COLUMN]) for r in man_rows]
    aut_keys = [build_uid(r[AUTO_LAST_NAME], r[AUTO_FIRST_NAME], r[AUTO_PRT_ID], r[AUTO_ROLE])
                for r in aut_rows]
    mcount, acount = Counter(man_keys), Counter(aut_keys)
    matched = {k for k in mcount if k and mcount[k] == 1 and acount.get(k) == 1}

    def ctx(row, side, which):
        col = {"status": CONTEXT_STATUS, "dob": CONTEXT_DOB, "case": CONTEXT_CASE}[which][0 if side == "man" else 1]
        return row.get(col) if col in row else None

    def role_of(row, side):
        return clean_id(row.get(MANUAL_ROLE if side == "man" else AUTO_ROLE))

    def prt_of(row, side):
        p = clean_id(row.get(MANUAL_PRT_ID if side == "man" else AUTO_PRT_ID))
        return p.zfill(PRT_ID_WIDTH) if p.isdigit() and PRT_ID_WIDTH else p

    def case_of(row, side):
        return clean_id(ctx(row, side, "case"))

    cases = Counter(case_of(r, "man") for r in man_rows if case_of(r, "man"))
    if not cases:
        cases = Counter(case_of(r, "aut") for r in aut_rows if case_of(r, "aut"))
    case_number = cases.most_common(1)[0][0] if cases else ""
    if len(cases) > 1:
        print(f"  WARNING: {len(cases)} distinct case numbers in the files: {dict(cases)}")

    # ---- population differences ---------------------------------------------
    pop = []

    def pop_row(row, side, pid, status, notes=""):
        dob = ctx(row, side, "dob")
        pop.append({
            "case": case_of(row, side) or case_number, "load": load_type, "pid": pid,
            "prt": prt_of(row, side), "role": role_of(row, side), "status": status,
            "pstatus": out_val(ctx(row, side, "status")), "dob": out_val(dob), "notes": notes,
        })

    # duplicates (one line per duplicated key)
    for side, keys, rows, cnt, other, label in (
            ("man", man_keys, man_rows, mcount, acount, "Duplicate in PBGC"),
            ("aut", aut_keys, aut_rows, acount, mcount, "Duplicate in Automation")):
        seen = set()
        for k, row in zip(keys, rows):
            if cnt[k] > 1 and k not in seen:
                seen.add(k)
                orig = row[MANUAL_KEY_COLUMN] if side == "man" else k
                note = f"{cnt[k]} rows with identical key; excluded from value comparison"
                pop_row(row, side, clean_id(orig) or k, label, note)

    # only-in rows (key absent from the other file) and role mismatches
    man_only_rows = [(k, r) for k, r in zip(man_keys, man_rows) if mcount[k] == 1 and k not in acount]
    aut_only_rows = [(k, r) for k, r in zip(aut_keys, aut_rows) if acount[k] == 1 and k not in mcount]
    man_roles_by_prt, aut_roles_by_prt = defaultdict(set), defaultdict(set)
    for k, r in man_only_rows:
        man_roles_by_prt[prt_of(r, "man")].add(role_of(r, "man"))
    for k, r in aut_only_rows:
        aut_roles_by_prt[prt_of(r, "aut")].add(role_of(r, "aut"))

    for side, rows_, other_roles, label_other, only_label in (
            ("man", man_only_rows, aut_roles_by_prt, "Automation", "Only in PBGC"),
            ("aut", aut_only_rows, man_roles_by_prt, "PBGC", "Only in Automation")):
        for k, r in rows_:
            prt, role = prt_of(r, side), role_of(r, side)
            other = {x for x in other_roles.get(prt, set()) if x != role} if prt else set()
            pid = clean_id(r[MANUAL_KEY_COLUMN]) if side == "man" else k
            pid = pid or k or "(blank key)"
            if other:
                pop_row(r, side, pid, "Role Mismatch",
                        f"Same PRT_ID found with role {', '.join(sorted(other))} in {label_other} file")
            else:
                pop_row(r, side, pid, only_label, "" if k else "Key could not be built (name/PRT_ID/role blank)")
    pop.sort(key=lambda d: (STATUS_ORDER.index(d["status"]), d["pid"]))

    # ---- value comparison -----------------------------------------------------
    man_by_key = {k: r for k, r in zip(man_keys, man_rows) if k in matched}
    aut_by_key = {k: r for k, r in zip(aut_keys, aut_rows) if k in matched}
    stats = {el["name"]: Counter() for el in both}
    value_rows, part_rows = [], []
    full_match = 0
    diff_cells = 0

    for k in sorted(matched):
        mr, ar = man_by_key[k], aut_by_key[k]
        pid = clean_id(mr[MANUAL_KEY_COLUMN]) or k
        case = case_of(mr, "man") or case_number
        n_diff = n_miss_aut = n_miss_pbgc = 0
        diff_names = []
        for el in both:
            a, b = mr[el["man_col"]], ar[el["aut_col"]]
            status, note = compare_values(el, a, b)
            stats[el["name"]]["compared"] += 1
            if counts_as_diff(status):
                n_diff += 1
                diff_names.append(el["name"])
                if status == "Missing in Automation":
                    n_miss_aut += 1
                elif status == "Missing in PBGC":
                    n_miss_pbgc += 1
            else:
                stats[el["name"]]["matched"] += 1
            if status not in ("Match", "Both Blank"):
                value_rows.append({
                    "case": case, "pid": pid, "element": el["name"],
                    "pbgc": out_val(a), "auto": out_val(b), "diff_type": status, "notes": note,
                })
        diff_cells += n_diff
        if n_diff == 0:
            full_match += 1
        else:
            part_rows.append({
                "case": case, "load": load_type, "pid": pid, "compared": n_both,
                "different": n_diff, "miss_aut": n_miss_aut, "miss_pbgc": n_miss_pbgc,
                "diff_elements": ", ".join(diff_names),
            })
    # worst offenders first
    part_rows.sort(key=lambda d: (-d["different"], d["pid"]))
    elem_order = {el["name"]: i for i, el in enumerate(both + man_only + aut_only)}
    value_rows.sort(key=lambda d: (d["pid"], elem_order.get(d["element"], 0)))

    # ---- element-level table ------------------------------------------------------
    # Column position: compare the order of the shared columns in each file and flag only the
    # ones that sit outside the longest run that is in the same relative order in both files.
    man_pos = {el["name"]: list(man.columns).index(el["man_col"]) + 1 for el in both}
    aut_pos = {el["name"]: list(aut.columns).index(el["aut_col"]) + 1 for el in both}
    man_seq = sorted((el["name"] for el in both), key=man_pos.get)
    aut_seq = sorted((el["name"] for el in both), key=aut_pos.get)
    in_order = set()
    for blk in difflib.SequenceMatcher(None, man_seq, aut_seq, autojunk=False).get_matching_blocks():
        in_order.update(man_seq[blk.a:blk.a + blk.size])

    def fill(rows, col):
        return sum(0 if is_blank(r[col]) else 1 for r in rows) / len(rows) if rows else None

    def vals(rows, col):
        return [r[col] for r in rows]

    man_only_norm = {el["name"]: norm_header(el["name"]) for el in man_only}
    aut_only_norm = {el["name"]: norm_header(el["name"]) for el in aut_only}

    def fuzzy(name, pool):
        best = difflib.get_close_matches(norm_header(name), list(pool.values()), n=1, cutoff=0.75)
        if not best:
            return None
        orig = [k for k, v in pool.items() if v == best[0]][0]
        return orig, difflib.SequenceMatcher(None, norm_header(name), best[0]).ratio()

    elements = []
    for el in both:
        notes = []
        if norm_header(el["man_col"]) != norm_header(el["aut_col"]):
            notes.append(f"PBGC column: '{el['man_col']}'")
        if el["name"] not in in_order:
            notes.append(f"Out of sequence vs. other shared columns (PBGC col {man_pos[el['name']]}, "
                         f"Automation col {aut_pos[el['name']]})")
        if el["key"] in ZERO_AS_BLANK_ELEMENTS:
            notes.append("0 treated as blank")
        if (VALUE_MAPS.get(el["key"]) or el["key"] in CODE_BEFORE_DASH_ELEMENTS
                or el["key"] in STATE_ELEMENTS):
            notes.append("Code/description equivalence applied")
        st = stats[el["name"]]
        elements.append({
            "name": el["name"], "in_p": True, "in_a": True, "status": "In Both",
            "dt_p": infer_type(vals(man_rows, el["man_col"])), "dt_a": infer_type(vals(aut_rows, el["aut_col"])),
            "fill_p": fill(man_rows, el["man_col"]), "fill_a": fill(aut_rows, el["aut_col"]),
            "compared": st["compared"], "matched": st["matched"], "notes": "; ".join(notes),
        })
    for el in man_only:
        notes = ["Not produced by automation"]
        if norm_header(el["man_col"]) == norm_header(MANUAL_KEY_COLUMN):
            notes = [f"Match key. Not in automation file; derived for matching as LAST+FIRST+last {UID_PRT_DIGITS} "
                     f"of PRT_ID+first {UID_ROLE_CHARS} of CUST_TYPE"]
        else:
            f = fuzzy(el["name"], aut_only_norm)
            if f:
                notes.append(f"Possible rename of {f[0]}? (fuzzy match {f[1]:.2f})")
        elements.append({
            "name": el["name"], "in_p": True, "in_a": False, "status": "Missing from Automation",
            "dt_p": infer_type(vals(man_rows, el["man_col"])), "dt_a": "",
            "fill_p": fill(man_rows, el["man_col"]), "fill_a": None,
            "compared": 0, "matched": 0, "notes": "; ".join(notes),
        })
    for el in aut_only:
        notes = ["Automation adds column not in PBGC layout"]
        f = fuzzy(el["name"], man_only_norm)
        if f:
            notes.append(f"Possible rename of {f[0]}? (fuzzy match {f[1]:.2f})")
        elements.append({
            "name": el["name"], "in_p": False, "in_a": True, "status": "Extra in Automation",
            "dt_p": "", "dt_a": infer_type(vals(aut_rows, el["aut_col"])),
            "fill_p": None, "fill_a": fill(aut_rows, el["aut_col"]),
            "compared": 0, "matched": 0, "notes": "; ".join(notes),
        })

    # ---- result ------------------------------------------------------------------
    r = LoadResult()
    r.plan, r.load, r.case = plan_name, load_type, case_number
    r.rows_pbgc, r.rows_auto, r.rows_matched = len(man_rows), len(aut_rows), len(matched)
    r.elements, r.pop, r.value_rows, r.part_rows = elements, pop, value_rows, part_rows
    r.n_in_both = n_both
    r.n_missing = len(man_only)
    r.n_extra = len(aut_only)
    r.full_match = full_match
    total_cells = len(matched) * n_both
    r.cell_match = (total_cells - diff_cells) / total_cells if total_cells else None

    pct_found = sum(1 for k in aut_keys if k in mcount) / len(aut_keys) if aut_keys else 0
    print(f"  rows: PBGC={r.rows_pbgc}, Automation={r.rows_auto}, matched on Participant_ID={r.rows_matched}")
    print(f"  derived automation IDs found in manual file: {pct_found:.1%}")
    if r.rows_matched == 0 or pct_found < 0.5:
        print("  WARNING: few Participant_IDs matched. Check the Unique ID derivation "
              "(UID_* settings / build_uid) against your real data.")
    print(f"  population diffs={len(pop)}, value-diff rows={len(value_rows)}, "
          f"participants with differences={len(part_rows)}")
    return r


# =============================================================================
# Workbook writer
# =============================================================================
NAVY, LIGHT = "1F3864", "D9E1F2"
ARIAL = "Arial"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED_FONT = Font(name=ARIAL, color="9C0006")


# Result of every formula cell, computed in Python while writing: {(sheet title, "E2"): value}.
# openpyxl saves formulas WITHOUT a result, so anything that reads the file before Excel recalculates
# (pandas, formula libraries, previewers) would see blanks. inject_cached_values() stores these results.
CACHE = {}


def put(ws, row, col, value, nf=None, bold=False, italic=False, color=None, border=False,
        align=None, wrap=None, fill=None, size=10, formula=False, cached=None):
    c = ws.cell(row=row, column=col)
    if formula and cached is not None:
        CACHE[(ws.title, c.coordinate)] = cached
    if isinstance(value, str) and not formula:
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        c.value = value
        if value.startswith("="):
            c.data_type = "s"
    else:
        c.value = value
    c.font = Font(name=ARIAL, size=size, bold=bold, italic=italic, color=color)
    if nf:
        c.number_format = nf
    if border:
        c.border = BOX
    if align or wrap is not None:
        c.alignment = Alignment(horizontal=align, wrap_text=wrap, vertical="center" if wrap else None)
    if fill:
        c.fill = PatternFill("solid", start_color=fill)
    return c


def title(ws, text):
    put(ws, 1, 1, text, bold=True, color=NAVY, size=16)
    ws.row_dimensions[1].height = 20


def header(ws, row, labels, border=False, height=31.5):
    for i, lab in enumerate(labels, 1):
        put(ws, row, i, lab, bold=True, color="FFFFFF", fill=NAVY, align="center", wrap=True, border=border)
    ws.row_dimensions[row].height = height


def print_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def setup(ws, tab, freeze, widths):
    print_setup(ws)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = freeze
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w


def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=None, showFirstColumn=False, showLastColumn=False,
                                      showRowStripes=True, showColumnStripes=False)
    ws.add_table(t)


def nf_pct():
    return "0.0%"


# Cell_Match_% is written as 0 when a plan has no matched participants (so the weighted
# ALL PLANS formula never touches a blank/text cell); the format displays that zero as "n/a".
NF_PCT_NA = '0.0%;-0.0%;"n/a"'


def inject_cached_values(path, sheet_names):
    """Store the Python-computed result next to each formula so the file reads correctly everywhere.
    Excel still recalculates everything on open (fullCalcOnLoad)."""
    by_sheet = defaultdict(dict)
    for (title, coord), v in CACHE.items():
        by_sheet[sheet_names.index(title) + 1][coord] = v
    pat = re.compile(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f><v></v></c>', re.S)

    def make_sub(values):
        def sub(m):
            coord, attrs, f = m.group(1), m.group(2), m.group(3)
            if coord not in values:
                return m.group(0)
            v = values[coord]
            if isinstance(v, bool):
                return m.group(0)
            if isinstance(v, numbers.Real):
                return f'<c r="{coord}"{attrs}><f>{f}</f><v>{repr(float(v)) if isinstance(v, float) else int(v)}</v></c>'
            return f'<c r="{coord}"{attrs} t="str"><f>{f}</f><v>{xml_escape(str(v))}</v></c>'
        return sub

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if m and int(m.group(1)) in by_sheet:
                data = pat.sub(make_sub(by_sheet[int(m.group(1))]), data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp.name, path)


def write_workbook(results):
    CACHE.clear()
    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    ws_sum = wb.create_sheet("Summary_by_Plan")
    ws_pop = wb.create_sheet("Population_Diff")
    ws_el = wb.create_sheet("Element_Diff")
    ws_h9 = wb.create_sheet("Heatmap_9A")
    ws_hp = wb.create_sheet("Heatmap_Payee")
    ws_v9 = wb.create_sheet("Value_Diff_9A")
    ws_vp = wb.create_sheet("Value_Diff_Payee")
    ws_ps = wb.create_sheet("Participant_Summary")

    write_summary(ws_sum, results)
    write_population(ws_pop, results)
    write_elements(ws_el, results)
    write_heatmap(ws_h9, results, "9A")
    write_heatmap(ws_hp, results, "Payee")
    write_value_diff(ws_v9, results, "9A", "tblValueDiff9A")
    write_value_diff(ws_vp, results, "Payee", "tblValueDiffPayee")
    write_participants(ws_ps, results)
    write_readme(readme, results)

    wb.save(OUTPUT_PATH)
    inject_cached_values(OUTPUT_PATH, wb.sheetnames)
    print(f"\nSaved: {OUTPUT_PATH}")


# ---- Summary_by_Plan --------------------------------------------------------
def write_summary(ws, results):
    title(ws, "Summary by Plan")
    heads = ["Case_Number", "Load_Type", "Rows_PBGC", "Rows_Automation", "Rows_Matched", "Rows_Only_PBGC",
             "Rows_Only_Automation", "Population_Match_%", "Elements_Compared", "Elements_Missing_in_Automation",
             "Elements_Extra_in_Automation", "Participants_Fully_Matching", "Participants_With_Any_Diff",
             "Cell_Match_%", "Plan_Name"]
    header(ws, 3, heads, border=True, height=45)
    first, last = 4, 3 + len(results)

    # order: by plan as listed, 9A before Payee
    plan_order = {p["plan_name"]: i for i, p in enumerate(PLANS)}
    rows = sorted(results, key=lambda r: (plan_order[r.plan], LOAD_TYPES.index(r.load)))
    last = 3 + len(rows)
    sums = Counter()
    n_weighted = 0.0
    for i, r in enumerate(rows):
        n = first + i
        e_, f_, g_ = r.rows_matched, r.rows_pbgc - r.rows_matched, r.rows_auto - r.rows_matched
        h_ = 0 if e_ + f_ + g_ == 0 else e_ / (e_ + f_ + g_)
        m_ = e_ - r.full_match
        n_ = r.cell_match if r.cell_match is not None else 0
        cache = {6: f_, 7: g_, 8: h_, 13: m_}
        for k, v in ((3, r.rows_pbgc), (4, r.rows_auto), (5, e_), (6, f_), (7, g_), (10, r.n_missing),
                     (11, r.n_extra), (12, r.full_match), (13, m_)):
            sums[k] += v
        sums["maxI"] = max(sums["maxI"], r.n_in_both)
        n_weighted += n_ * e_
        vals = [
            (1, r.case, None), (2, r.load, None), (3, r.rows_pbgc, "#,##0"), (4, r.rows_auto, "#,##0"),
            (5, r.rows_matched, "#,##0"),
            (6, f"=C{n}-E{n}", "#,##0"), (7, f"=D{n}-E{n}", "#,##0"),
            (8, f"=IF(E{n}+F{n}+G{n}=0,0,E{n}/(E{n}+F{n}+G{n}))", nf_pct()),
            (9, r.n_in_both, None), (10, r.n_missing, None), (11, r.n_extra, None),
            (12, r.full_match, "#,##0"), (13, f"=E{n}-L{n}", "#,##0"),
            (14, n_, NF_PCT_NA), (15, r.plan, None),
        ]
        for col, v, nf in vals:
            put(ws, n, col, v, nf=nf, border=True, formula=isinstance(v, str) and v.startswith("="),
                cached=cache.get(col))

    put(ws, 2, 1, "ALL PLANS", bold=True, fill=LIGHT, border=True)
    put(ws, 2, 2, None, bold=True, fill=LIGHT, border=True)
    rng = lambda c: f"{c}{first}:{c}{last}"
    tot = {
        3: (f"=SUM({rng('C')})", "#,##0"), 4: (f"=SUM({rng('D')})", "#,##0"),
        5: (f"=SUM({rng('E')})", "#,##0"), 6: (f"=SUM({rng('F')})", "#,##0"),
        7: (f"=SUM({rng('G')})", "#,##0"), 8: ("=IF(E2+F2+G2=0,0,E2/(E2+F2+G2))", nf_pct()),
        9: (f"=MAX({rng('I')})", None), 10: (f"=SUM({rng('J')})", None), 11: (f"=SUM({rng('K')})", None),
        12: (f"=SUM({rng('L')})", "#,##0"), 13: (f"=SUM({rng('M')})", "#,##0"),
        14: (f"=IF(SUM({rng('E')})=0,0,SUMPRODUCT({rng('N')},{rng('E')})/SUM({rng('E')}))", NF_PCT_NA),
        15: (None, None),
    }
    e2 = sums[5]
    tot_cache = {
        3: sums[3], 4: sums[4], 5: sums[5], 6: sums[6], 7: sums[7],
        8: 0 if e2 + sums[6] + sums[7] == 0 else e2 / (e2 + sums[6] + sums[7]),
        9: sums["maxI"], 10: sums[10], 11: sums[11], 12: sums[12], 13: sums[13],
        14: n_weighted / e2 if e2 else 0,
    }
    for col, (v, nf) in tot.items():
        put(ws, 2, col, v, nf=nf, bold=True, fill=LIGHT, border=True, formula=True, cached=tot_cache.get(col))

    setup(ws, NAVY, "C4", {"A": 25.7, "B": 13.9, "C": 15.0, "D": 18.0, "E": 17.5, "F": 17.3, "G": 23.0,
                           "H": 20.0, "I": 19.0, "J": 32.0, "K": 29.0, "L": 29.0, "M": 27.0, "N": 15.9,
                           "O": 16.0})
    ws.auto_filter.ref = f"A3:O{last}"
    ws.conditional_formatting.add(f"H{first}:H{last}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))
    ws.conditional_formatting.add(f"J{first}:K{last}", CellIsRule(
        operator="greaterThan", formula=["0"], font=RED_FONT,
        fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")))
    ws.conditional_formatting.add(f"N{first}:N{last}", ColorScaleRule(
        start_type="num", start_value=0.9, start_color="F8696B", mid_type="num", mid_value=0.96,
        mid_color="FFEB84", end_type="num", end_value=1, end_color="63BE7B"))


# ---- Population_Diff --------------------------------------------------------
def write_population(ws, results):
    title(ws, "Population Differences (Row-Level)")
    heads = ["Case_Number", "Load_Type", "Participant_ID", "PRT_ID", "Role", "Population_Status",
             "Participant_Status (context)", "DOB (context)", "Notes"]
    header(ws, 3, heads)
    r_ = 4
    for res in sorted(results, key=lambda r: LOAD_TYPES.index(r.load)):
        for d in res.pop:
            for col, v in enumerate([d["case"], d["load"], d["pid"], d["prt"], d["role"], d["status"],
                                     d["pstatus"], d["dob"], d["notes"] or None], 1):
                put(ws, r_, col, v)
            r_ += 1
    last = r_ - 1
    setup(ws, "C00000", "D4", {"A": 14, "B": 10, "C": 26, "D": 14, "E": 8, "F": 24, "G": 26, "H": 16, "I": 60})
    if last >= 4:
        add_table(ws, "tblPopDiff", f"A3:I{last}")
        ws.conditional_formatting.add(f"F4:F{last}", FormulaRule(formula=['LEFT($F4,7)="Only in"'], font=RED_FONT))


# ---- Element_Diff -------------------------------------------------------------
def aggregate_elements(results, load):
    agg = OrderedDict()
    for res in (r for r in results if r.load == load):
        for i, e in enumerate(res.elements):
            a = agg.setdefault(e["name"], {
                "name": e["name"], "in_p": False, "in_a": False, "dt_p": set(), "dt_a": set(),
                "fp_n": 0, "fp_d": 0, "fa_n": 0, "fa_d": 0, "cmp": 0, "mat": 0, "notes": [], "idx": i})
            a["in_p"] |= e["in_p"]
            a["in_a"] |= e["in_a"]
            if e["in_p"]:
                a["dt_p"].add(e["dt_p"])
                a["fp_n"] += (e["fill_p"] or 0) * res.rows_pbgc
                a["fp_d"] += res.rows_pbgc
            if e["in_a"]:
                a["dt_a"].add(e["dt_a"])
                a["fa_n"] += (e["fill_a"] or 0) * res.rows_auto
                a["fa_d"] += res.rows_auto
            a["cmp"] += e["compared"]
            a["mat"] += e["matched"]
            if e["notes"] and e["notes"] not in a["notes"]:
                a["notes"].append(e["notes"])
    out = []
    for a in agg.values():
        a["status"] = ("In Both" if a["in_p"] and a["in_a"] else
                       "Missing from Automation" if a["in_p"] else "Extra in Automation")
        out.append(a)
    rank = {"In Both": 0, "Missing from Automation": 1, "Extra in Automation": 2}
    out.sort(key=lambda a: (rank[a["status"]], a["idx"]))
    return out


def _dt(s):
    s = {x for x in s if x != "Empty"} if len(s) > 1 else s
    if not s:
        return ""
    return next(iter(s)) if len(s) == 1 else "Mixed"


def write_elements(ws, results):
    title(ws, "Element Differences (Column-Level)")
    heads = ["Load_Type", "Element_Name", "In_PBGC", "In_Automation", "Column_Status", "Data_Type_PBGC",
             "Data_Type_Automation", "Fill_Rate_PBGC", "Fill_Rate_Automation", "Value_Match_%", "Notes"]
    header(ws, 3, heads)
    r_ = 4
    schema = {}
    for load in LOAD_TYPES:
        aggs = aggregate_elements(results, load)
        schema[load] = (sum(1 for a in aggs if a["status"] == "In Both"), len(aggs))
        for a in aggs:
            vals = [
                load, a["name"], "Y" if a["in_p"] else "N", "Y" if a["in_a"] else "N", a["status"],
                _dt(a["dt_p"]) or None, _dt(a["dt_a"]) or None,
                a["fp_n"] / a["fp_d"] if a["fp_d"] else None,
                a["fa_n"] / a["fa_d"] if a["fa_d"] else None,
                a["mat"] / a["cmp"] if a["cmp"] else None,
                "; ".join(a["notes"]) or None,
            ]
            for col, v in enumerate(vals, 1):
                put(ws, r_, col, v, nf=nf_pct() if col in (8, 9, 10) else None)
            r_ += 1
    last = r_ - 1
    setup(ws, "C00000", "C4", {"A": 11, "B": 26, "C": 9, "D": 13, "E": 24, "F": 15, "G": 20, "H": 14,
                               "I": 20, "J": 14, "K": 70})
    if last >= 4:
        add_table(ws, "tblElemDiff", f"A3:K{last}")
        ws.conditional_formatting.add(f"E4:E{last}", FormulaRule(
            formula=['$E4<>"In Both"'], font=RED_FONT,
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")))
        ws.conditional_formatting.add(f"J4:J{last}", ColorScaleRule(
            start_type="num", start_value=0.85, start_color="F8696B", mid_type="num", mid_value=0.95,
            mid_color="FFEB84", end_type="num", end_value=1, end_color="63BE7B"))
    rr = max(last, 3) + 2
    for load in LOAD_TYPES:
        put(ws, rr, 1, f"Schema match rate ({load}):", bold=True)
        f = (f'=IF(COUNTIFS(A4:A{max(last, 4)},"{load}")=0,"",'
             f'COUNTIFS(A4:A{max(last, 4)},"{load}",E4:E{max(last, 4)},"In Both")/'
             f'COUNTIFS(A4:A{max(last, 4)},"{load}"))')
        in_both, total = schema[load]
        put(ws, rr, 4, f, nf=nf_pct(), bold=True, formula=True, cached=(in_both / total if total else ""))
        rr += 1


# ---- Heatmaps -----------------------------------------------------------------
def write_heatmap(ws, results, load):
    title(ws, f"Element x Plan Value-Match % ({load})")
    plan_order = [p["plan_name"] for p in PLANS]
    res_by_plan = OrderedDict((p, next((r for r in results if r.plan == p and r.load == load), None))
                              for p in plan_order)
    plans = [p for p, r in res_by_plan.items() if r is not None]
    ncol = len(plans) + 2
    header(ws, 3, ["Element_Name"] + plans + ["All Plans (avg)"], border=True, height=31.5)

    names = []
    for p in plans:
        for e in res_by_plan[p].elements:
            if e["status"] == "In Both" and e["name"] not in names:
                names.append(e["name"])
    widths = {"A": 30}
    for i in range(2, ncol + 1):
        widths[get_column_letter(i)] = 16
    setup(ws, "ED7D31", "B4", widths)
    if not names:
        put(ws, 4, 1, f"No {load} files supplied for this run.", italic=True, color="7F7F7F")
        return
    lastc = get_column_letter(ncol - 1)
    mean = lambda xs: (sum(xs) / len(xs)) if xs else ""
    grid = []   # grid[row][plan] value or None
    for i, nm in enumerate(names):
        n = 4 + i
        put(ws, n, 1, nm, bold=True)
        row_vals = []
        for j, p in enumerate(plans):
            e = next((x for x in res_by_plan[p].elements if x["name"] == nm and x["status"] == "In Both"), None)
            v = e["matched"] / e["compared"] if e and e["compared"] else None
            row_vals.append(v)
            put(ws, n, 2 + j, v, nf=nf_pct(), align="center", border=True)
        grid.append(row_vals)
        put(ws, n, ncol, f'=IFERROR(AVERAGE(B{n}:{lastc}{n}),"")', nf=nf_pct(), bold=True,
            align="center", border=True, formula=True, cached=mean([x for x in row_vals if x is not None]))
    last = 3 + len(names)
    pa = last + 1
    put(ws, pa, 1, "Plan avg", bold=True, fill=LIGHT)
    row_avgs = [mean([x for x in rv if x is not None]) for rv in grid]
    for j in range(2, ncol + 1):
        L = get_column_letter(j)
        col_vals = ([rv[j - 2] for rv in grid] if j < ncol else row_avgs)
        put(ws, pa, j, f'=IFERROR(AVERAGE({L}4:{L}{last}),"")', nf=nf_pct(), bold=True, align="center",
            border=True, fill=LIGHT, formula=True,
            cached=mean([x for x in col_vals if x is not None and x != ""]))
    ws.conditional_formatting.add(f"B4:{get_column_letter(ncol)}{pa}", ColorScaleRule(
        start_type="num", start_value=0.75, start_color="F8696B", mid_type="num", mid_value=0.92,
        mid_color="FFEB84", end_type="num", end_value=1, end_color="63BE7B"))


# ---- Value_Diff ---------------------------------------------------------------
def write_value_diff(ws, results, load, table_name):
    title(ws, f"Value Differences - {load}")
    header(ws, 3, ["Case_Number", "Participant_ID", "Element_Name", "PBGC_Value", "Automation_Value",
                   "Diff_Type", "Numeric_Diff", "Notes", "Root_Cause_Tag", "Triage_Status"])
    r_ = 4
    for res in (r for r in results if r.load == load):
        for d in res.value_rows:
            for col, v in ((1, d["case"]), (2, d["pid"]), (3, d["element"]), (4, d["pbgc"]), (5, d["auto"]),
                           (6, d["diff_type"]), (8, d["notes"] or None), (10, "Open")):
                put(ws, r_, col, v, align="left" if col in (4, 5) else None)
            is_num = lambda x: isinstance(x, numbers.Real) and not isinstance(x, bool)
            put(ws, r_, 7, f'=IF(AND(ISNUMBER(D{r_}),ISNUMBER(E{r_})),E{r_}-D{r_},"")',
                nf="#,##0.00;\\-#,##0.00", formula=True,
                cached=(d["auto"] - d["pbgc"]) if is_num(d["pbgc"]) and is_num(d["auto"]) else "")
            r_ += 1
    last = r_ - 1
    setup(ws, "C00000", "C4", {"A": 14, "B": 26, "C": 24, "D": 22, "E": 22, "F": 24, "G": 14, "H": 40,
                               "I": 18, "J": 14})
    if last >= 4:
        add_table(ws, table_name, f"A3:J{last}")
        ws.conditional_formatting.add(f"F4:F{last}", FormulaRule(
            formula=['LEFT($F4,7)="Missing"'], font=Font(name=ARIAL, bold=True, color="9C0006")))
        ws.conditional_formatting.add(f"F4:F{last}", FormulaRule(
            formula=['$F4="Format Only"'], font=Font(name=ARIAL, italic=True, color="7F7F7F")))
    else:
        put(ws, 4, 1, f"No {load} differences (or no {load} files supplied for this run).",
            italic=True, color="7F7F7F")


# ---- Participant_Summary -----------------------------------------------------
def write_participants(ws, results):
    title(ws, "Participant Summary")
    header(ws, 3, ["Case_Number", "Load_Type", "Participant_ID", "Elements_Compared", "Elements_Different",
                   "Elements_Missing_in_Automation", "Elements_Missing_in_PBGC", "Match_%", "Diff_Elements"])
    rows = [d for res in results for d in res.part_rows]
    # worst first across all loads / plans
    rows.sort(key=lambda d: (-d["different"] / d["compared"] if d["compared"] else 0, -d["different"], d["pid"]))
    for i, d in enumerate(rows):
        n = 4 + i
        for col, v in enumerate([d["case"], d["load"], d["pid"], d["compared"], d["different"],
                                 d["miss_aut"], d["miss_pbgc"]], 1):
            put(ws, n, col, v)
        put(ws, n, 8, f"=IF(D{n}=0,0,1-E{n}/D{n})", nf=nf_pct(), formula=True,
            cached=0 if d["compared"] == 0 else 1 - d["different"] / d["compared"])
        put(ws, n, 9, d["diff_elements"])
    last = 3 + len(rows)
    setup(ws, "548235", "D4", {"A": 14, "B": 10, "C": 26, "D": 12, "E": 12, "F": 16, "G": 14, "H": 10, "I": 80})
    if rows:
        add_table(ws, "tblPartSummary", f"A3:I{last}")
        ws.conditional_formatting.add(f"H4:H{last}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))
    else:
        put(ws, 4, 1, "No participants with differences.", italic=True, color="7F7F7F")


# ---- README ------------------------------------------------------------------
def write_readme(ws, results):
    print_setup(ws)
    ws.sheet_properties.tabColor = "7F7F7F"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110
    put(ws, 1, 1, "PBGC Load File Alignment Analysis", bold=True, color=NAVY, size=16)
    ws.row_dimensions[1].height = 20
    put(ws, 3, 1, "Workbook contents (click to jump)", bold=True)
    toc = [
        ("Summary_by_Plan", "One row per Case_Number x Load_Type. Population + element + cell-level alignment scores. Source for slide stats."),
        ("Population_Diff", "Objective 1. Row-level: participants in only one file, duplicate keys, role mismatches."),
        ("Element_Diff", "Objective 2. Column-level: elements missing / extra, position, fill rates, value match % per element."),
        ("Heatmap_9A", "Element x Plan value-match % (9A). Systemic vs. plan-specific issues at a glance."),
        ("Heatmap_Payee", "Element x Plan value-match % (Payee)."),
        ("Value_Diff_9A", "Objective 3. One row per participant x element disagreement (9A). Filterable."),
        ("Value_Diff_Payee", "Objective 3. One row per participant x element disagreement (Payee)."),
        ("Participant_Summary", "One row per participant with any difference; worst offenders first."),
    ]
    for i, (sheet, desc) in enumerate(toc):
        r = 4 + i
        c = put(ws, r, 1, sheet, color="0563C1")
        c.font = Font(name=ARIAL, size=10, color="0563C1", underline="single")
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sheet}'!A1", display=sheet)
        put(ws, r, 2, desc)

    put(ws, 14, 1, "Key definitions", bold=True)
    defs = [
        ("Participant_ID", f"Manual file '{MANUAL_KEY_COLUMN}' (LAST+FIRST+last {UID_PRT_DIGITS} of PRT_ID+first {UID_ROLE_CHARS} of CUST_TYPE, e.g. SMITHJANE9999P). "
                           "Derived the same way for the automation file. Primary match key within a Case_Number and Load_Type."),
        ("Case_Number", "PBGC plan case number (CASE_ID_NUMBER)."),
        ("Load_Type", "9A or Payee (one load file of each type per plan)."),
        ("PBGC / Automation", "PBGC = the manual file. Automation = the automation file."),
    ]
    r = 15
    for k, v in defs:
        put(ws, r, 1, k, bold=True)
        put(ws, r, 2, v)
        r += 1

    r += 1
    put(ws, r, 1, "Status vocabularies", bold=True)
    r += 1
    for k, v in [
        ("Population_Status", "Only in PBGC | Only in Automation | Duplicate in PBGC | Duplicate in Automation | Role Mismatch"),
        ("Column_Status", "In Both | Missing from Automation | Extra in Automation"),
        ("Diff_Type", "Value Mismatch | Missing in Automation | Missing in PBGC | Format Only"),
    ]:
        put(ws, r, 1, k, bold=True)
        put(ws, r, 2, v)
        r += 1

    r += 1
    put(ws, r, 1, "How it is calculated", bold=True)
    r += 1
    notes = [
        ("Rows_Matched", "Participant_IDs that appear exactly once in each file. Rows_Only_PBGC / Rows_Only_Automation are the remaining rows "
                         "(only-in, duplicate-key and role-mismatch rows), so they equal Rows - Rows_Matched."),
        ("Duplicates", "Participant_IDs repeated within one file are listed on Population_Diff and excluded from value comparison until resolved."),
        ("Role Mismatch", "A row whose key is in only one file, where the other file has the same PRT_ID under a different CUST_TYPE (one line per side)."),
        ("Equivalences", "Blank = null; text is trimmed and case-insensitive; dates are parsed before comparing; numbers use a "
                         f"{NUMERIC_TOLERANCE} tolerance; leading zeros / punctuation differences are 'Format Only'; code tables in VALUE_MAPS "
                         "(Sex, Country, State, Eligibility Status, ACH_ACCOUNT_TYPE) are treated as equivalent."),
        ("Format Only", "Listed on Value_Diff but "
                        + ("counted as a difference." if FORMAT_ONLY_COUNTS_AS_DIFF else "NOT counted as a difference in Cell_Match_% or Participant_Summary.")),
        ("Elements_Different", "On Participant_Summary: all counted differences (value mismatches + missing); the two Missing columns are subsets of it."),
        ("Cell_Match_%", "Matched cells / (Rows_Matched x Elements_Compared). Both-blank cells count as matches. "
                         "Shows 'n/a' when a plan has no matched participants (stored as 0 and weighted by Rows_Matched, so it does not affect ALL PLANS)."),
        ("Hard-coded values", "Row counts, element counts and match % on Summary_by_Plan and the heatmaps are computed by the script and written as values; "
                              "totals, % columns, heatmap averages and Numeric_Diff are live formulas."),
        ("Triage columns", "Root_Cause_Tag is left blank for analysts; Triage_Status starts at 'Open'."),
    ]
    for k, v in notes:
        put(ws, r, 1, k, bold=True)
        put(ws, r, 2, v)
        r += 1

    r += 1
    put(ws, r, 1, "Run information", bold=True)
    r += 1
    put(ws, r, 1, "Generated", bold=True)
    put(ws, r, 2, datetime.now().strftime("%Y-%m-%d %H:%M"))
    r += 1
    for res in sorted(results, key=lambda x: (x.plan, LOAD_TYPES.index(x.load))):
        put(ws, r, 1, f"{res.plan} / {res.load}", bold=True)
        put(ws, r, 2, f"Case {res.case}: {res.rows_pbgc} PBGC rows, {res.rows_auto} Automation rows, {res.rows_matched} matched.")
        r += 1


# =============================================================================
# Main
# =============================================================================
def main():
    results = []
    for plan in PLANS:
        for load, key in (("9A", "9a"), ("Payee", "payee")):
            files = plan.get(key)
            if files:
                results.append(analyze_load(plan["plan_name"], load, files["manual"], files["automation"]))
    if not results:
        raise SystemExit("No file pairs configured in PLANS.")
    write_workbook(results)


if __name__ == "__main__":
    main()
