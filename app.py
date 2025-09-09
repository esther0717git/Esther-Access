import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import re
import unicodedata
from datetime import datetime

# -------------------------------------------------
# Streamlit page config
# -------------------------------------------------
st.set_page_config(page_title="Data Center Format Converter 🌟 Esther", layout="centered")

# -------------------------------------------------
# Helpers
# -------------------------------------------------
def safe_company_name(df: pd.DataFrame) -> str:
    co_series = df.get("company full name")
    if co_series is not None:
        co_series = co_series.dropna()
        if not co_series.empty:
            return str(co_series.iloc[0])
    return "UnknownCompany"

def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Replace NaN/NA/Inf with None so we can write clean blanks to Excel."""
    return df.replace({pd.NA: None, np.nan: None, np.inf: None, -np.inf: None})

def write_row(ws, r, row_values, cell_fmt):
    """Write a row, using write_blank() for None/NaN/Inf."""
    for c, val in enumerate(row_values):
        is_blank = (
            val is None
            or (isinstance(val, float) and (math.isnan(val) or math.isinf(val)))
        )
        if is_blank:
            ws.write_blank(r, c, None, cell_fmt)
        else:
            ws.write(r, c, val, cell_fmt)

def make_safe_filename(stem: str, ext: str = ".xlsx") -> str:
    """
    Keep only alphanumeric, space, underscore (no hyphens).
    Convert any other character (including '-') to underscore.
    Collapse repeats and trim.
    """
    norm = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    norm = norm.replace("-", "_")
    norm = re.sub(r"[^A-Za-z0-9 _]", "_", norm)
    norm = re.sub(r"[ _]+", "_", norm).strip("_")
    return f"{norm}{ext}"

def safe_sheet_name(name: str) -> str:
    """Excel sheet names must be <= 31 chars and cannot contain : \ / ? * [ ]"""
    name = re.sub(r'[:\\/?*\[\]]', '_', str(name))
    return name[:31] if len(name) > 31 else name

def clean_phone(x, target_len=8):
    """Return an 8-digit phone from mixed inputs.
    - Handles floats like 86984997.0 without adding a trailing 0
    - Strips all non-digits
    - If longer than target_len, keep the LAST target_len digits
    """
    if pd.isna(x):
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, float):
        if np.isnan(x) or np.isinf(x):
            return ""
        s = str(int(x))
    else:
        s = re.sub(r"\D+", "", str(x))
    if target_len and len(s) > target_len:
        s = s[-target_len:]
    return s

def full_name_from(df: pd.DataFrame) -> pd.Series:
    if "full name as per nric" in df.columns:
        return df["full name as per nric"]
    first = df.get("first name as per nric", "")
    last  = df.get("middle and last name as per nric", "")
    return (first.fillna("").astype(str) + " " + last.fillna("").astype(str)).str.strip()

def as_str_or_blank(series: pd.Series | None, length: int) -> pd.Series:
    """Convert a Series to string values but keep missing as '' (never 'nan')."""
    if series is None:
        return pd.Series([""] * length)
    return series.apply(lambda x: "" if pd.isna(x) else str(x))

# -------------------------------------------------
# RC preset(s)
# -------------------------------------------------
RC_PRESETS = {
    "Default (as requested)": {
        "Visitor Category": "Visitor Access",     # A
        "Visitor Email": "liangwy@sea.com",       # D
        "Designation (E.g. Supervisor, Engineer) (UDF)": "Contractor",  # H
        "Purpose of Visit (UDF)": "access + relocation",                 # I
        "Level (UDF)": "3,4",                     # J
        "Location (UDF)": "All Halls",            # K
        "Rack ID (E.g. 71A01, 71A02) (UDF)": "31A10",                    # L
        "Host (UDF)": "Esther",                   # M
    }
}

# -------------------------------------------------
# AT Format
# -------------------------------------------------
def convert_to_at_dc(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["first name as per nric"].notna()]
    try:
        df_out = pd.DataFrame({
            "First Name":        df["first name as per nric"],
            "Last Name":         df["middle and last name as per nric"],
            "Email Address":     ["liangwy@sea.com"] * len(df),
            "Company":           df["company full name"],
            "Other IC Number":   df["ic (last 3 digits and suffix) 123a"]
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column: {e}")
        return None, None

# -------------------------------------------------
# SG DRT Format (13-column)
# -------------------------------------------------
def convert_to_sg_drt_dc(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["first name as per nric"].notna()]
    try:
        df_out = pd.DataFrame({
            "First Name":    df["first name as per nric"],
            "Last Name":     df["middle and last name as per nric"],
            "Email Address": ["liangwy@sea.com"] * len(df),
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column (SG DRT): {e}")
        return None, None

# -------------------------------------------------
# US DRT Format (11-column; E=First, F=Last)
# -------------------------------------------------
def convert_to_us_drt_dc(df: pd.DataFrame):
    raw = df.dropna(how="all").copy()
    if raw.shape[1] < 6:
        st.error("❌ US DRT sheet too narrow — expected ≥ 6 columns (E & F needed).")
        return None, None

    first = raw.iloc[:, 4]  # Column E (0-based index 4)
    last  = raw.iloc[:, 5]  # Column F (0-based index 5)

    mask = first.notna() & (first.astype(str).str.strip() != "")
    first = first[mask].astype(str).str.strip()
    last  = last[mask].fillna("").astype(str).str.strip()

    df_out = pd.DataFrame({
        "First Name":    first,
        "Last Name":     last,
        "Email Address": ["liangwy@sea.com"] * mask.sum(),
    })

    # Company may not exist in this layout; use safe fallback
    return sanitize_df(df_out), "UnknownCompany"

# -------------------------------------------------
# EQ Format
# -------------------------------------------------
def convert_to_eq(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["first name as per nric"].notna()]
    try:
        df_out = pd.DataFrame({
            "Legal First Name":       df["first name as per nric"],
            "Legal Last Name":        df["middle and last name as per nric"],
            "Company":                df["company full name"],
            "Email (Optional)":       ["liangwy@sea.com"] * len(df),
            "Country Code (Optional)": ["" for _ in range(len(df))],
            "Mobile Phone (Optional)": ["" for _ in range(len(df))]
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column: {e}")
        return None, None

# -------------------------------------------------
# STT Loyang Format
# -------------------------------------------------
def convert_to_sttly(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["full name as per nric"].notna()]
    try:
        nationality = df.get("nationality (country name)", pd.Series([""] * len(df)))
        mobile = df.get("mobile number", pd.Series([""] * len(df)))
        df_out = pd.DataFrame({
            "Name*": full_name_from(df),
            "Company*": df["company full name"],
            "Type (Visitor, Contractor)*": df["company full name"].apply(
                lambda x: "Visitor" if "sea" in str(x).lower() else "Contractor"
            ),
            "ID No.* (Only last 4 characters will be stored.)": df["ic (last 3 digits and suffix) 123a"],
            "Country (Will default to Singapore if left as blank)": nationality,
            "Business Email (optional)": ["" for _ in range(len(df))],
            "Business Phone* (If your number is not local, please input IDD Code without the +)":
                mobile.apply(clean_phone),
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column: {e}")
        return None, None

# -------------------------------------------------
# Racks Central (RC)
# -------------------------------------------------
def convert_to_rc(df: pd.DataFrame, preset_name="Default (as requested)"):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")

    # keep only rows with a non-empty name
    initial_names = full_name_from(df)
    mask = initial_names.notna() & (initial_names.astype(str).str.strip() != "")
    df = df[mask].copy()

    preset = RC_PRESETS.get(preset_name, RC_PRESETS["Default (as requested)"])

    try:
        length = len(df)
        name_series = full_name_from(df).fillna("").astype(str).str.strip()
        nric_series = as_str_or_blank(df.get("ic (last 3 digits and suffix) 123a"), length)
        mobile = df.get("mobile number", pd.Series([""] * length)).apply(clean_phone)
        plates = (
            as_str_or_blank(df.get("vehicle plate number"), length)
            .str.replace(";", ",", regex=False)
            .str.strip()
        )
        company = as_str_or_blank(df.get("company full name"), length)

        df_out = pd.DataFrame({
            "Visitor Category": [preset["Visitor Category"]] * length,
            "Visitor Name": name_series,
            "Visitor NRIC/Passport": nric_series,
            "Visitor Email": [preset["Visitor Email"]] * length,
            "Visitor Contact No": mobile,
            "Visitor Vehicle Number": plates,
            "Visitor Company (UDF)": company,
            "Designation (E.g. Supervisor, Engineer) (UDF)":
                [preset["Designation (E.g. Supervisor, Engineer) (UDF)"]] * length,
            "Purpose of Visit (UDF)": [preset["Purpose of Visit (UDF)"]] * length,
            "Level (UDF)": [preset["Level (UDF)"]] * length,
            "Location (UDF)": [preset["Location (UDF)"]] * length,
            "Rack ID (E.g. 71A01, 71A02) (UDF)":
                [preset["Rack ID (E.g. 71A01, 71A02) (UDF)"]] * length,
            "Host (UDF)": [preset["Host (UDF)"]] * length,
        })

        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column for RC: {e}")
        return None, None

# -------------------------------------------------
# App UI
# -------------------------------------------------
st.title(" DC Access 🌟 Esther 🌟")

uploaded_file = st.file_uploader("Upload the original visitor list (.xlsx)", type=["xlsx"])

format_type = st.selectbox(
    "Select the Data Center format to convert to",
    [
        "AT",
        "SG DRT",   # 13-col DRT (SIN12/Shopee)
        "US DRT",   # 11-col DRT (e.g., DFW14)
        "EQSG4|SG5|USDA11|USDC15",
        "STTLY",
        "RC",
    ]
)

rc_preset_name = None
if format_type == "RC":
    rc_preset_name = st.selectbox(
        "RC defaults",
        list(RC_PRESETS.keys()),
        index=list(RC_PRESETS.keys()).index("Default (as requested)")
    )

if uploaded_file and format_type:
    df = pd.read_excel(uploaded_file)

    if format_type == "AT":
        converted_df, company_name = convert_to_at_dc(df)
    elif format_type == "SG DRT":
        converted_df, company_name = convert_to_sg_drt_dc(df)
    elif format_type == "US DRT":
        converted_df, company_name = convert_to_us_drt_dc(df)
    elif format_type == "EQSG4|SG5|USDA11|USDC15":
        converted_df, company_name = convert_to_eq(df)
    elif format_type == "STTLY":
        converted_df, company_name = convert_to_sttly(df)
    elif format_type == "RC":
        converted_df, company_name = convert_to_rc(df, preset_name=rc_preset_name)
    else:
        converted_df, company_name = None, None

    if converted_df is not None:
        output = io.BytesIO()
        with pd.ExcelWriter(
            output,
            engine="xlsxwriter",
            engine_kwargs={"options": {"nan_inf_to_errors": True}}
        ) as writer:
            # Excel-safe sheet/tab name (≤ 31 chars)
            sheet = safe_sheet_name(format_type)
            converted_df.to_excel(writer, index=False, sheet_name=sheet)
            wb = writer.book
            ws = writer.sheets[sheet]

            header_fmt = wb.add_format({
                "bold": True, "border": 1,
                "align": "center", "valign": "vcenter",
                "bg_color": "#548135", "font_color": "white"
            })
            cell_fmt = wb.add_format({
                "border": 1, "align": "center", "valign": "vcenter"
            })

            # Re-write headers with formatting
            for col_idx, col_name in enumerate(converted_df.columns):
                ws.write(0, col_idx, col_name, header_fmt)

            # Body rows (handle blanks properly)
            for r, row in enumerate(converted_df.itertuples(index=False, name=None), start=1):
                write_row(ws, r, row, cell_fmt)

            # Autosize columns
            for i, col in enumerate(converted_df.columns):
                data_max = converted_df[col].astype(str).map(len).max() if not converted_df.empty else 0
                ws.set_column(i, i, max(len(col), data_max) + 2)

            ws.set_default_row(18)

        output.seek(0)
        date_str = datetime.today().strftime("%Y%m%d")
        stem = f"Upload_{format_type}_{company_name or 'UnknownCompany'}_{date_str}"
        fname = make_safe_filename(stem, ".xlsx")

        st.success("✅ Conversion completed! Download below:")
        st.download_button(
            "📥 Download Converted Excel",
            data=output,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
