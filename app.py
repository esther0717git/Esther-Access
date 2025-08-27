import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import re
import unicodedata
from datetime import datetime

# -------------------------------
# Helpers
# -------------------------------
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
    Keep only alphanumeric, space, underscore.
    Convert hyphens and all other characters to underscore.
    Collapse repeats and trim.
    """
    # fold to ASCII (drop accents/emoji)
    norm = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    # replace hyphens with underscore explicitly
    norm = norm.replace("-", "_")
    # allow only [A-Za-z0-9 _]; replace others with underscore
    norm = re.sub(r"[^A-Za-z0-9 _]", "_", norm)
    # collapse multiple spaces/underscores
    norm = re.sub(r"[ _]+", "_", norm).strip("_")
    return f"{norm}{ext}"

# -------------------------------
# AT Format
# -------------------------------
def convert_to_at_dc(df):
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

# -------------------------------
# DRT Format
# -------------------------------
def convert_to_drt_dc(df):
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["first name as per nric"].notna()]
    try:
        df_out = pd.DataFrame({
            "First Name":    df["first name as per nric"],
            "Last Name":     df["middle and last name as per nric"],
            "Email Address": ["chenh@sea.com"] * len(df),
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column: {e}")
        return None, None

# -------------------------------
# EQ Format
# -------------------------------
def convert_to_eq(df):
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["first name as per nric"].notna()]
    try:
        df_out = pd.DataFrame({
            "Legal First Name":       df["first name as per nric"],
            "Legal Last Name":        df["middle and last name as per nric"],
            "Company":                df["company full name"],
            "Email (Optional)":       ["chenh@sea.com"] * len(df),
            "Country Code (Optional)": ["" for _ in range(len(df))],
            "Mobile Phone (Optional)": ["" for _ in range(len(df))]
        })
        return sanitize_df(df_out), safe_company_name(df)
    except KeyError as e:
        st.error(f"❌ Missing expected column: {e}")
        return None, None

# -------------------------------
# STTLY Format
# -------------------------------
def convert_to_sttly(df):
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")
    df = df[df["full name as per nric"].notna()]
    try:
        nationality = df.get("nationality (country name)", pd.Series([""] * len(df)))
        mobile = df.get("mobile number", pd.Series([""] * len(df)))

        def clean_phone(x):
            if pd.isna(x):
                return ""
            return "".join(ch for ch in str(x) if ch.isdigit())

        df_out = pd.DataFrame({
            "Name*": df["full name as per nric"],
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

# -------------------------------
# Streamlit App
# -------------------------------
st.set_page_config(page_title="Data Center Format Converter 🌟 Murphy", layout="centered")
st.title("📮 DC Access 🌟 Murphy 🌟")

uploaded_file = st.file_uploader("Upload the original visitor list (.xlsx)", type=["xlsx"])
format_type = st.selectbox(
    "Select the Data Center format to convert to",
    ["AT", "DRT", "EQ", "STTLY"]
)

if uploaded_file and format_type:
    df = pd.read_excel(uploaded_file)

    # route to the correct converter
    if format_type == "AT":
        converted_df, company_name = convert_to_at_dc(df)
    elif format_type == "DRT":
        converted_df, company_name = convert_to_drt_dc(df)
    elif format_type == "EQ":
        converted_df, company_name = convert_to_eq(df)
    elif format_type == "STTLY":
        converted_df, company_name = convert_to_sttly(df)
    else:
        converted_df, company_name = None, None

    # export & styling
    if converted_df is not None:
        output = io.BytesIO()
        with pd.ExcelWriter(
            output,
            engine="xlsxwriter",
            engine_kwargs={"options": {"nan_inf_to_errors": True}}
        ) as writer:
            sheet = format_type
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

            # headers
            for col_idx, col_name in enumerate(converted_df.columns):
                ws.write(0, col_idx, col_name, header_fmt)

            # rows
            for r, row in enumerate(converted_df.itertuples(index=False, name=None), start=1):
                write_row(ws, r, row, cell_fmt)

            # auto-size columns
            for i, col in enumerate(converted_df.columns):
                data_max = converted_df[col].astype(str).map(len).max() if not converted_df.empty else 0
                ws.set_column(i, i, max(len(col), data_max) + 2)

            ws.set_default_row(18)

        output.seek(0)
        date_str = datetime.today().strftime("%Y%m%d")
        # Build a compliant filename: only A–Z a–z 0–9 space underscore; no hyphens
        stem = f"Upload_{format_type}_{company_name or 'UnknownCompany'}_{date_str}"
        fname = make_safe_filename(stem, ".xlsx")

        st.success("✅ Conversion completed! Download below:")
        st.download_button(
            "📥 Download Converted Excel",
            data=output,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


