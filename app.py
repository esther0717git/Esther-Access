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
    Keep only alphanumeric, space, underscore (no hyphens).
    Convert any other character (including '-') to underscore.
    Collapse repeats and trim.
    """
    norm = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    norm = norm.replace("-", "_")
    norm = re.sub(r"[^A-Za-z0-9 _]", "_", norm)
    norm = re.sub(r"[ _]+", "_", norm).strip("_")
    return f"{norm}{ext}"

def clean_phone(x):
    if pd.isna(x):
        return ""
    return "".join(ch for ch in str(x) if ch.isdigit())

def full_name_from(df: pd.DataFrame) -> pd.Series:
    if "full name as per nric" in df.columns:
        return df["full name as per nric"]
    # fallback: stitch from parts if available
    first = df.get("first name as per nric", "")
    last  = df.get("middle and last name as per nric", "")
    return (first.fillna("").astype(str) + " " + last.fillna("").astype(str)).str.strip()

# -------------------------------
# RC preset(s) - add more if needed
# -------------------------------
RC_PRESETS = {
    "Default (as requested)": {
        # A, D, H, I, J, K, L, M are fixed from this preset
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

# -------------------------------
# RC Format (uses preset for A, D, H, I, J, K, L, M)
# -------------------------------
def convert_to_rc(df, preset_name="Default (as requested)"):
    """
    Build RC sheet:
      A Visitor Category              <- preset
      B Visitor Name                  <- from file
      C Visitor NRIC/Passport         <- from file
      D Visitor Email                 <- preset
      E Visitor Contact No            <- from file (digits only)
      F Visitor Vehicle Number        <- from file (';' -> ',')
      G Visitor Company (UDF)         <- from file
      H Designation (UDF)             <- preset
      I Purpose of Visit (UDF)        <- preset
      J Level (UDF)                   <- preset
      K Location (UDF)                <- preset
      L Rack ID (UDF)                 <- preset
      M Host (UDF)                    <- preset
    """
    df.columns = df.columns.str.strip().str.lower()
    df = df.dropna(how="all")

    name_series = full_name_from(df)
    df = df[name_series.notna() & (name_series.astype(str).str.strip() != "")].copy()

    preset = RC_PRESETS.get(preset_name, RC_PRESETS["Default (as requested)"])

    try:
        mobile = df.get("m
