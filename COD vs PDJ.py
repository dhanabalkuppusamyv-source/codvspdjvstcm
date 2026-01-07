import streamlit as st
import pandas as pd
import numpy as np
import os, io, re, unicodedata
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ==============================
# Style & helpers (UI)
# ==============================
st.set_page_config(page_title="COD Compare", layout="wide")

logo_url = "https://raw.githubusercontent.com/Uthraa-18/cod-compare-app/refs/heads/main/image.png"

st.markdown(
    f"""
    <style>
    .corner-logo {{
        position: fixed;
        top: 50px;
        right: 20px;
        width: 120px;
        z-index: 9999;
    }}
    </style>

    <img src="{logo_url}" class="corner-logo">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <style>
    .section-title {
        font-size: 1.35rem;
        font-weight: 700;
        display: inline-flex;
        align-items: center;
        gap: .5rem;
        margin: .25rem 0 .5rem 0;
    }
    .info-dot {
        display:inline-block;
        font-size: 0.95rem;
        line-height: 1;
        padding: .1rem .35rem;
        border-radius: 999px;
        border: 1px solid #aaa;
        color: #333;
        cursor: help;
    }
    .subtle {
        font-size: 0.95rem;
        color: #555;
        margin-top: .25rem;
    }
    .small-input .stNumberInput > div > div > input {
        font-size: .9rem;
    }
    .block-container {
        padding-top: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def header_with_tip(text: str, tip: str):
    st.markdown(
        f"<div class='section-title'>{text}"
        f"<span class='info-dot' title='{tip}'>ⓘ</span></div>",
        unsafe_allow_html=True
    )

# ==============================
# Regex & utilities
# ==============================
RE_PM = re.compile(r'(?:±|\+/-)\s*(\d+(?:[.,]\d+)?)', re.I)
RE_SIGNED = re.compile(r'^[\+\-]?\s*\d+(?:[.,]\d+)?$')
RE_NUM = re.compile(r'[-+]?\d+(?:[.,]\d+)?')

def to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return None

def norm(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip().replace("’", "'")

def get_ext(name):
    return os.path.splitext(name)[-1].lower()

def read_all_sheets(name, file_bytes):
    engine = "xlrd" if get_ext(name) == ".xls" else "openpyxl"
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    return {
        s: pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, engine=engine, header=None)
        for s in xls.sheet_names
    }

# ==============================
# COD extraction helpers
# ==============================
def find_codification_value_below(cod_sheets, label="codification", scan_down=30):
    target = norm(label)
    for sname, df in cod_sheets.items():
        R, C = df.shape
        for r in range(R):
            for c in range(C):
                if norm(df.iat[r, c]) == target:
                    for rr in range(r + 1, min(R, r + 1 + scan_down)):
                        if norm(df.iat[rr, c]) != "":
                            return sname, str(df.iat[rr, c]).strip(), r, c
    return None, None, None, None

def find_stacked_anchor_vertical(df, words, max_gap=10):
    R, C = df.shape
    W = [w.lower() for w in words]
    for c in range(C):
        starts = [r for r in range(R) if W[0] in norm(df.iat[r, c])]
        for r0 in starts:
            rcur = r0
            ok = True
            for w in W[1:]:
                found = False
                for rr in range(rcur + 1, min(R, rcur + 1 + max_gap)):
                    if w in norm(df.iat[rr, c]):
                        rcur = rr
                        found = True
                        break
                if not found:
                    ok = False
                    break
            if ok:
                return rcur, c
    return None, None

def first_number_below(df, start_row, col, right_span=12, down_rows=4):
    R, C = df.shape
    for rr in range(start_row + 1, min(R, start_row + 1 + down_rows)):
        for cc in range(col, min(C, col + right_span)):
            s = "" if pd.isna(df.iat[rr, cc]) else str(df.iat[rr, cc])

            m = RE_PM.search(s)
            if m:
                v = to_float(m.group(1))
                if v is not None:
                    return v, rr, cc

            if norm(s) in {"±", "+/-"}:
                for cc2 in range(cc + 1, min(C, cc + 4)):
                    s2 = "" if pd.isna(df.iat[rr, cc2]) else str(df.iat[rr, cc2])
                    m2 = RE_NUM.search(s2)
                    if m2:
                        v = to_float(m2.group(0))
                        if v is not None:
                            return v, rr, cc2

            m3 = RE_NUM.search(s)
            if m3:
                v = to_float(m3.group(0))
                if v is not None:
                    return v, rr, cc
    return None, None, None

def two_signed_values_below_same_column(df, start_row, col, max_rows=8):
    R, _ = df.shape
    vals = []
    for rr in range(start_row + 1, min(R, start_row + 1 + max_rows)):
        s = "" if pd.isna(df.iat[rr, col]) else str(df.iat[rr, col]).strip()
        if RE_SIGNED.match(s):
            x = to_float(s)
            if x is not None:
                vals.append(x)
    for v in vals:
        if -v in vals:
            return abs(v), -abs(v)
    return None, None

# ==============================
# PDJ / TCM extraction
# ==============================
def row_numbers(df, r):
    nums = []
    row = df.iloc[r, :].tolist()
    for v in row:
        s = "" if pd.isna(v) else str(v)
        for m in RE_NUM.findall(s):
            x = to_float(m)
            if x is not None:
                nums.append(x)
    return nums

def sheet_numbers(df):
    nums = []
    for rr in range(df.shape[0]):
        nums.extend(row_numbers(df, rr))
    return nums

def find_key_positions(df, key):
    pos = []
    R, C = df.shape
    for r in range(R):
        for c in range(C):
            if str(df.iat[r, c]).strip() == key:
                pos.append((r, c))
    return pos

def approx_equal(a, b, tol):
    return abs(a - b) <= tol

def contains_value_eps(nums, val, tol):
    return any(approx_equal(x, val, tol) for x in nums)

def contains_pm_pair_eps(nums, mag, tol):
    return (
        any(approx_equal(x, abs(mag), tol) for x in nums) and
        any(approx_equal(x, -abs(mag), tol) for x in nums)
    )

def fmt_pm(m):
    s = f"{abs(m):.2f}".rstrip("0").rstrip(".")
    return f"+/- {s}"

# ==============================
# App UI
# ==============================
st.title("🔎 COD, PDJ, TCM Automatic Validation")

header_with_tip(
    "What this does",
    "Extracts Nominal & Tolerance from COD and compares PDJ/TCM rows."
)

st.caption("Epsilon allows 1.41 ≈ 1.4")

with st.container():
    st.markdown("<div class='small-input'>", unsafe_allow_html=True)
    eps = st.number_input(
        "Numeric tolerance (epsilon)",
        0.0, 0.2, 0.02, 0.01
    )
    st.markdown("</div>", unsafe_allow_html=True)

header_with_tip("Upload COD workbook (.xls/.xlsx)", "Reads COD data")
cod_file = st.file_uploader("", type=["xls", "xlsx"])

header_with_tip("Upload PDJ / TCM files", "Multiple files supported")
other_files = st.file_uploader("", type=["xls", "xlsx"], accept_multiple_files=True)

# ==============================
# Main logic
# ==============================
if cod_file and other_files:
    cod_bytes = cod_file.read()
    cod_sheets = read_all_sheets(cod_file.name, cod_bytes)

    s_cod, key_value, _, _ = find_codification_value_below(cod_sheets)
    if not key_value:
        st.error("Codification not found")
        st.stop()

    df_cod = cod_sheets[s_cod]

    nr, nc = find_stacked_anchor_vertical(df_cod, ["objectif", "nominal", "jeu"])
    cod_nominal, _, _ = first_number_below(df_cod, nr, nc)

    tr, tc = find_stacked_anchor_vertical(df_cod, ["calcul", "disp"])
    pm, _, _ = first_number_below(df_cod, tr, tc)
    tol_mag = abs(pm)

    st.success("COD reference extracted successfully")

    st.write("**Nominal:**", cod_nominal)
    st.write("**Tolerance:**", fmt_pm(tol_mag))
