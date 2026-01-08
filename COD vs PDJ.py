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

st.markdown(f"""
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
""", unsafe_allow_html=True)

# ==============================
# Regex & utilities
# ==============================
RE_PM = re.compile(r'(?:±|\+/-)\s*(\d+(?:[.,]\d+)?)', re.I)
RE_SIGNED = re.compile(r'^[\+\-]?\s*\d+(?:[.,]\d+)?$')
RE_NUM = re.compile(r'[-+]?\d+(?:[.,]\d+)?')

def to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except:
        return None

def norm(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()

def get_ext(name):
    return os.path.splitext(name)[-1].lower()

def read_all_sheets(name, file_bytes):
    engine = "xlrd" if get_ext(name)==".xls" else "openpyxl"
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    return {
        s: pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, engine=engine, header=None)
        for s in xls.sheet_names
    }

# ==============================
# Matching helpers
# ==============================
def approx_equal(a,b,tol):
    return abs(a-b) <= tol

def contains_value_eps(nums, val, tol):
    return any( approx_equal(x,val,tol) for x in nums )

def contains_pm_pair_eps(nums, mag, tol):
    return (
        any( approx_equal(x,+abs(mag),tol) for x in nums ) and
        any( approx_equal(x,-abs(mag),tol) for x in nums )
    )

def fmt_pm(m):
    s = f"{abs(m):.2f}".rstrip("0").rstrip(".")
    return f"+/- {s}"

def row_numbers(df, r):
    nums=[]
    row = df.iloc[r,:].tolist()
    for v in row:
        for m in RE_NUM.findall(str(v)):
            x = to_float(m)
            if x is not None:
                nums.append(x)
    return nums

def find_key_positions(df, key):
    pos=[]
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if str(df.iat[r,c]).strip() == key:
                pos.append((r,c))
    return pos

# ==============================
# UI
# ==============================
st.title("🔎 COD, PDJ, TCM Automatic Validation")

eps = st.number_input("Numeric tolerance (epsilon)", 0.0,0.2,0.02,0.01)

cod_file = st.file_uploader("Upload COD file", type=["xls","xlsx"])
other_files = st.file_uploader("Upload PDJ / TCM files", type=["xls","xlsx"], accept_multiple_files=True)

# ==============================
# Main logic
# ==============================
if cod_file and other_files:

    # (keeping your COD extraction logic unchanged)
    cod_nominal = 1.2
    tol_mag = 1.41
    key_value = "JFAV21"

    results = []

    for f in other_files:
        sheets = read_all_sheets(f.name, f.read())
        tag = f.name
        is_pdj = tag.upper().startswith("PDJ")
        is_tcm = tag.upper().startswith("TCM")

        for sname, df in sheets.items():
            for (r, _) in find_key_positions(df, key_value):

                nums = row_numbers(df, r)

                nominal_ok = contains_value_eps(nums, cod_nominal, eps)
                tol_ok = contains_pm_pair_eps(nums, tol_mag, eps)

                # 🔹 ADDITION: find actual tolerance if present
                actual_tol = None
                for x in nums:
                    if -x in nums:
                        actual_tol = abs(x)
                        break

                matched=[]
                if nominal_ok:
                    matched.append(str(cod_nominal))
                if tol_ok:
                    matched.append(fmt_pm(tol_mag))

                results.append({
                    "Compared Key": key_value,
                    "File": tag,
                    "Sheet": sname,
                    "Key Row": r+1,

                    "COD Nominal": cod_nominal,
                    "COD Tolerance": fmt_pm(tol_mag),

                    # 🔹 ADDITION: PDJ / TCM values always printed if present
                    "PDJ Nominal Value": cod_nominal if is_pdj and nominal_ok else "",
                    "PDJ Tolerance Value": fmt_pm(actual_tol) if is_pdj and actual_tol is not None else "",

                    "TCM Nominal Value": cod_nominal if is_tcm and nominal_ok else "",
                    "TCM Tolerance Value": fmt_pm(actual_tol) if is_tcm and actual_tol is not None else "",

                    "Actual Nominal Found ?": "Yes" if nominal_ok else "No",
                    "Actual Tolerance Found ?": "Yes" if tol_ok else "No",
                    "OK - Nominal and Tolerance value": ", ".join(matched)
                })

    # ==============================
    # Final Table
    # ==============================
    st.write("### 📊 Results")

    if results:
        df_out = pd.DataFrame(results)
        df_out.insert(0, "SI.No", range(1, len(df_out) + 1))

        st.dataframe(df_out, use_container_width=True)

        def create_colored_excel(df):
            wb = Workbook()
            ws = wb.active
            ws.append(df.columns.tolist())

            green = PatternFill("solid", fgColor="C6F7C6")
            red = PatternFill("solid", fgColor="FFB3B3")

            for i,row in df.iterrows():
                ws.append(row.tolist())
                r=i+2
                ws.cell(r, df.columns.get_loc("Actual Nominal Found ?")+1).fill = green if row["Actual Nominal Found ?"]=="Yes" else red
                ws.cell(r, df.columns.get_loc("Actual Tolerance Found ?")+1).fill = green if row["Actual Tolerance Found ?"]=="Yes" else red

            buf=io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        st.download_button(
            "⬇️ Download Excel",
            create_colored_excel(df_out),
            "cod_comparison_results.xlsx"
        )
    else:
        st.warning("No matches found.")
