import streamlit as st
import pandas as pd
import numpy as np
import os, io, re, unicodedata
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ==============================
# Page setup
# ==============================
st.set_page_config(page_title="COD Compare", layout="wide")
st.title("🔎 COD, PDJ, TCM Automatic Validation")

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

def read_all_sheets(name, file_bytes):
    engine = "xlrd" if name.lower().endswith(".xls") else "openpyxl"
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    return {
        s: pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, engine=engine, header=None)
        for s in xls.sheet_names
    }

# ==============================
# COD extraction
# ==============================
def find_codification_value_below(cod_sheets):
    for sname, df in cod_sheets.items():
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                if norm(df.iat[r, c]) == "codification":
                    for rr in range(r + 1, min(df.shape[0], r + 30)):
                        if norm(df.iat[rr, c]):
                            return sname, str(df.iat[rr, c]).strip()
    return None, None

def find_stacked_anchor_vertical(df, words):
    for c in range(df.shape[1]):
        for r in range(df.shape[0]):
            if words[0] in norm(df.iat[r, c]):
                rr = r
                ok = True
                for w in words[1:]:
                    found = False
                    for x in range(rr + 1, min(df.shape[0], rr + 10)):
                        if w in norm(df.iat[x, c]):
                            rr = x
                            found = True
                            break
                    if not found:
                        ok = False
                        break
                if ok:
                    return rr, c
    return None, None

def first_number_below(df, r, c):
    for rr in range(r + 1, min(df.shape[0], r + 6)):
        for cc in range(c, min(df.shape[1], c + 10)):
            s = "" if pd.isna(df.iat[rr, cc]) else str(df.iat[rr, cc])
            m = RE_NUM.search(s)
            if m:
                return abs(to_float(m.group()))
    return None

def two_signed_values_below(df, r, c):
    vals = []
    for rr in range(r + 1, min(df.shape[0], r + 10)):
        s = "" if pd.isna(df.iat[rr, c]) else str(df.iat[rr, c])
        if RE_SIGNED.match(s):
            vals.append(to_float(s))
    for v in vals:
        if -v in vals:
            return abs(v)
    return None

# ==============================
# PDJ / TCM helpers
# ==============================
def row_numbers(df, r):
    nums = []
    for v in df.iloc[r].tolist():
        if v is None or pd.isna(v):
            continue
        for m in RE_NUM.findall(str(v)):
            nums.append(to_float(m))
    return nums

def contains_value(nums, val, eps):
    return any(abs(x - val) <= eps for x in nums if x is not None)

def contains_pm(nums, val, eps):
    return (
        any(abs(x - val) <= eps for x in nums if x is not None) and
        any(abs(x + val) <= eps for x in nums if x is not None)
    )

def fmt_pm(v):
    return f"+/- {v}"

# ==============================
# UI
# ==============================
eps = st.number_input("Numeric tolerance (epsilon)", 0.0, 0.2, 0.02, 0.01)

cod_file = st.file_uploader("Upload COD file", type=["xls", "xlsx"])
other_files = st.file_uploader(
    "Upload PDJ / TCM files", type=["xls", "xlsx"], accept_multiple_files=True
)

# ==============================
# MAIN LOGIC
# ==============================
if cod_file and other_files:

    cod_sheets = read_all_sheets(cod_file.name, cod_file.read())
    cod_sheet, key_value = find_codification_value_below(cod_sheets)

    df_cod = cod_sheets[cod_sheet]

    nr, nc = find_stacked_anchor_vertical(df_cod, ["objectif", "nominal", "jeu"])
    tr, tc = find_stacked_anchor_vertical(df_cod, ["calcul", "disp"])

    cod_nominal = first_number_below(df_cod, nr, nc)
    cod_tol = two_signed_values_below(df_cod, tr, tc) or first_number_below(df_cod, tr, tc)

    rows = []

    for f in other_files:
        sheets = read_all_sheets(f.name, f.read())
        is_pdj = f.name.upper().startswith("PDJ")
        is_tcm = f.name.upper().startswith("TCM")

        for sname, df in sheets.items():
            for r in range(df.shape[0]):
                if key_value in df.iloc[r].astype(str).tolist():

                    nums = row_numbers(df, r)
                    nom_ok = contains_value(nums, cod_nominal, eps)
                    tol_ok = contains_pm(nums, cod_tol, eps)

                    rows.append({
                        "type": "PDJ" if is_pdj else "TCM",
                        "nom_ok": nom_ok,
                        "tol_ok": tol_ok
                    })

    # ==============================
    # TRY (2) AGGREGATION
    # ==============================
    pdj_nom = pdj_tol = tcm_nom = tcm_tol = ""

    nominal_found = False
    tolerance_found = False

    for r in rows:
        if r["type"] == "PDJ":
            if r["nom_ok"]: pdj_nom = cod_nominal
            if r["tol_ok"]: pdj_tol = fmt_pm(cod_tol)
        if r["type"] == "TCM":
            if r["nom_ok"]: tcm_nom = cod_nominal
            if r["tol_ok"]: tcm_tol = fmt_pm(cod_tol)

        if r["nom_ok"]: nominal_found = True
        if r["tol_ok"]: tolerance_found = True

    ok_vals = []
    nok_vals = []

    if nominal_found:
        ok_vals.append(str(cod_nominal))
    else:
        nok_vals.append(str(cod_nominal))

    if tolerance_found:
        ok_vals.append(fmt_pm(cod_tol))
    else:
        nok_vals.append(fmt_pm(cod_tol))

    df_out = pd.DataFrame([{
        "Sl.No": 1,
        "COD Nominal": cod_nominal,
        "COD Tolerance": fmt_pm(cod_tol),
        "PDJ Nominal Value": pdj_nom,
        "PDJ Tolerance Value": pdj_tol,
        "TCM Nominal Value": tcm_nom,
        "TCM Tolerance Value": tcm_tol,
        "Actual Nominal Found ?": "Yes" if nominal_found else "No",
        "Actual Tolerance Found ?": "Yes" if tolerance_found else "No",
        "OK - Nominal and Tolerance value": ", ".join(ok_vals),
        "NOK - value": ", ".join(nok_vals)
    }])

    st.dataframe(df_out, use_container_width=True)

    # ==============================
    # Excel Export
    # ==============================
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(df_out.columns.tolist())

    green = PatternFill("solid", fgColor="C6F7C6")
    red = PatternFill("solid", fgColor="FFB3B3")

    for i, row in df_out.iterrows():
        ws.append(row.tolist())
        r = i + 2
        ws.cell(r, df_out.columns.get_loc("Actual Nominal Found ?") + 1).fill = (
            green if row["Actual Nominal Found ?"] == "Yes" else red
        )
        ws.cell(r, df_out.columns.get_loc("Actual Tolerance Found ?") + 1).fill = (
            green if row["Actual Tolerance Found ?"] == "Yes" else red
        )

    output = io.BytesIO()
    wb.save(output)

    st.download_button(
        "⬇️ Download Excel",
        output.getvalue(),
        "cod_comparison_results_Try2.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
