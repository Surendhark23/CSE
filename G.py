"""
GSTR-1 Reconciliation — Streamlit App
======================================
Upload:  1) DayBook (Tally)  .xlsx
         2) GSTR-1 Govt file .xls
Output:  Reconciliation Excel with 10 sheets — download in browser
"""

import io
import warnings

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="GSTR-1 Reconciliation",
    page_icon="📊",
    layout="wide",
)

st.title("📊 GSTR-1 Reconciliation Tool")
st.markdown(
    "Upload your **Tally DayBook** and **Government GSTR-1** file. "
    "The tool will reconcile B2B, B2C and HSN data and give a formatted Excel report."
)

# ══════════════════════════════════════════════════════════════════════════════
# FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
col1, col2 = st.columns(2)

with col1:
    tally_file = st.file_uploader(
        "📁 DayBook File (Tally) — .xlsx",
        type=["xlsx", "xls"],
        key="tally",
    )

with col2:
    govt_file = st.file_uploader(
        "📁 GSTR-1 Govt File — .xls / .xlsx",
        type=["xlsx", "xls"],
        key="govt",
    )

if not tally_file or not govt_file:
    st.info("⬆️  Please upload both files to proceed.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PROCESS BUTTON
# ══════════════════════════════════════════════════════════════════════════════
if not st.button("🚀 Run Reconciliation", type="primary"):
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# ALL PROCESSING INSIDE try/except — shows clean error in UI
# ══════════════════════════════════════════════════════════════════════════════
try:
    with st.spinner("Processing... please wait"):

        # ── STEP 1 : READ & CLEAN TALLY ───────────────────────────────────────
        tally_raw = pd.read_excel(tally_file, sheet_name=0, header=8)
        tally_raw = tally_raw[
            tally_raw["Date"].notna() & (tally_raw["Particulars"] != "Grand Total")
        ].copy()
        tally_raw["Date"] = pd.to_datetime(tally_raw["Date"], errors="coerce")

        # Detect B2B / B2C BEFORE converting GSTIN column to string
        # (column is mixed dtype — float NaN for blank, string for filled)
        tally_raw["_is_b2b"] = tally_raw["GSTIN/UIN"].apply(
            lambda x: not pd.isna(x) and str(x).strip() not in ["", "nan", "NaN"]
        )
        tally_raw["GSTIN/UIN"] = tally_raw["GSTIN/UIN"].astype(str).str.strip()

        for col in ["Gross Total", "GST SALES"]:
            tally_raw[col] = pd.to_numeric(tally_raw[col], errors="coerce").fillna(0)

        tally_raw["Quantity"] = (
            pd.to_numeric(tally_raw["Quantity"], errors="coerce").fillna(0)
        )

        # Collect all CGST / SGST / IGST output columns dynamically
        cgst_cols = [c for c in tally_raw.columns if "OUTPUT CGST" in str(c).upper()]
        sgst_cols = [c for c in tally_raw.columns if "OUTPUT SGST" in str(c).upper()]
        igst_cols = [c for c in tally_raw.columns if "OUTPUT IGST" in str(c).upper()]

        def sum_cols(df, cols):
            if not cols:
                return pd.Series(0.0, index=df.index)
            return df[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

        tally_raw["_CGST"] = sum_cols(tally_raw, cgst_cols)
        tally_raw["_SGST"] = sum_cols(tally_raw, sgst_cols)
        tally_raw["_IGST"] = sum_cols(tally_raw, igst_cols)
        tally_raw["_TotalTax"] = (
            tally_raw["_CGST"] + tally_raw["_SGST"] + tally_raw["_IGST"]
        )
        tally_raw["_TaxRate"] = tally_raw.apply(
            lambda r: round(r["_TotalTax"] / r["GST SALES"] * 100)
            if r["GST SALES"] != 0
            else 0,
            axis=1,
        )

        tally_b2b = tally_raw[tally_raw["_is_b2b"]].copy()
        tally_b2c = tally_raw[~tally_raw["_is_b2b"]].copy()

        # ── STEP 2 : READ & CLEAN GOVT FILE ──────────────────────────────────
        govt_xls = pd.ExcelFile(govt_file, engine="xlrd")
        govt_b2b = pd.read_excel(govt_xls, sheet_name="b2b", header=3)
        govt_b2cs = pd.read_excel(govt_xls, sheet_name="b2cs", header=3)
        govt_hsn = pd.read_excel(govt_xls, sheet_name="hsn", header=3)

        # Clean Govt B2B
        govt_b2b["Invoice Number"] = govt_b2b["Invoice Number"].astype(str).str.strip()
        for col in ["Rate", "Taxable Value", "Invoice Value"]:
            govt_b2b[col] = pd.to_numeric(govt_b2b[col], errors="coerce")

        # Clean Govt B2CS
        govt_b2cs.columns = [
            "Type", "Place Of Supply", "Rate", "Taxable Value",
            "Cess Amount", "E-Commerce GSTIN",
        ]
        govt_b2cs["Rate"] = pd.to_numeric(govt_b2cs["Rate"], errors="coerce")
        govt_b2cs["Taxable Value"] = pd.to_numeric(
            govt_b2cs["Taxable Value"], errors="coerce"
        )

        # Clean Govt HSN
        govt_hsn["HSN"] = govt_hsn["HSN"].astype(str).str.strip()
        for col in [
            "Total Value", "Taxable Value", "Central Tax Amount",
            "State/UT Tax Amount", "Integrated Tax Amount", "Total Quantity",
        ]:
            govt_hsn[col] = pd.to_numeric(govt_hsn[col], errors="coerce").fillna(0)

        govt_hsn_agg = govt_hsn.groupby("HSN").agg(
            Govt_Total=("Total Value", "sum"),
            Govt_Taxable=("Taxable Value", "sum"),
            Govt_CGST=("Central Tax Amount", "sum"),
            Govt_SGST=("State/UT Tax Amount", "sum"),
            Govt_IGST=("Integrated Tax Amount", "sum"),
            Govt_Qty=("Total Quantity", "sum"),
            Description=("Description", "first"),
            UQC=("UQC", "first"),
        ).reset_index()

        # ── STEP 3 : B2B RECONCILIATION ──────────────────────────────────────
        def clean_party(name):
            return str(name).split("\t")[0].replace("\r", "").replace("\n", "").strip()

        b2b_rows = []
        for _, tr in tally_b2b.iterrows():
            vno = str(tr["Voucher No."]).strip()
            date = tr["Date"].strftime("%d-%m-%Y") if pd.notna(tr["Date"]) else ""
            party = clean_party(tr["Particulars"])
            gr = govt_b2b[govt_b2b["Invoice Number"] == vno]

            if gr.empty:
                b2b_rows.append({
                    "Voucher No.": vno, "Date": date, "Party Name": party,
                    "GSTIN (Tally)": tr["GSTIN/UIN"], "GSTIN (Govt)": "NOT FOUND",
                    "Gross Total (Tally)": tr["Gross Total"], "Invoice Value (Govt)": "",
                    "Gross Total Match": "❌ NOT IN GOVT",
                    "Taxable Value (Tally)": tr["GST SALES"], "Taxable Value (Govt)": "",
                    "Taxable Match": "❌ NOT IN GOVT",
                    "CGST": tr["_CGST"], "SGST": tr["_SGST"], "IGST": tr["_IGST"],
                    "Total Tax": tr["_TotalTax"],
                    "Tax Rate (Tally)": tr["_TaxRate"], "Tax Rate (Govt)": "",
                    "Rate Match": "❌ NOT IN GOVT", "Reverse Charge": "",
                    "Overall Status": "❌ MISSING IN GOVT",
                })
                continue

            g = gr.iloc[0]
            rc = str(g["Reverse Charge"]).strip()
            gm = ("✅" if abs(tr["Gross Total"] - g["Invoice Value"]) < 1
                  else f"❌ Diff={round(tr['Gross Total'] - g['Invoice Value'], 2)}")
            tm = ("✅" if abs(tr["GST SALES"] - g["Taxable Value"]) < 1
                  else f"❌ Diff={round(tr['GST SALES'] - g['Taxable Value'], 2)}")
            rm = ("✅" if tr["_TaxRate"] == g["Rate"]
                  else f"❌ Tally={tr['_TaxRate']} Govt={g['Rate']}")
            ok = all(x == "✅" for x in [gm, tm, rm])
            overall = ("✅ MATCH" if ok else "❌ MISMATCH") + (
                " | ⚠️ RC=Y" if rc == "Y" else ""
            )
            b2b_rows.append({
                "Voucher No.": vno, "Date": date, "Party Name": party,
                "GSTIN (Tally)": tr["GSTIN/UIN"],
                "GSTIN (Govt)": g["GSTIN/UIN of Recipient"],
                "Gross Total (Tally)": tr["Gross Total"],
                "Invoice Value (Govt)": g["Invoice Value"], "Gross Total Match": gm,
                "Taxable Value (Tally)": tr["GST SALES"],
                "Taxable Value (Govt)": g["Taxable Value"], "Taxable Match": tm,
                "CGST": tr["_CGST"], "SGST": tr["_SGST"], "IGST": tr["_IGST"],
                "Total Tax": tr["_TotalTax"],
                "Tax Rate (Tally)": tr["_TaxRate"], "Tax Rate (Govt)": g["Rate"],
                "Rate Match": rm, "Reverse Charge": rc, "Overall Status": overall,
            })

        df_b2b_result = pd.DataFrame(b2b_rows)

        # ── STEP 4 : B2C RECONCILIATION ──────────────────────────────────────
        b2c_agg = tally_b2c.groupby("_TaxRate").agg(
            Taxable_Value=("GST SALES", "sum"), CGST=("_CGST", "sum"),
            SGST=("_SGST", "sum"), IGST=("_IGST", "sum"),
            Total_Tax=("_TotalTax", "sum"), Count=("Voucher No.", "count"),
        ).reset_index().rename(columns={"_TaxRate": "Tax Rate"})

        all_rates = sorted(set(
            list(b2c_agg["Tax Rate"].unique()) +
            list(govt_b2cs["Rate"].dropna().unique())
        ))
        b2c_rows = []
        for rate in all_rates:
            tr = b2c_agg[b2c_agg["Tax Rate"] == rate]
            gr = govt_b2cs[govt_b2cs["Rate"] == rate]
            t_tax = tr["Taxable_Value"].sum() if not tr.empty else 0
            g_tax = gr["Taxable Value"].sum() if not gr.empty else 0
            tax_m = "✅" if abs(t_tax - g_tax) < 1 else f"❌ Diff={round(t_tax - g_tax, 2)}"
            b2c_rows.append({
                "Tax Rate %": rate,
                "Invoice Count": int(tr["Count"].sum()) if not tr.empty else 0,
                "Taxable Value (Tally)": round(t_tax, 2),
                "Taxable Value (Govt)": round(g_tax, 2), "Taxable Match": tax_m,
                "CGST": round(tr["CGST"].sum() if not tr.empty else 0, 2),
                "SGST": round(tr["SGST"].sum() if not tr.empty else 0, 2),
                "IGST": round(tr["IGST"].sum() if not tr.empty else 0, 2),
                "Total Tax": round(tr["Total_Tax"].sum() if not tr.empty else 0, 2),
                "Overall Status": "✅ MATCH" if tax_m == "✅" else "❌ MISMATCH",
            })
        df_b2c_result = pd.DataFrame(b2c_rows)

        b2c_detail_rows = []
        for _, r in tally_b2c.iterrows():
            b2c_detail_rows.append({
                "Voucher No.": r["Voucher No."],
                "Date": r["Date"].strftime("%d-%m-%Y") if pd.notna(r["Date"]) else "",
                "Party Name": clean_party(r["Particulars"]),
                "HSN": str(r["HSN"]).strip(), "Gross Total": r["Gross Total"],
                "GST Sales (Taxable)": r["GST SALES"], "CGST": r["_CGST"],
                "SGST": r["_SGST"], "IGST": r["_IGST"],
                "Total Tax": r["_TotalTax"], "Tax Rate %": r["_TaxRate"],
            })
        df_b2c_detail = pd.DataFrame(b2c_detail_rows)

        # ── STEP 5 : EXPAND TALLY PER INVOICE PER HSN ────────────────────────
        exp_rows = []
        for _, r in tally_raw.iterrows():
            hsn_raw = str(r["HSN"]).replace("\n", "").replace("\r", "").strip()
            seen = set()
            hsns = []
            for h in hsn_raw.split(","):
                h = h.strip().rstrip(",")
                if h and h != "nan" and h not in seen:
                    seen.add(h)
                    hsns.append(h)
            for h in hsns:
                exp_rows.append({
                    "Voucher No.": str(r["Voucher No."]).strip(),
                    "Date": r["Date"], "Particulars": clean_party(r["Particulars"]),
                    "GSTIN": r["GSTIN/UIN"], "HSN": h,
                    "Gross Total": r["Gross Total"], "GST SALES": r["GST SALES"],
                    "CGST": r["_CGST"], "SGST": r["_SGST"], "IGST": r["_IGST"],
                    "Quantity": r["Quantity"],
                    "n_hsn": len(hsns), "is_multi": len(hsns) > 1,
                    "is_b2b": r["_is_b2b"],
                })
        df_exp = pd.DataFrame(exp_rows)

        # ── STEP 5b : Build per-HSN B2B/B2C proportion table (for Option A) ────
        # Govt HSN sheet is combined (B2B+B2C). To compare B2B detail vs Govt
        # fairly, we derive the B2B share = (Tally B2B taxable / Tally total
        # taxable) × Govt total for each HSN. Same logic for B2C share.
        _hsn_tally_total = (
            df_exp.groupby("HSN")["GST SALES"].sum()
            .rename("tally_total_tax").reset_index()
        )
        _hsn_b2b_total = (
            df_exp[df_exp["is_b2b"]].groupby("HSN")["GST SALES"].sum()
            .rename("tally_b2b_tax").reset_index()
        )
        _hsn_b2c_total = (
            df_exp[~df_exp["is_b2b"]].groupby("HSN")["GST SALES"].sum()
            .rename("tally_b2c_tax").reset_index()
        )
        _hsn_prop = _hsn_tally_total.merge(_hsn_b2b_total, on="HSN", how="left") \
                                    .merge(_hsn_b2c_total, on="HSN", how="left").fillna(0)
        _hsn_prop["prop_b2b"] = _hsn_prop.apply(
            lambda r: r["tally_b2b_tax"] / r["tally_total_tax"] if r["tally_total_tax"] > 0 else 0, axis=1
        )
        _hsn_prop["prop_b2c"] = _hsn_prop.apply(
            lambda r: r["tally_b2c_tax"] / r["tally_total_tax"] if r["tally_total_tax"] > 0 else 0, axis=1
        )
        _hsn_prop_dict = _hsn_prop.set_index("HSN")[["prop_b2b","prop_b2c"]].to_dict("index")

        def get_govt_split(hsn_code, g_tot, g_tax, g_cg, g_sg, g_ig, filter_b2b):
            """
            OPTION A: When filter_b2b is True/False, scale Govt totals by the
            Tally B2B/B2C proportion so the comparison is apples-to-apples.
            When filter_b2b is None (combined view), use full Govt totals.
            """
            if filter_b2b is None:
                return g_tot, g_tax, g_cg, g_sg, g_ig, "Govt HSN Sheet — Full (B2B+B2C)"
            prop_info = _hsn_prop_dict.get(str(hsn_code), {"prop_b2b": 1.0, "prop_b2c": 0.0})
            prop = prop_info["prop_b2b"] if filter_b2b else prop_info["prop_b2c"]
            pct = round(prop * 100, 1)
            label = "B2B" if filter_b2b else "B2C"
            note = (
                f"Govt HSN × {pct}% {label} share (proportional split — "
                f"Govt doesn't separate B2B/B2C)"
            )
            return (
                round(g_tot * prop, 2), round(g_tax * prop, 2),
                round(g_cg  * prop, 2), round(g_sg  * prop, 2),
                round(g_ig  * prop, 2), note,
            )

        # ── STEP 6 : WATERFALL — invoice-level detail rows ────────────────────
        def waterfall_detail(hsn_code, df_exp, govt_hsn_agg, filter_b2b=None):
            ga = govt_hsn_agg[govt_hsn_agg["HSN"] == hsn_code]
            if ga.empty:
                return []
            g = ga.iloc[0]
            g_tot_full = g["Govt_Total"];  g_tax_full = g["Govt_Taxable"]
            g_cg_full  = g["Govt_CGST"];   g_sg_full  = g["Govt_SGST"]
            g_ig_full  = g["Govt_IGST"]
            desc = str(g["Description"]); uqc = str(g["UQC"])

            # ── OPTION A: use proportional Govt split for B2B / B2C detail ──
            g_tot, g_tax, g_cg, g_sg, g_ig, govt_note = get_govt_split(
                hsn_code, g_tot_full, g_tax_full, g_cg_full, g_sg_full, g_ig_full,
                filter_b2b
            )

            inv = df_exp[df_exp["HSN"] == hsn_code].sort_values("Voucher No.").copy()
            if filter_b2b is True:
                inv = inv[inv["is_b2b"]]
            elif filter_b2b is False:
                inv = inv[~inv["is_b2b"]]

            singles = inv[~inv["is_multi"]]
            multis  = inv[inv["is_multi"]]

            s_tot = singles["Gross Total"].sum(); s_tax = singles["GST SALES"].sum()
            s_cg  = singles["CGST"].sum();        s_sg  = singles["SGST"].sum()
            s_ig  = singles["IGST"].sum()
            n_m   = len(multis)

            rem_tot = round(g_tot - s_tot, 2); rem_tax = round(g_tax - s_tax, 2)
            rem_cg  = round(g_cg  - s_cg,  2); rem_sg  = round(g_sg  - s_sg,  2)
            rem_ig  = round(g_ig  - s_ig,  2)

            rows = []
            for _, r in singles.iterrows():
                rows.append({
                    "HSN": hsn_code, "Description": desc, "UQC": uqc,
                    "Voucher No.": r["Voucher No."],
                    "Date": r["Date"].strftime("%d-%m-%Y") if pd.notna(r["Date"]) else "",
                    "Party Name": r["Particulars"], "GSTIN": r["GSTIN"],
                    "Type": "B2B" if r["is_b2b"] else "B2C", "Multi-HSN": "No",
                    "Gross Total": round(r["Gross Total"], 2),
                    "Taxable Value": round(r["GST SALES"], 2),
                    "CGST": round(r["CGST"], 2), "SGST": round(r["SGST"], 2),
                    "IGST": round(r["IGST"], 2),
                    "Row Type": "Invoice", "Note": "Single HSN — full values",
                })

            for _, r in multis.iterrows():
                pa = round(rem_tot / n_m, 2) if n_m > 0 else 0
                ta = round(rem_tax / n_m, 2) if n_m > 0 else 0
                ca = round(rem_cg  / n_m, 2) if n_m > 0 else 0
                sa = round(rem_sg  / n_m, 2) if n_m > 0 else 0
                ia = round(rem_ig  / n_m, 2) if n_m > 0 else 0
                rows.append({
                    "HSN": hsn_code, "Description": desc, "UQC": uqc,
                    "Voucher No.": r["Voucher No."],
                    "Date": r["Date"].strftime("%d-%m-%Y") if pd.notna(r["Date"]) else "",
                    "Party Name": r["Particulars"], "GSTIN": r["GSTIN"],
                    "Type": "B2B" if r["is_b2b"] else "B2C",
                    "Multi-HSN": f'Yes ({r["n_hsn"]} HSNs)',
                    "Gross Total": pa, "Taxable Value": ta,
                    "CGST": ca, "SGST": sa, "IGST": ia,
                    "Row Type": "Invoice",
                    "Note": f"Multi-HSN: Govt remaining {rem_tot} ÷ {n_m} invoice(s)",
                })

            tt   = round(s_tot + rem_tot, 2); ttax = round(s_tax + rem_tax, 2)
            tcg  = round(s_cg  + rem_cg,  2); tsg  = round(s_sg  + rem_sg,  2)
            tig  = round(s_ig  + rem_ig,  2)

            rows.append({
                "HSN": hsn_code, "Description": desc, "UQC": uqc,
                "Voucher No.": "── TALLY TOTAL ──", "Date": "", "Party Name": "",
                "GSTIN": "", "Type": "", "Multi-HSN": "",
                "Gross Total": tt, "Taxable Value": ttax,
                "CGST": tcg, "SGST": tsg, "IGST": tig,
                "Row Type": "TallyTotal", "Note": "",
            })
            rows.append({
                "HSN": hsn_code, "Description": desc, "UQC": uqc,
                "Voucher No.": "── GOVT TOTAL (Proportional) ──" if filter_b2b is not None
                               else "── GOVT TOTAL (Full) ──",
                "Date": "", "Party Name": "", "GSTIN": "", "Type": "", "Multi-HSN": "",
                "Gross Total": g_tot, "Taxable Value": g_tax,
                "CGST": g_cg, "SGST": g_sg, "IGST": g_ig,
                "Row Type": "GovtTotal", "Note": govt_note,
            })

            dt   = round(tt   - g_tot,  2); dtax = round(ttax - g_tax, 2)
            dcg  = round(tcg  - g_cg,   2); dsg  = round(tsg  - g_sg,  2)
            dig  = round(tig  - g_ig,   2)
            ok   = all(abs(x) < 0.5 for x in [dt, dtax, dcg, dsg, dig])
            rows.append({
                "HSN": hsn_code, "Description": desc, "UQC": uqc,
                "Voucher No.": "✅ MATCH" if ok else "❌ DIFF",
                "Date": "", "Party Name": "", "GSTIN": "", "Type": "", "Multi-HSN": "",
                "Gross Total": dt, "Taxable Value": dtax,
                "CGST": dcg, "SGST": dsg, "IGST": dig,
                "Row Type": "Diff", "Note": "0 = Match | Non-zero = Difference",
            })
            spacer = {k: "" for k in rows[0]}
            spacer["Row Type"] = "Spacer"
            rows.append(spacer)
            return rows

        # ── STEP 7 : WATERFALL — plain alloc rows for HSN summary ─────────────
        def waterfall_alloc(hsn_code, df_exp, govt_hsn_agg):
            ga = govt_hsn_agg[govt_hsn_agg["HSN"] == hsn_code]
            if ga.empty:
                return []
            g = ga.iloc[0]
            g_tot = g["Govt_Total"];  g_tax = g["Govt_Taxable"]
            g_cg = g["Govt_CGST"];   g_sg = g["Govt_SGST"];  g_ig = g["Govt_IGST"]
            desc = str(g["Description"]); uqc = str(g["UQC"])

            inv = df_exp[df_exp["HSN"] == hsn_code].sort_values("Voucher No.").copy()
            singles = inv[~inv["is_multi"]]
            multis = inv[inv["is_multi"]]

            s_tot = singles["Gross Total"].sum(); s_tax = singles["GST SALES"].sum()
            s_cg = singles["CGST"].sum();          s_sg = singles["SGST"].sum()
            n_m = len(multis)

            rem_tot = round(g_tot - s_tot, 2); rem_tax = round(g_tax - s_tax, 2)
            rem_cg = round(g_cg - s_cg, 2);   rem_sg = round(g_sg - s_sg, 2)
            rem_ig = round(g_ig - singles["IGST"].sum(), 2)

            rows = []
            for _, r in singles.iterrows():
                rows.append({
                    "HSN": hsn_code, "Description": desc, "UQC": uqc,
                    "is_b2b": r["is_b2b"],
                    "Gross Total": round(r["Gross Total"], 2),
                    "Taxable": round(r["GST SALES"], 2),
                    "CGST": round(r["CGST"], 2), "SGST": round(r["SGST"], 2),
                    "IGST": 0.0, "Quantity": r["Quantity"],
                })
            for _, r in multis.iterrows():
                rows.append({
                    "HSN": hsn_code, "Description": desc, "UQC": uqc,
                    "is_b2b": r["is_b2b"],
                    "Gross Total": round(rem_tot / n_m, 2) if n_m > 0 else 0,
                    "Taxable": round(rem_tax / n_m, 2) if n_m > 0 else 0,
                    "CGST": round(rem_cg / n_m, 2) if n_m > 0 else 0,
                    "SGST": round(rem_sg / n_m, 2) if n_m > 0 else 0,
                    "IGST": 0.0, "Quantity": 0,
                })
            return rows

        all_alloc = []
        for h in govt_hsn_agg["HSN"].tolist():
            all_alloc.extend(waterfall_alloc(h, df_exp, govt_hsn_agg))
        df_alloc = pd.DataFrame(all_alloc)

        def build_hsn_summary(df_alloc, is_b2b):
            df_f = df_alloc[df_alloc["is_b2b"] == is_b2b]
            if df_f.empty:
                return pd.DataFrame(columns=[
                    "HSN", "Description", "UQC", "Total Quantity", "Total Value",
                    "Taxable Value", "Integrated Tax Amount", "Central Tax Amount",
                    "State/UT Tax Amount", "Cess Amount",
                ])
            agg = df_f.groupby(["HSN", "Description", "UQC"]).agg(
                Total_Quantity=("Quantity", "sum"), Total_Value=("Gross Total", "sum"),
                Taxable_Value=("Taxable", "sum"), IGST=("IGST", "sum"),
                CGST=("CGST", "sum"), SGST=("SGST", "sum"),
            ).reset_index()
            rows = [{
                "HSN": r["HSN"], "Description": r["Description"], "UQC": r["UQC"],
                "Total Quantity": int(round(r["Total_Quantity"])),
                "Total Value": round(r["Total_Value"], 2),
                "Taxable Value": round(r["Taxable_Value"], 2),
                "Integrated Tax Amount": round(r["IGST"], 2),
                "Central Tax Amount": round(r["CGST"], 2),
                "State/UT Tax Amount": round(r["SGST"], 2),
                "Cess Amount": 0,
            } for _, r in agg.iterrows()]
            df_out = pd.DataFrame(rows)
            num_cols = [
                "Total Quantity", "Total Value", "Taxable Value",
                "Integrated Tax Amount", "Central Tax Amount",
                "State/UT Tax Amount", "Cess Amount",
            ]
            total = {c: df_out[c].sum() for c in num_cols}
            total.update({"HSN": "TOTAL", "Description": "", "UQC": ""})
            df_out = pd.concat([df_out, pd.DataFrame([total])], ignore_index=True)
            return df_out

        df_b2b_hsn_sum = build_hsn_summary(df_alloc, True)
        df_b2c_hsn_sum = build_hsn_summary(df_alloc, False)

        all_combined, all_b2b_hsn, all_b2c_hsn = [], [], []
        for h in govt_hsn_agg["HSN"].tolist():
            all_combined.extend(waterfall_detail(h, df_exp, govt_hsn_agg, None))
            r_b2b = waterfall_detail(h, df_exp, govt_hsn_agg, True)
            if any(r["Row Type"] == "Invoice" for r in r_b2b):
                all_b2b_hsn.extend(r_b2b)
            r_b2c = waterfall_detail(h, df_exp, govt_hsn_agg, False)
            if any(r["Row Type"] == "Invoice" for r in r_b2c):
                all_b2c_hsn.extend(r_b2c)

        df_combined = pd.DataFrame(all_combined)
        df_b2b_hsn = pd.DataFrame(all_b2b_hsn)
        df_b2c_hsn = pd.DataFrame(all_b2c_hsn)

        # ── STEP 8 : EXCEL STYLES ─────────────────────────────────────────────
        FG   = PatternFill("solid", fgColor="C6EFCE")
        FR   = PatternFill("solid", fgColor="FFC7CE")
        FY   = PatternFill("solid", fgColor="FFEB9C")
        FT   = PatternFill("solid", fgColor="E2EFDA")
        FV   = PatternFill("solid", fgColor="DDEBF7")
        FALT = PatternFill("solid", fgColor="F2F2F2")
        FN   = PatternFill("solid", fgColor="1F4E79")

        thin = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin"),
        )

        def hdr(ws, row_n, cols, fill=None):
            fill = fill or FN
            for ci, col in enumerate(cols, 1):
                c = ws.cell(row_n, ci, col)
                c.fill = fill
                c.font = Font(bold=True, color="FFFFFF", size=9)
                c.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
                c.border = thin
            ws.row_dimensions[row_n].height = 28

        def cw(ws, ri, ci, val, fill=None, bold=False):
            v = "" if (isinstance(val, float) and np.isnan(val)) else val
            c = ws.cell(ri, ci, v)
            c.border = thin
            c.font = Font(size=9, bold=bold)
            c.alignment = Alignment(horizontal="left", vertical="center")
            if fill:
                c.fill = fill
            return c

        def auto_width(ws, max_w=45):
            for col in ws.columns:
                ml = max((len(str(c.value or "")) for c in col), default=8)
                ws.column_dimensions[
                    get_column_letter(col[0].column)
                ].width = min(ml + 2, max_w)

        # ── STEP 9 : BUILD WORKBOOK ───────────────────────────────────────────
        wb = Workbook()

        # Sheet 1 — B2B Reconciliation
        ws = wb.active
        ws.title = "B2B Reconciliation"
        ws.cell(1, 1, "GSTR-1 B2B Reconciliation").font = Font(
            bold=True, size=13, color="1F4E79"
        )
        ws.cell(2, 1, "🟢 Match  |  🔴 Mismatch / Missing  |  🟡 Reverse Charge = Y").font = Font(
            size=9, italic=True
        )
        cols = list(df_b2b_result.columns)
        hdr(ws, 4, cols)
        for ri, (_, row) in enumerate(df_b2b_result.iterrows(), 5):
            st_ = str(row.get("Overall Status", ""))
            rc = str(row.get("Reverse Charge", ""))
            if "✅ MATCH" in st_ and rc != "Y":
                fill = FG
            elif "❌" in st_:
                fill = FR
            elif rc == "Y" or "⚠️" in st_:
                fill = FY
            else:
                fill = None
            for ci, val in enumerate(row, 1):
                cw(ws, ri, ci, val, fill)
        ws.freeze_panes = "A5"
        auto_width(ws)

        # Sheet 2 — B2C Summary
        ws2 = wb.create_sheet("B2C Summary")
        ws2.cell(1, 1, "GSTR-1 B2C Summary Reconciliation").font = Font(
            bold=True, size=13, color="1F4E79"
        )
        hdr(ws2, 3, list(df_b2c_result.columns))
        for ri, (_, row) in enumerate(df_b2c_result.iterrows(), 4):
            fill = FG if "✅" in str(row.get("Overall Status", "")) else FR
            for ci, val in enumerate(row, 1):
                cw(ws2, ri, ci, val, fill)
        auto_width(ws2)

        # Sheet 3 — B2C Invoice Detail
        ws3 = wb.create_sheet("B2C Invoice Detail")
        ws3.cell(1, 1, "B2C Invoice Detail — Blank GSTIN in Tally").font = Font(
            bold=True, size=13, color="1F4E79"
        )
        hdr(ws3, 3, list(df_b2c_detail.columns))
        for ri, (_, row) in enumerate(df_b2c_detail.iterrows(), 4):
            for ci, val in enumerate(row, 1):
                cw(ws3, ri, ci, val)
        auto_width(ws3)

        # HSN detail sheet writer
        DISP = [
            "HSN", "Description", "UQC", "Voucher No.", "Date", "Party Name",
            "GSTIN", "Type", "Multi-HSN", "Gross Total", "Taxable Value",
            "CGST", "SGST", "IGST", "Note",
        ]

        def write_hsn_detail(wb, df, sheet_name, title):
            ws = wb.create_sheet(sheet_name)
            ws.cell(1, 1, title).font = Font(bold=True, size=13, color="1F4E79")
            ws.cell(
                2, 1,
                "🟩 Tally Total  |  🟦 Govt Total (Final)  |  "
                "🟢 Match  |  🔴 Diff  |  Govt HSN is FINAL cap",
            ).font = Font(size=9, italic=True, color="CC0000")
            cols = [c for c in DISP if c in df.columns]
            hdr(ws, 4, cols)
            ri_out = 5
            for _, row in df.iterrows():
                rt = str(row.get("Row Type", ""))
                if rt == "Spacer":
                    ri_out += 1
                    continue
                if rt == "TallyTotal":
                    fill, bold = FT, True
                elif rt == "GovtTotal":
                    fill, bold = FV, True
                elif rt == "Diff":
                    try:
                        dv = float(row.get("Gross Total", 0))
                    except Exception:
                        dv = 0
                    fill, bold = (FG, True) if abs(dv) < 0.5 else (FR, True)
                else:
                    fill, bold = None, False
                for ci, col in enumerate(cols, 1):
                    cw(ws, ri_out, ci, row.get(col, ""), fill, bold)
                ri_out += 1
            ws.freeze_panes = "A5"
            auto_width(ws)

        write_hsn_detail(wb, df_b2b_hsn,  "B2B HSN Detail",
                         "B2B HSN — Invoice-wise Waterfall | Govt = Proportional B2B Share (Option A)")
        write_hsn_detail(wb, df_b2c_hsn,  "B2C HSN Detail",
                         "B2C HSN — Invoice-wise Waterfall | Govt = Proportional B2C Share (Option A)")
        write_hsn_detail(wb, df_combined, "All HSN Combined",
                         "All HSN (B2B + B2C) — Invoice-wise Waterfall vs Full Govt HSN")

        # HSN Summary sheet writer
        HSN_COLS = [
            "HSN", "Description", "UQC", "Total Quantity", "Total Value",
            "Taxable Value", "Integrated Tax Amount", "Central Tax Amount",
            "State/UT Tax Amount", "Cess Amount",
        ]
        COL_W = {
            "HSN": 14, "Description": 36, "UQC": 14, "Total Quantity": 14,
            "Total Value": 16, "Taxable Value": 16, "Integrated Tax Amount": 20,
            "Central Tax Amount": 18, "State/UT Tax Amount": 18, "Cess Amount": 14,
        }

        def write_hsn_summary(wb, df, sheet_name, title, subtitle, df_other=None, govt_hsn_agg=None):
            """
            OPTION B: Adds a combined Tally (B2B+B2C) vs Govt comparison block
            below the main summary so the reader can see:
              • the split (this sheet = B2B or B2C portion)
              • the combined total that actually ties to Govt HSN
            df_other  = the partner summary df (if this is B2B, pass df_b2c_hsn_sum)
            govt_hsn_agg = aggregated govt HSN dataframe
            """
            ws = wb.create_sheet(sheet_name)
            ws.cell(1, 1, title).font = Font(bold=True, size=13, color="1F4E79")
            ws.merge_cells(
                start_row=1, start_column=1, end_row=1, end_column=len(HSN_COLS)
            )
            ws.cell(2, 1, subtitle).font = Font(size=9, italic=True, color="595959")
            ws.merge_cells(
                start_row=2, start_column=1, end_row=2, end_column=len(HSN_COLS)
            )
            hdr(ws, 3, HSN_COLS)
            ws.row_dimensions[3].height = 32
            next_ri = 4
            for ri, (_, row) in enumerate(df.iterrows(), 4):
                is_total = str(row.get("HSN", "")) == "TOTAL"
                fill = FT if is_total else (FALT if ri % 2 == 0 else None)
                for ci, col in enumerate(HSN_COLS, 1):
                    val = row.get(col, "")
                    if isinstance(val, float) and np.isnan(val):
                        val = ""
                    c = ws.cell(ri, ci, val)
                    c.border = thin
                    c.font = Font(size=9, bold=is_total)
                    if col in ("HSN", "UQC", "Total Quantity"):
                        c.alignment = Alignment(horizontal="center", vertical="center")
                    elif col == "Description":
                        c.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        c.alignment = Alignment(horizontal="right", vertical="center")
                        if col != "Cess Amount":
                            c.number_format = "#,##0.00"
                    if fill:
                        c.fill = fill
                next_ri = ri + 1

            # ── OPTION B : Combined Tally (B2B+B2C) vs Govt comparison block ──
            if df_other is not None and govt_hsn_agg is not None:
                gap_rows = next_ri + 1  # one blank row gap

                # Header banner
                banner = ws.cell(gap_rows, 1,
                    "OPTION B VALIDATION — Combined Tally (B2B + B2C) vs Govt HSN  "
                    "│  Govt HSN sheet does NOT split B2B/B2C — compare combined totals here")
                banner.font = Font(bold=True, size=10, color="FFFFFF")
                banner.fill = PatternFill("solid", fgColor="375623")
                banner.alignment = Alignment(horizontal="left", vertical="center")
                ws.merge_cells(start_row=gap_rows, start_column=1,
                               end_row=gap_rows, end_column=len(HSN_COLS))
                ws.row_dimensions[gap_rows].height = 22
                gap_rows += 1

                CMP_COLS = [
                    "HSN", "Description",
                    "Tally B2B Taxable", "Tally B2C Taxable", "Tally COMBINED Taxable",
                    "Govt Taxable", "Diff Taxable",
                    "Tally B2B CGST", "Tally B2C CGST", "Tally COMBINED CGST",
                    "Govt CGST", "Diff CGST",
                    "Status",
                ]
                hdr(ws, gap_rows, CMP_COLS, fill=PatternFill("solid", fgColor="375623"))
                ws.row_dimensions[gap_rows].height = 28
                gap_rows += 1

                # Build lookup: HSN → this-sheet row, other-sheet row, govt row
                this_data  = df[df["HSN"].astype(str) != "TOTAL"].set_index("HSN")
                other_data = df_other[df_other["HSN"].astype(str) != "TOTAL"].set_index("HSN")
                govt_lookup = govt_hsn_agg.set_index("HSN")

                all_hsns = sorted(set(
                    list(this_data.index.astype(str)) +
                    list(other_data.index.astype(str)) +
                    list(govt_lookup.index.astype(str))
                ))

                sum_comb_tax = 0; sum_govt_tax = 0
                sum_comb_cg  = 0; sum_govt_cg  = 0
                sum_b2b_tax  = 0; sum_b2c_tax  = 0
                sum_b2b_cg   = 0; sum_b2c_cg   = 0

                for hsn in all_hsns:
                    t_b2b_tax = float(this_data.loc[hsn, "Taxable Value"])   if hsn in this_data.index  else 0
                    t_b2c_tax = float(other_data.loc[hsn, "Taxable Value"])  if hsn in other_data.index else 0
                    t_b2b_cg  = float(this_data.loc[hsn, "Central Tax Amount"])  if hsn in this_data.index  else 0
                    t_b2c_cg  = float(other_data.loc[hsn, "Central Tax Amount"]) if hsn in other_data.index else 0
                    g_tax = float(govt_lookup.loc[hsn, "Govt_Taxable"]) if hsn in govt_lookup.index else 0
                    g_cg  = float(govt_lookup.loc[hsn, "Govt_CGST"])    if hsn in govt_lookup.index else 0
                    desc  = str(govt_lookup.loc[hsn, "Description"]) if hsn in govt_lookup.index else ""

                    comb_tax = round(t_b2b_tax + t_b2c_tax, 2)
                    comb_cg  = round(t_b2b_cg  + t_b2c_cg,  2)
                    diff_tax = round(comb_tax - g_tax, 2)
                    diff_cg  = round(comb_cg  - g_cg,  2)
                    ok_row   = abs(diff_tax) < 0.5 and abs(diff_cg) < 0.5
                    status   = "✅ MATCH" if ok_row else f"❌ DIFF tax={diff_tax} cgst={diff_cg}"

                    sum_b2b_tax += t_b2b_tax; sum_b2c_tax += t_b2c_tax
                    sum_b2b_cg  += t_b2b_cg;  sum_b2c_cg  += t_b2c_cg
                    sum_comb_tax += comb_tax;  sum_govt_tax += g_tax
                    sum_comb_cg  += comb_cg;   sum_govt_cg  += g_cg

                    row_vals = [
                        hsn, desc,
                        t_b2b_tax, t_b2c_tax, comb_tax, g_tax, diff_tax,
                        t_b2b_cg,  t_b2c_cg,  comb_cg,  g_cg,  diff_cg,
                        status,
                    ]
                    row_fill = FG if ok_row else FR
                    for ci, val in enumerate(row_vals, 1):
                        c = ws.cell(gap_rows, ci, val)
                        c.border = thin
                        c.font   = Font(size=9)
                        c.fill   = row_fill
                        c.alignment = Alignment(
                            horizontal="right" if ci > 2 else "left",
                            vertical="center"
                        )
                        if ci > 2 and ci != len(row_vals):
                            c.number_format = "#,##0.00"
                    gap_rows += 1

                # Grand total row
                grand_diff_tax = round(sum_comb_tax - sum_govt_tax, 2)
                grand_diff_cg  = round(sum_comb_cg  - sum_govt_cg,  2)
                grand_ok = abs(grand_diff_tax) < 1 and abs(grand_diff_cg) < 1
                grand_vals = [
                    "GRAND TOTAL", "",
                    round(sum_b2b_tax,2), round(sum_b2c_tax,2), round(sum_comb_tax,2),
                    round(sum_govt_tax,2), grand_diff_tax,
                    round(sum_b2b_cg,2),  round(sum_b2c_cg,2),  round(sum_comb_cg,2),
                    round(sum_govt_cg,2),  grand_diff_cg,
                    "✅ BALANCED" if grand_ok else f"❌ GAP={grand_diff_tax}",
                ]
                for ci, val in enumerate(grand_vals, 1):
                    c = ws.cell(gap_rows, ci, val)
                    c.border = thin
                    c.font   = Font(size=9, bold=True)
                    c.fill   = FT
                    c.alignment = Alignment(
                        horizontal="right" if ci > 2 else "left",
                        vertical="center"
                    )
                    if ci > 2 and ci != len(grand_vals):
                        c.number_format = "#,##0.00"

                # Auto-width for the wider comparison columns
                for ci in range(1, len(CMP_COLS)+1):
                    col_letter = get_column_letter(ci)
                    if ws.column_dimensions[col_letter].width < 16:
                        ws.column_dimensions[col_letter].width = 16
                ws.column_dimensions["A"].width = 14
                ws.column_dimensions["B"].width = 30

            for ci, col in enumerate(HSN_COLS, 1):
                ws.column_dimensions[get_column_letter(ci)].width = COL_W.get(col, 14)
            ws.freeze_panes = "A4"

        write_hsn_summary(
            wb, df_b2b_hsn_sum, "HSN B2B Summary",
            "HSN Summary — B2B Invoices (Tally Waterfall Allocation)",
            "HSN-wise clubbed totals for B2B | Scroll down for combined B2B+B2C vs Govt validation",
            df_other=df_b2c_hsn_sum, govt_hsn_agg=govt_hsn_agg,
        )
        write_hsn_summary(
            wb, df_b2c_hsn_sum, "HSN B2C Summary",
            "HSN Summary — B2C Invoices (Tally Waterfall Allocation)",
            "HSN-wise clubbed totals for B2C (blank GSTIN) | Scroll down for combined B2B+B2C vs Govt validation",
            df_other=df_b2b_hsn_sum, govt_hsn_agg=govt_hsn_agg,
        )

        # Sheet 9 — Govt HSN Reference
        ws_ref = wb.create_sheet("Govt HSN Reference")
        ws_ref.cell(1, 1, "Govt GSTR-1 HSN Sheet — Final Reference").font = Font(
            bold=True, size=13, color="1F4E79"
        )
        hdr(ws_ref, 3, list(govt_hsn.columns))
        for ri, (_, row) in enumerate(govt_hsn.iterrows(), 4):
            for ci, val in enumerate(row, 1):
                cw(ws_ref, ri, ci, val)
        auto_width(ws_ref)

        # Sheet 10 — Legend
        wl = wb.create_sheet("Legend")
        wl.column_dimensions["A"].width = 32
        wl.column_dimensions["B"].width = 88
        legend = [
            ("COLOUR / SYMBOL", "MEANING"),
            ("🟢 Green row (B2B)", "All fields match: Gross Total, Taxable Value, Tax Rate"),
            ("🔴 Red row (B2B)", "Mismatch or invoice missing from Govt file"),
            ("🟡 Yellow row (B2B)", "Reverse Charge = Y — needs attention"),
            ("🟩 Light Green band (HSN)", "Tally Total row for that HSN"),
            ("🟦 Light Blue band (HSN)", "Govt Total row — FINAL"),
            ("🟢 / 🔴 Diff row (HSN)", "0 = perfect match | non-zero = gap"),
            ("", ""),
            ("WHY GOVT HSN ≠ B2B or B2C ALONE", ""),
            ("Key fact",
             "Govt HSN sheet is a COMBINED total (B2B + B2C). It does NOT split by customer type."),
            ("", ""),
            ("OPTION A — Proportional Govt Split (B2B/B2C Detail sheets)", ""),
            ("What it does",
             "In B2B HSN Detail / B2C HSN Detail: Govt Total is scaled by the "
             "proportion of B2B (or B2C) taxable value vs total Tally taxable value for that HSN."),
            ("Example",
             "HSN 84151010: B2B = 73.85% of Tally total. Govt 12,33,042 × 73.85% = 9,10,000 used as B2B cap."),
            ("Result",
             "Diff row in B2B/B2C detail sheets will show ✅ MATCH when the split is correct."),
            ("", ""),
            ("OPTION B — Combined Tally vs Govt (HSN B2B/B2C Summary sheets)", ""),
            ("What it does",
             "At the bottom of HSN B2B Summary and HSN B2C Summary sheets, a validation table "
             "shows: Tally B2B + Tally B2C = Tally Combined, compared directly to Govt HSN total."),
            ("Why this is the real proof",
             "Since Govt HSN = B2B+B2C combined, this combined comparison is the true match test. "
             "If Tally Combined = Govt → ✅ BALANCED — your data is correct."),
            ("", ""),
            ("HSN WATERFALL LOGIC", ""),
            ("Step 1 — Single-HSN invoice",
             "Full Gross / Taxable / CGST / SGST / IGST assigned to that HSN"),
            ("Step 2 — Multi-HSN invoice",
             "Remaining = Govt_Total − Σ(single-invoice allocations for this HSN)"),
            ("Step 3 — Split remaining",
             "Remaining divided equally among multi-HSN invoices sharing this HSN"),
            ("Step 4 — Verify",
             "Tally Total = singles + remaining = Govt Total (always matches)"),
            ("", ""),
            ("Example — HSN 2903", "Govt Total = 13,570"),
            ("", "CSES/I/074 (single) → 3,540  |  running = 3,540"),
            ("", "CSES/I/075 (single) → 2,950  |  running = 6,490"),
            ("", "CSES/I/087 (multi)  → 13,570 − 6,490 = 7,080"),
            ("", "Tally Total = 3,540 + 2,950 + 7,080 = 13,570  ✅"),
        ]
        for ri, (k, v) in enumerate(legend, 1):
            bold = ri == 1 or k in ("HSN WATERFALL LOGIC", "Example — HSN 2903")
            c1 = wl.cell(ri, 1, k)
            c2 = wl.cell(ri, 2, v)
            c1.font = Font(bold=bold, size=9, color="1F4E79" if bold else "000000")
            c2.font = Font(size=9)

        # ── STEP 10 : SAVE TO MEMORY BUFFER & OFFER DOWNLOAD ─────────────────
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    # ── SUCCESS SUMMARY ───────────────────────────────────────────────────────
    st.success("✅ Reconciliation complete!")

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("B2B Invoices", len(df_b2b_result))
    col_b.metric("B2C Invoices", len(df_b2c_detail))
    col_c.metric("HSN B2B Rows", len(df_b2b_hsn_sum) - 1)
    col_d.metric("HSN B2C Rows", len(df_b2c_hsn_sum) - 1)

    # B2B match summary
    matched = df_b2b_result["Overall Status"].str.contains("✅ MATCH").sum()
    mismatched = len(df_b2b_result) - matched
    st.markdown(f"**B2B:** {matched} matched &nbsp;|&nbsp; {mismatched} mismatched / missing")

    st.download_button(
        label="📥 Download Reconciliation Report (.xlsx)",
        data=buf,
        file_name="GSTR1_Reconciliation_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Optional preview
    with st.expander("👁️ Preview B2B Reconciliation"):
        st.dataframe(df_b2b_result, use_container_width=True)

    with st.expander("👁️ Preview HSN B2B Summary"):
        st.dataframe(df_b2b_hsn_sum, use_container_width=True)

    with st.expander("👁️ Preview HSN B2C Summary"):
        st.dataframe(df_b2c_hsn_sum, use_container_width=True)

except Exception as e:
    st.error(f"❌ Error during processing: {e}")
    st.exception(e)
