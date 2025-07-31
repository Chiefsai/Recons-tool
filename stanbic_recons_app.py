import streamlit as st
import pandas as pd
import os
import re
import tempfile
from datetime import datetime
from PIL import Image
import plotly.express as px

# === Page Config & Style ===
st.set_page_config(page_title="Stanbic Bank Recons Tool", layout="wide")

st.markdown("""
    <style>
        .main {
            background-color: #f4f8ff;
            padding: 20px;
            border-radius: 12px;
        }
        .stApp {
            background-color: #e6f0ff;
        }
        .footer-note {
            position: fixed;
            bottom: 10px;
            right: 25px;
            font-size: 13px;
            color: #003366;
        }
        .center-title {
            display: flex;
            justify-content: center;
            align-items: center;
            font-size: 28px;
            font-weight: bold;
            color: #002060;
        }
    </style>
""", unsafe_allow_html=True)

# === Header with Logo and Title ===
logo_path = "C:/Users/a217202/OneDrive - Standard Bank/Desktop/New folder (2)/icon.jpeg"
col1, col2, col3 = st.columns([1, 4, 1])
with col1:
    st.image(logo_path, width=100)
with col2:
    st.markdown("<div class='center-title'>Stanbic Bank Recon Tool</div>", unsafe_allow_html=True)
with col3:
    st.write("")

# === File Reconciliation Functions ===
def normalize_isin(isin):
    if pd.isna(isin): return ""
    return re.sub(r"\W+", "", str(isin)).strip().upper()

def clean_number(value):
    if pd.isna(value): return None
    value_str = re.sub(r"[^\d.\-]", "", str(value))
    return pd.to_numeric(value_str, errors="coerce")

def reconcile(csv_df, excel_df):
    csv_df["Narration"] = csv_df["Unnamed: 5"].astype(str)
    csv_df["ISIN"] = csv_df["Unnamed: 7"].apply(normalize_isin)
    csv_df["Total Face Value CSD"] = csv_df["Unnamed: 20"].apply(clean_number)

    excel_df["PRODUCT_CODE.ISIN"] = excel_df["PRODUCT_CODE.ISIN"].apply(normalize_isin)
    excel_df["Calypso Position"] = excel_df["Position"].apply(clean_number)

    aligned_records = []
    current_narration = ""
    for _, row in csv_df.iterrows():
        narration, isin, face_value = row["Narration"], row["ISIN"], row["Total Face Value CSD"]
        if pd.notna(narration): current_narration = narration
        if isin:
            aligned_records.append({"ISIN": isin, "Narration": current_narration, "Total Face Value CSD": None})
        elif pd.notna(face_value) and aligned_records:
            aligned_records[-1]["Total Face Value CSD"] = face_value

    stanbic_records = [r for r in aligned_records if "STANBIC BANK GHANA LIMITED" in r["Narration"] and pd.notna(r["Total Face Value CSD"])]

    forward_matched, forward_unmatched = [], []
    for record in stanbic_records:
        isin, csv_value = record["ISIN"], record["Total Face Value CSD"]
        excel_row = excel_df[excel_df["PRODUCT_CODE.ISIN"] == isin]
        if not excel_row.empty:
            excel_value = excel_row.iloc[0]["Calypso Position"]
            if pd.notna(excel_value) and round(csv_value, 2) == round(excel_value, 2):
                forward_matched.append({"PRODUCT_CODE.ISIN": isin, "Total Face Value CSD": csv_value, "Calypso Position": excel_value})
            else:
                forward_unmatched.append({
                    "PRODUCT_CODE.ISIN": isin,
                    "Total Face Value CSD": csv_value,
                    "Calypso Position": excel_value,
                    "Reason": "Value mismatch",
                    "Difference (CSD - Calypso)": round(csv_value - (excel_value if pd.notna(excel_value) else 0), 2)
                })
        else:
            forward_unmatched.append({
                "PRODUCT_CODE.ISIN": isin,
                "Total Face Value CSD": csv_value,
                "Calypso Position": None,
                "Reason": "ISIN not found in Calypso",
                "Difference (CSD - Calypso)": round(csv_value, 2)
            })

    reverse_matched, reverse_unmatched = [], []
    for _, row in excel_df.iterrows():
        isin, excel_value = row["PRODUCT_CODE.ISIN"], row["Calypso Position"]
        match_found = False
        for r in stanbic_records:
            if isin == r["ISIN"]:
                csv_value = r["Total Face Value CSD"]
                if pd.notna(excel_value) and pd.notna(csv_value):
                    if round(excel_value, 2) == round(csv_value, 2):
                        reverse_matched.append({"PRODUCT_CODE.ISIN": isin, "Calypso Position": excel_value, "Total Face Value CSD": csv_value})
                    else:
                        reverse_unmatched.append({
                            "PRODUCT_CODE.ISIN": isin,
                            "Calypso Position": excel_value,
                            "Total Face Value CSD": csv_value,
                            "Reason": "Value mismatch",
                            "Difference (Calypso - CSD)": round(excel_value - csv_value, 2)
                        })
                    match_found = True
                    break
        if not match_found:
            reverse_unmatched.append({
                "PRODUCT_CODE.ISIN": isin,
                "Calypso Position": excel_value,
                "Total Face Value CSD": None,
                "Reason": "ISIN not found in CSD",
                "Difference (Calypso - CSD)": round(excel_value, 2)
            })

    forward_un_df = pd.DataFrame(forward_unmatched)
    reverse_un_df = pd.DataFrame(reverse_unmatched)

    total_csv_all = sum([r["Total Face Value CSD"] for r in stanbic_records if pd.notna(r["Total Face Value CSD"])] or [0])
    total_calypso_all = excel_df["Calypso Position"].sum()

    forward_summary = pd.DataFrame({
        "Insight": [
            "CSD ISINs Evaluated", "Matched Records", "Unmatched Records", "Matching Success Rate (%)",
            "Total Face Value CSD - Unmatched Only", "Calypso Position - Unmatched Only",
            "Difference Total - Unmatched Only", "Total Face Value CSD - All", "Total Calypso Position - All",
            "Difference Total - All"
        ],
        "Value": [
            len(forward_matched) + len(forward_unmatched),
            len(forward_matched),
            len(forward_unmatched),
            round(len(forward_matched) / (len(forward_matched) + len(forward_unmatched)) * 100, 2) if (len(forward_matched) + len(forward_unmatched)) > 0 else 0,
            f"{forward_un_df['Total Face Value CSD'].sum():,.2f}",
            f"{forward_un_df['Calypso Position'].sum():,.2f}",
            f"{forward_un_df['Difference (CSD - Calypso)'].sum():,.2f}",
            f"{total_csv_all:,.2f}",
            f"{total_calypso_all:,.2f}",
            f"{total_csv_all - total_calypso_all:,.2f}"
        ]
    })

    reverse_summary = pd.DataFrame({
        "Insight": [
            "Calypso ISINs Evaluated", "Matched Records", "Unmatched Records", "Matching Success Rate (%)",
            "Total Face Value CSD - Unmatched Only", "Calypso Position - Unmatched Only",
            "Difference Total - Unmatched Only", "Total Face Value CSD - All", "Total Calypso Position - All",
            "Difference Total - All"
        ],
        "Value": [
            len(reverse_matched) + len(reverse_unmatched),
            len(reverse_matched),
            len(reverse_unmatched),
            round(len(reverse_matched) / (len(reverse_matched) + len(reverse_unmatched)) * 100, 2) if (len(reverse_matched) + len(reverse_unmatched)) > 0 else 0,
            f"{reverse_un_df['Total Face Value CSD'].sum():,.2f}",
            f"{reverse_un_df['Calypso Position'].sum():,.2f}",
            f"{reverse_un_df['Difference (Calypso - CSD)'].sum():,.2f}",
            f"{total_csv_all:,.2f}",
            f"{total_calypso_all:,.2f}",
            f"{total_calypso_all - total_csv_all:,.2f}"
        ]
    })

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
            pd.DataFrame(forward_matched).to_excel(writer, sheet_name="CSD-Calypso matched", index=False)
            forward_un_df.to_excel(writer, sheet_name="CSD-Calypso unmatched", index=False)
            pd.DataFrame(reverse_matched).to_excel(writer, sheet_name="Calypso-CSD matched", index=False)
            reverse_un_df.to_excel(writer, sheet_name="Calypso-CSD unmatched", index=False)
            forward_summary.to_excel(writer, sheet_name="Summary Report", startrow=0, index=False)
            reverse_summary.to_excel(writer, sheet_name="Summary Report", startrow=len(forward_summary) + 3, index=False)
        return tmp.name, forward_summary, reverse_summary, forward_un_df, reverse_un_df

# === UI Logic ===
st.markdown("---")
st.subheader("📁 Upload Files for Reconciliation")

csv_file = st.file_uploader("Upload CSD CSV file", type=["csv"])
excel_file = st.file_uploader("Upload Calypso Excel file", type=["xlsx"])

if csv_file and excel_file:
    try:
        csv_df = pd.read_csv(csv_file)
        excel_df = pd.read_excel(excel_file)

        report_path, forward_summary, reverse_summary, forward_un_df, reverse_un_df = reconcile(csv_df, excel_df)

        st.success("✅ Reconciliation complete.")

        st.subheader("📊 CSD vs Calypso Summary")
        st.dataframe(forward_summary)

        st.subheader("📊 Calypso vs CSD Summary")
        st.dataframe(reverse_summary)

        st.subheader("❌ CSD to Calypso Unmatched Records")
        st.dataframe(forward_un_df)

        st.subheader("❌ Calypso to CSD Unmatched Records")
        st.dataframe(reverse_un_df)

        with open(report_path, "rb") as f:
            st.download_button("📥 Download Full Reconciliation Report", f, file_name="Reconciliation_Report.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error during reconciliation: {e}")
else:
    st.info("⬆️ Please upload both the CSD CSV and Calypso Excel files to begin.")
