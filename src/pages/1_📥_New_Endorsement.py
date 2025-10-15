import streamlit as st
import pandas as pd
import io
from msoffcrypto import OfficeFile
from io import BytesIO
from pandas.tseries.offsets import DateOffset
from dotenv import load_dotenv, find_dotenv
import os
from datetime import datetime

# === App configuration ===
st.set_page_config(
    page_title="📆 Monthly Endorsement",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📆 Monthly Endorsement Automation Tool")
st.markdown("Upload your required files below to begin processing.")

# === Load environment variables ===
env_path = find_dotenv()
if not env_path:
    st.error("❌ .env file not found in project directory.")
    st.stop()

load_dotenv(env_path)

DEFAULT_PASSWORD = os.getenv("DEFAULT_PASSWORD")

if not DEFAULT_PASSWORD:
    st.error("❌ Environment variable DEFAULT_PASSWORD not found or empty. Check your .env file.")
    st.stop()


# === Template Header Definition ===
TEMPLATE_HEADERS = [
    "LAST BARCODE DATE", "LAST BARCODE", "PTP DATE", "AGENT", "CLASSIFICATION", "ENDO DATE", "DATE REFERRED",
    "CTL2", "CTL3", "CTL4", "DEBTOR ID", "LAN", "NAME", "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL",
    "MONTHLY AMORTIZATION", "INTEREST", "LPC", "INSURANCE", "PREPAYMENT", "CU PAYMENT",
    "LAST PAYMENT DATE", "PREM AMT", "PROD TYPE", "LPC YTD", "Rate", "Repricing Date", "DPD",
    "LOAN MATURITY", "DUE DATE", "OLDEST DUE DATE", "NEXT DUE DATE", "ADA SHORTAGE", "UNIT", "EMAIL",
    "ALTERNATIVE EMAIL ADDRESS", "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", "BUS_NO_ALS",
    "LANDLINE_NO_ALS", "CO BORROWER", "CO BORROWER MOBILE_ALFES", "CO BORROWER LANDLINE__ALFES",
    "CO BORROWER EMAIL"
]

# === Helper: read Excel/CSV safely ===
def read_file(uploaded_file, password=None, sheet_name=None, header_row=0):
    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()

        if file_ext == "csv":
            return pd.read_csv(uploaded_file)

        elif file_ext in ["xls", "xlsx"]:
            if password:
                decrypted = io.BytesIO()
                office_file = OfficeFile(uploaded_file)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                return pd.read_excel(decrypted, sheet_name=sheet_name, header=header_row, engine="openpyxl")
            else:
                return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, engine="openpyxl")
        else:
            st.error("Unsupported file format.")
            return None

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        return None


# === Step 1: Upload Files ===
st.subheader("Step 1: Upload Required Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**📘 Monthly ENDORSEMENT FILE (Sheet 2)**")
    endorsement_file = st.file_uploader("Upload Monthly ENDORSEMENT FILE", type=["xlsx", "xls", "csv"], key="endorsement")

with col2:
    st.markdown("**📗 TAD FILE (Sheet 1, starts at A4)**")
    tad_file = st.file_uploader("Upload Monthly TAD FILE", type=["xlsx", "xls", "csv"], key="tad")

# === Load files ===
if endorsement_file or tad_file:
    col1, col2 = st.columns(2)

    with col1:
        if endorsement_file:
            df1 = read_file(endorsement_file, password=DEFAULT_PASSWORD, sheet_name="M1 AUTO SPM", header_row=0)
            if df1 is not None:
                st.session_state["endorsement_df"] = df1
                st.success(f"✅ ENDORSEMENT FILE uploaded:\n**{endorsement_file.name}**")

    with col2:
        if tad_file:
            df2 = read_file(tad_file, password=DEFAULT_PASSWORD, sheet_name=0, header_row=3)
            if df2 is not None:
                st.session_state["tad_df"] = df2
                st.success(f"✅ TAD FILE uploaded:\n**{tad_file.name}**")


# === Step 2: Align TAD with Template ===
if "tad_df" in st.session_state:
    st.subheader("⚙️ Step 2: Align TAD Data with Template")

    tad_df = st.session_state["tad_df"]
    tad_df.columns = tad_df.columns.str.strip().str.upper()

    tad_to_template_map = {
        "DATE REFERRED": "DATE REFERRED",
        "CTL2": "CTL2",
        "CTL3": "CTL3",
        "CTL4": "CTL4",
        "LAN": "LAN",
        "PAST DUE": "PAST DUE",
        "PAYOFF AMOUNT": "PAYOFF AMOUNT",
        "PRINCIPAL": "PRINCIPAL",
        "INTEREST": "INTEREST",
        "LPC": "LPC",
        "INSURANCE": "INSURANCE",
        "PREPAYMENT": "PREPAYMENT",
        "CU PAYMENT AMT": "CU PAYMENT",
        "LST BAL CHG DT": "LAST PAYMENT DATE",
        "PREM AMT": "PREM AMT",
        "PROD TYPE": "PROD TYPE",
        "LPC YTD": "LPC YTD",
        "RATE": "Rate",
        "REPRICING DATE": "Repricing Date",
        "DPD": "DPD",
        "ADA SHORTAGE": "ADA SHORTAGE"
    }

    aligned_df = pd.DataFrame(columns=TEMPLATE_HEADERS)
    for tad_col, template_col in tad_to_template_map.items():
        aligned_df[template_col] = tad_df[tad_col] if tad_col in tad_df.columns else ""

    if "LAN" in tad_df.columns:
        aligned_df["LAN"] = tad_df["LAN"]

    st.success("✅ TAD data successfully aligned with template based on unique key LAN!")

    with st.expander("📄 Aligned TAD Preview", expanded=True):
        st.dataframe(aligned_df.head(20))

    csv = aligned_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download Aligned TAD as CSV",
        data=csv,
        file_name="TAD_Aligned_Template.csv",
        mime="text/csv"
    )

    st.session_state["aligned_tad"] = aligned_df


# === Step 3: Merge TAD + ENDORSEMENT ===
st.subheader("Step 3: Merge Files")

if st.button("🚀 Merge Monthly Files"):
    if "endorsement_df" in st.session_state and "aligned_tad" in st.session_state:
        tad_df = st.session_state["aligned_tad"]
        endorsement_df = st.session_state["endorsement_df"]

        # Normalize key for consistency
        tad_df["LAN"] = tad_df["LAN"].astype(str).str.strip()
        endorsement_df.columns = endorsement_df.columns.str.upper().str.strip()

        if "ACCTNUM" in endorsement_df.columns:
            endorsement_df = endorsement_df.rename(columns={"ACCTNUM": "LAN"})

        endorsement_df["LAN"] = endorsement_df["LAN"].astype(str).str.strip()

        # Merge
        merged_df = pd.merge(
            tad_df,
            endorsement_df,
            on="LAN",
            how="left",
            suffixes=("", "_ENDORSEMENT")
        )

        # Map fields
        column_map = {
            "MOAMORT_ALFES": "MONTHLY AMORTIZATION",
            "OLDEST_DUE_DATE": "OLDEST DUE DATE",
            "SHORT_DESCRIPTION": "UNIT",
            "EMAIL_ALS": "EMAIL",
            "EMAIL_ALFES": "ALTERNATIVE EMAIL ADDRESS",
            "MOBILE_NO_ALS": "MOBILE_ALS",
            "MOBILE_ALFES": "MOBILE_ALFES",
            "PRIMARY_NO_ALS": "PRIMARY_NO_ALS",
            "BUS_NO_ALS": "BUS_NO_ALS",
            "LANDLINE_NO_ALFES": "LANDLINE_NO_ALS",
            "COMAKER_NAME_ALFES": "CO BORROWER",
            "COMAKER_MOBILE_ALFES": "CO BORROWER MOBILE_ALFES",
            "COMAKER_LANDLINE_ALFES": "CO BORROWER LANDLINE__ALFES",
            "COMAKER_EMAIL_ALFES": "CO BORROWER EMAIL",
            "NAME_ALS": "NAME"
        }

        for source_col, target_col in column_map.items():
            if source_col in merged_df.columns:
                merged_df[target_col] = merged_df[source_col]

        # Add fixed date fields
        today = datetime.today().strftime("%m/%d/%Y")
        merged_df["ENDO DATE"] = today
        merged_df["DATE REFERRED"] = today
        merged_df["CLASSIFICATION"] = "NEW ENDO"

        # Compute due dates
        try:
            merged_df["OLDEST DUE DATE"] = pd.to_datetime(merged_df["OLDEST DUE DATE"], errors="coerce")
            merged_df["DUE DATE"] = merged_df["OLDEST DUE DATE"].dt.day
            merged_df["NEXT DUE DATE"] = merged_df["OLDEST DUE DATE"] + DateOffset(months=1)

            date_columns = ["OLDEST DUE DATE", "NEXT DUE DATE", "DATE REFERRED", "ENDO DATE", "LAST PAYMENT DATE"]
            for col in date_columns:
                merged_df[col] = pd.to_datetime(merged_df[col], errors="coerce").dt.strftime("%m/%d/%Y")

            merged_df["DUE DATE"] = merged_df["DUE DATE"].fillna("").astype("Int64").astype(str).replace("<NA>", "")
        except Exception as e:
            st.warning(f"⚠️ Could not compute due dates: {e}")

        # Final cleanup
        final_df = pd.DataFrame(columns=TEMPLATE_HEADERS)
        for col in TEMPLATE_HEADERS:
            final_df[col] = merged_df[col] if col in merged_df.columns else ""

        st.success("✅ Successfully merged Monthly TAD + ENDORSEMENT → Active List created!")

        st.dataframe(final_df.head(20))

        # Download Excel
        output = BytesIO()
        final_df.to_excel(output, index=False, engine="openpyxl")
        st.download_button(
            label="💾 Download Monthly Active List",
            data=output.getvalue(),
            file_name="Monthly_Active_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("⚠️ Please upload and align both TAD + Endorsement files before merging.")
