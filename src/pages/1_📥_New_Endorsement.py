import streamlit as st
import pandas as pd
import io
from msoffcrypto import OfficeFile
from io import BytesIO
from pandas.tseries.offsets import DateOffset

# === App configuration ===
st.set_page_config(
    page_title="Work Automation App",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📂 Work Automation App – File Processing Tool")
st.markdown("Upload your files below to begin automation.")


# === Default password for Excel ===
DEFAULT_PASSWORD = "BPI_SPM2025"  # Replace with actual Excel sheet password

# === Template Header Definition ===
TEMPLATE_HEADERS = [
    "LAST BARCODE", "DATE", "PTP DATE", "AGENT", "CLASSIFICATION", "ENDO DATE", "DATE REFERRED",
    "CTL2", "CTL3", "CTL4", "DEBTOR ID", "LAN", "NAME", "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL",
    "MONTHLY AMORTIZATION", "INTEREST", "LPC", "INSURANCE", "PREPAYMENT", "CU PAYMENT",
    "LAST PAYMENT DATE", "PREM AMT", "PROD TYPE", "LPC YTD", "Rate", "Repricing Date", "DPD",
    "LOAN MATURITY", "DUE DATE", "OLDEST DUE DATE", "NEXT DUE DATE", "ADA SHORTAGE", "UNIT", "EMAIL",
    "ALTERNATIVE EMAIL ADDRESS", "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", "BUS_NO_ALS",
    "LANDLINE_NO_ALS", "CO BORROWER", "CO BORROWER MOBILE_ALFES", "CO BORROWER LANDLINE__ALFES",
    "CO BORROWER EMAIL"
]

# === Columns present in TAD ===
TAD_HEADERS = [
    "DATE REFERRED", "CTL2", "CTL3", "CTL4", "LAN", "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL",
    "INTEREST", "LPC", "INSURANCE", "PREPAYMENT", "CU PAYMENT AMT", "LST BAL CHG DT", "PREM AMT",
    "PROD TYPE", "LPC YTD", "Rate", "Repricing Date", "DPD", "ADA SHORTAGE"
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


# === Step 1: Upload Files (side-by-side layout) ===
st.subheader("Step 1: Upload Required Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**📘 ENDORSEMENT FILE (Sheet 2)**")
    endorsement_file = st.file_uploader("Upload ENDORSEMENT FILE", type=["xlsx", "xls", "csv"], key="endorsement")

with col2:
    st.markdown("**📗 TAD FILE (Sheet 1, starts at A4)**")
    tad_file = st.file_uploader("Upload TAD FILE", type=["xlsx", "xls", "csv"], key="tad")

# === Load files ===
if endorsement_file:
    df1 = read_file(endorsement_file, password=DEFAULT_PASSWORD, sheet_name="M1 AUTO SPM", header_row=0)
    if df1 is not None:
        st.session_state["endorsement_df"] = df1
        st.success(f"✅ ENDORSEMENT FILE uploaded: {endorsement_file.name} (Sheet 2)")

if tad_file:
    df2 = read_file(tad_file, password=DEFAULT_PASSWORD, sheet_name="SPM Endo", header_row=3)
    if df2 is not None:
        st.session_state["tad_df"] = df2
        st.success(f"✅ TAD FILE uploaded: {tad_file.name} (Sheet 1, starts at A4)")

# === Step 2: Align TAD with Template ===
if "tad_df" in st.session_state:
    st.subheader("⚙️ Step 2: Align TAD Data with Template")

    tad_df = st.session_state["tad_df"]

    # Normalize TAD columns
    tad_df.columns = tad_df.columns.str.strip().str.upper()

    # --- Define mapping between TAD headers and Template headers ---
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
        "CU PAYMENT AMT": "CU PAYMENT",              # ✅ Custom mapping
        "LST BAL CHG DT": "LAST PAYMENT DATE",       # ✅ Custom mapping
        "PREM AMT": "PREM AMT",
        "PROD TYPE": "PROD TYPE",
        "LPC YTD": "LPC YTD",
        "RATE": "Rate",
        "REPRICING DATE": "Repricing Date",
        "DPD": "DPD",
        "ADA SHORTAGE": "ADA SHORTAGE"
    }

    # --- Build aligned DataFrame ---
    aligned_df = pd.DataFrame(columns=TEMPLATE_HEADERS)

    for tad_col, template_col in tad_to_template_map.items():
        if tad_col in tad_df.columns:
            aligned_df[template_col] = tad_df[tad_col]
        else:
            aligned_df[template_col] = ""

    # --- Preserve the unique key (LAN) ---
    if "LAN" in tad_df.columns:
        aligned_df["LAN"] = tad_df["LAN"]

    # --- Display result ---
    st.success("✅ TAD data successfully aligned with template based on unique key LAN!")

    with st.expander("📄 Aligned TAD Preview", expanded=True):
        st.dataframe(aligned_df.head(20))

    # --- Option to download aligned file ---
    csv = aligned_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download Aligned TAD as CSV",
        data=csv,
        file_name="TAD_Aligned_Template.csv",
        mime="text/csv"
    )

    # Store aligned dataframe for merging step
    st.session_state["aligned_tad"] = aligned_df


# === Step 2.5: Upload Masterlist (for classification tracking) ===
st.subheader("Step 2.5: Upload Masterlist (for classification tracking)")
masterlist_file = st.file_uploader("📘 Upload MASTERLIST Excel", type=["xls", "xlsx", "csv"], key="masterlist")

if masterlist_file:
    masterlist_df = read_file(masterlist_file, password=DEFAULT_PASSWORD)
    if masterlist_df is not None:
        st.session_state["masterlist_df"] = masterlist_df
        st.success(f"✅ Masterlist uploaded: {masterlist_file.name}")

    
# === Step 3: Merge Files ===
st.subheader("Step 3: Merge Files")

if st.button("🚀 Merge with ENDORSEMENT FILE"):
    if "endorsement_df" in st.session_state and "aligned_tad" in st.session_state:
        tad_df = st.session_state["aligned_tad"]
        endorsement_df = st.session_state["endorsement_df"]

        # --- Template headers (final output order) ---
        TEMPLATE_HEADERS = [
            "LAST BARCODE","DATE","PTP DATE","AGENT","CLASSIFICATION","ENDO DATE","DATE REFERRED",
            "CTL2","CTL3","CTL4","DEBTOR ID","LAN","NAME","PAST DUE","PAYOFF AMOUNT","PRINCIPAL",
            "MONTHLY AMORTIZATION","INTEREST","LPC","INSURANCE","PREPAYMENT","CU PAYMENT",
            "LAST PAYMENT DATE","PREM AMT","PROD TYPE","LPC YTD","Rate","Repricing Date","DPD",
            "LOAN MATURITY","DUE DATE","OLDEST DUE DATE","NEXT DUE DATE","ADA SHORTAGE","UNIT",
            "EMAIL","ALTERNATIVE EMAIL ADDRESS","MOBILE_ALS","MOBILE_ALFES","PRIMARY_NO_ALS",
            "BUS_NO_ALS","LANDLINE_NO_ALS","CO BORROWER","CO BORROWER MOBILE_ALFES",
            "CO BORROWER LANDLINE__ALFES","CO BORROWER EMAIL"
        ]

        # --- Rename ACCTNUM -> LAN for merge consistency ---
        if "ACCTNUM" in endorsement_df.columns:
            endorsement_df = endorsement_df.rename(columns={"ACCTNUM": "LAN"})

        # --- Merge on LAN key ---
        merged_df = pd.merge(
            tad_df,
            endorsement_df,
            on="LAN",
            how="left",
            suffixes=("", "_ENDORSEMENT")
        )

        # --- Map ENDORSEMENT fields to TEMPLATE headers ---
        column_map = {
            "CTL4": "CTL4",
            "MOAMORT_ALFES": "MONTHLY AMORTIZATION",
            "Oldest_Due_date": "OLDEST DUE DATE",
            "DPD": "DPD",
            "short_description": "UNIT",
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

        # === Classification Logic Integration ===
        from datetime import datetime
        today = datetime.today().strftime("%Y-%m-%d")

        if "masterlist_df" in st.session_state:
            masterlist_df = st.session_state["masterlist_df"]

            if "LAN" in masterlist_df.columns:
                # --- Classification (NEW ENDO or REENDO) ---
                merged_df["CLASSIFICATION"] = merged_df["LAN"].apply(
                    lambda lan: "REENDO" if lan in masterlist_df["LAN"].values else "NEW ENDO"
                )

                # --- DATE REFERRED Logic ---
                def get_date_referred(row):
                    if row["CLASSIFICATION"] == "NEW ENDO":
                        return today
                    else:
                        prev_date = masterlist_df.loc[
                            masterlist_df["LAN"] == row["LAN"], "DATE REFERRED"
                        ]
                        return prev_date.values[0] if not prev_date.empty else today

                merged_df["DATE REFERRED"] = merged_df.apply(get_date_referred, axis=1)

            else:
                st.warning("⚠️ Masterlist file missing 'LAN' column. Classification skipped.")
                merged_df["CLASSIFICATION"] = "NEW ENDO"
                merged_df["DATE REFERRED"] = today

        else:
            st.info("ℹ️ No masterlist uploaded. All accounts marked as NEW ENDO.")
            merged_df["CLASSIFICATION"] = "NEW ENDO"
            merged_df["DATE REFERRED"] = today

        # --- ENDO DATE always today ---
        merged_df["ENDO DATE"] = today


        # --- Derived columns ---


    try:
        # 1️⃣ Parse OLDEST DUE DATE to datetime safely
        merged_df["OLDEST DUE DATE"] = pd.to_datetime(
            merged_df["OLDEST DUE DATE"], errors="coerce"
        )

        # 2️⃣ DUE DATE = DAY(OLDEST DUE DATE)
        merged_df["DUE DATE"] = merged_df["OLDEST DUE DATE"].dt.day

        # 3️⃣ NEXT DUE DATE = EDATE(OLDEST DUE DATE, 1)
        merged_df["NEXT DUE DATE"] = merged_df["OLDEST DUE DATE"] + DateOffset(months=1)

        # 4️⃣ Format all date-related columns to short date (mm/dd/yyyy)
        date_columns = [
            "OLDEST DUE DATE", "NEXT DUE DATE",
            "DATE REFERRED", "ENDO DATE", "LAST PAYMENT DATE"
        ]

        for col in date_columns:
            if col in merged_df.columns:
                merged_df[col] = pd.to_datetime(merged_df[col], errors="coerce").dt.strftime("%m/%d/%Y")

        # 5️⃣ Ensure DUE DATE is integer only (no decimals, no NaN)
        merged_df["DUE DATE"] = (
            merged_df["DUE DATE"]
            .astype("Int64")
            .astype(str)
            .replace("<NA>", "")
        )

    except Exception as e:
        st.warning(f"⚠️ Could not compute due dates automatically: {e}")


        # --- Final cleanup and alignment ---
        final_df = pd.DataFrame(columns=TEMPLATE_HEADERS)
        for col in TEMPLATE_HEADERS:
            if col in merged_df.columns:
                final_df[col] = merged_df[col]
            else:
                final_df[col] = ""

        st.success("✅ Successfully merged TAD + ENDORSEMENT using LAN ↔ ACCTNUM key!")

        # --- Dashboard Summary ---
        st.markdown("### 📊 Dashboard Summary")

        merged_df["PAST DUE"] = pd.to_numeric(merged_df["PAST DUE"], errors="coerce").fillna(0)
        total_endo = endorsement_df.shape[0]
        with_past_due = (merged_df["PAST DUE"] > 0).sum()
        pout = (merged_df["PAST DUE"] == 0).sum()
        active = with_past_due  # active = accounts with past due > 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🆕 New ENDO", f"{total_endo:,}")
        col2.metric("📅 With Past Dues", f"{with_past_due:,}")
        col3.metric("💤 POUT (0 Past Due)", f"{pout:,}")
        col4.metric("✅ ACTIVE", f"{active:,}")

        # --- Display and Download Section ---
        st.subheader("📋 Merged Preview")
        st.dataframe(final_df.head(15))

    
        output = BytesIO()
        final_df.to_excel(output, index=False, engine="openpyxl")
        st.download_button(
            label="💾 Download Final Merged Excel",
            data=output.getvalue(),
            file_name="merged_TAD_ENDORSEMENT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("⚠️ Please upload and align both files before merging.")
from utils import process_excel_file, compare_excel_files, merge_excel_files