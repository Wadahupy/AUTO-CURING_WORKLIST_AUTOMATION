import streamlit as st
import pandas as pd
import sys
import io
from pathlib import Path
from datetime import datetime
from io import BytesIO
from msoffcrypto import OfficeFile
from dotenv import load_dotenv, find_dotenv
import os

# Add parent directory to Python path
sys.path.append(str(Path(__file__).parent.parent))

# Configure pandas to handle future warnings
pd.set_option('future.no_silent_downcasting', True)

from utils import read_excel_file, generate_download_button, show_dataframe_preview


# === Page Config ===
st.set_page_config(page_title="🔁 Daily TAD Update", layout="wide")
st.title("🔁 Daily TAD Update and Comparison")

# === Load environment variables ===
env_path = find_dotenv()
if not env_path:
    st.error("❌ .env file not found in project directory.")
    st.stop()

load_dotenv(env_path)

# === Get password from .env ===
DEFAULT_PASSWORD = os.getenv("DEFAULT_PASSWORD")

if not DEFAULT_PASSWORD:
    st.error("❌ Environment variable DEFAULT_PASSWORD not found or empty. Check your .env file.")
    st.stop()

# === Login Gate ===
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align:center;'>🔐 Login Required</h2>", unsafe_allow_html=True)
    password = st.text_input("Enter Password", type="password")

    if st.button("Login", use_container_width=True):
        if password == DEFAULT_PASSWORD:
            st.session_state["logged_in"] = True
            st.success("✅ Access granted! Loading the app...")
            st.rerun()
        else:
            st.error("❌ Incorrect password. Please try again.")
    st.stop()  # Stop execution here if not logged in

# === Logout button ===
st.sidebar.button("🔓 Logout", on_click=lambda: st.session_state.update({"logged_in": False}))

# === Standard Header Format ===
STANDARD_HEADERS = [
    "LAST BARCODE DATE", "LAST BARCODE", "PTP DATE", "AGENT", "CLASSIFICATION",
    "ENDO DATE", "DATE REFERRED", "CTL2", "CTL3", "CTL4", "DEBTOR ID",
    "LAN", "NAME", "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL",
    "MONTHLY AMORTIZATION", "INTEREST", "LPC", "INSURANCE", "PREPAYMENT",
    "CU PAYMENT", "LAST PAYMENT DATE", "PREM AMT", "PROD TYPE",
    "LPC YTD", "RATE", "REPRICING DATE", "DPD", "LOAN MATURITY",
    "DUE DATE", "OLDEST DUE DATE", "NEXT DUE DATE", "ADA SHORTAGE",
    "UNIT", "EMAIL", "ALTERNATIVE EMAIL ADDRESS", "MOBILE_ALS",
    "MOBILE_ALFES", "PRIMARY_NO_ALS", "BUS_NO_ALS", "LANDLINE_NO_ALS",
    "CO BORROWER", "CO BORROWER MOBILE_ALFES", "CO BORROWER LANDLINE__ALFES",
    "CO BORROWER EMAIL"
]

# Date columns that need formatting
DATE_COLUMNS = [
    "DATE", "PTP DATE", "ENDO DATE", "DATE REFERRED", "LAST PAYMENT DATE",
    "LOAN MATURITY", "OLDEST DUE DATE", "NEXT DUE DATE", "REPRICING DATE"
]

# === Helper Function to Read Excel/CSV (Handles Encrypted + Unencrypted) ===
def read_file(uploaded_file, password=None, sheet_name=None, header_row=0):
    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        if file_ext == "csv":
            return pd.read_csv(uploaded_file)
        elif file_ext in ["xls", "xlsx"]:
            decrypted = io.BytesIO()
            try:
                office_file = OfficeFile(uploaded_file)
                office_file.load_key(password=password or DEFAULT_PASSWORD)
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                data = pd.read_excel(decrypted, sheet_name=sheet_name, header=header_row, engine="openpyxl")
            except Exception:
                uploaded_file.seek(0)
                data = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, engine="openpyxl")

            if isinstance(data, dict):
                data = list(data.values())[0]
            return data
        else:
            st.error("Unsupported file format.")
            return None
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        return None


def standardize_headers(df):
    """Ensure dataframe has all standard headers in correct order."""
    if df is None or df.empty:
        return pd.DataFrame(columns=STANDARD_HEADERS)
    
    df = df.copy()
    df.columns = df.columns.map(str).str.strip().str.upper()  # Force all column names to string
    
    standardized = pd.DataFrame(columns=STANDARD_HEADERS)
    for col in STANDARD_HEADERS:
        if col in df.columns:
            standardized[col] = df[col]
        else:
            standardized[col] = pd.NA
    return standardized


def format_dates(df):
    """Format all date columns to mm/dd/yyyy without time."""
    if df is None or df.empty:
        return df
    
    df = df.copy()
    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%m/%d/%Y", errors="coerce").dt.strftime("%m/%d/%Y")
            df[col] = df[col].fillna("")  # Replace NaT with empty string
    return df



def align_headers(df, template_headers):
    """Ensure all template columns exist and are in correct order"""
    df = df.copy()
    for col in template_headers:
        if col not in df.columns:
            df[col] = ""
    return df[template_headers]


# === Upload Section ===
st.subheader("📤 Step 1: Upload Required Files")

col1, col2, col3 = st.columns(3)
with col1:
    yesterday_file = st.file_uploader("📘 Yesterday's Active List", type=["xls", "xlsx", "csv"], key="yesterday")
with col2:
    today_tad_file = st.file_uploader("📗 Today's TAD File (Sheet 1, starts A4)", type=["xls", "xlsx", "csv"], key="today_tad")
with col3:
    masterlist_file = st.file_uploader("📒 Masterlist (for revive lookups)", type=["xls", "xlsx", "csv"], key="masterlist")

import re

def validate_filename(filename: str, expected_type: str) -> bool:
    filename_upper = filename.upper()

    if expected_type == "active":
        # ACTIVE WORKLIST mmddyy
        return bool(re.match(r"^ACTIVE FILES \d{6}", filename_upper))
    elif expected_type == "tad":
        # TAD_SPM M1_mm.dd.yyyy
        return bool(re.match(r"^TAD_SPM M1_\d{2}\.\d{2}\.\d{4}", filename_upper))
    elif expected_type == "master":
        # MASTERLIST mmddyyyy
        return bool(re.match(r"^MASTERLIST \d{8}", filename_upper))
    return False


def load_and_validate(uploaded_file, expected_type, password=None, sheet_name=None, header_row=0):
    """Validate file name and read if valid."""
    if uploaded_file is None:
        return None

    if not validate_filename(uploaded_file.name, expected_type):
        st.error(f"❌ Invalid filename for {expected_type.upper()} file.\n\n"
                 f"Expected format:\n"
                 f"- ACTIVE WORKLIST → ACTIVE FILES mmddyy\n"
                 f"- TAD FILE → TAD_SPM M1_mm.dd.yyyy\n"
                 f"- MASTERLIST → MASTERLIST mmddyyyy")
        return None

    df = read_file(uploaded_file, password=password, sheet_name=sheet_name, header_row=header_row)
    return df


# === Load & Validate Each File ===
if yesterday_file:
    df = load_and_validate(yesterday_file, "active")
    if df is not None:
        df.columns = df.columns.astype(str).str.strip().str.upper()
        if "LAN" in df.columns:
            df["LAN"] = df["LAN"].astype(str).str.strip()
            df = df.drop_duplicates(subset=["LAN"], keep="last")
        st.session_state["yesterday_df"] = df
        st.success("✅ Yesterday's Active List loaded successfully.")

if today_tad_file:
    df = load_and_validate(today_tad_file, "tad", password=DEFAULT_PASSWORD, sheet_name=0, header_row=3)
    if df is not None:
        df.columns = df.columns.astype(str).str.strip().str.upper()
        if "LAN" in df.columns:
            df["LAN"] = df["LAN"].astype(str).str.strip()
            df = df.drop_duplicates(subset=["LAN"], keep="last")
        st.session_state["today_df"] = df
        st.success("✅ Today's TAD File loaded successfully.")

if masterlist_file:
    df = load_and_validate(masterlist_file, "master")
    if df is not None:
        df.columns = df.columns.astype(str).str.strip().str.upper()
        if "LAN" in df.columns:
            df["LAN"] = df["LAN"].astype(str).str.strip()
            df = df.drop_duplicates(subset=["LAN"], keep="last")
        st.session_state["masterlist_df"] = df
        st.success("✅ Masterlist loaded successfully.")



# === Processing Logic ===
    if st.button("🚀 Process Daily Updates"):
        # Reset previous results
        st.session_state["processing_complete"] = False

        # --- Check required files ---
        if "yesterday_df" not in st.session_state or "today_df" not in st.session_state:
            # Clear any previously stored results
            for key in ["final_active_list", "pullout_accounts", "for_update_accounts", "for_upload_accounts", "revive_display", "metrics"]:
                st.session_state.pop(key, None)

            st.warning("⚠️ Please upload both Yesterday's Active List and Today's TAD file before processing.")
            st.stop()

        # Continue processing only if both files exist
        yesterday_active = st.session_state["yesterday_df"].copy()
        today_tad = st.session_state["today_df"].copy()
        masterlist_df = st.session_state.get("masterlist_df", pd.DataFrame())

        # ✅ (The rest of your processing logic continues here as is)


        # Normalize columns
        yesterday_active.columns = yesterday_active.columns.str.strip().str.upper()
        today_tad.columns = today_tad.columns.str.strip().str.upper()
        if not masterlist_df.empty:
            masterlist_df.columns = masterlist_df.columns.str.strip().str.upper()

        # Standardize TAD column names
        today_tad.rename(columns={
            "LST BAL CHG DT": "LAST PAYMENT DATE",
            "CU PAYMENT AMT": "CU PAYMENT",
            "CU PAYMENT AMOUNT": "CU PAYMENT",
            "RATE": "RATE",
            "REPRICING DATE": "REPRICING DATE"
        }, inplace=True)

        # Column mapping
        columns_to_update = [
            "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL", "INTEREST", "LPC",
            "INSURANCE", "CU PAYMENT", "PREM AMT", "LAST PAYMENT DATE",
            "PROD TYPE", "LPC YTD", "RATE", "REPRICING DATE", "DPD", "ADA SHORTAGE"
        ]

        # Clean numeric columns
        def clean_numeric(df, columns):
            df = df.copy()
            for col in columns:
                if col in df.columns:
                    # Convert to string first
                    df[col] = df[col].astype(str)
                    
                    # Clean the strings
                    df[col] = (df[col]
                             .str.replace('₱', '', regex=False)
                             .str.replace('$', '', regex=False)
                             .str.replace(',', '', regex=False)
                             .str.replace('-', '', regex=False)
                             .str.strip())
                    
                    # Convert to numeric
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            return df

        numeric_cols = ["PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL", "INTEREST", "LPC",
                       "INSURANCE", "CU PAYMENT", "PREM AMT", "LPC YTD", "RATE"]

        today_tad = clean_numeric(today_tad, numeric_cols)
        yesterday_active = clean_numeric(yesterday_active, numeric_cols)

        # Update Active List columns based on TAD file
        updated_active = yesterday_active.copy()

        if "LAN" not in updated_active.columns or "LAN" not in today_tad.columns:
            st.error("❌ 'LAN' column missing. Cannot proceed.")
            st.stop()

        # Ensure all template headers exist
        for col in STANDARD_HEADERS:
            if col not in updated_active.columns:
                updated_active[col] = pd.NA

        # Get available TAD columns
        tad_cols_available = [col for col in columns_to_update if col in today_tad.columns]
        tad_update_cols = ["LAN"] + tad_cols_available
        tad_for_update = today_tad[tad_update_cols].copy()

        # Drop old columns and merge
        cols_to_drop = [col for col in tad_cols_available if col in updated_active.columns]
        updated_active.drop(columns=cols_to_drop, inplace=True)
        updated_active = pd.merge(updated_active, tad_for_update, on="LAN", how="left")

        # Filter accounts
        tad_with_pastdue = today_tad[today_tad["PAST DUE"] > 0].copy()
        tad_pastdue_count = len(tad_with_pastdue)

        updated_active["PAST DUE"] = pd.to_numeric(updated_active.get("PAST DUE", 0), errors="coerce").fillna(0)
        active_with_pastdue = updated_active[updated_active["PAST DUE"] > 0].copy()
        active_pastdue_count = len(active_with_pastdue)

        # Pullout accounts
        pullout_accounts = updated_active[updated_active["PAST DUE"] == 0].copy()
        pullout_count = len(pullout_accounts)

        # Remove pullout from updated active
        updated_active = updated_active[updated_active["PAST DUE"] > 0].copy()

        # Identify REVIVE accounts
        yesterday_all_lans = set(yesterday_active["LAN"].astype(str).values)
        tad_pastdue_lans = set(tad_with_pastdue["LAN"].astype(str).values)
        revive_lans = tad_pastdue_lans - yesterday_all_lans
        
        revive_accounts_tad = today_tad[today_tad["LAN"].astype(str).isin(revive_lans)].copy()
        revive_count = len(revive_accounts_tad)

        # Attach REVIVE accounts and lookup from MASTERLIST
        # ---------- STEP: Attach REVIVE Accounts (XLOOKUP from MASTERLIST) ----------
        if revive_count > 0:
            revive_accounts = revive_accounts_tad.copy()

            # Ensure template headers exist
            for col in STANDARD_HEADERS:
                if col not in revive_accounts.columns:
                    revive_accounts[col] = pd.NA

            # Classification and ENDO date
            revive_accounts["CLASSIFICATION"] = "REENDO"
            revive_accounts["ENDO DATE"] = datetime.today().strftime("%m/%d/%Y")

            # Lookup data from Masterlist
            if not masterlist_df.empty and "LAN" in masterlist_df.columns:
                ml = masterlist_df.copy()
                ml.columns = ml.columns.str.strip().str.upper()
                ml["LAN"] = ml["LAN"].astype(str).str.strip()

                # Merge Masterlist data
                revive_merged = pd.merge(
                    revive_accounts, ml, on="LAN", how="left", suffixes=("", "_ML")
                )

                # Replace invalid date values
                invalid_dates = ["0", "0/00/0000", "00/00/00", "NaT", "nan", "None", "", "1970-01-01", "01/01/1970",  "1/1/1970"]

                # 🧩 Fix DATE REFERRED
                if "DATE REFERRED_ML" in revive_merged.columns:
                    mask_ref = revive_merged["DATE REFERRED"].astype(str).str.strip().isin(invalid_dates) | revive_merged["DATE REFERRED"].isna()
                    revive_merged.loc[mask_ref, "DATE REFERRED"] = revive_merged.loc[mask_ref, "DATE REFERRED_ML"]

                # 🧩 Fix OLDEST DUE DATE (must come from masterlist)
                if "OLDEST DUE DATE_ML" in revive_merged.columns and "OLDEST DUE DATE" in revive_merged.columns:
                    # Convert TAD OLDEST DUE DATE to datetime with specific format
                    revive_merged["OLDEST DUE DATE"] = pd.to_datetime(revive_merged["OLDEST DUE DATE"], format="%m/%d/%Y", errors="coerce")

                    # Convert Masterlist Excel serial to datetime
                    def excel_date_to_datetime(val):
                        try:
                            val_float = float(val)
                            return pd.Timestamp('1899-12-30') + pd.to_timedelta(val_float, unit='D')
                        except:
                            return pd.NaT

                    revive_merged["OLDEST DUE DATE_ML"] = revive_merged["OLDEST DUE DATE_ML"].apply(excel_date_to_datetime)

                    # Pick the latest date
                    revive_merged["OLDEST DUE DATE"] = revive_merged[["OLDEST DUE DATE", "OLDEST DUE DATE_ML"]].max(axis=1)

                    # Calculate NEXT DUE DATE (+1 month)
                    revive_merged["NEXT DUE DATE"] = revive_merged["OLDEST DUE DATE"] + pd.DateOffset(months=1)

                    # Calculate DUE DATE as day of OLDEST DUE DATE
                    revive_merged["DUE DATE"] = revive_merged["OLDEST DUE DATE"].dt.day

                    # Format OLDEST DUE DATE and NEXT DUE DATE to mm/dd/yyyy strings
                    revive_merged["OLDEST DUE DATE"] = revive_merged["OLDEST DUE DATE"].dt.strftime("%m/%d/%Y").fillna("")
                    revive_merged["NEXT DUE DATE"] = revive_merged["NEXT DUE DATE"].dt.strftime("%m/%d/%Y").fillna("")

                    # Drop helper column
                    revive_merged.drop(columns=["OLDEST DUE DATE_ML"], inplace=True)




                # Fill all other missing columns from masterlist
                for col in STANDARD_HEADERS:
                    ml_col = f"{col}_ML"
                    if ml_col in revive_merged.columns:
                        # Convert both columns to the same type before filling
                        col_type = revive_merged[col].dtype
                        revive_merged[ml_col] = revive_merged[ml_col].astype(col_type)
                        
                        # Use loc for assignment to maintain types
                        mask = revive_merged[col].isna()
                        revive_merged.loc[mask, col] = revive_merged.loc[mask, ml_col]
                        
                        # Drop the masterlist column
                        revive_merged.drop(columns=[ml_col], inplace=True, errors="ignore")

                # 🧹 Format all date columns uniformly
                for col in DATE_COLUMNS:
                    if col in revive_merged.columns:
                        revive_merged[col] = pd.to_datetime(
                            revive_merged[col], errors="coerce"
                        ).dt.strftime("%m/%d/%Y")
                        revive_merged[col] = revive_merged[col].replace("NaT", "").fillna("")

                revive_accounts = revive_merged

            # Align and append to active list
            for col in STANDARD_HEADERS:
                if col not in updated_active.columns:
                    updated_active[col] = pd.NA

            updated_active = pd.concat(
                [updated_active, revive_accounts[updated_active.columns].copy()],
                ignore_index=True
            )


        # Format dates in all dataframes
        updated_active = format_dates(updated_active)
        pullout_accounts = format_dates(pullout_accounts)
        if revive_count > 0:
            revive_accounts_tad = format_dates(revive_accounts_tad)

        # Apply standard headers
        final_active_list = standardize_headers(updated_active)
        pullout_accounts = standardize_headers(pullout_accounts)
        
        # FOR UPDATE: existing accounts excluding revive
        for_update_accounts = updated_active[~updated_active["LAN"].astype(str).isin(revive_lans)].copy()
        for_update_accounts = standardize_headers(for_update_accounts)
        
        # FOR UPLOAD: all revive accounts
        tad_all_lans = set(today_tad["LAN"].astype(str).values)
        new_lans = tad_all_lans - yesterday_all_lans
        for_upload_accounts = updated_active[updated_active["LAN"].astype(str).isin(new_lans)].copy()
        for_upload_accounts = standardize_headers(for_upload_accounts)
        
        if revive_count > 0:
            revive_display = standardize_headers(revive_accounts_tad)
        else:
            revive_display = standardize_headers(pd.DataFrame())

        # Store results in session state
        st.session_state["final_active_list"] = final_active_list
        st.session_state["pullout_accounts"] = pullout_accounts
        st.session_state["for_update_accounts"] = for_update_accounts
        st.session_state["for_upload_accounts"] = for_upload_accounts
        st.session_state["revive_display"] = revive_display
        st.session_state["metrics"] = {
            "total_tad": len(today_tad),
            "tad_pastdue_count": tad_pastdue_count,
            "active_pastdue_count": active_pastdue_count,
            "pullout_count": pullout_count,
            "revive_count": revive_count,
            "forupdate_count": len(for_update_accounts)
        }
        st.session_state["processing_complete"] = True

# === Display Results (outside the button) ===
if st.session_state.get("processing_complete", False):
    metrics = st.session_state["metrics"]
    final_active_list = st.session_state["final_active_list"]
    pullout_accounts = st.session_state["pullout_accounts"]
    for_update_accounts = st.session_state["for_update_accounts"]
    for_upload_accounts = st.session_state["for_upload_accounts"]
    revive_display = st.session_state["revive_display"]

    # Dashboard Summary
    st.subheader(f"📅 {datetime.today():%b %d, %Y}: UPDATING TAD")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total TAD", f"{metrics['total_tad']:,}")
        st.metric("TAD with Past Due", f"{metrics['tad_pastdue_count']:,}")
    with col2:
        st.metric("Active with Past Due", f"{metrics['active_pastdue_count']:,}")
        st.metric("PULLOUT", f"{metrics['pullout_count']:,}")
    with col3:
        st.metric("REVIVE Accounts", f"{metrics['revive_count']:,}")
        st.metric("FOR UPDATE", f"{metrics['forupdate_count']:,}")

    # Excel Download Helper
    def to_excel(df):
        output = io.BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        return output.getvalue()

    # Download Section
    st.markdown("### 💾 Download Processed Files")
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "📘 UPDATED ACTIVE LIST", 
            to_excel(align_headers(final_active_list, STANDARD_HEADERS)), 
            file_name=f"ACTIVE WORKLIST {datetime.today():%m%d%y}.xlsx",
            use_container_width=True
        )
        st.download_button(
            "💤 PULLOUT Accounts", 
            to_excel(align_headers(pullout_accounts, STANDARD_HEADERS)), 
            file_name=f"PULLED OUT {datetime.today():%m%d%y}.xlsx",
            use_container_width=True
        )
        st.download_button(
            "🔁 REVIVE Accounts", 
            to_excel(align_headers(revive_display, STANDARD_HEADERS) if metrics["revive_count"] > 0 else pd.DataFrame(columns=STANDARD_HEADERS)), 
            file_name=f"REVIVE ACCOUNTS RAW {datetime.today():%m%d%y}.xlsx",
            use_container_width=True
        )
    
    with col2:
        st.download_button(
            "📝 FOR UPDATE", 
            to_excel(align_headers(for_update_accounts, STANDARD_HEADERS)), 
            file_name=f"FOR UPDATE {datetime.today():%m%d%y}.xlsx",
            use_container_width=True
        )
        st.download_button(
            "📤 FOR UPLOAD (New/Reendo)", 
            to_excel(align_headers(for_upload_accounts, STANDARD_HEADERS) if metrics["revive_count"] > 0 else pd.DataFrame(columns=STANDARD_HEADERS)), 
            file_name=f"FOR UPLOAD {datetime.today():%m%d%y}.xlsx",
            use_container_width=True
        )

    st.success("✅ Daily TAD update processed successfully!")
    
    # Show preview of key datasets
    with st.expander("👁️ Preview: REVIVE Accounts"):
        if metrics["revive_count"] > 0:
            st.dataframe(revive_display.head(10))
        else:
            st.info("No revive accounts found")
            
    with st.expander("👁️ Preview: Updated Active List (first 10)"):
        st.dataframe(final_active_list.head(10))

else:
    if st.session_state.get("processing_complete") is False:
        st.warning("⚠️ Please upload both Yesterday's Active List and Today's TAD file.")