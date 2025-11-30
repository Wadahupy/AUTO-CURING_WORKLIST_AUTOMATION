import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from msoffcrypto import OfficeFile
from pandas.tseries.offsets import DateOffset
from dotenv import load_dotenv, find_dotenv
import os

# === Page Config ===
st.set_page_config(page_title="🗓️ Weekly Endorsement", layout="wide")
st.title("🗓️ Weekly Endorsement (With Existing Masterlist)")
st.markdown("Upload **TAD**, **M1 Auto Endorsement**, and **Masterlist** — processing runs automatically after all files are uploaded.")

# === Load environment variables ===
try:
    DEFAULT_PASSWORD = st.secrets["DEFAULT_PASSWORD"]
except Exception:
    st.error("❌ Streamlit secret `DEFAULT_PASSWORD` not found. Please add it in the app settings under 'Edit Secrets'.")
    st.stop()
# === Login Gate ===

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


# === Template Headers ===
TEMPLATE_HEADERS = [
    "LAST BARCODE DATE", "LAST BARCODE", "PTP DATE", "AGENT", "CLASSIFICATION", "ENDO DATE", "DATE REFERRED",
    "CTL2", "CTL3", "CTL4", "DEBTOR ID", "LAN", "NAME", "PAST DUE", "PAYOFF AMOUNT", "PRINCIPAL",
    "MONTHLY AMORTIZATION", "INTEREST", "LPC", "INSURANCE", "PREPAYMENT", "CU PAYMENT",
    "LAST PAYMENT DATE", "PREM AMT", "PROD TYPE", "LPC YTD", "RATE", "REPRICING DATE", "DPD",
    "LOAN MATURITY", "DUE DATE", "OLDEST DUE DATE", "NEXT DUE DATE", "ADA SHORTAGE", "UNIT", "EMAIL",
    "ALTERNATIVE EMAIL ADDRESS", "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", "BUS_NO_ALS",
    "LANDLINE_NO_ALS", "CO BORROWER", "CO BORROWER MOBILE_ALFES", "CO BORROWER LANDLINE__ALFES",
    "CO BORROWER EMAIL"
]

DATE_COLUMNS = [
    "DATE", "PTP DATE", "ENDO DATE", "DATE REFERRED", "LAST PAYMENT DATE",
    "LOAN MATURITY", "OLDEST DUE DATE", "NEXT DUE DATE", "REPRICING DATE"
]

# === Helper Functions ===
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

def format_dates(series):
    """Format datetime series to mm/dd/yyyy"""
    return pd.to_datetime(series, errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "").fillna("")

def ensure_columns(df, cols):
    """Ensure all columns exist in dataframe"""
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    return df

def to_excel_bytes(df):
    """Convert dataframe to Excel bytes"""
    output = io.BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output.getvalue()

# === Upload Section ===
st.subheader("📤 Step 1: Upload Required Files")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**📗 TAD File**")
    st.caption("Sheet 1, starts at A4 (Encrypted)")
    tad_file = st.file_uploader("Upload TAD File", type=["xls", "xlsx", "csv"], key="tad")
    
    if tad_file:
        df = read_file(tad_file, password=DEFAULT_PASSWORD, sheet_name=0, header_row=3)
        if df is not None:
            df.columns = df.columns.str.strip().str.upper()
            if "LAN" in df.columns:
                df["LAN"] = df["LAN"].astype(str).str.strip()
                df = df.drop_duplicates(subset=["LAN"], keep="last")
            st.session_state["tad_df"] = df
            st.success(f"✅ {tad_file.name}")

with col2:
    st.markdown("**📘 M1 Auto Endorsement**")
    st.caption("Sheet: M1 AUTO SPM (Encrypted)")
    endorsement_file = st.file_uploader("Upload Endorsement File", type=["xls", "xlsx", "csv"], key="endorsement")
    
    if endorsement_file:
        df = read_file(endorsement_file, password=DEFAULT_PASSWORD, sheet_name="M1 AUTO SPM", header_row=0)
        if df is not None:
            df.columns = df.columns.str.strip().str.upper()
            if "ACCTNUM" in df.columns:
                df.rename(columns={"ACCTNUM": "LAN"}, inplace=True)
            if "LAN" in df.columns:
                df["LAN"] = df["LAN"].astype(str).str.strip()
                df = df.drop_duplicates(subset=["LAN"], keep="last")
            st.session_state["endorsement_df"] = df
            st.success(f"✅ {endorsement_file.name}")

with col3:
    st.markdown("**📒 Masterlist**")
    st.caption("Existing masterlist (Not Encrypted)")
    masterlist_file = st.file_uploader("Upload Masterlist", type=["xls", "xlsx", "csv"], key="masterlist")
    
    if masterlist_file:
        df = read_file(masterlist_file, password=None)
        if df is not None:
            df.columns = df.columns.str.strip().str.upper()
            if "LAN" in df.columns:
                df["LAN"] = df["LAN"].astype(str).str.strip()
                df = df.drop_duplicates(subset=["LAN"], keep="last")
            st.session_state["masterlist_df"] = df
            st.success(f"✅ {masterlist_file.name}")

# === Auto-Processing ===
if all(key in st.session_state for key in ["tad_df", "endorsement_df", "masterlist_df"]):
    
    st.markdown("---")
    st.subheader("⚙️ Processing Files...")
    
    try:
        # Load dataframes
        tad_df = st.session_state["tad_df"].copy()
        endorsement_df = st.session_state["endorsement_df"].copy()
        masterlist_df = st.session_state["masterlist_df"].copy()

        # Normalize columns
        tad_df.columns = tad_df.columns.str.strip().str.upper()
        endorsement_df.columns = endorsement_df.columns.str.strip().str.upper()
        masterlist_df.columns = masterlist_df.columns.str.strip().str.upper()

        # === Step 2: Build ACTIVE FILE from TAD ===
        # Map TAD columns to template
        tad_column_map = {
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
            "RATE": "RATE",
            "REPRICING DATE": "REPRICING DATE",
            "DPD": "DPD",
            "ADA SHORTAGE": "ADA SHORTAGE"
        }

        # Build active file from TAD
        active_file = pd.DataFrame()
        for tad_col, template_col in tad_column_map.items():
            if tad_col in tad_df.columns:
                active_file[template_col] = tad_df[tad_col]
            else:
                active_file[template_col] = ""

        # Ensure LAN exists
        if "LAN" not in active_file.columns:
            st.error("❌ LAN column missing in TAD file. Cannot proceed.")
            st.stop()
        
        active_file["LAN"] = active_file["LAN"].astype(str).str.strip()

        # === Step 3: Fill missing values from M1 Auto Endorsement ===
        endorsement_column_map = {
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

        # Merge with endorsement data
        endorsement_df["LAN"] = endorsement_df["LAN"].astype(str).str.strip()
        active_file = pd.merge(active_file, endorsement_df, on="LAN", how="left", suffixes=("", "_ENDO"))

        # Apply endorsement mappings
        for src_col, tgt_col in endorsement_column_map.items():
            if src_col in active_file.columns:
                if tgt_col not in active_file.columns:
                    active_file[tgt_col] = active_file[src_col]
                else:
                    active_file[tgt_col] = active_file[tgt_col].fillna(active_file[src_col])

        # === Step 4: Classification and DATE REFERRED Logic ===
        masterlist_df["LAN"] = masterlist_df["LAN"].astype(str).str.strip()
        masterlist_lans = set(masterlist_df["LAN"].values)
        
        today_str = datetime.today().strftime("%m/%d/%Y")
        
        # Classification
        active_file["CLASSIFICATION"] = active_file["LAN"].apply(
            lambda lan: "REENDO" if lan in masterlist_lans else "NEW ENDO"
        )
        
        # ENDO DATE = today for all
        active_file["ENDO DATE"] = today_str
        
        # DATE REFERRED logic
        def get_date_referred(row):
            if row["CLASSIFICATION"] == "NEW ENDO":
                return today_str
            else:
                # REENDO: lookup from masterlist
                master_row = masterlist_df[masterlist_df["LAN"] == row["LAN"]]
                if not master_row.empty and "DATE REFERRED" in master_row.columns:
                    date_ref = master_row["DATE REFERRED"].values[0]
                    if pd.notna(date_ref) and str(date_ref).strip() != "":
                        # Format if it's a date
                        try:
                            return pd.to_datetime(date_ref).strftime("%m/%d/%Y")
                        except:
                            return str(date_ref)
                return today_str  # Fallback if not found in masterlist
        
        active_file["DATE REFERRED"] = active_file.apply(get_date_referred, axis=1)

        # === Step 5: Validate DATE REFERRED and ENDO DATE Consistency ===
        suspicious = active_file[
            (active_file["CLASSIFICATION"] == "REENDO") & 
            (active_file["DATE REFERRED"] == active_file["ENDO DATE"])
        ]
        
        if not suspicious.empty:
            st.error("❌ Validation Failed: Some REENDO accounts have DATE REFERRED equal to ENDO DATE.")
            st.error("This indicates the Masterlist may be incorrect or missing DATE REFERRED values.")
            st.markdown("**Suspicious Accounts:**")
            st.dataframe(suspicious[["LAN", "NAME", "CLASSIFICATION", "DATE REFERRED", "ENDO DATE"]].head(20))
            st.warning("⚠️ Please check and upload the correct Masterlist, then try again.")
            st.stop()

        # === Compute Due Dates ===
        try:
            active_file["OLDEST DUE DATE"] = pd.to_datetime(active_file.get("OLDEST DUE DATE"), errors="coerce")
            active_file["DUE DATE"] = active_file["OLDEST DUE DATE"].dt.day
            active_file["NEXT DUE DATE"] = active_file["OLDEST DUE DATE"] + DateOffset(months=1)
            
            # Format all date columns
            for col in DATE_COLUMNS:
                if col in active_file.columns:
                    active_file[col] = format_dates(active_file[col])
            
            # DUE DATE as integer string
            active_file["DUE DATE"] = active_file["DUE DATE"].fillna("").astype(str).replace("nan", "").replace(".0", "")
        except Exception as e:
            st.warning(f"⚠️ Could not compute due dates: {e}")

        # === Ensure all template columns exist ===
        active_file = ensure_columns(active_file, TEMPLATE_HEADERS)
        final_active = active_file[TEMPLATE_HEADERS].copy()

        # === Step 6: Create Outputs ===
        
        # 6.1: FOR UPLOAD (NEW ENDO only)
        for_upload = final_active[final_active["CLASSIFICATION"] == "NEW ENDO"].copy()
        
        # 6.2: FOR UPDATE (REENDO only)
        for_update = final_active[final_active["CLASSIFICATION"] == "REENDO"].copy()
        
        # 6.3: Consolidated MASTERLIST
        # The consolidated masterlist combines the original masterlist with ALL accounts from the active list
        # Use ALL columns from final_active (complete data including all mapped fields)
        active_for_master = final_active.copy()
        
        # Ensure masterlist has all columns from active list
        for col in active_for_master.columns:
            if col not in masterlist_df.columns:
                masterlist_df[col] = ""
        
        # Combine: original masterlist + all active list accounts
        consolidated_masterlist = pd.concat([masterlist_df, active_for_master], ignore_index=True)
        
        # Keep duplicates - do not remove them (allows tracking of multiple entries per LAN)
        
        # Format all date columns in consolidated masterlist to MM/DD/YYYY
        for col in DATE_COLUMNS:
            if col in consolidated_masterlist.columns:
                consolidated_masterlist[col] = format_dates(consolidated_masterlist[col])

        # === Display Results ===
        st.success("✅ Processing complete!")
        
        # Metrics Dashboard
        st.subheader(f"📊 Dashboard Summary — {datetime.today():%b %d, %Y}")
        
        total_accounts = len(final_active)
        new_endo_count = (final_active["CLASSIFICATION"] == "NEW ENDO").sum()
        reendo_count = (final_active["CLASSIFICATION"] == "REENDO").sum()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Accounts", f"{total_accounts:,}")
        col2.metric("FOR UPLOAD (NEW ENDO)", f"{new_endo_count:,}", 
                   delta=f"{(new_endo_count/total_accounts*100):.1f}%" if total_accounts else "0%")
        col3.metric("FOR UPDATE (REENDO)", f"{reendo_count:,}",
                   delta=f"{(reendo_count/total_accounts*100):.1f}%" if total_accounts else "0%")
        col4.metric("Masterlist Size", f"{len(consolidated_masterlist):,}")

        # === Download Section ===
        st.markdown("---")
        st.subheader("💾 Download Files")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.download_button(
                "📘 Download ACTIVE FILE",
                to_excel_bytes(final_active),
                file_name=f"Active_File_{datetime.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            st.download_button(
                "📒 Download CONSOLIDATED MASTERLIST",
                to_excel_bytes(consolidated_masterlist),
                file_name=f"Masterlist_{datetime.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col3:
            st.download_button(
                "📤 Download FOR UPLOAD (NEW ENDO)",
                to_excel_bytes(for_upload),
                file_name=f"For_Upload_{datetime.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col4:
            st.download_button(
                "✏️ Download FOR UPDATE (REENDO)",
                to_excel_bytes(for_update),
                file_name=f"For_Update_{datetime.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # === Preview Section ===
        st.markdown("---")
        
        with st.expander("👁️ Preview: Active File (first 20 rows)", expanded=True):
            st.dataframe(final_active.head(20))
        
        with st.expander("👁️ Preview: FOR UPLOAD (NEW ENDO accounts)"):
            st.dataframe(for_upload.head(20))
        
        with st.expander("👁️ Preview: FOR UPDATE (REENDO accounts)"):
            st.dataframe(for_update.head(20))
        
        with st.expander("👁️ Preview: Consolidated Masterlist (first 20 rows)"):
            st.dataframe(consolidated_masterlist.head(20))
        
        with st.expander("📊 Classification Breakdown"):
            classification_counts = final_active["CLASSIFICATION"].value_counts()
            st.bar_chart(classification_counts)

    except Exception as e:
        st.error(f"❌ Processing Error: {e}")
        st.exception(e)

else:
    st.info("ℹ️ Please upload all three files (TAD, M1 Auto Endorsement, and Masterlist) to begin processing.")