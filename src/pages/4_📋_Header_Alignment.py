import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import sys
from datetime import datetime
import re

# Add parent directory to path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

# Import from header_alignment module
from utils.header_alignment import (
    align_headers,
    get_alignment_report,
    STANDARD_HEADERS,
    ALIGNMENT_MAP
)


def format_date_referred(df: pd.DataFrame) -> pd.DataFrame:
    """
    Format DATE REFFERED column to MM/MMMM/YYYY format if it exists
    """
    col_name = "DATE REFFERED"

    if col_name in df.columns:
        # Convert to datetime, handling various formats
        df[col_name] = pd.to_datetime(
            df[col_name],
            errors="coerce",
            infer_datetime_format=True
        )
        # Format as MM/MMMM/YYYY (e.g., 01/January/2026), keeping NaT values as empty strings
        df[col_name] = df[col_name].dt.strftime("%m/%B/%Y").fillna("")

    return df


def detect_file_type(filename: str) -> str:
    """
    Detect if the file is FOR UPLOAD or FOR UPDATE based on filename.
    
    Args:
        filename: The uploaded filename
        
    Returns:
        'FOR_UPLOAD', 'FOR_UPDATE', or 'UNKNOWN'
    """
    filename_upper = filename.upper()
    
    if 'FOR UPLOAD' in filename_upper or 'FORUPLOAD' in filename_upper:
        return 'FOR_UPLOAD'
    elif 'FOR UPDATE' in filename_upper or 'FORUPDATE' in filename_upper:
        return 'FOR_UPDATE'
    else:
        return 'UNKNOWN'


def generate_output_filename(input_filename: str) -> str:
    """
    Generate the output filename based on the input filename type.
    
    Args:
        input_filename: The uploaded filename
        
    Returns:
        The formatted output filename
    """
    file_type = detect_file_type(input_filename)
    today = datetime.now().strftime('%m%d%Y')
    
    if file_type == 'FOR_UPLOAD':
        return f"BPI_AUTOCURING_FORUPLOADS_{today}"
    elif file_type == 'FOR_UPDATE':
        return f"BPI_AUTOCURING_FORUPDATES_{today}"
    else:
        # Default naming if file type cannot be detected
        return f"BPI_AUTOCURING_ALIGNED_{today}"


def render_header_alignment_tool():
    """Render the header alignment tool in Streamlit"""
    
    st.header("📋 Column Header Alignment Tool")
    st.markdown("""
    Upload CSV or Excel files for alignment. You can upload:
    - **FOR UPLOAD file** only
    - **FOR UPDATE file** only  
    - **Both files** together
    """)
    
    # Initialize session state for tracking uploaded files
    if "uploaded_for_upload" not in st.session_state:
        st.session_state["uploaded_for_upload"] = None
    if "uploaded_for_update" not in st.session_state:
        st.session_state["uploaded_for_update"] = None
    
    # File upload section
    st.subheader("📤 Step 1: Upload Files")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📘 FOR UPLOAD File**")
        for_upload_file = st.file_uploader(
            "Upload FOR UPLOAD file",
            type=["csv", "xlsx", "xls"],
            key="for_upload_uploader",
            help="File containing accounts to upload"
        )
        if for_upload_file:
            st.session_state["uploaded_for_upload"] = for_upload_file
            st.success(f"✅ FOR UPLOAD: {for_upload_file.name}")
    
    with col2:
        st.markdown("**📗 FOR UPDATE File**")
        for_update_file = st.file_uploader(
            "Upload FOR UPDATE file",
            type=["csv", "xlsx", "xls"],
            key="for_update_uploader",
            help="File containing accounts to update"
        )
        if for_update_file:
            st.session_state["uploaded_for_update"] = for_update_file
            st.success(f"✅ FOR UPDATE: {for_update_file.name}")
    
    # Check if any file is uploaded
    has_for_upload = st.session_state["uploaded_for_upload"] is not None
    has_for_update = st.session_state["uploaded_for_update"] is not None
    
    if not has_for_upload and not has_for_update:
        st.info("👆 Upload at least one file to begin")
        return None
    
    # Process files
    st.markdown("---")
    st.subheader("⚙️ Step 2: Process & Align")
    
    processed_files = {}
    
    # Process FOR UPLOAD file
    if has_for_upload:
        uploaded_file = st.session_state["uploaded_for_upload"]
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                selected_sheet = sheet_names[0] if len(sheet_names) == 1 else st.selectbox("Select sheet (FOR UPLOAD)", sheet_names)
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Align headers
            aligned_df = align_headers(df)
            aligned_df = format_date_referred(aligned_df)
            processed_files["FOR_UPLOAD"] = {
                "df": aligned_df,
                "filename": "BPI_AUTOCURING_FORUPLOADS_" + datetime.now().strftime('%m%d%Y'),
                "status": "✅"
            }
            st.success(f"✅ FOR UPLOAD aligned: {len(aligned_df)} records")
        except Exception as e:
            st.error(f"❌ Error processing FOR UPLOAD: {e}")
    
    # Process FOR UPDATE file
    if has_for_update:
        uploaded_file = st.session_state["uploaded_for_update"]
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                selected_sheet = sheet_names[0] if len(sheet_names) == 1 else st.selectbox("Select sheet (FOR UPDATE)", sheet_names)
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Align headers
            aligned_df = align_headers(df)
            aligned_df = format_date_referred(aligned_df)
            processed_files["FOR_UPDATE"] = {
                "df": aligned_df,
                "filename": "BPI_AUTOCURING_FORUPDATES_" + datetime.now().strftime('%m%d%Y'),
                "status": "✅"
            }
            st.success(f"✅ FOR UPDATE aligned: {len(aligned_df)} records")
        except Exception as e:
            st.error(f"❌ Error processing FOR UPDATE: {e}")
    
    if not processed_files:
        st.warning("⚠️ No files were successfully processed")
        return None
    
    # Store processed files in session state for download persistence
    st.session_state["processed_files"] = processed_files
    st.session_state["processing_complete"] = True
    
    return processed_files


# Display previews and downloads if processing is complete
if __name__ == "__main__":
    st.set_page_config(page_title="Header Alignment Tool", layout="wide")
    
    # Initialize session state flags
    if "processing_complete" not in st.session_state:
        st.session_state["processing_complete"] = False
    
    # Render main tool
    processed_files = render_header_alignment_tool()
    
    # Display previews and downloads if files were processed
    if st.session_state.get("processing_complete", False) and "processed_files" in st.session_state:
        processed_files = st.session_state["processed_files"]
        
        st.markdown("---")
        st.subheader("👁️ Step 3: Preview Aligned Files")
        
        # Show previews
        if "FOR_UPLOAD" in processed_files:
            with st.expander("📘 FOR UPLOAD Preview", expanded=True):
                st.dataframe(processed_files["FOR_UPLOAD"]["df"].head(20))
        
        if "FOR_UPDATE" in processed_files:
            with st.expander("📗 FOR UPDATE Preview", expanded=True):
                st.dataframe(processed_files["FOR_UPDATE"]["df"].head(20))
        
        # Download section
        st.markdown("---")
        st.subheader("💾 Step 4: Download Aligned Files")
        
        col1, col2 = st.columns(2)
        
        if "FOR_UPLOAD" in processed_files:
            file_data = processed_files["FOR_UPLOAD"]
            with col1:
                st.markdown("**📤 FOR UPLOAD**")
                
                # Excel download
                excel_buffer = BytesIO()
                file_data["df"].to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
                st.download_button(
                    label="📥 Download as Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"{file_data['filename']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="for_upload_xlsx"
                )
                
                # CSV download
                csv_buffer = file_data["df"].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv_buffer,
                    file_name=f"{file_data['filename']}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="for_upload_csv"
                )
        
        if "FOR_UPDATE" in processed_files:
            file_data = processed_files["FOR_UPDATE"]
            with col2:
                st.markdown("**📝 FOR UPDATE**")
                
                # Excel download
                excel_buffer = BytesIO()
                file_data["df"].to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
                st.download_button(
                    label="📥 Download as Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"{file_data['filename']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="for_update_xlsx"
                )
                
                # CSV download
                csv_buffer = file_data["df"].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv_buffer,
                    file_name=f"{file_data['filename']}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="for_update_csv"
                )
