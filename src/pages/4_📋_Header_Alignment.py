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
    Upload a CSV or Excel file with arbitrary column names. 
    This tool will automatically map your columns to the standard format.
    """)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload your data file",
        type=["csv", "xlsx", "xls"],
        help="CSV or Excel file with data to align"
    )
    
    if uploaded_file is None:
        st.info("👆 Upload a file to begin")
        return None
    
    # Read file
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # For Excel files, allow sheet selection
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "Select sheet",
                    sheet_names,
                    help="Choose which sheet to use"
                )
            else:
                selected_sheet = sheet_names[0]
            
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        return None
    
    st.success(f"✅ File loaded: {uploaded_file.name}")
    
    # Detect file type and show it
    detected_type = detect_file_type(uploaded_file.name)
    output_filename = generate_output_filename(uploaded_file.name)
    
    type_emoji = "📤" if detected_type == 'FOR_UPLOAD' else "✏️" if detected_type == 'FOR_UPDATE' else "📋"
    type_label = "FOR UPLOAD" if detected_type == 'FOR_UPLOAD' else "FOR UPDATE" if detected_type == 'FOR_UPDATE' else "UNKNOWN TYPE"
    
    st.info(f"{type_emoji} **Detected Type:** {type_label}\n\n**Output Filename:** `{output_filename}`")
    
    # Generate alignment report
    report = get_alignment_report(df)
    
    
    # Align headers
    try:
        aligned_df = align_headers(df)
        st.success("✅ Headers aligned successfully!")
    except Exception as e:
        st.error(f"❌ Error aligning headers: {e}")
        return None
    
    # Show aligned data preview
    with st.expander("📋 Aligned Data Preview", expanded=True):
        st.dataframe(aligned_df.head(20))
    
    # Download options
    st.subheader("💾 Download Aligned File")
    
    # Generate appropriate output filename
    output_base_name = generate_output_filename(uploaded_file.name)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Excel download
        excel_buffer = BytesIO()
        aligned_df.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0)
        
        st.download_button(
            label="📥 Download as Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"{output_base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # CSV download
        csv_buffer = aligned_df.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label="📥 Download as CSV (.csv)",
            data=csv_buffer,
            file_name=f"{output_base_name}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    return aligned_df


if __name__ == "__main__":
    st.set_page_config(page_title="Header Alignment Tool", layout="wide")
    
    render_header_alignment_tool()
