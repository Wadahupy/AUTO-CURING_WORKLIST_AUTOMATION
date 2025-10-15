import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import io
from msoffcrypto import OfficeFile
from typing import Optional, Dict, Union

def read_excel_file(uploaded_file: BytesIO, 
                   password: Optional[str] = None,
                   sheet_name: Optional[str] = None,
                   header_row: int = 0) -> Optional[pd.DataFrame]:
    """
    Safely read an Excel file with optional password protection.
    
    Args:
        uploaded_file: The uploaded file object
        password: Optional password for protected files
        sheet_name: Name or index of the sheet to read
        header_row: Row number to use as header (0-based)
        
    Returns:
        DataFrame or None if reading fails
    """
    try:
        if uploaded_file is None:
            return None
            
        file_ext = uploaded_file.name.split(".")[-1].lower()
        
        if file_ext == "csv":
            return pd.read_csv(uploaded_file)
            
        elif file_ext in ["xls", "xlsx"]:
            try:
                # Try reading as encrypted first
                decrypted = io.BytesIO()
                office_file = OfficeFile(uploaded_file)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                
                data = pd.read_excel(
                    decrypted,
                    sheet_name=sheet_name,
                    header=header_row,
                    engine="openpyxl"
                )
                
            except Exception as e:
                # If decrypt fails, try reading directly
                uploaded_file.seek(0)
                data = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=header_row,
                    engine="openpyxl"
                )
            
            # Handle dict result (multiple sheets)
            if isinstance(data, dict):
                if sheet_name and sheet_name in data:
                    return data[sheet_name]
                else:
                    return list(data.values())[0]
            return data
            
        else:
            st.error(f"❌ Unsupported file format: {file_ext}")
            return None
            
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        return None

def generate_download_button(df: pd.DataFrame,
                           button_text: str,
                           file_name: str,
                           file_type: str = "excel") -> None:
    """
    Create a download button for a DataFrame.
    
    Args:
        df: DataFrame to download
        button_text: Text to display on the button
        file_name: Name of the downloaded file
        file_type: 'csv' or 'excel'
    """
    try:
        if file_type == "csv":
            data = df.to_csv(index=False).encode("utf-8")
            mime = "text/csv"
        else:
            output = BytesIO()
            df.to_excel(output, index=False, engine="openpyxl")
            data = output.getvalue()
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            
        st.download_button(
            label=button_text,
            data=data,
            file_name=file_name,
            mime=mime,
            use_container_width=True
        )
    except Exception as e:
        st.error(f"❌ Error generating download: {str(e)}")

def show_dataframe_preview(df: pd.DataFrame,
                         title: str,
                         num_rows: int = 20) -> None:
    """Display a preview of a DataFrame with expandable section."""
    with st.expander(f"📋 {title}", expanded=True):
        st.dataframe(df.head(num_rows))

def process_excel_file(df: pd.DataFrame) -> pd.DataFrame:
    """
    Process the uploaded Excel file.
    This is a template function that you can modify based on your specific needs.
    
    Args:
        df (pd.DataFrame): Input DataFrame from the uploaded Excel file
        
    Returns:
        pd.DataFrame: Processed DataFrame
    """
    # Create a copy of the DataFrame to avoid modifying the original
    processed_df = df.copy()
    
    # Example processing steps (customize these based on your needs):
    # 1. Remove duplicates
    processed_df = processed_df.drop_duplicates()
    
    # 2. Handle missing values (example)
    processed_df = processed_df.fillna('')
    
    # 3. Add placeholder for additional processing
    # Add your custom processing logic here
    
    return processed_df

def compare_excel_files(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    """
    Compare two Excel files and highlight differences.
    This is a template function for future implementation.
    
    Args:
        df1 (pd.DataFrame): First DataFrame
        df2 (pd.DataFrame): Second DataFrame
        
    Returns:
        pd.DataFrame: DataFrame highlighting the differences
    """
    # Placeholder for comparison logic
    # Implement your comparison logic here
    pass

def merge_excel_files(df_list: list[pd.DataFrame], merge_on: str) -> pd.DataFrame:
    """
    Merge multiple Excel files based on a common column.
    This is a template function for future implementation.
    
    Args:
        df_list (list[pd.DataFrame]): List of DataFrames to merge
        merge_on (str): Column name to merge on
        
    Returns:
        pd.DataFrame: Merged DataFrame
    """
    # Placeholder for merging logic
    # Implement your merging logic here
    pass