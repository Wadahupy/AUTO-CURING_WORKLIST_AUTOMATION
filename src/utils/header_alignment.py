import pandas as pd
from typing import Dict, List, Optional


# Standard headers in required order - only necessary columns
STANDARD_HEADERS = [
    "LAN",
    "CH CODE"
    "NAME",
    "CTL4",
    "PAST DUE",
    "PAYOFF AMOUNT",
    "PRINCIPAL",
    "LPC",
    "ADA SHORTAGE",
    "EMAIL",
    "MOBILE_ALS",
    "MOBILE_ALFES",
    "PRIMARY_NO_ALS",
    "BUS_NO_ALS",
    "LANDLINE_NO_ALS",
    "DATE REFERRED",
    "UNIT",
    "DPD"
]

# Mapping from standard column names to possible input column names
# Key = standard column name
# Value = list of possible input column names
ALIGNMENT_MAP = {
    "LAN": ["LAN", "ACCOUNT NUMBER", "ACCTNUM"],
    "CH CODE": ["LAN", "ch code"],
    "NAME": ["NAME", "DEBTOR NAME", "BORROWER NAME"],
    "CTL4": ["CTL4"],
    "PAST DUE": ["PAST DUE", "OVERDUE AMOUNT"],
    "PAYOFF AMOUNT": ["PAYOFF AMOUNT", "PAYOFF AMT"],
    "PRINCIPAL": ["PRINCIPAL", "PRINCIPAL AMOUNT"],
    "LPC": ["LPC", "LOAN PRINCIPAL CONTRACTED"],
    "ADA SHORTAGE": ["ADA SHORTAGE", "ADA SHORT"],
    "EMAIL": ["EMAIL", "EMAIL_ALS", "BORROWER EMAIL"],
    "MOBILE_ALS": ["MOBILE_ALS", "MOBILE NO ALS", "MOBILE NUMBER"],
    "MOBILE_ALFES": ["MOBILE_ALFES", "MOBILE NO ALFES"],
    "PRIMARY_NO_ALS": ["PRIMARY_NO_ALS", "PRIMARY NO"],
    "BUS_NO_ALS": ["BUS_NO_ALS", "BUSINESS NO"],
    "LANDLINE_NO_ALS": ["LANDLINE_NO_ALS", "LANDLINE NO ALFES", "LANDLINE"],
    "DATE REFERRED": ["DATE REFERRED", "REFERRAL DATE"],
    "UNIT": ["UNIT", "SHORT DESCRIPTION", "UNIT DESCRIPTION"],
    "DPD": ["DPD", "DAYS PAST DUE"]
}


def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize input DataFrame column names to uppercase and strip whitespace.
    
    Args:
        df: Input DataFrame
        
    Returns:
        DataFrame with normalized column names
    """
    df.columns = df.columns.astype(str).str.strip().str.upper()
    return df


def find_column_in_dataframe(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
    """
    Find a column in the DataFrame that matches any of the possible names.
    
    Args:
        df: Input DataFrame with normalized column names
        possible_names: List of possible column name variations
        
    Returns:
        The actual column name in the DataFrame, or None if not found
    """
    # Normalize possible names to uppercase
    normalized_possible = [name.upper().strip() for name in possible_names]
    
    for col in df.columns:
        if col.upper().strip() in normalized_possible:
            return col
    
    return None


def format_date_column(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    """
    Format a date column to m/dd/yyyy format without time component.
    
    Args:
        df: DataFrame containing the date column
        column_name: Name of the column to format
        
    Returns:
        DataFrame with formatted date column
    """
    if column_name not in df.columns:
        return df
    
    try:
        # Convert to datetime, handling various formats
        df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
        
        # Format as m/dd/yyyy (m = month without leading zero, dd = day with leading zero, yyyy = year)
        # Using strftime with custom formatting
        df[column_name] = df[column_name].apply(
            lambda x: x.strftime('%-m/%d/%Y').lstrip('0') if pd.notna(x) else ''
        )
        
        # For Windows compatibility (strftime doesn't support %-m)
        # Apply manual leading zero removal for month
        df[column_name] = df[column_name].apply(
            lambda x: x.lstrip('0') if isinstance(x, str) and x.startswith('0') else x
        )
        
    except Exception as e:
        # Fallback: try direct string manipulation
        try:
            df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
            df[column_name] = df[column_name].dt.strftime('%m/%d/%Y')
            # Remove leading zeros from month
            df[column_name] = df[column_name].str.replace(r'^0(\d/)', r'\1', regex=True)
            # Replace NaT with empty string
            df[column_name] = df[column_name].replace('NaT', '')
        except Exception as e2:
            print(f"Warning: Could not format date column '{column_name}': {e2}")
    
    return df


def align_headers(df: pd.DataFrame, custom_map: Optional[Dict[str, List[str]]] = None) -> pd.DataFrame:
    """
    Align input DataFrame columns to standardized header format.
    
    Args:
        df: Input DataFrame with arbitrary column names
        custom_map: Optional custom mapping dictionary. If None, uses ALIGNMENT_MAP
        
    Returns:
        DataFrame with standardized headers in correct order, missing columns filled with empty strings
    """
    # Normalize input column names
    df = normalize_column_names(df)
    
    # Use provided map or default
    mapping = custom_map if custom_map is not None else ALIGNMENT_MAP
    
    # Create output DataFrame with standard headers
    aligned_df = pd.DataFrame(columns=STANDARD_HEADERS)
    
    # Map each standard header to input data
    for standard_col in STANDARD_HEADERS:
        possible_input_cols = mapping.get(standard_col, [])
        
        # Ensure it's a list
        if isinstance(possible_input_cols, str):
            possible_input_cols = [possible_input_cols]
        
        # Find matching column in input DataFrame
        input_col = find_column_in_dataframe(df, possible_input_cols)
        
        if input_col and input_col in df.columns:
            aligned_df[standard_col] = df[input_col]
        else:
            # Fill with empty strings if column not found
            aligned_df[standard_col] = ""
    
    # Format DATE REFERRED column
    aligned_df = format_date_column(aligned_df, "DATE REFERRED")
    
    return aligned_df


def align_and_export(
    input_file: str,
    output_file: str,
    sheet_name: Optional[str] = None,
    header_row: int = 0,
    custom_map: Optional[Dict[str, List[str]]] = None,
    file_format: str = "excel"
) -> bool:
    """
    Read input file, align headers, and export to output file.
    
    Args:
        input_file: Path to input CSV or Excel file
        output_file: Path to output file
        sheet_name: Sheet name for Excel files (None for first sheet)
        header_row: Row number containing headers (0-indexed)
        custom_map: Optional custom mapping dictionary
        file_format: Output format ('excel' or 'csv')
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Read input file
        if input_file.lower().endswith('.csv'):
            df = pd.read_csv(input_file)
        else:  # Excel file
            df = pd.read_excel(input_file, sheet_name=sheet_name, header=header_row)
        
        # Align headers
        aligned_df = align_headers(df, custom_map)
        
        # Export to output file
        if file_format.lower() == 'csv':
            aligned_df.to_csv(output_file, index=False)
        else:  # Excel
            aligned_df.to_excel(output_file, index=False, engine='openpyxl')
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False


def get_alignment_report(df: pd.DataFrame, custom_map: Optional[Dict[str, List[str]]] = None) -> Dict:
    """
    Generate a report showing which input columns map to which standard columns.
    
    Args:
        df: Input DataFrame
        custom_map: Optional custom mapping dictionary
        
    Returns:
        Dictionary with mapping report
    """
    df = normalize_column_names(df)
    mapping = custom_map if custom_map is not None else ALIGNMENT_MAP
    
    report = {
        "found_mappings": {},
        "missing_columns": [],
        "input_columns": list(df.columns),
        "standard_columns": STANDARD_HEADERS
    }
    
    for standard_col in STANDARD_HEADERS:
        possible_input_cols = mapping.get(standard_col, [])
        if isinstance(possible_input_cols, str):
            possible_input_cols = [possible_input_cols]
        
        input_col = find_column_in_dataframe(df, possible_input_cols)
        
        if input_col:
            report["found_mappings"][standard_col] = input_col
        else:
            report["missing_columns"].append(standard_col)
    
    return report
