"""
Simple transformer module that mimics the transformation logic from excel_processor.py

This module provides simplified versions of the key transformation functions
used in the excel processing workflow.
"""

import re
import pandas as pd
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Optional, Dict, List


# Constants from excel_processor.py
TARGET_TERMS = {12, 24, 36, 48, 60}

BASE_COLS = [
    'Start Month',
    'State', 
    'Utility',
    'Congestion Zone',
    'Load Factor',
    'Term',
    'Product',
    '0-200,000',
]

MASTER_COLS = [
    'ID', 'Price_Date', 'Date', 'Zone', 'Load', 'REP1', 'Term', 'Min_MWh', 'Max_MWh',
    'Daily_No_Ruc', 'RUC_Nodal', 'Daily', 'Com_Disc', 'HOA_Disc', 'Broker_Fee', 'Meter_Fee', 'Max_Meters'
]


def create_master_table_backup(master_path: Path) -> Optional[Path]:
    """
    Create a timestamped backup copy of the master table before modifications.

    Args:
        master_path: Path to the master table file

    Returns:
        Path to the backup file if successful, None if failed
    """
    if not master_path.exists():
        print(f"Master table not found for backup: {master_path}")
        return None

    try:
        # Create backup directory if it doesn't exist
        backup_dir = master_path.parent / "backups"
        backup_dir.mkdir(exist_ok=True)

        # Generate timestamped backup filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{master_path.stem}_backup_{timestamp}{master_path.suffix}"
        backup_path = backup_dir / backup_name

        # Copy the master table to backup location
        shutil.copy2(master_path, backup_path)
        print(f"Created backup: {backup_path}")
        return backup_path

    except Exception as e:
        print(f"Failed to create backup of master table: {e}")
        return None


def parse_term_to_int(val) -> Optional[int]:
    """
    Parse a term value to integer, extracting numeric value from various formats.
    
    Args:
        val: Term value (could be int, float, string, etc.)
        
    Returns:
        Integer term value if valid and in TARGET_TERMS, None otherwise
    """
    if pd.isna(val) or val is None:
        return None
    
    if isinstance(val, (int, float)):
        return int(val) if float(int(val)) == val else None
    
    s = str(val).strip()
    if s.isdigit():
        return int(s)
    
    # Extract first number from string
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else None


def filter_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filter a DataFrame to keep only 'Fixed Price' products with valid terms.
    
    Args:
        df: Input DataFrame with product and term columns
        
    Returns:
        Filtered DataFrame with BASE_COLS columns
    """
    # Find product and term columns (case-insensitive)
    prod_col = next((col for col in df.columns if col.strip().lower() in ['product', 'products']), None)
    term_col = next((col for col in df.columns if col.strip().lower() in ['term', 'terms']), None)
    
    if prod_col is None or term_col is None:
        return pd.DataFrame(columns=BASE_COLS)
    
    # Filter for 'Fixed Price' products
    prod_mask = df[prod_col].astype(str).str.strip().str.lower().eq('fixed price')
    
    # Filter for valid terms
    term_vals = df[term_col].map(parse_term_to_int)
    term_mask = term_vals.isin(TARGET_TERMS)
    
    # Get available columns from BASE_COLS
    cols = [c for c in BASE_COLS if c in df.columns]
    filtered = df.loc[prod_mask & term_mask, cols].copy()
    
    return filtered


def normalize_column_name(name: str) -> str:
    """Normalize column name for flexible matching."""
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def find_column(df: pd.DataFrame, candidates: List[str], allow_contains: bool = False) -> Optional[str]:
    """
    Find a column in DataFrame using candidate names.
    
    Args:
        df: DataFrame to search
        candidates: List of candidate column names
        allow_contains: Whether to allow partial matches
        
    Returns:
        Actual column name if found, None otherwise
    """
    # Build normalization map
    norm_map = {normalize_column_name(col): col for col in df.columns}
    
    # Try exact matches first
    for candidate in candidates:
        key = normalize_column_name(candidate)
        if key in norm_map:
            return norm_map[key]
    
    # Try contains matches if allowed
    if allow_contains:
        for key_norm, orig_col in norm_map.items():
            for candidate in candidates:
                if normalize_column_name(candidate) in key_norm:
                    return orig_col
    
    return None


def parse_zone_from_description(text: str) -> str:
    """Extract zone name from description by removing load factor suffix."""
    s = str(text or '').strip()
    for token in [' Low Load Factor', ' Medium Load Factor', ' High Load Factor']:
        if s.endswith(token):
            return s[:-len(token)]
    return s


def parse_load_factor_from_description(text: str) -> str:
    """Extract load factor from description text."""
    s = str(text or '').strip()
    if 'Low Load Factor' in s:
        return 'LOW'
    elif 'Medium Load Factor' in s:
        return 'MED'
    elif 'High Load Factor' in s:
        return 'HIGH'
    return ''


def transform_to_master_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transform DataFrame to master table format.
    
    This is a simplified version of hda_matrix_to_master_cols from excel_processor.py
    
    Args:
        df: Input DataFrame with various column formats
        
    Returns:
        DataFrame with master table column structure
    """
    # Column candidates for flexible matching
    desc_candidates = ['MatrixDescription', 'Matrix Description', 'Description', 'Desc']
    price_candidates = ['Price', 'Price $', 'Price($)', 'Matrix Price', 'Rate', 'Base Price']
    green_candidates = ['GreenPrice', 'Green Price']
    term_candidates = ['TermCode', 'Term', 'Term Months', 'Term_Months', 'TermLength', 'Term (months)']
    start_candidates = ['StartDate', 'Start Date', 'StartMonth', 'Start Month', 'Delivery Start', 'First Delivery']
    zone_candidates = ['Zone', 'Congestion Zone']
    lf_candidates = ['Load Factor', 'LoadFactor', 'LF']
    product_candidates = ['Product', 'Products']
    
    # Find columns
    c_desc = find_column(df, desc_candidates, allow_contains=True)
    c_price = find_column(df, price_candidates, allow_contains=True)
    c_green = find_column(df, green_candidates, allow_contains=True)
    c_term = find_column(df, term_candidates, allow_contains=True)
    c_start = find_column(df, start_candidates, allow_contains=True)
    c_zone = find_column(df, zone_candidates, allow_contains=True)
    c_lf = find_column(df, lf_candidates, allow_contains=True)
    c_prod = find_column(df, product_candidates, allow_contains=True)
    
    # Create output DataFrame
    out = pd.DataFrame(columns=MASTER_COLS)
    
    # Filter for Fixed Price products if product column exists
    work_df = df
    if c_prod is not None:
        mask_fp = work_df[c_prod].astype(str).str.strip().str.lower().eq('fixed price')
        work_df = work_df[mask_fp]
    
    # Check required fields
    if (c_price is None and c_green is None) or c_term is None or c_start is None:
        return out
    if c_desc is None and (c_zone is None and c_lf is None):
        return out

    # Get the number of rows we'll be working with
    num_rows = len(work_df)
    if num_rows == 0:
        return out

    # Map to master table columns
    out = pd.DataFrame(index=range(num_rows), columns=MASTER_COLS)
    out['ID'] = None  # Will be set during append

    # Dates
    today = date.today()
    out['Price_Date'] = today  # Column B - today's date for all rows
    start_dates = pd.to_datetime(work_df[c_start], errors='coerce')
    out['Date'] = start_dates.dt.date  # Column C - start date from input
    
    # Zone and Load Factor
    if c_desc is not None:
        out['Zone'] = work_df[c_desc].map(parse_zone_from_description)
        out['Load'] = work_df[c_desc].map(parse_load_factor_from_description)
    else:
        out['Zone'] = work_df[c_zone].astype(str) if c_zone is not None else ''
        out['Load'] = work_df[c_lf].astype(str) if c_lf is not None else ''
    
    # REP1 - hardcoded
    out['REP1'] = 'HUDSON'
    
    # Term with filtering
    def term_to_int(v):
        try:
            iv = int(float(v))
            return iv if iv in TARGET_TERMS else None
        except Exception:
            m = re.search(r"(\d+)", str(v))
            if m:
                iv = int(m.group(1))
                return iv if iv in TARGET_TERMS else None
            return None
    
    out['Term'] = work_df[c_term].map(term_to_int)
    
    # Usage tiers
    out['Min_MWh'] = 0
    out['Max_MWh'] = 1000
    
    # Price columns
    price_series = pd.to_numeric(work_df[c_price], errors='coerce') if c_price is not None else pd.Series(dtype='float64')
    if price_series.isna().all() and c_green is not None:
        price_series = pd.to_numeric(work_df[c_green], errors='coerce')
    
    out['Daily_No_Ruc'] = price_series
    out['RUC_Nodal'] = 0
    out['Daily'] = price_series
    
    # Default values
    out['Com_Disc'] = 0
    out['HOA_Disc'] = 0
    out['Broker_Fee'] = 0
    out['Meter_Fee'] = 0
    out['Max_Meters'] = 5
    
    # Keep only rows with valid terms
    out = out[out['Term'].notna()].copy()
    
    return out


def transform_to_base_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transform DataFrame to BASE_COLS format.
    
    Args:
        df: Input DataFrame
        
    Returns:
        DataFrame with BASE_COLS structure
    """
    # This is a simplified transformation - you can expand based on your needs
    out = pd.DataFrame(columns=BASE_COLS)
    
    # Map common columns if they exist
    if 'Start Month' in df.columns:
        out['Start Month'] = df['Start Month']
    if 'State' in df.columns:
        out['State'] = df['State']
    else:
        out['State'] = 'TX'  # Default
    
    if 'Utility' in df.columns:
        out['Utility'] = df['Utility']
    if 'Congestion Zone' in df.columns:
        out['Congestion Zone'] = df['Congestion Zone']
    if 'Load Factor' in df.columns:
        out['Load Factor'] = df['Load Factor']
    if 'Term' in df.columns:
        out['Term'] = df['Term'].map(parse_term_to_int)
    if 'Product' in df.columns:
        out['Product'] = df['Product']
    else:
        out['Product'] = 'Fixed Price'  # Default
    
    if '0-200,000' in df.columns:
        out['0-200,000'] = df['0-200,000']
    
    return out


# Example usage functions
def example_filter_usage():
    """Example of how to use the filter_sheet function."""
    # Create sample data
    data = {
        'Product': ['Fixed Price', 'Variable', 'Fixed Price'],
        'Term': [12, 24, 36],
        'Start Month': ['2024-01-01', '2024-02-01', '2024-03-01'],
        'State': ['TX', 'TX', 'TX'],
        'Congestion Zone': ['HOUSTON', 'NORTH', 'SOUTH'],
        '0-200,000': [75.5, 80.2, 72.1]
    }
    df = pd.DataFrame(data)

    # Apply filter
    filtered = filter_sheet(df)
    print("Filtered data:")
    print(filtered)
    return filtered


def example_backup_usage():
    """Example of how to use the backup functionality."""
    # Example master table path
    master_path = Path("2-copy-reformat/Master-Table.xlsx")

    if master_path.exists():
        print(f"Creating backup of {master_path}")
        backup_path = create_master_table_backup(master_path)
        if backup_path:
            print(f"Backup created successfully: {backup_path}")
            return backup_path
        else:
            print("Backup creation failed")
            return None
    else:
        print(f"Master table not found at {master_path}")
        return None


def safe_append_example(data_df, master_path):
    """
    Example of safely appending data with backup.

    Args:
        data_df: DataFrame to append
        master_path: Path to master table
    """
    master_path = Path(master_path)

    # Create backup before any modifications
    backup_path = create_master_table_backup(master_path)
    if backup_path is None:
        print("Warning: Could not create backup, proceeding anyway...")
    else:
        print(f"Master table backed up to: {backup_path}")

    # Transform data to master format
    transformed_data = transform_to_master_format(data_df)

    if transformed_data.empty:
        print("No data to append after transformation")
        return 0

    print(f"Ready to append {len(transformed_data)} rows to master table")
    # Here you would call your actual append function
    # For example: append_to_master_table(transformed_data, master_path)

    return len(transformed_data)


if __name__ == "__main__":
    print("Transformer module - mimics excel_processor.py transformation logic")
    print("Available functions:")
    print("- create_master_table_backup(master_path)")
    print("- parse_term_to_int(val)")
    print("- filter_sheet(df)")
    print("- transform_to_master_format(df)")
    print("- transform_to_base_format(df)")
    print("- safe_append_example(data_df, master_path)")
    print("\nRunning examples...")
    print("\n1. Filter example:")
    example_filter_usage()
    print("\n2. Backup example:")
    example_backup_usage()
