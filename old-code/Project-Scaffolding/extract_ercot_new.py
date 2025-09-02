import sys
from pathlib import Path

try:
    import pandas as pd
except Exception as e:
    print('DEPENDENCY_ERROR')
    print('pandas is not available:', e)
    sys.exit(10)


def main():
    """
    Extract all information from ERCOT-new.xlsx and create ERCOT-new-copy.xlsx
    """
    root = Path('.').resolve()
    
    # Look for ERCOT-new.xlsx file
    src_file = root / 'ERCOT-new.xlsx'
    
    if not src_file.exists():
        print('NOT_FOUND')
        print(f'ERCOT-new.xlsx not found in {root}')
        sys.exit(3)

    dst_file = root / 'ERCOT-new-copy.xlsx'

    print(f'Source file: {src_file}')
    print(f'Destination file: {dst_file}')

    # Read all sheets as DataFrames
    try:
        print('Reading Excel file...')
        # sheet_name=None loads all sheets into a dict
        sheets = pd.read_excel(src_file, sheet_name=None)
        print(f'Found {len(sheets)} sheet(s):')
        for sheet_name in sheets.keys():
            print(f'  - {sheet_name}: {sheets[sheet_name].shape[0]} rows, {sheets[sheet_name].shape[1]} columns')
    except Exception as e:
        print('READ_ERROR')
        print(str(e))
        sys.exit(2)

    # Write all sheets to the new file
    try:
        print('Writing to new Excel file...')
        with pd.ExcelWriter(dst_file, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                # Excel sheet names are limited to 31 chars
                safe_name = str(sheet_name)[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
                print(f'  - Copied sheet "{sheet_name}" ({df.shape[0]} rows, {df.shape[1]} columns)')
    except Exception as e:
        print('WRITE_ERROR')
        print(str(e))
        sys.exit(2)

    print('SUCCESS')
    print(f'All data successfully extracted from {src_file.name} to {dst_file.name}')
    
    # Display summary of extracted data
    print('\n=== EXTRACTION SUMMARY ===')
    total_rows = sum(df.shape[0] for df in sheets.values())
    total_cols = sum(df.shape[1] for df in sheets.values())
    print(f'Total sheets: {len(sheets)}')
    print(f'Total rows across all sheets: {total_rows}')
    print(f'Total columns across all sheets: {total_cols}')
    
    # Show first few rows of each sheet for verification
    print('\n=== SAMPLE DATA FROM EACH SHEET ===')
    for sheet_name, df in sheets.items():
        print(f'\nSheet: {sheet_name}')
        print(f'Columns: {list(df.columns)}')
        if not df.empty:
            print('First 3 rows:')
            print(df.head(3).to_string())
        else:
            print('(Empty sheet)')


if __name__ == '__main__':
    main()
