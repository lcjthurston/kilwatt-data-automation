import pandas as pd
from pathlib import Path

def read_excel_file(file_path: Path) -> pd.DataFrame:
    """Read an Excel file into a pandas DataFrame."""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()  # Return an empty DataFrame on error

def main():
    file_path = Path('2-copy-reformat/Master-Table.xlsx')
    df = read_excel_file(file_path)
    print(df.head())

if __name__ == "__main__":
    main()
