import sys
from pathlib import Path

try:
    import pandas as pd
except Exception as e:
    print('DEPENDENCY_ERROR')
    print('pandas is not available:', e)
    sys.exit(10)


def find_product_column(df):
    """Return the column name that represents Product/Products, or None if not found."""
    if df is None or df.empty:
        return None
    norm_map = {col: str(col).strip().lower() for col in df.columns}
    # Try exact matches first
    for orig, norm in norm_map.items():
        if norm in ("product", "products"):
            return orig
    # Fallback: any column that contains the word 'product'
    for orig, norm in norm_map.items():
        if "product" in norm:
            return orig
    return None


def main():
    root = Path('.').resolve()
    src = root / 'ERCOT-new.xlsx'
    if not src.exists():
        print('NOT_FOUND')
        print(f'ERCOT-new.xlsx not found in {root}')
        sys.exit(3)

    dst = root / 'ERCOT-new-fixed-price.xlsx'

    try:
        sheets = pd.read_excel(src, sheet_name=None)
    except Exception as e:
        print('READ_ERROR')
        print(str(e))
        sys.exit(2)

    kept_any = False
    summary = []

    try:
        with pd.ExcelWriter(dst, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                total = len(df)
                prod_col = find_product_column(df)
                if prod_col is None:
                    # No product column: write empty sheet with original columns
                    filtered = df.iloc[0:0]
                    kept = 0
                    note = 'No Product/Products column found'
                else:
                    series = df[prod_col].astype(str).str.strip().str.lower()
                    mask = series == 'fixed price'
                    filtered = df[mask]
                    kept = len(filtered)
                    kept_any = kept_any or kept > 0
                    note = f'Filtered on column "{prod_col}"'

                safe_name = str(sheet_name)[:31]
                filtered.to_excel(writer, sheet_name=safe_name, index=False)
                summary.append((sheet_name, total, kept, note))
    except Exception as e:
        print('WRITE_ERROR')
        print(str(e))
        sys.exit(2)

    print('SUCCESS')
    print(f'Filtered rows with Product == "Fixed Price" (case-insensitive) into: {dst.name}')
    for sheet_name, total, kept, note in summary:
        print(f'- {sheet_name}: kept {kept} of {total} rows | {note}')

    if not kept_any:
        print('WARNING: No rows matched "Fixed Price" across all sheets.')


if __name__ == '__main__':
    main()

