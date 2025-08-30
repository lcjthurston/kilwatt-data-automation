import sys
from pathlib import Path
import re

try:
    import pandas as pd
except Exception as e:
    print('DEPENDENCY_ERROR')
    print('pandas is not available:', e)
    sys.exit(10)

TARGET_TERMS = {12, 24, 36, 48, 60}


def find_col(df, preferred_names):
    """Find a column by a set of preferred names; fall back to contains match."""
    norm_map = {col: str(col).strip().lower() for col in df.columns}
    # exact match
    for orig, norm in norm_map.items():
        if norm in preferred_names:
            return orig
    # contains match
    for orig, norm in norm_map.items():
        if any(name in norm for name in preferred_names):
            return orig
    return None


def parse_term_to_int(val):
    """Try to extract an integer month value from a cell (e.g., 12, '12', '12 Months', '12 mo')."""
    if pd.isna(val):
        return None
    # If already an int-like
    if isinstance(val, (int,)):
        return val
    # If float and integral
    if isinstance(val, float):
        if pd.isna(val):
            return None
        ival = int(val)
        return ival if float(ival) == val else None
    s = str(val).strip()
    # direct numeric string
    if s.isdigit():
        try:
            return int(s)
        except Exception:
            return None
    # extract first number
    m = re.search(r"(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def main():
    root = Path('.').resolve()
    src = root / 'ERCOT-new.xlsx'
    if not src.exists():
        print('NOT_FOUND')
        print(f'ERCOT-new.xlsx not found in {root}')
        sys.exit(3)

    dst = root / 'ERCOT-new-product-term.xlsx'

    try:
        sheets = pd.read_excel(src, sheet_name=None)
    except Exception as e:
        print('READ_ERROR')
        print(str(e))
        sys.exit(2)

    summary = []
    kept_any = False

    try:
        with pd.ExcelWriter(dst, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                total = len(df)
                # Find relevant columns
                prod_col = find_col(df, {"product", "products"})
                term_col = find_col(df, {"term", "terms"})

                if prod_col is None or term_col is None:
                    filtered = df.iloc[0:0]
                    kept = 0
                    note = (
                        f"Missing column(s): "
                        f"{'Product' if prod_col is None else ''}"
                        f"{' and ' if (prod_col is None and term_col is None) else ''}"
                        f"{'Term' if term_col is None else ''}"
                    )
                else:
                    # Product mask
                    prod_mask = (
                        df[prod_col]
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .eq('fixed price')
                    )

                    # Term mask: map values to integers and check membership
                    term_vals = df[term_col].map(parse_term_to_int)
                    term_mask = term_vals.isin(TARGET_TERMS)

                    mask = prod_mask & term_mask
                    filtered = df[mask]
                    kept = len(filtered)
                    kept_any = kept_any or kept > 0
                    note = f'Filtered on columns "{prod_col}" and "{term_col}"'

                safe_name = str(sheet_name)[:31]
                filtered.to_excel(writer, sheet_name=safe_name, index=False)
                summary.append((sheet_name, total, kept, note))
    except Exception as e:
        print('WRITE_ERROR')
        print(str(e))
        sys.exit(2)

    print('SUCCESS')
    print('Applied filters: Product == "Fixed Price" and Term in {12,24,36,48,60} (robust to strings like "12 Months").')
    print(f'Output: {dst.name}')
    for sheet_name, total, kept, note in summary:
        print(f'- {sheet_name}: kept {kept} of {total} rows | {note}')

    if not kept_any:
        print('WARNING: No rows matched the combined filters across all sheets.')


if __name__ == '__main__':
    main()

