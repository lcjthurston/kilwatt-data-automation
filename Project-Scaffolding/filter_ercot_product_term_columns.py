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
DESIRED_COLUMNS = [
    'Start Month',
    'State',
    'Utility',
    'Congestion Zone',
    'Load Factor',
    'Term',
    'Product',
    '0-200,000',
]


def find_col(df, preferred_names):
    """Find a column by a set of preferred names; fall back to contains match (case-insensitive)."""
    norm_map = {col: str(col).strip().lower() for col in df.columns}
    preferred_lower = {n.lower() for n in preferred_names}
    # exact match
    for orig, norm in norm_map.items():
        if norm in preferred_lower:
            return orig
    # contains match
    for orig, norm in norm_map.items():
        if any(name in norm for name in preferred_lower):
            return orig
    return None


def parse_term_to_int(val):
    """Extract an integer month value from a cell (e.g., 12, '12', '12 Months', '12 mo')."""
    if pd.isna(val):
        return None
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        if pd.isna(val):
            return None
        ival = int(val)
        return ival if float(ival) == val else None
    s = str(val).strip()
    if s.isdigit():
        try:
            return int(s)
        except Exception:
            return None
    m = re.search(r"(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def select_columns(df, desired_names):
    """Return list of actual DataFrame columns corresponding to desired_names, preserving order.
    Matches are case-insensitive by exact name first, then contains as fallback.
    """
    selected = []
    norm_map = {str(col).strip().lower(): col for col in df.columns}
    used = set()
    # exact case-insensitive matches first
    for name in desired_names:
        key = name.strip().lower()
        if key in norm_map:
            col = norm_map[key]
            selected.append(col)
            used.add(col)
        else:
            # fallback contains search
            found = None
            for col in df.columns:
                if col in used:
                    continue
                if key in str(col).strip().lower():
                    found = col
                    break
            if found is not None:
                selected.append(found)
                used.add(found)
            else:
                # not found, skip
                pass
    return selected


def main():
    root = Path('.').resolve()
    src = root / 'ERCOT-new.xlsx'
    if not src.exists():
        print('NOT_FOUND')
        print(f'ERCOT-new.xlsx not found in {root}')
        sys.exit(3)

    dst = root / 'ERCOT-new-product-term-selected-columns.xlsx'

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
                # Identify filter columns
                prod_col = find_col(df, {"Product", "Products"})
                term_col = find_col(df, {"Term", "Terms"})

                if prod_col is None or term_col is None:
                    filtered = df.iloc[0:0]
                    kept = 0
                    note = (
                        f"Missing column(s): "
                        f"{'Product' if prod_col is None else ''}"
                        f"{' and ' if (prod_col is None and term_col is None) else ''}"
                        f"{'Term' if term_col is None else ''}"
                    )
                    selected_cols = []
                else:
                    # Apply filters
                    prod_mask = (
                        df[prod_col].astype(str).str.strip().str.lower() == 'fixed price'
                    )
                    term_vals = df[term_col].map(parse_term_to_int)
                    term_mask = term_vals.isin(TARGET_TERMS)
                    mask = prod_mask & term_mask

                    filtered = df[mask]
                    kept = len(filtered)
                    kept_any = kept_any or kept > 0
                    note = f'Filtered on columns "{prod_col}" and "{term_col}"'

                    # Select only desired columns (that exist), in requested order
                    selected_cols = select_columns(filtered, DESIRED_COLUMNS)
                    filtered = filtered[selected_cols]

                safe_name = str(sheet_name)[:31]
                filtered.to_excel(writer, sheet_name=safe_name, index=False)
                summary.append((sheet_name, total, kept, note, selected_cols))
    except Exception as e:
        print('WRITE_ERROR')
        print(str(e))
        sys.exit(2)

    print('SUCCESS')
    print('Applied filters: Product == "Fixed Price" and Term in {12,24,36,48,60}.')
    print(f'Output: {dst.name}')
    for sheet_name, total, kept, note, cols in summary:
        col_list = ', '.join(map(str, cols)) if cols else '(no columns)'
        print(f'- {sheet_name}: kept {kept} of {total} rows | {note} | columns: {col_list}')

    if not kept_any:
        print('WARNING: No rows matched the combined filters across all sheets.')


if __name__ == '__main__':
    main()

