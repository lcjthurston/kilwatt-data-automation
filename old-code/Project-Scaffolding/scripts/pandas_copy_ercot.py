import sys
from pathlib import Path

try:
    import pandas as pd
except Exception as e:
    print('DEPENDENCY_ERROR')
    print('pandas is not available:', e)
    sys.exit(10)


def main():
    root = Path('.').resolve()
    matches = list(root.rglob('ERCOT.xlsx'))
    if not matches:
        print('NOT_FOUND')
        sys.exit(3)

    src = matches[0]
    dst = src.with_name('ERCOT-copy.xlsx')

    # Read all sheets as DataFrames
    try:
        # sheet_name=None loads all sheets into a dict
        sheets = pd.read_excel(src, sheet_name=None)
    except Exception as e:
        print('READ_ERROR')
        print(str(e))
        sys.exit(2)

    try:
        with pd.ExcelWriter(dst, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                # Excel sheet names are limited to 31 chars
                safe_name = str(sheet_name)[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
    except Exception as e:
        print('WRITE_ERROR')
        print(str(e))
        sys.exit(2)

    print('COPIED')
    print(str(src))
    print(str(dst))


if __name__ == '__main__':
    main()

