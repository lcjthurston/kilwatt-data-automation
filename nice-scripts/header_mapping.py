import pandas as pd
from pathlib import Path
p = Path('HDA/DAILY PRICING - HUDSON - TEMPLATE - 2021 (1).xlsx')
xls = pd.ExcelFile(p)
print('SHEETS:', xls.sheet_names)
# Try to read IMPORT sheet
sheet = None
for s in xls.sheet_names:
    if str(s).strip().lower() == 'import':
        sheet = s
        break
print('IMPORT SHEET:', sheet)
if sheet:
    df = pd.read_excel(p, sheet_name=sheet, header=0)
    print('IMPORT HEADERS:', list(df.columns))
    print('IMPORT SAMPLE ROWS:', df.head(3).to_dict('records'))