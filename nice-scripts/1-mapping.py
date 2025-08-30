import pandas as pd
from pathlib import Path
p = Path('1-original-excel-data/DAILY PRICING - new.xlsx')
df = pd.read_excel(p, sheet_name=0)
print('COLS:', list(df.columns))
print(df.head(2).to_dict('records'))