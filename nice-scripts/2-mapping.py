import pandas as pd, re
from pathlib import Path

def parse_zone_and_load(desc: str):
    s = str(desc or '').strip()
    load = None
    for token in [' Low Load Factor',' Medium Load Factor',' High Load Factor']:
        if s.endswith(token):
            base = s[:-len(token)]
            load = token.strip().split()[0].upper()  # LOW/MEDIUM/HIGH
            return base, load
    return s, ''

src_xlsm = Path('HDA/HudsonMatrixPrices08272025020701PM.xlsm')
master_path = Path('1-original-excel-data/DAILY PRICING - new.xlsx')

# Read hudson Matrix Table
all_sheets = pd.read_excel(src_xlsm, sheet_name=None)
mt_name = next(n for n in all_sheets.keys() if str(n).strip().lower() == 'matrix table')
df = all_sheets[mt_name]

# Build mapped df for master B..Q
out = pd.DataFrame()
out['Price_Date'] = pd.to_datetime(df['CreatedDate'], errors='coerce').dt.date
out['Date'] = pd.to_datetime(df['StartDate'], errors='coerce').dt.date
zone, load = zip(*df['MatrixDescription'].map(parse_zone_and_load))
out['Zone'] = [z.strip().upper() for z in zone]
out['REP1'] = 'HUDSON'
out['Load'] = [l if l else '' for l in load]
out['Term'] = pd.to_numeric(df['TermCode'], errors='coerce').astype('Int64')
out['Min_MWh'] = 0
out['Max_MWh'] = 1000
# price in cents to $/MWh (x10)
out['Daily_No_Ruc'] = pd.to_numeric(df['Price'], errors='coerce') * 10
out['RUC_Nodal'] = 0
out['Daily'] = out['Daily_No_Ruc']
out['Com_Disc'] = 0
out['HOA_Disc'] = 0
out['Broker_Fee'] = 0
out['Meter_Fee'] = 0
out['Max_Meters'] = 5

# Read master
m = pd.read_excel(master_path)
# Compute next IDs
a = m['ID']
next_id = (a.max() if len(a)>0 else 0) + 1
out.insert(0,'ID', range(next_id, next_id + len(out)))

# Concat and write to temp file
final_df = pd.concat([m, out], ignore_index=True)
print('About to write rows:', len(out), '-> final rows:', len(final_df))
final_df.to_excel('1-original-excel-data/DAILY PRICING - new.xlsx', index=False)
print('WROTE')