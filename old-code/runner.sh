#!/bin/bash
python - << 'PY'
import pandas as pd
from pathlib import Path
import excel_reader as xr

input_path = Path('HudsonMatrixPrices08272025020701PM.xlsm')
print(f"Using input: {input_path.resolve()}")

try:
    df = xr.transform_input_to_master_df(input_path)
except Exception as e:
    print('ERROR while transforming:', e)
else:
    print('Shape:', df.shape)
    print('Columns:', list(df.columns))
    print('Head:')
    print(df.head(10))
    print('\nDtypes:')
    print(df.dtypes)
    print('\nValue counts: Term')
    print(df['Term'].value_counts(dropna=False).head(10))
    print('\nValue counts: Zone')
    print(df['Zone'].value_counts(dropna=False).head(10))
    print('\nValue counts: Load')
    print(df['Load'].value_counts(dropna=False).head(10))
PY