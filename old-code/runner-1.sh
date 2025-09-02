python - << 'PY'
from pathlib import Path
import pandas as pd
import excel_reader as xr
from excel_processor import write_updated_master_copy

input_path = Path('HudsonMatrixPrices08272025020701PM.xlsm')
master_dir = Path('2-copy-reformat')
master_filename = 'Master-Table.xlsx'
out_filename = 'master-file-updated.xlsx'

print('Input file exists:', input_path.exists(), input_path)
print('Master file exists:', (master_dir / master_filename).exists(), master_dir / master_filename)

try:
    df = xr.transform_input_to_master_df(input_path)
    print('Transformed shape:', df.shape)
    print('Sample rows:')
    print(df.head(5))
except Exception as e:
    print('ERROR during transform:', e)
    raise

try:
    out_path = write_updated_master_copy(df, master_dir=master_dir, master_filename=master_filename, out_filename=out_filename)
    exists = Path(out_path).exists()
    print('Output written:', out_path, 'exists=', exists)
except Exception as e:
    print('ERROR during write:', e)
    raise
PY