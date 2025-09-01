python - << 'PY'
from pathlib import Path
import sys

print('CWD:', Path.cwd())

try:
    import excel_reader as xr
    from excel_processor import write_updated_master_copy
except Exception as e:
    print('IMPORT_ERROR:', e)
    sys.exit(2)

# 1) Ensure master table downloaded (rename existing if present)
master_dir = Path('2-copy-reformat')
master_name = 'Master-Table.xlsx'
print('Ensuring master table in', master_dir/master_name)
try:
    downloaded = xr.ensure_master_table_downloaded(master_dir, master_file_name=master_name,
                                                   parent_folder_override='/Kilowatt/Client Pricing Sheets')
    print('Downloaded/Ensured path:', downloaded)
except Exception as e:
    print('ERROR ensure_master_table_downloaded:', e)
    downloaded = master_dir / master_name

# 2) Transform Hudson input
input_path = Path('HudsonMatrixPrices08272025020701PM.xlsm')
print('Input exists?', input_path.exists(), input_path)

try:
    df = xr.transform_input_to_master_df(input_path, master_path=downloaded)
    print('Transformed rows/cols:', df.shape)
except Exception as e:
    print('ERROR transform_input_to_master_df:', e)
    raise

# 3) Write updated master copy without modifying original
try:
    out_path = write_updated_master_copy(df, master_dir=master_dir, master_filename=master_name,
                                         out_filename='master-file-updated.xlsx')
    print('Output written:', out_path, 'exists=', Path(out_path).exists())
except Exception as e:
    print('ERROR write_updated_master_copy:', e)
    raise
PY