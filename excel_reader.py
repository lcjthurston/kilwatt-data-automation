from __future__ import annotations
import os
from datetime import datetime
try:
    import requests
    from dotenv import load_dotenv
    from graph_auth import acquire_graph_token
    load_dotenv()
except Exception:
    # Allow excel_reader to be imported even if SharePoint deps are missing
    requests = None
    acquire_graph_token = None

import re
from datetime import date as date_today
from pathlib import Path
from typing import Optional, Dict, List

import pandas as pd

# Master table schema (17 columns)
MASTER_COLS: List[str] = [
    'ID', 'Price_Date', 'Date', 'Zone', 'Load', 'REP1', 'Term', 'Min_MWh', 'Max_MWh',
    'Daily_No_Ruc', 'RUC_Nodal', 'Daily', 'Com_Disc', 'HOA_Disc', 'Broker_Fee', 'Meter_Fee', 'Max_Meters'
]

# Allowed contract terms
TARGET_TERMS = {12, 24, 36, 48, 60}


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())


def _find(norm_map: Dict[str, str], candidates: List[str]) -> Optional[str]:
    for key in candidates:
        if key in norm_map:
            return norm_map[key]
    return None


def _max_id_from_master(master_path: Optional[Path]) -> int:
    if not master_path:
        return 0
    p = Path(master_path)
    if not p.exists():
        return 0
    try:
        df = pd.read_excel(p, usecols=[0])  # Column A expected to be ID
        if df.empty:
            return 0
        ser = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        ser = ser.dropna()
        return int(ser.max()) if not ser.empty else 0
    except Exception:
        return 0


def _parse_zone_from_col_e(val: str) -> str:
    """Extract the word that precedes the word 'zone' in Column E and map per spec.
    If no such pattern exists, return 'NA'.
    Mapping: North->NORTH, West->WEST, South->SOUTH, Houston->COAST, else NA.
    """
    s = str(val or '').strip()
    if not s:
        return 'NA'
    # Strictly require '<word> zone' pattern
    m = re.search(r"\b([A-Za-z]+)\s+zone\b", s, flags=re.IGNORECASE)
    if not m:
        return 'NA'
    w = m.group(1).strip().lower()
    if w == 'north':
        return 'NORTH'
    if w == 'west':
        return 'WEST'
    if w == 'south':
        return 'SOUTH'
    if w == 'houston':
        return 'COAST'
    return 'NA'


def _parse_load_from_col_e(val: str) -> str:
    """Extract HIGH/MED/LOW from Column E text; else NA."""
    s = str(val or '').lower()
    if 'high' in s:
        return 'HIGH'
    if 'med' in s:
        return 'MED'
    if 'low' in s:
        return 'LOW'
    return 'NA'


def transform_input_to_master_df(
    input_path: Path | str,
    *,
    master_path: Optional[Path | str] = None,
    start_id: Optional[int] = None,
    multiply_price_by_100: bool = False,
) -> pd.DataFrame:
    """Read an input Excel file and return a DataFrame in master table format.

    Implements Program Specification.txt:
    - 17 columns in master schema
    - Column A (ID) sequential
    - Column B (Price_Date) = today
    - Column C (Date) from input Column J
    - Column D (Zone) parsed from input Column E per mapping
    - Column E (Load) parsed from input Column E (HIGH/MED/LOW/NA)
    - Column F (REP1) = 'HUDSON'
    - Column G (Term) from input Column D, filtered to {12,24,36,48,60}
    - Column H (Min_MWh) = 0
    - Column I (Max_MWh) = 1000
    - Column J (Daily_No_Ruc) = derived price (if available) multiplied by 10, else 0
    Remaining columns defaulted as before.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input Excel not found: {input_path}")

    # Read all sheets and concatenate rows
    sheets = pd.read_excel(input_path, sheet_name=None)
    if not sheets:
        return pd.DataFrame(columns=MASTER_COLS)

    frames: List[pd.DataFrame] = []

    for _, df in sheets.items():
        if df is None or df.empty:
            continue

        work_df = df.copy()

        # Resolve column names by absolute positions when available
        def col_at(idx: int) -> Optional[str]:
            return work_df.columns[idx] if idx < work_df.shape[1] else None

        cE = col_at(4)   # Column E: zone+load descriptor
        cD = col_at(3)   # Column D: term
        cJ = col_at(9)   # Column J: start date
        cH = col_at(7)   # Column H: source for Daily_No_Ruc

        # Build a normalized map for optional filters
        raw_cols = list(work_df.columns)
        norm_map = {_norm(c): c for c in raw_cols}

        # Product filter if present (kept for safety)
        c_prod = _find(norm_map, ['product'])
        if c_prod is not None:
            mask_fp = work_df[c_prod].astype(str).str.strip().str.lower().eq('fixed price')
            work_df = work_df[mask_fp]
        if work_df.empty:
            continue

        # Column G: Term from Column D
        def term_to_int(v):
            try:
                iv = int(float(v))
                return iv if iv in TARGET_TERMS else None
            except Exception:
                m = re.search(r"(\d+)", str(v))
                if m:
                    iv = int(m.group(1))
                    return iv if iv in TARGET_TERMS else None
                return None

        if cD is not None:
            terms = work_df[cD].map(term_to_int)
        else:
            terms = pd.Series([None] * len(work_df), index=work_df.index)

        # Column C: Start Date from Column J
        if cJ is not None:
            start_dates = pd.to_datetime(work_df[cJ], errors='coerce').dt.date
        else:
            start_dates = pd.Series([pd.NaT] * len(work_df), index=work_df.index)

        # Column D/E: Zone and Load from Column E text
        if cE is not None:
            zone_series = work_df[cE].map(_parse_zone_from_col_e)
            load_series = work_df[cE].map(_parse_load_from_col_e)
        else:
            zone_series = pd.Series(['NA'] * len(work_df), index=work_df.index)
            load_series = pd.Series(['NA'] * len(work_df), index=work_df.index)

        # Column J: Daily_No_Ruc = 1000 * Column H of input
        if cH is not None:
            base_h = pd.to_numeric(work_df[cH], errors='coerce').fillna(0.0)
            daily_no_ruc = (base_h * 1000.0)
        else:
            daily_no_ruc = pd.Series([0.0] * len(work_df), index=work_df.index, dtype='float64')

        # Build output slice
        out = pd.DataFrame(index=work_df.index, columns=MASTER_COLS)
        out['ID'] = None  # set later
        out['Price_Date'] = date_today.today()
        out['Date'] = start_dates
        out['Zone'] = zone_series
        out['Load'] = load_series
        out['REP1'] = 'HUDSON'
        out['Term'] = terms
        out['Min_MWh'] = 0
        out['Max_MWh'] = 1000
        out['Daily_No_Ruc'] = daily_no_ruc
        # Column K: always $0.00 => numeric 0.00
        out['RUC_Nodal'] = 0.0
        # Column L: sum of Columns J and K (Daily_No_Ruc + RUC_Nodal)
        out['Daily'] = (
            pd.to_numeric(out['Daily_No_Ruc'], errors='coerce').fillna(0.0)
            + pd.to_numeric(out['RUC_Nodal'], errors='coerce').fillna(0.0)
        )
        # Columns M..P: always $0.00 => numeric 0.00
        out['Com_Disc'] = 0.0
        out['HOA_Disc'] = 0.0
        out['Broker_Fee'] = 0.0
        out['Meter_Fee'] = 0.0
        # Column Q: always 5
        out['Max_Meters'] = 5

        # Keep only rows with valid terms
        out = out[out['Term'].notna()].copy()
        if not out.empty:
            frames.append(out)

    if not frames:
        return pd.DataFrame(columns=MASTER_COLS)

    dest = pd.concat(frames, ignore_index=True)

    # Assign sequential ID
    base = start_id if start_id is not None else (_max_id_from_master(Path(master_path)) + 1 if master_path else 1)
    dest['ID'] = range(base, base + len(dest))

    # Ensure numeric dtypes where sensible
    for col in ['Term', 'Min_MWh', 'Max_MWh', 'Daily_No_Ruc', 'RUC_Nodal', 'Daily', 'Com_Disc', 'HOA_Disc', 'Broker_Fee', 'Meter_Fee', 'Max_Meters']:
        if col in dest.columns:
            dest[col] = pd.to_numeric(dest[col], errors='coerce')

# --- SharePoint helpers (download master table with rename-on-exist) ---

def _download_sharepoint_file(file_name: str, download_path: Path | str) -> Optional[Path]:
    """Download a file from SharePoint using Microsoft Graph.
    Mirrors excel_processor.download_sharepoint_file but scoped here.
    """
    if requests is None or acquire_graph_token is None:
        print("SharePoint dependencies not available. Install requests, python-dotenv, and configure graph_auth.")
        return None

    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")
    sharepoint_folder = os.getenv("SHAREPOINT_UPLOAD_FOLDER")

    if not all([tenant_id, client_id, client_secret, site_hostname, site_path, sharepoint_folder]):
        print("ERROR: Missing SharePoint configuration in .env (TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_HOSTNAME, SITE_PATH, SHAREPOINT_UPLOAD_FOLDER)")
        return None

    download_path = Path(download_path)
    try:
        token = acquire_graph_token(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token['access_token']}"}

        # Get site id
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:/sites/{site_path}"
        site_resp = requests.get(site_url, headers=headers)
        site_resp.raise_for_status()
        site_id = site_resp.json()['id']

        # Get drive id
        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_resp = requests.get(drive_url, headers=headers)
        drive_resp.raise_for_status()
        drive_id = drive_resp.json()['id']

        # Build download URL
        file_path = f"{sharepoint_folder}/{file_name}".replace('//','/')
        if file_path.startswith('/'):
            file_path = file_path[1:]
        file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content"

        resp = requests.get(file_url, headers=headers)
        if resp.status_code == 200:
            download_path.parent.mkdir(parents=True, exist_ok=True)
            with open(download_path, 'wb') as f:
                f.write(resp.content)
            print(f"Downloaded SharePoint file to: {download_path}")
            return download_path
        else:
            print(f"Failed to download. Status: {resp.status_code}\n{resp.text}")
            return None
    except Exception as e:
        print(f"Error during SharePoint download: {e}")
        return None


def ensure_master_table_downloaded(dest_dir: Path | str,
                                   master_file_name: str = 'Master-Table.xlsx',
                                   parent_folder_override: Optional[str] = None) -> Optional[Path]:
    """Ensure the master table is downloaded into dest_dir.
    - If a file with that name already exists, rename it by appending a timestamp.
    - Then download fresh from SharePoint.
    - If parent_folder_override is provided, temporarily override SHAREPOINT_UPLOAD_FOLDER.
    Returns the path to the downloaded file, or None on failure.
    """
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    out_path = dest_dir / master_file_name

    # If file exists, rename it with date-time suffix
    if out_path.exists():
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        renamed = out_path.with_name(f"{out_path.stem}_{ts}{out_path.suffix}")
        out_path.rename(renamed)
        print(f"Existing master renamed to: {renamed}")

    original_folder = os.getenv("SHAREPOINT_UPLOAD_FOLDER")
    try:
        if parent_folder_override:
            os.environ["SHAREPOINT_UPLOAD_FOLDER"] = parent_folder_override
        return _download_sharepoint_file(master_file_name, out_path)
    finally:
        if parent_folder_override is not None and original_folder is not None:
            os.environ["SHAREPOINT_UPLOAD_FOLDER"] = original_folder


if __name__ == "__main__":
    # Minimal manual test: read an input file path and print the transformed shape
    # Example: python excel_reader.py
    sample = None  # Set to a local .xlsx path if you want to quick-test
    if sample:
        df_master = transform_input_to_master_df(sample)
        print(df_master.head())
        print(df_master.dtypes)
        print(f"Rows: {len(df_master)} | Columns: {list(df_master.columns)}")
