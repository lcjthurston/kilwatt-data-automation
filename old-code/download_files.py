#!/usr/bin/env python3
"""
Download SharePoint files script + chain transform/append/upload.

- Downloads the master table file "DAILY PRICING - new.xlsx" and the Hudson input file
  "HudsonMatrixPrices08272025020701PM.xlsm" from SharePoint into the new_files directory.
- Transforms the Hudson input into the 17-col master schema and writes an updated master copy
  (master-file-updated.xlsx) in new_files WITHOUT modifying the downloaded master file.
- Uploads master-file-updated.xlsx back to SharePoint (default folder: /Kilowatt/Client Pricing Sheets).

Usage:
    python download_files.py
"""

import os
import sys
from pathlib import Path
from datetime import datetime

try:
    import requests
    from dotenv import load_dotenv
    from graph_auth import acquire_graph_token
    import pandas as pd
    import excel_reader as xr
    from excel_processor import write_updated_master_copy
except ImportError as e:
    print('DEPENDENCY_ERROR: Required libraries are not installed.')
    print('Please install: pip install requests python-dotenv msal pandas openpyxl')
    print(e)
    sys.exit(1)

# Load environment variables
load_dotenv()


def download_sharepoint_file(file_name: str, download_path: Path, sharepoint_folder_override: str = None) -> bool:
    """Download a file from SharePoint using Microsoft Graph API.

    Args:
        file_name: Name of the file to download from SharePoint
        download_path: Local path where to save the file
        sharepoint_folder_override: Optional override for SHAREPOINT_UPLOAD_FOLDER

    Returns:
        bool: True if successful, False otherwise
    """
    # Get credentials from environment
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")
    sharepoint_folder = sharepoint_folder_override or os.getenv("SHAREPOINT_UPLOAD_FOLDER")

    if not all([tenant_id, client_id, client_secret, site_hostname, site_path, sharepoint_folder]):
        print("ERROR: Missing SharePoint configuration in environment variables.")
        print("Required: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_HOSTNAME, SITE_PATH, SHAREPOINT_UPLOAD_FOLDER")
        return False

    try:
        print(f"Downloading {file_name} from SharePoint...")

        # Acquire access token
        token = acquire_graph_token(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token['access_token']}"}

        # Get the site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:/sites/{site_path}"
        site_response = requests.get(site_url, headers=headers)

        if site_response.status_code != 200:
            print(f"Failed to get site info. Status: {site_response.status_code}")
            print(f"Response: {site_response.text}")
            return False

        site_data = site_response.json()
        site_id = site_data['id']
        print(f"Found SharePoint site ID: {site_id}")

        # Get the default drive (document library)
        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_response = requests.get(drive_url, headers=headers)

        if drive_response.status_code != 200:
            print(f"Failed to get drive info. Status: {drive_response.status_code}")
            print(f"Response: {drive_response.text}")
            return False

        drive_data = drive_response.json()
        drive_id = drive_data['id']
        print(f"Found drive ID: {drive_id}")

        # Construct the file path and download URL
        file_path = f"{sharepoint_folder}/{file_name}".replace('//', '/')
        if file_path.startswith('/'):
            file_path = file_path[1:]  # Remove leading slash

        file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content"
        print(f"Downloading from: {file_url}")

        # Download the file
        response = requests.get(file_url, headers=headers)

        if response.status_code == 200:
            # Ensure directory exists
            download_path.parent.mkdir(parents=True, exist_ok=True)

            with open(download_path, "wb") as local_file:
                local_file.write(response.content)
            print(f"Successfully downloaded: {download_path}")
            return True
        else:
            print(f"Failed to download file. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False

    except Exception as e:
        print(f"Error during SharePoint download: {str(e)}")
        return False


def upload_sharepoint_file(local_path: Path, remote_name: str, sharepoint_folder_override: str = None) -> bool:
    """Upload a local file to SharePoint (overwrites if exists).
    Default folder is SHAREPOINT_UPLOAD_FOLDER unless overridden.
    """
    # Get credentials from environment
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")
    sharepoint_folder = sharepoint_folder_override or os.getenv("SHAREPOINT_UPLOAD_FOLDER")

    if not all([tenant_id, client_id, client_secret, site_hostname, site_path, sharepoint_folder]):
        print("ERROR: Missing SharePoint configuration in environment variables.")
        return False

    if not Path(local_path).exists():
        print(f"ERROR: Local file not found: {local_path}")
        return False

    try:
        print(f"Uploading {local_path} to SharePoint folder: {sharepoint_folder} as {remote_name} ...")
        token = acquire_graph_token(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token['access_token']}", "Content-Type": "application/octet-stream"}

        # Resolve site/drive
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:/sites/{site_path}"
        site_response = requests.get(site_url, headers=headers)
        if site_response.status_code != 200:
            print(f"Failed to get site info. Status: {site_response.status_code}")
            print(f"Response: {site_response.text}")
            return False
        site_id = site_response.json()['id']

        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_response = requests.get(drive_url, headers=headers)
        if drive_response.status_code != 200:
            print(f"Failed to get drive info. Status: {drive_response.status_code}")
            print(f"Response: {drive_response.text}")
            return False
        drive_id = drive_response.json()['id']

        # Build upload path and PUT
        remote_path = f"{sharepoint_folder}/{remote_name}".replace('//','/')
        if remote_path.startswith('/'):
            remote_path = remote_path[1:]
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{remote_path}:/content"

        with open(local_path, 'rb') as f:
            data = f.read()
        resp = requests.put(upload_url, headers=headers, data=data)
        if resp.status_code in (200, 201):
            print(f"Successfully uploaded to: {remote_path}")
            return True
        else:
            print(f"Failed to upload. Status: {resp.status_code}\n{resp.text}")
            return False
    except Exception as e:
        print(f"Error during SharePoint upload: {e}")
        return False


def rename_existing_file(file_path: Path) -> None:
    """Rename existing file by appending timestamp if it exists."""
    if file_path.exists():
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        renamed_path = file_path.with_name(f"{file_path.stem}_{timestamp}{file_path.suffix}")
        file_path.rename(renamed_path)
        print(f"Existing file renamed to: {renamed_path}")


def main():
    """Main function to download both files from SharePoint and chain transform/append/upload."""
    # Define target directory
    new_files_dir = Path("new_files")

    # Define files to download
    master_filename_remote = "DAILY PRICING - new.xlsx"
    hudson_filename_remote = "HudsonMatrixPrices08272025020701PM.xlsm"

    files_to_download = [
        {
            "name": master_filename_remote,
            "local_name": master_filename_remote,
            "folder_override": "/Kilowatt/Client Pricing Sheets"  # Master table location
        },
        {
            "name": hudson_filename_remote,
            "local_name": hudson_filename_remote,
            "folder_override": None  # Use default SHAREPOINT_UPLOAD_FOLDER
        }
    ]

    print(f"Starting download process to: {new_files_dir.absolute()}")
    print("=" * 60)

    success_count = 0

    for file_info in files_to_download:
        file_name = file_info["name"]
        local_name = file_info["local_name"]
        folder_override = file_info["folder_override"]

        # Define local path
        local_path = new_files_dir / local_name

        # Rename existing file if present
        rename_existing_file(local_path)

        # Download the file
        print(f"\nDownloading: {file_name}")
        if download_sharepoint_file(file_name, local_path, folder_override):
            success_count += 1
            print(f"✓ Successfully downloaded: {local_path}")
        else:
            print(f"✗ Failed to download: {file_name}")

        print("-" * 40)

    print(f"\nDownload Summary:")
    print(f"Successfully downloaded: {success_count}/{len(files_to_download)} files")
    print(f"Files saved to: {new_files_dir.absolute()}")

    if success_count != len(files_to_download):
        print("Some downloads failed. Check the output above for details.")
        return 1

    # Chain: transform Hudson and write updated master copy (non-destructive)
    master_local_path = new_files_dir / master_filename_remote
    hudson_local_path = new_files_dir / hudson_filename_remote

    print("\nStarting transformation of Hudson input...")
    try:
        df_master = xr.transform_input_to_master_df(hudson_local_path, master_path=master_local_path)
        print(f"Transformed DataFrame shape: {df_master.shape}")
    except Exception as e:
        print("ERROR during transform_input_to_master_df:", e)
        return 2

    # Write updated copy (does not modify the downloaded master file)
    try:
        out_path = write_updated_master_copy(
            df_master,
            master_dir=new_files_dir,
            master_filename=master_filename_remote,
            out_filename='master-file-updated.xlsx'
        )
        print(f"Updated master copy written to: {out_path}")
    except Exception as e:
        print("ERROR during write_updated_master_copy:", e)
        return 3

    # Upload the updated master copy back to SharePoint (default to parent folder)
    upload_folder = os.getenv('MASTER_UPLOAD_FOLDER', '/Kilowatt/Client Pricing Sheets')
    print(f"\nUploading updated master to SharePoint folder: {upload_folder}")
    try:
        uploaded = upload_sharepoint_file(Path(out_path), 'master-file-updated.xlsx', sharepoint_folder_override=upload_folder)
        if uploaded:
            print("Upload completed successfully.")
        else:
            print("Upload failed.")
            return 4
    except Exception as e:
        print("ERROR during upload:", e)
        return 4

    print("\nAll steps completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
