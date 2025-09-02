import os
import requests
from dotenv import load_dotenv
from graph_auth import acquire_graph_token

# Load environment variables
load_dotenv()

def download_sharepoint_file(file_name):
    """Download a file from SharePoint using Microsoft Graph API"""

    # Get credentials from environment
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")
    sharepoint_folder = os.getenv("SHAREPOINT_UPLOAD_FOLDER")

    # Acquire access token
    token = acquire_graph_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token['access_token']}"}

    try:
        # First, get the site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:/sites/{site_path}"
        site_response = requests.get(site_url, headers=headers)

        if site_response.status_code != 200:
            print(f"Failed to get site info. Status: {site_response.status_code}")
            print(f"Response: {site_response.text}")
            return False

        site_data = site_response.json()
        site_id = site_data['id']
        print(f"Found site ID: {site_id}")

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
        print(f"Attempting to download from: {file_url}")

        # Download the file
        response = requests.get(file_url, headers=headers)

        if response.status_code == 200:
            with open(file_name, "wb") as local_file:
                local_file.write(response.content)
            print(f"Successfully downloaded: {file_name}")
            return True
        else:
            print(f"Failed to download file. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False

    except Exception as e:
        print(f"Error during download: {str(e)}")
        return False

# Example usage
if __name__ == "__main__":
    # Download the specific file mentioned in your SharePoint link
    file_name = "HudsonMatrixPrices08272025020701PM.xlsm"
    download_sharepoint_file(file_name)
