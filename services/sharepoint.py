import requests
import os
import json
from .auth import get_access_token

TENANT_NAME = os.getenv("TENANT_NAME")
SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
FOLDER = os.getenv("SHAREPOINT_FOLDER")

def get_site_id():
    token = get_access_token()
    
    # If the user specifically wants the root site (Common for "Communication site")
    if SITE_NAME in ["root", "Communication site", ""] or not SITE_NAME:
        url = "https://graph.microsoft.com/v1.0/sites/root"
    else:
        url = f"https://graph.microsoft.com/v1.0/sites/{TENANT_NAME}.sharepoint.com:/sites/{SITE_NAME}"
    
    res = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"}
    )
    
    if res.status_code != 200:
        # Fallback: Search for the site if direct path fails
        search_url = f"https://graph.microsoft.com/v1.0/sites?search={SITE_NAME}"
        search_res = requests.get(search_url, headers={"Authorization": f"Bearer {token}"})
        if search_res.status_code == 200:
            sites = search_res.json().get('value', [])
            if sites:
                return sites[0]["id"]
        
        raise Exception(f"Failed to get site ID for '{SITE_NAME}': {res.text}")
    
    return res.json()["id"]


def get_drive_id(site_id):
    token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    res = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"}
    )
    if res.status_code != 200:
        raise Exception(f"Failed to get drive ID: {res.text}")
    return res.json()["id"]

def upload_file_to_sharepoint(file_data, filename, subfolder=None):
    """
    Uploads file data (bytes) to SharePoint.
    subfolder: optional subfolder inside the main SharePoint folder (e.g. employee doc folder)
    """
    token = get_access_token()
    site_id = get_site_id()
    drive_id = get_drive_id(site_id)

    target_path = f"{FOLDER}/{subfolder}/{filename}" if subfolder else f"{FOLDER}/{filename}"
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{target_path}:/content"

    res = requests.put(
        upload_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        },
        data=file_data
    )

    if res.status_code not in [200, 201]:
        raise Exception(f"Failed to upload to SharePoint: {res.text}")

    data = res.json()
    return {
        "url": data.get("webUrl"),
        "id": data.get("id")
    }

def upload_json_to_sharepoint(json_data, filename, subfolder=None):
    """Uploads JSON content to SharePoint"""
    content = json.dumps(json_data, indent=2).encode('utf-8')
    return upload_file_to_sharepoint(content, filename, subfolder)
