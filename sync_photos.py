"""
sync_photos.py
--------------
Downloads website photos from the DataHub SharePoint site and writes
photos.json, which the static site reads at page load.

SharePoint source folders (relative to the Shared Documents library root):
    Textbooks/Electra Site/Home Page Pic   -> first image (alphabetical) shown on Home
    Textbooks/Electra Site/Photo Gallery   -> all images shown in the gallery carousel

Authentication: Foxtrot Report Automation Entra app (client credentials),
same pattern as the other Foxtrot automation repos. Requires env vars:
    TENANT_ID, CLIENT_ID, CLIENT_SECRET
"""

import json
import os
import shutil
import sys
from urllib.parse import quote

import requests

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SP_SITE = "foxtrotaviationcom.sharepoint.com:/sites/DataHub:"

HOME_PIC_FOLDER = "Textbooks/Electra Site/Home Page Pic"
GALLERY_FOLDER = "Textbooks/Electra Site/Photo Gallery"

LOCAL_HOME_DIR = "images/home"
LOCAL_GALLERY_DIR = "images/gallery"
MANIFEST_PATH = "photos.json"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".svg"}

GRAPH = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------
def get_token() -> str:
    tenant = os.environ["TENANT_ID"]
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": os.environ["CLIENT_ID"],
            "client_secret": os.environ["CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    if not resp.ok:
        print(f"Token request failed ({resp.status_code}): {resp.text}")
    resp.raise_for_status()
    return resp.json()["access_token"]


def get_drive_id(headers: dict) -> str:
    site = requests.get(f"{GRAPH}/sites/{SP_SITE}", headers=headers, timeout=30)
    site.raise_for_status()
    site_id = site.json()["id"]

    drive = requests.get(f"{GRAPH}/sites/{site_id}/drive", headers=headers, timeout=30)
    drive.raise_for_status()
    return drive.json()["id"]


def list_folder_images(drive_id: str, folder_path: str, headers: dict) -> list[dict]:
    """Return image files in a SharePoint folder, sorted alphabetically.

    An empty list is returned if the folder does not exist or has no images.
    """
    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(folder_path)}:/children"
    items: list[dict] = []

    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 404:
            print(f"  Folder not found (treating as empty): {folder_path}")
            return []
        resp.raise_for_status()
        payload = resp.json()
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")

    images = [
        item
        for item in items
        if "file" in item
        and os.path.splitext(item["name"])[1].lower() in IMAGE_EXTENSIONS
    ]
    images.sort(key=lambda item: item["name"].lower())
    return images


def download_item(item: dict, dest_dir: str) -> str:
    """Download a drive item to dest_dir; returns the local relative path."""
    download_url = item["@microsoft.graph.downloadUrl"]
    local_path = os.path.join(dest_dir, item["name"])
    with requests.get(download_url, stream=True, timeout=120) as resp:
        resp.raise_for_status()
        with open(local_path, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=1 << 16):
                fh.write(chunk)
    print(f"  Downloaded: {local_path}")
    return local_path.replace(os.sep, "/")


def reset_dir(path: str) -> None:
    """Delete and recreate a local image folder so removed SharePoint files disappear."""
    shutil.rmtree(path, ignore_errors=True)
    os.makedirs(path, exist_ok=True)
    # Keep the folder present in git even when empty
    open(os.path.join(path, ".gitkeep"), "w").close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = get_drive_id(headers)

    # --- Home page picture: first image alphabetically, or nothing ---
    print(f"Syncing home page pic from: {HOME_PIC_FOLDER}")
    reset_dir(LOCAL_HOME_DIR)
    home_images = list_folder_images(drive_id, HOME_PIC_FOLDER, headers)
    home_pic_path = None
    if home_images:
        home_pic_path = download_item(home_images[0], LOCAL_HOME_DIR)
    else:
        print("  No home page picture found.")

    # --- Gallery: every image in the folder ---
    print(f"Syncing gallery from: {GALLERY_FOLDER}")
    reset_dir(LOCAL_GALLERY_DIR)
    gallery_paths = [
        download_item(item, LOCAL_GALLERY_DIR)
        for item in list_folder_images(drive_id, GALLERY_FOLDER, headers)
    ]
    if not gallery_paths:
        print("  No gallery photos found.")

    # --- Manifest ---
    manifest = {"homePagePic": home_pic_path, "gallery": gallery_paths}
    with open(MANIFEST_PATH, "w") as fh:
        json.dump(manifest, fh, indent=2)
    print(f"Wrote {MANIFEST_PATH}: 1 home pic = {bool(home_pic_path)}, "
          f"{len(gallery_paths)} gallery photos")

    return 0


if __name__ == "__main__":
    sys.exit(main())
