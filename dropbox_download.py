"""
Downloads Masterdata File.xlsx from Dropbox.
Used by GitHub Actions before running generate_dashboard.py.
"""
import dropbox
import os
from pathlib import Path

DROPBOX_PATH = "/Admission Dashboard/Masterdata File.xlsx"
LOCAL_PATH   = Path("Masterdata File.xlsx")


def download():
    app_key       = os.environ["DROPBOX_APP_KEY"]
    app_secret    = os.environ["DROPBOX_APP_SECRET"]
    refresh_token = os.environ["DROPBOX_REFRESH_TOKEN"]

    print(f"Connecting to Dropbox...")
    dbx = dropbox.Dropbox(
        app_key=app_key,
        app_secret=app_secret,
        oauth2_refresh_token=refresh_token
    )

    print(f"Downloading {DROPBOX_PATH} ...")
    _, response = dbx.files_download(DROPBOX_PATH)
    LOCAL_PATH.write_bytes(response.content)
    print(f"Saved {len(response.content):,} bytes → {LOCAL_PATH}")


if __name__ == "__main__":
    download()
