"""
Deploys dashboard_share.html to Netlify using the file-digest API.
This correctly sets Content-Type: text/html so browsers render it properly.
"""
import requests
import hashlib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import config

SHARE_HTML = Path(r"C:\Users\guruv\Desktop\Office\Admission Dashboard\dashboard_share.html")


def sha1(content: bytes) -> str:
    return hashlib.sha1(content).hexdigest()


def deploy():
    if not config.NETLIFY_TOKEN or not config.NETLIFY_SITE_ID:
        print("  [Netlify] Token or Site ID missing in config.py — skipping.")
        return False

    if not SHARE_HTML.exists():
        print("  [Netlify] dashboard_share.html not found — skipping.")
        return False

    print(f"  [Netlify] Deploying to {config.NETLIFY_SITE_URL} ...")

    content  = SHARE_HTML.read_bytes()
    file_sha = sha1(content)
    auth     = {"Authorization": f"Bearer {config.NETLIFY_TOKEN}"}

    # ── Step 1: create deploy with file manifest ──────────────────────────────
    r1 = requests.post(
        f"https://api.netlify.com/api/v1/sites/{config.NETLIFY_SITE_ID}/deploys",
        headers={**auth, "Content-Type": "application/json"},
        json={"files": {"/index.html": file_sha}},
        timeout=30,
    )
    if r1.status_code not in (200, 201):
        print(f"  [Netlify] Deploy create failed ({r1.status_code}): {r1.text[:200]}")
        return False

    deploy_data = r1.json()
    deploy_id   = deploy_data["id"]
    required    = deploy_data.get("required", [])

    # ── Step 2: upload file only if Netlify requests it ──────────────────────
    # If required=[] Netlify already has the file cached — deploy is done.
    # If our SHA is in required — upload it now.
    if file_sha in required:
        r2 = requests.put(
            f"https://api.netlify.com/api/v1/deploys/{deploy_id}/files/index.html",
            headers={**auth, "Content-Type": "text/html; charset=utf-8"},
            data=content,
            timeout=60,
        )
        if r2.status_code not in (200, 201):
            print(f"  [Netlify] File upload failed ({r2.status_code}): {r2.text[:200]}")
            return False
    elif not required:
        # File already cached — Netlify will use it. Force upload anyway
        # by sending the file content so the deploy can transition to ready.
        r2 = requests.put(
            f"https://api.netlify.com/api/v1/deploys/{deploy_id}/files/index.html",
            headers={**auth, "Content-Type": "text/html; charset=utf-8"},
            data=content,
            timeout=60,
        )
        # 422 here just means Netlify didn't need it — that's fine

    print(f"  [Netlify] Deploy successful — live at: {config.NETLIFY_SITE_URL}")
    return True


if __name__ == "__main__":
    deploy()
