"""
Deploys dashboard_share.html to Netlify using the file-digest API.
Includes a _headers file to force Content-Type: text/html on all browsers/devices.
"""
import requests
import hashlib
import sys
import os
from pathlib import Path

# Load config — from config.py locally, or from environment variables in GitHub Actions
sys.path.insert(0, str(Path(__file__).parent))
try:
    import config as _config
    _cfg_token   = _config.NETLIFY_TOKEN
    _cfg_site_id = _config.NETLIFY_SITE_ID
    _cfg_site_url = _config.NETLIFY_SITE_URL
except ImportError:
    _cfg_token = _cfg_site_id = _cfg_site_url = ""

_IN_CLOUD = os.environ.get("GITHUB_ACTIONS") == "true"

class config:
    NETLIFY_TOKEN    = os.environ.get("NETLIFY_TOKEN")    or _cfg_token
    NETLIFY_SITE_ID  = os.environ.get("NETLIFY_SITE_ID")  or _cfg_site_id
    NETLIFY_SITE_URL = os.environ.get("NETLIFY_SITE_URL") or _cfg_site_url or "https://mit-admissions-dashboard.netlify.app"

SHARE_HTML = (
    Path("dashboard_share.html")
    if _IN_CLOUD
    else Path(r"C:\Users\guruv\Desktop\Office\Admission Dashboard\dashboard_share.html")
)

# _headers file content — tells Netlify: always serve index.html as text/html
HEADERS_CONTENT = (
    "/index.html\n"
    "  Content-Type: text/html; charset=utf-8\n"
    "  X-Content-Type-Options: nosniff\n"
).encode("utf-8")


def sha1(content: bytes) -> str:
    return hashlib.sha1(content).hexdigest()


def upload_file(deploy_id: str, path: str, content: bytes,
                content_type: str, auth: dict) -> bool:
    """Upload a single file to a Netlify deploy."""
    r = requests.put(
        f"https://api.netlify.com/api/v1/deploys/{deploy_id}/files{path}",
        headers={**auth, "Content-Type": content_type},
        data=content,
        timeout=60,
    )
    # 422 = Netlify already has this file cached — that's fine
    if r.status_code not in (200, 201, 422):
        print(f"  [Netlify] Upload {path} failed ({r.status_code}): {r.text[:200]}")
        return False
    return True


def deploy():
    if not config.NETLIFY_TOKEN or not config.NETLIFY_SITE_ID:
        print("  [Netlify] Token or Site ID missing in config.py — skipping.")
        return False

    if not SHARE_HTML.exists():
        print("  [Netlify] dashboard_share.html not found — skipping.")
        return False

    print(f"  [Netlify] Deploying to {config.NETLIFY_SITE_URL} ...")

    html_content    = SHARE_HTML.read_bytes()
    html_sha        = sha1(html_content)
    headers_sha     = sha1(HEADERS_CONTENT)
    auth            = {"Authorization": f"Bearer {config.NETLIFY_TOKEN}"}

    # ── Step 1: create deploy with BOTH files in manifest ────────────────────
    r1 = requests.post(
        f"https://api.netlify.com/api/v1/sites/{config.NETLIFY_SITE_ID}/deploys",
        headers={**auth, "Content-Type": "application/json"},
        json={"files": {
            "/index.html": html_sha,
            "/_headers":   headers_sha,
        }},
        timeout=30,
    )
    if r1.status_code not in (200, 201):
        print(f"  [Netlify] Deploy create failed ({r1.status_code}): {r1.text[:200]}")
        return False

    deploy_data = r1.json()
    deploy_id   = deploy_data["id"]
    required    = set(deploy_data.get("required", []))

    # ── Step 2: upload whichever files Netlify needs ──────────────────────────
    # Always attempt both — 422 is silently accepted if already cached.
    upload_file(deploy_id, "/index.html", html_content,
                "text/html; charset=utf-8", auth)
    upload_file(deploy_id, "/_headers",   HEADERS_CONTENT,
                "text/plain", auth)

    print(f"  [Netlify] Deploy successful — live at: {config.NETLIFY_SITE_URL}")
    return True


if __name__ == "__main__":
    deploy()
