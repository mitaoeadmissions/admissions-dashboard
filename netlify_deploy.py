"""
Deploys dashboard_share.html to Netlify as index.html.
Called automatically by watch.py after every Excel save.
"""
import requests
import zipfile
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import config

SHARE_HTML = Path(r"C:\Users\guruv\Desktop\Office\Admission Dashboard\dashboard_share.html")


def deploy():
    if not config.NETLIFY_TOKEN or not config.NETLIFY_SITE_ID:
        print("  [Netlify] Token or Site ID missing in config.py — skipping deploy.")
        return False

    if not SHARE_HTML.exists():
        print("  [Netlify] dashboard_share.html not found — skipping deploy.")
        return False

    print(f"  [Netlify] Deploying to {config.NETLIFY_SITE_URL} ...")

    # Pack the share HTML as index.html inside a zip
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(str(SHARE_HTML), arcname="index.html")
    buf.seek(0)

    headers = {
        "Authorization": f"Bearer {config.NETLIFY_TOKEN}",
        "Content-Type": "application/zip",
    }

    r = requests.post(
        f"https://api.netlify.com/api/v1/sites/{config.NETLIFY_SITE_ID}/deploys",
        headers=headers,
        data=buf.read(),
        timeout=60,
    )

    if r.status_code in (200, 201):
        state = r.json().get("state", "?")
        print(f"  [Netlify] Deploy successful (state: {state})")
        print(f"  [Netlify] Live at: {config.NETLIFY_SITE_URL}")
        return True
    else:
        print(f"  [Netlify] Deploy failed ({r.status_code}): {r.text[:200]}")
        return False


if __name__ == "__main__":
    deploy()
