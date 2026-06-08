"""
Triggers the GitHub Actions workflow immediately.
Dashboard updates in ~20 seconds after running this.
"""
import requests
import sys
import time
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
GITHUB_TOKEN = "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"          # paste your GitHub Personal Access Token here
GITHUB_OWNER = "mitaoeadmissions"
GITHUB_REPO  = "admissions-dashboard"
WORKFLOW     = "update_dashboard.yml"
BRANCH       = "main"
# ─────────────────────────────────────────────────────────────────────────────

def trigger():
    if not GITHUB_TOKEN:
        print("ERROR: Paste your GitHub token in trigger_now.py (GITHUB_TOKEN = '...')")
        input("Press Enter to close...")
        sys.exit(1)

    print("=" * 50)
    print("  Triggering dashboard update...")
    print("=" * 50)

    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/actions/workflows/{WORKFLOW}/dispatches"
    r = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "Accept": "application/vnd.github+json",
        },
        json={"ref": BRANCH},
        timeout=15,
    )

    if r.status_code == 204:
        print("\n  SUCCESS! Dashboard will update in ~20 seconds.")
        print(f"\n  Live URL: https://mitaoeadmissions.github.io/admissions-dashboard/")
        print("\n  You can close this window.")
    else:
        print(f"\n  ERROR ({r.status_code}): {r.text[:200]}")
        print("\n  Check your GitHub token and try again.")

    try:
        input("\nPress Enter to close...")
    except (EOFError, OSError):
        pass

if __name__ == "__main__":
    trigger()
