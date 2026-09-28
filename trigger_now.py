"""
Generates the dashboard from your local Excel file and pushes it live.
Run this whenever you want to update the dashboard with the latest data.
"""
import sys, os, subprocess, shutil, tempfile
from pathlib import Path
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as _cfg

GITHUB_TOKEN = _cfg.GITHUB_TOKEN
REPO_URL     = f"https://{GITHUB_TOKEN}@github.com/mitaoeadmissions/admissions-dashboard.git"
AUTOMATE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE   = Path(r"C:\Users\guruv\Dropbox\Admission Dashboard\Masterdata File.xlsx")
OUTPUT_DIR   = Path(r"C:\Users\guruv\OneDrive\Desktop\Office\Admission Dashboard")

def sep(title=""):
    print("=" * 50)
    if title:
        print(f"  {title}")
        print("=" * 50)

def run(cmd, **kwargs):
    return subprocess.run(cmd, capture_output=True, text=True, **kwargs)

def main():
    sep("MITAOE Dashboard Updater")
    print()

    # ── 1. Check Excel exists ────────────────────────────────────────────────
    if not EXCEL_FILE.exists():
        print(f"  ERROR: Excel file not found at:\n  {EXCEL_FILE}")
        input("\nPress Enter to close...")
        sys.exit(1)

    mtime = datetime.fromtimestamp(EXCEL_FILE.stat().st_mtime)
    print(f"  Excel file : {EXCEL_FILE.name}")
    print(f"  Modified   : {mtime.strftime('%d-%m-%Y %H:%M:%S')}")
    print()

    # ── 2. Generate dashboard from local Excel ───────────────────────────────
    print("  Step 1/3 — Generating dashboard from Excel...")
    gen_script = AUTOMATE_DIR / "generate_dashboard.py"
    result = run([sys.executable, str(gen_script)], cwd=str(AUTOMATE_DIR))
    if result.returncode != 0:
        print("  ERROR during generation:")
        print(result.stderr[-600:])
        input("\nPress Enter to close...")
        sys.exit(1)

    # Extract the latest date from output
    for line in result.stdout.splitlines():
        if "Latest date" in line or "Totals" in line:
            print(f"    {line.strip()}")
    print()

    # ── 3. Clone gh-pages and swap index.html ───────────────────────────────
    print("  Step 2/3 — Preparing release...")
    live_html = OUTPUT_DIR / "dashboard.html"
    if not live_html.exists():
        print(f"  ERROR: Generated file not found: {live_html}")
        input("\nPress Enter to close...")
        sys.exit(1)

    tmp = Path(tempfile.mkdtemp())
    try:
        clone = run(["git", "clone", "--branch", "gh-pages", "--single-branch", REPO_URL, str(tmp)])
        if clone.returncode != 0:
            print("  ERROR cloning gh-pages:", clone.stderr[:200])
            input("\nPress Enter to close...")
            sys.exit(1)

        shutil.copy2(str(live_html), str(tmp / "index.html"))

        today = datetime.today().strftime("%Y-%m-%d")
        for cmd in [
            ["git", "-C", str(tmp), "config", "user.name", "Guruv"],
            ["git", "-C", str(tmp), "config", "user.email", "admissions@mitaoe.ac.in"],
            ["git", "-C", str(tmp), "add", "index.html"],
            ["git", "-C", str(tmp), "commit", "-m", f"Update dashboard {today} [skip ci]"],
        ]:
            run(cmd)

        # ── 4. Push ──────────────────────────────────────────────────────────
        print("  Step 3/3 — Publishing...")
        push = run(["git", "-C", str(tmp), "push", REPO_URL, "gh-pages", "--force"])
        if push.returncode != 0:
            print("  ERROR pushing:", push.stderr[:300])
            input("\nPress Enter to close...")
            sys.exit(1)

    finally:
        shutil.rmtree(str(tmp), ignore_errors=True)

    # ── Done ─────────────────────────────────────────────────────────────────
    print()
    sep()
    print("  SUCCESS! Dashboard is now live.")
    print()
    print("  URL: https://mitaoeadmissions.github.io/admissions-dashboard/")
    print()
    print("  Refresh the dashboard in your browser to see the latest data.")
    sep()
    try:
        input("\nPress Enter to close...")
    except (EOFError, OSError):
        pass

if __name__ == "__main__":
    main()
