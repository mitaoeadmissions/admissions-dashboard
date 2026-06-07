"""
Watches Masterdata File.xlsx for changes, regenerates the dashboard,
and auto-deploys to Netlify so the team always sees the latest data.

Run:  python watch.py
Local preview:  http://localhost:8000/dashboard.html
Team URL:       https://mit-admissions-dashboard.netlify.app
"""
import time
import threading
import webbrowser
import json
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import generate_dashboard
import netlify_deploy

WATCH_FILE   = generate_dashboard.EXCEL_FILE
SERVE_DIR    = generate_dashboard.OUTPUT_HTML.parent
DASHBOARD    = generate_dashboard.OUTPUT_HTML.name
VERSION_FILE = SERVE_DIR / "version.json"
PORT         = 8000


def write_version():
    VERSION_FILE.write_text(json.dumps({"v": str(int(time.time()))}), encoding="utf-8")


# ── HTTP server (local preview, no-cache) ─────────────────────────────────────

class NoCacheHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(SERVE_DIR), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, format, *args):
        pass  # keep terminal clean


# ── Excel watcher ─────────────────────────────────────────────────────────────

class ExcelHandler(FileSystemEventHandler):
    def __init__(self):
        self._last_trigger = 0

    def on_modified(self, event):
        if Path(event.src_path).resolve() == WATCH_FILE.resolve():
            now = time.time()
            if now - self._last_trigger < 3:
                return
            self._last_trigger = now
            print(f"\n[{time.strftime('%H:%M:%S')}] Change detected in {WATCH_FILE.name}")
            time.sleep(1.0)
            self._run_pipeline()

    def on_created(self, event):
        self.on_modified(event)

    def _run_pipeline(self):
        # Step 1 — regenerate HTML
        try:
            print("  [1/3] Regenerating dashboard...")
            generate_dashboard.generate()
            write_version()
            print("  [1/3] Done — local dashboard updated.")
        except Exception as e:
            print(f"  [1/3] ERROR generating dashboard: {e}")
            return

        # Step 2 — deploy to Netlify in a background thread (non-blocking)
        print("  [2/3] Deploying to Netlify...")
        def deploy():
            try:
                netlify_deploy.deploy()
                print("  [3/3] Team URL updated successfully.")
            except Exception as e:
                print(f"  [3/3] Netlify deploy error: {e}")
        threading.Thread(target=deploy, daemon=True).start()


# ── startup ───────────────────────────────────────────────────────────────────

def main():
    import config
    print("=" * 62)
    print("  MIT Admissions Dashboard — Watcher + Auto-Deploy")
    print("=" * 62)
    print(f"  Watching : {WATCH_FILE.name}")
    print(f"  Local    : http://localhost:{PORT}/{DASHBOARD}")
    print(f"  Team URL : {config.NETLIFY_SITE_URL}")
    print(f"  Press Ctrl+C to stop.")
    print("=" * 62)

    # Generate + deploy on startup
    print("\n  Starting up — generating and deploying initial dashboard...")
    try:
        generate_dashboard.generate()
        write_version()
        print("  Dashboard generated.")
    except Exception as e:
        print(f"  Warning on generation: {e}")

    try:
        netlify_deploy.deploy()
    except Exception as e:
        print(f"  Warning on initial deploy: {e}")

    # Start local server
    server = HTTPServer(("localhost", PORT), NoCacheHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()

    # Open local preview in browser
    webbrowser.open(f"http://localhost:{PORT}/{DASHBOARD}")
    print(f"\n  Local browser opened.")
    print(f"  Share this with your team: {config.NETLIFY_SITE_URL}\n")

    # Start file watcher
    handler  = ExcelHandler()
    observer = Observer()
    observer.schedule(handler, str(WATCH_FILE.parent), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        server.shutdown()
        print("\nStopped.")
    observer.join()


if __name__ == "__main__":
    main()
