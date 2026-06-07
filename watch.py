"""
Watches Masterdata File.xlsx for changes and regenerates dashboard.html automatically.
Also runs a local web server so the browser always gets the freshest file (no caching).

Run:  python watch.py
Then open:  http://localhost:8000
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

WATCH_FILE   = generate_dashboard.EXCEL_FILE
SERVE_DIR    = generate_dashboard.OUTPUT_HTML.parent   # folder containing dashboard.html
DASHBOARD    = generate_dashboard.OUTPUT_HTML.name
VERSION_FILE = SERVE_DIR / "version.json"
PORT         = 8000

def write_version():
    """Write a version token the browser polls to detect new data."""
    VERSION_FILE.write_text(json.dumps({"v": str(int(time.time()))}), encoding="utf-8")


# ── HTTP server with cache disabled ──────────────────────────────────────────

class NoCacheHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(SERVE_DIR), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, format, *args):
        pass  # suppress request logs to keep terminal clean


# ── Excel file watcher ────────────────────────────────────────────────────────

class ExcelHandler(FileSystemEventHandler):
    def __init__(self):
        self._last_trigger = 0

    def on_modified(self, event):
        if Path(event.src_path).resolve() == WATCH_FILE.resolve():
            now = time.time()
            if now - self._last_trigger < 3:
                return
            self._last_trigger = now
            print(f"\n[{time.strftime('%H:%M:%S')}] Change detected -- regenerating dashboard...")
            time.sleep(1.0)
            try:
                generate_dashboard.generate()
                write_version()
                print(f"  Done. Browser will reload automatically.")
            except Exception as e:
                print(f"  ERROR: {e}")

    def on_created(self, event):
        self.on_modified(event)


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  MIT Admissions Dashboard -- Live Watcher + Server")
    print("=" * 60)
    print(f"  Watching : {WATCH_FILE.name}")
    print(f"  Serving  : http://localhost:{PORT}/{DASHBOARD}")
    print(f"  Press Ctrl+C to stop.")
    print("=" * 60)

    # Generate once on startup
    try:
        generate_dashboard.generate()
        write_version()
        print(f"\n  Initial dashboard generated OK.")
    except Exception as e:
        print(f"\n  Warning on startup: {e}")

    # Start HTTP server in background thread
    server = HTTPServer(("localhost", PORT), NoCacheHandler)
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    print(f"\n  Server running at http://localhost:{PORT}/{DASHBOARD}")

    # Open browser automatically
    url = f"http://localhost:{PORT}/{DASHBOARD}"
    webbrowser.open(url)
    print(f"  Browser opened. If it didn't open, go to: {url}")
    print()

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
