"""
Windows/desktop launcher for the DRB Subjective Drivability Tool.

PyInstaller bundles this script as the executable entry point. It starts the
Streamlit server and opens the app in the default browser.
"""
from __future__ import annotations

import os
import sys
import webbrowser
from threading import Timer


def resource_path(relative: str) -> str:
    """Resolve a bundled resource path (PyInstaller one-file/one-dir)."""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)


def main() -> int:
    import streamlit.web.cli as stcli

    app_path = resource_path("app.py")
    if not os.path.isfile(app_path):
        print(f"ERROR: app.py not found at {app_path}", file=sys.stderr)
        return 1

    # Persist sessions next to the .exe, not inside the temp extract folder.
    if getattr(sys, "frozen", False):
        app_root = os.path.dirname(os.path.abspath(sys.executable))
        os.environ["DRB_APP_DIR"] = app_root
        os.chdir(app_root)

    port = os.environ.get("DRB_PORT", "8501")
    url = f"http://localhost:{port}"

    # Open the browser once the server has a moment to bind.
    Timer(1.5, lambda: webbrowser.open(url)).start()

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.headless=true",
        f"--server.port={port}",
        "--server.address=localhost",
    ]
    return stcli.main()


if __name__ == "__main__":
    raise SystemExit(main())
