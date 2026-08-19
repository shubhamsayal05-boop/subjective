"""
Windows .exe entry point for Subjective Spreadsheet Tool.

Build on Windows with:  build_windows.bat
"""

from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path


def _resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _open_browser(port: str) -> None:
    time.sleep(2.5)
    webbrowser.open(f"http://127.0.0.1:{port}")


def main() -> None:
    root = _resource_root()
    os.chdir(root)

    app_file = root / "app.py"
    if not app_file.is_file():
        print(f"Error: app.py not found at {app_file}")
        input("Press Enter to exit...")
        sys.exit(1)

    port = os.environ.get("SUBJECTIVE_APP_PORT", "8501")

    print("=" * 60)
    print("  Subjective Spreadsheet Tool")
    print("  Starting local server...")
    print(f"  Open in browser: http://127.0.0.1:{port}")
    print("  Close this window to stop the application.")
    print("=" * 60)

    threading.Thread(target=_open_browser, args=(port,), daemon=True).start()

    sys.argv = [
        "streamlit",
        "run",
        str(app_file),
        "--global.developmentMode",
        "false",
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
        "--server.port",
        port,
        "--server.address",
        "127.0.0.1",
    ]

    try:
        from streamlit.web import cli as stcli

        sys.exit(stcli.main())
    except Exception as exc:
        print(f"Failed to start application: {exc}")
        input("Press Enter to exit...")
        sys.exit(1)


if __name__ == "__main__":
    main()
