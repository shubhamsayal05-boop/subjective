"""
Windows/desktop launcher for the DRB Subjective Drivability Tool.

PyInstaller bundles this script as the executable entry point. It starts the
Streamlit server and opens the app in the default browser.

The app is configured for fully offline use (no Wi-Fi / internet required).
Everything runs locally on 127.0.0.1.
"""
from __future__ import annotations

import os
import sys
import webbrowser
from threading import Timer
from urllib.error import URLError
from urllib.request import urlopen

HOST = "127.0.0.1"


def resource_path(relative: str) -> str:
    """Resolve a bundled resource path (PyInstaller one-file/one-dir)."""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)


def configure_offline_runtime(app_root: str | None = None) -> None:
    """Apply settings so Streamlit never needs the internet."""
    os.environ.setdefault("NO_PROXY", "127.0.0.1,localhost")
    os.environ.setdefault("no_proxy", "127.0.0.1,localhost")

    offline_env = {
        "STREAMLIT_BROWSER_GATHER_USAGE_STATS": "false",
        "STREAMLIT_GLOBAL_DEVELOPMENT_MODE": "false",
        "STREAMLIT_SERVER_HEADLESS": "true",
        "STREAMLIT_SERVER_ADDRESS": HOST,
        "STREAMLIT_BROWSER_SERVER_ADDRESS": HOST,
        "STREAMLIT_CLIENT_TOOLBAR_MODE": "viewer",
        "STREAMLIT_CLIENT_SHOW_ERROR_LINKS": "false",
        "STREAMLIT_SERVER_ENABLE_CORS": "false",
        "STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION": "false",
        "STREAMLIT_SERVER_FILE_WATCHER_TYPE": "none",
        "MPLBACKEND": "Agg",
    }
    for key, value in offline_env.items():
        os.environ.setdefault(key, value)

    if app_root:
        os.environ.setdefault("DRB_APP_DIR", app_root)
        config_dir = os.path.join(app_root, ".streamlit")
        bundled_config = resource_path(".streamlit")
        if os.path.isdir(bundled_config) and not os.path.isdir(config_dir):
            os.makedirs(app_root, exist_ok=True)
            import shutil
            shutil.copytree(bundled_config, config_dir)


def _write_browser_shortcut(app_root: str, port: str) -> None:
    """Create a local shortcut users can open when offline (no Wi-Fi needed)."""
    url = f"http://{HOST}:{port}/"
    shortcut = os.path.join(app_root, "Open DRB Tool.url")
    if os.path.isfile(shortcut):
        return
    try:
        with open(shortcut, "w", encoding="utf-8") as f:
            f.write("[InternetShortcut]\n")
            f.write(f"URL={url}\n")
    except OSError:
        pass


def _open_browser_when_ready(url: str, attempts: int = 30) -> None:
    """Wait for the local Streamlit server, then open the browser."""
    for _ in range(attempts):
        try:
            with urlopen(url, timeout=1):
                webbrowser.open(url)
                return
        except (URLError, OSError):
            pass
        import time
        time.sleep(0.5)
    print(f"Open this address in your browser: {url}", flush=True)


def main() -> int:
    import streamlit.web.cli as stcli

    app_path = resource_path("app.py")
    if not os.path.isfile(app_path):
        print(f"ERROR: app.py not found at {app_path}", file=sys.stderr)
        return 1

    app_root = None
    if getattr(sys, "frozen", False):
        app_root = os.path.dirname(os.path.abspath(sys.executable))
        os.chdir(app_root)

    configure_offline_runtime(app_root)

    port = os.environ.get("DRB_PORT", "8501")
    url = f"http://{HOST}:{port}"

    if app_root:
        _write_browser_shortcut(app_root, port)

    print("=" * 60, flush=True)
    print(" DRB Subjective Tool — running offline on this PC", flush=True)
    print(f" Open in your browser: {url}", flush=True)
    print(" Wi-Fi is NOT required. If the browser does not open,", flush=True)
    print(f" double-click 'Open DRB Tool.url' next to the .exe", flush=True)
    print("=" * 60, flush=True)

    Timer(1.0, lambda: _open_browser_when_ready(url)).start()

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        f"--browser.serverAddress={HOST}",
        f"--server.address={HOST}",
        f"--server.port={port}",
        "--server.headless=true",
        "--client.toolbarMode=viewer",
        "--client.showErrorLinks=false",
        "--server.fileWatcherType=none",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
    ]
    return stcli.main()


if __name__ == "__main__":
    raise SystemExit(main())
