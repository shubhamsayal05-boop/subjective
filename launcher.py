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

HOST = "127.0.0.1"


def resource_path(relative: str) -> str:
    """Resolve a bundled resource path (PyInstaller one-file/one-dir)."""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)


def configure_offline_runtime(app_root: str | None = None) -> None:
    """Apply settings so Streamlit never needs the internet."""
    # Avoid corporate proxies intercepting localhost when offline.
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
        # Let Streamlit pick up bundled .streamlit/config.toml when present.
        config_dir = os.path.join(app_root, ".streamlit")
        bundled_config = resource_path(".streamlit")
        if os.path.isdir(bundled_config) and not os.path.isdir(config_dir):
            os.makedirs(app_root, exist_ok=True)
            import shutil
            shutil.copytree(bundled_config, config_dir)


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

    # Open the browser once the server has a moment to bind.
    Timer(1.5, lambda: webbrowser.open(url)).start()

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
