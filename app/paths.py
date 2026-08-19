"""Resolve application paths for development and PyInstaller builds."""

import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


def bundled_dir() -> Path:
    return app_root() / "app" / "bundled"


def main_script() -> Path:
    return app_root() / "app.py"
