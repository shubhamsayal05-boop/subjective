"""JSON session persistence — the tool is fully Excel-independent."""
import json
import os
import sys
from datetime import datetime


def app_dir() -> str:
    """Directory for writable app data (sessions folder).

    When frozen as a .exe, sessions live next to the executable so they
    survive restarts and are easy to find/back up. The launcher sets
    DRB_APP_DIR before Streamlit starts so this stays correct even when
    modules are loaded from PyInstaller's temp extract folder.
    """
    env = os.environ.get("DRB_APP_DIR")
    if env:
        return env
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def session_dir() -> str:
    """Return the sessions folder, creating it if needed."""
    path = os.path.join(app_dir(), "sessions")
    os.makedirs(path, exist_ok=True)
    return path


def session_filename(vehicle: dict, variant: str) -> str:
    parts = [vehicle.get("Model Year") or "0",
             vehicle.get("Vehicle Line") or "0",
             variant,
             vehicle.get("Last 4 of VIN") or "0",
             vehicle.get("date") or datetime.now().strftime("%m%d%y")]
    safe = "_".join(str(p).replace(" ", "").replace("/", "-") for p in parts)
    return f"{safe}.json"


def save_session(state: dict, name: str | None = None) -> str:
    name = name or session_filename(state.get("vehicle", {}), state.get("variant", "NA"))
    if not name.endswith(".json"):
        name += ".json"
    path = os.path.join(session_dir(), name)
    payload = dict(state)
    payload["_saved_at"] = datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1)
    return path


def list_sessions() -> list[str]:
    d = session_dir()
    try:
        return sorted(f for f in os.listdir(d) if f.endswith(".json"))
    except OSError:
        return []


def load_session(name: str) -> dict:
    with open(os.path.join(session_dir(), name), encoding="utf-8") as f:
        return json.load(f)
