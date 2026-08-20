"""JSON session persistence — the tool is fully Excel-independent."""
import json
import os
from datetime import datetime

SESSION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sessions")
os.makedirs(SESSION_DIR, exist_ok=True)


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
    path = os.path.join(SESSION_DIR, name)
    payload = dict(state)
    payload["_saved_at"] = datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1)
    return path


def list_sessions() -> list[str]:
    return sorted(f for f in os.listdir(SESSION_DIR) if f.endswith(".json"))


def load_session(name: str) -> dict:
    with open(os.path.join(SESSION_DIR, name), encoding="utf-8") as f:
        return json.load(f)
