"""Shared filename / INCA label token rules."""
from datetime import datetime


def vehicle_token(vehicle: dict) -> str:
    """INCA/AVL recorder suffix used on test sheet filenames.

    Format: ModelYear_EngineDisp_Transmission_VehicleLine_Last4VIN_MMDDYY
    (Vehicle Line = LB; Model Code is intentionally excluded.)
    """
    parts = [
        vehicle.get("Model Year") or "0",
        vehicle.get("Engine Disp.") or "0",
        vehicle.get("Transmission") or "0",
        vehicle.get("Vehicle Line") or "0",
        vehicle.get("Last 4 of VIN") or "0",
        vehicle.get("date") or datetime.now().strftime("%m%d%y"),
    ]
    return "_".join(str(p).replace(" ", "").replace("/", "-") for p in parts)


def inca_label(prefix: str, vehicle: dict) -> str:
    return f"{prefix}{vehicle_token(vehicle)}"


def session_filename(vehicle: dict, variant: str) -> str:
  """Auto-name for saved JSON sessions."""
  parts = [
      vehicle.get("Model Year") or "0",
      vehicle.get("Vehicle Line") or "0",
      variant,
      vehicle.get("Last 4 of VIN") or "0",
      vehicle.get("date") or datetime.now().strftime("%m%d%y"),
  ]
  safe = "_".join(str(p).replace(" ", "").replace("/", "-") for p in parts)
  return f"{safe}.json"


def export_basename(variant: str, vehicle: dict) -> str:
    return f"DRB_{variant}_{vehicle_token(vehicle)}"
