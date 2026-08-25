"""Export filled session data to Excel (.xlsx) using the original DRB templates."""
from __future__ import annotations

import io
import os
import re
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

import config as C
import naming

TEMPLATES = {
    "AT": "Subjective_SprdSheet_072926.xlsm",
    "BEV": "BEV_Subjective_SprdSheet_072926.xlsm",
    "CVT": "CVT Subjective_SprdSheet_072926.xlsm",
}

TEST_SHEETS = {
    "driveaway": "Driveaway-sweeps",
    "driveaway_ess": "Driveaway ESS",
    "decel_cstdown": "Decel Cstdowns",
    "decel_cstdown_bev": "Decel Cstdowns",
    "decel_opd": "Decel OPD",
    "uss_manual": "USS-Manual (Opt.)",
    "rtito": "RTITO_AD",
    "ti_cstspd": "TI_CstSpd ",
    "to_cstspd": "TO_CstSpd",
    "rrl": "RRL (with HS)",
    "hill_start": "RRL (with HS)",
    "gs_static": "Garage Shifts",
    "gs_rolling": "Garage Shifts",
    "gs_extra": "Garage Shifts",
    "tito_opt": "TITO(Opt)",
    "cstspd_manual": "CstSpd",
    "cstspd_drive": "CstSpd",
    "accel_fl": "Acceleration(Opt)",
    "accel_plcp": "Acceleration(Opt)",
    "accel_plrp": "Acceleration(Opt)",
    "stationary_ess": "Stationary Test",
    "stationary_ac": "Stationary Test",
    "stationary_load": "Stationary Test",
    "stationary_thr": "Stationary Test",
    "kickdowns": "Kickdowns",
}

FILL = {
    "green": PatternFill("solid", fgColor="00B050"),
    "yellow": PatternFill("solid", fgColor="FFC000"),
    "red": PatternFill("solid", fgColor="FF3B30"),
    "grey": PatternFill("solid", fgColor="D6D6D6"),
    "white": PatternFill("solid", fgColor="FFFFFF"),
}


def resource_path(name: str) -> str:
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def cell_severity(code: str) -> str:
    s = str(code).strip()
    if s in ("", "None", "nan"):
        return "clean"
    if re.fullmatch(r"10|[1-9](\.\d+)?", s):
        r = float(s)
        return "clean" if r >= C.SUBJ_GREEN else ("event" if r >= C.SUBJ_YELLOW else "bad")
    if "!" in s:
        return "bad"
    return "event"


def sev_fill(sev: str, tested: bool):
    if sev == "bad":
        return FILL["red"]
    if sev == "event":
        return FILL["yellow"]
    return FILL["green"] if tested else FILL["grey"]


def _pedal_excel_value(pedal: str):
    if pedal == "No Pedal":
        return 0
    if pedal.endswith("%"):
        val = pedal.rstrip("%").strip()
        if val.replace(".", "").isdigit():
            return float(val) / 100.0
    if "99%" in pedal:
        return "99% (Opt.)"
    if "100%" in pedal or "detent" in pedal.lower():
        return 1
    return pedal


def _vehicle_token(vehicle: dict) -> str:
    return naming.vehicle_token(vehicle)


def _test_def(key: str, variant: str) -> dict:
    d = dict(C.TESTS[key])
    ov = C.VARIANT_OVERRIDES.get(variant, {})
    if key in ov:
        d.update(ov[key])
    return d


def _run_rows(key: str, variant: str) -> list[str]:
    if key in C.RUN_TESTS:
        return [f"{r['pedal']} · {r['brake']} ({r['ptype']})" for r in C.DRIVEAWAY_RUNS]
    return _test_def(key, variant)["rows"]


def _find_header_row(ws, col_names: list[str]) -> tuple[int, dict[str, int]] | None:
    targets = {_norm(c) for c in col_names}
    for r in range(1, min(ws.max_row, 40) + 1):
        mapping: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            raw = ws.cell(r, c).value
            if raw is None:
                continue
            label = _norm(str(raw).split("\n")[0])
            for name in col_names:
                n = _norm(name)
                if label == n or n in label or label in n:
                    mapping[name] = c
        if "Complete" in mapping or any(k in mapping for k in col_names):
            if len(mapping) >= 2:
                return r, mapping
    return None


def _set_cell(ws, row: int, col: int, value, fill=None):
    _safe_set(ws, row, col, value)
    cell = ws.cell(row, col)
    from openpyxl.cell.cell import MergedCell
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    if fill is not None:
        cell.fill = copy(fill)


def _fill_filename_cells(ws, prefix: str, vehicle: dict):
    token = _vehicle_token(vehicle)
    label = f"{prefix}{token}"
    for r in range(1, 20):
        for c in range(1, 8):
            v = ws.cell(r, c).value
            if v and "filename" in _norm(v):
                ws.cell(r, c + 1).value = label


def _fill_driveaway_sheet(ws, test_data: dict, variant: str):
    rows = _run_rows("driveaway", variant)
    cols = _test_def("driveaway", variant)["cols"]
    header = _find_header_row(ws, ["Complete"] + cols)
    if not header:
        return
    _, col_map = header
    complete_col = col_map.get("Complete", 4)

    excel_rows: list[tuple[int, float | str, str]] = []
    for r in range(8, min(ws.max_row, 120)):
        ped = ws.cell(r, 2).value
        brk = ws.cell(r, 3).value
        if brk:
            excel_rows.append((r, ped, str(brk).strip()))

    run_by_label = {f"{x['pedal']} · {x['brake']} ({x['ptype']})": x for x in C.DRIVEAWAY_RUNS}

    for label in rows:
        run = run_by_label[label]
        target_ped = _pedal_excel_value(run["pedal"])
        target_brk = run["brake"]
        match_row = None
        for er, ped, brk in excel_rows:
            ped_ok = False
            if isinstance(target_ped, (int, float)) and isinstance(ped, (int, float)):
                ped_ok = abs(float(ped) - float(target_ped)) < 0.001
            else:
                ped_ok = _norm(ped) == _norm(target_ped)
            if ped_ok and _norm(brk) == _norm(target_brk):
                match_row = er
                break
        if match_row is None:
            continue
        if test_data["complete"].get(label):
            _set_cell(ws, match_row, complete_col, "x")
        tg = str(test_data["topgear"].get(label, "")).strip()
        # Top gear stored after event columns when a spare column exists.
        tg_col = max(col_map.values(), default=0) + 1
        if tg and tg_col <= ws.max_column + 3:
            _set_cell(ws, match_row, tg_col, tg)
        for col_name in cols:
            cidx = col_map.get(col_name)
            if not cidx:
                continue
            code = test_data["codes"][label][col_name]
            tested = bool(test_data["complete"].get(label))
            sev = cell_severity(code)
            _set_cell(ws, match_row, cidx, code or None,
                      sev_fill(sev, tested and (code or tested)))


def _fill_simple_grid(ws, rows: list[str], cols: list[str], test_data: dict, *,
                      row_col: int = 1, start_search: int = 1):
    header = _find_header_row(ws, ["Complete"] + cols)
    if not header:
        header = _find_header_row(ws, cols)
    if not header:
        return
    header_row, col_map = header
    complete_col = col_map.get("Complete")

    row_lookup: dict[str, int] = {}
    for r in range(header_row + 1, min(ws.max_row, header_row + 200)):
        for c in (row_col, row_col + 1, 2):
            val = ws.cell(r, c).value
            if val is None:
                continue
            row_lookup[_norm(val)] = r

    for row_label in rows:
        ridx = row_lookup.get(_norm(row_label))
        if ridx is None:
            # partial match
            for k, rr in row_lookup.items():
                if _norm(row_label) in k or k in _norm(row_label):
                    ridx = rr
                    break
        if ridx is None:
            continue
        if complete_col and test_data["complete"].get(row_label):
            _set_cell(ws, ridx, complete_col, "x")
        for col_name in cols:
            cidx = col_map.get(col_name)
            if not cidx:
                for k, idx in col_map.items():
                    if _norm(col_name) == _norm(k) or _norm(col_name) in _norm(k):
                        cidx = idx
                        break
            if not cidx:
                continue
            code = test_data["codes"][row_label][col_name]
            tested = bool(test_data["complete"].get(row_label))
            sev = cell_severity(code)
            _set_cell(ws, ridx, cidx, code or None, sev_fill(sev, tested and (code or tested)))


def _safe_set(ws, row: int, col: int, value):
    cell = ws.cell(row, col)
    from openpyxl.cell.cell import MergedCell
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    cell.value = value


def _fill_vehicle_summary(wb, vehicle: dict, variant: str):
    if "DRB - Summary" not in wb.sheetnames:
        return
    ws = wb["DRB - Summary"]
    mapping = {
        "Model Year:": vehicle.get("Model Year", ""),
        "Vehicle Line:": vehicle.get("Vehicle Line", ""),
        "Engine:": vehicle.get("Engine Disp.", ""),
        "Transmisson:": vehicle.get("Transmission", ""),
        "Software Level:": vehicle.get("Software Level", ""),
        "Vin+(Last 3)": vehicle.get("Last 4 of VIN", ""),
    }
    for r in range(1, 6):
        for c in range(1, 20):
            label = ws.cell(r, c).value
            if not label:
                continue
            key = str(label).strip()
            for prefix, val in mapping.items():
                if key.startswith(prefix.split(":")[0]):
                    _safe_set(ws, r, c + 1, val)


def _fill_drb_driveaway_summary(wb, test_data: dict, variant: str):
    if "DRB - Summary" not in wb.sheetnames or "driveaway" not in C.VARIANT_TESTS.get(variant, []):
        return
    ws = wb["DRB - Summary"]
    cols = _test_def("driveaway", variant)["cols"]

    # Header row with "Pedal %" and shift columns.
    header_row = None
    col_map: dict[str, int] = {}
    for r in range(1, 30):
        if ws.cell(r, 2).value and "pedal" in _norm(ws.cell(r, 2).value):
            header_row = r
            for c in range(3, 40):
                v = ws.cell(r, c).value
                if v and _norm(v) not in ("subjective", "frequency", "objective"):
                    col_map[str(v).strip()] = c
            break
    if not header_row:
        return

    # Import aggregate logic inline (mirror app.py rules).
    groups = {}
    for label, done in test_data["complete"].items():
        pedal = label.split(" · ")[0]
        g = groups.setdefault(pedal, dict(runs=0, ev={c: 0 for c in cols}, bad={c: False for c in cols}))
        if done:
            g["runs"] += 1
        for c in cols:
            sev = cell_severity(test_data["codes"][label][c])
            if sev != "clean":
                g["ev"][c] += 1
                if sev == "bad":
                    g["bad"][c] = True

    pedal_rows = {}
    for r in range(header_row + 1, header_row + 30):
        ped = ws.cell(r, 2).value
        if ped is not None:
            pedal_rows[r] = ped

    for r, ped_val in pedal_rows.items():
        pedal_key = None
        for p in C.DRIVEAWAY_PEDAL_ORDER:
            if _pedal_excel_value(p) == ped_val or _norm(str(ped_val)) == _norm(_pedal_excel_value(p)):
                pedal_key = p
                break
        if pedal_key not in groups:
            continue
        g = groups[pedal_key]
        for col_name, start_c in col_map.items():
            if col_name not in cols:
                continue
            sev = "bad" if g["bad"][col_name] else ("event" if g["ev"][col_name] else "clean")
            tested = g["runs"] > 0
            n_runs = max(g["runs"], 1) if tested else 0
            freq = round(100.0 * g["ev"][col_name] / n_runs, 0) if n_runs else None
            subj = "" if sev == "clean" and tested else ("!" if sev == "bad" else ("x" if sev == "event" else "w"))
            _set_cell(ws, r, start_c, subj, sev_fill(sev, tested))
            if start_c + 1 <= ws.max_column:
                _set_cell(ws, r, start_c + 1, freq if freq is not None else "w",
                          sev_fill(sev, tested))


def _fill_test_sheet(wb, key: str, variant: str, test_data: dict, vehicle: dict):
    sheet = TEST_SHEETS.get(key)
    if not sheet or sheet not in wb.sheetnames:
        return
    ws = wb[sheet]
    d = _test_def(key, variant)
    _fill_filename_cells(ws, d["prefix"], vehicle)
    rows = _run_rows(key, variant)
    cols = d["cols"]
    if key == "driveaway":
        _fill_driveaway_sheet(ws, test_data, variant)
    elif key == "gs_static":
        _fill_simple_grid(ws, rows, cols, test_data, row_col=2, start_search=7)
    elif key == "gs_rolling":
        _fill_simple_grid(ws, rows, cols, test_data, row_col=2, start_search=12)
    elif key in ("stationary_ess", "stationary_ac", "stationary_load", "stationary_thr"):
        _fill_simple_grid(ws, rows, cols, test_data, row_col=1, start_search=1)
    else:
        _fill_simple_grid(ws, rows, cols, test_data)


def export_session_to_excel(variant: str, vehicle: dict, data: dict) -> bytes:
    """Build an .xlsx workbook from the original template filled with session data."""
    template = resource_path(TEMPLATES[variant])
    wb = openpyxl.load_workbook(template, keep_vba=False, data_only=False)

    _fill_vehicle_summary(wb, vehicle, variant)

    for key in C.VARIANT_TESTS[variant]:
        test_data = data.get(key)
        if not test_data:
            continue
        _fill_test_sheet(wb, key, variant, test_data, vehicle)

    if "driveaway" in data:
        _fill_drb_driveaway_summary(wb, data["driveaway"], variant)

    # Metadata sheet for traceability.
    if "Export Info" in wb.sheetnames:
        ws = wb["Export Info"]
    else:
        ws = wb.create_sheet("Export Info")
    ws["A1"] = "Exported from DRB Python Tool"
    ws["A2"] = f"Variant: {variant}"
    ws["A3"] = f"Exported at: {datetime.now().isoformat(timespec='seconds')}"
    ws["A4"] = f"Vehicle token: {_vehicle_token(vehicle)}"

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.read()


def export_filename(variant: str, vehicle: dict) -> str:
    return f"{naming.export_basename(variant, vehicle)}.xlsx"
