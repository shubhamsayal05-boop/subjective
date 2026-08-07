from pathlib import Path

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"

VERSIONS = {
    "Base (General Transmission)": "Subjective_SprdSheet_072926.xlsm",
    "BEV (Battery Electric)": "BEV_Subjective_SprdSheet_072926.xlsm",
    "CVT (CVT Transmission)": "CVT Subjective_SprdSheet_072926.xlsm",
}

# Sheet display order and grouping (matches Excel tab bar layout)
SHEET_GROUPS = [
    ("Report", ["Test Summary -->", "DRB - Summary", "DRB - Color Chart", "Optional Test - Summary", "Engine - Summary"]),
    ("DRB Tests", ["DRB Testing -->", "Driveaway-sweeps", "Driveaway ESS", "Decel Cstdowns", "Decel OPD", "USS-Manual (Opt.)", "RTITO_AD", "TI_CstSpd ", "TO_CstSpd", "RRL (with HS)", "Garage Shifts", "Kickdowns"]),
    ("Engine Tests", ["ENG Testing -->", "TITO(Opt)", "CstSpd", "Acceleration(Opt)", "Stationary Test", "Engine - Color Chart"]),
    ("Reference", ["Pedal Geo", "Test Details -->", "PV Max", "Acronyms"]),
]


def get_template_path(version_label: str) -> Path:
    filename = VERSIONS[version_label]
    return TEMPLATES_DIR / filename


DEFAULT_SHEET = "DRB - Summary"


def ordered_sheets(sheetnames: list[str]) -> list[str]:
    known = set(sheetnames)
    ordered: list[str] = []
    for _, sheets in SHEET_GROUPS:
        for s in sheets:
            if s in known:
                ordered.append(s)
    for s in sheetnames:
        if s not in ordered:
            ordered.append(s)
    return ordered
