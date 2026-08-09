from app.cell_store import CellStore, _coord_str
from app.conditional_formatting import rating_color


def build_grid_data(store: CellStore, sheet: str, height: int = 700) -> dict:
    meta = store.sheet_meta(sheet)
    bounds = meta.get("bounds")
    if not bounds:
        return {
            "sheet": sheet,
            "minRow": 1,
            "maxRow": 1,
            "minCol": 1,
            "maxCol": 1,
            "cells": {},
            "merges": [],
            "colWidths": {},
            "height": height,
        }

    min_r, max_r, min_c, max_c = bounds
    cells: dict[str, dict] = {}
    sheet_cells = meta.get("cells", {})

    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            coord = _coord_str(c, r)
            cell_def = sheet_cells.get(coord, {})
            display = store.display_value(sheet, coord)
            editable = store.is_editable(sheet, coord)
            info: dict = {"v": display, "e": editable}
            style = cell_def.get("style", {})
            bg = style.get("bg") or rating_color(display)
            if bg:
                info["bg"] = bg
            cells[coord] = info

    return {
        "sheet": sheet,
        "minRow": min_r,
        "maxRow": max_r,
        "minCol": min_c,
        "maxCol": max_c,
        "cells": cells,
        "merges": meta.get("merges", []),
        "colWidths": meta.get("col_widths", {}),
        "height": height,
    }
