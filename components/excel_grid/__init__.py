import os

import streamlit.components.v1 as components

_FRONTEND = os.path.join(os.path.dirname(__file__), "frontend")
_excel_grid = components.declare_component("excel_grid", path=_FRONTEND)


def excel_grid(grid_data: dict, height: int = 720, key: str | None = None) -> dict | None:
    """Interactive spreadsheet grid with double-click in-cell editing."""
    result = _excel_grid(grid_data=grid_data, height=height, key=key, default=None)
    if result is None:
        return None
    return dict(result)
