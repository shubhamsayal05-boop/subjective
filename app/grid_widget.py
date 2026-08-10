"""Interactive sheet grid with double-click editing (streamlit-aggrid)."""

from typing import Any

import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from st_aggrid.shared import GridUpdateMode

from openpyxl.utils import get_column_letter

from app.cell_store import CellStore


def build_sheet_dataframe(store: CellStore, sheet: str) -> tuple[pd.DataFrame, list[str], dict[str, str]]:
    meta = store.sheet_meta(sheet)
    bounds = meta.get("bounds")
    if not bounds:
        return pd.DataFrame(), [], {}

    min_r, max_r, min_c, max_c = bounds
    columns = [get_column_letter(c) for c in range(min_c, max_c + 1)]
    rows: list[dict[str, Any]] = []

    for r in range(min_r, max_r + 1):
        row: dict[str, Any] = {"Row": r}
        for c in range(min_c, max_c + 1):
            letter = get_column_letter(c)
            coord = f"{letter}{r}"
            row[letter] = store.display_value(sheet, coord)
            row[f"__e_{letter}"] = store.is_editable(sheet, coord)
        rows.append(row)

    return pd.DataFrame(rows), columns, {"min_r": min_r, "min_c": min_c, "sheet": sheet}


_EDITABLE_JS = JsCode(
    """
    function(params) {
        const field = params.colDef.field;
        if (!field || field === 'Row') return false;
        const key = '__e_' + field;
        return params.data[key] === true;
    }
    """
)


def render_editable_grid(
    store: CellStore,
    sheet: str,
    height: int = 680,
) -> dict[str, str] | None:
    df, data_columns, ctx = build_sheet_dataframe(store, sheet)
    if df.empty:
        return None

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        editable=False,
        filterable=False,
        sortable=False,
        resizable=True,
        wrapText=False,
        autoHeight=False,
        singleClickEdit=False,
    )
    gb.configure_column("Row", editable=False, width=52, pinned="left", cellStyle={"backgroundColor": "#f0f0f0"})
    for col in data_columns:
        gb.configure_column(
            col,
            editable=_EDITABLE_JS,
            width=88,
            cellStyle=JsCode(
                f"""
                function(params) {{
                    if (params.data['__e_{col}']) {{
                        return {{backgroundColor: '#ffffff'}};
                    }}
                    return {{backgroundColor: '#f3f3f3'}};
                }}
                """
            ),
        )
    for col in data_columns:
        gb.configure_column(f"__e_{col}", hide=True)

    gb.configure_grid_options(
        domLayout="normal",
        suppressMovableColumns=True,
        enterNavigatesVertically=True,
        enterNavigatesVerticallyAfterEdit=True,
        stopEditingWhenCellsLoseFocus=True,
    )

    grid_options = gb.build()
    response = AgGrid(
        df,
        gridOptions=grid_options,
        height=height,
        width="100%",
        update_mode=GridUpdateMode.VALUE_CHANGED,
        allow_unsafe_jscode=True,
        theme="streamlit",
        key=f"aggrid_{ctx['sheet']}",
    )

    if response is None:
        return None

    updated = response.get("data")
    if updated is None:
        return None

    edits: dict[str, str] = {}
    min_r = ctx["min_r"]
    for row_data in updated:
        r = row_data.get("Row")
        if r is None:
            continue
        for col in data_columns:
            if not row_data.get(f"__e_{col}"):
                continue
            coord = f"{col}{r}"
            new_val = row_data.get(col, "")
            old_val = store.display_value(sheet, coord)
            new_str = "" if new_val is None else str(new_val)
            if new_str != old_val:
                edits[coord] = new_str

    return edits if edits else None
