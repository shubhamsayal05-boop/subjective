from html import escape
from typing import Any

from openpyxl.utils import get_column_letter

from app.conditional_formatting import rating_color
from app.utils import color_to_css, column_width_px, format_display_value, row_height_px
from app.workbook_manager import WorkbookManager


def _merged_map(ws) -> tuple[dict[tuple[int, int], tuple[int, int, int, int]], set[tuple[int, int]]]:
    merged: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    covered: set[tuple[int, int]] = set()
    for mrange in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = mrange.bounds
        merged[(min_row, min_col)] = (min_row, min_col, max_row, max_col)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) != (min_row, min_col):
                    covered.add((r, c))
    return merged, covered


def _cell_bg(cell, display_value: Any) -> str | None:
    if cell.fill and cell.fill.patternType and cell.fill.patternType != "none":
        bg = color_to_css(cell.fill.fgColor)
        if bg:
            return bg
    return rating_color(display_value)


def render_sheet_html(manager: WorkbookManager, sheet: str) -> str:
    ws = manager.wb[sheet]
    bounds = manager.sheet_bounds(sheet)
    if bounds is None:
        return "<p><em>This sheet is empty.</em></p>"

    min_r, max_r, min_c, max_c = bounds
    merged, covered = _merged_map(ws)

    col_widths = "".join(
        f"<col style='width:{column_width_px(ws, c)}px'>" for c in range(min_c, max_c + 1)
    )
    header_cells = "".join(
        f"<th>{get_column_letter(c)}</th>" for c in range(min_c, max_c + 1)
    )

    body_rows: list[str] = []
    for r in range(min_r, max_r + 1):
        cells: list[str] = []
        for c in range(min_c, max_c + 1):
            if (r, c) in covered:
                continue
            coord = f"{get_column_letter(c)}{r}"
            cell = ws[coord]
            display = manager.get_display_value(sheet, coord)
            text = escape(format_display_value(display))

            attrs = [f"title='{coord}'"]
            bg = _cell_bg(cell, display)
            if bg:
                attrs.append(f"style='background:{bg}'")
            if manager.is_formula_cell(sheet, coord):
                attrs.append("class='formula'")

            rowspan, colspan = 1, 1
            if (r, c) in merged:
                min_row, min_col, max_row, max_col = merged[(r, c)]
                rowspan = max_row - min_row + 1
                colspan = max_col - min_col + 1
            if rowspan > 1:
                attrs.append(f"rowspan='{rowspan}'")
            if colspan > 1:
                attrs.append(f"colspan='{colspan}'")

            cells.append(f"<td {' '.join(attrs)}>{text}</td>")

        body_rows.append(f"<tr><th class='row-hdr'>{r}</th>{''.join(cells)}</tr>")

    return f"""
    <div class="excel-sheet-wrap">
      <table class="excel-sheet">
        <colgroup><col class="row-col">{col_widths}</colgroup>
        <thead><tr><th class="row-hdr"></th>{header_cells}</tr></thead>
        <tbody>{''.join(body_rows)}</tbody>
      </table>
    </div>
    """


def render_sheet_css() -> str:
    return """
    <style>
      .excel-sheet-wrap {
        overflow: auto; max-height: 72vh;
        border: 1px solid #a6a6a6; background: #e6e6e6; padding: 6px;
      }
      .excel-sheet {
        border-collapse: collapse; background: white;
        table-layout: fixed; font-family: Calibri, Arial, sans-serif; font-size: 11px;
      }
      .excel-sheet td, .excel-sheet th {
        border: 1px solid #b4b4b4; padding: 1px 4px;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        max-width: 200px; box-sizing: border-box;
      }
      .excel-sheet thead th { background: #217346; color: white; font-size: 10px; }
      .excel-sheet .row-hdr, .excel-sheet .row-col { background: #f0f0f0; color: #333; font-size: 10px; width: 32px; }
      .excel-sheet td.formula { background: #f3f3f3; }
    </style>
    """
