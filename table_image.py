"""Render pandas tables as JPEG images for download."""
from __future__ import annotations

import io

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.table import Table


def _cell_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


def dataframe_to_jpeg(
    df: pd.DataFrame,
    *,
    title: str = "",
    cell_colors: dict[tuple[int, int], str] | None = None,
    index: bool = True,
) -> bytes:
    """Return a JPEG byte stream for a dataframe (optionally with per-cell colors)."""
    display = df.copy()
    if not index:
        display = display.reset_index(drop=True)

    nrows, ncols = display.shape
    if index:
        ncols += 1

    fig_w = min(max(6, ncols * 1.05), 48)
    fig_h = min(max(1.5, (nrows + 1) * 0.38 + (0.8 if title else 0)), 80)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=150)
    ax.axis("off")
    if title:
        ax.set_title(title, fontsize=11, fontweight="bold", pad=10)

    table: Table = ax.table(
        cellText=[[""] * ncols for _ in range(nrows + 1)],
        loc="center",
        cellLoc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(max(5, min(9, int(220 / max(nrows, ncols)))))

    headers = (["Row"] if index else []) + [str(c) for c in display.columns]
    for c, header in enumerate(headers):
        cell = table[(0, c)]
        cell.get_text().set_text(header)
        cell.set_facecolor("#1f4e79")
        cell.get_text().set_color("white")
        cell.get_text().set_weight("bold")

    for r in range(nrows):
        row_label = str(display.index[r]) if index else str(r + 1)
        col_offset = 0
        if index:
            cell = table[(r + 1, 0)]
            cell.get_text().set_text(row_label)
            cell.set_facecolor("#eef2f7")
            col_offset = 1
        for c, col_name in enumerate(display.columns):
            cell = table[(r + 1, c + col_offset)]
            text = _cell_text(display.iloc[r, c])
            cell.get_text().set_text(text)
            color = (cell_colors or {}).get((r, c))
            if color:
                cell.set_facecolor(color)
            elif (r + c) % 2 == 0:
                cell.set_facecolor("#fafafa")

    table.scale(1, 1.25)
    buf = io.BytesIO()
    fig.savefig(buf, format="jpeg", bbox_inches="tight", facecolor="white", pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf.read()
