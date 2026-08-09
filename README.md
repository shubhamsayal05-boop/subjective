# Subjective Spreadsheet Tool

**100% Python** — no Excel runtime, no slow formula engine startup.

Templates, formulas, and layout are bundled as JSON (`app/bundled/*.json`) with a native Python formula evaluator (~0.2s recalc).

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Features

- **Instant load** (~0.01s data + ~0.2s formulas)
- **Double-click** cells to edit (Excel-like grid)
- **Base / BEV / CVT** versions on main screen
- **Download .xlsx** export

## Editing

1. Double-click a white (editable) cell
2. Type value (e.g. `b`, `nd`, `!c`)
3. Enter to save — summaries update automatically

## Re-extract templates (optional)

If you update the source `.xlsm` files in `templates/`:

```bash
python scripts/extract_bundles.py
```

## Architecture

| File | Purpose |
|------|---------|
| `app/bundled/*.json` | Sheet data, styles, formulas (pre-extracted) |
| `app/cell_store.py` | In-memory workbook |
| `app/formula_engine.py` | Python IF/COUNTIF/COUNTA/MAX/INDEX/MATCH/… |
| `components/excel_grid/` | Double-click grid UI |
| `app.py` | Streamlit app |
