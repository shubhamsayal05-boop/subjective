# Subjective Spreadsheet Tool (Streamlit)

Standalone **web-based** subjective driveability test tool. No Microsoft Excel installation required.

## Versions (select on main screen)

| Version | Use case |
|---------|----------|
| Base (General Transmission) | Standard template |
| BEV (Battery Electric) | Battery electric vehicles |
| CVT (CVT Transmission) | CVT transmissions |

## Editing (Excel-like)

- **Double-click** any editable (white) cell to edit
- **Enter** or click away to save
- **Tab** moves to the next editable cell
- Grey cells are formula/calculated — use **Recalculate** after edits

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501

## Performance

- **First open:** ~5 seconds (spreadsheet loads from bundled templates)
- **Recalculate:** first time loads formula engine (~25s, cached); then ~2s
- Without recalculate, saved template values are shown instantly

## Export

Download `.xlsm` from the sidebar. Open in Excel only if you need native Excel printing — the web tool is fully standalone.

## Structure

```
app.py                     # Streamlit UI
components/excel_grid/     # Double-click in-cell editor
app/workbook_manager.py    # Data + optional live formulas
app/grid_builder.py        # Grid JSON for the editor
templates/                 # Bundled spreadsheet templates (internal)
```
