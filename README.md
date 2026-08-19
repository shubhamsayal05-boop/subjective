# Subjective Spreadsheet Tool

**100% Python** — no Excel runtime, no slow formula engine startup.

Templates, formulas, and layout are bundled as JSON (`app/bundled/*.json`) with a native Python formula evaluator (~0.2s recalc).

## Windows .exe (standalone app)

### Download pre-built .exe (easiest)

1. Open the [Actions tab](https://github.com/shubhamsayal05-boop/subjective/actions/workflows/build-windows-exe.yml) on GitHub.
2. Click the latest successful **Build Windows EXE** run.
3. Under **Artifacts**, download `SubjectiveSpreadsheetTool-windows-zip`.
4. Extract the zip, then double-click `SubjectiveSpreadsheetTool\SubjectiveSpreadsheetTool.exe`.

### Build locally on Windows

Python 3.10+ required only for building:

```bat
build_windows.bat
```

Output:

```
dist\SubjectiveSpreadsheetTool\SubjectiveSpreadsheetTool.exe
```

- Copy the whole `SubjectiveSpreadsheetTool` folder to any Windows machine.
- Double-click `SubjectiveSpreadsheetTool.exe` — browser opens automatically.
- **No Python or Excel** needed on the target PC.
- Close the console window to stop the app.

Manual build:

```bat
pip install -r requirements-build.txt
pyinstaller build_exe.spec --noconfirm
```

## Run (developer)

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
| `app/grid_widget.py` | AgGrid double-click editor |
| `app.py` | Streamlit app |
