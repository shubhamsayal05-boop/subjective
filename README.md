# Subjective Spreadsheet Tool (Streamlit)

Python Streamlit version of the subjective driveability test Excel spreadsheets. One unified app with the same worksheets, formulas, layout, and design as the original `.xlsm` files.

## Versions

On the main screen, choose which template to work with:

| Version | Original file |
|---------|----------------|
| Base (General Transmission) | `Subjective_SprdSheet_072926.xlsm` |
| BEV (Battery Electric) | `BEV_Subjective_SprdSheet_072926.xlsm` |
| CVT (CVT Transmission) | `CVT Subjective_SprdSheet_072926.xlsm` |

## Features

- Same sheet names and tab order as Excel
- Excel-like grid preview (merged cells, colors, column widths)
- Live formula recalculation (DRB Summary, Color Charts, filenames, etc.)
- Cell editor and batch edit for subjective entry sheets
- PV Max dropdown for transmission selector (cell A1)
- Download updated workbook as `.xlsm` (formulas preserved)

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

## Usage

1. Select **Base**, **BEV**, or **CVT** at the top of the main page.
2. Open a worksheet from the sidebar (same names as Excel tabs).
3. Enter subjective ratings and metadata using **Edit cell values** or **Quick edit**.
4. Summary sheets (`DRB - Summary`, color charts) update after each apply/recalculate.
5. Use **Download Excel (.xlsm)** to save your work.

## Project layout

```
app.py                  # Streamlit entry point
app/
  config.py             # Version paths and sheet groups
  workbook_manager.py   # Load/edit/recalculate/export
  html_renderer.py      # Excel-like HTML grid
  conditional_formatting.py
  utils.py
templates/              # Original .xlsm templates (unchanged)
```

## Notes

- First load of a version takes ~20–30 seconds while the formula engine initializes (cached afterward).
- Recalculation after edits typically takes ~1–2 seconds.
- Open downloaded `.xlsm` files in Excel to see fully calculated values with native Excel rendering.
