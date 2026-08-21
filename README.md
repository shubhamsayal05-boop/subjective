# DRB Subjective Drivability Tool (Python)

Standalone Python port of the 072926 Subjective Spreadsheets (AT / BEV / CVT).
No Excel required — all data is stored as JSON.

## Run (Python)
    pip install -r requirements.txt
    streamlit run app.py

## Build Windows .exe

**Important:** The build scripts (`build_exe.bat`, `drb_tool.spec`, etc.) are in the project repo. If you only copied `app.py` / `config.py`, the build will fail — download or clone the **full folder**, including the three `.xlsm` template files.

### Easiest — download pre-built .exe (no build needed)
1. Open GitHub → **Actions** → **Build Windows EXE**
2. Click the latest green run → scroll to **Artifacts**
3. Download **`DRB_Subjective_Tool-Windows`**
4. Unzip and run **`DRB_Subjective_Tool.exe`**

Direct link (latest successful run on the build branch):  
https://github.com/shubhamsayal05-boop/subjective/actions/workflows/build-exe.yml

### Option A — build on your Windows PC
1. Install [Python 3.10+](https://www.python.org/downloads/) and check **Add Python to PATH**.
2. Open the **full project folder** (must contain `build_exe.bat`, `drb_tool.spec`, and the `.xlsm` files).
3. Double-click **`build_exe.bat`** (or run **`build_exe.ps1`** in PowerShell).
4. Wait 5–10 minutes. When finished, open **`dist\DRB_Subjective_Tool.exe`**.

If the build fails, open **`build.log`** in the same folder — it contains the full error.

Common fixes:
- **"Python is not installed"** → Reinstall Python and check "Add to PATH"
- **"Missing required file"** → You don't have the full project folder
- **PyInstaller error** → Run `python -m pip install -r requirements.txt -r requirements-build.txt` then try again
- **Antivirus blocked the build** → Temporarily allow the project folder, then rebuild

### Option B — manual build (on Windows)
    pip install -r requirements.txt -r requirements-build.txt
    pyinstaller --noconfirm --clean drb_tool.spec

The executable launches the tool in your default web browser. Saved sessions are stored in a `sessions` folder next to the `.exe`.

**Offline / track use:** The tool runs entirely on your laptop — no Wi-Fi or internet is required. Double-click the `.exe`, wait for your browser to open at `http://127.0.0.1:8501`, and start testing. If your browser shows an "offline" banner, ignore it; localhost still works without internet.

### Option C — download from GitHub Actions
Open the **Actions** tab → **Build Windows EXE** → download the **`DRB_Subjective_Tool-Windows`** artifact from the latest green run.

## Workflow (matches the Excel sheets exactly)
1. **Home** — pick the transmission/propulsion (AT, BEV, CVT), enter vehicle info
   (feeds every INCA/AVL recorder label), save/load/import/export sessions.
2. **Test sheets** (sidebar) — same entry model as Excel: type shorthand
   defect codes in the event cells (`sj`, `!b`, `db-c`; blank = clean pass),
   tick **Complete (x)** per run row, enter the achieved **Top Gear** per run
   (driveaway). The Driveaway sheet carries the full pedal x brake-level run
   matrix (No Pedal ... 100%, No/Light/Normal/Medium brake) verbatim from the
   072926 template. A numeric 1-10 in a cell acts as an explicit rating.
3. **DRB Summary** — colors are DERIVED, like the spreadsheet: green =
   completed run with no code, yellow = event (marginal), red = `!` bad event,
   grey = not tested. Frequency = flagged runs / completed runs (<=25 green,
   26-49 yellow, >=50 red). The Block Pedal [Sweeps] Summary shows
   Subjective | Frequency pairs per shift plus TOP GEAR = MAX over that pedal
   step's runs (the Python equivalent of =MAX('Driveaway-sweeps'!Dxx:Dyy)).
   Exports: session JSON, **Excel (.xlsx)** filled from the original template,
   and **JPG** downloads on every table.
4. **Filenames** — full INCA/AVL label list (prefix + vehicle token + date).
5. **Acronyms** — the DRB shorthand dictionary (nd, b, c, st, bb, !, db, …).

Variant differences match the Excel tools: BEV gets Decel OPD + brake-pedal
coastdowns over speed bands (no ESS/Stationary/Kickdowns, driveaway columns
become accel bands); AT/CVT keep ESS, Stationary, Kickdowns and gear columns.

Sessions auto-name as MY_Line_Variant_VIN_date.json in ./sessions/.
Test grids/axes live in `config.py` — edit there to add pedals, speeds, or tests.
