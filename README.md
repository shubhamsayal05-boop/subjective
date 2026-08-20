# DRB Subjective Drivability Tool (Python)

Standalone Python port of the 072926 Subjective Spreadsheets (AT / BEV / CVT).
No Excel required — all data is stored as JSON.

## Run
    pip install -r requirements.txt
    streamlit run app.py

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
   Exports: session JSON + flat findings CSV.
4. **Filenames** — full INCA/AVL label list (prefix + vehicle token + date).
5. **Acronyms** — the DRB shorthand dictionary (nd, b, c, st, bb, !, db, …).

Variant differences match the Excel tools: BEV gets Decel OPD + brake-pedal
coastdowns over speed bands (no ESS/Stationary/Kickdowns, driveaway columns
become accel bands); AT/CVT keep ESS, Stationary, Kickdowns and gear columns.

Sessions auto-name as MY_Line_Variant_VIN_date.json in ./sessions/.
Test grids/axes live in `config.py` — edit there to add pedals, speeds, or tests.
