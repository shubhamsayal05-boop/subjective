"""
DRB Subjective Drivability Tool — configuration
All structures extracted from Subjective_SprdSheet_072926.xlsm (Base/ICE-AT),
BEV_Subjective_SprdSheet_072926.xlsm and CVT_Subjective_SprdSheet_072926.xlsm.
"""

APP_TITLE = "DRB Subjective Drivability Tool"
TEMPLATE_REV = "072926 (Python port)"

# ----------------------------------------------------------------------------- variants
VARIANTS = {
    "AT (Step Automatic / ICE)": "AT",
    "BEV (Battery Electric)": "BEV",
    "CVT (Continuously Variable)": "CVT",
}

# ----------------------------------------------------------------------------- rating scales
RATING_OPTIONS = ["w"] + [str(i) for i in range(1, 11)]   # 'w' = waiting / not tested

SUBJ_GREEN = 8      # >= 8  acceptable
SUBJ_YELLOW = 6     # 6-7   marginal ; <=5 unacceptable
FREQ_GREEN = 25     # <= 25 % acceptable
FREQ_YELLOW = 49    # 26-49 % marginal ; >= 50 % unacceptable

COLORS = {
    "green": "#00b050",    # Excel-style strong green
    "yellow": "#ffc000",   # strong amber
    "red": "#ff3b30",      # strong red
    "grey": "#d6d6d6",
    "white": "#ffffff",
}

# ----------------------------------------------------------------------------- shared axes
UPSHIFTS = ["Launch", "1-2", "2-3", "3-4", "4-5", "5-6", "6-7", "7-8", "8-9"]
COASTDOWNS = ["9-8", "8-7", "7-6", "6-5", "5-4", "4-3", "3-2", "2-1"]
DRIVEAWAY_COLS = UPSHIFTS + COASTDOWNS          # one grid: Upshifts + Coastdowns, like Excel

# RTITO_AD: TI/TO speed pairs (last two pairs orange/optional in Excel)
RTITO_COLS = ["TI 6mph/10kph", "TO 31mph/50kph", "TI 19mph/30kph", "TO 43mph/70kph",
              "TI 31mph/50kph", "TO 56mph/90kph", "TI 43mph/70kph", "TO 68mph/110kph",
              "TI 56mph/90kph", "TO 78mph/125kph", "TI 68mph/110kph (Opt.)", "TO 86mph/140kph (Opt.)",
              "TI 81mph/130kph (Opt.)", "TO 100mph/160kph (Opt.)"]
TI_CST_COLS = ["TI 6mph/10kph", "TI 19mph/30kph", "TI 31mph/50kph", "TI 43mph/70kph",
               "TI 56mph/90kph", "TI 68mph/110kph (Opt.)", "TI 81mph/130kph (Opt.)"]
TO_CST_COLS = ["TO 28mph/45kph", "TO 37mph/60kph", "TO 50mph/80kph", "TO 62mph/100kph",
               "TO 75mph/120kph", "TO 81mph/130kph (Opt.)", "TO 93mph/150kph (Opt.)"]

def _runs4(label):        # Fast x2 + Slow x2, Excel row pattern
    return [f"{label} Fast #1", f"{label} Fast #2", f"{label} Slow #1", f"{label} Slow #2"]

TITO_ROWS = (_runs4("20%") + _runs4("30%") + _runs4("40%") + _runs4("50%") + _runs4("70%")
             + ["100%/detent #1", "100%/detent #2", "100%/detent active #1", "100%/detent active #2"])

TITO_RPM_COLS = ["TI 1500", "TO 2500", "TI 2000", "TO 3000", "TI 2500", "TO 3500",
                 "TI 3000", "TO 4000", "TI 3500", "TO 4500", "TI 4000", "TO 5000",
                 "TI 4500", "TO 5500", "TI 5000", "TO 6000"]

def _gears(pedal, gears):
    out = []
    for g in gears:
        out += [f"{pedal} · {g} #1", f"{pedal} · {g} #2"]
    return out

TITO_OPT_ROWS = (_gears("25%", ["2nd", "3rd"]) + _gears("50%", ["2nd", "3rd", "4th"])
                 + _gears("75%", ["2nd", "3rd"]) + _gears("100%", ["2nd", "3rd", "4th"]))


# ---- BEV / CVT driveaway speed-range columns (extracted from the workbooks;
#      note: the template genuinely has no 35-20 decel band)
SPEED_ACCEL_CVT = ["Launch (0-5 mph / 0-8 kph)", "5-20 mph (8-32 kph)", "20-35 mph (32-56 kph)",
                   "35-50 mph (56-81 kph)", "50-65 mph (81-105 kph)", "65-75 mph (105-121 kph)",
                   "75-85 mph (121-137 kph)", "85-93 mph (137-150 kph) (Opt.)"]
SPEED_DECEL_CVT = ["93-85 mph (150-137 kph) (Opt.)", "85-75 mph (137-121 kph)",
                   "75-65 mph (121-105 kph)", "65-50 mph (105-81 kph)",
                   "50-35 mph (81-56 kph)", "20-5 mph (32-8 kph)", "5-0 mph (8-0 kph)"]
SPEED_ACCEL_BEV = ["Launch (0-5 mph / 0-8 kph)", "5-20 mph (8-32 kph)", "20-35 mph (32-56 kph)",
                   "35-50 mph (56-81 kph)", "50-65 mph (81-105 kph)", "65-75 mph (105-121 kph)",
                   "75-85 mph (121-137 kph) (Opt.)", "85-93 mph (137-150 kph) (Opt.)"]
SPEED_DECEL_BEV = ["93-85 mph (150-137 kph) (Opt.)", "85-75 mph (137-121 kph) (Opt.)",
                   "75-65 mph (121-105 kph)", "65-50 mph (105-81 kph)",
                   "50-35 mph (81-56 kph)", "20-5 mph (32-8 kph)", "5-0 mph (8-0 kph)"]
BEV_DECEL_BANDS = ["90-75 mph (145-121 kph)", "75-60 mph (121-97 kph)", "60-45 mph (97-72 kph)",
                   "45-30 mph (72-48 kph)", "30-15 mph (48-24 kph)", "15-0 mph (24-0 kph)"]

BEV_SPEED_BANDS = ["90-75 mph", "75-60 mph", "60-45 mph", "45-30 mph", "30-15 mph", "15-0 mph"]
BEV_BRAKE_PEDALS = ["0%", "10%", "15%", "20%", "25%", "30%", "35%", "40%", "50%", "70%", "100%"]
OPD_PEDALS = ["5%", "10%", "15%", "20%", "25%", "30%"]

ESS_PEDALS = ["0%", "5%", "10%", "15%", "20%", "25%", "30%", "35%", "40%", "45%",
              "50%", "60%", "70%", "80%", "100%"]
ESS_ROWS = ["0% · Run 1", "0% · Run 2"] + [f"{p} · Run {i}" for p in ESS_PEDALS[1:] for i in (1, 2, 3, 4)]

USS_ROWS = [f"{p} · {m} #{i}" for p in ("30%", "50%", "100%") for m in ("Normal", "Sport") for i in (1, 2)]

# ----------------------------------------------------------------------------- test definitions
TESTS = {
    "driveaway": dict(
        name="Driveaway – Sweeps",
        rows=None, cols=DRIVEAWAY_COLS,
        prefix="DriveAway_10_1_",
        note="Accel & Brake Patterns → Upshifts (Launch…8-9) then Coastdowns (9-8…2-1) "
             "in the same run. If possible perform all sweeps on the same day. "
             "NOTE: scrub the response-delay criteria on driveaways for ramp pedals.",
    ),
    "driveaway_ess": dict(
        name="Driveaway – ESS (Engine Start/Stop)",
        rows=ESS_ROWS, cols=["Start", "Launch"],
        prefix="DriveAway_ESS_10_1_",
        note="ESS launches, 10 s stops. Select Drive, accelerate to 10 mph, brake to "
             "auto-stop, release and launch. Brake-release→tip-in timing: fast <300 ms "
             "(white) · 500 ms–1 s per header note. Rate the engine Start and the Launch.",
    ),
    "decel_cstdown": dict(
        name="Decel – Coastdowns (USS 0%)",
        rows=["No Brake #1", "No Brake #2",
              "-0.7 m/s² (-0.07g) #1", "-0.7 m/s² (-0.07g) #2",
              "-1.6 m/s² (-0.16g) #1", "-1.6 m/s² (-0.16g) #2",
              "-2.4 m/s² (-0.24g) #1", "-2.4 m/s² (-0.24g) #2"],
        cols=COASTDOWNS,
        prefix="Decel_nobrk_1_",
        note="Closed-pedal (USS 0%) coastdowns; two runs per braking level. "
             "Rate each downshift 9-8 … 2-1.",
    ),
    "uss_manual": dict(
        name="USS – Manual Upshifts (Opt.)",
        rows=[  # gas block
              "30% gas (US2500/CD2000) · Normal", "30% gas (US2500/CD2000) · Sport",
              "50% gas (US4000/CD2500) · Normal", "50% gas (US4000/CD2500) · Sport",
              "100% gas (Peak/No brake) · Normal", "100% gas (Peak/No brake) · Sport",
                # diesel block
              "30% diesel (US2000/CD1500) · Normal", "30% diesel (US2000/CD1500) · Sport",
              "50% diesel (US3000/CD1500) · Normal", "50% diesel (US3000/CD1500) · Sport",
              "100% diesel (Peak/CD2000) · Normal", "100% diesel (Peak/CD2000) · Sport"],
        cols=DRIVEAWAY_COLS,
        prefix="USS_Manual_",
        note="Optional manual upshift + coastdown scoring, Normal & Sport modes. "
             "Gas block (US 2500/4000 rpm) and diesel block (US 2000/3000 rpm) — "
             "fill the block matching the powertrain.",
    ),
    "rtito": dict(
        name="RTITO_AD – Rolling TI/TO at Deceleration",
        rows=TITO_ROWS, cols=RTITO_COLS,
        prefix="RTITO_AD_6_20pct_",
        note="TI at speed, TO at the paired speed while decelerating. Fast ×2 then "
             "Slow ×2 per pedal; cooldown lap after 40% and 70%. Orange pairs optional.",
    ),
    "ti_cstspd": dict(
        name="Tip-In at Constant Speed",
        rows=TITO_ROWS, cols=TI_CST_COLS,
        prefix="TI_CstSpd_6_20pct_",
        note="Stabilise at each speed, tip in to target pedal. Fast ×2 / Slow ×2 per "
             "pedal; 68 & 81 mph columns optional.",
    ),
    "to_cstspd": dict(
        name="Tip-Out at Constant Speed",
        rows=["100% TO #1", "100% TO #2", "100% TO #3", "100% TO #4"],
        cols=TO_CST_COLS,
        prefix="TO_CstSpd_28_RL_",
        note="From 100% pedal (Ctrl+U detent) stabilised at speed, tip out; rate the "
             "decel transition. 81 & 93 mph columns optional.",
    ),
    "rrl": dict(
        name="RRL – Rolling Re-Launch",
        rows=["30% · 30→3 mph #1", "30% · 30→3 mph #2",
              "30% · 30→6 mph #1", "30% · 30→6 mph #2",
              "50% · 30→3 mph #1", "50% · 30→3 mph #2",
              "50% · 30→6 mph #1", "50% · 30→6 mph #2"],
        cols=["TI", "TO"],
        prefix="RRL_30pct_3mph_",
        note="Rolling start: decel 30 mph to target speed, tip in before 5 kph. "
             "Rate TI and TO.",
    ),
    "hill_start": dict(
        name="Hill Start (Opt., Maserati)",
        rows=["25% · D · 15%", "25% · R · 15%", "25% · D · 22%",
              "50% · D · 15%", "50% · D · 22%",
              "75% · D · 15%", "75% · D · 22%",
              "100% · D · 15%", "100% · D · 22%"],
        cols=["TI", "TO"],
        prefix="RRL_HS_",
        note="Optional hill-start driveaway on 15% / 22% gradients, D and R.",
    ),
    "gs_static": dict(
        name="Garage Shifts – Static",
        rows=["0%, Firm Brake [Static]"],
        cols=["P-R", "R-P", "P-D", "D-P", "N-D", "D-N", "N-R", "R-N", "D-R", "R-D"],
        prefix="GS_Static_DRD_",
        note="Stationary, firm brake. PRP · PDP · NDN · NRN · DRD shift pairs.",
    ),
    "gs_rolling": dict(
        name="Garage Shifts – Rolling",
        rows=["0% Released Brake (3mph/5kph)", "0% Released Brake (max creep)",
              "10% Released Brake (3mph/5kph)"],
        cols=["D-R", "R-D"],
        prefix="GS_RL_0pct_",
        note="Rolling DRD at 3 mph / max creep / 10% pedal.",
    ),
    "gs_extra": dict(
        name="Garage Shifts – Extra (Opt.)",
        rows=["P-D [0%]", "D-R [0%]", "R-D [0%]", "P-R [0%]", "N-D [0%]",
              "P-D [5%]", "D-R [5%]", "R-D [5%]", "P-R [5%]", "N-D [5%]"],
        cols=["Shift"],
        prefix="Extra_GS_0%_",
        note="Optional extra garage shifts at 0% and 5% pedal.",
    ),
    "tito_opt": dict(
        name="TITO (Opt.) – Engine Tip-In/Tip-Out",
        rows=TITO_OPT_ROWS, cols=TITO_RPM_COLS,
        prefix="TITO_M2_25pct_",
        note="Fixed-gear TI/TO across the RPM ladder (min TI 1500, max TO 6500 rpm).",
    ),
    "cstspd_manual": dict(
        name="Constant Speed – Manual Mode",
        # per-speed gear subset, verbatim from the sheet (each column tests only its gears)
        rows=["Gear set A", "Gear set B", "Gear set C", "Gear set D"],
        cols=["Creep", "18.6mph/30kph", "31.1mph/50kph", "43.5mph/70kph", "62.1mph/100kph"],
        prefix="CS_M12_",
        col_gears={
            "Creep":            ["M1", "M2", "M1", "M2"],
            "18.6mph/30kph":    ["M2", "M3", "M2", "M3"],
            "31.1mph/50kph":    ["M3", "M4", "M3", "M4"],
            "43.5mph/70kph":    ["M4", "M5", "M4", "M5"],
            "62.1mph/100kph":   ["M5", "M6", "M7", "M5", "M6", "M7"],  # 62 mph adds M6/M7
        },
        note="Hold each speed in the gears listed for that column (Ctrl+U detent). "
             "Cell label shows which manual gear the run is in.",
    ),
    "cstspd_drive": dict(
        name="Constant Speed – In Drive",
        rows=["D"],
        cols=["Creep", "6mph/10kph", "12mph/20kph", "25mph/40kph", "37mph/60kph",
              "50mph/80kph", "62mph/100kph", "75mph/120kph"],
        prefix="CstSpd_D_25mph_",
        note="Hold each speed in D; rate hunting, surge, drivability.",
    ),
    "accel_fl": dict(
        name="Acceleration – Full Load (Opt.)",
        rows=["1st", "2nd", "3rd", "4th", "5th"],
        cols=["Full Load #1", "Full Load #2"],
        prefix="FullLoad_M1_",
        note="Full-load pulls per gear, two runs each.",
    ),
    "accel_plcp": dict(
        name="Acceleration – Part Load Constant Pedal (Opt.)",
        rows=["20% #1", "20% #2", "30% #1", "30% #2", "40% #1", "40% #2",
              "50% #1", "50% #2", "75% #1", "75% #2", "90% #1", "90% #2"],
        cols=["2nd Gear", "3rd Gear"],
        prefix="PLCP_M2_20pct_",
        note="Constant-pedal part-load accelerations in 2nd and 3rd.",
    ),
    "accel_plrp": dict(
        name="Acceleration – Part Load Rising Pedal (Opt.)",
        rows=["Run #1", "Run #2"],
        cols=["2nd Gear", "3rd Gear"],
        prefix="PLRP_M2_",
        note="Rising-pedal part-load accelerations.",
    ),
    "stationary_ess": dict(
        name="Stationary – Engine Manual Start/Stop",
        rows=["1st", "2nd", "3rd (AC ON)", "4th (AC ON)"],
        cols=["Engine On", "AC ON", "AC OFF", "Engine OFF"],
        prefix="EngineOnOff_",
        note="Park, manual start/stop attempts. Not done on hybrids.",
    ),
    "stationary_ac": dict(
        name="Stationary – Idle AC On/Off",
        rows=["1st", "2nd"], cols=["AC ON", "AC OFF"],
        prefix="IdleACOnOff_",
        note="Idle AC on/off attempts.",
    ),
    "stationary_load": dict(
        name="Stationary – Idle Electric Load",
        rows=["1st", "2nd"], cols=["Load ON", "Load OFF"],
        prefix="IdleElectricLoad_",
        note="Idle electric loading attempts.",
    ),
    "stationary_thr": dict(
        name="Stationary – Throttle Response (Opt.)",
        rows=["25% #1", "25% #2", "50% #1", "50% #2", "75% #1", "75% #2"],
        cols=["Throttle Response"],
        prefix="ThrRes_25pct_",
        note="Stationary throttle-response blips per pedal step.",
    ),
    "kickdowns": dict(
        name="Kickdowns",
        rows=["1 gear kickdown Fast Tip In", "1 gear kickdown Gradual Tip In",
              "2 gear kickdown Fast Tip In", "2 gear kickdown Gradual Tip In",
              "3 gear kickdown Fast Tip In", "3 gear kickdown Gradual Tip In",
              "4 gear kickdown Fast Tip In", "4 gear kickdown Gradual Tip In",
              "5 gear kickdown Fast Tip In", "5 gear kickdown Gradual Tip In"],
        cols=["10mph", "20mph", "30mph", "40mph", "50mph", "60mph", "70mph"],
        prefix="ThrRes_25pct_",
        note="Note any issues with kickdowns (Excel: free-text issue grid).",
    ),
    # ---- BEV-specific
    "decel_cstdown_bev": dict(
        name="Decel – Coastdowns (constant brake pedal)",
        rows=[f"{p} #{i}" for p in ("0%", "10%", "15%", "20%", "25%", "30%", "35%", "40%")
              for i in (1, 2)],
        cols=BEV_DECEL_BANDS,
        prefix="Decel_nobrk_1_",
        note="Constant brake pedal (%) through each deceleration band, two runs per "
             "pedal. NOTE: do all regen pedals.",
    ),
    "decel_opd": dict(
        name="Decel – OPD (One-Pedal Drive)",
        rows=[f"{p} #{i}" for p in ("5%", "10%", "15%", "20%", "25%", "30%") for i in (1, 2)],
        cols=BEV_DECEL_BANDS,
        prefix="Decel_OPD_1_",
        note="Constant speed to tip-out pedal (%), rate regen decel through each band, "
             "two runs per pedal. NOTE: do all regen pedals.",
    ),
}

# extra INCA labels for the filename generator page
EXTRA_PREFIXES = {
    "RTITO_AD extra 10/20%": "Extra_RTITO_AD_10_20pct",
    "GS DRD after movement": "Extra_GS_DRD_AftMov_",
    "GS RNDN": "GS_RNDN_",
    "GS RL max creep": "GS_RLMC_0pct_",
    "GS RL 10%": "GS_RL_10pct_",
    "GS extra 5%": "Extra_GS_5%_",
    "Idle AC On/Off": "IdleACOnOff_",
    "Idle electric load": "IdleElectricLoad_",
}

# ----------------------------------------------------------------------------- variant → tests
_AT_CVT = ["driveaway", "driveaway_ess", "decel_cstdown", "uss_manual", "rtito",
           "ti_cstspd", "to_cstspd", "rrl", "hill_start", "gs_static", "gs_rolling",
           "gs_extra", "tito_opt", "cstspd_manual", "cstspd_drive", "accel_fl",
           "accel_plcp", "accel_plrp", "stationary_ess", "stationary_ac",
           "stationary_load", "stationary_thr", "kickdowns"]
VARIANT_TESTS = {
    "AT":  list(_AT_CVT),
    "CVT": list(_AT_CVT),
    "BEV": ["driveaway", "decel_cstdown_bev", "decel_opd", "rtito", "ti_cstspd",
            "to_cstspd", "rrl", "gs_static", "gs_rolling", "cstspd_drive", "accel_plcp"],
}

# Per-variant test overrides (extracted from the BEV / CVT workbooks):
# both replace the gear-shift columns with speed ranges; the DRB Summary's last
# column is "Speed" (top speed reached) instead of "Gear".
VARIANT_OVERRIDES = {
    "BEV": {"driveaway": dict(cols=SPEED_ACCEL_BEV + SPEED_DECEL_BEV,
                              note="Accel & Brake Patterns → Acceleration Speed Ranges, "
                                   "then Deceleration Speed Ranges, in the same run. "
                                   "Same-day sweeps if possible; scrub response-delay "
                                   "criteria on ramp pedals.")},
    "CVT": {"driveaway": dict(cols=SPEED_ACCEL_CVT + SPEED_DECEL_CVT,
                              note="Accel & Brake Patterns → Upshifts (speed ranges), "
                                   "then Coastdowns (speed ranges), in the same run. "
                                   "Same-day sweeps if possible; scrub response-delay "
                                   "criteria on ramp pedals.")},
}
TOPGEAR_LABEL = {"AT": "TOP GEAR", "BEV": "TOP SPEED", "CVT": "TOP SPEED"}
BEV_OVERRIDES = VARIANT_OVERRIDES["BEV"]          # backward-compat alias

# ----------------------------------------------------------------------------- acronyms
ACRONYMS = [
    ("nosedive", "nd", ""),
    ("bump", "b", ""),
    ("clunk", "c", ""),
    ("stumble", "st", "feeling nosedive, bobble, bump, delay all at once"),
    ("bobble", "bb", "two bumps in quick succession"),
    ("delay", "d", ""),
    ("deceleration feeling", "decel", "slight G decrease, decrease in acceleration feel"),
    ("shudder", "sh", ""),
    ("jerk", "j", ""),
    ("push", "p", ""),
    ("lunge", "l", "longer push, elongated g increase"),
    ("wheel spin", "ws", ""),
]
MODIFIERS = [
    ("late-***", "l", "within 1–5 s after maneuver"),
    ("slight-***", "s", "a slight event"),
    ("double-***", "db", "same event back to back"),
    ("bad event", "!", "prefix — e.g. !b = bad bump"),
    ("very bad event", "!!", ""),
]

VEHICLE_FIELDS = ["Model Year", "Vehicle Line", "Engine Disp.", "Transmission",
                  "Model Code", "Prototype Level", "Software Level", "Last 4 of VIN"]

# Display labels for the Home page (internal dict keys stay as in VEHICLE_FIELDS).
VEHICLE_FIELD_LABELS = {
    "Vehicle Line": "Vehicle Line (LB)",
    "Model Code": "Model Code (optional — not used in filenames)",
}


# ----------------------------------------------------------------------------- Excel-faithful entry model
# Cells hold SHORTHAND DEFECT CODES (e.g. "sj", "!b", "db-c"), NOT numbers.
# Blank cell on a completed run = clean pass (green on summary).
# Severity: any "!" in a code -> red (bad event); otherwise -> yellow (event/marginal).
# A numeric cell (1-10) is accepted as an explicit subjective override.
# Complete column: mark "x" per run row. TOP GEAR: achieved gear per run row
# (summary shows MAX per pedal step, exactly like =MAX('Driveaway-sweeps'!Dxx:Dyy)).

# Run rows — verbatim from Subjective_SprdSheet_072926.xlsm 'Driveaway-sweeps'
# ptype: "Step" (white rows, step-in pedal) or "Ramp" (blue rows, ramp-in pedal:
# brake to tip-in). base: row fill in the template (white / blue / orange for the
# optional 99% pre-detent block). pv: pedal-voltage reference (column A).
DRIVEAWAY_RUNS = [
    dict(pedal="No Pedal", brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="5%",   brake="No Brake",     ptype="Step", base="white"),
    dict(pedal="5%",   brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="10%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="10%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="10%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="10%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="15%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="15%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="15%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="15%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="20%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="20%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="20%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="20%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="25%",  brake="No Brake",     ptype="Step", base="white"),
    dict(pedal="25%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="25%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="25%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="30%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="30%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="30%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="30%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="35%",  brake="No Brake",     ptype="Step", base="white"),
    dict(pedal="35%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="35%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="35%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="40%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="40%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="40%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="40%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="45%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="45%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="45%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="45%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="50%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="50%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="50%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="50%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="60%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="60%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="60%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="60%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="70%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="70%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="70%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="70%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="80%",  brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="80%",  brake="Normal Brake", ptype="Step", base="white"),
    dict(pedal="80%",  brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="80%",  brake="Medium Brake", ptype="Ramp", base="blue"),
    dict(pedal="99% pre-detent (Opt.)", brake="Light Brake",  ptype="Step", base="orange"),
    dict(pedal="99% pre-detent (Opt.)", brake="Normal Brake", ptype="Step", base="orange"),
    dict(pedal="99% pre-detent (Opt.)", brake="Normal Brake", ptype="Ramp", base="orange"),
    dict(pedal="99% pre-detent (Opt.)", brake="Medium Brake", ptype="Ramp", base="orange"),
    dict(pedal="100% / through detent", brake="No Brake",     ptype="Step", base="white"),
    dict(pedal="100% / through detent", brake="Light Brake",  ptype="Step", base="white"),
    dict(pedal="100% / through detent", brake="Normal Brake", ptype="Ramp", base="blue"),
    dict(pedal="100% / through detent", brake="Medium Brake", ptype="Ramp", base="blue"),
]
DRIVEAWAY_PEDAL_ORDER = ["No Pedal", "5%", "10%", "15%", "20%", "25%", "30%", "35%",
                         "40%", "45%", "50%", "60%", "70%", "80%",
                         "99% pre-detent (Opt.)", "100% / through detent"]

# Pedal-voltage reference (column A of the sheet); WOT = 3.832 V
DRIVEAWAY_PV = {"5%": 0.205, "10%": 0.396, "15%": 0.596, "20%": 0.777, "25%": 0.958,
                "30%": 1.158, "35%": 1.354, "40%": 1.535, "45%": 1.740, "50%": 1.916,
                "60%": 2.297, "70%": 2.874, "80%": 3.055}
DRIVEAWAY_WOT_PV = 3.832

ROW_BASE_COLORS = {"white": "#ffffff", "blue": "#c6d9f0", "orange": "#ffe08a"}

# tests that use the multi-run (pedal x brake) layout with a TOP GEAR column
RUN_TESTS = {"driveaway"}


# ----------------------------------------------------------------------------- per-row comment column
# Tests whose Excel sheet has a per-row "Comments" column (verified across AT/BEV/CVT).
COMMENT_COL_TESTS = {
    "driveaway", "decel_cstdown", "decel_cstdown_bev", "decel_opd", "uss_manual",
    "rtito", "ti_cstspd", "to_cstspd", "rrl", "tito_opt",
    "cstspd_manual", "cstspd_drive", "stationary_ess",
}

# ----------------------------------------------------------------------------- Pedal Geo procedure
PEDAL_GEO_STEPS = [
    ("Hardware Setup", [
        "Grab a Pedal Geo kit (with 410 included).",
        "Mount the arm bar onto the steering wheel; move the wheel to the highest, "
        "most telescoped position and get the string-pot wire in line with the pedal.",
        "Tape the load cell to the accel pedal (top flat first, then bottom with force).",
        "Hook the string pot to the load cell.",
        "Connect the 410 and string-pot connectors to the 593-D & 610.",
    ]),
    ("Software Setup", [
        "Import the pre-existing workspace for the Pedal Geo.",
        "Hardware page: ensure 410, 610, 593-D connected; 593-D has CAN monitoring.",
        "Set up sensor calibrations if needed (Load Cell = Auto Phys).",
        "Zero the sensors.",
        "Open the experiment, hit record, confirm a dot appears; add AccelPdlPosn_OBD "
        "or AccelPdlPosn signals.",
    ]),
    ("Gathering Data", [
        "Once recording is active, apply the pedal within 8 seconds.",
        "The pedal cannot be depressed prior to hitting record.",
        "Fully release the pedal / don't touch the load cell momentarily between presses.",
        "Record 2 separate files of 3 depressions each.",
    ]),
    ("Data Checking", [
        "Open the BMC script and run the Pedal Geo portion of the tool.",
        "Select the 2 data files and the correct signals.",
    ]),
]

# PV Max reference: output pedal % -> pedal voltage for the common 3.832 V (B) map
PV_MAX_3832B = {0: 0.0, 5: 0.205, 10: 0.396, 15: 0.596, 20: 0.777, 25: 0.958, 30: 1.158,
                35: 1.354, 40: 1.535, 45: 1.740, 50: 1.916, 60: 2.297, 70: 2.693,
                75: 2.874, 80: 3.055, 90: 3.451, 100: 3.832}
