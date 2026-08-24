"""
DRB Subjective Drivability Tool — standalone Python port of the 072926
Subjective Spreadsheets (AT / BEV / CVT).  No Excel dependency.

Entry model matches the Excel workflow exactly:
  * Type SHORTHAND DEFECT CODES in the event cells ("sj", "!b", "db-c").
    Blank cell on a completed run = clean pass.
  * Mark "x" (checkbox) in Complete per run row.
  * Enter the achieved TOP GEAR per run row (driveaway).
  * DRB Summary derives the colors:  green = complete & clean,
    yellow = event (marginal), red = "!" bad event, grey = not tested.
    Frequency = flagged runs / completed runs (<=25 green, 26-49 yellow, >=50 red).
    TOP GEAR on the summary = MAX over the pedal step's run rows.

Run:  streamlit run app.py
"""
import json
import re
from datetime import datetime

import pandas as pd
import streamlit as st

import config as C
import storage

try:
    import excel_export as _excel_export
    EXCEL_EXPORT_OK = True
    EXCEL_EXPORT_ERROR = ""
except Exception as exc:  # noqa: BLE001
    _excel_export = None
    EXCEL_EXPORT_OK = False
    EXCEL_EXPORT_ERROR = str(exc)

try:
    import table_image as _table_image
    TABLE_IMAGE_OK = True
    TABLE_IMAGE_ERROR = ""
except Exception as exc:  # noqa: BLE001
    _table_image = None
    TABLE_IMAGE_OK = False
    TABLE_IMAGE_ERROR = str(exc)

st.set_page_config(
    page_title=C.APP_TITLE,
    layout="wide",
    page_icon="🚗",
    menu_items={
        "Get help": None,
        "Report a bug": None,
        "About": f"{C.APP_TITLE} — rev {C.TEMPLATE_REV}. Runs fully offline; no internet required.",
    },
)

# ============================================================================ severity / colors
def cell_severity(code: str) -> str:
    """'' -> clean · '!' anywhere -> bad(red) · numeric -> explicit rating · else event(yellow)."""
    s = str(code).strip()
    if s in ("", "None", "nan"):
        return "clean"
    if re.fullmatch(r"10|[1-9](\.\d+)?", s):          # explicit 1-10 override
        r = float(s)
        return "clean" if r >= C.SUBJ_GREEN else ("event" if r >= C.SUBJ_YELLOW else "bad")
    if "!" in s:
        return "bad"
    return "event"


def sev_color(sev: str, tested: bool) -> str:
    if sev == "bad":
        return C.COLORS["red"]
    if sev == "event":
        return C.COLORS["yellow"]
    return C.COLORS["green"] if tested else C.COLORS["grey"]


def freq_color(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return C.COLORS["grey"]
    if v <= C.FREQ_GREEN:
        return C.COLORS["green"]
    if v <= C.FREQ_YELLOW:
        return C.COLORS["yellow"]
    return C.COLORS["red"]


# ============================================================================ state
def test_def(key: str, variant: str) -> dict:
    d = dict(C.TESTS[key])
    ov = C.VARIANT_OVERRIDES.get(variant, {})
    if key in ov:
        d.update(ov[key])
    return d


def run_label(run: dict) -> str:
    return f"{run['pedal']} · {run['brake']} ({run['ptype']})"


def is_accel_col(c) -> bool:
    cs = str(c)
    if cs in C.UPSHIFTS or cs.startswith("Launch"):
        return True
    if cs in C.COASTDOWNS:
        return False
    import re as _re
    ms = _re.search(r"(\d+)-(\d+)\s*mph", cs)
    return bool(ms) and int(ms.group(2)) > int(ms.group(1))


def col_group_header(cols, variant):
    """Return (accel_label, n_accel, decel_label, n_decel) for the grouping banner."""
    acc = [c for c in cols if is_accel_col(c)]
    dec = [c for c in cols if not is_accel_col(c)]
    if variant == "BEV":
        return ("Acceleration Speed Ranges", len(acc), "Deceleration Speed Ranges", len(dec))
    return ("Upshifts", len(acc), "Coastdowns", len(dec))


def run_rows(key: str) -> list[str]:
    if key in C.RUN_TESTS:
        return [run_label(r) for r in C.DRIVEAWAY_RUNS]
    return None


RUN_BY_LABEL = {run_label(r): r for r in C.DRIVEAWAY_RUNS}


def reset_editor_state(prefixes=("base_", "ed_", "tg_summary")):
    """Drop editor widget/base state so grids rebuild from stored data
    (call after loading/importing a session)."""
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in prefixes):
            del st.session_state[k]


def ensure_state():
    ss = st.session_state
    ss.setdefault("variant", None)
    ss.setdefault("vehicle", {f: "" for f in C.VEHICLE_FIELDS})
    ss["vehicle"].setdefault("date", datetime.now().strftime("%m%d%y"))
    ss.setdefault("data", {})
    ss.setdefault("session_name", "")


def ensure_test(key: str, variant: str):
    d = test_def(key, variant)
    rows = run_rows(key) or d["rows"]
    cols = d["cols"]
    t = st.session_state["data"].setdefault(key, {})
    if ("codes" not in t or list(t["codes"].keys()) != list(rows)
            or (t["codes"] and list(next(iter(t["codes"].values())).keys()) != list(cols))):
        t["codes"] = {r: {c: "" for c in cols} for r in rows}
        t["objective"] = {r: {c: "" for c in cols} for r in rows}
        t["complete"] = {r: False for r in rows}
        t["topgear"] = {r: "" for r in rows}
        t["comments"] = t.get("comments", [])
        for k in list(st.session_state.keys()):
            if k.startswith(f"base_{key}_") or k.startswith(f"ed_{key}_"):
                del st.session_state[k]
    t.setdefault("row_comments", {})
    if "objective" not in t or list(t["objective"].keys()) != list(rows):
        t["objective"] = {r: {c: t.get("objective", {}).get(r, {}).get(c, "") for c in cols}
                          for r in rows}
    if key in C.RUN_TESTS:
        t.setdefault("topgear_summary", {p: "" for p in C.DRIVEAWAY_PEDAL_ORDER})
        for p in C.DRIVEAWAY_PEDAL_ORDER:
            t["topgear_summary"].setdefault(p, "")
    return t, d, rows, cols


def vehicle_token() -> str:
    v = st.session_state["vehicle"]
    parts = [v.get("Model Year") or "0", v.get("Engine Disp.") or "0",
             v.get("Transmission") or "0", v.get("Model Code") or "0",
             v.get("Last 4 of VIN") or "0", v.get("date") or "0"]
    return "_".join(str(p).replace(" ", "") for p in parts)


def inca_label(prefix: str) -> str:
    return f"{prefix}{vehicle_token()}"


# ============================================================================ summary math
def aggregate_runs(t, cols):
    t_ref = t
    """Group driveaway run rows by pedal step -> per-shift subj severity, freq %, top gear.

    Excel behaviour: a pedal row's colors activate once its TOP GEAR is filled
    (or any of its runs is marked complete). Green = tested & clean.
    """
    groups = {}
    for label, done in t["complete"].items():
        pedal = label.split(" · ")[0]
        g = groups.setdefault(pedal, dict(runs=0, ev={c: 0 for c in cols},
                                          bad={c: False for c in cols}, gears=[]))
        if done:
            g["runs"] += 1
        for c in cols:
            sev = cell_severity(t["codes"][label][c])
            if sev != "clean":
                g["ev"][c] += 1
                if sev == "bad":
                    g["bad"][c] = True
        tg = str(t["topgear"].get(label, "")).strip()
        if tg.replace(".", "").isdigit():
            g["gears"].append(int(float(tg)))
    out = {}
    for pedal in C.DRIVEAWAY_PEDAL_ORDER:
        if pedal not in groups:
            continue
        g = groups[pedal]
        manual_tg = str(t.get("topgear_summary", {}).get(pedal, "")).strip()
        manual_val = int(float(manual_tg)) if manual_tg.replace(".", "").isdigit() else None
        topgear = manual_val if manual_val is not None else (max(g["gears"]) if g["gears"] else 0)
        tested = g["runs"] > 0 or manual_val is not None or bool(g["gears"])
        row = {}
        obj_map = {}
        for c in cols:
            # objective roll-up: worst verdict across the pedal's runs for this shift
            verdicts = [t_ref["objective"][lbl][c]
                        for lbl in t_ref["objective"]
                        if lbl.split(" · ")[0] == pedal and t_ref["objective"][lbl][c]]
            if "F" in verdicts:
                obj_map[c] = "F"
            elif "M" in verdicts:
                obj_map[c] = "M"
            elif "P" in verdicts:
                obj_map[c] = "P"
            else:
                obj_map[c] = ""
        for c in cols:
            sev = "bad" if g["bad"][c] else ("event" if g["ev"][c] else "clean")
            # gear masking: a shift "a-b" only exists if the achieved top gear >= b.
            # (only for single-digit gear labels — BEV speed-band columns are untouched)
            reachable = True
            cs = str(c).strip()
            m = re.fullmatch(r"(\d)-(\d)", cs)
            if m and topgear:
                a, b = int(m.group(1)), int(m.group(2))
                reachable = topgear >= (b if b > a else a)   # upshift a-b vs coastdown a-b
            else:
                ms = re.search(r"(\d+)-(\d+)\s*mph", cs)   # speed band (BEV/CVT)
                if ms and topgear:
                    a, b = int(ms.group(1)), int(ms.group(2))
                    # accel band a<b: entered it if top speed > a
                    # decel band a>b: entered it if top speed >= a
                    reachable = (topgear > a) if b > a else (topgear >= a)
            cell_tested = tested and reachable
            if sev == "clean" and not reachable:
                cell_tested = False                      # blank/grey beyond top gear
            n_runs = max(g["runs"], 1) if cell_tested else 0
            freq = round(100.0 * g["ev"][c] / n_runs, 0) if n_runs else None
            row[c] = dict(sev=sev, tested=cell_tested, freq=freq)
        out[pedal] = dict(cells=row, runs=g["runs"], topgear=topgear, obj=obj_map)
    return out


def simple_status(t, rows, cols):
    """Per-cell severity for non-run tests (row complete flag gates 'tested')."""
    out = {}
    for r in rows:
        tested = t["complete"].get(r, False)
        out[r] = {c: dict(sev=cell_severity(t["codes"][r][c]), tested=tested) for c in cols}
    return out


def styled_codes(t, rows, cols, key=None):
    df = pd.DataFrame(t["codes"]).T.reindex(index=rows, columns=cols)
    status = simple_status(t, rows, cols)
    is_run = key in C.RUN_TESTS if key else False

    def color(v, r, c):
        s = status[r][c]
        if s["sev"] == "clean" and not s["tested"] and is_run:
            base = C.ROW_BASE_COLORS[RUN_BY_LABEL[r]["base"]]   # Excel row fill
            return f"background-color: {base}"
        return f"background-color: {sev_color(s['sev'], s['tested'])}"

    sty = df.style
    for r in rows:
        for c in cols:
            sty = sty.map(lambda v, r=r, c=c: color(v, r, c),
                          subset=pd.IndexSlice[[r], [c]])
    return sty


def test_stats(t, rows, cols, key):
    n_runs = sum(1 for v in t["complete"].values() if v)
    events = bads = 0
    for r in rows:
        for c in cols:
            sev = cell_severity(t["codes"][r][c])
            if sev == "event":
                events += 1
            elif sev == "bad":
                bads += 1
    return dict(runs=n_runs, total_rows=len(rows), events=events, bads=bads)


def grid_cell_colors(t, rows, cols, key=None) -> dict[tuple[int, int], str]:
    """Background colors for JPG export of event grids."""
    colors: dict[tuple[int, int], str] = {}
    is_run = key in C.RUN_TESTS if key else False
    status = simple_status(t, rows, cols)
    for ri, r in enumerate(rows):
        for ci, c in enumerate(cols):
            s = status[r][c]
            if s["sev"] == "clean" and not s["tested"] and is_run:
                base = C.ROW_BASE_COLORS[RUN_BY_LABEL[r]["base"]]
                colors[(ri, ci)] = base
            else:
                colors[(ri, ci)] = sev_color(s["sev"], s["tested"])
    return colors


def show_table_with_jpg(df: pd.DataFrame, key: str, *, title: str = "",
                        cell_colors: dict[tuple[int, int], str] | None = None,
                        index: bool = True, render: bool = True, **dataframe_kwargs):
    """Render a table; JPG is created only when the user clicks download."""
    if render:
        st.dataframe(df, **dataframe_kwargs)
    if not TABLE_IMAGE_OK:
        st.caption("JPG download unavailable in this build.")
        return

    ready_key = f"jpg_ready_{key}"
    if st.button("Download table as JPG", key=f"mkjpg_{key}"):
        with st.spinner("Creating image..."):
            try:
                st.session_state[ready_key] = _table_image.dataframe_to_jpeg(
                    df, title=title or key, cell_colors=cell_colors, index=index)
            except Exception as exc:  # noqa: BLE001
                st.error(f"Could not create JPG: {exc}")
                return
    if ready_key in st.session_state:
        st.download_button(
            "Save JPG file",
            data=st.session_state[ready_key],
            file_name=f"{key}.jpg",
            mime="image/jpeg",
            key=f"dljpg_{key}",
        )


@st.cache_data(show_spinner="Building Excel file...")
def _build_excel_bytes(variant: str, vehicle: dict, data_json: str) -> bytes:
    if not EXCEL_EXPORT_OK:
        raise RuntimeError(EXCEL_EXPORT_ERROR or "Excel export module not available")
    data = json.loads(data_json)
    return _excel_export.export_session_to_excel(variant, vehicle, data)


def render_export_panel(variant: str, vehicle: dict, *, location: str = "summary"):
    """Prominent JSON / Excel export controls (Excel built only on request)."""
    st.subheader("Export")
    c1, c2 = st.columns(2)
    with c1:
        payload = {k: st.session_state[k] for k in ("variant", "vehicle", "data")}
        st.download_button(
            "Download session JSON",
            json.dumps(payload, indent=1),
            file_name=f"DRB_{variant}_{vehicle_token()}.json",
            key=f"json_{location}",
        )
    with c2:
        if not EXCEL_EXPORT_OK:
            st.error(f"Excel export unavailable: {EXCEL_EXPORT_ERROR}")
        else:
            ready_key = f"xlsx_ready_{location}"
            if st.button("Export as Excel (.xlsx)", key=f"mkxlsx_{location}", type="primary"):
                with st.spinner("Building Excel workbook (may take ~10 seconds)..."):
                    try:
                        st.session_state[ready_key] = _build_excel_bytes(
                            variant, vehicle, json.dumps(st.session_state["data"]))
                        st.session_state[f"xlsx_name_{location}"] = _excel_export.export_filename(
                            variant, vehicle)
                    except Exception as exc:  # noqa: BLE001
                        st.error(f"Excel export failed: {exc}")
            if ready_key in st.session_state:
                st.download_button(
                    "Save Excel file",
                    data=st.session_state[ready_key],
                    file_name=st.session_state.get(f"xlsx_name_{location}", "DRB_export.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dlxlsx_{location}",
                )
                st.caption("Filled copy of the original DRB spreadsheet template.")
    if TABLE_IMAGE_OK:
        st.caption("Use **Download table as JPG** under any table when you need a screenshot.")


# ============================================================================ pages
def page_home():
    st.title("🚗 " + C.APP_TITLE)
    st.caption(f"Template rev. {C.TEMPLATE_REV} — standalone, Excel-independent")

    st.subheader("1 · Select transmission / propulsion")
    labels = list(C.VARIANTS.keys())
    current = st.session_state["variant"]
    idx = labels.index(next((k for k, v in C.VARIANTS.items() if v == current), labels[0])) if current else 0
    choice = st.radio("Which variant are you working on?", labels, index=idx, horizontal=True)
    if C.VARIANTS[choice] != current:
        st.session_state["variant"] = C.VARIANTS[choice]
        reset_editor_state()
        st.rerun()
    variant = st.session_state["variant"]
    st.success(f"Active variant: **{choice}** — {len(C.VARIANT_TESTS[variant])} test sheets enabled")

    st.subheader("2 · Vehicle information")
    st.caption("Feeds every INCA/AVL recorder label and the summary header — fill this first.")
    v = st.session_state["vehicle"]
    cols = st.columns(4)
    for i, f in enumerate(C.VEHICLE_FIELDS):
        v[f] = cols[i % 4].text_input(f, v.get(f, ""))
    v["date"] = st.text_input("Test date (MMDDYY)", v.get("date", ""))
    st.code(f"Recorder token:  {vehicle_token()}", language=None)

    st.subheader("3 · Session")
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        name = st.text_input("Save as (blank = auto)", st.session_state.get("session_name", ""))
        if st.button("💾 Save session"):
            payload = {k: st.session_state[k] for k in ("variant", "vehicle", "data")}
            path = storage.save_session(payload, name or None)
            st.session_state["session_name"] = name
            st.success(f"Saved → {path}")
    with c2:
        pick = st.selectbox("Load existing session", ["—"] + storage.list_sessions())
        if st.button("📂 Load") and pick != "—":
            loaded = storage.load_session(pick)
            for k in ("variant", "vehicle", "data"):
                if k in loaded:
                    st.session_state[k] = loaded[k]
            reset_editor_state()
            st.rerun()
    with c3:
        payload = {k: st.session_state[k] for k in ("variant", "vehicle", "data")}
        st.download_button("⬇️ Export session (JSON)", json.dumps(payload, indent=1),
                           file_name=storage.session_filename(v, variant))
        up = st.file_uploader("Import session (JSON)", type="json", label_visibility="collapsed")
        if up is not None and st.button("Import uploaded file"):
            loaded = json.load(up)
            for k in ("variant", "vehicle", "data"):
                if k in loaded:
                    st.session_state[k] = loaded[k]
            reset_editor_state()
            st.rerun()


def page_test(key: str):
    variant = st.session_state["variant"]
    t, d, rows, cols = ensure_test(key, variant)
    st.title(d["name"])
    st.caption(d["note"])
    st.code(f"INCA/AVL filename:  {inca_label(d['prefix'])}", language=None)
    if key in C.RUN_TESTS:
        st.markdown(
            f"**WOT PV = {C.DRIVEAWAY_WOT_PV} V** &nbsp;·&nbsp; "
            "⬜ **white rows = Step-In pedal** &nbsp;·&nbsp; "
            "🟦 **blue rows = Ramp-In pedal** (brake to tip-in) &nbsp;·&nbsp; "
            "🟧 **orange rows = 99% prior to detent (optional, if the pedal has a detent)**")
    d_full = C.TESTS.get(key, {})
    if d_full.get("col_gears"):
        gear_hint = " · ".join(f"{c}: {'/'.join(sorted(set(g)))}"
                               for c, g in d_full["col_gears"].items())
        st.caption(f"Manual gears per speed → {gear_hint}")
    st.markdown(
        "**How to fill:** type shorthand codes in the event cells "
        "(`sj` slight jerk · `!b` bad bump · `db-c` double clunk — see Acronyms). "
        "Leave blank if the event was clean. Tick **Complete** for each run performed"
        + (f". Enter the achieved **{C.TOPGEAR_LABEL.get(variant, 'Top Gear').title()}** per run."
           if key in C.RUN_TESTS else "."))

    # editable grid: (PV | Pedal | Brake | Type) + Complete + (Top Gear) + event cells
    df = pd.DataFrame(t["codes"]).T.reindex(index=rows, columns=cols).fillna("")
    colcfg = {}
    if key in C.RUN_TESTS:
        pv_seen = set()
        pv_col, ped_col, brk_col, typ_col = [], [], [], []
        for r in rows:
            run = RUN_BY_LABEL[r]
            pv = C.DRIVEAWAY_PV.get(run["pedal"])
            pv_col.append(f"{pv:.3f}" if pv is not None and run["pedal"] not in pv_seen else "")
            pv_seen.add(run["pedal"])
            ped_col.append(run["pedal"])
            brk_col.append(run["brake"])
            typ_col.append("Ramp-In (blue)" if run["ptype"] == "Ramp" else
                           ("Step-In ·99% Opt." if run["base"] == "orange" else "Step-In"))
        df.insert(0, "PV [V]", pv_col)
        df.insert(1, "Pedal", ped_col)
        df.insert(2, "Braking Level", brk_col)
        df.insert(3, "Pedal Type", typ_col)
        df.insert(4, "Complete", [bool(t["complete"].get(r, False)) for r in rows])
        df.insert(5, "Top Gear", [str(t["topgear"].get(r, "")) for r in rows])
        df.index = pd.RangeIndex(len(rows))   # hide long labels; identity via columns
        colcfg["PV [V]"] = st.column_config.TextColumn("PV [V]", width="small", disabled=True)
        colcfg["Pedal"] = st.column_config.TextColumn("Pedal", width="small", disabled=True)
        colcfg["Braking Level"] = st.column_config.TextColumn("Braking Level", width="small", disabled=True)
        colcfg["Pedal Type"] = st.column_config.TextColumn("Pedal Type", width="small", disabled=True)
        colcfg["Top Gear"] = st.column_config.TextColumn(
            C.TOPGEAR_LABEL.get(variant, "Top Gear").title(), width="small")
    else:
        df.insert(0, "Complete", [bool(t["complete"].get(r, False)) for r in rows])
    colcfg["Complete"] = st.column_config.CheckboxColumn("Complete (x)", width="small")
    for c in cols:
        colcfg[c] = st.column_config.TextColumn(c, width="small")
    if key in C.COMMENT_COL_TESTS:
        df["Comments"] = [t.get("row_comments", {}).get(r, "") for r in rows]
        colcfg["Comments"] = st.column_config.TextColumn("Comments", width="medium")

    # column-group banner (Upshifts | Coastdowns  /  Accel | Decel ranges)
    acc_lbl, n_acc, dec_lbl, n_dec = col_group_header(cols, variant)
    if n_acc and n_dec:
        lead = 6 if key in C.RUN_TESTS else 1     # PV/Pedal/Brake/Type/Complete/TopGear vs Complete
        st.markdown(
            f"<div style='display:flex;gap:2px;margin-bottom:-8px;font-weight:700'>"
            f"<div style='flex:{lead}'></div>"
            f"<div style='flex:{n_acc};text-align:center;background:#dce6f2;"
            f"border:1px solid #9db8d6;padding:2px'>{acc_lbl}</div>"
            f"<div style='flex:{n_dec};text-align:center;background:#f2e6dc;"
            f"border:1px solid #d6b89d;padding:2px'>{dec_lbl}</div></div>",
            unsafe_allow_html=True)

    # STABLE base data: the same object must be passed to st.data_editor on every
    # rerun while the user is editing, otherwise Streamlit resets its pending edits
    # (values "disappear"). The editor's widget state (ed_key) lives only while this
    # page is rendered — when it's absent (first visit, or returning after visiting
    # another page), rebuild the base from memory so saved entries show up.
    base_id = f"base_{key}_{variant}"
    ed_key = f"ed_{key}_{variant}"
    if ed_key not in st.session_state or base_id not in st.session_state:
        st.session_state[base_id] = df
    base_df = st.session_state[base_id]

    editor_data = base_df
    if key in C.RUN_TESTS:
        _bases = [C.ROW_BASE_COLORS[RUN_BY_LABEL[r]["base"]] for r in rows]

        def _row_fill(s):
            color = _bases[s.name]
            css = f"background-color: {color}" if color != "#ffffff" else ""
            return [css] * len(s)

        editor_data = base_df.style.apply(_row_fill, axis=1)
    edited = st.data_editor(editor_data, column_config=colcfg, width="stretch",
                            key=f"ed_{key}_{variant}", height=min(38 * len(rows) + 40, 900),
                            hide_index=(key in C.RUN_TESTS))
    for i, r in enumerate(rows):
        rec = edited.iloc[i]
        t["complete"][r] = bool(rec["Complete"])
        if key in C.RUN_TESTS:
            t["topgear"][r] = str(rec["Top Gear"]).strip()
        for c in cols:
            t["codes"][r][c] = str(rec[c]).strip()
        if key in C.COMMENT_COL_TESTS:
            t["row_comments"][r] = str(rec["Comments"]).strip()

    codes_preview = pd.DataFrame(t["codes"]).T.reindex(index=rows, columns=cols).fillna("")
    with st.expander("Download test sheet as JPG"):
        show_table_with_jpg(
            codes_preview,
            f"{key}_{variant}_grid",
            title=d["name"],
            cell_colors=grid_cell_colors(t, rows, cols, key),
            index=True,
            render=False,
        )

    with st.expander("Live color status (derived, same rules as DRB Summary)"):
        st.dataframe(styled_codes(t, rows, cols, key), width="stretch")

    with st.expander("Objective verdicts (measured pass / marginal / fail per shift)"):
        st.caption("Optional: log the objective (measured) result per cell — "
                   "P = pass 🟩 · M = marginal 🟨 · F = fail 🟥. Feeds the "
                   "Objective column on the DRB Summary.")
        obj_base_id = f"objbase_{key}_{variant}"
        obj_ed_key = f"objed_{key}_{variant}"
        obj_df = pd.DataFrame(t["objective"]).T.reindex(index=rows, columns=cols).fillna("")
        obj_df.index = pd.RangeIndex(len(rows)) if key in C.RUN_TESTS else rows
        if obj_ed_key not in st.session_state or obj_base_id not in st.session_state:
            st.session_state[obj_base_id] = obj_df
        obj_cfg = {c: st.column_config.SelectboxColumn(c, options=["", "P", "M", "F"],
                                                       width="small") for c in cols}
        obj_edited = st.data_editor(st.session_state[obj_base_id], column_config=obj_cfg,
                                    width="stretch", key=obj_ed_key,
                                    height=min(38 * len(rows) + 40, 480),
                                    hide_index=(key in C.RUN_TESTS))
        for i, r in enumerate(rows):
            for c in cols:
                t["objective"][r][c] = str(obj_edited.iloc[i][c]).strip()
        obj_colors = {}
        obj_map = {"P": C.COLORS["green"], "M": C.COLORS["yellow"], "F": C.COLORS["red"]}
        for ri, r in enumerate(rows):
            for ci, c in enumerate(cols):
                v = t["objective"][r][c]
                if v:
                    obj_colors[(ri, ci)] = obj_map.get(v, C.COLORS["grey"])
        show_table_with_jpg(
            pd.DataFrame(t["objective"]).T.reindex(index=rows, columns=cols).fillna(""),
            f"{key}_{variant}_objective",
            title=f"{d['name']} — objective verdicts",
            cell_colors=obj_colors,
            index=(key not in C.RUN_TESTS),
            render=False,
        )

    st.markdown("**Sheet notes**")
    new = st.text_input("Add note", key=f"cmt_{key}")
    if st.button("Add note", key=f"cmtbtn_{key}") and new.strip():
        t["comments"].append({"time": datetime.now().strftime("%H:%M:%S"), "text": new.strip()})
        st.rerun()
    for cm in reversed(t["comments"]):
        st.markdown(f"- `{cm['time']}` {cm['text']}")


def driveaway_summary_block(t, cols):
    # --- editable TOP GEAR per pedal step (Excel: filling it activates the row colors)
    c_grid, c_tg = st.columns([6, 1])
    with c_tg:
        st.markdown(f"**{C.TOPGEAR_LABEL.get(st.session_state['variant'], 'TOP GEAR')}**")
        if "tg_summary" not in st.session_state or "tg_summary_base" not in st.session_state:
            st.session_state["tg_summary_base"] = pd.DataFrame(
                {"Top Gear": [str(t["topgear_summary"].get(p, ""))
                              for p in C.DRIVEAWAY_PEDAL_ORDER]},
                index=C.DRIVEAWAY_PEDAL_ORDER)
        tg_df = st.session_state["tg_summary_base"]
        tg_ed = st.data_editor(tg_df, width="stretch", key="tg_summary",
                               column_config={"Top Gear": st.column_config.TextColumn(
                                   "Gear", width="small")},
                               height=38 * len(C.DRIVEAWAY_PEDAL_ORDER) + 40)
        for p in C.DRIVEAWAY_PEDAL_ORDER:
            t["topgear_summary"][p] = str(tg_ed.loc[p, "Top Gear"]).strip()
        show_table_with_jpg(tg_ed, "driveaway_topgear_summary",
                            title="Block Pedal Top Gear summary", index=True, render=False)

    agg = aggregate_runs(t, cols)
    pedals = [p for p in C.DRIVEAWAY_PEDAL_ORDER if p in agg]

    def obj_color(v):
        return {"P": C.COLORS["green"], "M": C.COLORS["yellow"],
                "F": C.COLORS["red"]}.get(v, C.COLORS["grey"])

    def color_grid(sub_cols, title, height_rows=None):
        header = []
        for c in sub_cols:
            header += [f"{c} S", f"{c} F%", f"{c} Obj"]
        data, colors = [], []
        for p in pedals:
            row, crow = [], []
            for c in sub_cols:
                cell = agg[p]["cells"][c]
                ov = agg[p].get("obj", {}).get(c, "")
                row += ["", "", ov]
                crow += [sev_color(cell["sev"], cell["tested"]),
                         freq_color(cell["freq"]), obj_color(ov)]
            data.append(row)
            colors.append(crow)
        df = pd.DataFrame(data, index=pedals, columns=header)
        cdf = pd.DataFrame(colors, index=pedals, columns=header)
        st.markdown(title)
        st.dataframe(df.style.apply(lambda s: [f"background-color: {cdf.loc[s.name, c]}"
                                               for c in df.columns], axis=1),
                     width="stretch", height=38 * len(pedals) + 40)
        flat_colors = {
            (ri, ci): cdf.iloc[ri, ci]
            for ri in range(len(pedals))
            for ci in range(len(header))
        }
        slug = re.sub(r"[^a-zA-Z0-9]+", "_", title)[:40].strip("_").lower()
        show_table_with_jpg(
            df,
            f"driveaway_{slug}",
            title=title.strip("*"),
            cell_colors=flat_colors,
            index=True,
            render=False,
        )

    def _is_accel(c):
        cs = str(c)
        if cs in C.UPSHIFTS or cs.startswith("Launch"):
            return True
        if cs in C.COASTDOWNS:
            return False
        ms = re.search(r"(\d+)-(\d+)\s*mph", cs)
        return bool(ms) and int(ms.group(2)) > int(ms.group(1))

    up_cols = [c for c in cols if _is_accel(c)]
    cd_cols = [c for c in cols if not _is_accel(c)]
    with c_grid:
        color_grid(up_cols, "**Upshifts / Acceleration — Subjective | Frequency** "
                            "(colors only — 🟩 clean · 🟨 event · 🟥 bad · ⬜ not tested)")
        if cd_cols:
            color_grid(cd_cols, "**Coastdowns / Deceleration [CD Sum] — Subjective | Frequency**")


def page_summary():
    variant = st.session_state["variant"]
    st.title("DRB Summary — " + [k for k, v in C.VARIANTS.items() if v == variant][0])
    v = st.session_state["vehicle"]
    st.caption(" · ".join(f"{f}: {v.get(f) or '—'}" for f in
                          ("Model Year", "Vehicle Line", "Engine Disp.", "Transmission",
                           "Software Level", "Last 4 of VIN")))

    render_export_panel(variant, v, location="summary_top")

    st.markdown("Legend: 🟩 complete & clean · 🟨 event (marginal) · 🟥 `!` bad event · "
                "⬜ not tested — Frequency: ≤25% 🟩 · 26–49% 🟨 · ≥50% 🟥")

    # ---- roll-up table
    rows_out = []
    for key in C.VARIANT_TESTS[variant]:
        t, d, rows, cols = ensure_test(key, variant)
        s = test_stats(t, rows, cols, key)
        status = ("🟥 Bad events" if s["bads"] else
                  "🟨 Events" if s["events"] else
                  "🟩 Clean" if s["runs"] else "⬜ Waiting")
        rows_out.append({"Test": d["name"],
                         "Runs complete": f"{s['runs']}/{s['total_rows']}",
                         "Events": s["events"], "Bad (!)": s["bads"],
                         "Notes": len(t["comments"]), "Status": status})
    rollup_df = pd.DataFrame(rows_out)
    show_table_with_jpg(
        rollup_df,
        f"summary_rollup_{variant}",
        title="DRB Summary — test roll-up",
        index=False,
        width="stretch",
        hide_index=True,
    )

    # ---- driveaway block (Excel Block Pedal [Sweeps] Summary)
    if "driveaway" in C.VARIANT_TESTS[variant]:
        st.subheader("Block Pedal [Sweeps] Summary — Subjective | Frequency | Top Gear")
        t, d, rows, cols = ensure_test("driveaway", variant)
        driveaway_summary_block(t, cols)

    # ---- per-test color charts
    st.subheader("Color charts")
    for key in C.VARIANT_TESTS[variant]:
        t, d, rows, cols = ensure_test(key, variant)
        with st.expander(d["name"]):
            preview = pd.DataFrame(t["codes"]).T.reindex(index=rows, columns=cols).fillna("")
            show_table_with_jpg(
                preview,
                f"chart_{key}_{variant}",
                title=d["name"],
                cell_colors=grid_cell_colors(t, rows, cols, key),
                index=True,
                width="stretch",
            )

    variant = st.session_state["variant"]
    st.title("INCA / AVL Filename Generator")
    st.caption("Copy the label and paste it into INCA/AVL before recording.")
    rows = [{"Test": test_def(k, variant)["name"], "Label": inca_label(test_def(k, variant)["prefix"])}
            for k in C.VARIANT_TESTS[variant]]
    rows += [{"Test": f"(extra) {n}", "Label": inca_label(p)} for n, p in C.EXTRA_PREFIXES.items()]
    fn_df = pd.DataFrame(rows)
    show_table_with_jpg(fn_df, f"filenames_{variant}", title="INCA / AVL Filenames", index=False,
                        width="stretch")


def page_pv_max():
    st.title("PV Max — Pedal Voltage Reference")
    st.caption("Output pedal % → pedal voltage (3.832 V 'B' map). WOT = 3.832 V.")
    rows = [{"Pedal %": k, "PV [V]": v} for k, v in C.PV_MAX_3832B.items()]
    show_table_with_jpg(pd.DataFrame(rows), "pv_max", title="PV Max — Pedal Voltage Reference",
                        index=False, width="stretch")
    st.info("Full PV Max table (all voltage variants: 3.001 / 3.182 / 3.832 / 4.008 / "
            "4.477) lives in the workbook; this is the map the Driveaway PV column uses.")


def page_pedal_geo():
    st.title("Pedal Geo — Measurement Procedure")
    st.caption("Load-cell + string-pot pedal force/travel capture (410 / 610 / 593-D).")
    for section, steps in C.PEDAL_GEO_STEPS:
        st.subheader(section)
        for i, s in enumerate(steps, 1):
            st.markdown(f"{i}. {s}")


def page_acronyms():
    st.title("Comment Shorthand — Acronyms")
    st.markdown("**General classifications**")
    acr_df = pd.DataFrame(C.ACRONYMS, columns=["Descriptor", "Shorthand", "Definition"])
    show_table_with_jpg(acr_df, "acronyms_general", title="Acronyms — General", index=False)
    st.markdown("**Modifiers**")
    mod_df = pd.DataFrame(C.MODIFIERS, columns=["Variation", "Shorthand", "Definition"])
    show_table_with_jpg(mod_df, "acronyms_modifiers", title="Acronyms — Modifiers", index=False)
    st.info("Cell examples: `sj` slight jerk (yellow) · `!b` bad bump (red) · "
            "`db-c` double clunk (yellow) · `l-sh` late shudder (yellow) · "
            "blank on a completed run = clean (green). "
            "A number 1–10 in a cell is treated as an explicit subjective rating.")


# ============================================================================ main
def _sidebar_name(key, variant):
    return test_def(key, variant)["name"]


ensure_state()
with st.sidebar:
    st.markdown(f"### {C.APP_TITLE}")
    variant = st.session_state["variant"]
    if variant:
        vlabel = [k for k, v in C.VARIANTS.items() if v == variant][0]
        st.markdown(f"**Variant:** {vlabel}")
        pages = (["🏠 Home / Vehicle"]
                 + [f"🧪 {_sidebar_name(k, variant)}" for k in C.VARIANT_TESTS[variant]]
                 + ["📊 DRB Summary", "🏷️ Filenames (INCA/AVL)",
                    "⚡ PV Max", "📐 Pedal Geo", "🔤 Acronyms"])
        sel = st.radio("Navigate", pages, label_visibility="collapsed")
    else:
        sel = "🏠 Home / Vehicle"
        st.info("Select a transmission on the Home page to unlock the test sheets.")
    st.divider()
    if st.button("💾 Quick-save session", width="stretch") and variant:
        payload = {k: st.session_state[k] for k in ("variant", "vehicle", "data")}
        st.success(storage.save_session(payload))
    if variant:
        st.caption("Excel & JPG exports are on the **DRB Summary** page and under each table.")

if sel == "🏠 Home / Vehicle":
    page_home()
elif sel == "📊 DRB Summary":
    page_summary()
elif sel == "🏷️ Filenames (INCA/AVL)":
    page_filenames()
elif sel == "⚡ PV Max":
    page_pv_max()
elif sel == "📐 Pedal Geo":
    page_pedal_geo()
elif sel == "🔤 Acronyms":
    page_acronyms()
else:
    name = sel.removeprefix("🧪 ")
    variant = st.session_state["variant"]
    key = next(k for k in C.VARIANT_TESTS[variant] if test_def(k, variant)["name"] == name)
    page_test(key)
