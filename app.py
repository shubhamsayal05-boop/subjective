import streamlit as st
import formulas

from app.config import DEFAULT_SHEET, get_template_path, ordered_sheets, VERSIONS
from app.grid_builder import build_grid_data
from app.workbook_manager import WorkbookManager
from components.excel_grid import excel_grid


@st.cache_resource(show_spinner=False)
def load_formula_model(template_path: str) -> formulas.ExcelModel:
    return formulas.ExcelModel().loads(template_path).finish()


@st.cache_resource(show_spinner="Loading spreadsheet...")
def get_manager(version_label: str) -> WorkbookManager:
    return WorkbookManager(get_template_path(version_label))


def ensure_formula_engine(manager: WorkbookManager, version_label: str) -> None:
    if manager.has_formula_engine():
        return
    path = str(get_template_path(version_label))
    with st.spinner("Starting formula engine (first time ~25s, then cached)..."):
        model = load_formula_model(path)
        manager.attach_formula_model(model)


def init_session_state() -> None:
    defaults = {
        "version": list(VERSIONS.keys())[0],
        "loaded_version": None,
        "active_sheet": DEFAULT_SHEET,
        "load_error": None,
        "live_calc": False,
        "last_edit_sig": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def ensure_version(version_label: str) -> WorkbookManager | None:
    try:
        if st.session_state.loaded_version != version_label:
            st.session_state.loaded_version = version_label
            st.session_state.version = version_label
            st.session_state.live_calc = False
            st.session_state.last_edit_sig = None
            sheets = ordered_sheets(get_manager(version_label).list_sheets())
            if st.session_state.active_sheet not in sheets:
                st.session_state.active_sheet = (
                    DEFAULT_SHEET if DEFAULT_SHEET in sheets else sheets[0]
                )
        return get_manager(version_label)
    except Exception as exc:
        st.session_state.load_error = str(exc)
        return None


def apply_grid_edits(manager: WorkbookManager, sheet: str, edits: dict) -> bool:
    if not edits:
        return False
    edit_sig = str(sorted(edits.items()))
    if st.session_state.last_edit_sig == edit_sig:
        return False
    for coord, value in edits.items():
        if manager.is_editable(sheet, coord):
            manager.set_cell_value(sheet, coord, value)
    st.session_state.last_edit_sig = edit_sig
    if manager.has_formula_engine():
        manager.recalculate()
        st.session_state.live_calc = True
    return True


def main() -> None:
    st.set_page_config(
        page_title="Subjective Spreadsheet Tool",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_session_state()

    st.markdown(
        """
        <style>
          .block-container { padding-top: 0.75rem; max-width: 100%; }
          div[data-testid="stSidebar"] { background: #f7f7f7; }
          .banner {
            background: linear-gradient(90deg, #217346, #185c37);
            color: white; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px;
            font-size: 14px;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Subjective Spreadsheet Tool")
    st.caption(
        "Standalone web tool — no Microsoft Excel required. "
        "Double-click any white cell to edit, like Excel."
    )

    version_labels = list(VERSIONS.keys())
    selected_version = st.selectbox(
        "Select spreadsheet version",
        version_labels,
        index=version_labels.index(st.session_state.version),
        key="version_selector",
    )

    manager = ensure_version(selected_version)
    if manager is None:
        st.error(f"Failed to load: {st.session_state.load_error}")
        return

    sheets = ordered_sheets(manager.list_sheets())
    if st.session_state.active_sheet not in sheets:
        st.session_state.active_sheet = (
            DEFAULT_SHEET if DEFAULT_SHEET in sheets else sheets[0]
        )

    mode = "Live formulas" if st.session_state.live_calc else "Fast preview"
    st.markdown(
        f"<div class='banner'>"
        f"<b>Version:</b> {selected_version} &nbsp;|&nbsp; "
        f"<b>Sheet:</b> {st.session_state.active_sheet} &nbsp;|&nbsp; "
        f"<b>Mode:</b> {mode}"
        f"</div>",
        unsafe_allow_html=True,
    )

    if not st.session_state.live_calc:
        st.caption(
            "Grey cells are calculated automatically. "
            "After editing, click **Recalculate** in the sidebar to update formula cells."
        )

    with st.sidebar:
        st.header("Sheets")
        for group_name, group_sheets in [
            (
                "Report",
                [
                    "Test Summary -->",
                    "DRB - Summary",
                    "DRB - Color Chart",
                    "Optional Test - Summary",
                    "Engine - Summary",
                ],
            ),
            (
                "DRB Tests",
                [
                    "DRB Testing -->",
                    "Driveaway-sweeps",
                    "Driveaway ESS",
                    "Decel Cstdowns",
                    "Decel OPD",
                    "USS-Manual (Opt.)",
                    "RTITO_AD",
                    "TI_CstSpd ",
                    "TO_CstSpd",
                    "RRL (with HS)",
                    "Garage Shifts",
                    "Kickdowns",
                ],
            ),
            (
                "Engine Tests",
                [
                    "ENG Testing -->",
                    "TITO(Opt)",
                    "CstSpd",
                    "Acceleration(Opt)",
                    "Stationary Test",
                    "Engine - Color Chart",
                ],
            ),
            ("Reference", ["Pedal Geo", "Test Details -->", "PV Max", "Acronyms"]),
        ]:
            visible = [s for s in group_sheets if s in sheets]
            if not visible:
                continue
            st.subheader(group_name)
            for name in visible:
                if st.button(name, key=f"nav_{name}", use_container_width=True):
                    st.session_state.active_sheet = name
                    st.session_state.last_edit_sig = None
                    st.rerun()

        st.divider()
        if st.button("Recalculate all sheets", type="primary", use_container_width=True):
            ensure_formula_engine(manager, selected_version)
            with st.spinner("Recalculating..."):
                manager.recalculate()
            st.session_state.live_calc = True
            st.session_state.last_edit_sig = None
            st.rerun()

        st.download_button(
            "Download .xlsm",
            data=manager.export_bytes(),
            file_name=manager.book_name.replace(".xlsm", "_export.xlsm"),
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
        )

        st.divider()
        st.caption("Acronyms: nd, b, c, st, j, sh — prefix ! for bad events")

    active_sheet = st.session_state.active_sheet
    grid_data = build_grid_data(manager, active_sheet, height=700)

    grid_result = excel_grid(
        grid_data,
        height=720,
        key=f"grid_{selected_version}_{active_sheet}",
    )

    if grid_result and grid_result.get("edits"):
        if apply_grid_edits(manager, active_sheet, grid_result["edits"]):
            st.rerun()


if __name__ == "__main__":
    main()
