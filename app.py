import streamlit as st

from app.cell_store import CellStore, load_bundle
from app.config import DEFAULT_SHEET, ordered_sheets, VERSIONS
from app.formula_engine import recalculate
from app.grid_widget import render_editable_grid


@st.cache_resource(show_spinner=False)
def load_bundle_cached(version_key: str) -> dict:
    return load_bundle(version_key)


def get_store(version_key: str) -> CellStore:
    if (
        st.session_state.get("store") is None
        or st.session_state.get("store_version") != version_key
    ):
        store = CellStore(load_bundle_cached(version_key))
        recalculate(store)
        st.session_state.store = store
        st.session_state.store_version = version_key
        st.session_state.last_edit_sig = None
    return st.session_state.store


def init_session_state() -> None:
    defaults = {
        "version_key": VERSIONS[list(VERSIONS.keys())[0]],
        "active_sheet": DEFAULT_SHEET,
        "last_edit_sig": None,
        "store": None,
        "store_version": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def apply_edits(store: CellStore, sheet: str, edits: dict[str, str]) -> bool:
    if not edits:
        return False
    sig = str(sorted(edits.items()))
    if st.session_state.last_edit_sig == sig:
        return False
    for coord, value in edits.items():
        if store.is_editable(sheet, coord):
            store.set_user_value(sheet, coord, value)
    recalculate(store)
    st.session_state.last_edit_sig = sig
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
          div[data-testid="stSidebar"] { background-color: #f7f7f7; }
          .banner {
            background: linear-gradient(90deg, #217346, #185c37);
            color: white; padding: 10px 14px; border-radius: 8px; margin-bottom: 8px;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Subjective Spreadsheet Tool")
    st.caption("Pure Python — double-click a white cell to edit (like Excel).")

    labels = list(VERSIONS.keys())
    current_label = next(
        k for k, v in VERSIONS.items() if v == st.session_state.version_key
    )
    label = st.selectbox(
        "Select spreadsheet version",
        labels,
        index=labels.index(current_label),
        key="version_label",
    )
    version_key = VERSIONS[label]
    if version_key != st.session_state.version_key:
        st.session_state.version_key = version_key
        st.session_state.store = None

    store = get_store(version_key)
    sheets = ordered_sheets(store.list_sheets())

    if st.session_state.active_sheet not in sheets:
        st.session_state.active_sheet = (
            DEFAULT_SHEET if DEFAULT_SHEET in sheets else sheets[0]
        )

    active_sheet = st.session_state.active_sheet

    st.markdown(
        f"<div class='banner'>"
        f"<b>Version:</b> {label} &nbsp;|&nbsp; "
        f"<b>Sheet:</b> {active_sheet} &nbsp;|&nbsp; "
        f"<b>Engine:</b> Python (instant)"
        f"</div>",
        unsafe_allow_html=True,
    )

    st.caption("White cells = editable (double-click). Grey cells = calculated.")

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
        if st.button("Recalculate", type="primary", use_container_width=True):
            recalculate(store)
            st.session_state.last_edit_sig = None
            st.rerun()

        st.download_button(
            "Download .xlsx",
            data=store.export_xlsx_bytes(),
            file_name=f"{store.book_name.replace('.xlsm', '')}_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    meta = store.sheet_meta(active_sheet)
    if not meta.get("bounds"):
        st.info("This sheet is empty in the template.")
        return

    edits = render_editable_grid(store, active_sheet)
    if edits:
        if apply_edits(store, active_sheet, edits):
            st.rerun()


if __name__ == "__main__":
    main()
