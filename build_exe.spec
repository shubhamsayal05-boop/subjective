# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — run on Windows: pyinstaller build_exe.spec

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None
project_root = Path(SPECPATH)

# Collect Streamlit + AgGrid assets (static files, metadata)
streamlit_datas, streamlit_binaries, streamlit_hiddenimports = collect_all("streamlit")
aggrid_datas, aggrid_binaries, aggrid_hiddenimports = collect_all("st_aggrid")

hiddenimports = (
    streamlit_hiddenimports
    + aggrid_hiddenimports
    + collect_submodules("streamlit")
    + [
        "streamlit.web.cli",
        "streamlit.runtime.scriptrunner.magic_funcs",
        "pandas",
        "openpyxl",
        "altair",
        "click",
        "tornado",
        "watchdog",
        "pyarrow",
        "PIL",
        "packaging",
    ]
)

datas = [
    (str(project_root / "app.py"), "."),
    (str(project_root / "app"), "app"),
] + streamlit_datas + aggrid_datas

a = Analysis(
    [str(project_root / "run_app.py")],
    pathex=[str(project_root)],
    binaries=streamlit_binaries + aggrid_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SubjectiveSpreadsheetTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="SubjectiveSpreadsheetTool",
)
