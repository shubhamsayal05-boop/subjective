# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for DRB Subjective Drivability Tool (Windows .exe)."""

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None
root = Path(SPECPATH)

app_datas = [
    (str(root / "app.py"), "."),
    (str(root / "config.py"), "."),
    (str(root / "storage.py"), "."),
    (str(root / "excel_export.py"), "."),
    (str(root / "table_image.py"), "."),
    (str(root / ".streamlit"), ".streamlit"),
    (str(root / "Subjective_SprdSheet_072926.xlsm"), "."),
    (str(root / "BEV_Subjective_SprdSheet_072926.xlsm"), "."),
    (str(root / "CVT Subjective_SprdSheet_072926.xlsm"), "."),
]

datas = list(app_datas)
binaries = []
hiddenimports = [
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
    "pandas",
    "pyarrow",
    "numpy",
    "altair",
    "watchdog",
    "tornado",
    "click",
    "packaging",
    "packaging.version",
    "packaging.specifiers",
    "packaging.requirements",
    "importlib_metadata",
    "tzdata",
    "openpyxl",
    "matplotlib",
    "PIL",
    "excel_export",
    "table_image",
]

for pkg in ("streamlit", "altair", "pyarrow", "pandas"):
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

hiddenimports += collect_submodules("streamlit")

a = Analysis(
    [str(root / "launcher.py")],
    pathex=[str(root)],
    binaries=binaries,
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="DRB_Subjective_Tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
