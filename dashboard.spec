# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_submodules,
    copy_metadata
)


# ==============================
# STREAMLIT
# ==============================

streamlit_datas = collect_data_files("streamlit")
streamlit_hiddenimports = collect_submodules("streamlit")
streamlit_metadata = copy_metadata("streamlit")


# ==============================
# PANDAS
# ==============================

pandas_datas = collect_data_files("pandas")
pandas_hiddenimports = collect_submodules("pandas")


# ==============================
# PLOTLY
# ==============================

plotly_datas = collect_data_files("plotly")
plotly_hiddenimports = collect_submodules("plotly")


# ==============================
# OPENPYXL
# ==============================

openpyxl_datas = collect_data_files("openpyxl")
openpyxl_hiddenimports = collect_submodules("openpyxl")


# ==============================
# ANALISIS
# ==============================

a = Analysis(
    ["run_app.py"],

    pathex=[],

    binaries=[],

    datas=(
        streamlit_datas
        + streamlit_metadata
        + pandas_datas
        + plotly_datas
        + openpyxl_datas
        + [
            ("app.py", "."),
	    ("image.jpeg",".")
        ]
    ),

    hiddenimports=(
        streamlit_hiddenimports
        + pandas_hiddenimports
        + plotly_hiddenimports
        + openpyxl_hiddenimports
    ),

    hookspath=[],

    hooksconfig={},

    runtime_hooks=[],

    excludes=[],

    noarchive=False,
)


# ==============================
# PYZ
# ==============================

pyz = PYZ(
    a.pure
)


# ==============================
# EJECUTABLE
# ==============================

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],

    name="Dashboard Autoinspecciones",

    debug=False,

    bootloader_ignore_signals=False,

    strip=False,

    upx=True,

    console=True,

    disable_windowed_traceback=False,

    argv_emulation=False,

    target_arch=None,

    codesign_identity=None,

    entitlements_file=None,
)