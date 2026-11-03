# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build definition for WenShu Converter."""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [("assets", "assets")]
datas += collect_data_files("docx")
datas += collect_data_files("pdf2docx")
datas += collect_data_files("fontTools")

hiddenimports = [
    "pymupdf",
    "fitz",
    "cv2",
    "docx",
    "lxml",
    "fontTools.ttLib",
    "comtypes",
    "comtypes.client",
    "pdf_flow",
]
hiddenimports += collect_submodules("pdf2docx")

excluded = [
    "onnxruntime",
    "pymupdf_layout",
    "pymupdf4llm",
    "networkx",
    "pandas",
    "matplotlib",
    "scipy",
    "IPython",
    "pytest",
    "tkinter",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtMultimedia",
    "PySide6.QtCharts",
    "PySide6.QtDataVisualization",
    "PySide6.QtPdf",
    "PySide6.QtSql",
    "PySide6.QtTest",
    "PySide6.QtOpenGL",
    "PySide6.Qt3DCore",
]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="WenShuConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon="assets/app.ico",
    version="version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="WenShuConverter",
)
