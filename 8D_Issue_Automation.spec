# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

app_dir = Path(SPECPATH) / "app"

datas = []
binaries = []
hiddenimports = [
    p.stem
    for p in app_dir.glob("*.py")
    if p.name != "__init__.py"
]

# Bundle packages that carry runtime data/native files. tkinterdnd2 in
# particular needs its Tcl/Tk DnD payload inside the frozen application.
for package in ("pptx", "openpyxl", "tkinterdnd2", "PIL"):
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(package)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

a = Analysis(
    [str(app_dir / "main_enterprise_final.py")],
    pathex=[str(app_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=sorted(set(hiddenimports)),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="8D_Issue_Automation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
