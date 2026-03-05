# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

def _find_project_root() -> Path:
    """Find project root reliably even if spec is invoked from different working dirs."""
    starts = [Path.cwd(), Path(SPECPATH).resolve()]
    for s in list(starts):
        starts.extend(list(s.parents)[:6])
    seen = set()
    for base in starts:
        key = str(base)
        if key in seen:
            continue
        seen.add(key)
        if (base / "app" / "__main__.py").exists():
            return base
        if (base / "maintenance_reporter" / "app" / "__main__.py").exists():
            return base / "maintenance_reporter"
    return Path(SPECPATH).resolve().parents[2]

project_root = _find_project_root()
entry = project_root / "app" / "__main__.py"

block_cipher = None

a = Analysis(
    [str(entry)],
    pathex=[str(project_root)],
    binaries=[],
    datas=[],
    hiddenimports=[
        # tkinter is usually detected, but keep it safe for some environments
        "tkinter",
        "tkinter.filedialog",
        "tkinter.messagebox",
    ],
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
    name="MaintenanceReporter",
    icon=str(project_root/"assets"/"app.ico"),
    version=str(project_root/"build"/"pyinstaller"/"version_info.txt"),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="MaintenanceReporter",
)