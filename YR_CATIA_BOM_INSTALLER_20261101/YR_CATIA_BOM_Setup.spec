# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


installer_root = Path(SPECPATH).resolve()
installer_entry = installer_root / "installer_gui.py"
app_bundle = installer_root / "dist" / "YR_CATIA_BOM_Manager"
icon_ico = installer_root / "assets" / "YR.ico"
icon_png = installer_root / "assets" / "YR_icon.png"
version_file = installer_root / "version_info_setup.txt"

if not app_bundle.is_dir():
    raise FileNotFoundError(
        f"Application bundle not found: {app_bundle}. Build YR_CATIA_BOM_Manager.spec first."
    )

hidden_imports = [
    "pythoncom",
    "pywintypes",
    "win32com.client",
    "win32com.client.dynamic",
]

a = Analysis(
    [str(installer_entry)],
    pathex=[str(installer_root)],
    binaries=[],
    datas=[
        (str(icon_ico), "assets"),
        (str(icon_png), "assets"),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
a.datas += Tree(str(app_bundle), prefix="payload/YR_CATIA_BOM_Manager")
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="YR_CATIA_BOM_Setup",
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
    icon=str(icon_ico),
    version=str(version_file),
)
