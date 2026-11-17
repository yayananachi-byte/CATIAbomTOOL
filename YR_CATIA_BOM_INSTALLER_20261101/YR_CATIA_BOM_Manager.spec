# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


installer_root = Path(SPECPATH).resolve()
project_root = installer_root.parent
gui_entry = project_root / "CATIA_BOM_GUI" / "catia_bom_gui.py"
converter = project_root / "模板转化" / "bom_to_control_table.py"
icon_ico = installer_root / "assets" / "YR.ico"
version_file = installer_root / "version_info_app.txt"

hidden_imports = [
    "catia_bom_export_v2",
    "gwm_properties",
    "geely_data",
    # 模板转换器作为数据文件动态加载，PyInstaller 不会自动分析其新增依赖。
    "PIL.ImageDraw",
    "pythoncom",
    "pywintypes",
    "win32com.client",
    "win32com.client.dynamic",
    "win32com.client.gencache",
]

a = Analysis(
    [str(gui_entry)],
    # GUI 内核优先，避免打包到仓库根目录的兼容入口。
    pathex=[str(project_root / "CATIA_BOM_GUI"), str(project_root)],
    binaries=[],
    datas=[
        (str(converter), "模板转化"),
        (str(icon_ico), "assets"),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="YR_CATIA_BOM_Manager",
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

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="YR_CATIA_BOM_Manager",
)
