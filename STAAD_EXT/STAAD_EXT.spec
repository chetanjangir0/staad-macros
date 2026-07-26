# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all


comtypes_data, comtypes_binaries, comtypes_hidden_imports = collect_all("comtypes")

analysis = Analysis(
    ["src/staad_ext/frozen.py"],
    pathex=[],
    binaries=comtypes_binaries,
    datas=comtypes_data,
    hiddenimports=comtypes_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="STAAD_EXT",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

collection = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="STAAD_EXT",
)
