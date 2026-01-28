# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/brusli/coding/goodlist_script/Стокова Разписка скртип/desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Lagardere',
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
    icon=['/Users/brusli/coding/goodlist_script/Стокова Разписка скртип/icon.icns'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Lagardere',
)
app = BUNDLE(
    coll,
    name='Lagardere.app',
    icon='/Users/brusli/coding/goodlist_script/Стокова Разписка скртип/icon.icns',
    bundle_identifier=None,
)
