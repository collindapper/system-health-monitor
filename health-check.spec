# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['health-check.py'],
    pathex=[],
    binaries=[],
    datas=[('favicon.ico', '.'), ('C:\\Users\\dappec\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\emoji\\unicode_codes', 'emoji/unicode_codes'), ('hp-hpia-5.3.2.exe', '.'), ('system_update_5.08.03.59.exe', '.'), ('get_passcode_v2.py', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='health-check',
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
    icon=['favicon.ico'],
)
