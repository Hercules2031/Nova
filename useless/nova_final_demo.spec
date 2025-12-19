# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['nova_final_demo.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['shipmonk_cred', 'shopify_cred', 'amazon_cred', 'amazon_inventory', 'amazon_order', 'shipmonk_inventory', 'shipmonk_order', 'shopify'],
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
    name='nova_final_demo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
