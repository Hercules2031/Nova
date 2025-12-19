# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 添加所有需要的數據文件
datas = [
    # 如果有圖標文件可以添加
    # ('icon.ico', '.'),
    # ('config.ini', '.'),
]

# 添加隱藏的導入模塊（如果有的話）
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'pandas',
    'numpy',
    'openpyxl',
    'openpyxl.styles',
    'csv',
    'os',
    'sys',
    'ast',
    'datetime',
    'threading',
    'warnings',
    # 如果你的 import 模塊有問題，添加在這裡
    # 'amazon_inventory',
    # 'shipmonk_order',
    # 'shipmonk_inventory',
    # 'shopify',
]

a = Analysis(
    ['your_main_file.py'],  # 你的主程序文件
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

# 打包成單個 EXE 文件
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Nova_Excel_Viewer',  # EXE 文件名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 使用 UPX 壓縮（需要安裝 upx）
    console=False,  # 設置為 True 可以看到 console 輸出，False 隱藏 console
    icon=None,  # 如果有圖標文件可以添加 'icon.ico'
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
    name='Nova_Excel_Viewer',
)