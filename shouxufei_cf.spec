# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 添加必要的隐藏导入
hiddenimports = [
    'pandas',
    'openpyxl',
    'openpyxl.styles',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'numpy',  # pandas依赖
    'numpy.core._methods',
    'numpy.lib.format',
    'pandas._libs.tslibs.base',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.timezones',
    'pandas._libs.skiplist',
]

a = Analysis(
    ['shouxufei_cf.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

# 添加必要的二进制文件和数据文件
# a.binaries = a.binaries + []
# a.datas = a.datas + []

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='预支处理工具',
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
    icon=None,
)
