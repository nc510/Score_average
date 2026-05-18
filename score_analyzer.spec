# -*- mode: python ; coding: utf-8 -*-

import sys
import os

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath('.')],
    binaries=[
        ('vv/Library/bin/tcl86t.dll', '.'),
        ('vv/Library/bin/tk86t.dll', '.'),
        ('vv/Library/bin/ffi.dll', '.'),
        ('vv/Library/bin/libcrypto-3-x64.dll', '.'),
        ('vv/Library/bin/libssl-3-x64.dll', '.'),
        ('vv/Library/bin/liblzma.dll', '.'),
        ('vv/Library/bin/libbz2.dll', '.'),
        ('vv/Library/bin/bzip2.dll', '.'),
        ('vv/Library/bin/libexpat.dll', '.'),
        ('vv/Library/bin/sqlite3.dll', '.'),
    ],
    datas=[
        ('ico.png', '.'),
        ('qr_codes.py', '.'),
        ('vv/Library/lib/tcl8.6', 'tcl8.6'),
        ('vv/Library/lib/tk8.6', 'tk8.6'),
    ],
    hiddenimports=['pandas', 'numpy', 'xlrd', 'openpyxl', 'PIL', 'tkinter', 'tkinter.ttk'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=True,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='成绩统计分析系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='ico.png',
)
