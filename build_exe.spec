# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src\\app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'matplotlib',
        'matplotlib.backends.backend_qtagg',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'unittest',
        'email',
        'http',
        'xml',
        'pydoc',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TrackMyWorkout',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # ללא חלון קונסול
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # נוסיף אייקון אם יש
    version_info={
        'CompanyName': 'Fitness Tracker',
        'FileDescription': 'אפליקציית מעקב אימונים',
        'FileVersion': '1.0.0.0',
        'InternalName': 'TrackMyWorkout',
        'LegalCopyright': '© 2025',
        'OriginalFilename': 'TrackMyWorkout.exe',
        'ProductName': 'Track My Workout',
        'ProductVersion': '1.0.0.0',
    },
)
