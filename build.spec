# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Smart PDF Data Extractor Pro
# Optimized for size and performance

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include data files if needed (currently none required)
    ],
    hiddenimports=[
        'pdfplumber',
        'pytesseract',
        'PIL',
        'openpyxl',
        'google.auth',
        'google.auth.transport.requests',
        'google.oauth2.credentials',
        'google.oauth2.service_account',
        'googleapiclient.discovery',
        'googleapiclient.http',
        'tkinter',
        're',
        'os',
        'threading',
        'tempfile',
        'shutil',
        'pathlib',
        'datetime',
        'typing',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unnecessary modules to reduce size
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'pytest',
        'IPython',
        'jupyter',
        'notebook',
        'sphinx',
        'setuptools',
        'distutils',
        'wheel',
        'pip',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF_Extractor_Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Use UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI only)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,  # Optional app icon
    version='version_info.txt' if os.path.exists('version_info.txt') else None,
)