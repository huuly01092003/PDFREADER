# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file - OPTIMIZED VERSION
# Smart PDF Data Extractor Pro v3.0

import sys
import os

block_cipher = None

# Hidden imports - chỉ giữ những gì thực sự cần
hidden_imports = [
    'pdfplumber',
    'pdfplumber.utils',
    'pytesseract',
    'PIL',
    'PIL._imaging',
    'openpyxl',
    'openpyxl.styles',
    'google.auth',
    'google.auth.transport.requests',
    'google.oauth2.service_account',
    'googleapiclient.discovery',
    'googleapiclient.http',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
]

# Exclude modules - loại bỏ thư viện không cần thiết
excludes = [
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
    'test',
    'unittest',
    'email',
    'http',
    'urllib3',
    'xml',
    'pydoc',
    'doctest',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Không cần thêm data files - app tự tạo
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Loại bỏ các file .pyc và __pycache__
a.datas = [x for x in a.datas if not x[0].startswith('__pycache__')]

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
    upx=True,  # Bật UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Không hiện console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
    version='version_info.txt' if os.path.exists('version_info.txt') else None,
    # Optimization flags
    optimize=2,  # Bytecode optimization level
)