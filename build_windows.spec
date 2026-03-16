# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller spec file untuk PKBM Generator
Mode: --onefile (single EXE, semua bundled di dalam)
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
project_path = os.path.abspath(SPECPATH)

# ============================================================
# DATA FILES - semua folder yang harus ter-bundle dalam EXE
# ============================================================
datas = [
    ('assets', 'assets'),      # Logo, fonts, signatures
    ('lib', 'lib'),            # Generator modules
    ('data', 'data'),          # Template Excel files
]

# ============================================================
# HIDDEN IMPORTS
# ============================================================
hiddenimports = [
    'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox',
    'PIL', 'PIL.Image', 'PIL.ImageTk',
    'reportlab', 'reportlab.pdfgen', 'reportlab.pdfgen.canvas',
    'reportlab.lib', 'reportlab.lib.pagesizes', 'reportlab.lib.units',
    'reportlab.lib.colors', 'reportlab.lib.styles', 'reportlab.lib.enums',
    'reportlab.platypus', 'reportlab.platypus.tables', 'reportlab.platypus.paragraph',
    'reportlab.pdfbase', 'reportlab.pdfbase.ttfonts', 'reportlab.pdfbase.pdfmetrics',
    'pandas', 'numpy', 'openpyxl', 'openpyxl.cell', 'openpyxl.styles',
    'docx', 'docx.shared', 'docx.enum.text',
    'cv2', 'threading', 'shutil', 'datetime', 'pathlib',
]

hiddenimports += collect_submodules('reportlab')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('PIL')

datas += collect_data_files('reportlab')

# ============================================================
# ANALYSIS
# ============================================================
a = Analysis(
    ['pkbm_generator_app.py'],
    pathex=[project_path],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'pytest',
        'setuptools', 'distutils',
        'test', 'tests', 'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ============================================================
# ONE-FOLDER BUNDLE (--onedir mode)
# ============================================================
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PKBM_Generator',
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
    icon='assets/logo.png' if os.path.exists('assets/logo.png') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PKBM_Generator'
)
