# -*- mode: python ; coding: utf-8 -*-

"""
Simplified PyInstaller spec file untuk PKBM Generator - macOS
Minimal dependencies, excludes problematic modules
"""

import os
import sys

block_cipher = None
project_path = os.path.abspath(SPECPATH)

# ============================================================
# DATA FILES
# ============================================================
datas = [
    ('assets', 'assets'),
    ('lib', 'lib'),
    ('data', 'data'),
]

# ============================================================
# MINIMAL HIDDEN IMPORTS
# ============================================================
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'PIL._tkinter_finder',
]

# ============================================================
# COMPREHENSIVE EXCLUDES
# ============================================================
excludes = [
    # Development tools
    'IPython', 'jupyter', 'notebook', 'nbformat', 'nbconvert', 'ipykernel',
    'jedi', 'parso', 'pygments', 'wcwidth', 'prompt_toolkit',
    
    # Testing frameworks
    'pytest', 'unittest', 'test', 'tests', '_pytest',
    
    # Build tools (causes distutils conflict)
    'setuptools', 'distutils', '_distutils_hack', 'pkg_resources',
    
    # Scientific computing (not needed)
    'matplotlib', 'scipy', 'numba', 'llvmlite',
    
    # Unused packages
    'qrcode', 'docx', 'python-docx',
    
    # Networking/web (not needed)
    'requests', 'urllib3', 'certifi', 'charset_normalizer',
    'zmq', 'tornado', 'asyncio',
    
    # Other heavy packages
    'lxml.html', 'jsonschema', 'referencing',
]

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
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ============================================================
# ONE-FOLDER BUNDLE
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

# ============================================================
# macOS APP BUNDLE
# ============================================================
app = BUNDLE(
    coll,
    name='PKBM_Generator.app',
    icon='assets/logo.icns' if os.path.exists('assets/logo.icns') else None,
    bundle_identifier='com.pkbm.generator',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSHighResolutionCapable': 'True',
        'CFBundleName': 'PKBM Generator',
        'CFBundleDisplayName': 'PKBM Generator',
        'CFBundleVersion': '1.2.0',
        'CFBundleShortVersionString': '1.2.0',
        'LSMinimumSystemVersion': '10.13.0',
        'NSRequiresAquaSystemAppearance': 'False',
    },
)
