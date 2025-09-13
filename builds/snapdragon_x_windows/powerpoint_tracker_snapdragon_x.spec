# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Add the presentation package to the path
# Get the project root directory (3 levels up from this spec file)
spec_dir = os.path.dirname(os.path.abspath(SPEC))
project_root = os.path.dirname(os.path.dirname(spec_dir))
presentation_path = os.path.join(project_root, 'presentation')
sys.path.insert(0, presentation_path)

# Collect all data files and submodules for dependencies
datas = []
hiddenimports = []

# Add presentation_detector data files
try:
    datas += collect_data_files('presentation_detector')
    hiddenimports += collect_submodules('presentation_detector')
except:
    pass

# Add pptx data files
try:
    datas += collect_data_files('pptx')
    hiddenimports += collect_submodules('pptx')
except:
    pass

# Add opencv data files
try:
    datas += collect_data_files('cv2')
    hiddenimports += collect_submodules('cv2')
except:
    pass

# Add PIL/Pillow data files
try:
    datas += collect_data_files('PIL')
    hiddenimports += collect_submodules('PIL')
except:
    pass

# Add numpy data files
try:
    datas += collect_data_files('numpy')
    hiddenimports += collect_submodules('numpy')
except:
    pass

# Add pytesseract data files
try:
    datas += collect_data_files('pytesseract')
    hiddenimports += collect_submodules('pytesseract')
except:
    pass

# Add psutil data files
try:
    datas += collect_data_files('psutil')
    hiddenimports += collect_submodules('psutil')
except:
    pass

# Add platform-specific imports for Windows ARM64
try:
    hiddenimports += ['win32gui', 'win32process', 'win32api', 'win32con', 'win32ctypes']
except:
    pass

# Main analysis configuration
a = Analysis(
    [os.path.join(presentation_path, 'main_app.py')],
    pathex=[presentation_path],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'pandas', 'jupyter', 'notebook',
        'IPython', 'tkinter.test', 'test', 'unittest', 'doctest',
        'Cocoa', 'Quartz', 'Foundation', 'AppKit',  # macOS specific
        'pyobjc', 'pyobjc_core', 'pyobjc_framework'  # macOS specific
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# Remove duplicate entries
a.datas = list(set(a.datas))
a.binaries = list(set(a.binaries))

# PYZ configuration
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# EXE configuration for Snapdragon X Windows ARM64
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PowerPointTracker_SnapdragonX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='arm64',  # Snapdragon X ARM64 architecture
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path if you have one
    version=None,  # Add version info if needed
)

# Optional: Create a directory distribution instead of single file
# Uncomment the following lines if you prefer a directory distribution
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='PowerPointTracker_SnapdragonX'
# )
