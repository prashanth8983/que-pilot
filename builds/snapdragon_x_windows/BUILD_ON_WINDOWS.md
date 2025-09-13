# Building Windows .exe on Windows Laptop

This guide provides the exact commands to build the PowerPoint Tracker Windows .exe file on a Windows laptop.

## Prerequisites

- Windows 11 (recommended) or Windows 10
- Python 3.11 or 3.12 installed
- At least 4GB RAM
- 10GB free disk space

## Step-by-Step Commands

### 1. Open Command Prompt or PowerShell

```cmd
# Open Command Prompt as Administrator (recommended)
# Or use PowerShell
```

### 2. Navigate to Project Directory

```cmd
# Navigate to your project folder
cd C:\path\to\que-pilot\builds\snapdragon_x_windows
```

### 3. Create Virtual Environment

```cmd
# Create virtual environment
python -m venv venv_windows

# Activate virtual environment
venv_windows\Scripts\activate
```

### 4. Install Dependencies

```cmd
# Upgrade pip
python -m pip install --upgrade pip

# Install PyInstaller
pip install pyinstaller>=5.13.0

# Install core dependencies
pip install python-pptx>=0.6.21
pip install opencv-python>=4.8.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install numpy>=1.24.0
pip install psutil>=5.9.0

# Install Windows-specific dependencies
pip install pywin32>=306
```

### 5. Create Windows .exe Spec File

Create a file called `powerpoint_tracker_windows.spec` with this content:

```python
# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Add the presentation package to the path
presentation_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(SPEC))), 'presentation')
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

# Add other dependencies
try:
    datas += collect_data_files('pptx')
    hiddenimports += collect_submodules('pptx')
except:
    pass

try:
    datas += collect_data_files('cv2')
    hiddenimports += collect_submodules('cv2')
except:
    pass

try:
    datas += collect_data_files('PIL')
    hiddenimports += collect_submodules('PIL')
except:
    pass

try:
    datas += collect_data_files('numpy')
    hiddenimports += collect_submodules('numpy')
except:
    pass

try:
    datas += collect_data_files('pytesseract')
    hiddenimports += collect_submodules('pytesseract')
except:
    pass

try:
    datas += collect_data_files('psutil')
    hiddenimports += collect_submodules('psutil')
except:
    pass

# Add Windows-specific imports
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

# EXE configuration for Windows ARM64 (.exe file)
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
    icon=None,
    version=None,
)
```

### 6. Build the Windows .exe

```cmd
# Build using the spec file
pyinstaller --clean --noconfirm powerpoint_tracker_windows.spec
```

### 7. Verify the Build

```cmd
# Check if the .exe file was created
dir dist\PowerPointTracker_SnapdragonX.exe

# Check file properties
powershell "Get-ItemProperty 'dist\PowerPointTracker_SnapdragonX.exe' | Select-Object Name, Length, CreationTime"
```

### 8. Test the Executable

```cmd
# Run the executable to test
dist\PowerPointTracker_SnapdragonX.exe
```

## Alternative: Direct PyInstaller Command

If you prefer not to use a spec file, you can use this direct command:

```cmd
pyinstaller ^
    --clean ^
    --onefile ^
    --windowed ^
    --target-arch arm64 ^
    --name "PowerPointTracker_SnapdragonX" ^
    --add-data "..\presentation;presentation" ^
    --hidden-import "presentation_detector" ^
    --hidden-import "tkinter" ^
    --hidden-import "pptx" ^
    --hidden-import "cv2" ^
    --hidden-import "PIL" ^
    --hidden-import "numpy" ^
    --hidden-import "pytesseract" ^
    --hidden-import "psutil" ^
    --hidden-import "win32gui" ^
    --hidden-import "win32process" ^
    --hidden-import "win32api" ^
    --hidden-import "win32con" ^
    --hidden-import "win32ctypes" ^
    --exclude-module "matplotlib" ^
    --exclude-module "scipy" ^
    --exclude-module "pandas" ^
    --exclude-module "jupyter" ^
    --exclude-module "IPython" ^
    --exclude-module "test" ^
    --exclude-module "unittest" ^
    --exclude-module "doctest" ^
    --exclude-module "Cocoa" ^
    --exclude-module "Quartz" ^
    --exclude-module "Foundation" ^
    --exclude-module "AppKit" ^
    --exclude-module "pyobjc" ^
    --exclude-module "pyobjc_core" ^
    --exclude-module "pyobjc_framework" ^
    "..\presentation\main_app.py"
```

## Troubleshooting

### Common Issues

**1. Module not found errors:**
```cmd
# Add missing modules to hidden-imports
--hidden-import "missing_module_name"
```

**2. File not found errors:**
```cmd
# Check the path to presentation folder
dir ..\presentation\main_app.py
```

**3. Permission errors:**
```cmd
# Run Command Prompt as Administrator
```

**4. PyInstaller not found:**
```cmd
# Reinstall PyInstaller
pip install --upgrade pyinstaller
```

### Build Output

After successful build, you'll find:
- `dist\PowerPointTracker_SnapdragonX.exe` - Main executable
- `build\` - Temporary build files (can be deleted)

### File Properties

The created .exe file should be:
- **Size:** ~60-80MB
- **Format:** Windows PE executable
- **Architecture:** ARM64 (for Snapdragon X)
- **Type:** GUI application (no console window)

## Deployment

1. Copy `PowerPointTracker_SnapdragonX.exe` to your Snapdragon X Windows device
2. Ensure PowerPoint is installed on the target system
3. Run the executable (no installation required)
4. Grant necessary permissions for window detection

## Success Indicators

✅ **Build successful if:**
- No error messages during build
- `dist\PowerPointTracker_SnapdragonX.exe` exists
- File size is reasonable (60-80MB)
- Executable runs without immediate crashes

❌ **Build failed if:**
- Error messages during build
- Missing .exe file
- Executable crashes immediately
- Import errors when running

## Next Steps

After building successfully:
1. Test on Snapdragon X Windows device
2. Verify PowerPoint integration
3. Check window detection functionality
4. Create deployment package if needed
