@echo off
REM Build script for Windows .exe on Windows laptop
REM This script automates the build process on Windows

echo 🚀 Building PowerPoint Tracker Windows .exe
echo ===========================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python is not installed or not in PATH
    echo    Please install Python 3.11 or 3.12 first
    pause
    exit /b 1
)

REM Check if we're in the right directory
if not exist "..\presentation\main_app.py" (
    echo ❌ Error: presentation\main_app.py not found
    echo    Please run this script from builds\snapdragon_x_windows\ directory
    pause
    exit /b 1
)

echo ✅ Python found
echo ✅ Project structure verified

REM Create virtual environment if it doesn't exist
if not exist "venv_windows" (
    echo 📦 Creating virtual environment...
    python -m venv venv_windows
    if errorlevel 1 (
        echo ❌ Failed to create virtual environment
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo 🔧 Activating virtual environment...
call venv_windows\Scripts\activate.bat

REM Upgrade pip
echo ⬆️  Upgrading pip...
python -m pip install --upgrade pip

REM Install dependencies
echo 📥 Installing dependencies...
pip install pyinstaller>=5.13.0
pip install python-pptx>=0.6.21
pip install opencv-python>=4.8.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install numpy>=1.24.0
pip install psutil>=5.9.0
pip install pywin32>=306

REM Create the spec file
echo 📝 Creating PyInstaller spec file...
(
echo # -*- mode: python ; coding: utf-8 -*-
echo.
echo import os
echo import sys
echo from PyInstaller.utils.hooks import collect_data_files, collect_submodules
echo.
echo # Add the presentation package to the path
echo presentation_path = os.path.join^(os.path.dirname^(os.path.dirname^(os.path.abspath^(SPEC^)^)^)^, 'presentation'^)
echo sys.path.insert^(0, presentation_path^)
echo.
echo # Collect all data files and submodules for dependencies
echo datas = []
echo hiddenimports = []
echo.
echo # Add presentation_detector data files
echo try:
echo     datas += collect_data_files^('presentation_detector'^)
echo     hiddenimports += collect_submodules^('presentation_detector'^)
echo except:
echo     pass
echo.
echo # Add other dependencies
echo try:
echo     datas += collect_data_files^('pptx'^)
echo     hiddenimports += collect_submodules^('pptx'^)
echo except:
echo     pass
echo.
echo try:
echo     datas += collect_data_files^('cv2'^)
echo     hiddenimports += collect_submodules^('cv2'^)
echo except:
echo     pass
echo.
echo try:
echo     datas += collect_data_files^('PIL'^)
echo     hiddenimports += collect_submodules^('PIL'^)
echo except:
echo     pass
echo.
echo try:
echo     datas += collect_data_files^('numpy'^)
echo     hiddenimports += collect_submodules^('numpy'^)
echo except:
echo     pass
echo.
echo try:
echo     datas += collect_data_files^('pytesseract'^)
echo     hiddenimports += collect_submodules^('pytesseract'^)
echo except:
echo     pass
echo.
echo try:
echo     datas += collect_data_files^('psutil'^)
echo     hiddenimports += collect_submodules^('psutil'^)
echo except:
echo     pass
echo.
echo # Add Windows-specific imports
echo try:
echo     hiddenimports += ['win32gui', 'win32process', 'win32api', 'win32con', 'win32ctypes']
echo except:
echo     pass
echo.
echo # Main analysis configuration
echo a = Analysis^(
echo     [os.path.join^(presentation_path, 'main_app.py'^)],
echo     pathex=[presentation_path],
echo     binaries=[],
echo     datas=datas,
echo     hiddenimports=hiddenimports,
echo     hookspath=[],
echo     hooksconfig={},
echo     runtime_hooks=[],
echo     excludes=[
echo         'matplotlib', 'scipy', 'pandas', 'jupyter', 'notebook',
echo         'IPython', 'tkinter.test', 'test', 'unittest', 'doctest',
echo         'Cocoa', 'Quartz', 'Foundation', 'AppKit',  # macOS specific
echo         'pyobjc', 'pyobjc_core', 'pyobjc_framework'  # macOS specific
echo     ],
echo     win_no_prefer_redirects=False,
echo     win_private_assemblies=False,
echo     cipher=None,
echo     noarchive=False,
echo ^)
echo.
echo # Remove duplicate entries
echo a.datas = list^(set^(a.datas^)^)
echo a.binaries = list^(set^(a.binaries^)^)
echo.
echo # PYZ configuration
echo pyz = PYZ^(a.pure, a.zipped_data, cipher=None^)
echo.
echo # EXE configuration for Windows ARM64 ^(.exe file^)
echo exe = EXE^(
echo     pyz,
echo     a.scripts,
echo     a.binaries,
echo     a.zipfiles,
echo     a.datas,
echo     [],
echo     name='PowerPointTracker_SnapdragonX',
echo     debug=False,
echo     bootloader_ignore_signals=False,
echo     strip=False,
echo     upx=True,
echo     upx_exclude=[],
echo     runtime_tmpdir=None,
echo     console=False,  # GUI application
echo     disable_windowed_traceback=False,
echo     argv_emulation=False,
echo     target_arch='arm64',  # Snapdragon X ARM64 architecture
echo     codesign_identity=None,
echo     entitlements_file=None,
echo     icon=None,
echo     version=None,
echo ^)
) > powerpoint_tracker_windows.spec

REM Build the executable
echo 🔨 Building Windows .exe executable...
echo    This may take several minutes...
pyinstaller --clean --noconfirm powerpoint_tracker_windows.spec

REM Check if build was successful
if exist "dist\PowerPointTracker_SnapdragonX.exe" (
    echo.
    echo ✅ Build successful!
    echo 📁 Executable location: dist\PowerPointTracker_SnapdragonX.exe
    
    REM Get file size
    for %%A in ("dist\PowerPointTracker_SnapdragonX.exe") do set size=%%~zA
    set /a sizeMB=%size%/1024/1024
    
    echo 📊 File size: %sizeMB% MB
    echo 🏗️  Architecture: ARM64 ^(Snapdragon X^)
    echo 🎯 Format: Windows .exe
    echo.
    echo 🚀 The Windows .exe is ready for deployment!
    echo.
    echo 📋 Deployment notes:
    echo    • The .exe file is self-contained and doesn't require Python installation
    echo    • It includes all necessary dependencies
    echo    • Compatible with Windows 11 on Snapdragon X ARM64 processors
    echo    • May require Visual C++ Redistributable for ARM64 on target system
    echo.
    echo 🧪 Testing recommendations:
    echo    • Test on actual Snapdragon X Windows device
    echo    • Verify PowerPoint integration works correctly
    echo    • Check window detection functionality
    echo.
    
    REM Create deployment package
    echo 📦 Creating deployment package...
    if not exist "PowerPointTracker_SnapdragonX_Windows_Package" mkdir "PowerPointTracker_SnapdragonX_Windows_Package"
    copy "dist\PowerPointTracker_SnapdragonX.exe" "PowerPointTracker_SnapdragonX_Windows_Package\"
    
    REM Create README
    (
    echo PowerPoint Tracker for Snapdragon X Windows ARM64
    echo =================================================
    echo.
    echo This package contains the PowerPoint Tracker application as a Windows .exe file
    echo compiled for Snapdragon X Windows ARM64.
    echo.
    echo Files:
    echo - PowerPointTracker_SnapdragonX.exe: Main Windows executable
    echo.
    echo System Requirements:
    echo - Windows 11 on Snapdragon X ARM64 processor
    echo - PowerPoint installed on target system
    echo - Visual C++ Redistributable for ARM64 ^(may be required^)
    echo.
    echo Usage:
    echo 1. Copy PowerPointTracker_SnapdragonX.exe to your Snapdragon X Windows device
    echo 2. Double-click the .exe file to run ^(no installation required^)
    echo 3. Ensure PowerPoint is installed on target system
    echo 4. Grant necessary permissions for window detection
    echo.
    echo Features:
    echo - Automatic PowerPoint window detection
    echo - Real-time slide tracking
    echo - Cross-platform compatibility
    echo - Self-contained Windows executable
    echo.
    echo For support and updates, please refer to the project documentation.
    ) > "PowerPointTracker_SnapdragonX_Windows_Package\README.txt"
    
    echo 📦 Deployment package created: PowerPointTracker_SnapdragonX_Windows_Package
    echo    Contains .exe file and deployment instructions
    
) else (
    echo ❌ Build failed!
    echo    Check the error messages above for details
    echo    Common issues:
    echo    • Missing dependencies
    echo    • PyInstaller version compatibility
    echo    • Path issues with presentation directory
    pause
    exit /b 1
)

echo.
echo 🎉 Build process completed!
echo.
echo 📋 Next steps:
echo 1. Test the .exe file on a Snapdragon X Windows device
echo 2. Verify PowerPoint integration works correctly
echo 3. Check window detection functionality
echo.
pause
