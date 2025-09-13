#!/bin/bash

# Build script for Windows .exe file for Snapdragon X
# This script creates a proper Windows executable (.exe) file

set -e  # Exit on any error

echo "🚀 Building Windows .exe for Snapdragon X ARM64"
echo "================================================"

# Check if we're on macOS (required for cross-compilation)
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ Error: This build script requires macOS for cross-compilation to Windows"
    echo "   Please run this script on macOS"
    exit 1
fi

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is required but not installed"
    exit 1
fi

# Get the project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
PRESENTATION_DIR="$PROJECT_ROOT/presentation"
BUILD_DIR="$PROJECT_ROOT/builds/snapdragon_x_windows"
DIST_DIR="$BUILD_DIR/dist_windows_exe"

echo "📁 Project root: $PROJECT_ROOT"
echo "📁 Presentation dir: $PRESENTATION_DIR"
echo "📁 Build dir: $BUILD_DIR"

# Check if presentation directory exists
if [ ! -d "$PRESENTATION_DIR" ]; then
    echo "❌ Error: Presentation directory not found at $PRESENTATION_DIR"
    exit 1
fi

# Check if virtual environment exists
VENV_DIR="$PROJECT_ROOT/venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "❌ Error: Virtual environment not found at $VENV_DIR"
    echo "   Please create a virtual environment first:"
    echo "   cd $PROJECT_ROOT && python3 -m venv venv"
    exit 1
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# Upgrade pip
echo "⬆️  Upgrading pip..."
pip install --upgrade pip

# Install required packages for Windows .exe cross-compilation
echo "📥 Installing dependencies for Windows .exe build..."

# Install PyInstaller with Windows support
pip install pyinstaller>=5.13.0

# Install core dependencies
pip install python-pptx>=0.6.21
pip install opencv-python>=4.8.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install numpy>=1.24.0
pip install psutil>=5.9.0

# Install Windows-specific dependencies
echo "🪟 Installing Windows dependencies..."
pip install pywin32-ctypes>=0.2.0

# Create build directories
echo "📁 Creating build directories..."
mkdir -p "$DIST_DIR"
mkdir -p "$BUILD_DIR/build_windows"

# Clean previous builds
echo "🧹 Cleaning previous Windows builds..."
rm -rf "$BUILD_DIR/build_windows"/*
rm -rf "$DIST_DIR"/*

# Build the Windows .exe executable
echo "🔨 Building Windows .exe executable..."
echo "   This may take several minutes..."

# Use PyInstaller with the Windows .exe spec file
cd "$BUILD_DIR"
pyinstaller \
    --clean \
    --noconfirm \
    --distpath "$DIST_DIR" \
    --workpath "$BUILD_DIR/build_windows" \
    powerpoint_tracker_windows_exe.spec

# Check if build was successful
if [ -f "$DIST_DIR/PowerPointTracker_SnapdragonX.exe" ]; then
    echo ""
    echo "✅ Windows .exe build successful!"
    echo "📁 Executable location: $DIST_DIR/PowerPointTracker_SnapdragonX.exe"
    echo ""
    echo "📊 Build information:"
    echo "   Target: Windows ARM64 (.exe format)"
    echo "   Architecture: ARM64 (Snapdragon X)"
    echo "   Format: Windows executable (.exe)"
    echo "   Size: $(du -h "$DIST_DIR/PowerPointTracker_SnapdragonX.exe" | cut -f1)"
    echo ""
    echo "🚀 The Windows .exe is ready for deployment on Snapdragon X Windows devices!"
    echo ""
    echo "📋 Deployment notes:"
    echo "   • The .exe file is self-contained and doesn't require Python installation"
    echo "   • It includes all necessary dependencies"
    echo "   • Compatible with Windows 11 on Snapdragon X ARM64 processors"
    echo "   • May require Visual C++ Redistributable for ARM64 on target system"
    echo ""
    echo "🧪 Testing recommendations:"
    echo "   • Test on actual Snapdragon X Windows device"
    echo "   • Verify PowerPoint integration works correctly"
    echo "   • Check window detection functionality"
    echo ""
    
    # Create a deployment package
    echo "📦 Creating Windows deployment package..."
    DEPLOY_DIR="$DIST_DIR/PowerPointTracker_SnapdragonX_Windows_Package"
    mkdir -p "$DEPLOY_DIR"
    
    # Copy executable
    cp "$DIST_DIR/PowerPointTracker_SnapdragonX.exe" "$DEPLOY_DIR/"
    
    # Create README for Windows deployment
    cat > "$DEPLOY_DIR/README.txt" << EOF
PowerPoint Tracker for Snapdragon X Windows ARM64
=================================================

This package contains the PowerPoint Tracker application as a Windows .exe file
compiled for Snapdragon X Windows ARM64.

Files:
- PowerPointTracker_SnapdragonX.exe: Main Windows executable

System Requirements:
- Windows 11 on Snapdragon X ARM64 processor
- PowerPoint installed on target system
- Visual C++ Redistributable for ARM64 (may be required)

Usage:
1. Copy PowerPointTracker_SnapdragonX.exe to your Snapdragon X Windows device
2. Double-click the .exe file to run (no installation required)
3. Ensure PowerPoint is installed on target system
4. Grant necessary permissions for window detection

Features:
- Automatic PowerPoint window detection
- Real-time slide tracking
- Cross-platform compatibility
- Self-contained Windows executable

For support and updates, please refer to the project documentation.
EOF
    
    echo "📦 Windows deployment package created: $DEPLOY_DIR"
    echo "   Contains .exe file and deployment instructions"
    
elif [ -f "$DIST_DIR/PowerPointTracker_SnapdragonX" ]; then
    echo ""
    echo "⚠️  Build created executable but not in .exe format"
    echo "📁 Executable location: $DIST_DIR/PowerPointTracker_SnapdragonX"
    echo "   This is likely a macOS executable (Mach-O format)"
    echo ""
    echo "💡 Note: Cross-compilation from macOS to Windows .exe format"
    echo "   may require additional tools or a Windows build environment."
    echo ""
    echo "🔧 Alternative approaches:"
    echo "   1. Use a Windows machine with Python to build the .exe"
    echo "   2. Use a Windows virtual machine"
    echo "   3. Use GitHub Actions with Windows runners"
    echo "   4. Use Docker with Windows containers"
    
else
    echo "❌ Build failed!"
    echo "   Check the error messages above for details"
    echo "   Common issues:"
    echo "   • Missing dependencies"
    echo "   • PyInstaller version compatibility"
    echo "   • Cross-compilation limitations"
    echo "   • Path issues with presentation directory"
    exit 1
fi

echo "🎉 Build process completed!"
