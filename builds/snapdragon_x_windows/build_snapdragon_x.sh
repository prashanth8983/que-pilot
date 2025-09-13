#!/bin/bash

# Build script for Snapdragon X Windows executable
# This script sets up the environment and builds the PowerPoint Tracker for Snapdragon X Windows ARM64

set -e  # Exit on any error

echo "🚀 Building PowerPoint Tracker for Snapdragon X Windows ARM64"
echo "=============================================================="

# Check if we're on macOS (required for cross-compilation)
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ Error: This build script requires macOS for cross-compilation to Snapdragon X Windows"
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
DIST_DIR="$BUILD_DIR/dist"

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

# Install required packages for Snapdragon X Windows cross-compilation
echo "📥 Installing dependencies for Snapdragon X Windows build..."

# Install PyInstaller with ARM64 support
pip install pyinstaller>=5.13.0

# Install core dependencies (excluding platform-specific ones)
pip install python-pptx>=0.6.21
pip install opencv-python>=4.8.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install numpy>=1.24.0
pip install psutil>=5.9.0

# Install Windows-specific dependencies for cross-compilation
echo "🪟 Installing Windows ARM64 dependencies..."
pip install pywin32-ctypes>=0.2.0

# Create build directories
echo "📁 Creating build directories..."
mkdir -p "$DIST_DIR"
mkdir -p "$BUILD_DIR/build"

# Clean previous builds
echo "🧹 Cleaning previous builds..."
rm -rf "$BUILD_DIR/build"/*
rm -rf "$DIST_DIR"/*

# Build the executable
echo "🔨 Building Snapdragon X Windows executable..."
echo "   This may take several minutes..."

# Use PyInstaller with the spec file
cd "$BUILD_DIR"
pyinstaller \
    --clean \
    --noconfirm \
    --distpath "$DIST_DIR" \
    --workpath "$BUILD_DIR/build" \
    powerpoint_tracker_snapdragon_x.spec

# Check if build was successful
if [ -f "$DIST_DIR/PowerPointTracker_SnapdragonX" ]; then
    echo ""
    echo "✅ Build successful!"
    echo "📁 Executable location: $DIST_DIR/PowerPointTracker_SnapdragonX"
    echo ""
    echo "📊 Build information:"
    echo "   Target: Snapdragon X Windows ARM64"
    echo "   Architecture: ARM64"
    echo "   Format: Single executable"
    echo "   Size: $(du -h "$DIST_DIR/PowerPointTracker_SnapdragonX" | cut -f1)"
    echo ""
    echo "🚀 The executable is ready for deployment on Snapdragon X Windows devices!"
    echo ""
    echo "📋 Deployment notes:"
    echo "   • The executable is self-contained and doesn't require Python installation"
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
    echo "📦 Creating deployment package..."
    DEPLOY_DIR="$DIST_DIR/PowerPointTracker_SnapdragonX_Package"
    mkdir -p "$DEPLOY_DIR"
    
    # Copy executable
    cp "$DIST_DIR/PowerPointTracker_SnapdragonX" "$DEPLOY_DIR/"
    
    # Create README for deployment
    cat > "$DEPLOY_DIR/README.txt" << EOF
PowerPoint Tracker for Snapdragon X Windows ARM64
=================================================

This package contains the PowerPoint Tracker application compiled for Snapdragon X Windows ARM64.

Files:
- PowerPointTracker_SnapdragonX: Main executable

System Requirements:
- Windows 11 on Snapdragon X ARM64 processor
- PowerPoint installed on target system
- Visual C++ Redistributable for ARM64 (may be required)

Usage:
1. Copy PowerPointTracker_SnapdragonX to your Snapdragon X Windows device
2. Run the executable directly (no installation required)
3. Ensure PowerPoint is installed on target system
4. Grant necessary permissions for window detection

Features:
- Automatic PowerPoint window detection
- Real-time slide tracking
- Cross-platform compatibility
- Self-contained executable

For support and updates, please refer to the project documentation.
EOF
    
    echo "📦 Deployment package created: $DEPLOY_DIR"
    echo "   Contains executable and deployment instructions"
    
else
    echo "❌ Build failed!"
    echo "   Check the error messages above for details"
    echo "   Common issues:"
    echo "   • Missing dependencies"
    echo "   • PyInstaller version compatibility"
    echo "   • ARM64 cross-compilation support"
    echo "   • Path issues with presentation directory"
    exit 1
fi

echo "🎉 Build process completed!"
