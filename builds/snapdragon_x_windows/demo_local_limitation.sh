#!/bin/bash

# Demonstration of local Windows .exe build limitation
# This shows why PyInstaller on macOS cannot create true Windows .exe files

set -e  # Exit on any error

echo "🔬 Demonstrating Local Windows .exe Build Limitation"
echo "===================================================="

# Get the project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
BUILD_DIR="$PROJECT_ROOT/builds/snapdragon_x_windows"
DIST_DIR="$BUILD_DIR/dist_demo"

echo "📁 Project root: $PROJECT_ROOT"
echo "📁 Build dir: $BUILD_DIR"

# Activate virtual environment
VENV_DIR="$PROJECT_ROOT/venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "❌ Error: Virtual environment not found"
    exit 1
fi

echo "🔧 Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# Create demo directory
mkdir -p "$DIST_DIR"

# Try to build with PyInstaller
echo "🔨 Attempting to build with PyInstaller..."
echo "   Target: Windows .exe file"
echo "   Platform: macOS (host)"
echo ""

cd "$BUILD_DIR"
pyinstaller \
    --clean \
    --onefile \
    --windowed \
    --name "PowerPointTracker_Demo" \
    --distpath "$DIST_DIR" \
    --workpath "$BUILD_DIR/build_demo" \
    --add-data "../../presentation:presentation" \
    --hidden-import "presentation_detector" \
    --hidden-import "tkinter" \
    --hidden-import "pptx" \
    --hidden-import "cv2" \
    --hidden-import "PIL" \
    --hidden-import "numpy" \
    --hidden-import "pytesseract" \
    --hidden-import "psutil" \
    "../../presentation/main_app.py"

# Check what was actually created
echo ""
echo "🔍 Analyzing the created file..."
if [ -f "$DIST_DIR/PowerPointTracker_Demo" ]; then
    echo "📁 File created: $DIST_DIR/PowerPointTracker_Demo"
    echo "📊 File size: $(du -h "$DIST_DIR/PowerPointTracker_Demo" | cut -f1)"
    echo ""
    echo "🔬 File type analysis:"
    file "$DIST_DIR/PowerPointTracker_Demo"
    echo ""
    echo "📋 File header (first 16 bytes):"
    hexdump -C "$DIST_DIR/PowerPointTracker_Demo" | head -1
    echo ""
    echo "❌ CONCLUSION: This is a macOS executable (Mach-O format), NOT a Windows .exe file!"
    echo ""
    echo "🔍 Why this happens:"
    echo "   • PyInstaller creates executables for the platform it's running on"
    echo "   • macOS PyInstaller → macOS executable (Mach-O)"
    echo "   • Windows PyInstaller → Windows executable (.exe)"
    echo "   • Linux PyInstaller → Linux executable (ELF)"
    echo ""
    echo "💡 This is why we need alternative approaches:"
    echo "   1. GitHub Actions (Windows runner)"
    echo "   2. Docker with Windows containers"
    echo "   3. Windows virtual machine"
    echo "   4. Windows machine directly"
    
else
    echo "❌ No file was created"
fi

echo ""
echo "🎯 RECOMMENDATION:"
echo "   Use GitHub Actions for reliable Windows .exe builds:"
echo "   ./setup_windows_build.sh"
