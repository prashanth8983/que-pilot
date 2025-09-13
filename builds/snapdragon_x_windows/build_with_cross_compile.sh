#!/bin/bash

# Experimental: Attempt cross-compilation using various tools
# This is experimental and may not work reliably

set -e  # Exit on any error

echo "🔬 Experimental: Cross-compilation attempt"
echo "========================================="

echo "⚠️  WARNING: This is experimental and may not work!"
echo "   PyInstaller on macOS cannot create true Windows .exe files"
echo "   This script attempts various workarounds"
echo ""

# Check if we're on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ Error: This script is designed for macOS"
    exit 1
fi

# Get the project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
BUILD_DIR="$PROJECT_ROOT/builds/snapdragon_x_windows"
DIST_DIR="$BUILD_DIR/dist_cross_compile"

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

# Method 1: Try with different PyInstaller options
echo "🔬 Method 1: PyInstaller with Windows-specific options..."
mkdir -p "$DIST_DIR/method1"

cd "$BUILD_DIR"
pyinstaller \
    --clean \
    --onefile \
    --windowed \
    --name "PowerPointTracker_SnapdragonX_Method1" \
    --distpath "$DIST_DIR/method1" \
    --workpath "$BUILD_DIR/build_method1" \
    --add-data "../presentation:presentation" \
    --hidden-import "presentation_detector" \
    --hidden-import "tkinter" \
    --hidden-import "pptx" \
    --hidden-import "cv2" \
    --hidden-import "PIL" \
    --hidden-import "numpy" \
    --hidden-import "pytesseract" \
    --hidden-import "psutil" \
    --exclude-module "matplotlib" \
    --exclude-module "scipy" \
    --exclude-module "pandas" \
    --exclude-module "jupyter" \
    --exclude-module "IPython" \
    --exclude-module "test" \
    --exclude-module "unittest" \
    --exclude-module "doctest" \
    --exclude-module "Cocoa" \
    --exclude-module "Quartz" \
    --exclude-module "Foundation" \
    --exclude-module "AppKit" \
    --exclude-module "pyobjc" \
    --exclude-module "pyobjc_core" \
    --exclude-module "pyobjc_framework" \
    "../presentation/main_app.py"

# Check result
if [ -f "$DIST_DIR/method1/PowerPointTracker_SnapdragonX_Method1" ]; then
    echo "📁 Method 1 result: $DIST_DIR/method1/PowerPointTracker_SnapdragonX_Method1"
    file "$DIST_DIR/method1/PowerPointTracker_SnapdragonX_Method1"
else
    echo "❌ Method 1 failed"
fi

# Method 2: Try with console mode
echo ""
echo "🔬 Method 2: PyInstaller with console mode..."
mkdir -p "$DIST_DIR/method2"

pyinstaller \
    --clean \
    --onefile \
    --console \
    --name "PowerPointTracker_SnapdragonX_Method2" \
    --distpath "$DIST_DIR/method2" \
    --workpath "$BUILD_DIR/build_method2" \
    --add-data "../presentation:presentation" \
    --hidden-import "presentation_detector" \
    --hidden-import "tkinter" \
    --hidden-import "pptx" \
    --hidden-import "cv2" \
    --hidden-import "PIL" \
    --hidden-import "numpy" \
    --hidden-import "pytesseract" \
    --hidden-import "psutil" \
    --exclude-module "matplotlib" \
    --exclude-module "scipy" \
    --exclude-module "pandas" \
    --exclude-module "jupyter" \
    --exclude-module "IPython" \
    --exclude-module "test" \
    --exclude-module "unittest" \
    --exclude-module "doctest" \
    --exclude-module "Cocoa" \
    --exclude-module "Quartz" \
    --exclude-module "Foundation" \
    --exclude-module "AppKit" \
    --exclude-module "pyobjc" \
    --exclude-module "pyobjc_core" \
    --exclude-module "pyobjc_framework" \
    "../presentation/main_app.py"

# Check result
if [ -f "$DIST_DIR/method2/PowerPointTracker_SnapdragonX_Method2" ]; then
    echo "📁 Method 2 result: $DIST_DIR/method2/PowerPointTracker_SnapdragonX_Method2"
    file "$DIST_DIR/method2/PowerPointTracker_SnapdragonX_Method2"
else
    echo "❌ Method 2 failed"
fi

# Method 3: Try with different target architecture
echo ""
echo "🔬 Method 3: PyInstaller with explicit architecture..."
mkdir -p "$DIST_DIR/method3"

pyinstaller \
    --clean \
    --onefile \
    --windowed \
    --target-arch arm64 \
    --name "PowerPointTracker_SnapdragonX_Method3" \
    --distpath "$DIST_DIR/method3" \
    --workpath "$BUILD_DIR/build_method3" \
    --add-data "../presentation:presentation" \
    --hidden-import "presentation_detector" \
    --hidden-import "tkinter" \
    --hidden-import "pptx" \
    --hidden-import "cv2" \
    --hidden-import "PIL" \
    --hidden-import "numpy" \
    --hidden-import "pytesseract" \
    --hidden-import "psutil" \
    --exclude-module "matplotlib" \
    --exclude-module "scipy" \
    --exclude-module "pandas" \
    --exclude-module "jupyter" \
    --exclude-module "IPython" \
    --exclude-module "test" \
    --exclude-module "unittest" \
    --exclude-module "doctest" \
    --exclude-module "Cocoa" \
    --exclude-module "Quartz" \
    --exclude-module "Foundation" \
    --exclude-module "AppKit" \
    --exclude-module "pyobjc" \
    --exclude-module "pyobjc_core" \
    --exclude-module "pyobjc_framework" \
    "../presentation/main_app.py"

# Check result
if [ -f "$DIST_DIR/method3/PowerPointTracker_SnapdragonX_Method3" ]; then
    echo "📁 Method 3 result: $DIST_DIR/method3/PowerPointTracker_SnapdragonX_Method3"
    file "$DIST_DIR/method3/PowerPointTracker_SnapdragonX_Method3"
else
    echo "❌ Method 3 failed"
fi

echo ""
echo "🔍 Summary of cross-compilation attempts:"
echo "=========================================="

for method in method1 method2 method3; do
    if [ -d "$DIST_DIR/$method" ]; then
        echo "📁 $method:"
        ls -la "$DIST_DIR/$method/"
        if [ -f "$DIST_DIR/$method/PowerPointTracker_SnapdragonX_${method^}" ]; then
            file "$DIST_DIR/$method/PowerPointTracker_SnapdragonX_${method^}"
        fi
        echo ""
    fi
done

echo "💡 Conclusion:"
echo "   All methods will create macOS executables (Mach-O format)"
echo "   PyInstaller on macOS cannot create true Windows .exe files"
echo ""
echo "🎯 Recommended alternatives:"
echo "   1. Use GitHub Actions (easiest)"
echo "   2. Use Docker with Windows containers"
echo "   3. Use Windows virtual machine"
echo "   4. Use Windows machine directly"
echo ""
echo "🔗 For GitHub Actions setup, run:"
echo "   ./setup_windows_build.sh"
