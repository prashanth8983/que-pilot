#!/bin/bash

# Setup script for Windows .exe build via GitHub Actions
# This script helps set up the GitHub Actions workflow for building Windows .exe files

set -e  # Exit on any error

echo "🚀 Setting up Windows .exe build via GitHub Actions"
echo "=================================================="

# Get the project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
GITHUB_DIR="$PROJECT_ROOT/.github"
WORKFLOWS_DIR="$GITHUB_DIR/workflows"

echo "📁 Project root: $PROJECT_ROOT"
echo "📁 GitHub directory: $GITHUB_DIR"

# Create GitHub Actions directory structure
echo "📁 Creating GitHub Actions directory structure..."
mkdir -p "$WORKFLOWS_DIR"

# Check if the workflow file already exists
if [ -f "$WORKFLOWS_DIR/build-windows-exe.yml" ]; then
    echo "✅ GitHub Actions workflow already exists"
else
    echo "❌ GitHub Actions workflow not found"
    echo "   Please ensure the workflow file is committed to the repository"
fi

# Check if we're in a git repository
if [ ! -d "$PROJECT_ROOT/.git" ]; then
    echo "❌ Error: Not in a git repository"
    echo "   Please initialize git and commit your changes first:"
    echo "   git init"
    echo "   git add ."
    echo "   git commit -m 'Initial commit'"
    exit 1
fi

# Check git status
echo "📋 Checking git status..."
git status --porcelain

# Provide instructions for setting up GitHub Actions
echo ""
echo "🔧 To build Windows .exe files, follow these steps:"
echo ""
echo "1. 📤 Push your code to GitHub:"
echo "   git remote add origin <your-github-repo-url>"
echo "   git push -u origin main"
echo ""
echo "2. 🚀 Trigger the GitHub Actions workflow:"
echo "   • Go to your GitHub repository"
echo "   • Click on 'Actions' tab"
echo "   • Select 'Build Windows .exe for Snapdragon X' workflow"
echo "   • Click 'Run workflow' button"
echo ""
echo "3. ⬇️  Download the built .exe file:"
echo "   • Wait for the workflow to complete"
echo "   • Go to the 'Actions' tab"
echo "   • Click on the completed workflow run"
echo "   • Download the 'PowerPointTracker_SnapdragonX_Windows_EXE' artifact"
echo ""
echo "4. 🎯 Alternative: Automatic builds on push:"
echo "   • The workflow will automatically run when you push changes to:"
echo "     - presentation/ directory"
echo "     - builds/snapdragon_x_windows/ directory"
echo "     - .github/workflows/build-windows-exe.yml"
echo ""

# Check if GitHub CLI is available
if command -v gh &> /dev/null; then
    echo "🔧 GitHub CLI detected! You can also trigger the workflow via CLI:"
    echo "   gh workflow run 'Build Windows .exe for Snapdragon X'"
    echo ""
else
    echo "💡 Install GitHub CLI for easier workflow management:"
    echo "   brew install gh"
    echo "   gh auth login"
    echo ""
fi

# Create a simple test script
echo "🧪 Creating test script for Windows .exe..."
cat > "$PROJECT_ROOT/test_windows_exe.py" << 'EOF'
#!/usr/bin/env python3
"""
Test script for Windows .exe functionality
This script can be used to verify the Windows .exe works correctly
"""

import sys
import os

def test_imports():
    """Test that all required modules can be imported"""
    try:
        # Test core imports
        import tkinter as tk
        print("✅ tkinter imported successfully")
        
        import presentation_detector
        print("✅ presentation_detector imported successfully")
        
        from presentation_detector import PowerPointTracker, PowerPointWindowDetector
        print("✅ PowerPointTracker and PowerPointWindowDetector imported successfully")
        
        # Test other dependencies
        import pptx
        print("✅ python-pptx imported successfully")
        
        import cv2
        print("✅ opencv-python imported successfully")
        
        import PIL
        print("✅ Pillow imported successfully")
        
        import numpy
        print("✅ numpy imported successfully")
        
        import psutil
        print("✅ psutil imported successfully")
        
        # Test Windows-specific imports
        try:
            import win32gui
            print("✅ win32gui imported successfully")
        except ImportError:
            print("⚠️  win32gui not available (expected on non-Windows systems)")
        
        print("\n🎉 All imports successful! The Windows .exe should work correctly.")
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def test_basic_functionality():
    """Test basic functionality without GUI"""
    try:
        from presentation_detector import PowerPointWindowDetector
        
        # Test window detector initialization
        detector = PowerPointWindowDetector()
        print("✅ PowerPointWindowDetector initialized successfully")
        
        # Test getting windows (should work on any platform)
        windows = detector.get_powerpoint_windows()
        print(f"✅ Found {len(windows)} PowerPoint windows")
        
        return True
        
    except Exception as e:
        print(f"❌ Functionality test failed: {e}")
        return False

if __name__ == "__main__":
    print("🧪 Testing Windows .exe functionality...")
    print("=" * 50)
    
    # Test imports
    imports_ok = test_imports()
    
    if imports_ok:
        # Test basic functionality
        functionality_ok = test_basic_functionality()
        
        if functionality_ok:
            print("\n🎉 All tests passed! The Windows .exe is ready for deployment.")
            sys.exit(0)
        else:
            print("\n❌ Functionality tests failed.")
            sys.exit(1)
    else:
        print("\n❌ Import tests failed.")
        sys.exit(1)
EOF

chmod +x "$PROJECT_ROOT/test_windows_exe.py"

echo "✅ Test script created: test_windows_exe.py"
echo ""
echo "🎯 Next steps:"
echo "1. Commit and push your changes to GitHub"
echo "2. Run the GitHub Actions workflow"
echo "3. Download the Windows .exe file"
echo "4. Test the .exe file on a Snapdragon X Windows device"
echo ""
echo "📚 For more information, see:"
echo "   • builds/snapdragon_x_windows/BUILD_SNAPDRAGON_X.md"
echo "   • .github/workflows/build-windows-exe.yml"
echo ""
echo "🎉 Setup complete!"
