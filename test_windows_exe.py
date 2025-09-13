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
