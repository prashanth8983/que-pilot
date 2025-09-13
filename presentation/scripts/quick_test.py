#!/usr/bin/env python3
"""
Quick functionality test for PowerPoint Tracker
"""

import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """Test that all modules can be imported"""
    print("🔍 Testing imports...")
    
    try:
        from presentation_detector import PowerPointTracker
        print("✅ PowerPointTracker imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import PowerPointTracker: {e}")
        return False
    
    try:
        from presentation_detector import PowerPointWindowDetector
        print("✅ PowerPointWindowDetector imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import PowerPointWindowDetector: {e}")
        return False
    
    return True

def test_basic_functionality():
    """Test basic functionality without PowerPoint files"""
    print("\n🧪 Testing basic functionality...")
    
    try:
        from presentation_detector import PowerPointTracker
        
        # Test initialization without file
        tracker = PowerPointTracker()
        print("✅ Tracker initialized without file")
        
        # Test initialization with auto-detect
        tracker = PowerPointTracker(auto_detect=True)
        print("✅ Tracker initialized with auto-detect")
        
        # Test basic properties
        assert tracker.current_slide_index == 0
        assert tracker.total_slides == 0
        assert tracker.auto_detect == True
        print("✅ Basic properties working")
        
        return True
        
    except Exception as e:
        print(f"❌ Basic functionality test failed: {e}")
        return False

def test_window_detection():
    """Test window detection functionality"""
    print("\n🪟 Testing window detection...")
    
    try:
        from presentation_detector import PowerPointWindowDetector
        
        detector = PowerPointWindowDetector()
        print("✅ Window detector initialized")
        
        # Test PowerPoint window identification
        is_ppt = detector.is_powerpoint_window("Presentation.pptx - Microsoft PowerPoint", "Microsoft PowerPoint")
        assert is_ppt == True
        print("✅ PowerPoint window identification working")
        
        # Test slide info extraction
        slide_info = detector.extract_slide_info_from_title("Slide 1 of 10")
        assert slide_info['current_slide'] == 1
        assert slide_info['total_slides'] == 10
        print("✅ Slide info extraction working")
        
        # Test getting windows (may return empty list if no PowerPoint running)
        windows = detector.get_powerpoint_windows()
        print(f"✅ Found {len(windows)} PowerPoint windows")
        
        return True
        
    except Exception as e:
        print(f"❌ Window detection test failed: {e}")
        return False

def test_dependencies():
    """Test that all required dependencies are available"""
    print("\n📦 Testing dependencies...")
    
    dependencies = [
        ('pptx', 'python-pptx'),
        ('cv2', 'opencv-python'),
        ('PIL', 'Pillow'),
        ('numpy', 'numpy'),
        ('pytesseract', 'pytesseract'),
        ('psutil', 'psutil'),
    ]
    
    all_good = True
    
    for module_name, package_name in dependencies:
        try:
            __import__(module_name)
            print(f"✅ {package_name} available")
        except ImportError:
            print(f"❌ {package_name} not available")
            all_good = False
    
    # Test platform-specific dependencies
    import platform
    system = platform.system()
    
    if system == "Darwin":  # macOS
        try:
            import Cocoa
            print("✅ Cocoa (macOS) available")
        except ImportError:
            print("⚠️ Cocoa (macOS) not available - window detection may be limited")
    
    elif system == "Windows":
        try:
            import win32gui
            print("✅ win32gui (Windows) available")
        except ImportError:
            print("⚠️ win32gui (Windows) not available - window detection may be limited")
    
    return all_good

def test_gui_availability():
    """Test GUI framework availability"""
    print("\n🖥️ Testing GUI availability...")
    
    try:
        import tkinter as tk
        print("✅ tkinter available")
        
        # Try to create a simple window (don't show it)
        root = tk.Tk()
        root.withdraw()  # Hide the window
        root.destroy()
        print("✅ tkinter window creation working")
        
        return True
        
    except Exception as e:
        print(f"❌ GUI test failed: {e}")
        return False

def main():
    """Run all quick tests"""
    print("🚀 PowerPoint Tracker - Quick Test")
    print("=" * 40)
    
    tests = [
        ("Import Test", test_imports),
        ("Basic Functionality", test_basic_functionality),
        ("Window Detection", test_window_detection),
        ("Dependencies", test_dependencies),
        ("GUI Availability", test_gui_availability),
    ]
    
    results = []
    
    for test_name, test_func in tests:
        print(f"\n📋 Running {test_name}...")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ {test_name} crashed: {e}")
            results.append((test_name, False))
    
    # Print summary
    print("\n" + "=" * 40)
    print("📊 Test Summary:")
    
    passed = 0
    total = len(results)
    
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"   {test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\nResults: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed! The PowerPoint Tracker is ready to use.")
        print("\n🚀 Next steps:")
        print("   • Run the GUI: python main_app.py")
        print("   • Try the demo: python auto_detect_demo.py")
        print("   • Run full tests: python test_app.py")
        return True
    else:
        print(f"\n⚠️ {total - passed} test(s) failed. Check the output above for details.")
        print("\n💡 Common solutions:")
        print("   • Install missing dependencies: pip install -r requirements.txt")
        print("   • Install Tesseract OCR for your platform")
        print("   • Check platform-specific dependencies")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
