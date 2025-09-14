#!/usr/bin/env python3
"""
Debug script for macOS PowerPoint .ppt file detection.
This will help diagnose issues with .ppt file detection specifically.
"""

import sys
import subprocess
import time
from pathlib import Path

# Add the src directory to the path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_applescript_basic():
    """Test basic AppleScript PowerPoint connection"""
    print("=== Testing Basic AppleScript Connection ===")

    applescript = '''
    tell application "Microsoft PowerPoint"
        try
            return "PowerPoint is running: " & (count of presentations) & " presentations open"
        on error errMsg
            return "Error: " & errMsg
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', applescript],
                              capture_output=True, text=True, timeout=10)
        print(f"✓ AppleScript result: {result.stdout.strip()}")
        return result.returncode == 0
    except Exception as e:
        print(f"✗ AppleScript failed: {e}")
        return False

def test_presentation_details():
    """Get detailed information about the current presentation"""
    print("\n=== Getting Presentation Details ===")

    applescript = '''
    tell application "Microsoft PowerPoint"
        try
            if (count of presentations) > 0 then
                set currentPresentation to active presentation
                set presentationName to name of currentPresentation
                set presentationPath to full name of currentPresentation
                set totalSlides to count of slides of currentPresentation
                set fileFormat to file format of currentPresentation

                return presentationName & "|" & presentationPath & "|" & totalSlides & "|" & fileFormat
            else
                return "No presentations open"
            end if
        on error errMsg
            return "Error getting details: " & errMsg
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', applescript],
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            parts = result.stdout.strip().split("|")
            if len(parts) >= 4:
                print(f"✓ Presentation Name: {parts[0]}")
                print(f"✓ Path: {parts[1]}")
                print(f"✓ Total Slides: {parts[2]}")
                print(f"✓ File Format: {parts[3]}")
                return parts
            else:
                print(f"! Response: {result.stdout.strip()}")
        return None
    except Exception as e:
        print(f"✗ Failed to get presentation details: {e}")
        return None

def test_slide_detection_methods():
    """Test different methods to detect current slide"""
    print("\n=== Testing Slide Detection Methods ===")

    methods = [
        ("Slideshow Window", '''
        tell application "Microsoft PowerPoint"
            try
                if (count of slide show windows) > 0 then
                    set currentSlide to slide number of slide of slide show view of slide show window 1
                    return "slideshow|" & currentSlide
                else
                    return "slideshow|no_slideshow_window"
                end if
            on error errMsg
                return "slideshow|error: " & errMsg
            end try
        end tell
        '''),

        ("Document Window View", '''
        tell application "Microsoft PowerPoint"
            try
                if (count of document windows) > 0 then
                    set docWin to document window 1
                    set viewObj to view of docWin
                    set slideIdx to slide index of viewObj
                    return "docview|" & slideIdx
                else
                    return "docview|no_document_window"
                end if
            on error errMsg
                return "docview|error: " & errMsg
            end try
        end tell
        '''),

        ("Selection Method", '''
        tell application "Microsoft PowerPoint"
            try
                if (count of document windows) > 0 then
                    set selectionObj to selection of document window 1
                    if (count of slides of selectionObj) > 0 then
                        set selectedSlide to slide 1 of selectionObj
                        set slideNum to slide number of selectedSlide
                        return "selection|" & slideNum
                    else
                        return "selection|no_slides_selected"
                    end if
                else
                    return "selection|no_document_window"
                end if
            on error errMsg
                return "selection|error: " & errMsg
            end try
        end tell
        ''')
    ]

    results = {}
    for method_name, script in methods:
        try:
            result = subprocess.run(['osascript', '-e', script],
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                response = result.stdout.strip()
                print(f"✓ {method_name}: {response}")
                results[method_name] = response
            else:
                print(f"✗ {method_name}: Failed")
        except Exception as e:
            print(f"✗ {method_name}: {e}")

    return results

def test_window_detection():
    """Test Quartz window detection"""
    print("\n=== Testing Quartz Window Detection ===")

    try:
        from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

        window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)
        ppt_windows = []

        for window_info in window_list:
            window_title = window_info.get('kCGWindowName', '')
            owner_name = window_info.get('kCGWindowOwnerName', '')
            bounds = window_info.get('kCGWindowBounds', {})

            if owner_name == 'Microsoft PowerPoint':
                ppt_windows.append({
                    'title': window_title,
                    'owner': owner_name,
                    'size': f"{bounds.get('Width', 0)}x{bounds.get('Height', 0)}"
                })

        print(f"✓ Found {len(ppt_windows)} PowerPoint windows:")
        for i, window in enumerate(ppt_windows):
            print(f"  {i+1}. '{window['title']}' ({window['size']})")

        return ppt_windows

    except ImportError:
        print("✗ Quartz not available. Install: pip install pyobjc-framework-Quartz")
        return []
    except Exception as e:
        print(f"✗ Window detection failed: {e}")
        return []

def test_existing_detector():
    """Test our existing detector implementation"""
    print("\n=== Testing Existing Detector Implementation ===")

    try:
        from core.presentation.detector import PowerPointWindowDetector

        detector = PowerPointWindowDetector()
        print(f"✓ Detector initialized")

        # Test window detection
        windows = detector.get_powerpoint_windows()
        print(f"✓ Found {len(windows)} PowerPoint windows via detector")

        for window in windows:
            print(f"  - {window.title} ({window.app_name})")

        # Test slide info
        if windows:
            active = detector.get_active_powerpoint_window()
            if active:
                print(f"✓ Active window: {active.title}")

                # Test AppleScript method
                applescript_info = detector.get_powerpoint_slide_info_macos()
                print(f"✓ AppleScript info: {applescript_info}")

                # Test title extraction
                slide_info = detector.extract_slide_info_from_title(active.title)
                print(f"✓ Title extraction: {slide_info}")

        return True

    except Exception as e:
        print(f"✗ Existing detector failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("macOS PowerPoint .ppt Detection Debug Tool")
    print("=" * 50)
    print("Make sure PowerPoint is open with a .ppt file!")
    print()

    # Run all tests
    applescript_works = test_applescript_basic()

    if applescript_works:
        presentation_details = test_presentation_details()
        slide_methods = test_slide_detection_methods()

    windows = test_window_detection()
    detector_works = test_existing_detector()

    print("\n" + "=" * 50)
    print("SUMMARY")
    print("=" * 50)
    print(f"AppleScript Connection: {'✓' if applescript_works else '✗'}")
    print(f"Window Detection: {'✓' if windows else '✗'}")
    print(f"Existing Detector: {'✓' if detector_works else '✗'}")

    if not applescript_works:
        print("\n⚠️  PowerPoint may not be running or AppleScript access denied")
        print("   Try: System Preferences → Security & Privacy → Privacy → Automation")

    if applescript_works and not windows:
        print("\n⚠️  AppleScript works but no windows detected")
        print("   This suggests a Quartz/window detection issue")

if __name__ == "__main__":
    main()