#!/usr/bin/env python3
"""
Debug script to test PowerPoint slide detection.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

def debug_powerpoint_detection():
    """Debug the PowerPoint detection system."""
    print("🔍 DEBUGGING POWERPOINT SLIDE DETECTION")
    print("=" * 60)

    try:
        from core.presentation.detector import PowerPointWindowDetector
        print("✅ Successfully imported PowerPointWindowDetector")

        # Create detector
        detector = PowerPointWindowDetector()
        print("✅ Created PowerPointWindowDetector instance")

        # Test 1: Check for PowerPoint windows
        print("\n1️⃣ CHECKING FOR POWERPOINT WINDOWS:")
        windows = detector.get_powerpoint_windows()
        print(f"   Found {len(windows)} PowerPoint windows")

        for i, window in enumerate(windows):
            print(f"   Window {i+1}: {window}")

        # Test 2: Get active window
        print("\n2️⃣ CHECKING ACTIVE POWERPOINT WINDOW:")
        active_window = detector.get_active_powerpoint_window()
        if active_window:
            print(f"   Active window: {active_window}")
            print(f"   Title: '{active_window.title}'")

            # Test 3: Extract slide info from title
            print("\n3️⃣ EXTRACTING SLIDE INFO FROM TITLE:")
            slide_info = detector.extract_slide_info_from_title(active_window.title)
            print(f"   Slide info from title: {slide_info}")
        else:
            print("   No active PowerPoint window found")

        # Test 4: Direct AppleScript detection (macOS)
        print("\n4️⃣ TESTING DIRECT APPLESCRIPT DETECTION:")
        try:
            slide_info = detector.get_powerpoint_slide_info_macos()
            print(f"   AppleScript result: {slide_info}")

            if slide_info and slide_info.get('current_slide'):
                current_slide = slide_info.get('current_slide')
                total_slides = slide_info.get('total_slides')
                presentation_name = slide_info.get('presentation_name', 'Unknown')
                mode = slide_info.get('mode', 'unknown')
                slide_text = slide_info.get('slide_text', '')

                print(f"   📊 Current slide: {current_slide}")
                print(f"   📊 Total slides: {total_slides}")
                print(f"   📊 Presentation: {presentation_name}")
                print(f"   📊 Mode: {mode}")
                if slide_text:
                    print(f"   📝 Slide text: {slide_text[:100]}{'...' if len(slide_text) > 100 else ''}")
                else:
                    print(f"   📝 No slide text detected")

                if current_slide and total_slides:
                    print(f"   ✅ AppleScript detection: SUCCESS ({current_slide}/{total_slides})")
                else:
                    print(f"   ⚠️ AppleScript detection: PARTIAL (missing slide numbers)")
            else:
                print(f"   ❌ AppleScript detection: FAILED")

        except Exception as e:
            print(f"   💥 AppleScript error: {e}")

        # Test 5: Current slide info method
        print("\n5️⃣ TESTING get_current_slide_info() METHOD:")
        try:
            current_info = detector.get_current_slide_info()
            if current_info:
                print(f"   ✅ get_current_slide_info() returned: {current_info}")
            else:
                print(f"   ❌ get_current_slide_info() returned None")
        except Exception as e:
            print(f"   💥 get_current_slide_info() error: {e}")

    except ImportError as e:
        print(f"❌ Import error: {e}")
        return
    except Exception as e:
        print(f"❌ General error: {e}")
        import traceback
        traceback.print_exc()
        return

    print(f"\n" + "=" * 60)
    print("🎯 DEBUGGING COMPLETE")
    print("=" * 60)

    print(f"\n🔧 **NEXT STEPS:**")
    print(f"1. ✅ If AppleScript detection shows slide info → Sync logic needs fixing")
    print(f"2. ❌ If AppleScript detection fails → Need to check PowerPoint status")
    print(f"3. ⚠️ If partial detection → Need to improve AppleScript reliability")
    print(f"4. 🔄 Run this with PowerPoint open and on different slides to test")

if __name__ == "__main__":
    debug_powerpoint_detection()