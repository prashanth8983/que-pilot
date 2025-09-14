#!/usr/bin/env python3
"""
Debug script to understand why PPT is loaded but UI doesn't reflect it.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

def debug_presentation_loading():
    """Debug the presentation loading and UI state."""
    print("🔍 DEBUGGING PPT LOADING VS UI STATE")
    print("=" * 50)

    try:
        from services.presentation_service import PresentationService
        from core.presentation.detector import PowerPointWindowDetector
        print("✅ Successfully imported modules")
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return

    # Initialize services
    pres_service = PresentationService()
    detector = PowerPointWindowDetector()

    print(f"\n📊 INITIAL STATE:")
    print(f"   current_presentation_id: {pres_service.current_presentation_id}")
    print(f"   is_presenting: {pres_service.is_presenting}")

    # Try to load the sample PPT directly
    sample_ppt = "./sample_0.ppt"
    if os.path.exists(sample_ppt):
        print(f"\n🎯 TESTING DIRECT PPT LOADING:")
        print(f"   Loading: {sample_ppt}")

        try:
            success = pres_service.load_presentation(sample_ppt)
            print(f"   Load result: {success}")

            if success:
                print(f"   ✅ Direct loading successful!")
                summary = pres_service.get_presentation_summary()
                print(f"   📊 Summary: {summary}")

                slide_info = pres_service.get_current_slide_info()
                print(f"   📄 Slide info: {slide_info}")
            else:
                print(f"   ❌ Direct loading failed")

        except Exception as e:
            print(f"   💥 Error during direct loading: {e}")
            import traceback
            traceback.print_exc()

    # Try auto-detection
    print(f"\n🔍 TESTING AUTO-DETECTION:")
    try:
        # Reset state first
        pres_service.clear_presentation()
        print(f"   State cleared, current_presentation_id: {pres_service.current_presentation_id}")

        success = pres_service.auto_detect_presentation()
        print(f"   Auto-detect result: {success}")

        if success:
            print(f"   ✅ Auto-detection successful!")
            summary = pres_service.get_presentation_summary()
            print(f"   📊 Summary: {summary}")
        else:
            print(f"   ⚠️ Auto-detection failed")

    except Exception as e:
        print(f"   💥 Error during auto-detection: {e}")
        import traceback
        traceback.print_exc()

    # Test PowerPoint window detection
    print(f"\n🪟 TESTING POWERPOINT WINDOW DETECTION:")
    try:
        windows = detector.get_powerpoint_windows()
        print(f"   Found {len(windows)} PowerPoint windows")

        for i, window in enumerate(windows):
            print(f"   Window {i+1}: {window.title}")
            slide_info = detector.extract_slide_info_from_title(window.title)
            print(f"      Slide info: {slide_info}")

        if detector.system == "Darwin":
            print(f"   Testing macOS AppleScript detection...")
            macos_info = detector.get_powerpoint_slide_info_macos()
            print(f"   macOS detection: {macos_info}")

    except Exception as e:
        print(f"   💥 Error in window detection: {e}")
        import traceback
        traceback.print_exc()

    print(f"\n" + "=" * 50)
    print("🎯 DIAGNOSIS SUMMARY")
    print("=" * 50)

    # Check if presentation is loaded but UI doesn't know
    if pres_service.current_presentation_id:
        print("✅ Presentation IS loaded in service")
        print("   → UI should show presentation info")
        print("   → Check LiveView.refresh_presentation_data()")
    else:
        print("❌ No presentation loaded in service")
        print("   → This explains why UI shows 'waiting'")
        print("   → Auto-detection or file loading failed")

    # Check if the issue is in the UI update mechanism
    print(f"\nUI UPDATE MECHANISM CHECK:")
    print(f"   1. Presentation loaded: {'✅' if pres_service.current_presentation_id else '❌'}")
    print(f"   2. Callbacks registered: {len(pres_service.slide_change_callbacks)} slide, {len(pres_service.presentation_load_callbacks)} presentation")
    print(f"   3. Is presenting: {'✅' if pres_service.is_presenting else '❌'}")


if __name__ == "__main__":
    try:
        debug_presentation_loading()
    except Exception as e:
        print(f"\n💥 Debug script failed: {e}")
        import traceback
        traceback.print_exc()