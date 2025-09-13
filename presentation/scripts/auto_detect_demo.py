#!/usr/bin/env python3
"""
PowerPoint Auto-Detection Demo Script

This script demonstrates the automatic detection and monitoring capabilities
of the PowerPoint Tracker application.
"""

import time
import sys
from presentation_detector import PowerPointTracker, PowerPointWindowDetector

def demo_window_detection():
    """Demonstrate basic window detection"""
    print("🔍 PowerPoint Window Detection Demo")
    print("=" * 50)
    
    detector = PowerPointWindowDetector()
    
    # Get all PowerPoint windows
    windows = detector.get_powerpoint_windows()
    
    if windows:
        print(f"Found {len(windows)} PowerPoint window(s):")
        for i, window in enumerate(windows, 1):
            print(f"  {i}. {window}")
            slide_info = detector.extract_slide_info_from_title(window.title)
            if slide_info['current_slide']:
                print(f"     → Slide: {slide_info['current_slide']}")
                if slide_info['total_slides']:
                    print(f"       Total: {slide_info['total_slides']}")
                print(f"       Mode: {slide_info['mode']}")
    else:
        print("❌ No PowerPoint windows found.")
        print("   Please open a PowerPoint presentation and try again.")
        return False
    
    return True

def demo_auto_tracking():
    """Demonstrate automatic presentation tracking"""
    print("\n📊 Auto-Tracking Demo")
    print("=" * 50)
    
    # Create tracker with auto-detection
    tracker = PowerPointTracker(auto_detect=True)
    
    # Try to auto-load presentation
    if tracker.auto_load_presentation():
        print(f"✅ Auto-loaded presentation: {tracker.ppt_path}")
        print(f"   Total slides: {tracker.get_total_slides()}")
        
        # Enable auto-sync
        tracker.enable_auto_sync()
        print("✅ Auto-sync enabled")
        
        # Get window info
        window_info = tracker.get_window_info()
        if window_info:
            print(f"📱 Connected to: {window_info['window_title']}")
            slide_info = window_info['slide_info']
            if slide_info.get('current_slide'):
                print(f"📍 Current slide: {slide_info['current_slide']}")
                if slide_info.get('total_slides'):
                    print(f"📊 Total slides: {slide_info['total_slides']}")
        
        return True
    else:
        print("❌ Could not auto-load presentation")
        print("   Make sure PowerPoint is open with a presentation")
        return False

def demo_monitoring():
    """Demonstrate live monitoring"""
    print("\n👁️ Live Monitoring Demo")
    print("=" * 50)
    print("This will monitor PowerPoint for 30 seconds...")
    print("Try changing slides in PowerPoint to see the detection!")
    print("Press Ctrl+C to stop early.\n")
    
    detector = PowerPointWindowDetector()
    start_time = time.time()
    last_slide = None
    
    try:
        while time.time() - start_time < 30:  # Monitor for 30 seconds
            window = detector.get_active_powerpoint_window()
            
            if window:
                slide_info = detector.extract_slide_info_from_title(window.title)
                current_slide = slide_info.get('current_slide')
                
                if current_slide and current_slide != last_slide:
                    timestamp = time.strftime("%H:%M:%S")
                    presentation = slide_info.get('presentation_name', 'Unknown')
                    mode = slide_info.get('mode', 'unknown')
                    
                    print(f"[{timestamp}] 📄 Slide {current_slide} in '{presentation}' ({mode} mode)")
                    last_slide = current_slide
            else:
                if last_slide is not None:
                    timestamp = time.strftime("%H:%M:%S")
                    print(f"[{timestamp}] ❌ PowerPoint window lost")
                    last_slide = None
            
            time.sleep(1)  # Check every second
            
    except KeyboardInterrupt:
        print("\n\n⏹️ Monitoring stopped by user")
    
    print("\n✅ Monitoring demo completed")

def demo_sync_functionality():
    """Demonstrate sync functionality"""
    print("\n🔄 Sync Functionality Demo")
    print("=" * 50)
    
    tracker = PowerPointTracker(auto_detect=True)
    
    if not tracker.auto_load_presentation():
        print("❌ Could not auto-load presentation for sync demo")
        return False
    
    tracker.enable_auto_sync()
    print("✅ Auto-sync enabled")
    
    # Manual sync
    print("🔄 Performing manual sync...")
    if tracker.sync_with_powerpoint_window():
        current = tracker.get_current_slide_number()
        total = tracker.get_total_slides()
        print(f"✅ Synced to slide {current}/{total}")
    else:
        print("❌ Manual sync failed")
    
    # Get comprehensive info
    window_info = tracker.get_window_info()
    if window_info:
        print(f"\n📊 Window Information:")
        print(f"   Title: {window_info['window_title']}")
        print(f"   App: {window_info['app_name']}")
        if window_info['window_position']:
            print(f"   Position: {window_info['window_position']}")
        if window_info['window_size']:
            print(f"   Size: {window_info['window_size']}")
    
    return True

def main():
    """Main demo function"""
    print("🎯 PowerPoint Auto-Detection Demo")
    print("=" * 60)
    print("This demo will show you the automatic detection capabilities")
    print("of the PowerPoint Tracker application.")
    print()
    
    # Check if PowerPoint is running
    print("Step 1: Checking for PowerPoint windows...")
    if not demo_window_detection():
        print("\n💡 Tip: Open PowerPoint with a presentation and run this demo again.")
        return
    
    # Demo auto-tracking
    print("\nStep 2: Testing auto-tracking...")
    if not demo_auto_tracking():
        print("\n💡 Tip: Make sure PowerPoint has a presentation open.")
        return
    
    # Demo sync
    print("\nStep 3: Testing sync functionality...")
    demo_sync_functionality()
    
    # Ask user if they want to see monitoring
    print("\nStep 4: Live monitoring demo")
    response = input("Would you like to see live monitoring? (y/n): ").lower().strip()
    
    if response in ['y', 'yes']:
        demo_monitoring()
    else:
        print("⏭️ Skipping monitoring demo")
    
    print("\n🎉 Demo completed!")
    print("\n📋 Summary of capabilities demonstrated:")
    print("   ✅ Automatic PowerPoint window detection")
    print("   ✅ Auto-loading of presentation files")
    print("   ✅ Real-time slide synchronization")
    print("   ✅ Window information extraction")
    print("   ✅ Live monitoring (if selected)")
    
    print("\n🚀 Ready to use the PowerPoint Tracker!")
    print("   Run: python main_app.py")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Demo interrupted by user. Goodbye!")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Demo error: {e}")
        sys.exit(1)
