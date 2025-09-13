import platform
import time
import re
from typing import Optional, List, Dict, Tuple
import psutil

class WindowInfo:
    def __init__(self, window_id, title: str, app_name: str, position: Tuple[int, int] = None, size: Tuple[int, int] = None):
        self.window_id = window_id
        self.title = title
        self.app_name = app_name
        self.position = position
        self.size = size

    def __str__(self):
        return f"WindowInfo(id={self.window_id}, title='{self.title}', app='{self.app_name}')"

class PowerPointWindowDetector:
    def __init__(self):
        self.system = platform.system()
        self.powerpoint_process_names = [
            "Microsoft PowerPoint",
            "PowerPoint",
            "POWERPNT.EXE",
            "powerpnt.exe"
        ]
        self.current_window = None
        self.last_slide_info = None

    def is_powerpoint_window(self, window_title: str, app_name: str) -> bool:
        powerpoint_indicators = [
            "PowerPoint",
            "Microsoft PowerPoint",
            ".ppt",
            ".pptx",
            "Slide",
            "Presentation"
        ]

        title_lower = window_title.lower()
        app_lower = app_name.lower()

        for indicator in powerpoint_indicators:
            if indicator.lower() in title_lower or indicator.lower() in app_lower:
                return True
        return False

    def extract_slide_info_from_title(self, title: str) -> Dict:
        slide_info = {
            'current_slide': None,
            'total_slides': None,
            'presentation_name': None,
            'mode': 'unknown'
        }

        # For macOS, if title is empty or doesn't contain slide info, try AppleScript
        if self.system == "Darwin" and (not title or not any(pattern in title for pattern in ['Slide', '/', 'of'])):
            applescript_info = self.get_powerpoint_slide_info_macos()
            if applescript_info['current_slide']:
                return applescript_info

        # Fall back to title parsing for Windows or when AppleScript fails
        slide_patterns = [
            r'Slide (\d+) of (\d+)',
            r'(\d+)/(\d+)',
            r'(\d+) of (\d+)',
            r'Slide (\d+)',
        ]

        for pattern in slide_patterns:
            match = re.search(pattern, title, re.IGNORECASE)
            if match:
                groups = match.groups()
                slide_info['current_slide'] = int(groups[0])
                if len(groups) > 1:
                    slide_info['total_slides'] = int(groups[1])
                break

        if 'Slide Show' in title:
            slide_info['mode'] = 'slideshow'
        elif 'Normal' in title or 'Edit' in title:
            slide_info['mode'] = 'edit'
        else:
            slide_info['mode'] = 'normal'

        if ' - ' in title:
            parts = title.split(' - ')
            for part in parts:
                if '.ppt' in part.lower():
                    slide_info['presentation_name'] = part.strip()
                    break

        return slide_info

    def get_powerpoint_windows(self) -> List[WindowInfo]:
        if self.system == "Darwin":
            return self._get_powerpoint_windows_macos()
        elif self.system == "Windows":
            return self._get_powerpoint_windows_windows()
        else:
            return []

    def _get_powerpoint_windows_macos(self) -> List[WindowInfo]:
        windows = []
        try:
            from Cocoa import NSWorkspace, NSRunningApplication
            from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID
            import Cocoa

            workspace = NSWorkspace.sharedWorkspace()
            running_apps = workspace.runningApplications()

            powerpoint_apps = []
            for app in running_apps:
                app_name = app.localizedName()
                bundle_id = app.bundleIdentifier()
                if (any(pp_name.lower() in app_name.lower() for pp_name in self.powerpoint_process_names) or
                    (bundle_id and 'powerpoint' in bundle_id.lower())):
                    powerpoint_apps.append(app)

            if not powerpoint_apps:
                return windows

            window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)

            # Get presentation info from running app if possible
            presentation_title = self._get_powerpoint_presentation_title_macos()

            for window_info in window_list:
                window_title = window_info.get('kCGWindowName', '')
                owner_name = window_info.get('kCGWindowOwnerName', '')
                window_id = window_info.get('kCGWindowNumber')
                layer = window_info.get('kCGWindowLayer', 0)
                bounds = window_info.get('kCGWindowBounds', {})

                # Check if this is a PowerPoint window
                is_ppt_window = False

                # Method 1: Direct title/owner check
                if self.is_powerpoint_window(window_title, owner_name):
                    is_ppt_window = True

                # Method 2: PowerPoint app with significant size (likely main window)
                elif (owner_name == 'Microsoft PowerPoint' and
                      bounds.get('Width', 0) > 400 and bounds.get('Height', 0) > 300 and
                      layer == 0):  # Main layer
                    is_ppt_window = True

                # Use detected presentation title if window title is empty (applies to both methods)
                if is_ppt_window and not window_title and presentation_title:
                    window_title = presentation_title

                if is_ppt_window:
                    position = (bounds.get('X', 0), bounds.get('Y', 0))
                    size = (bounds.get('Width', 0), bounds.get('Height', 0))

                    windows.append(WindowInfo(
                        window_id=window_id,
                        title=window_title,
                        app_name=owner_name,
                        position=position,
                        size=size
                    ))

        except ImportError:
            print("macOS window detection requires pyobjc-framework-Cocoa and pyobjc-framework-Quartz")
            print("Install with: pip install pyobjc-framework-Cocoa pyobjc-framework-Quartz")
        except Exception as e:
            print(f"Error detecting macOS windows: {e}")

        return windows

    def _get_powerpoint_presentation_title_macos(self) -> str:
        """Try to get the current presentation title using AppleScript"""
        try:
            import subprocess

            # AppleScript to get PowerPoint presentation info
            # Use the same slide detection logic as get_powerpoint_slide_info_macos
            slide_info = self.get_powerpoint_slide_info_macos()
            if slide_info:
                presentation_name = slide_info.get('presentation_name', '')
                current_slide = slide_info.get('current_slide', 1)
                total_slides = slide_info.get('total_slides', 0)
                return f"{presentation_name} - Slide {current_slide} of {total_slides}"

        except Exception:
            pass  # Silently fail if AppleScript doesn't work

        return ""

    def get_powerpoint_slide_info_macos(self) -> Dict:
        """Get detailed slide information via AppleScript with improved reliability"""
        try:
            import subprocess

            # Simplified AppleScript that focuses on the most reliable methods first
            applescript = '''
            tell application "Microsoft PowerPoint"
                try
                    if (count of presentations) > 0 then
                        set currentPresentation to active presentation
                        set presentationName to name of currentPresentation
                        set totalSlides to count of slides of currentPresentation
                        set currentSlideNum to 1
                        set currentMode to "normal"
                        set slideDetected to false

                        -- Method 1: Check slide show mode first (most reliable when in presentation mode)
                        try
                            if (count of slide show windows) > 0 then
                                set currentSlideNum to slide number of slide of slide show view of slide show window 1
                                set currentMode to "slideshow"
                                set slideDetected to true
                            end if
                        on error
                            -- Slideshow method failed, continue to next method
                        end try

                        -- Method 2: Document window view (reliable for edit mode)
                        if not slideDetected then
                            try
                                if (count of document windows) > 0 then
                                    set docWin to document window 1
                                    set viewObj to view of docWin

                                    -- Try getting slide index from view
                                    try
                                        set slideIdx to slide index of viewObj
                                        if slideIdx > 0 and slideIdx <= totalSlides then
                                            set currentSlideNum to slideIdx
                                            set slideDetected to true
                                        end if
                                    on error
                                        -- slide index method failed
                                    end try

                                    -- Alternative: try getting slide from view
                                    if not slideDetected then
                                        try
                                            set slideRef to slide of viewObj
                                            if slideRef is not missing value then
                                                set currentSlideNum to slide number of slideRef
                                                set slideDetected to true
                                            end if
                                        on error
                                            -- slide of view method failed
                                        end try
                                    end if
                                end if
                            on error
                                -- Document window method failed
                            end try
                        end if

                        -- Method 3: Selection-based detection (fallback)
                        if not slideDetected then
                            try
                                if (count of document windows) > 0 then
                                    set docWin to document window 1
                                    set selectionObj to selection of docWin

                                    -- Try to get slide from selection
                                    try
                                        if (count of slides of selectionObj) > 0 then
                                            set selectedSlide to slide 1 of selectionObj
                                            set currentSlideNum to slide number of selectedSlide
                                            set slideDetected to true
                                        end if
                                    on error
                                        -- Selection method failed
                                    end try
                                end if
                            on error
                                -- Selection fallback failed
                            end try
                        end if

                        -- Check for compatibility mode indicators
                        try
                            if presentationName contains ".ppt" and not (presentationName contains ".pptx") then
                                set currentMode to "compatibility"
                            end if

                            if (count of document windows) > 0 then
                                set winTitle to name of document window 1
                                if winTitle contains "Compatibility Mode" then
                                    set currentMode to "compatibility"
                                end if
                            end if
                        on error
                            -- Mode detection failed, keep default
                        end try

                        -- If no slide was detected, default to slide 1 but mark as limited
                        if not slideDetected then
                            set currentSlideNum to 1
                            set currentMode to "limited"
                        end if

                        return presentationName & "|" & currentSlideNum & "|" & totalSlides & "|" & currentMode

                    else
                        return "no_presentation|||"
                    end if
                on error errMsg
                    return "error|" & errMsg & "||"
                end try
            end tell
            '''

            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=5  # Increased timeout for reliability
            )

            if result.returncode == 0:
                output = result.stdout.strip()
                if output:
                    parts = output.split("|")
                    if len(parts) >= 4:
                        # Handle error cases
                        if parts[0] == "no_presentation":
                            return {
                                'current_slide': None,
                                'total_slides': None,
                                'presentation_name': None,
                                'mode': 'no_presentation'
                            }
                        elif parts[0] == "error":
                            return {
                                'current_slide': None,
                                'total_slides': None,
                                'presentation_name': None,
                                'mode': 'error'
                            }

                        # Normal response processing
                        try:
                            current_slide = int(parts[1]) if parts[1] and parts[1].isdigit() else None
                            total_slides = int(parts[2]) if parts[2] and parts[2].isdigit() else None

                            return {
                                'current_slide': current_slide,
                                'total_slides': total_slides,
                                'presentation_name': parts[0] if parts[0] else None,
                                'mode': parts[3] if len(parts) > 3 else 'normal'
                            }
                        except (ValueError, IndexError):
                            # Parse error
                            pass

        except subprocess.TimeoutExpired:
            return {
                'current_slide': None,
                'total_slides': None,
                'presentation_name': None,
                'mode': 'timeout'
            }
        except Exception as e:
            return {
                'current_slide': None,
                'total_slides': None,
                'presentation_name': None,
                'mode': 'exception'
            }

        return {
            'current_slide': None,
            'total_slides': None,
            'presentation_name': None,
            'mode': 'unknown'
        }

    def _get_powerpoint_windows_windows(self) -> List[WindowInfo]:
        windows = []
        try:
            import win32gui
            import win32process

            def enum_window_callback(hwnd, windows_list):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    _, process_id = win32process.GetWindowThreadProcessId(hwnd)

                    try:
                        process = psutil.Process(process_id)
                        app_name = process.name()

                        if self.is_powerpoint_window(window_title, app_name):
                            rect = win32gui.GetWindowRect(hwnd)
                            position = (rect[0], rect[1])
                            size = (rect[2] - rect[0], rect[3] - rect[1])

                            windows_list.append(WindowInfo(
                                window_id=hwnd,
                                title=window_title,
                                app_name=app_name,
                                position=position,
                                size=size
                            ))
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                return True

            win32gui.EnumWindows(enum_window_callback, windows)

        except ImportError:
            print("Windows window detection requires pywin32")
            print("Install with: pip install pywin32")
        except Exception as e:
            print(f"Error detecting Windows windows: {e}")

        return windows

    def get_active_powerpoint_window(self) -> Optional[WindowInfo]:
        windows = self.get_powerpoint_windows()

        if not windows:
            return None

        if len(windows) == 1:
            return windows[0]

        for window in windows:
            if 'Slide Show' in window.title:
                return window

        largest_window = max(windows, key=lambda w: (w.size[0] * w.size[1]) if w.size else 0)
        return largest_window

    def monitor_powerpoint_window(self, callback=None, interval: float = 1.0):
        print(f"Starting PowerPoint window monitoring (interval: {interval}s)")
        print("Press Ctrl+C to stop monitoring")

        try:
            while True:
                current_window = self.get_active_powerpoint_window()

                if current_window:
                    slide_info = self.extract_slide_info_from_title(current_window.title)

                    if self.current_window is None or current_window.title != self.current_window.title:
                        print(f"\n🎯 PowerPoint window detected:")
                        print(f"   Title: {current_window.title}")
                        print(f"   App: {current_window.app_name}")

                        if slide_info['current_slide']:
                            print(f"   Current Slide: {slide_info['current_slide']}")
                            if slide_info['total_slides']:
                                print(f"   Total Slides: {slide_info['total_slides']}")
                            if slide_info['presentation_name']:
                                print(f"   Presentation: {slide_info['presentation_name']}")
                            print(f"   Mode: {slide_info['mode']}")

                        self.current_window = current_window
                        self.last_slide_info = slide_info

                        if callback:
                            callback(current_window, slide_info)

                    elif (self.last_slide_info and slide_info['current_slide'] and
                          slide_info['current_slide'] != self.last_slide_info.get('current_slide')):
                        print(f"\n📄 Slide changed: {self.last_slide_info.get('current_slide', '?')} → {slide_info['current_slide']}")
                        self.last_slide_info = slide_info

                        if callback:
                            callback(current_window, slide_info)

                elif self.current_window is not None:
                    print("\n❌ PowerPoint window lost")
                    self.current_window = None
                    self.last_slide_info = None

                time.sleep(interval)

        except KeyboardInterrupt:
            print("\n\nMonitoring stopped by user")
        except Exception as e:
            print(f"\nMonitoring error: {e}")

    def get_current_slide_info(self) -> Optional[Dict]:
        window = self.get_active_powerpoint_window()
        if window:
            return self.extract_slide_info_from_title(window.title)
        return None

def main():
    detector = PowerPointWindowDetector()

    print("PowerPoint Window Auto-Detection Demo")
    print("=" * 40)

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

        print("\nStarting continuous monitoring...")
        detector.monitor_powerpoint_window()
    else:
        print("No PowerPoint windows found.")
        print("\nTip: Open a PowerPoint presentation and try again.")

if __name__ == "__main__":
    main()