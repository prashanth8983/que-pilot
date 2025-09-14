import platform
import time
import re
from typing import Optional, List, Dict, Tuple

# Import psutil with fallback
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    print("psutil not available. Windows process detection may be limited.")
    psutil = None
    PSUTIL_AVAILABLE = False

# Import the new screen detector
try:
    from .screen_detector import PowerPointScreenDetector
    SCREEN_DETECTOR_AVAILABLE = True
except ImportError as e:
    print(f"Screen detector not available: {e}")
    print("Install OCR dependencies with: pip install -r requirements-ocr.txt")
    SCREEN_DETECTOR_AVAILABLE = False
    PowerPointScreenDetector = None

# Import the specialized PPT detector
try:
    from .ppt_detector import PPTDetector
    PPT_DETECTOR_AVAILABLE = True
except ImportError as e:
    print(f"PPT detector not available: {e}")
    PPT_DETECTOR_AVAILABLE = False
    PPTDetector = None

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

        # Initialize screen detector if available
        self.screen_detector = None
        if SCREEN_DETECTOR_AVAILABLE:
            try:
                self.screen_detector = PowerPointScreenDetector()
            except Exception as e:
                print(f"Failed to initialize screen detector: {e}")
                self.screen_detector = None

        # Initialize PPT detector if available
        self.ppt_detector = None
        if PPT_DETECTOR_AVAILABLE:
            try:
                self.ppt_detector = PPTDetector()
            except Exception as e:
                print(f"Failed to initialize PPT detector: {e}")
                self.ppt_detector = None

    def is_powerpoint_window(self, window_title: str, app_name: str) -> bool:
        powerpoint_indicators = [
            "PowerPoint",
            "Microsoft PowerPoint",
            ".ppt",
            ".pptx",
            "Slide",
            "Presentation",
            "Slide Show"
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
            try:
                applescript_info = self.get_powerpoint_slide_info_macos()
                if applescript_info and applescript_info.get('current_slide'):
                    return applescript_info
            except Exception as e:
                print(f"AppleScript slide info failed: {e}")

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

        if not slide_info['current_slide'] and 'Slide Show' in title:
            slide_info['mode'] = 'slideshow'
            # Try to extract presentation name from titles like "Slide Show - My Presentation.pptx"
            if ' - ' in title:
                parts = title.split(' - ')
                if len(parts) > 1:
                    # Find the part that is likely the presentation name
                    for part in parts:
                        if '.ppt' in part.lower():
                            slide_info['presentation_name'] = part.strip()
                            break
                    if not slide_info['presentation_name']:
                         slide_info['presentation_name'] = parts[1].strip()

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
            from Cocoa import NSWorkspace
            from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

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
            presentation_title = self._get_powerpoint_presentation_title_macos()

            for window_info in window_list:
                window_title = window_info.get('kCGWindowName', '')
                owner_name = window_info.get('kCGWindowOwnerName', '')
                window_id = window_info.get('kCGWindowNumber')
                layer = window_info.get('kCGWindowLayer', 0)
                bounds = window_info.get('kCGWindowBounds', {})
                
                is_ppt_window = self.is_powerpoint_window(window_title, owner_name) or \
                                (owner_name == 'Microsoft PowerPoint' and bounds.get('Width', 0) > 400 and layer == 0)

                if is_ppt_window:
                    if not window_title and presentation_title:
                        window_title = presentation_title
                    
                    position = (bounds.get('X', 0), bounds.get('Y', 0))
                    size = (bounds.get('Width', 0), bounds.get('Height', 0))

                    windows.append(WindowInfo(
                        window_id=window_id, title=window_title, app_name=owner_name,
                        position=position, size=size
                    ))
        except ImportError:
            print("macOS window detection requires pyobjc-framework-Cocoa and pyobjc-framework-Quartz")
        except Exception as e:
            print(f"Error detecting macOS windows: {e}")
        return windows

    def _get_powerpoint_presentation_title_macos(self) -> str:
        try:
            slide_info = self.get_powerpoint_slide_info_macos()
            if slide_info:
                name = slide_info.get('presentation_name', '')
                current = slide_info.get('current_slide', 1)
                total = slide_info.get('total_slides', 0)
                mode = slide_info.get('mode', 'normal')

                if mode == 'slideshow':
                    return f"Slide Show - {name} - Slide {current} of {total}"
                else:
                    return f"{name} - Slide {current} of {total}"
        except Exception:
            pass
        return ""

    def get_powerpoint_slide_info_macos(self) -> Dict:
        try:
            import subprocess
            # Enhanced AppleScript with special handling for .ppt files
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
                        set slideText to ""
                        set isPptFile to false

                        -- Check if this is a .ppt file (compatibility mode)
                        try
                            if presentationName contains ".ppt" and not (presentationName contains ".pptx") then
                                set isPptFile to true
                                set currentMode to "compatibility"
                            end if
                        end try

                        -- Method 1: Try slideshow mode first
                        try
                            if (count of slide show windows) > 0 then
                                set currentSlideNum to slide number of slide of slide show view of slide show window 1
                                set currentMode to "slideshow"
                                set slideDetected to true
                                -- Get slide text in slideshow mode
                                try
                                    set currentSlide to slide currentSlideNum of currentPresentation
                                    repeat with shp in shapes of currentSlide
                                        if has text frame shp then
                                            set slideText to slideText & text of text frame of shp & " "
                                        end if
                                    end repeat
                                end try
                            end if
                        end try

                        -- Method 2: Enhanced document window approach for .ppt files
                        if not slideDetected then
                            try
                                if (count of document windows) > 0 then
                                    set docWin to document window 1

                                    -- For .ppt files, try different approaches
                                    if isPptFile then
                                        -- .ppt files may have different view object behavior
                                        try
                                            set viewType to view type of view of docWin
                                            -- Try to get slide from current view
                                            if viewType is not missing value then
                                                set viewObj to view of docWin
                                                try
                                                    -- Try getting slide index differently for .ppt
                                                    set currentSlideRef to slide of viewObj
                                                    if currentSlideRef is not missing value then
                                                        set currentSlideNum to slide number of currentSlideRef
                                                        set slideDetected to true
                                                    end if
                                                on error
                                                    -- Fallback: try slide index property
                                                    try
                                                        set slideIdx to slide index of viewObj
                                                        if slideIdx is not missing value and slideIdx > 0 then
                                                            set currentSlideNum to slideIdx
                                                            set slideDetected to true
                                                        end if
                                                    end try
                                                end try
                                            end if
                                        end try
                                    else
                                        -- Standard approach for .pptx files
                                        set viewObj to view of docWin
                                        try
                                            set slideIdx to slide index of viewObj
                                            if slideIdx > 0 and slideIdx <= totalSlides then
                                                set currentSlideNum to slideIdx
                                                set slideDetected to true
                                            end if
                                        on error
                                            try
                                                set slideRef to slide of viewObj
                                                if slideRef is not missing value then
                                                    set currentSlideNum to slide number of slideRef
                                                    set slideDetected to true
                                                end if
                                            end try
                                        end try
                                    end if

                                    -- Get slide text if we detected a slide
                                    if slideDetected then
                                        try
                                            set currentSlide to slide currentSlideNum of currentPresentation
                                            repeat with shp in shapes of currentSlide
                                                if has text frame shp then
                                                    set slideText to slideText & text of text frame of shp & " "
                                                end if
                                            end repeat
                                        end try
                                    end if
                                end if
                            end try
                        end if

                        -- Method 3: Selection-based detection (works better for .ppt sometimes)
                        if not slideDetected then
                            try
                                if (count of document windows) > 0 then
                                    set selectionObj to selection of document window 1
                                    if (count of slides of selectionObj) > 0 then
                                        set selectedSlide to slide 1 of selectionObj
                                        set currentSlideNum to slide number of selectedSlide
                                        set slideDetected to true

                                        -- Get slide text from selection
                                        try
                                            repeat with shp in shapes of selectedSlide
                                                if has text frame shp then
                                                    set slideText to slideText & text of text frame of shp & " "
                                                end if
                                            end repeat
                                        end try
                                    end if
                                end if
                            end try
                        end if

                        -- Method 4: Fallback for .ppt files - use first slide if nothing else works
                        if not slideDetected and isPptFile then
                            try
                                set currentSlideNum to 1
                                set slideDetected to true
                                set currentMode to "ppt_fallback"
                                -- Try to get text from first slide
                                try
                                    set currentSlide to slide 1 of currentPresentation
                                    repeat with shp in shapes of currentSlide
                                        if has text frame shp then
                                            set slideText to slideText & text of text frame of shp & " "
                                        end if
                                    end repeat
                                end try
                            end try
                        end if

                        -- Clean up slide text
                        if slideText is not "" then
                            set slideText to (characters 1 through (length of slideText) of slideText) as string
                        end if

                        -- Set final mode
                        if not slideDetected then
                            set currentMode to "limited"
                        else if isPptFile then
                            set currentMode to "ppt_compatibility"
                        end if

                        return presentationName & "|" & currentSlideNum & "|" & totalSlides & "|" & currentMode & "|" & slideText
                    else
                        return "no_presentation||||"
                    end if
                on error errMsg
                    return "error|" & errMsg & "|||"
                end try
            end tell
            '''
            result = subprocess.run(['osascript', '-e', applescript], capture_output=True, text=True, timeout=8)
            if result.returncode == 0 and result.stdout.strip():
                parts = result.stdout.strip().split("|")
                if len(parts) >= 4:
                    if parts[0] in ["no_presentation", "error"]:
                        print(f"PowerPoint detection: {parts[0]} - {parts[1] if len(parts) > 1 else 'No details'}")
                        return {'mode': parts[0]}

                    slide_text = parts[4] if len(parts) > 4 else ""
                    current_slide = int(parts[1]) if parts[1].isdigit() else None

                    # Log current slide and text to console
                    if current_slide and slide_text.strip():
                        print(f"\n=== SLIDE {current_slide} CONTENT ===")
                        print(f"Text: {slide_text.strip()[:200]}{'...' if len(slide_text.strip()) > 200 else ''}")
                        print("=" * 30)
                    elif current_slide:
                        print(f"\n=== SLIDE {current_slide} ===")
                        print("No text content detected")
                        print("=" * 30)

                    return {
                        'current_slide': current_slide,
                        'total_slides': int(parts[2]) if parts[2].isdigit() else None,
                        'presentation_name': parts[0] or None,
                        'mode': parts[3] or 'normal',
                        'slide_text': slide_text.strip()
                    }
        except (subprocess.TimeoutExpired, Exception) as e:
            print(f"PowerPoint detection error: {e}")
            return {'mode': 'exception'}
        return {'mode': 'unknown'}

    def _get_powerpoint_windows_windows(self) -> List[WindowInfo]:
        windows = []
        try:
            import win32gui
            import win32process

            def enum_window_callback(hwnd, windows_list):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        if PSUTIL_AVAILABLE:
                            app_name = psutil.Process(pid).name()
                        else:
                            app_name = "unknown"

                        if self.is_powerpoint_window(title, app_name):
                            rect = win32gui.GetWindowRect(hwnd)
                            pos = (rect[0], rect[1])
                            size = (rect[2] - rect[0], rect[3] - rect[1])
                            windows_list.append(WindowInfo(hwnd, title, app_name, pos, size))
                    except Exception as e:
                        if PSUTIL_AVAILABLE:
                            try:
                                # Try without psutil if it fails
                                if self.is_powerpoint_window(title, "PowerPoint"):
                                    rect = win32gui.GetWindowRect(hwnd)
                                    pos = (rect[0], rect[1])
                                    size = (rect[2] - rect[0], rect[3] - rect[1])
                                    windows_list.append(WindowInfo(hwnd, title, "PowerPoint", pos, size))
                            except:
                                pass
                return True
            win32gui.EnumWindows(enum_window_callback, windows)
        except ImportError:
            print("Windows detection requires: pip install pywin32")
        except Exception as e:
            print(f"Error detecting Windows windows: {e}")
        return windows

    def get_active_powerpoint_window(self) -> Optional[WindowInfo]:
        windows = self.get_powerpoint_windows()
        if not windows: return None
        if len(windows) == 1: return windows[0]
        for w in windows:
            if 'Slide Show' in w.title: return w
        return max(windows, key=lambda w: (w.size[0] * w.size[1]) if w.size else 0)

    def monitor_powerpoint_window(self, callback=None, interval: float = 1.0):
        print(f"Starting PowerPoint monitoring (interval: {interval}s). Press Ctrl+C to stop.")
        try:
            while True:
                window = self.get_active_powerpoint_window()
                if window:
                    slide_info = self.extract_slide_info_from_title(window.title)
                    if not self.current_window or window.title != self.current_window.title:
                        print(f"\nDetected: {window.title} ({slide_info.get('current_slide', '?')}/{slide_info.get('total_slides', '?')})")
                        self.current_window, self.last_slide_info = window, slide_info
                        if callback: callback(window, slide_info)
                    elif self.last_slide_info and slide_info.get('current_slide') != self.last_slide_info.get('current_slide'):
                        print(f"\n📄 Slide changed: {self.last_slide_info.get('current_slide', '?')} → {slide_info.get('current_slide')}")
                        self.last_slide_info = slide_info
                        if callback: callback(window, slide_info)
                elif self.current_window:
                    print("\nPowerPoint window lost.")
                    self.current_window, self.last_slide_info = None, None
                time.sleep(interval)
        except KeyboardInterrupt:
            print("\nMonitoring stopped.")
        except Exception as e:
            print(f"\nMonitoring error: {e}")

    def get_current_slide_info(self) -> Optional[Dict]:
        # On macOS, try PPT detector first if available (it doesn't need window detection)
        if self.system == "Darwin" and self.ppt_detector:
            try:
                ppt_info = self.ppt_detector.detect_current_slide_simple()
                if ppt_info and ppt_info.get('current_slide') and ppt_info.get('is_ppt'):
                    print(f"\n=== PPT DETECTION ===")
                    print(f"📄 Presentation: {ppt_info.get('presentation_name', 'Unknown')}")
                    print(f"📊 Slide: {ppt_info.get('current_slide')}/{ppt_info.get('total_slides')}")
                    print(f"🔧 Method: {ppt_info.get('detection_method')}")

                    slide_text = ppt_info.get('slide_text', '')
                    if slide_text:
                        if slide_text.startswith('[') and slide_text.endswith(']'):
                            print(f"⚠️  {slide_text}")
                        else:
                            print(f"📝 Content: {slide_text}")
                    else:
                        print(f"📝 Content: [No text extracted]")
                    print("=" * 50)

                    # Convert to standard format
                    return {
                        'current_slide': ppt_info.get('current_slide'),
                        'total_slides': ppt_info.get('total_slides'),
                        'presentation_name': ppt_info.get('presentation_name'),
                        'mode': 'ppt_specialized',
                        'slide_text': ppt_info.get('slide_text', ''),
                        'detection_method': ppt_info.get('detection_method', 'ppt_specialized')
                    }
            except Exception as e:
                print(f"PPT specialized detection failed: {e}")

        window = self.get_active_powerpoint_window()
        if not window:
            # If no window but we're on macOS, still try regular AppleScript
            if self.system == "Darwin":
                return self.extract_slide_info_from_title("")
            return None

        # First try the existing method (AppleScript on macOS, title parsing on Windows)
        slide_info = self.extract_slide_info_from_title(window.title)

        # If we couldn't get slide info and screen detector is available, use OCR fallback
        if (not slide_info.get('current_slide') or not slide_info.get('slide_text', '').strip()) and self.screen_detector:
            try:
                screen_info = self.screen_detector.get_current_slide_info()
                if screen_info and screen_info.slide_number:
                    # Merge the information
                    slide_info['current_slide'] = screen_info.slide_number
                    slide_info['slide_text'] = screen_info.content
                    slide_info['slide_title'] = screen_info.title
                    slide_info['detection_method'] = 'screen_ocr'
                    slide_info['ocr_confidence'] = screen_info.confidence_score

                    print(f"\n=== OCR DETECTION ===")
                    print(f"Slide {screen_info.slide_number}: {screen_info.title}")
                    print(f"Confidence: {screen_info.confidence_score:.1f}%")
                    print(f"Content: {screen_info.content[:100]}...")
                    print("=" * 30)
            except Exception as e:
                print(f"Screen detection fallback failed: {e}")

        return slide_info

    def get_current_slide_info_with_content(self) -> Optional[Dict]:
        """Get slide info with enhanced content detection using OCR when available"""
        return self.get_current_slide_info()

if __name__ == "__main__":
    detector = PowerPointWindowDetector()
    detector.monitor_powerpoint_window()
