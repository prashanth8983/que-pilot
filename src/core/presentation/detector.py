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
            # Enhanced AppleScript with better slide detection
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

                        -- Method 2: Try document window view
                        if not slideDetected then
                            try
                                if (count of document windows) > 0 then
                                    set docWin to document window 1
                                    set viewObj to view of docWin

                                    -- Try multiple approaches to get current slide
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

                                    -- Get slide text in normal mode
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

                        -- Method 3: Try selection-based detection
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

                        -- Clean up slide text
                        set slideText to (characters 1 through (length of slideText) of slideText) as string

                        -- Determine compatibility mode
                        try
                            if presentationName contains ".ppt" and not (presentationName contains ".pptx") then
                                set currentMode to "compatibility"
                            end if
                            if (count of document windows) > 0 and (name of document window 1 contains "Compatibility Mode") then
                                set currentMode to "compatibility"
                            end if
                        end try

                        if not slideDetected then
                            set currentMode to "limited"
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
                        app_name = psutil.Process(pid).name()
                        if self.is_powerpoint_window(title, app_name):
                            rect = win32gui.GetWindowRect(hwnd)
                            pos = (rect[0], rect[1])
                            size = (rect[2] - rect[0], rect[3] - rect[1])
                            windows_list.append(WindowInfo(hwnd, title, app_name, pos, size))
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
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
        window = self.get_active_powerpoint_window()
        return self.extract_slide_info_from_title(window.title) if window else None

if __name__ == "__main__":
    detector = PowerPointWindowDetector()
    detector.monitor_powerpoint_window()
