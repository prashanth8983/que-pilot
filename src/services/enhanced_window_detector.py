"""
Enhanced PowerPoint Window Detector

Improved version of the window detector with better error handling and slide image support.
"""

import platform
import time
import re
import subprocess
import logging
from typing import Optional, List, Dict, Tuple
from dataclasses import dataclass

@dataclass
class SlideInfo:
    """Information about the current slide."""
    current_slide: int
    total_slides: int
    presentation_name: str
    mode: str
    title: str = ""
    has_content: bool = False

@dataclass
class WindowInfo:
    """Information about a PowerPoint window."""
    window_id: str
    title: str
    app_name: str
    position: Tuple[int, int] = (0, 0)
    size: Tuple[int, int] = (0, 0)
    is_slideshow: bool = False

class EnhancedPowerPointDetector:
    """Enhanced PowerPoint window detection with better slide tracking."""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.system = platform.system()
        self.last_slide_info = None

    def get_current_slide_info(self) -> Optional[SlideInfo]:
        """Get comprehensive information about the current slide."""
        if self.system == "Darwin":
            return self._get_slide_info_macos()
        elif self.system == "Windows":
            return self._get_slide_info_windows()
        else:
            self.logger.warning(f"Unsupported system: {self.system}")
            return None

    def _get_slide_info_macos(self) -> Optional[SlideInfo]:
        """Get slide info on macOS using AppleScript."""
        try:
            # Enhanced AppleScript with better error handling
            applescript = '''
            tell application "Microsoft PowerPoint"
                try
                    if (count of presentations) > 0 then
                        set currentPresentation to active presentation
                        set presentationName to name of currentPresentation
                        set totalSlides to count of slides of currentPresentation
                        set currentSlideNum to 1
                        set currentMode to "normal"
                        set slideTitle to ""
                        set hasContent to false

                        -- Try multiple methods to get current slide number
                        try
                            -- Method 1: Slideshow window
                            if (count of slide show windows) > 0 then
                                set slideShowWin to slide show window 1
                                set slideShowView to slide show view of slideShowWin
                                set currentSlideNum to slide number of slide of slideShowView
                                set currentMode to "slideshow"
                            else
                                -- Method 2: Document window
                                if (count of document windows) > 0 then
                                    set docWin to document window 1
                                    set viewObj to view of docWin

                                    -- Try to get slide index from view
                                    try
                                        set slideIdx to slide index of viewObj
                                        if slideIdx > 0 and slideIdx <= totalSlides then
                                            set currentSlideNum to slideIdx
                                        end if
                                    end try

                                    -- Try to get current slide from selection
                                    try
                                        set selectionObj to selection of docWin
                                        if (count of slides of selectionObj) > 0 then
                                            set selectedSlide to slide 1 of selectionObj
                                            set currentSlideNum to slide number of selectedSlide
                                        end if
                                    end try
                                end if
                            end if
                        end try

                        -- Try to get slide title and content info
                        try
                            if currentSlideNum > 0 and currentSlideNum <= totalSlides then
                                set currentSlide to slide currentSlideNum of currentPresentation
                                set shapeCount to count of shapes of currentSlide

                                if shapeCount > 0 then
                                    set hasContent to true
                                    -- Try to get title from first text shape
                                    repeat with i from 1 to shapeCount
                                        set currentShape to shape i of currentSlide
                                        if has text frame of currentShape then
                                            set textFrame to text frame of currentShape
                                            if (count of characters of (text of textFrame)) > 0 then
                                                set slideTitle to text of textFrame
                                                exit repeat
                                            end if
                                        end if
                                    end repeat
                                end if
                            end if
                        end try

                        return presentationName & "|" & currentSlideNum & "|" & totalSlides & "|" & currentMode & "|" & slideTitle & "|" & hasContent
                    else
                        return "no_presentation|0|0|none||false"
                    end if
                on error errMsg
                    return "error|0|0|error|" & errMsg & "|false"
                end try
            end tell
            '''

            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=10)

            if result.returncode == 0 and result.stdout.strip():
                parts = result.stdout.strip().split("|")
                if len(parts) >= 6:
                    if parts[0] == "no_presentation":
                        self.logger.info("No PowerPoint presentation detected")
                        return None
                    elif parts[0] == "error":
                        self.logger.warning(f"PowerPoint AppleScript error: {parts[4] if len(parts) > 4 else 'Unknown'}")
                        return None

                    return SlideInfo(
                        current_slide=int(parts[1]) if parts[1].isdigit() else 1,
                        total_slides=int(parts[2]) if parts[2].isdigit() else 1,
                        presentation_name=parts[0],
                        mode=parts[3],
                        title=parts[4] if len(parts) > 4 else "",
                        has_content=parts[5].lower() == "true" if len(parts) > 5 else False
                    )

        except subprocess.TimeoutExpired:
            self.logger.warning("PowerPoint detection timed out")
        except Exception as e:
            self.logger.error(f"PowerPoint detection failed: {e}")

        return None

    def _get_slide_info_windows(self) -> Optional[SlideInfo]:
        """Get slide info on Windows using COM interface."""
        try:
            import win32com.client

            # Connect to PowerPoint
            try:
                ppt = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                self.logger.info("PowerPoint not running or not accessible")
                return None

            if ppt.Presentations.Count == 0:
                self.logger.info("No presentations open")
                return None

            # Get active presentation
            presentation = ppt.ActivePresentation
            presentation_name = presentation.Name
            total_slides = presentation.Slides.Count

            # Get current slide
            try:
                if ppt.SlideShowWindows.Count > 0:
                    # Slideshow mode
                    slideshow = ppt.SlideShowWindows(1)
                    current_slide = slideshow.View.CurrentShowPosition
                    mode = "slideshow"
                else:
                    # Normal mode - try to get from active window
                    if ppt.ActiveWindow:
                        try:
                            current_slide = ppt.ActiveWindow.Selection.SlideRange(1).SlideNumber
                            mode = "normal"
                        except:
                            current_slide = 1
                            mode = "normal"
                    else:
                        current_slide = 1
                        mode = "normal"
            except:
                current_slide = 1
                mode = "normal"

            # Try to get slide title
            slide_title = ""
            has_content = False
            try:
                if 1 <= current_slide <= total_slides:
                    slide = presentation.Slides(current_slide)
                    has_content = slide.Shapes.Count > 0

                    # Try to get title from first text shape
                    for shape in slide.Shapes:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            slide_title = shape.TextFrame.TextRange.Text.strip()
                            break
            except:
                pass

            return SlideInfo(
                current_slide=current_slide,
                total_slides=total_slides,
                presentation_name=presentation_name,
                mode=mode,
                title=slide_title,
                has_content=has_content
            )

        except ImportError:
            self.logger.error("pywin32 not available for Windows PowerPoint detection")
        except Exception as e:
            self.logger.error(f"Windows PowerPoint detection failed: {e}")

        return None

    def get_powerpoint_windows(self) -> List[WindowInfo]:
        """Get information about all PowerPoint windows."""
        windows = []

        if self.system == "Darwin":
            windows = self._get_windows_macos()
        elif self.system == "Windows":
            windows = self._get_windows_windows()

        return windows

    def _get_windows_macos(self) -> List[WindowInfo]:
        """Get PowerPoint windows on macOS."""
        windows = []
        try:
            from Cocoa import NSWorkspace
            from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

            # Get PowerPoint windows
            window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)

            for window_info in window_list:
                owner_name = window_info.get('kCGWindowOwnerName', '')
                window_title = window_info.get('kCGWindowName', '')
                window_id = window_info.get('kCGWindowNumber', 0)
                bounds = window_info.get('kCGWindowBounds', {})

                if 'powerpoint' in owner_name.lower() or 'Microsoft PowerPoint' in owner_name:
                    is_slideshow = 'slide show' in window_title.lower()

                    windows.append(WindowInfo(
                        window_id=str(window_id),
                        title=window_title,
                        app_name=owner_name,
                        position=(bounds.get('X', 0), bounds.get('Y', 0)),
                        size=(bounds.get('Width', 0), bounds.get('Height', 0)),
                        is_slideshow=is_slideshow
                    ))

        except ImportError:
            self.logger.error("PyObjC not available for macOS window detection")
        except Exception as e:
            self.logger.error(f"macOS window detection failed: {e}")

        return windows

    def _get_windows_windows(self) -> List[WindowInfo]:
        """Get PowerPoint windows on Windows."""
        windows = []
        try:
            import win32gui
            import win32process
            import psutil

            def enum_window_callback(hwnd, windows_list):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)

                    try:
                        process = psutil.Process(pid)
                        app_name = process.name()

                        if 'powerpnt' in app_name.lower() or 'powerpoint' in window_title.lower():
                            rect = win32gui.GetWindowRect(hwnd)
                            is_slideshow = 'slide show' in window_title.lower()

                            windows_list.append(WindowInfo(
                                window_id=str(hwnd),
                                title=window_title,
                                app_name=app_name,
                                position=(rect[0], rect[1]),
                                size=(rect[2] - rect[0], rect[3] - rect[1]),
                                is_slideshow=is_slideshow
                            ))
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass

                return True

            win32gui.EnumWindows(enum_window_callback, windows)

        except ImportError:
            self.logger.error("pywin32 not available for Windows window detection")
        except Exception as e:
            self.logger.error(f"Windows window detection failed: {e}")

        return windows

    def monitor_slide_changes(self, callback=None, interval: float = 2.0):
        """Monitor for slide changes and call callback when detected."""
        self.logger.info(f"Starting slide monitoring (interval: {interval}s)")

        try:
            while True:
                current_info = self.get_current_slide_info()

                if current_info:
                    # Check if slide changed
                    if (not self.last_slide_info or
                        current_info.current_slide != self.last_slide_info.current_slide or
                        current_info.presentation_name != self.last_slide_info.presentation_name):

                        self.logger.info(f"Slide change detected: {current_info.presentation_name} "
                                       f"slide {current_info.current_slide}/{current_info.total_slides}")

                        if callback:
                            callback(current_info)

                        self.last_slide_info = current_info

                time.sleep(interval)

        except KeyboardInterrupt:
            self.logger.info("Slide monitoring stopped")
        except Exception as e:
            self.logger.error(f"Slide monitoring error: {e}")

    def is_powerpoint_available(self) -> bool:
        """Check if PowerPoint is available and running."""
        slide_info = self.get_current_slide_info()
        return slide_info is not None

    def can_generate_slide_images(self) -> bool:
        """Check if slide image generation is possible."""
        # For now, return False to use window-based tracking
        # This avoids the "Cannot generate slide image" error
        return False

# Global instance
enhanced_detector = EnhancedPowerPointDetector()