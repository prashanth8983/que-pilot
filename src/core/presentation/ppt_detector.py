"""
Specialized PowerPoint .ppt file detector for macOS.
This handles the specific quirks of older .ppt files in PowerPoint.
"""

import subprocess
from typing import Dict, Optional

class PPTDetector:
    """Specialized detector for .ppt files on macOS using AppleScript"""

    def __init__(self):
        self.last_known_slide = 1
        self.presentation_info = None

    def get_basic_info(self) -> Optional[Dict]:
        """Get basic presentation information"""
        applescript = '''
        tell application "Microsoft PowerPoint"
            try
                if (count of presentations) > 0 then
                    set currentPresentation to active presentation
                    set presentationName to name of currentPresentation
                    set totalSlides to count of slides of currentPresentation
                    set windowCount to count of document windows
                    set slideshowCount to count of slide show windows

                    return presentationName & "|" & totalSlides & "|" & windowCount & "|" & slideshowCount
                else
                    return "no_presentation|||"
                end if
            on error errMsg
                return "error|" & errMsg & "||"
            end try
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0 and result.stdout.strip():
                parts = result.stdout.strip().split("|")
                if len(parts) >= 4:
                    return {
                        'presentation_name': parts[0],
                        'total_slides': int(parts[1]) if parts[1].isdigit() else 0,
                        'document_windows': int(parts[2]) if parts[2].isdigit() else 0,
                        'slideshow_windows': int(parts[3]) if parts[3].isdigit() else 0,
                        'is_ppt': '.ppt' in parts[0] and '.pptx' not in parts[0]
                    }
        except Exception as e:
            print(f"Failed to get basic info: {e}")
        return None

    def detect_current_slide_simple(self) -> Optional[Dict]:
        """Simple slide detection for .ppt files"""
        basic_info = self.get_basic_info()
        if not basic_info:
            return None

        # For slideshow mode
        if basic_info['slideshow_windows'] > 0:
            slide_info = self._detect_slideshow_slide()
            if slide_info:
                slide_info.update(basic_info)
                return slide_info

        # For normal mode - try different approaches
        slide_info = self._detect_normal_slide()
        if slide_info:
            slide_info.update(basic_info)
            return slide_info

        # Fallback: assume slide 1 for .ppt files
        if basic_info['is_ppt']:
            print("Using fallback detection for .ppt file")
            return {
                'current_slide': self.last_known_slide,
                'slide_text': '',
                'detection_method': 'ppt_fallback',
                **basic_info
            }

        return None

    def _detect_slideshow_slide(self) -> Optional[Dict]:
        """Detect current slide in slideshow mode"""
        applescript = '''
        tell application "Microsoft PowerPoint"
            try
                if (count of slide show windows) > 0 then
                    set slideNum to slide number of slide of slide show view of slide show window 1
                    return "slideshow|" & slideNum
                else
                    return "slideshow|no_window"
                end if
            on error errMsg
                return "slideshow|error|" & errMsg
            end try
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                parts = result.stdout.strip().split("|")
                if len(parts) >= 2 and parts[0] == "slideshow" and parts[1].isdigit():
                    slide_num = int(parts[1])
                    self.last_known_slide = slide_num
                    return {
                        'current_slide': slide_num,
                        'detection_method': 'slideshow',
                        'slide_text': self._get_slide_text(slide_num)
                    }
        except Exception as e:
            print(f"Slideshow detection failed: {e}")
        return None

    def _detect_normal_slide(self) -> Optional[Dict]:
        """Detect current slide in normal editing mode"""
        # Method 1: Try to detect from window title if it contains slide info
        title_info = self._get_window_title_info()
        if title_info:
            return title_info

        # Method 2: Try selection-based detection
        selection_info = self._detect_from_selection()
        if selection_info:
            return selection_info

        # Method 3: Try view-based detection (simpler approach)
        view_info = self._detect_from_view()
        if view_info:
            return view_info

        return None

    def _get_window_title_info(self) -> Optional[Dict]:
        """Try to extract slide info from window title"""
        applescript = '''
        tell application "Microsoft PowerPoint"
            try
                if (count of document windows) > 0 then
                    set winTitle to name of document window 1
                    return "title|" & winTitle
                else
                    return "title|no_window"
                end if
            on error errMsg
                return "title|error|" & errMsg
            end try
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                parts = result.stdout.strip().split("|", 1)
                if len(parts) >= 2:
                    title = parts[1]
                    # Look for slide patterns in title
                    import re
                    slide_patterns = [
                        r'Slide (\d+)',
                        r'(\d+) of \d+',
                        r'(\d+)/\d+'
                    ]

                    for pattern in slide_patterns:
                        match = re.search(pattern, title, re.IGNORECASE)
                        if match:
                            slide_num = int(match.group(1))
                            self.last_known_slide = slide_num
                            return {
                                'current_slide': slide_num,
                                'detection_method': 'window_title',
                                'slide_text': self._get_slide_text(slide_num)
                            }
        except Exception as e:
            print(f"Window title detection failed: {e}")
        return None

    def _detect_from_selection(self) -> Optional[Dict]:
        """Try to detect slide from current selection"""
        applescript = '''
        tell application "Microsoft PowerPoint"
            try
                if (count of document windows) > 0 then
                    set selObj to selection of document window 1
                    if (count of slides of selObj) > 0 then
                        set selSlide to slide 1 of selObj
                        set slideNum to slide number of selSlide
                        return "selection|" & slideNum
                    else
                        return "selection|no_slides"
                    end if
                else
                    return "selection|no_window"
                end if
            on error errMsg
                return "selection|error|" & errMsg
            end try
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                parts = result.stdout.strip().split("|")
                if len(parts) >= 2 and parts[0] == "selection" and parts[1].isdigit():
                    slide_num = int(parts[1])
                    self.last_known_slide = slide_num
                    return {
                        'current_slide': slide_num,
                        'detection_method': 'selection',
                        'slide_text': self._get_slide_text(slide_num)
                    }
        except Exception as e:
            print(f"Selection detection failed: {e}")
        return None

    def _detect_from_view(self) -> Optional[Dict]:
        """Simple view-based detection"""
        applescript = '''
        tell application "Microsoft PowerPoint"
            try
                if (count of document windows) > 0 then
                    set docWin to document window 1

                    -- Try to get any slide reference we can
                    try
                        set viewObj to view of docWin
                        set currentSlide to slide of viewObj
                        if currentSlide is not missing value then
                            set slideNum to slide number of currentSlide
                            return "view|" & slideNum
                        else
                            return "view|no_slide_ref"
                        end if
                    on error
                        return "view|error_getting_slide"
                    end try
                else
                    return "view|no_window"
                end if
            on error errMsg
                return "view|error|" & errMsg
            end try
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', applescript],
                                  capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                parts = result.stdout.strip().split("|")
                if len(parts) >= 2 and parts[0] == "view" and parts[1].isdigit():
                    slide_num = int(parts[1])
                    self.last_known_slide = slide_num
                    return {
                        'current_slide': slide_num,
                        'detection_method': 'view',
                        'slide_text': self._get_slide_text(slide_num)
                    }
        except Exception as e:
            print(f"View detection failed: {e}")
        return None

    def _get_slide_text(self, slide_num: int) -> str:
        """Get text content from a specific slide with enhanced extraction"""
        # Create working AppleScript based on test results
        applescript = f'''tell application "Microsoft PowerPoint"
    try
        set currentPresentation to active presentation
        set targetSlide to slide {slide_num} of currentPresentation
        set slideText to ""
        set shapeCount to count of shapes of targetSlide

        repeat with i from 1 to shapeCount
            try
                set currentShape to shape i of targetSlide
                try
                    set shapeText to text of text frame of currentShape
                    if shapeText is not "" then
                        set slideText to slideText & shapeText & " | "
                    end if
                on error
                    -- This shape doesn't have text
                end try
            on error
                -- Skip this shape entirely
            end try
        end repeat

        if slideText is not "" then
            return slideText
        else
            return "[No text found in " & shapeCount & " shapes on slide " & {slide_num} & "]"
        end if
    on error errMsg
        return "[AppleScript error: " & errMsg & "]"
    end try
end tell'''

        try:
            # Use heredoc approach to avoid AppleScript parsing issues
            result = subprocess.run(['osascript'], input=applescript,
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                text = result.stdout.strip()
                # Log the extraction attempt
                print(f"🔍 Text extraction for slide {slide_num}: {len(text)} characters")
                if text and not text.startswith('['):
                    print(f"📝 Content preview: {text[:150]}{'...' if len(text) > 150 else ''}")
                elif text.startswith('['):
                    print(f"ℹ️  Status: {text}")
                return text
            else:
                error_msg = f"[AppleScript execution failed: {result.stderr.strip()}]"
                print(f"❌ Text extraction error: {error_msg}")
                return error_msg
        except Exception as e:
            error_msg = f"[Python exception: {str(e)}]"
            print(f"❌ Text extraction exception: {error_msg}")
            return error_msg

    def monitor_slide_changes(self, callback, interval=2.0):
        """Monitor for slide changes"""
        import time
        import threading

        def monitor_loop():
            last_slide = None
            while True:
                try:
                    current_info = self.detect_current_slide_simple()
                    if current_info and current_info.get('current_slide') != last_slide:
                        last_slide = current_info.get('current_slide')
                        if callback:
                            callback(current_info)
                except Exception as e:
                    print(f"Monitoring error: {e}")
                time.sleep(interval)

        thread = threading.Thread(target=monitor_loop, daemon=True)
        thread.start()
        return thread

# Test function
if __name__ == "__main__":
    detector = PPTDetector()

    print("Testing PPT Detector...")
    info = detector.detect_current_slide_simple()

    if info:
        print(f"✓ Detection successful:")
        print(f"  Presentation: {info.get('presentation_name')}")
        print(f"  Current Slide: {info.get('current_slide')}")
        print(f"  Total Slides: {info.get('total_slides')}")
        print(f"  Method: {info.get('detection_method')}")
        print(f"  Is PPT: {info.get('is_ppt')}")

        text = info.get('slide_text', '')
        if text:
            print(f"  Text: {text[:100]}...")
    else:
        print("✗ Detection failed")