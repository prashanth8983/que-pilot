import platform
import time
import re
from typing import Optional, Dict, List, Tuple, Callable
import threading
import hashlib
from dataclasses import dataclass

# Core imports
import pyautogui
import cv2
import numpy as np
from PIL import Image

# Platform-specific imports will be loaded dynamically
pytesseract = None
win32gui = None
win32process = None
psutil = None

@dataclass
class SlideInfo:
    slide_number: Optional[int]
    title: Optional[str]
    content: str
    words: List[Dict]
    timestamp: float
    window_title: str
    confidence_score: float

class PowerPointScreenDetector:
    """
    Windows API-based PowerPoint slide detection using screen capture and OCR.
    Complements existing AppleScript and Win32 window detection methods.
    """

    def __init__(self):
        self.system = platform.system()
        self.ppt_windows = []
        self.current_slide_info = None
        self.previous_content_hash = ""
        self.monitoring_thread = None
        self.is_monitoring = False

        # OCR configuration
        self.ocr_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,!?;:()[]{}"-\' '

        # Import platform-specific modules
        self._initialize_platform_modules()

    def _initialize_platform_modules(self):
        """Initialize platform-specific modules"""
        global pytesseract, win32gui, win32process, psutil

        try:
            import pytesseract as pt
            pytesseract = pt
        except ImportError:
            print("Warning: pytesseract not available. Install with: pip install pytesseract")
            print("Also ensure Tesseract OCR is installed on your system")

        if self.system == "Windows":
            try:
                import win32gui as w32gui
                import win32process as w32process
                import psutil as ps
                win32gui = w32gui
                win32process = w32process
                psutil = ps
            except ImportError:
                print("Warning: Windows modules not available. Install with: pip install pywin32 psutil")

    def find_powerpoint_windows(self) -> List[Dict]:
        """Find all PowerPoint windows using platform-specific methods"""
        if self.system == "Windows":
            return self._find_windows_powerpoint()
        elif self.system == "Darwin":
            return self._find_macos_powerpoint()
        else:
            print(f"Platform {self.system} not supported for screen detection")
            return []

    def _find_windows_powerpoint(self) -> List[Dict]:
        """Find PowerPoint windows on Windows"""
        if not (win32gui and win32process and psutil):
            return []

        windows = []

        def enum_windows_callback(hwnd, windows_list):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if self._is_powerpoint_window(window_text):
                    try:
                        rect = win32gui.GetWindowRect(hwnd)
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        process_name = psutil.Process(pid).name()

                        windows_list.append({
                            'hwnd': hwnd,
                            'title': window_text,
                            'rect': rect,
                            'process': process_name,
                            'area': (rect[2] - rect[0]) * (rect[3] - rect[1])
                        })
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            return True

        win32gui.EnumWindows(enum_windows_callback, windows)
        self.ppt_windows = windows
        return windows

    def _find_macos_powerpoint(self) -> List[Dict]:
        """Find PowerPoint windows on macOS using Quartz"""
        try:
            from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

            windows = []
            window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)

            for window_info in window_list:
                window_title = window_info.get('kCGWindowName', '')
                owner_name = window_info.get('kCGWindowOwnerName', '')

                if self._is_powerpoint_window(window_title) or owner_name == 'Microsoft PowerPoint':
                    bounds = window_info.get('kCGWindowBounds', {})
                    if bounds.get('Width', 0) > 400:  # Filter out small windows
                        windows.append({
                            'title': window_title,
                            'rect': (bounds.get('X', 0), bounds.get('Y', 0),
                                   bounds.get('X', 0) + bounds.get('Width', 0),
                                   bounds.get('Y', 0) + bounds.get('Height', 0)),
                            'process': owner_name,
                            'area': bounds.get('Width', 0) * bounds.get('Height', 0)
                        })

            self.ppt_windows = windows
            return windows
        except ImportError:
            print("macOS window detection requires pyobjc-framework-Quartz")
            return []

    def _is_powerpoint_window(self, window_title: str) -> bool:
        """Check if a window title indicates PowerPoint"""
        if not window_title:
            return False

        powerpoint_indicators = [
            'PowerPoint', 'Slide Show', '.ppt', '.pptx',
            'Microsoft PowerPoint', 'Presentation'
        ]

        title_lower = window_title.lower()
        return any(indicator.lower() in title_lower for indicator in powerpoint_indicators)

    def capture_slide_area(self, window_info: Dict) -> Tuple[np.ndarray, Image.Image]:
        """Capture the slide area from PowerPoint window"""
        left, top, right, bottom = window_info['rect']

        # Add some padding to ensure we capture the full content
        padding = 5
        region = (max(0, left + padding),
                 max(0, top + padding),
                 max(1, right - left - 2*padding),
                 max(1, bottom - top - 2*padding))

        try:
            # Capture screenshot of the window area
            screenshot = pyautogui.screenshot(region=region)

            # Convert to OpenCV format
            opencv_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            return opencv_image, screenshot
        except Exception as e:
            print(f"Failed to capture screen area: {e}")
            # Return a small black image as fallback
            fallback_img = np.zeros((100, 100, 3), dtype=np.uint8)
            fallback_pil = Image.fromarray(fallback_img)
            return fallback_img, fallback_pil

    def extract_slide_content(self, image: np.ndarray) -> Dict:
        """Extract text from slide using OCR"""
        if pytesseract is None:
            return {
                'text': 'OCR not available',
                'words': [],
                'lines': [],
                'confidence': 0.0
            }

        try:
            # Convert to grayscale for better OCR
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

            # Apply image processing to improve OCR
            # Adaptive thresholding works better than simple binary threshold
            processed = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )

            # Apply slight blur to reduce noise
            processed = cv2.medianBlur(processed, 3)

            # Extract text using Tesseract
            text = pytesseract.image_to_string(processed, config=self.ocr_config)

            # Get detailed information including coordinates and confidence
            data = pytesseract.image_to_data(processed, output_type=pytesseract.Output.DICT)

            words = self._extract_words_with_positions(data)
            lines = self._extract_lines(data)

            # Calculate overall confidence
            confidences = [w['confidence'] for w in words if w['confidence'] > 0]
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0

            return {
                'text': text.strip(),
                'words': words,
                'lines': lines,
                'confidence': avg_confidence
            }

        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return {
                'text': f'OCR error: {str(e)}',
                'words': [],
                'lines': [],
                'confidence': 0.0
            }

    def _extract_words_with_positions(self, data: Dict) -> List[Dict]:
        """Extract words with their positions and confidence scores"""
        words = []
        n_boxes = len(data['level'])

        for i in range(n_boxes):
            confidence = int(data['conf'][i])
            if confidence > 30:  # Lower threshold for better detection
                text = data['text'][i].strip()
                if text:  # Only include non-empty text
                    word = {
                        'text': text,
                        'x': data['left'][i],
                        'y': data['top'][i],
                        'width': data['width'][i],
                        'height': data['height'][i],
                        'confidence': confidence
                    }
                    words.append(word)

        return words

    def _extract_lines(self, data: Dict) -> List[str]:
        """Extract text lines from OCR data"""
        lines = []
        current_line = []
        current_y = -1

        for i, level in enumerate(data['level']):
            if level == 4:  # Word level
                confidence = int(data['conf'][i])
                if confidence > 30:
                    text = data['text'][i].strip()
                    y_pos = data['top'][i]

                    # If this word is on a significantly different Y position, start new line
                    if current_y >= 0 and abs(y_pos - current_y) > 10:
                        if current_line:
                            lines.append(' '.join(current_line))
                            current_line = []

                    if text:
                        current_line.append(text)
                        current_y = y_pos

        # Add the last line
        if current_line:
            lines.append(' '.join(current_line))

        return lines

    def detect_slide_number(self, content: Dict) -> Optional[int]:
        """Try to detect slide number from the extracted content"""
        text = content['text']

        # Look for slide number patterns
        number_patterns = [
            r'Slide\s*(\d+)',
            r'(\d+)\s*of\s*\d+',
            r'(\d+)\s*/\s*\d+',
            r'^(\d+)$',  # Just a number on its own
        ]

        for pattern in number_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                try:
                    return int(match.group(1))
                except (ValueError, IndexError):
                    continue

        # Look in the bottom area words for slide numbers
        words = content['words']
        if words:
            # Sort words by Y position (bottom to top)
            bottom_words = sorted([w for w in words if w['confidence'] > 60],
                                key=lambda w: w['y'], reverse=True)

            # Check the bottom 20% of words
            bottom_count = max(1, len(bottom_words) // 5)
            for word in bottom_words[:bottom_count]:
                if word['text'].isdigit():
                    num = int(word['text'])
                    if 1 <= num <= 999:  # Reasonable slide number range
                        return num

        return None

    def extract_title(self, content: Dict) -> str:
        """Extract likely title from content"""
        lines = content['lines']
        if not lines:
            return "No Title Found"

        # Try to find the title (usually first non-empty line, large font)
        words = content['words']
        if words:
            # Find words in the top 30% of the image
            top_words = [w for w in words if w['y'] < (max(w['y'] for w in words) * 0.3)]
            if top_words:
                # Group words by line (similar Y positions)
                line_groups = {}
                for word in top_words:
                    y_key = word['y'] // 20  # Group by 20-pixel bands
                    if y_key not in line_groups:
                        line_groups[y_key] = []
                    line_groups[y_key].append(word)

                # Find the topmost line with reasonable text
                if line_groups:
                    top_line = min(line_groups.keys())
                    title_words = sorted(line_groups[top_line], key=lambda w: w['x'])
                    title = ' '.join(w['text'] for w in title_words)
                    if len(title.strip()) > 0 and len(title) < 150:
                        return title.strip()

        # Fall back to first line
        for line in lines:
            line = line.strip()
            if line and len(line) < 150:  # Reasonable title length
                return line

        return "No Title Found"

    def monitor_slides(self, callback: Callable[[SlideInfo], None], interval: float = 2.0):
        """Monitor slide changes and call callback when changes detected"""
        def monitor_loop():
            print(f"Starting PowerPoint screen monitoring (interval: {interval}s)")
            self.is_monitoring = True

            while self.is_monitoring:
                try:
                    windows = self.find_powerpoint_windows()

                    if windows:
                        # Prioritize slideshow windows, otherwise take the largest
                        target_window = None
                        for window in windows:
                            if 'Slide Show' in window['title']:
                                target_window = window
                                break

                        if not target_window:
                            target_window = max(windows, key=lambda w: w['area'])

                        # Capture and process the window
                        image, pil_image = self.capture_slide_area(target_window)
                        content = self.extract_slide_content(image)

                        # Create content hash to detect changes
                        content_hash = hashlib.md5(content['text'].encode()).hexdigest()

                        if content_hash != self.previous_content_hash and content['text'].strip():
                            slide_number = self.detect_slide_number(content)
                            title = self.extract_title(content)

                            slide_info = SlideInfo(
                                slide_number=slide_number,
                                title=title,
                                content=content['text'],
                                words=content['words'],
                                timestamp=time.time(),
                                window_title=target_window['title'],
                                confidence_score=content['confidence']
                            )

                            self.current_slide_info = slide_info
                            self.previous_content_hash = content_hash

                            # Call the callback
                            if callback:
                                callback(slide_info)

                except Exception as e:
                    print(f"Error in monitoring loop: {e}")

                time.sleep(interval)

        # Start monitoring in separate thread
        self.monitoring_thread = threading.Thread(target=monitor_loop, daemon=True)
        self.monitoring_thread.start()

        return self.monitoring_thread

    def stop_monitoring(self):
        """Stop the monitoring thread"""
        self.is_monitoring = False
        if self.monitoring_thread and self.monitoring_thread.is_alive():
            self.monitoring_thread.join(timeout=5)

    def get_current_slide_info(self) -> Optional[SlideInfo]:
        """Get current slide information (one-time capture)"""
        windows = self.find_powerpoint_windows()

        if not windows:
            return None

        # Prioritize slideshow windows
        target_window = None
        for window in windows:
            if 'Slide Show' in window['title']:
                target_window = window
                break

        if not target_window:
            target_window = max(windows, key=lambda w: w['area'])

        try:
            image, pil_image = self.capture_slide_area(target_window)
            content = self.extract_slide_content(image)

            slide_number = self.detect_slide_number(content)
            title = self.extract_title(content)

            return SlideInfo(
                slide_number=slide_number,
                title=title,
                content=content['text'],
                words=content['words'],
                timestamp=time.time(),
                window_title=target_window['title'],
                confidence_score=content['confidence']
            )

        except Exception as e:
            print(f"Failed to get current slide info: {e}")
            return None

# Example usage and testing
def example_callback(slide_info: SlideInfo):
    """Example callback function for slide changes"""
    print(f"\n=== SLIDE DETECTED ===")
    print(f"Slide {slide_info.slide_number}: {slide_info.title}")
    print(f"Window: {slide_info.window_title}")
    print(f"Confidence: {slide_info.confidence_score:.1f}%")
    print(f"Content preview: {slide_info.content[:200]}{'...' if len(slide_info.content) > 200 else ''}")
    print("=" * 50)

if __name__ == "__main__":
    # Test the screen detector
    detector = PowerPointScreenDetector()

    # Test one-time detection
    current_info = detector.get_current_slide_info()
    if current_info:
        example_callback(current_info)
    else:
        print("No PowerPoint presentation detected")

    # Test continuous monitoring (uncomment to test)
    # print("Starting continuous monitoring. Press Ctrl+C to stop.")
    # try:
    #     detector.monitor_slides(example_callback, interval=1.0)
    #     # Keep the main thread alive
    #     while True:
    #         time.sleep(1)
    # except KeyboardInterrupt:
    #     print("\nStopping monitor...")
    #     detector.stop_monitoring()