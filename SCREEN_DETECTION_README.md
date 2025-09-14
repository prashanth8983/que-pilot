# PowerPoint Screen Detection with OCR

This implementation adds Windows API-based screen capture and OCR functionality to detect PowerPoint slide numbers and content. It works as a fallback when traditional window title parsing or AppleScript methods fail.

## Features

- **Screen Capture**: Captures PowerPoint windows using platform-specific APIs
- **OCR Text Extraction**: Uses Tesseract OCR to extract slide content and numbers
- **Cross-Platform**: Supports both Windows (Win32 API) and macOS (Quartz)
- **Fallback Integration**: Automatically used when primary detection methods fail
- **Content Analysis**: Extracts slide titles, content, and slide numbers with confidence scores

## Installation

### Core Dependencies
```bash
# Install OCR dependencies
pip install -r requirements-ocr.txt

# Or install individually:
pip install pytesseract opencv-python pillow numpy pyautogui
```

### Platform-Specific Dependencies

**Windows:**
```bash
pip install pywin32 psutil
```

**macOS:**
```bash
pip install pyobjc-framework-Quartz pyobjc-framework-Cocoa
```

### System Requirements

You also need to install Tesseract OCR on your system:

- **Windows**: Download from [UB Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt-get install tesseract-ocr`

## Usage

### Standalone Screen Detector

```python
from src.core.presentation.screen_detector import PowerPointScreenDetector

# Initialize the detector
detector = PowerPointScreenDetector()

# Get current slide info (one-time detection)
slide_info = detector.get_current_slide_info()
if slide_info:
    print(f"Slide {slide_info.slide_number}: {slide_info.title}")
    print(f"Content: {slide_info.content}")
    print(f"Confidence: {slide_info.confidence_score}%")

# Monitor slides continuously
def on_slide_change(slide_info):
    print(f"New slide: {slide_info.slide_number}")

detector.monitor_slides(on_slide_change, interval=2.0)
```

### Integrated with Existing Detector

The screen detector is automatically integrated as a fallback:

```python
from src.core.presentation.detector import PowerPointWindowDetector

detector = PowerPointWindowDetector()
slide_info = detector.get_current_slide_info()

# If traditional methods fail, OCR is automatically used
if slide_info.get('detection_method') == 'screen_ocr':
    print("OCR detection was used")
    print(f"Confidence: {slide_info.get('ocr_confidence')}%")
```

## How It Works

### 1. Window Detection
- **Windows**: Uses Win32 API to enumerate visible windows
- **macOS**: Uses Quartz to get window information
- Filters for PowerPoint windows by title and process name

### 2. Screen Capture
- Captures the PowerPoint window area using pyautogui
- Applies padding to ensure full content capture
- Handles multiple monitor setups

### 3. OCR Processing
- Converts image to grayscale for better OCR accuracy
- Applies adaptive thresholding and noise reduction
- Uses Tesseract with optimized configuration
- Extracts text with position and confidence data

### 4. Content Analysis
- **Slide Numbers**: Searches for patterns like "Slide 5", "5/10", "5 of 10"
- **Titles**: Identifies likely titles from top-area text
- **Content**: Extracts full slide text content
- **Confidence**: Calculates average OCR confidence score

## Configuration

### OCR Settings
You can customize OCR behavior in `screen_detector.py`:

```python
# Current OCR configuration
self.ocr_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,!?;:()[]{}"-\' '
```

### Monitoring Intervals
Adjust detection frequency:

```python
# Check every 2 seconds (default)
detector.monitor_slides(callback, interval=2.0)

# More responsive (every 1 second)
detector.monitor_slides(callback, interval=1.0)
```

## Integration Points

The screen detector integrates with the existing system at these points:

1. **detector.py**: Fallback integration in `get_current_slide_info()`
2. **__init__.py**: Graceful import handling for missing dependencies
3. **WindowInfo**: Compatible with existing window detection structures

## Limitations

- **Performance**: OCR processing takes 1-2 seconds per detection
- **Accuracy**: Depends on slide design, font size, and image quality
- **Dependencies**: Requires additional packages and system setup
- **Privacy**: Screen capture may trigger security warnings on some systems

## Troubleshooting

### Common Issues

1. **"No module named 'pytesseract'"**
   - Install: `pip install pytesseract`
   - Install system Tesseract OCR

2. **"TesseractNotFoundError"**
   - Ensure Tesseract is installed and in PATH
   - On Windows, you may need to specify the path explicitly

3. **Low OCR accuracy**
   - Ensure PowerPoint window is clearly visible
   - Avoid overlapping windows
   - Check slide font sizes and contrast

4. **No windows detected**
   - Verify PowerPoint is running
   - Check platform-specific dependencies are installed
   - Try running with administrator/sudo privileges

### Debug Output

The system provides detailed logging:

```
=== OCR DETECTION ===
Slide 5: Introduction to Data Science
Confidence: 87.3%
Content: Welcome to our presentation on data science fundamentals...
==============================
```

## Files Structure

```
src/core/presentation/
├── screen_detector.py          # Main OCR detector implementation
├── detector.py                 # Integrated detector with OCR fallback
└── __init__.py                # Graceful dependency handling

requirements-ocr.txt            # OCR dependencies
test_screen_detection.py        # Comprehensive test suite
test_basic_detection.py         # Basic functionality test
SCREEN_DETECTION_README.md      # This documentation
```

## Future Enhancements

- **Performance**: Add image caching and incremental detection
- **Accuracy**: Improve OCR preprocessing and slide layout analysis
- **Features**: Add slide change detection, presenter notes extraction
- **Integration**: Direct integration with presentation software APIs