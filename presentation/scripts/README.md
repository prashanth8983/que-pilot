# PowerPoint Tracking Application

A Python GUI application for tracking PowerPoint presentations with automatic detection, OCR text extraction, and object detection capabilities.

## Features

- **PowerPoint File Loading**: Load and parse .pptx and .ppt files
- **Slide Navigation**: Navigate through slides (next/previous/jump to specific slide)
- **Current Slide Tracking**: Keep track of current slide position
- **🆕 Auto-Detection**: Automatically detect running PowerPoint windows (Mac & Windows)
- **🆕 Auto-Sync**: Sync with live PowerPoint presentations in real-time
- **🆕 Window Monitoring**: Monitor slide changes and presentation mode switches
- **Text Extraction**: Extract text content directly from PowerPoint slides
- **OCR Text Extraction**: Use OCR to extract text from slide images
- **Object Detection**: Detect and identify objects/shapes in slides using computer vision
- **Search Functionality**: Search for text across all slides
- **Cross-Platform**: Works on macOS and Windows
- **User-friendly GUI**: Easy-to-use graphical interface with real-time sync

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd que-pilot
```

2. Create and activate a virtual environment (recommended):
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required dependencies:
```bash
pip install -r requirements.txt
```

4. Install Tesseract OCR (required for OCR functionality):
   - **macOS**: `brew install tesseract`
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **Windows**: Download from [Tesseract Wiki](https://github.com/UB-Mannheim/tesseract/wiki)

## Usage

### Running the Application

1. Activate the virtual environment:
```bash
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Start the GUI application:
```bash
python main_app.py
```

**Or use the convenience script:**
```bash
./activate_and_run.sh main_app.py
```

3. Use the GUI application features:

   **🎯 Automatic Detection Features:**
   - **Background scanning**: Continuously scans for PowerPoint windows (every 3 seconds)
   - **Auto-file loading**: Automatically finds and loads presentation files without asking
   - **Real-time sync**: Auto-sync enabled by default for live slide tracking
   - **Manual sync**: 🔄 Sync Now button for on-demand synchronization
   - **Status feedback**: Color-coded status indicator with real-time updates

   **📊 PowerPoint Processing:**
   - Browse and select PowerPoint files with file dialog
   - Navigate through slides using Previous/Next buttons
   - Jump to specific slides with input field
   - View extracted text in tabbed interface
   - Extract OCR text with dedicated button
   - Detect objects with visual display
   - Search for text across all slides with results display

### Testing the Installation

Run the test script to verify everything is working:
```bash
python test_app.py
```

### 🎯 GUI Usage Guide

**Step 1: Launch the Application**
```bash
# Option 1: Quick launch
./launch.sh

# Option 2: Manual launch
source venv/bin/activate
python main_app.py
```

**Step 2: Automatic Detection (No Action Required)**
1. Open PowerPoint with a presentation
2. The app automatically detects and loads your presentation!
3. Watch the status bar for real-time updates:
   - "Scanning for PowerPoint..." → "Auto-loaded: presentation.pptx"
4. ✅ Auto-sync is enabled by default - slides sync automatically!

**Step 3: Use Real-time Features**
- **Auto-Sync**: Enabled by default - syncs every 2-3 seconds automatically
- **🔄 Sync Now**: Manual sync button for immediate sync
- **Status**: Color-coded feedback (green=synced, red=error, etc.)
- **Detection Button**: Shows current status (🔍 Scanning... / ✅ PowerPoint Found)
- Navigate slides in PowerPoint and watch the GUI update automatically!

**Manual File Loading (Optional)**
- Click **Browse PPT File** to select a different presentation manually
- The app will still auto-detect and sync with any open PowerPoint window

### Using the PowerPointTracker Class Programmatically

**Standard Usage:**
```python
from presentation_detector import PowerPointTracker

# Load a PowerPoint file
tracker = PowerPointTracker("path/to/your/presentation.pptx")

# Navigate slides
tracker.next_slide()
tracker.previous_slide()
tracker.go_to_slide(5)

# Get current slide information
current_slide = tracker.get_current_slide_number()
total_slides = tracker.get_total_slides()

# Extract text
slide_text = tracker.extract_slide_text()
ocr_text = tracker.extract_text_with_ocr()

# Detect objects
objects = tracker.detect_objects_in_slide()

# Search for text
matching_slides = tracker.search_text_in_slides("search term")

# Get comprehensive slide information
slide_info = tracker.get_slide_info()
```

**🆕 Auto-Detection Usage:**
```python
from presentation_detector import PowerPointTracker

# Create tracker with auto-detection enabled
tracker = PowerPointTracker(auto_detect=True)

# Auto-detect and load presentation
if tracker.auto_load_presentation():
    print(f"Auto-loaded: {tracker.ppt_path}")

    # Enable auto-sync with PowerPoint window
    tracker.enable_auto_sync()

    # Manually sync once
    tracker.sync_with_powerpoint_window()

    # Get window information
    window_info = tracker.get_window_info()
    print(f"Window: {window_info['window_title']}")

    # Start monitoring (blocking)
    def callback(tracker, window, slide_info):
        print(f"Slide changed to: {slide_info['current_slide']}")

    tracker.start_auto_monitoring(callback, interval=1.0)
```

**Window Detection Only:**
```python
from presentation_detector import PowerPointWindowDetector

detector = PowerPointWindowDetector()

# Get all PowerPoint windows
windows = detector.get_powerpoint_windows()
for window in windows:
    print(f"Found: {window.title}")

# Get the active window
active = detector.get_active_powerpoint_window()
if active:
    slide_info = detector.extract_slide_info_from_title(active.title)
    print(f"Current slide: {slide_info['current_slide']}")

# Monitor window changes
detector.monitor_powerpoint_window(interval=1.0)
```

## Project Structure

```
que-pilot/
├── presentation_detector/          # Main package
│   ├── __init__.py                # Package initialization
│   ├── ppt_tracker.py             # Core PowerPoint tracking
│   ├── window_detector.py         # Cross-platform window detection
│   ├── core/                      # Core functionality
│   │   └── __init__.py
│   ├── detectors/                 # Detection modules
│   │   └── __init__.py
│   ├── processors/               # Processing modules
│   │   └── __init__.py
│   └── utils/                    # Utility functions
│       └── __init__.py
├── main_app.py                   # GUI application
├── auto_detect_demo.py           # Demo script
├── test_app.py                   # Test suite
├── quick_test.py                 # Quick test script
├── launch.sh                     # Launch script
├── activate_and_run.sh           # Environment script
├── requirements.txt              # Python dependencies
├── venv/                         # Python virtual environment
├── LICENSE                       # MIT License
└── README.md                     # This file
```

## Dependencies

- **python-pptx**: PowerPoint file parsing
- **opencv-python**: Computer vision and image processing
- **pytesseract**: OCR text extraction
- **Pillow**: Image processing
- **numpy**: Numerical computations
- **pyobjc-framework-Cocoa**: macOS window detection
- **pyobjc-framework-Quartz**: macOS window management
- **pywin32**: Windows window detection (Windows only)
- **psutil**: Cross-platform process utilities
- **tkinter**: GUI framework (usually included with Python)

## Features Overview

### PowerPoint Processing
- Load and parse PowerPoint presentations
- Extract text content from slides
- Navigate through slides programmatically

### Auto-Detection & Sync
- **Cross-platform window detection** (macOS & Windows)
- **Real-time slide tracking** from live PowerPoint presentations
- **Automatic file discovery** based on window titles
- **Live synchronization** with presentation mode changes
- **One-click detection** with user-friendly GUI interface

### OCR Capabilities
- Convert slides to images
- Extract text using Tesseract OCR
- Handle various text formats and fonts

### Object Detection
- Detect shapes and objects in slides
- Provide bounding box coordinates
- Calculate object areas and properties

### Search and Navigation
- Search text across all slides
- Jump to specific slides
- Track current position in presentation

## Limitations

- Object detection is basic and primarily detects shapes/contours
- OCR accuracy depends on slide image quality
- Some advanced PowerPoint features may not be fully supported

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the LICENSE file for details.