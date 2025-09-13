"""
Presentation Detector Package

A comprehensive PowerPoint presentation tracking and detection system.

This package provides:
- PowerPoint file processing and analysis
- Cross-platform window detection and monitoring
- Real-time slide tracking and synchronization
- OCR text extraction capabilities
- Object detection in slides
- Search functionality across presentations

Modules:
- ppt_tracker: Core PowerPoint tracking functionality
- window_detector: Cross-platform window detection
"""

__version__ = "1.0.0"
__author__ = "Presentation Detector Team"
__description__ = "PowerPoint presentation tracking and detection system"

# Import main classes for easy access
from .ppt_tracker import PowerPointTracker
from .window_detector import PowerPointWindowDetector, WindowInfo

__all__ = [
    'PowerPointTracker',
    'PowerPointWindowDetector', 
    'WindowInfo'
]
