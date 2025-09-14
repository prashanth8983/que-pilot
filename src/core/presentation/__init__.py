"""Presentation handling core components."""

# Import detector first (minimal dependencies)
from .detector import PowerPointWindowDetector

# Import tracker and processor with dependency handling
try:
    from .tracker import PresentationTracker
    TRACKER_AVAILABLE = True
except ImportError as e:
    print(f"PresentationTracker not available: {e}")
    PresentationTracker = None
    TRACKER_AVAILABLE = False

try:
    from .processor import ContentProcessor
    PROCESSOR_AVAILABLE = True
except ImportError as e:
    print(f"ContentProcessor not available: {e}")
    ContentProcessor = None
    PROCESSOR_AVAILABLE = False

__all__ = ['PowerPointWindowDetector']
if TRACKER_AVAILABLE:
    __all__.append('PresentationTracker')
if PROCESSOR_AVAILABLE:
    __all__.append('ContentProcessor')
