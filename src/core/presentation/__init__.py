"""Presentation handling core components."""

from .tracker import PresentationTracker
from .detector import PowerPointWindowDetector
from .processor import ContentProcessor

__all__ = ['PresentationTracker', 'PowerPointWindowDetector', 'ContentProcessor']
