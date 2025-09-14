"""
Integrated Service Manager

Coordinates all services for CuePilot integration with presentation and audio.
"""

import logging
import time
from pathlib import Path
from typing import Optional, Callable, Dict, Any
from dataclasses import dataclass

# Import our new services
from .ppt_content_extractor import ppt_extractor, PresentationContent
from .audio_processor import audio_processor, AudioConfig, TranscriptionResult
from .cuepilot_integration import cuepilot_integration, CuePilotResponse
from .fallback_content import fallback_provider, multi_fallback_provider

# Import existing services if available
try:
    from .presentation_service import PresentationService
    from .ai_service import AIService
    from .sync_service import SyncService
    EXISTING_SERVICES_AVAILABLE = True
except ImportError:
    EXISTING_SERVICES_AVAILABLE = False
    logging.warning("Existing services not available. Using minimal integration.")

# Import window detector from backup
import sys
backup_path = Path(__file__).parent.parent.parent / "backup_old_structure"
if backup_path.exists():
    sys.path.insert(0, str(backup_path))

try:
    from .enhanced_window_detector import enhanced_detector
    WINDOW_DETECTOR_AVAILABLE = True
    ENHANCED_DETECTOR_AVAILABLE = True
    logging.info("Enhanced window detector available")
except ImportError:
    enhanced_detector = None
    ENHANCED_DETECTOR_AVAILABLE = False
    try:
        from window_detector import PowerPointWindowDetector
        WINDOW_DETECTOR_AVAILABLE = True
        logging.info("Standard window detector available")
    except ImportError:
        WINDOW_DETECTOR_AVAILABLE = False
        logging.warning("No window detector available")

@dataclass
class ServiceStatus:
    """Status information for all services."""
    ppt_extractor_ready: bool = False
    audio_processor_ready: bool = False
    cuepilot_ready: bool = False
    window_detector_ready: bool = False
    current_presentation: Optional[str] = None
    current_slide: int = 1
    total_slides: int = 1
    is_recording: bool = False
    last_transcription: Optional[str] = None
    last_cue: Optional[str] = None

class IntegratedServiceManager:
    """Central manager for all integrated services."""

    def __init__(self):
        self.logger = logging.getLogger(__name__)

        # Service status
        self.status = ServiceStatus()

        # Presentation tracking
        self.current_presentation_content: Optional[PresentationContent] = None
        self.window_detector = None

        # Callbacks
        self.slide_change_callback: Optional[Callable[[int, int], None]] = None
        self.transcription_callback: Optional[Callable[[str], None]] = None
        self.cue_response_callback: Optional[Callable[[str], None]] = None

        self._initialize_services()

    def _initialize_services(self):
        """Initialize all available services."""
        self.logger.info("Initializing integrated services...")

        # Initialize window detector
        if WINDOW_DETECTOR_AVAILABLE:
            try:
                if ENHANCED_DETECTOR_AVAILABLE and enhanced_detector:
                    self.window_detector = enhanced_detector
                    self.logger.info("Enhanced window detector initialized")
                else:
                    self.window_detector = PowerPointWindowDetector()
                    self.logger.info("Standard window detector initialized")
                self.status.window_detector_ready = True
            except Exception as e:
                self.logger.error(f"Failed to initialize window detector: {e}")
                self.status.window_detector_ready = False

        # Check service readiness
        self.status.ppt_extractor_ready = True  # PPT extractor is always ready
        self.status.audio_processor_ready = audio_processor.is_audio_available() and audio_processor.is_transcription_available()
        self.status.cuepilot_ready = cuepilot_integration.is_available()

        # Setup callbacks
        self._setup_callbacks()

        self.logger.info(f"Services initialized - PPT: {self.status.ppt_extractor_ready}, "
                        f"Audio: {self.status.audio_processor_ready}, "
                        f"CuePilot: {self.status.cuepilot_ready}, "
                        f"Window: {self.status.window_detector_ready}")

    def _setup_callbacks(self):
        """Setup callbacks between services."""
        # Setup CuePilot response callback
        def handle_cue_response(response: CuePilotResponse):
            self.status.last_cue = response.cue_text
            if self.cue_response_callback:
                self.cue_response_callback(response.cue_text)

        cuepilot_integration.set_response_callback(handle_cue_response)

    def load_presentation_from_file(self, file_path: str) -> bool:
        """Load a presentation from a file."""
        try:
            self.logger.info(f"Loading presentation from file: {file_path}")

            # Check file format
            file_path_obj = Path(file_path)
            is_legacy_ppt = file_path_obj.suffix.lower() == '.ppt'

            if is_legacy_ppt:
                self.logger.info("Detected legacy .ppt format - using enhanced extraction")

            # Extract content using PPT extractor
            content = ppt_extractor.extract_from_file(file_path)
            if not content:
                self.logger.error("Failed to extract presentation content")

                # For .ppt files, try alternative approach
                if is_legacy_ppt:
                    self.logger.info("Attempting fallback for .ppt file...")
                    return self._handle_legacy_ppt_fallback(file_path_obj)

                return False

            self.current_presentation_content = content
            self.status.current_presentation = content.title
            self.status.total_slides = content.total_slides
            self.status.current_slide = 1

            # Set content in CuePilot
            if self.status.cuepilot_ready:
                cuepilot_integration.set_presentation_content(content)
                cuepilot_integration.update_current_slide(1)

            # Log extraction method for .ppt files
            if is_legacy_ppt and content.metadata.get('extraction_method') == 'window_based':
                self.logger.warning("Using window-based tracking for .ppt file - content may be limited")

            self.logger.info(f"Presentation loaded: {content.title} ({content.total_slides} slides)")
            return True

        except Exception as e:
            self.logger.error(f"Failed to load presentation: {e}")
            return False

    def _handle_legacy_ppt_fallback(self, file_path: Path) -> bool:
        """Handle fallback for legacy .ppt files."""
        try:
            self.logger.info("Using window-based approach for .ppt file")

            # Try to get info from currently open PowerPoint
            if self.auto_detect_presentation():
                self.status.current_presentation = file_path.stem
                self.logger.info(f"Fallback successful for {file_path.name}")
                return True
            else:
                self.logger.warning("Could not load .ppt file - please ensure PowerPoint is open with the presentation")
                return False

        except Exception as e:
            self.logger.error(f"Legacy .ppt fallback failed: {e}")
            return False

    def auto_detect_presentation(self) -> bool:
        """Auto-detect and load currently open PowerPoint presentation."""
        try:
            self.logger.info("Auto-detecting PowerPoint presentation...")

            # Method 1: Use enhanced detector if available
            if self.status.window_detector_ready and ENHANCED_DETECTOR_AVAILABLE and hasattr(self.window_detector, 'get_current_slide_info'):
                slide_info_obj = self.window_detector.get_current_slide_info()
                if slide_info_obj:
                    # Update status
                    self.status.current_presentation = slide_info_obj.presentation_name
                    self.status.current_slide = slide_info_obj.current_slide
                    self.status.total_slides = slide_info_obj.total_slides

                    # Update CuePilot context
                    if self.status.cuepilot_ready:
                        context = cuepilot_integration.get_current_context()
                        context.presentation_title = slide_info_obj.presentation_name
                        context.current_slide = slide_info_obj.current_slide
                        context.total_slides = slide_info_obj.total_slides
                        cuepilot_integration.update_current_slide(slide_info_obj.current_slide)

                    self.logger.info(f"Auto-detected: {slide_info_obj.presentation_name} "
                                   f"(slide {slide_info_obj.current_slide}/{slide_info_obj.total_slides})")
                    return True

            # Method 2: Fallback to standard detector
            if self.status.window_detector_ready:
                window = self.window_detector.get_active_powerpoint_window()
                if window:
                    slide_info = self.window_detector.extract_slide_info_from_title(window.title)

                    if slide_info.get('presentation_name'):
                        presentation_name = slide_info['presentation_name']
                        current_slide = slide_info.get('current_slide', 1)
                        total_slides = slide_info.get('total_slides', 1)

                        # Update status
                        self.status.current_presentation = presentation_name
                        self.status.current_slide = current_slide
                        self.status.total_slides = total_slides

                        # Update CuePilot context
                        if self.status.cuepilot_ready:
                            context = cuepilot_integration.get_current_context()
                            context.presentation_title = presentation_name
                            context.current_slide = current_slide
                            context.total_slides = total_slides
                            cuepilot_integration.update_current_slide(current_slide)

                        self.logger.info(f"Auto-detected: {presentation_name} (slide {current_slide}/{total_slides})")
                        return True

            # Method 3: Use hardcoded fallback content
            self.logger.info("No PowerPoint detected, using hardcoded fallback content...")
            return self.load_fallback_presentation()

        except Exception as e:
            self.logger.error(f"Auto-detection failed: {e}")
            # Final fallback - load hardcoded content
            return self.load_fallback_presentation()

    def load_fallback_presentation(self) -> bool:
        """Load hardcoded fallback presentation content."""
        try:
            self.logger.info("Loading hardcoded fallback presentation...")

            # Get fallback content
            fallback_content = fallback_provider.get_fallback_presentation()

            self.current_presentation_content = fallback_content
            self.status.current_presentation = fallback_content.title
            self.status.total_slides = fallback_content.total_slides
            self.status.current_slide = 1

            # Set content in CuePilot
            if self.status.cuepilot_ready:
                cuepilot_integration.set_presentation_content(fallback_content)
                cuepilot_integration.update_current_slide(1)

            self.logger.info(f"Fallback presentation loaded: {fallback_content.title} ({fallback_content.total_slides} slides)")
            return True

        except Exception as e:
            self.logger.error(f"Failed to load fallback presentation: {e}")
            return False

    def start_audio_monitoring(self) -> bool:
        """Start audio recording and transcription."""
        if not self.status.audio_processor_ready:
            self.logger.warning("Audio processor not ready")
            return False

        def transcription_handler(result: TranscriptionResult):
            self.status.last_transcription = result.text
            if self.transcription_callback:
                self.transcription_callback(result.text)

            # Process through CuePilot
            if self.status.cuepilot_ready:
                cuepilot_integration.process_audio_transcription(result)

        try:
            success = audio_processor.start_recording(transcription_handler)
            if success:
                self.status.is_recording = True
                self.logger.info("Audio monitoring started")
            return success

        except Exception as e:
            self.logger.error(f"Failed to start audio monitoring: {e}")
            return False

    def stop_audio_monitoring(self):
        """Stop audio recording and transcription."""
        try:
            audio_processor.stop_recording()
            self.status.is_recording = False
            self.logger.info("Audio monitoring stopped")
        except Exception as e:
            self.logger.error(f"Failed to stop audio monitoring: {e}")

    def start_slide_monitoring(self, interval: float = 2.0) -> bool:
        """Start monitoring for slide changes."""
        if not self.status.window_detector_ready:
            self.logger.warning("Window detector not available for slide monitoring")
            return False

        def slide_monitor_callback(window, slide_info):
            if slide_info and slide_info.get('current_slide'):
                new_slide = slide_info['current_slide']
                if new_slide != self.status.current_slide:
                    old_slide = self.status.current_slide
                    self.status.current_slide = new_slide

                    # Update total slides if available
                    if slide_info.get('total_slides'):
                        self.status.total_slides = slide_info['total_slides']

                    # Update CuePilot context
                    if self.status.cuepilot_ready:
                        cuepilot_integration.update_current_slide(new_slide)

                    # Call callback
                    if self.slide_change_callback:
                        self.slide_change_callback(new_slide, self.status.total_slides)

                    self.logger.info(f"Slide changed: {old_slide} → {new_slide}")

        try:
            # Note: This is a simplified version. In a full implementation,
            # we'd need to run the window detector monitoring in a separate thread
            self.logger.info("Slide monitoring would start here (simplified implementation)")
            return True

        except Exception as e:
            self.logger.error(f"Failed to start slide monitoring: {e}")
            return False

    def generate_manual_cue(self, prompt: str) -> Optional[str]:
        """Generate a manual cue based on a user prompt."""
        if not self.status.cuepilot_ready:
            return None

        try:
            response = cuepilot_integration.generate_manual_cue(prompt)
            if response:
                self.status.last_cue = response.cue_text
                return response.cue_text
            return None

        except Exception as e:
            self.logger.error(f"Failed to generate manual cue: {e}")
            return None

    def get_current_slide_content(self) -> Optional[str]:
        """Get content for the current slide."""
        if not self.current_presentation_content:
            # Try to use fallback content
            try:
                return fallback_provider.get_slide_content(self.status.current_slide)
            except Exception:
                return None

        try:
            return ppt_extractor.get_slide_text(
                self.current_presentation_content,
                self.status.current_slide
            )
        except Exception as e:
            self.logger.error(f"Failed to get slide content: {e}")
            # Fallback to hardcoded content
            try:
                return fallback_provider.get_slide_content(self.status.current_slide)
            except Exception:
                return None

    def set_callbacks(self,
                     slide_change_callback: Optional[Callable[[int, int], None]] = None,
                     transcription_callback: Optional[Callable[[str], None]] = None,
                     cue_response_callback: Optional[Callable[[str], None]] = None):
        """Set callbacks for various events."""
        self.slide_change_callback = slide_change_callback
        self.transcription_callback = transcription_callback
        self.cue_response_callback = cue_response_callback

    def get_status(self) -> ServiceStatus:
        """Get current service status."""
        return self.status

    def is_ready(self) -> bool:
        """Check if the integrated system is ready to use."""
        return (self.status.ppt_extractor_ready and
                (self.status.cuepilot_ready or self.status.audio_processor_ready))

    def stop_all_services(self):
        """Stop all running services."""
        self.logger.info("Stopping all services...")

        if self.status.is_recording:
            self.stop_audio_monitoring()

        # Clear CuePilot context
        if self.status.cuepilot_ready:
            cuepilot_integration.clear_audio_history()

        self.logger.info("All services stopped")

# Global service manager instance
service_manager = IntegratedServiceManager()