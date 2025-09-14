"""
CuePilot Integration Service

Integrates the CuePilot LLM with presentation content and audio transcription.
"""

import logging
import time
import sys
from pathlib import Path
from typing import Optional, Dict, Any, List, Callable
from dataclasses import dataclass, field

# Add CuePilot to path
cuepilot_path = Path(__file__).parent.parent.parent / "CuePilot"
if cuepilot_path.exists():
    sys.path.insert(0, str(cuepilot_path))

try:
    from CuePilot.model import ModelInterface
    from CuePilot.main_utils import PresentationFlow, PresentationContentManagement
    CUEPILOT_AVAILABLE = True
except ImportError:
    CUEPILOT_AVAILABLE = False
    logging.warning("CuePilot modules not available")

from .ppt_content_extractor import PresentationContent, SlideContent
from .audio_processor import TranscriptionResult

@dataclass
class CueContext:
    """Context information for CuePilot processing."""
    current_slide: int = 1
    total_slides: int = 1
    presentation_title: str = ""
    slide_text: str = ""
    recent_audio: List[str] = field(default_factory=list)
    presenter_notes: str = ""
    session_history: List[Dict] = field(default_factory=list)

@dataclass
class CuePilotResponse:
    """Response from CuePilot processing."""
    cue_text: str
    confidence: float
    context_used: Dict[str, Any]
    timestamp: float
    response_type: str = "suggestion"  # suggestion, correction, enhancement

class CuePilotIntegration:
    """Service for integrating CuePilot with presentation and audio data."""

    def __init__(self, config_path: Optional[str] = None):
        self.logger = logging.getLogger(__name__)
        self.config_path = config_path

        # CuePilot components
        self.model_interface = None
        self.presentation_flow = None

        # Context management
        self.current_context = CueContext()
        self.presentation_content: Optional[PresentationContent] = None

        # Callbacks
        self.response_callback: Optional[Callable[[CuePilotResponse], None]] = None

        # Audio transcript buffer (keep last N transcripts)
        self.max_audio_history = 10

        self._initialize_cuepilot()

    def _initialize_cuepilot(self):
        """Initialize CuePilot components."""
        if not CUEPILOT_AVAILABLE:
            self.logger.error("CuePilot not available. Integration disabled.")
            return False

        try:
            # Initialize model interface
            self.model_interface = ModelInterface(model="MODEL_GEMMA34BE")

            # Initialize presentation flow
            self.presentation_flow = PresentationFlow()

            self.logger.info("CuePilot initialized successfully")
            return True

        except Exception as e:
            self.logger.error(f"Failed to initialize CuePilot: {e}")
            self.model_interface = None
            self.presentation_flow = None
            return False

    def set_presentation_content(self, content: PresentationContent):
        """Set the current presentation content."""
        self.presentation_content = content
        self.current_context.presentation_title = content.title
        self.current_context.total_slides = content.total_slides
        self.logger.info(f"Presentation content set: {content.title} ({content.total_slides} slides)")

        # Generate initial presentation flow if possible
        if self.presentation_flow:
            try:
                all_text = self._extract_all_slide_text(content)
                visual_description = self._generate_visual_description(content)

                self.presentation_flow.generate_flow(all_text, visual_description)
                self.logger.info("Generated presentation flow from content")
            except Exception as e:
                self.logger.warning(f"Could not generate presentation flow: {e}")

    def update_current_slide(self, slide_number: int):
        """Update the current slide context."""
        if not self.presentation_content:
            return

        if slide_number < 1 or slide_number > self.presentation_content.total_slides:
            self.logger.warning(f"Invalid slide number: {slide_number}")
            return

        self.current_context.current_slide = slide_number

        # Update slide text
        slide_content = self.presentation_content.slides[slide_number - 1]
        self.current_context.slide_text = self._format_slide_text(slide_content)
        self.current_context.presenter_notes = slide_content.notes

        self.logger.info(f"Updated context to slide {slide_number}")

    def process_audio_transcription(self, transcription: TranscriptionResult):
        """Process new audio transcription."""
        if not transcription.text.strip():
            return

        # Add to recent audio history
        self.current_context.recent_audio.append(transcription.text)

        # Maintain buffer size
        if len(self.current_context.recent_audio) > self.max_audio_history:
            self.current_context.recent_audio.pop(0)

        self.logger.info(f"Added audio transcription: {transcription.text[:50]}...")

        # Generate cue if we have sufficient context
        if self.model_interface and len(self.current_context.recent_audio) >= 2:
            try:
                response = self._generate_cue_response()
                if response and self.response_callback:
                    self.response_callback(response)
            except Exception as e:
                self.logger.error(f"Failed to generate cue response: {e}")

    def _generate_cue_response(self) -> Optional[CuePilotResponse]:
        """Generate a cue response based on current context."""
        if not self.model_interface or not self.current_context.recent_audio:
            return None

        try:
            # Prepare context for LLM
            context_prompt = self._build_context_prompt()

            # Call LLM
            messages = [{"role": "user", "content": context_prompt}]
            response = self.model_interface.chat_completion(messages, temperature=0.7)

            if response:
                return CuePilotResponse(
                    cue_text=response.strip(),
                    confidence=0.8,  # Default confidence
                    context_used={
                        "slide": self.current_context.current_slide,
                        "audio_segments": len(self.current_context.recent_audio),
                        "has_slide_text": bool(self.current_context.slide_text),
                        "has_notes": bool(self.current_context.presenter_notes)
                    },
                    timestamp=time.time()
                )

        except Exception as e:
            self.logger.error(f"Error generating cue response: {e}")

        return None

    def _build_context_prompt(self) -> str:
        """Build the context prompt for the LLM."""
        prompt_parts = [
            "You are an AI presentation assistant. Provide helpful, concise suggestions based on the following context:",
            "",
            f"Presentation: {self.current_context.presentation_title}",
            f"Current Slide: {self.current_context.current_slide} of {self.current_context.total_slides}",
        ]

        if self.current_context.slide_text:
            prompt_parts.extend([
                "",
                "Slide Content:",
                self.current_context.slide_text
            ])

        if self.current_context.presenter_notes:
            prompt_parts.extend([
                "",
                "Presenter Notes:",
                self.current_context.presenter_notes
            ])

        if self.current_context.recent_audio:
            prompt_parts.extend([
                "",
                "Recent Speech (what the presenter is saying):",
                "- " + "\n- ".join(self.current_context.recent_audio[-3:])  # Last 3 segments
            ])

        prompt_parts.extend([
            "",
            "Please provide a brief, helpful cue or suggestion (1-2 sentences) for the presenter based on:",
            "- How well their speech aligns with the slide content",
            "- Potential improvements or next points to cover",
            "- Timing and pacing suggestions",
            "",
            "Keep your response conversational and supportive."
        ])

        return "\n".join(prompt_parts)

    def _extract_all_slide_text(self, content: PresentationContent) -> str:
        """Extract all text from presentation for flow generation."""
        all_text = []
        for slide in content.slides:
            slide_text = self._format_slide_text(slide)
            if slide_text:
                all_text.append(f"Slide {slide.slide_number}: {slide_text}")
        return "\n\n".join(all_text)

    def _format_slide_text(self, slide: SlideContent) -> str:
        """Format slide content as text."""
        parts = []
        if slide.title:
            parts.append(f"Title: {slide.title}")
        if slide.text_content:
            parts.append("Content: " + " ".join(slide.text_content))
        return "\n".join(parts) if parts else ""

    def _generate_visual_description(self, content: PresentationContent) -> str:
        """Generate a basic visual description of the presentation."""
        descriptions = []
        for slide in content.slides:
            desc = f"Slide {slide.slide_number} ({slide.layout_name}): {slide.shape_count} elements"
            if slide.title:
                desc += f", titled '{slide.title}'"
            descriptions.append(desc)
        return "; ".join(descriptions)

    def set_response_callback(self, callback: Callable[[CuePilotResponse], None]):
        """Set callback for when cue responses are generated."""
        self.response_callback = callback

    def generate_manual_cue(self, prompt: str) -> Optional[CuePilotResponse]:
        """Generate a cue response for a manual prompt."""
        if not self.model_interface:
            return None

        try:
            # Build context with manual prompt
            context_prompt = f"""
            {self._build_context_prompt()}

            Additional Request: {prompt}

            Please provide a helpful response to the additional request in context of the presentation.
            """

            messages = [{"role": "user", "content": context_prompt}]
            response = self.model_interface.chat_completion(messages, temperature=0.7)

            if response:
                return CuePilotResponse(
                    cue_text=response.strip(),
                    confidence=0.9,
                    context_used={
                        "slide": self.current_context.current_slide,
                        "manual_prompt": True,
                        "prompt": prompt
                    },
                    timestamp=time.time(),
                    response_type="manual"
                )

        except Exception as e:
            self.logger.error(f"Error generating manual cue: {e}")

        return None

    def is_available(self) -> bool:
        """Check if CuePilot integration is available."""
        return CUEPILOT_AVAILABLE and self.model_interface is not None

    def get_current_context(self) -> CueContext:
        """Get current context information."""
        return self.current_context

    def clear_audio_history(self):
        """Clear the audio transcription history."""
        self.current_context.recent_audio.clear()
        self.logger.info("Audio history cleared")

# Service instance
cuepilot_integration = CuePilotIntegration()