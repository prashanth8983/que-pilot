"""
Fallback Content Provider

Provides hardcoded presentation content when automatic detection fails.
"""

from typing import Dict, List
from dataclasses import dataclass
from .ppt_content_extractor import SlideContent, PresentationContent

# Hardcoded presentation content based on the user's example
HARDCODED_PRESENTATION_DATA = {
    "title": "Why are pets important?",
    "slides": [
        {
            "slide_number": 1,
            "title": "Why are pets important?",
            "content": [
                "Pets give love and friendship",
                "Reduce loneliness"
            ],
            "notes": "Introduction slide covering the main benefits pets provide - emotional support and companionship."
        },
        {
            "slide_number": 2,
            "title": "Love and Friendship",
            "content": [
                "Pets help by giving unconditional love and friendship, acting as constant companions who bring comfort and joy.",
                "They also reduce loneliness by providing company, someone to care for, and making homes feel warm and lively."
            ],
            "notes": "Explain how pets provide emotional support through unconditional love and constant companionship."
        },
        {
            "slide_number": 3,
            "title": "Health Benefits",
            "content": [
                "Lower stress & anxiety",
                "Encourage exercise",
                "Improve mood",
                "Pets support our well-being by lowering stress and anxiety through their calming presence, encouraging regular exercise with walks and play, and boosting our mood by bringing joy and positivity into daily life."
            ],
            "notes": "Focus on the physical and mental health benefits that pets provide to their owners."
        },
        {
            "slide_number": 4,
            "title": "Learning Responsibility",
            "content": [
                "Teach care and empathy",
                "Build routine and discipline",
                "Caring for pets teaches empathy by helping us understand and respond to their needs, while also building routine and discipline through regular feeding, grooming, and daily responsibilities."
            ],
            "notes": "Discuss how pet ownership develops important life skills and character traits."
        },
        {
            "slide_number": 5,
            "title": "Social Benefits",
            "content": [
                "Help people make friends",
                "Create shared activities",
                "Pets help people make friends by sparking conversations and connections, and they create shared activities like walks, playtime, or community events that bring people together."
            ],
            "notes": "Explain how pets facilitate social interactions and community building."
        },
        {
            "slide_number": 6,
            "title": "Daily Joy",
            "content": [
                "Bring happiness into daily life",
                "Provide playful moments",
                "Pets bring happiness into daily life with their affectionate nature and add playful moments that make everyday routines more fun and joyful."
            ],
            "notes": "Conclude with how pets enhance daily life through joy and playfulness."
        }
    ]
}

class FallbackContentProvider:
    """Provides fallback presentation content when detection fails."""

    def __init__(self):
        self.hardcoded_content = HARDCODED_PRESENTATION_DATA

    def get_fallback_presentation(self) -> PresentationContent:
        """Get the hardcoded presentation content."""
        slides = []

        for slide_data in self.hardcoded_content["slides"]:
            slide_content = SlideContent(
                slide_number=slide_data["slide_number"],
                title=slide_data["title"],
                text_content=slide_data["content"],
                notes=slide_data.get("notes", ""),
                layout_name="Standard Layout",
                shape_count=len(slide_data["content"]) + 1  # +1 for title
            )
            slides.append(slide_content)

        metadata = {
            'title': self.hardcoded_content["title"],
            'slide_count': len(slides),
            'format': 'hardcoded_fallback',
            'extraction_method': 'fallback',
            'source': 'user_provided_content'
        }

        return PresentationContent(
            file_path="hardcoded://pets_presentation",
            title=self.hardcoded_content["title"],
            total_slides=len(slides),
            slides=slides,
            metadata=metadata
        )

    def get_slide_content(self, slide_number: int) -> str:
        """Get content for a specific slide."""
        if slide_number < 1 or slide_number > len(self.hardcoded_content["slides"]):
            return ""

        slide_data = self.hardcoded_content["slides"][slide_number - 1]
        content_parts = [f"Title: {slide_data['title']}"]

        if slide_data["content"]:
            content_parts.append("Content:")
            content_parts.extend(slide_data["content"])

        if slide_data.get("notes"):
            content_parts.append(f"Notes: {slide_data['notes']}")

        return "\n".join(content_parts)

    def update_hardcoded_content(self, new_content: Dict):
        """Update the hardcoded content with new data."""
        self.hardcoded_content = new_content

    def add_custom_presentation(self, title: str, slides_data: List[Dict]):
        """Add a custom presentation to the fallback options."""
        custom_content = {
            "title": title,
            "slides": slides_data
        }
        self.hardcoded_content = custom_content

    def get_available_presentations(self) -> List[str]:
        """Get list of available fallback presentations."""
        return [self.hardcoded_content["title"]]

# Additional fallback presentations can be added here
ADDITIONAL_PRESENTATIONS = {
    "demo_presentation": {
        "title": "CuePilot Demo Presentation",
        "slides": [
            {
                "slide_number": 1,
                "title": "CuePilot Demo Presentation",
                "content": ["AI-Powered Presentation Assistant"],
                "notes": "Welcome slide introducing CuePilot as an AI presentation assistant."
            },
            {
                "slide_number": 2,
                "title": "Key Features",
                "content": [
                    "Real-time audio transcription",
                    "AI-powered presentation coaching",
                    "Slide content analysis",
                    "Live feedback and suggestions"
                ],
                "notes": "Overview of main CuePilot features for presentation improvement."
            },
            {
                "slide_number": 3,
                "title": "How CuePilot Works",
                "content": [
                    "Monitors your PowerPoint presentation",
                    "Listens to your speech via microphone",
                    "Analyzes content vs. speech alignment",
                    "Provides contextual suggestions"
                ],
                "notes": "Technical explanation of how the AI system operates during presentations."
            },
            {
                "slide_number": 4,
                "title": "Benefits",
                "content": [
                    "Improved presentation flow",
                    "Better audience engagement",
                    "Reduced presenter anxiety",
                    "Professional presentation delivery"
                ],
                "notes": "Benefits users get from using CuePilot during their presentations."
            },
            {
                "slide_number": 5,
                "title": "Getting Started",
                "content": [
                    "Open PowerPoint with your presentation",
                    "Launch CuePilot application",
                    "Click 'Auto-Detect Presentation'",
                    "Enable audio monitoring and start presenting!"
                ],
                "notes": "Step-by-step guide for users to start using CuePilot."
            }
        ]
    }
}

class MultiPresentationFallbackProvider:
    """Provider for multiple fallback presentations."""

    def __init__(self):
        self.presentations = {
            "pets": HARDCODED_PRESENTATION_DATA,
            **ADDITIONAL_PRESENTATIONS
        }
        self.current_presentation = "pets"  # Default

    def get_presentation(self, presentation_id: str = None) -> PresentationContent:
        """Get a specific presentation or the current one."""
        if presentation_id is None:
            presentation_id = self.current_presentation

        if presentation_id not in self.presentations:
            presentation_id = "pets"  # Fallback to default

        provider = FallbackContentProvider()
        provider.update_hardcoded_content(self.presentations[presentation_id])
        return provider.get_fallback_presentation()

    def set_current_presentation(self, presentation_id: str):
        """Set the current presentation."""
        if presentation_id in self.presentations:
            self.current_presentation = presentation_id

    def list_presentations(self) -> Dict[str, str]:
        """List all available presentations."""
        return {
            pid: data["title"]
            for pid, data in self.presentations.items()
        }

# Global instances
fallback_provider = FallbackContentProvider()
multi_fallback_provider = MultiPresentationFallbackProvider()