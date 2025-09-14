"""
Application settings and configuration management.
"""

import os
from pathlib import Path
from typing import Dict, Any
from dataclasses import dataclass, field


@dataclass
class AppSettings:
    """Application settings container."""
    
    # Application info
    app_name: str = "AI Presentation Copilot"
    app_version: str = "1.0.0"
    
    # Window settings
    window_width: int = 1400
    window_height: int = 850
    window_min_width: int = 1200
    window_min_height: int = 750
    
    # Theme settings
    theme: str = "dracula"
    
    # AI/ML settings
    whisper_model: str = "medium"
    confidence_threshold: float = 0.75
    llm_temperature: float = 0.7
    
    # Presentation settings
    auto_sync_enabled: bool = True
    real_time_assistance: bool = True
    auto_answer_questions: bool = False
    
    # File paths
    models_dir: Path = field(default_factory=lambda: Path("assets/models"))
    icons_dir: Path = field(default_factory=lambda: Path("assets/icons"))
    styles_dir: Path = field(default_factory=lambda: Path("assets/styles"))
    
    # Performance settings
    sync_interval: float = 1.0
    transcription_buffer_size: int = 1024
    
    def __post_init__(self):
        """Initialize paths relative to project root."""
        project_root = Path(__file__).parent.parent
        self.models_dir = project_root / self.models_dir
        self.icons_dir = project_root / self.icons_dir
        self.styles_dir = project_root / self.styles_dir


@dataclass
class DraculaTheme:
    """Dracula theme color palette."""
    
    # Background colors
    bg_main: str = "#282A36"
    bg_secondary: str = "#21222C"
    bg_input_border: str = "#44475A"
    
    # Text colors
    text_primary: str = "#F8F8F2"
    text_secondary: str = "#6272A4"
    
    # Accent colors
    accent_primary: str = "#BD93F9"      # Purple
    accent_secondary: str = "#8BE9FD"    # Cyan
    accent_success: str = "#50FA7B"      # Green
    accent_warning: str = "#F1FA8C"      # Yellow
    accent_error: str = "#FF5555"        # Red
    accent_orange: str = "#FFB86C"       # Orange
    accent_pink: str = "#FF79C6"         # Pink
    
    # UI elements
    border_color: str = "#44475A"
    card_bg: str = "#383A47"
    input_bg: str = "#21222C"


# Global settings instance
settings = AppSettings()
dracula_theme = DraculaTheme()
