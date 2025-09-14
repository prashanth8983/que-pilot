"""
Main entry point for AI Presentation Copilot.
"""

import sys
import os
import logging
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# Add the src directory to the Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

# Import the main window from the new modular structure
from app.main_window import MainWindow

# Import config
try:
    from config.settings import settings, dracula_theme
except ImportError:
    # Fallback configuration
    class Settings:
        app_name = "AI Presentation Copilot"
        app_version = "1.0.0"
        window_width = 1200
        window_height = 720
        window_min_width = 900
        window_min_height = 600
    
    class DraculaTheme:
        bg_main = "#282a36"
        bg_secondary = "#21222c"
        bg_input_border = "#44475a"
        card_bg = "#44475a"
        border_color = "#6272a4"
        text_primary = "#f8f8f2"
        text_secondary = "#6272a4"
        accent_primary = "#bd93f9"
        accent_secondary = "#8be9fd"
        accent_success = "#50fa7b"
        accent_warning = "#ffb86c"
        accent_error = "#ff5555"
        input_bg = "#44475a"
    
    settings = Settings()
    dracula_theme = DraculaTheme()


def setup_logging():
    """Setup application logging."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler('ai_presentation_copilot.log'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logger = logging.getLogger(__name__)
    return logger


def check_dependencies():
    """Check if required dependencies are available."""
    logger = logging.getLogger(__name__)
    
    try:
        import PySide6
        logger.info("PySide6 is available")
    except ImportError:
        logger.error("PySide6 not found. Please install it with: pip install PySide6")
        return False
    
    try:
        import cv2
        logger.info("OpenCV is available")
    except ImportError:
        logger.warning("OpenCV not found. Some features may be limited.")
    
    try:
        import numpy
        logger.info("NumPy is available")
    except ImportError:
        logger.warning("NumPy not found. Some features may be limited.")
    
    try:
        import psutil
        logger.info("psutil is available")
    except ImportError:
        logger.warning("psutil not found. Some features may be limited.")
    
    return True


def check_system_compatibility():
    """Check system compatibility."""
    logger = logging.getLogger(__name__)
    
    # Check Python version
    if sys.version_info < (3, 8):
        logger.error(f"Python {sys.version_info.major}.{sys.version_info.minor} is not supported. Please use Python 3.8 or higher.")
        return False
    
    logger.info(f"Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} is compatible")
    
    # Check operating system
    import platform
    system = platform.system()
    logger.info(f"Running on {system} {platform.release()}")
    
    return True


def main():
    """Main application entry point."""
    # Setup logging
    logger = setup_logging()
    logger.info("Starting AI Presentation Copilot...")
    
    # Check system compatibility
    if not check_system_compatibility():
        logger.error("System compatibility check failed")
        return 1
    
    # Check dependencies
    if not check_dependencies():
        logger.error("Dependency check failed")
        return 1
    
    # Create QApplication
    app = QApplication(sys.argv)
    app.setApplicationName(settings.app_name)
    app.setApplicationVersion(settings.app_version)
    
    # Set application font
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Set application style
    app.setStyle('Fusion')
    
    # Set application palette for dark theme
    palette = app.palette()
    palette.setColor(palette.ColorRole.Window, dracula_theme.bg_main)
    palette.setColor(palette.ColorRole.WindowText, dracula_theme.text_primary)
    palette.setColor(palette.ColorRole.Base, dracula_theme.input_bg)
    palette.setColor(palette.ColorRole.AlternateBase, dracula_theme.bg_secondary)
    palette.setColor(palette.ColorRole.ToolTipBase, dracula_theme.card_bg)
    palette.setColor(palette.ColorRole.ToolTipText, dracula_theme.text_primary)
    palette.setColor(palette.ColorRole.Text, dracula_theme.text_primary)
    palette.setColor(palette.ColorRole.Button, dracula_theme.bg_secondary)
    palette.setColor(palette.ColorRole.ButtonText, dracula_theme.text_primary)
    palette.setColor(palette.ColorRole.BrightText, dracula_theme.accent_error)
    palette.setColor(palette.ColorRole.Link, dracula_theme.accent_primary)
    palette.setColor(palette.ColorRole.Highlight, dracula_theme.accent_primary)
    palette.setColor(palette.ColorRole.HighlightedText, dracula_theme.text_primary)
    app.setPalette(palette)
    
    try:
        # Create and show main window
        logger.info("Creating main window...")
        window = MainWindow()
        window.show()
        
        logger.info("Application started successfully")
        
        # Run application
        return app.exec()
        
    except Exception as e:
        logger.error(f"Failed to create main window: {e}")
        QMessageBox.critical(
            None, 
            "Error", 
            f"Failed to start application: {e}"
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
