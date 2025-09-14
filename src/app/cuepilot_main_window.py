"""
Enhanced Main Window with CuePilot Integration

Extends the existing main window to include CuePilot functionality.
"""

import sys
import logging
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QMessageBox, QButtonGroup,
    QTextEdit, QSplitter, QGroupBox, QCheckBox, QSlider, QComboBox
)
from PySide6.QtCore import Signal, Slot, QPoint, Qt, QSize, QTimer
from PySide6.QtGui import QFont, QIcon, QPixmap, QPainter

# Import the original main window
from .main_window import MainWindow

# Import our integrated service manager
try:
    from ..services.integrated_service_manager import service_manager, ServiceStatus
    CUEPILOT_AVAILABLE = True
except ImportError:
    CUEPILOT_AVAILABLE = False
    logging.warning("CuePilot integration not available")

class CuePilotWidget(QWidget):
    """Widget for CuePilot controls and display."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self.setup_ui()
        self.setup_connections()

    def setup_ui(self):
        """Setup the CuePilot control interface."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Title
        title = QLabel("AI CuePilot")
        title.setObjectName("cuePilotTitle")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Status Group
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout(status_group)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        status_layout.addWidget(self.status_label)

        self.slide_info = QLabel("Slide: - / -")
        self.slide_info.setObjectName("slideInfo")
        status_layout.addWidget(self.slide_info)

        layout.addWidget(status_group)

        # Controls Group
        controls_group = QGroupBox("Controls")
        controls_layout = QVBoxLayout(controls_group)

        # Audio monitoring
        audio_layout = QHBoxLayout()
        self.audio_checkbox = QCheckBox("Audio Monitoring")
        self.audio_checkbox.setEnabled(CUEPILOT_AVAILABLE)
        audio_layout.addWidget(self.audio_checkbox)
        controls_layout.addLayout(audio_layout)

        # Slide tracking
        slide_layout = QHBoxLayout()
        self.slide_checkbox = QCheckBox("Slide Tracking")
        self.slide_checkbox.setEnabled(CUEPILOT_AVAILABLE)
        slide_layout.addWidget(self.slide_checkbox)
        controls_layout.addLayout(slide_layout)

        # Auto-detect button
        self.detect_button = QPushButton("Auto-Detect Presentation")
        self.detect_button.setEnabled(CUEPILOT_AVAILABLE)
        controls_layout.addWidget(self.detect_button)

        layout.addWidget(controls_group)

        # Transcription Display
        transcription_group = QGroupBox("Live Transcription")
        transcription_layout = QVBoxLayout(transcription_group)

        self.transcription_display = QTextEdit()
        self.transcription_display.setMaximumHeight(100)
        self.transcription_display.setReadOnly(True)
        self.transcription_display.setPlaceholderText("Audio transcription will appear here...")
        transcription_layout.addWidget(self.transcription_display)

        layout.addWidget(transcription_group)

        # Cue Response Display
        cue_group = QGroupBox("AI Cues")
        cue_layout = QVBoxLayout(cue_group)

        self.cue_display = QTextEdit()
        self.cue_display.setMaximumHeight(150)
        self.cue_display.setReadOnly(True)
        self.cue_display.setPlaceholderText("AI suggestions will appear here...")
        cue_layout.addWidget(self.cue_display)

        # Manual prompt
        prompt_layout = QHBoxLayout()
        self.manual_prompt = QTextEdit()
        self.manual_prompt.setMaximumHeight(60)
        self.manual_prompt.setPlaceholderText("Ask for specific guidance...")
        self.ask_button = QPushButton("Ask AI")
        self.ask_button.setMaximumWidth(80)
        self.ask_button.setEnabled(CUEPILOT_AVAILABLE)

        prompt_layout.addWidget(self.manual_prompt)
        prompt_layout.addWidget(self.ask_button)
        cue_layout.addLayout(prompt_layout)

        layout.addWidget(cue_group)

        layout.addStretch()

    def setup_connections(self):
        """Setup signal connections."""
        if CUEPILOT_AVAILABLE:
            self.audio_checkbox.toggled.connect(self.toggle_audio_monitoring)
            self.slide_checkbox.toggled.connect(self.toggle_slide_tracking)
            self.detect_button.clicked.connect(self.auto_detect_presentation)
            self.ask_button.clicked.connect(self.ask_manual_question)

            # Setup service callbacks
            service_manager.set_callbacks(
                slide_change_callback=self.on_slide_change,
                transcription_callback=self.on_transcription,
                cue_response_callback=self.on_cue_response
            )

    def toggle_audio_monitoring(self, enabled):
        """Toggle audio monitoring on/off."""
        if not CUEPILOT_AVAILABLE:
            return

        if enabled:
            if service_manager.start_audio_monitoring():
                self.status_label.setText("Listening...")
            else:
                self.audio_checkbox.setChecked(False)
                QMessageBox.warning(self, "Audio Error",
                                  "Could not start audio monitoring. Check your microphone.")
        else:
            service_manager.stop_audio_monitoring()
            self.status_label.setText("Ready")

    def toggle_slide_tracking(self, enabled):
        """Toggle slide tracking on/off."""
        if not CUEPILOT_AVAILABLE:
            return

        if enabled:
            if service_manager.start_slide_monitoring():
                self.status_label.setText("Tracking slides...")
            else:
                self.slide_checkbox.setChecked(False)
                QMessageBox.warning(self, "Tracking Error",
                                  "Could not start slide tracking. Ensure PowerPoint is running.")
        else:
            self.status_label.setText("Ready")

    def auto_detect_presentation(self):
        """Auto-detect current presentation."""
        if not CUEPILOT_AVAILABLE:
            return

        self.status_label.setText("Detecting...")
        if service_manager.auto_detect_presentation():
            status = service_manager.get_status()
            self.status_label.setText(f"Detected: {status.current_presentation}")
            self.slide_info.setText(f"Slide: {status.current_slide} / {status.total_slides}")
        else:
            self.status_label.setText("No presentation found")
            QMessageBox.information(self, "Detection",
                                  "No PowerPoint presentation detected. Ensure PowerPoint is running.")

    def ask_manual_question(self):
        """Ask manual question to AI."""
        if not CUEPILOT_AVAILABLE:
            return

        prompt = self.manual_prompt.toPlainText().strip()
        if not prompt:
            return

        self.ask_button.setText("Asking...")
        self.ask_button.setEnabled(False)

        response = service_manager.generate_manual_cue(prompt)
        if response:
            self.cue_display.append(f"\n🤔 You: {prompt}")
            self.cue_display.append(f"🤖 AI: {response}")
            self.cue_display.append("─" * 40)
        else:
            self.cue_display.append(f"\n❌ Could not process: {prompt}")

        self.manual_prompt.clear()
        self.ask_button.setText("Ask AI")
        self.ask_button.setEnabled(True)

    def on_slide_change(self, slide_num, total_slides):
        """Handle slide change notification."""
        self.slide_info.setText(f"Slide: {slide_num} / {total_slides}")
        self.cue_display.append(f"\n📄 Moved to slide {slide_num}")

    def on_transcription(self, text):
        """Handle transcription update."""
        self.transcription_display.append(text)
        # Keep only last few lines to prevent overflow
        if self.transcription_display.document().lineCount() > 5:
            cursor = self.transcription_display.textCursor()
            cursor.movePosition(cursor.Start)
            cursor.select(cursor.LineUnderCursor)
            cursor.removeSelectedText()

    def on_cue_response(self, cue_text):
        """Handle AI cue response."""
        self.cue_display.append(f"\n💡 AI Suggestion: {cue_text}")
        self.cue_display.append("─" * 40)

    def load_presentation_file(self, file_path):
        """Load a presentation file."""
        if CUEPILOT_AVAILABLE and service_manager.load_presentation_from_file(file_path):
            status = service_manager.get_status()
            self.status_label.setText(f"Loaded: {status.current_presentation}")
            self.slide_info.setText(f"Slide: {status.current_slide} / {status.total_slides}")
            return True
        return False

class CuePilotMainWindow(MainWindow):
    """Enhanced main window with CuePilot integration."""

    def __init__(self):
        super().__init__()
        self.cuepilot_widget = None
        self.setup_cuepilot_integration()

    def setup_cuepilot_integration(self):
        """Setup CuePilot integration in the main window."""
        if not CUEPILOT_AVAILABLE:
            self.logger = logging.getLogger(__name__)
            self.logger.warning("CuePilot integration not available")
            return

        # Create CuePilot widget
        self.cuepilot_widget = CuePilotWidget(self)

        # Modify the main layout to include CuePilot
        self.integrate_cuepilot_widget()

    def integrate_cuepilot_widget(self):
        """Integrate CuePilot widget into the main window layout."""
        if not self.cuepilot_widget:
            return

        # Get the main content area
        central_widget = self.centralWidget()
        if not central_widget:
            return

        # Find the content area and replace with splitter
        main_layout = central_widget.layout()
        if not main_layout or main_layout.count() < 2:
            return

        # Get the content widget (should be the second item after title bar)
        content_widget = main_layout.itemAt(1).widget()
        if not content_widget:
            return

        content_layout = content_widget.layout()
        if not content_layout or content_layout.count() < 2:
            return

        # Get the right content widget
        right_content_widget = content_layout.itemAt(1).widget()
        if not right_content_widget:
            return

        # Remove from layout and create splitter
        content_layout.removeWidget(right_content_widget)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(right_content_widget)
        splitter.addWidget(self.cuepilot_widget)
        splitter.setSizes([700, 300])  # Give more space to main content

        content_layout.addWidget(splitter)

    def show_live_view(self, file_path=None):
        """Override to integrate with CuePilot."""
        # Call parent method
        result = super().show_live_view(file_path)

        # Load presentation in CuePilot if file_path provided
        if file_path and self.cuepilot_widget:
            self.cuepilot_widget.load_presentation_file(file_path)
        elif self.cuepilot_widget and CUEPILOT_AVAILABLE:
            # Try auto-detection
            self.cuepilot_widget.auto_detect_presentation()

        return result

    def stop_presentation(self):
        """Override to stop CuePilot services."""
        if CUEPILOT_AVAILABLE:
            service_manager.stop_all_services()
            if self.cuepilot_widget:
                self.cuepilot_widget.audio_checkbox.setChecked(False)
                self.cuepilot_widget.slide_checkbox.setChecked(False)
                self.cuepilot_widget.status_label.setText("Ready")

        # Call parent method
        super().stop_presentation()

def main():
    """Main entry point for CuePilot integrated application."""
    app = QApplication(sys.argv)
    app.setApplicationName("AI Presentation Copilot with CuePilot")

    # Set application font
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Create main window
    if CUEPILOT_AVAILABLE:
        window = CuePilotMainWindow()
        print("🚀 CuePilot integration loaded successfully!")
    else:
        window = MainWindow()
        print("⚠️  Running without CuePilot integration")

    window.show()
    return app.exec()

if __name__ == "__main__":
    sys.exit(main())