"""
Live view for the AI Presentation Copilot application.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QProgressBar, QFrame
)
from PySide6.QtCore import Slot, Qt, QSize, QTimer, Signal, QMetaObject, Q_ARG
from PySide6.QtGui import QFont

# Import config and widgets with fallbacks
try:
    from ...config.settings import dracula_theme
except ImportError:
    # Fallback theme
    class DraculaTheme:
        bg_main = "#282a36"
        bg_secondary = "#21222c"
        text_primary = "#f8f8f2"
        text_secondary = "#6272a4"
        accent_primary = "#bd93f9"
        accent_secondary = "#8be9fd"
        accent_success = "#50fa7b"
        accent_warning = "#ffb86c"
        accent_error = "#ff5555"
        card_bg = "#44475a"
    dracula_theme = DraculaTheme()

try:
    from ..widgets import CircularProgress
except ImportError:
    # Fallback circular progress
    from PySide6.QtWidgets import QProgressBar
    CircularProgress = QProgressBar

# Import icon loader
try:
    from ...core.utils.icon_loader import load_icon
except ImportError:
    # Fallback if icon loader fails
    from PySide6.QtGui import QIcon
    def load_icon(name, size=None, color=None):
        return QIcon(f"assets/icons/{name}.svg")


class LiveView(QWidget):
    """Live session view with real-time presentation assistance."""

    def __init__(self, ai_service, pres_service, sync_service, parent=None):
        super().__init__(parent)
        self.ai_service = ai_service
        self.pres_service = pres_service
        self.sync_service = sync_service

        # Track callback connections to avoid duplicates
        self.callbacks_connected = False
        self.slide_progress_timer = None

        self.init_ui()

    def init_ui(self):
        """Initialize the live view UI to match the modern SVG design."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Main content container with improved layout
        main_container = QFrame()
        main_container.setStyleSheet(f"""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #343746, stop:1 {dracula_theme.bg_main});
            border-radius: 12px;
        """)
        main_layout = QHBoxLayout(main_container)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Left Panel - Current Slide Info (narrower, modern sidebar)
        left_panel = self.create_left_panel()
        main_layout.addWidget(left_panel)

        # Center Panel - AI Assistance & Transcription (wider main area)
        center_panel = self.create_center_panel()
        main_layout.addWidget(center_panel)

        # Right Panel - Live Analytics (compact sidebar)
        right_panel = self.create_right_panel()
        main_layout.addWidget(right_panel)

        layout.addWidget(main_container)

    def create_left_panel(self):
        """Create the left panel with current slide info and speaker cues."""
        panel = QFrame()
        panel.setObjectName("leftPanel")
        panel.setFixedWidth(220)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(15, 20, 15, 20)
        layout.setSpacing(20)

        # Presentation title
        self.presentation_title = QLabel("No Presentation Loaded")
        self.presentation_title.setObjectName("presentationTitle")
        self.presentation_title.setWordWrap(True)
        layout.addWidget(self.presentation_title)

        layout.addSpacing(5)

        # Current Slide title
        current_slide_title = QLabel("Current Slide")
        current_slide_title.setObjectName("sectionTitle")
        layout.addWidget(current_slide_title)

        layout.addSpacing(10)

        # Speaker Cues section
        cues_title = QLabel("Speaker Cues")
        cues_title.setObjectName("subsectionTitle")
        layout.addWidget(cues_title)

        # Speaker cues content with bullet points
        cues_container = QFrame()
        cues_layout = QVBoxLayout(cues_container)
        cues_layout.setSpacing(10)
        cues_layout.setContentsMargins(0, 10, 0, 0)

        # Create individual cue items with bullet points
        self.cue_items = []
        default_cues = [
            "Load a presentation to see speaker cues",
            "Navigate to live mode to start presenting",
            "Cues will update based on your current slide"
        ]

        # Create a temporary speaker_cues widget to maintain compatibility
        self.speaker_cues = QTextEdit()
        self.speaker_cues.setPlainText("\n".join(f"• {cue}" for cue in default_cues))
        self.speaker_cues.setStyleSheet(f"""
            QTextEdit {{
                background-color: {getattr(dracula_theme, 'bg_secondary', '#21222c')};
                border: 1px solid {getattr(dracula_theme, 'bg_input_border', '#44475a')};
                border-radius: 8px;
                padding: 12px;
                color: {getattr(dracula_theme, 'text_primary', '#f8f8f2')};
                font-size: 13px;
                line-height: 1.4;
            }}
        """)
        self.speaker_cues.setMaximumHeight(120)
        layout.addWidget(self.speaker_cues)

        for cue_text in default_cues:
            cue_item = self.create_cue_item(cue_text)
            cues_layout.addWidget(cue_item)
            self.cue_items.append(cue_item)

        layout.addWidget(cues_container)

        # Remove next slide section - not needed

        layout.addStretch()
        return panel

    def create_cue_item(self, text):
        """Create a speaker cue item with bullet point."""
        container = QFrame()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Bullet point
        bullet = QLabel("•")
        bullet.setObjectName("cueBullet")
        bullet.setFixedSize(10, 20)
        layout.addWidget(bullet)

        # Cue text
        cue_text = QLabel(text)
        cue_text.setObjectName("cueText")
        cue_text.setWordWrap(True)
        layout.addWidget(cue_text)

        return container

    def create_center_panel(self):
        """Create the center panel with AI assistance and transcription."""
        panel = QFrame()
        panel.setObjectName("centerPanel")
        panel.setFixedWidth(340)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # AI Assistance Section
        ai_title = QLabel("AI Assistance")
        ai_title.setObjectName("sectionTitle")
        layout.addWidget(ai_title)

        # AI suggestion content
        self.ai_suggestion = QLabel()
        self.ai_suggestion.setObjectName("aiSuggestion")
        self.ai_suggestion.setWordWrap(True)
        self.ai_suggestion.setText("🤖 AI will provide real-time assistance based on your speech patterns and presentation context.")
        self.ai_suggestion.setMinimumHeight(80)
        layout.addWidget(self.ai_suggestion)

        # Divider line
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setObjectName("divider")
        layout.addWidget(divider)

        # Live Transcription Section
        transcript_title = QLabel("Live Transcription Feed")
        transcript_title.setObjectName("sectionTitle")
        layout.addWidget(transcript_title)

        # Transcription feed
        self.transcription_feed = QTextEdit()
        self.transcription_feed.setObjectName("transcriptionFeed")
        self.transcription_feed.setReadOnly(True)
        self.transcription_feed.setPlainText("🎤 Ready for live transcription...\nSpeak to see your words appear here in real-time.")
        self.transcription_feed.setMinimumHeight(300)
        layout.addWidget(self.transcription_feed)

        layout.addStretch()
        return panel

    def create_right_panel(self):
        """Create the right panel with live analytics."""
        panel = QFrame()
        panel.setObjectName("rightPanel")
        panel.setFixedWidth(230)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # Analytics title
        analytics_title = QLabel("Live Analytics")
        analytics_title.setObjectName("sectionTitle")
        layout.addWidget(analytics_title)

        layout.addSpacing(10)

        # Slide Progress
        slide_progress_container = self.create_analytics_item(
            "Slide Progress",
            "0 / 0",
            0,  # No progress initially
            "circular",
            dracula_theme.accent_primary
        )
        layout.addWidget(slide_progress_container)

        # Speaking Pace
        speaking_pace_container = self.create_analytics_item(
            "Speaking Pace",
            "0 WPM",
            0,  # No speaking yet
            "semicircle",
            dracula_theme.accent_success
        )
        layout.addWidget(speaking_pace_container)

        # Filler Words
        filler_words_container = self.create_analytics_item(
            "Filler Words",
            'Count: 0',
            0,  # No filler words yet
            "bar",
            dracula_theme.accent_error
        )
        layout.addWidget(filler_words_container)

        layout.addStretch()
        return panel

    def create_analytics_item(self, title, subtitle, value, chart_type, color):
        """Create an analytics item with chart."""
        container = QFrame()
        layout = QVBoxLayout(container)
        layout.setSpacing(10)

        # Title
        title_label = QLabel(title)
        title_label.setObjectName("analyticsLabel")
        layout.addWidget(title_label)

        if chart_type == "circular":
            # Circular progress (like slide progress)
            self.slide_progress = CircularProgress()
            self.slide_progress.setValue(value)
            self.slide_progress.setMaximum(100)
            self.slide_progress.setColor(color)
            self.slide_progress.setFixedSize(70, 70)

            chart_container = QHBoxLayout()
            chart_container.addWidget(self.slide_progress)

            # Text beside the chart
            text_container = QVBoxLayout()
            self.slide_progress_text = QLabel(subtitle)
            self.slide_progress_text.setObjectName("progressText")
            text_container.addWidget(self.slide_progress_text)
            chart_container.addLayout(text_container)

            layout.addLayout(chart_container)

        elif chart_type == "semicircle":
            # Semi-circular progress (like speaking pace)
            chart_container = QHBoxLayout()

            # Create a simple circular indicator
            self.speaking_pace = CircularProgress()
            self.speaking_pace.setValue(value)
            self.speaking_pace.setMaximum(100)
            self.speaking_pace.setColor(color)
            self.speaking_pace.setFixedSize(50, 50)
            chart_container.addWidget(self.speaking_pace)

            # Text beside the chart
            text_container = QVBoxLayout()
            self.speaking_pace_text = QLabel(subtitle)
            self.speaking_pace_text.setObjectName("progressText")
            text_container.addWidget(self.speaking_pace_text)
            chart_container.addLayout(text_container)

            layout.addLayout(chart_container)

        elif chart_type == "bar":
            # Progress bar (like filler words)
            self.filler_words_progress = QProgressBar()
            self.filler_words_progress.setObjectName("fillerWordsProgress")
            self.filler_words_progress.setValue(value)
            self.filler_words_progress.setMaximum(100)
            self.filler_words_progress.setFixedHeight(8)
            layout.addWidget(self.filler_words_progress)

            self.filler_words_text = QLabel(subtitle)
            self.filler_words_text.setObjectName("fillerWordsText")
            layout.addWidget(self.filler_words_text)

        return container

    def update_cue_items(self, cue_texts):
        """Update the speaker cue items with new text."""
        try:
            # Ensure we have the same number of cue items as texts
            max_items = max(len(cue_texts), len(self.cue_items))

            for i in range(max_items):
                if i < len(cue_texts) and i < len(self.cue_items):
                    # Update existing cue item
                    cue_text = cue_texts[i].strip()
                    if cue_text.startswith('•'):
                        cue_text = cue_text[1:].strip()  # Remove bullet if present

                    # Find the cue text label in the cue item
                    cue_item = self.cue_items[i]
                    layout = cue_item.layout()
                    if layout and layout.count() > 1:
                        text_label = layout.itemAt(1).widget()
                        if isinstance(text_label, QLabel):
                            text_label.setText(cue_text)
        except Exception as e:
            print(f"Error updating cue items: {e}")

    def apply_styles(self):
        """Apply the modern stylesheet to match the SVG design."""
        self.setStyleSheet(f"""
            /* Main panels */
            QFrame#leftPanel {{
                background-color: {dracula_theme.bg_secondary};
                border-radius: 12px;
            }}

            QFrame#centerPanel {{
                background-color: rgba(40, 42, 54, 0.7);
                border: 1px solid rgba(68, 71, 90, 0.5);
                border-radius: 12px;
                margin: 0 10px;
            }}

            QFrame#rightPanel {{
                background-color: {dracula_theme.bg_secondary};
                border-radius: 12px;
            }}

            /* Presentation title */
            QLabel#presentationTitle {{
                color: {dracula_theme.accent_primary};
                font-size: 16px;
                font-weight: bold;
                margin: 0;
                padding: 5px 0;
                border-bottom: 1px solid {getattr(dracula_theme, 'border_color', '#44475a')};
            }}

            /* Section titles */
            QLabel#sectionTitle {{
                color: {dracula_theme.accent_secondary};
                font-size: 14px;
                font-weight: bold;
                margin: 0;
                padding: 0;
            }}

            /* Subsection titles */
            QLabel#subsectionTitle {{
                color: {dracula_theme.accent_primary};
                font-size: 14px;
                font-weight: bold;
                margin: 0;
                padding: 0;
            }}

            /* Analytics labels */
            QLabel#analyticsLabel {{
                color: {dracula_theme.text_primary};
                font-size: 14px;
                font-weight: normal;
                margin: 0;
                padding: 0;
            }}

            /* Progress text */
            QLabel#progressText {{
                color: {dracula_theme.text_secondary};
                font-size: 12px;
                font-weight: normal;
                margin: 0;
                padding: 0;
            }}

            /* Next slide label */
            QLabel#nextSlideLabel {{
                color: {dracula_theme.text_secondary};
                font-size: 12px;
                font-weight: bold;
                text-transform: uppercase;
                letter-spacing: 1px;
                margin: 0;
                padding: 0;
            }}

            /* Next slide title */
            QLabel#nextSlideTitle {{
                color: {dracula_theme.text_primary};
                font-size: 14px;
                font-weight: normal;
                margin: 0;
                padding: 0;
            }}

            /* Cue bullets */
            QLabel#cueBullet {{
                color: {dracula_theme.accent_primary};
                font-size: 16px;
                font-weight: bold;
            }}

            /* Cue text */
            QLabel#cueText {{
                color: {dracula_theme.text_primary};
                font-size: 14px;
                font-weight: normal;
            }}

            /* AI suggestion */
            QLabel#aiSuggestion {{
                color: {dracula_theme.text_primary};
                font-size: 16px;
                font-weight: 500;
                line-height: 1.4;
                margin: 0;
                padding: 0;
            }}

            /* Transcription feed */
            QTextEdit#transcriptionFeed {{
                background-color: transparent;
                border: none;
                color: {dracula_theme.text_secondary};
                font-family: "Monaco", "Menlo", "Consolas", monospace;
                font-size: 12px;
                line-height: 1.5;
                padding: 5px;
            }}

            /* Divider lines */
            QFrame#divider {{
                color: {getattr(dracula_theme, 'bg_input_border', '#44475a')};
                background-color: {getattr(dracula_theme, 'bg_input_border', '#44475a')};
            }}

            /* Filler words progress */
            QProgressBar#fillerWordsProgress {{
                background-color: {getattr(dracula_theme, 'bg_input_border', '#44475a')};
                border: none;
                border-radius: 4px;
                text-align: center;
            }}

            QProgressBar#fillerWordsProgress::chunk {{
                background-color: {dracula_theme.accent_error};
                border-radius: 4px;
            }}

            /* Filler words text */
            QLabel#fillerWordsText {{
                color: {dracula_theme.text_secondary};
                font-size: 12px;
                margin: 0;
                padding: 0;
            }}
        """)

    def showEvent(self, event):
        """Apply styles when the view is shown."""
        super().showEvent(event)
        self.apply_styles()

    def setup_view(self):
        """Called by MainWindow right before this view is shown."""
        try:
            # Apply the styles
            self.apply_styles()

            # Reset AI metrics
            if hasattr(self.ai_service, 'reset_metrics'):
                self.ai_service.reset_metrics()

            # Connect service callbacks only once to avoid duplicates
            if not self.callbacks_connected and hasattr(self.pres_service, 'add_slide_change_callback'):
                self.pres_service.add_slide_change_callback(self.update_slide_info)
                self.callbacks_connected = True

            if hasattr(self.ai_service, 'add_transcription_callback'):
                self.ai_service.add_transcription_callback(self.update_transcription)
            if hasattr(self.ai_service, 'add_assistance_callback'):
                self.ai_service.add_assistance_callback(self.update_assistance)
            if hasattr(self.ai_service, 'add_analysis_callback'):
                self.ai_service.add_analysis_callback(self.update_analytics_from_analysis)

            # Start AI listening service for live transcription
            if hasattr(self.ai_service, 'start_listening'):
                try:
                    if self.ai_service.start_listening():
                        print("✅ Started live AI transcription service")
                        self.transcription_feed.setPlainText("🎤 Listening for speech... Say something to test transcription.")
                    else:
                        print("⚠️ Failed to start AI listening service")
                        self.transcription_feed.setPlainText("⚠️ Audio transcription not available. Please check microphone permissions.")
                except Exception as e:
                    print(f"Error starting AI service: {e}")
                    self.transcription_feed.setPlainText(f"⚠️ Error starting transcription: {e}")

            # Start real-time updates timer
            self.start_real_time_updates()

            # Get initial slide info safely
            self.refresh_presentation_data()

        except Exception as e:
            print(f"Error in setup_view: {e}")

    def start_real_time_updates(self):
        """Start timer for real-time presentation status updates."""
        if not self.slide_progress_timer:
            self.slide_progress_timer = QTimer()
            self.slide_progress_timer.timeout.connect(self.refresh_presentation_data)
            self.slide_progress_timer.start(1500)  # Update every 1.5 seconds for better responsiveness

    def stop_real_time_updates(self):
        """Stop real-time updates timer."""
        if self.slide_progress_timer:
            self.slide_progress_timer.stop()
            self.slide_progress_timer = None

    def refresh_presentation_data(self):
        """Safely refresh presentation data and update UI."""
        try:
            if not hasattr(self.pres_service, 'get_current_slide_info'):
                return

            # Try to sync with PowerPoint first for real-time updates
            if hasattr(self.pres_service, 'sync_with_powerpoint'):
                self.pres_service.sync_with_powerpoint()

            # Get current presentation status
            if hasattr(self.pres_service, 'get_presentation_summary'):
                summary = self.pres_service.get_presentation_summary()
                if summary and summary.get('presentation_id'):
                    # Update presentation title
                    presentation_id = summary.get('presentation_id', '')
                    if presentation_id:
                        # Clean up the presentation ID for display
                        display_name = presentation_id.replace('_', ' ').title()
                        self.presentation_title.setText(display_name)
                    else:
                        self.presentation_title.setText("Unknown Presentation")

                    # Update slide progress from summary
                    current_slide = summary.get('current_slide', 1)
                    total_slides = summary.get('total_slides', 1)

                    # Update progress indicators
                    self.slide_progress_text.setText(f"{current_slide} / {total_slides}")
                    if total_slides > 0:
                        progress_percent = int((current_slide / total_slides) * 100)
                        self.slide_progress.setValue(progress_percent)

            # Get detailed slide info for speaker cues
            slide_info = self.pres_service.get_current_slide_info()
            if slide_info:
                current_slide = slide_info.get('slide_number', 1)
                total_slides = slide_info.get('total_slides', 1)

                # Update UI with safe data
                self.update_slide_info_safe(current_slide, total_slides, slide_info)

        except Exception as e:
            print(f"Error refreshing presentation data: {e}")

    def update_slide_info_safe(self, current_slide, total_slides, slide_info):
        """Safely update slide information without errors."""
        try:
            # Update progress text and bar
            self.slide_progress_text.setText(f"{current_slide} / {total_slides}")

            if total_slides > 0:
                progress_percent = int((current_slide / total_slides) * 100)
                self.slide_progress.setValue(progress_percent)

            # Generate speaker cues with error handling
            slide_content = slide_info.get('text_content', '') if slide_info else ''

            if hasattr(self.ai_service, 'generate_slide_notes') and slide_content:
                try:
                    notes = self.ai_service.generate_slide_notes(slide_content, current_slide)
                    self.speaker_cues.setPlainText(notes)
                except:
                    # Fallback cues
                    self.speaker_cues.setPlainText(f"• Slide {current_slide} key points\n• Important details to emphasize\n• Transition to next section")
            else:
                # Default fallback cues
                self.speaker_cues.setPlainText(f"• Slide {current_slide} key points\n• Important details to emphasize\n• Transition to next section")

            # Next slide card has been removed - no longer updating slide preview

        except Exception as e:
            print(f"Error in update_slide_info_safe: {e}")

    def hideEvent(self, event):
        """Stop updates when view is hidden."""
        self.stop_real_time_updates()

        # Stop AI listening service when leaving live view
        if hasattr(self.ai_service, 'stop_listening'):
            try:
                self.ai_service.stop_listening()
                print("🛑 Stopped AI listening service")
            except Exception as e:
                print(f"Error stopping AI service: {e}")

        super().hideEvent(event)

    # --- SLOTS to update the UI ---

    @Slot(int, int, dict)
    def update_slide_info(self, current_slide, total_slides, slide_info):
        """Update slide information display."""
        self.slide_progress_text.setText(f"{current_slide} / {total_slides}")

        # Update slide progress
        if total_slides > 0:
            progress_percent = int((current_slide / total_slides) * 100)
            self.slide_progress.setValue(progress_percent)

        # Generate speaker cues if AI service is available
        slide_content = slide_info.get('text_content', '') if slide_info else ''

        if hasattr(self.ai_service, 'generate_slide_notes') and slide_content:
            try:
                notes = self.ai_service.generate_slide_notes(slide_content, current_slide)
                self.update_cue_items(notes.split('\n'))
            except:
                # Fallback cues
                self.update_cue_items([
                    f"Slide {current_slide} key points",
                    "Important details to emphasize",
                    "Transition to next section"
                ])
        else:
            # Default fallback cues
            self.update_cue_items([
                f"Slide {current_slide} key points",
                "Important details to emphasize",
                "Transition to next section"
            ])

    @Slot(str, float)
    def update_transcription(self, new_transcription, timestamp):
        """Update live transcription feed (thread-safe)."""
        # Use QMetaObject.invokeMethod to ensure UI updates happen on main thread
        QMetaObject.invokeMethod(
            self,
            "_update_transcription_ui",
            Qt.QueuedConnection,
            Q_ARG(str, new_transcription),
            Q_ARG(float, timestamp)
        )

    @Slot(str, float)
    def _update_transcription_ui(self, new_transcription, timestamp):
        """Internal method to update UI on main thread."""
        try:
            import time
            current_time = time.strftime("%H:%M:%S", time.localtime(timestamp))
            current_text = self.transcription_feed.toPlainText()
            new_line = f'[{current_time}] "{new_transcription}"'

            # Keep only last 20 lines to prevent memory issues
            lines = current_text.split('\n')
            if len(lines) >= 20:
                lines = lines[-19:]  # Keep last 19 lines
                current_text = '\n'.join(lines)

            self.transcription_feed.setPlainText(f"{current_text}\n{new_line}")

            # Auto-scroll to bottom
            scrollbar = self.transcription_feed.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        except Exception as e:
            print(f"Error updating transcription UI: {e}")

    @Slot(str, str, str)
    def update_assistance(self, assistance_text, trigger_type, context):
        """Update AI assistance display (thread-safe)."""
        QMetaObject.invokeMethod(
            self,
            "_update_assistance_ui",
            Qt.QueuedConnection,
            Q_ARG(str, assistance_text),
            Q_ARG(str, trigger_type),
            Q_ARG(str, context)
        )

    @Slot(str, str, str)
    def _update_assistance_ui(self, assistance_text, trigger_type, context):
        """Internal method to update assistance UI on main thread."""
        try:
            self.ai_suggestion.setText(assistance_text)
        except Exception as e:
            print(f"Error updating assistance UI: {e}")

    def update_analytics(self, wpm=None, filler_count=None):
        """Update speaking analytics."""
        if wpm:
            # Update speaking pace (assuming 200 WPM max for percentage calculation)
            pace_percentage = min(int((wpm / 200) * 100), 100)
            self.speaking_pace.setValue(pace_percentage)
            self.speaking_pace_text.setText(f"{wpm} WPM")

        if filler_count is not None:
            # Update filler words (assuming 20 max for percentage calculation)
            filler_percentage = min(int((filler_count / 20) * 100), 100)
            self.filler_words_progress.setValue(filler_percentage)
            # You could customize the filler words shown based on actual detected words
            self.filler_words_text.setText(f'Count: {filler_count} ("Um", "Ah")')

    def update_analytics_from_analysis(self, analysis_data):
        """Update analytics from AI service analysis data."""
        try:
            if 'pace_analysis' in analysis_data:
                pace_data = analysis_data['pace_analysis']
                if isinstance(pace_data, dict) and 'wpm' in pace_data:
                    wpm = pace_data['wpm']
                elif isinstance(pace_data, (int, float)):
                    wpm = pace_data
                else:
                    wpm = None

                if wpm:
                    pace_percentage = min(int((wpm / 200) * 100), 100)
                    self.speaking_pace.setValue(pace_percentage)
                    self.speaking_pace_text.setText(f"{int(wpm)} WPM")

            if 'filler_analysis' in analysis_data:
                filler_data = analysis_data['filler_analysis']
                if isinstance(filler_data, dict) and 'count' in filler_data:
                    filler_count = filler_data['count']
                elif isinstance(filler_data, (int, float)):
                    filler_count = filler_data
                else:
                    filler_count = None

                if filler_count is not None:
                    filler_percentage = min(int((filler_count / 20) * 100), 100)
                    self.filler_words_progress.setValue(filler_percentage)
                    self.filler_words_text.setText(f'Count: {int(filler_count)} ("Um", "Ah")')

        except Exception as e:
            print(f"Error updating analytics from analysis: {e}")