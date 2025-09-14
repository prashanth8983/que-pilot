"""
Main application window for AI Presentation Copilot.
"""

import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QMessageBox, QButtonGroup
)
from PySide6.QtCore import (
    Signal, Slot, QPoint, Qt, QSize
)
from PySide6.QtGui import QFont, QIcon, QPixmap, QPainter

# (Your service, config, and view imports remain the same)
# ...
# Import services - handle both relative and absolute imports
try:
    from ..services import PresentationService, AIService, SyncService
except ImportError:
    # Fallback for when running from different contexts
    import sys
    from pathlib import Path
    services_path = Path(__file__).parent.parent / "services"
    sys.path.insert(0, str(services_path))
    try:
        from presentation_service import PresentationService
        from ai_service import AIService
        from sync_service import SyncService
    except ImportError:
        # Create placeholder services if imports fail
        class PresentationService:
            def __init__(self): pass
            def load_presentation(self, path): return True
            def start_presentation(self): pass
            def stop_presentation(self): pass
            def add_slide_change_callback(self, callback): pass
            def add_presentation_load_callback(self, callback): pass
            def get_current_slide_info(self): return {'current_slide': 1, 'total_slides': 1}
        
        class AIService:
            def __init__(self): pass
            def start_listening(self): pass
            def stop_listening(self): pass
            def add_transcription_callback(self, callback): pass
            def add_assistance_callback(self, callback): pass
            def reset_metrics(self): pass
            def generate_slide_notes(self, content, slide_num): return f"• Slide {slide_num} content\n• Key points to highlight"
        
        class SyncService:
            def __init__(self): pass
            def start_sync(self): pass
            def stop_sync(self): pass
            def add_slide_sync_callback(self, callback): pass

# Import config - handle both relative and absolute imports
try:
    from ...config.settings import settings, dracula_theme
except ImportError:
    # Fallback for when running from different contexts
    import sys
    from pathlib import Path
    config_path = Path(__file__).parent.parent.parent / "config"
    sys.path.insert(0, str(config_path))
    try:
        from settings import settings, dracula_theme
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

# Import your new view modules
from .views.welcome_view import WelcomeView
from .views.plan_view import PlanView
from .views.live_view import LiveView

# Import the icon loader
try:
    from ..core.utils.icon_loader import load_icon
except ImportError:
    # Fallback if icon loader fails
    def load_icon(name, size=None, color=None):
        return QIcon(f"assets/icons/{name}.svg")
    
    def create_colored_icon(name, color, size=24):
        """Create a colored icon from SVG."""
        try:
            icon_path = f"assets/icons/{name}.svg"
            pixmap = QPixmap(size, size)
            pixmap.fill(Qt.transparent)
            
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # Load the original icon
            original_icon = QIcon(icon_path)
            original_pixmap = original_icon.pixmap(size, size)
            
            # Create a colored version by compositing
            painter.drawPixmap(0, 0, original_pixmap)
            painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
            painter.fillRect(pixmap.rect(), color)
            painter.end()
            
            return QIcon(pixmap)
        except Exception:
            # Fallback to default icon
            return QIcon("assets/icons/help-circle.svg")


class MainWindow(QMainWindow):
    """Main application window with modular view structure."""

    def __init__(self):
        super().__init__()
        
        # Initialize services
        self.ai_service = AIService()
        self.pres_service = PresentationService()
        self.sync_service = SyncService()

        # Connect services
        self.sync_service.set_presentation_service(self.pres_service)
        
        self.setWindowTitle(settings.app_name)
        self.setGeometry(100, 50, 900, 600)
        self.setMinimumSize(800, 500)

        # Set window flags for custom title bar
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        
        # Attribute for draggable window
        self.drag_pos = None
        
        self.init_ui()
        self.setup_connections()
        self.apply_global_styles()
        
    def init_ui(self):
        """Initialize the user interface."""
        # Main container with rounded corners and shadow
        self.container = QFrame()
        self.container.setObjectName("container")
        self.setCentralWidget(self.container)

        main_v_layout = QVBoxLayout(self.container)
        main_v_layout.setContentsMargins(0, 0, 0, 0)
        main_v_layout.setSpacing(0)

        # Add custom title bar
        self.create_title_bar()
        main_v_layout.addWidget(self.title_bar)

        # Content area
        content_widget = QWidget()
        main_v_layout.addWidget(content_widget)

        main_layout = QHBoxLayout(content_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Left sidebar
        self.create_sidebar()
        main_layout.addWidget(self.sidebar)
        
        # Right content area (views + status bar)
        right_content_widget = QWidget()
        right_content_layout = QVBoxLayout(right_content_widget)
        right_content_layout.setContentsMargins(0, 0, 0, 0)
        right_content_layout.setSpacing(0)
        
        self.stacked_widget = QStackedWidget()
        self.welcome_view = WelcomeView(self.pres_service)
        self.plan_view = PlanView()
        self.live_view = LiveView(self.ai_service, self.pres_service, self.sync_service)
        self.stacked_widget.addWidget(self.welcome_view)
        self.stacked_widget.addWidget(self.plan_view)
        self.stacked_widget.addWidget(self.live_view)
        
        right_content_layout.addWidget(self.stacked_widget)
        
        self.create_status_bar()
        right_content_layout.addWidget(self.status_bar)
        
        main_layout.addWidget(right_content_widget)

    def create_title_bar(self):
        """Create the custom title bar."""
        self.title_bar = QFrame()
        self.title_bar.setObjectName("titleBar")
        self.title_bar.setFixedHeight(50)
        title_layout = QHBoxLayout(self.title_bar)
        title_layout.setContentsMargins(15, 0, 5, 0)
        
        icon_label = QLabel()
        icon_label.setPixmap(load_icon("users", QSize(20, 20), dracula_theme.text_primary).pixmap(20, 20))
        icon_label.setStyleSheet("margin-right: 5px;")
        title_layout.addWidget(icon_label)
        
        title_label = QLabel(settings.app_name)
        title_label.setObjectName("titleLabel")
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        min_btn = QPushButton()
        min_btn.setIcon(load_icon("minus", QSize(16, 16), dracula_theme.text_primary))
        min_btn.clicked.connect(self.showMinimized)

        self.max_btn = QPushButton()
        self.max_btn.setIcon(load_icon("square", QSize(16, 16), dracula_theme.text_primary))
        self.max_btn.clicked.connect(self.toggle_maximize)

        close_btn = QPushButton()
        close_btn.setIcon(load_icon("x", QSize(16, 16), dracula_theme.text_primary))
        close_btn.clicked.connect(self.close)

        for btn, hover_color in [(min_btn, dracula_theme.accent_warning),
                                 (self.max_btn, dracula_theme.accent_success),
                                 (close_btn, dracula_theme.accent_error)]:
            btn.setFixedSize(30, 30)
            btn.setObjectName("windowControlBtn")
            # Store hover color in a dynamic property
            btn.setProperty("hoverColor", hover_color)
            title_layout.addWidget(btn)

    def create_sidebar(self):
        """Create the left sidebar navigation."""
        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(80)
        
        layout = QVBoxLayout(self.sidebar)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        layout.setSpacing(15)
        
        self.nav_group = QButtonGroup()
        self.nav_group.setExclusive(True)

        # Create navigation buttons with colored icons
        self.home_btn = self.create_nav_button("home", "Home", 0, self.show_welcome_view, True)
        self.plan_btn = self.create_nav_button("check", "Plan", 2, self.show_plan_view)
        self.live_session_btn = self.create_nav_button("mic", "Live", 1, self.show_live_view)
        
        layout.addWidget(self.home_btn)
        layout.addWidget(self.plan_btn)
        layout.addWidget(self.live_session_btn)

    def create_nav_button(self, icon_name, tooltip, btn_id, func, is_checked=False):
        """Helper to create a sidebar navigation button with colored icons."""
        btn = QPushButton()

        # Create colored icons for different states
        # Unselected: white color, Selected: same color as sidebar background
        normal_icon = load_icon(icon_name, QSize(24, 24), dracula_theme.text_primary)  # White for unselected
        hover_icon = load_icon(icon_name, QSize(24, 24), dracula_theme.text_primary)   # White on hover
        checked_icon = load_icon(icon_name, QSize(24, 24), dracula_theme.bg_secondary)  # Same as sidebar background when selected

        # Set the initial icon
        if is_checked:
            btn.setIcon(checked_icon)
        else:
            btn.setIcon(normal_icon)

        btn.setFixedSize(50, 50)
        btn.setCheckable(True)
        btn.setChecked(is_checked)
        btn.setToolTip(tooltip)
        btn.clicked.connect(func)

        # Store icon names for dynamic updates
        btn.icon_name = icon_name
        btn.normal_icon = normal_icon
        btn.hover_icon = hover_icon
        btn.checked_icon = checked_icon

        # Connect to state changes to update icon
        btn.toggled.connect(lambda checked: self.update_button_icon(btn, checked))

        self.nav_group.addButton(btn, btn_id)
        return btn
    
    def update_button_icon(self, btn, checked):
        """Update button icon based on state."""
        if checked:
            btn.setIcon(btn.checked_icon)
        else:
            btn.setIcon(btn.normal_icon)

    def create_status_bar(self):
        """Create the status bar."""
        self.status_bar = QFrame()
        self.status_bar.setObjectName("statusBar")
        self.status_bar.setFixedHeight(30)
        
        layout = QHBoxLayout(self.status_bar)
        layout.setContentsMargins(20, 0, 20, 0)
        
        self.status_text = QLabel("STATUS: Idle")
        self.status_text.setObjectName("statusText")
        layout.addWidget(self.status_text)
        layout.addStretch()
        
        self.stop_btn = QPushButton("Stop Session")
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.clicked.connect(self.stop_presentation)
        self.stop_btn.hide()
        layout.addWidget(self.stop_btn)
        
        version_label = QLabel(f"v{settings.app_version}")
        version_label.setObjectName("versionLabel")
        layout.addWidget(version_label)
        
    def setup_connections(self):
        self.welcome_view.navigate_to_plan.connect(self.show_plan_view)
        self.welcome_view.navigate_to_live.connect(lambda: self.show_live_view())  # Call without parameters
        self.welcome_view.navigate_to_live_with_plan.connect(self.show_live_view)
        self.welcome_view.auto_detect_requested.connect(self.handle_auto_detect)
        self.plan_view.presentation_started.connect(self.show_live_view)

    # --- DRAGGABLE WINDOW LOGIC ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Check if the click is within the title bar's geometry
            if self.title_bar.geometry().contains(event.pos()):
                self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()

    def mouseMoveEvent(self, event):
        if self.drag_pos and event.buttons() == Qt.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.drag_pos = None
        event.accept()
    
    def toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
            self.max_btn.setIcon(load_icon("square", QSize(16, 16), dracula_theme.text_primary))
        else:
            self.showMaximized()
            self.max_btn.setIcon(load_icon("minimize-2", QSize(16, 16), dracula_theme.text_primary))

    # --- Navigation and State Methods ---
    @Slot()
    def show_welcome_view(self):
        self.stacked_widget.setCurrentWidget(self.welcome_view)
        self.home_btn.setChecked(True)
        self.stop_btn.hide()
        self.update_status("Idle")
        
    @Slot()
    def show_plan_view(self):
        self.stacked_widget.setCurrentWidget(self.plan_view)
        self.plan_btn.setChecked(True)
        self.stop_btn.hide()
        self.update_status("Ready to Plan")
        
    @Slot()
    @Slot(str)
    def show_live_view(self, file_path=None):
        should_start = False

        if file_path:
            # Try to load the specified presentation
            if self.pres_service.load_presentation(file_path):
                should_start = True
            else:
                QMessageBox.warning(self, "Error", f"Could not load presentation: {file_path}")
                return
        elif self.pres_service.current_presentation_id:
            # Allow navigating to live view if a presentation is already loaded
            should_start = True
        else:
            # No presentation loaded - offer options to the user
            reply = QMessageBox.question(
                self,
                "No Presentation Loaded",
                "No presentation is currently loaded. Would you like to:\n\n"
                "• Click 'Yes' to auto-detect a running PowerPoint presentation\n"
                "• Click 'No' to go back and load a presentation file first",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                # Try auto-detection
                self.update_status("Detecting presentation...")
                if self.pres_service.auto_detect_presentation():
                    should_start = True
                    self.update_status("Presentation detected!")
                else:
                    QMessageBox.information(
                        self,
                        "Auto-Detection Failed",
                        "Could not detect a running PowerPoint presentation.\n\n"
                        "Please ensure:\n"
                        "• PowerPoint is running\n"
                        "• A presentation is open\n"
                        "• The presentation is not minimized\n\n"
                        "Or go back and load a presentation file manually."
                    )
                    return
            else:
                # User chose to go back
                return

        if should_start:
            self.pres_service.start_presentation()

            # Start sync service for real-time slide tracking
            self.sync_service.start_sync(1.5)  # Check every 1.5 seconds

            self.live_view.setup_view()  # This will start the AI service with proper error handling
            self.stacked_widget.setCurrentWidget(self.live_view)
            self.live_session_btn.setChecked(True)
            self.stop_btn.show()
            self.update_status("Listening...")
            
    def stop_presentation(self):
        self.ai_service.stop_listening()
        self.pres_service.stop_presentation()
        self.sync_service.stop_sync()  # Stop slide synchronization
        self.update_status("Presentation stopped")
        self.show_welcome_view()

    @Slot()
    def handle_auto_detect(self):
        """Handle automatic presentation detection request."""
        self.update_status("Detecting presentation...")

        try:
            success = self.pres_service.auto_detect_presentation()
            if success:
                self.update_status("Presentation detected! Ready to start live session.")
                # Show live view with the detected presentation
                self.show_live_view()
            else:
                self.update_status("No presentation detected. Please ensure PowerPoint is running.")
                QMessageBox.information(
                    self,
                    "Auto-Detection",
                    "No PowerPoint presentation was detected.\n\n"
                    "Please ensure that:\n"
                    "• PowerPoint is running\n"
                    "• A presentation is open\n"
                    "• The presentation is not minimized"
                )
        except Exception as e:
            self.update_status("Auto-detection failed")
            QMessageBox.warning(self, "Error", f"Auto-detection failed: {str(e)}")

    def update_status(self, text):
        self.status_text.setText(f"STATUS: {text}")
        
    def apply_global_styles(self):
        self.setStyleSheet(f"""
            #container {{
                background-color: {dracula_theme.bg_main};
                border-radius: 10px;
            }}
            #titleBar {{
                background-color: {dracula_theme.bg_secondary};
                border-bottom: 1px solid {dracula_theme.border_color};
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }}
            #titleLabel {{
                color: {dracula_theme.text_primary};
                font-size: 16px; font-weight: 600;
            }}
            #windowControlBtn {{
                background-color: transparent;
                color: {dracula_theme.text_primary};
                border: none;
                border-radius: 0px;
                font-size: 16px;
                font-weight: bold;
                margin: 2px;
            }}
            #windowControlBtn:hover {{
                background-color: {dracula_theme.bg_input_border};
                color: {dracula_theme.text_primary};
            }}
            #sidebar {{
                background-color: {dracula_theme.bg_secondary};
                border-right: 1px solid {dracula_theme.border_color};
            }}
            #sidebar QPushButton {{
                background-color: transparent; border: none;
                border-radius: 10px; font-size: 24px;
            }}
            #sidebar QPushButton:hover {{
                background-color: {dracula_theme.bg_input_border};
            }}
            #sidebar QPushButton:checked {{
                background-color: {dracula_theme.accent_primary};
            }}
            #sidebar QPushButton QIcon {{
                color: {dracula_theme.text_secondary};
            }}
            #sidebar QPushButton:hover QIcon {{
                color: {dracula_theme.text_primary};
            }}
            #sidebar QPushButton:checked QIcon {{
                color: {dracula_theme.text_primary};
            }}
            #statusBar {{
                background-color: {dracula_theme.bg_secondary};
                border-top: 1px solid {dracula_theme.border_color};
                border-bottom-left-radius: 10px;
                border-bottom-right-radius: 10px;
            }}
            #statusText, #versionLabel {{
                color: {dracula_theme.text_secondary};
                font-size: 12px;
            }}
            #stopButton {{
                background-color: {dracula_theme.accent_error};
                color: {dracula_theme.text_primary}; border: none;
                border-radius: 6px; padding: 4px 12px;
                font-size: 12px; font-weight: 600;
            }}
            #stopButton:hover {{ background-color: #FF6B6B; }}
            QToolTip {{
                background-color: {dracula_theme.card_bg}; color: {dracula_theme.text_primary};
                border: 1px solid {dracula_theme.accent_primary}; padding: 4px; border-radius: 4px;
            }}
        """)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
