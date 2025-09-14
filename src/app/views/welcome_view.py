"""
Welcome view for the AI Presentation Copilot application.
"""
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame
)
from PySide6.QtCore import Signal, Qt, Slot, QSize
from PySide6.QtGui import QIcon

# Import config and widgets with fallbacks
try:
    from config.settings import dracula_theme
except ImportError:
    # Fallback theme
    class DraculaTheme:
        bg_main = "#282a36"
        bg_secondary = "#21222c"
        text_primary = "#f8f8f2"
        text_secondary = "#6272a4"
        accent_primary = "#bd93f9"
    dracula_theme = DraculaTheme()

try:
    from ..widgets import DraculaButton
except ImportError:
    # Fallback button
    from PySide6.QtWidgets import QPushButton
    DraculaButton = QPushButton

try:
    from ...services import PresentationService
except ImportError:
    # Fallback service
    class PresentationService:
        def __init__(self): pass

# Import icon loader
try:
    from ...core.utils.icon_loader import load_icon
except ImportError:
    # Fallback if icon loader fails
    def load_icon(name, size=None):
        return QIcon()

class PlanCard(QFrame):
    """A clickable card widget to display a recent presentation plan."""
    
    # Signal emitted when the card is clicked, carrying the full file path
    clicked = Signal(str)

    def __init__(self, title, last_modified, file_path, parent=None):
        super().__init__(parent)
        self.plan_name = title
        self.file_path = file_path # Store the full file path
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("planCard")
        self.setFixedHeight(70)
        
        # Main layout
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 0, 20, 0)
        layout.setSpacing(15)

        # Icon
        icon_label = QLabel()
        icon_label.setPixmap(load_icon("file-text", QSize(24, 24)).pixmap(24, 24))
        icon_label.setStyleSheet(f"color: {dracula_theme.text_secondary};")
        layout.addWidget(icon_label)

        # Title and subtitle
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        text_layout.addStretch()
        title_label = QLabel(title)
        title_label.setObjectName("planCardTitle")
        subtitle_label = QLabel(last_modified)
        subtitle_label.setObjectName("planCardSubtitle")
        text_layout.addWidget(title_label)
        text_layout.addWidget(subtitle_label)
        text_layout.addStretch()
        
        layout.addLayout(text_layout)
        layout.addStretch()

        # Chevron icon
        chevron_label = QLabel()
        chevron_label.setPixmap(load_icon("chevron-right", QSize(20, 20)).pixmap(20, 20))
        chevron_label.setStyleSheet(f"color: {dracula_theme.text_primary};")
        layout.addWidget(chevron_label)
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet(f"""
            #planCard {{
                background-color: {dracula_theme.bg_secondary};
                border-radius: 10px;
                border: 1px solid transparent;
            }}
            #planCard:hover {{
                border: 1px solid {dracula_theme.accent_primary};
            }}
            #planCardTitle {{
                color: {dracula_theme.text_primary};
                font-size: 16px;
                font-weight: 500;
            }}
            #planCardSubtitle {{
                color: {dracula_theme.text_secondary};
                font-size: 13px;
            }}
        """)

    def mouseReleaseEvent(self, event):
        """Emit the clicked signal with the file path."""
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.file_path)
        super().mouseReleaseEvent(event)


class WelcomeView(QWidget):
    """Welcome page with dynamic recent plans and quick actions."""

    navigate_to_plan = Signal()
    navigate_to_live = Signal()
    navigate_to_live_with_plan = Signal(str)
    auto_detect_requested = Signal()
    
    def __init__(self, pres_service: PresentationService, parent=None):
        super().__init__(parent)
        self.pres_service = pres_service
        self.init_ui()
        self.populate_recent_plans()
        
    def init_ui(self):
        """Initialize the welcome view UI."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(50, 40, 50, 40)
        main_layout.setSpacing(30)
        main_layout.setAlignment(Qt.AlignTop)

        # Welcome Header
        welcome_title = QLabel("Welcome")
        welcome_title.setStyleSheet(f"color: {dracula_theme.text_primary}; font-size: 28px; font-weight: 700;")
        welcome_subtitle = QLabel("Manage your presentation plans and start a live session.")
        welcome_subtitle.setStyleSheet(f"color: {dracula_theme.text_secondary}; font-size: 16px;")
        main_layout.addWidget(welcome_title)
        main_layout.addWidget(welcome_subtitle)
        main_layout.addSpacing(20)

        # Recent Plans Header
        plans_header_layout = QHBoxLayout()
        plans_header_layout.setContentsMargins(0, 0, 0, 10)
        plans_title = QLabel("Recent Plans")
        plans_title.setStyleSheet(f"color: {dracula_theme.text_primary}; font-size: 20px; font-weight: 600;")
        plans_header_layout.addWidget(plans_title)
        plans_header_layout.addStretch()

        # Start assist button
        self.auto_detect_btn = DraculaButton("Start Assist", primary=False)
        self.auto_detect_btn.clicked.connect(self.auto_detect_requested.emit)
        plans_header_layout.addWidget(self.auto_detect_btn)

        self.create_new_plan_btn = DraculaButton("Create New Plan", primary=True)
        self.create_new_plan_btn.clicked.connect(self.navigate_to_plan.emit)
        plans_header_layout.addWidget(self.create_new_plan_btn)
        main_layout.addLayout(plans_header_layout)
        
        # A container for the dynamic plan cards
        self.plans_list_layout = QVBoxLayout()
        self.plans_list_layout.setSpacing(15)
        main_layout.addLayout(self.plans_list_layout)
        main_layout.addStretch()

    def populate_recent_plans(self):
        """Fetch recent plans from the service and display them."""
        # Clear any existing plan cards
        for i in reversed(range(self.plans_list_layout.count())): 
            widget_to_remove = self.plans_list_layout.itemAt(i).widget()
            if widget_to_remove:
                widget_to_remove.setParent(None)

        # In a real app, this method would be in your PresentationService
        recent_plans = self.get_recent_plans_from_disk()

        if not recent_plans:
            no_more_label = QLabel("No recent plans found in the project directory.")
            no_more_label.setAlignment(Qt.AlignCenter)
            no_more_label.setStyleSheet(f"color: {dracula_theme.text_secondary}; font-size: 14px; margin-top: 20px;")
            self.plans_list_layout.addWidget(no_more_label)
        else:
            for plan in recent_plans:
                card = PlanCard(plan['name'], plan['modified'], plan['path'])
                card.clicked.connect(self.load_plan)
                self.plans_list_layout.addWidget(card)

    def get_recent_plans_from_disk(self, count=5):
        """
        Placeholder function to find recent .pptx files.
        This logic should eventually live in your PresentationService.
        """
        from datetime import datetime
        # Scans the project's root directory for .pptx files
        project_root = Path(__file__).parent.parent.parent.parent
        files = list(project_root.glob("*.pptx"))
        
        if not files:
            return []

        # Sort files by modification time, newest first
        files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        
        plans = []
        for f in files[:count]:
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            plans.append({
                'name': f.stem,
                'modified': f"Last modified: {mtime.strftime('%b %d, %Y')}",
                'path': str(f)
            })
        return plans

    @Slot(str)
    def load_plan(self, file_path):
        """Handles the click from a PlanCard."""
        print(f"Loading plan from path: {file_path}")
        self.navigate_to_live_with_plan.emit(file_path)