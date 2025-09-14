"""
Plan view for the AI Presentation Copilot application.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QFileDialog, QMessageBox
)
from PySide6.QtCore import Signal, Qt
import os
import json
from pathlib import Path

from ..widgets import DraculaButton, DragDropArea


class PlanView(QWidget):
    """Plan page view for loading presentations and scripts."""
    
    # Signal to notify the main window to start and switch to the live view
    presentation_started = Signal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_file_path = None
        self.init_ui()
        
    def init_ui(self):
        """Initialize the plan view UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(25)

        # Title
        title = QLabel("Plan Your Presentation")
        title.setStyleSheet("""
            color: #f8f8f2;
            font-size: 24px;
            font-weight: 700;
            margin-bottom: 10px;
        """)
        layout.addWidget(title)
        
        main_h_layout = QHBoxLayout()
        main_h_layout.setSpacing(30)
        
        # Left section - Load Presentation File
        left_section = QWidget()
        left_layout = QVBoxLayout(left_section)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(20)
        
        # Section 1: Load Presentation File
        section1_title = QLabel("1. Load Presentation File")
        section1_title.setStyleSheet("""
            color: #bd93f9;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
        """)
        left_layout.addWidget(section1_title)
        
        self.drag_drop_area = DragDropArea("Drag & Drop .pptx file here")
        self.drag_drop_area.file_dropped.connect(self.set_file_path)
        left_layout.addWidget(self.drag_drop_area)
        
        browse_btn = DraculaButton("Browse Files", primary=False)
        browse_btn.clicked.connect(self.browse_presentation)
        browse_btn.setFixedWidth(150)
        left_layout.addWidget(browse_btn, 0, Qt.AlignCenter)
        
        left_layout.addStretch()
        
        # Right section - Add Speaker Script
        right_section = QWidget()
        right_layout = QVBoxLayout(right_section)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(20)
        
        # Section 2: Add Speaker Script
        section2_title = QLabel("2. Add Speaker Script (Optional)")
        section2_title.setStyleSheet("""
            color: #bd93f9;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
        """)
        right_layout.addWidget(section2_title)
        
        self.script_input = QTextEdit()
        self.script_input.setPlaceholderText("Paste your script here to get real-time cues...")
        self.script_input.setStyleSheet("""
            QTextEdit {
                background-color: #44475a;
                border: 1px solid #6272a4;
                border-radius: 8px;
                padding: 15px;
                color: #f8f8f2;
                font-size: 13px;
            }
            QTextEdit:focus { border-color: #bd93f9; }
        """)
        right_layout.addWidget(self.script_input)
        
        right_layout.addStretch()
        
        main_h_layout.addWidget(left_section, 1)
        main_h_layout.addWidget(right_section, 1)
        
        layout.addLayout(main_h_layout)
        
        # Action buttons at bottom
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)
        
        self.save_plan_btn = DraculaButton("Save Plan", primary=False)
        self.save_plan_btn.clicked.connect(self.save_plan)
        action_layout.addWidget(self.save_plan_btn)
        
        self.start_presentation_btn = DraculaButton("Start Presentation", primary=True)
        self.start_presentation_btn.clicked.connect(self.on_start_clicked)
        action_layout.addWidget(self.start_presentation_btn)
        
        layout.addLayout(action_layout)
        layout.addStretch()
        
    def set_file_path(self, path):
        """Set the current file path from drag & drop."""
        if path.endswith(('.pptx', '.ppt')):
            self._current_file_path = path
            # Update the drag drop area to show the loaded file
            self.drag_drop_area.set_file_loaded(os.path.basename(path))
        else:
            # Show error or reset
            self._current_file_path = None
            
    def browse_presentation(self):
        """Browse for presentation file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Presentation", "", "PowerPoint files (*.pptx *.ppt)"
        )
        if file_path:
            self.set_file_path(file_path)
            
    def save_plan(self):
        """Save the current plan."""
        try:
            if not self._current_file_path:
                QMessageBox.warning(
                    self,
                    "No Presentation",
                    "Please load a presentation file before saving the plan."
                )
                return

            # Create a plan data structure
            plan_data = {
                "presentation_file": self._current_file_path,
                "speaker_script": self.script_input.toPlainText(),
                "created_at": str(Path(self._current_file_path).stat().st_mtime),
                "plan_name": Path(self._current_file_path).stem
            }

            # Save plan as JSON file with same name as presentation
            plan_file_name = f"{Path(self._current_file_path).stem}_plan.json"
            project_root = Path(__file__).parent.parent.parent.parent
            plans_dir = project_root / "plans"
            plans_dir.mkdir(exist_ok=True)

            plan_file_path = plans_dir / plan_file_name

            with open(plan_file_path, 'w', encoding='utf-8') as f:
                json.dump(plan_data, f, indent=2, ensure_ascii=False)

            QMessageBox.information(
                self,
                "Plan Saved",
                f"Presentation plan saved successfully!\n\nSaved to: {plan_file_path}"
            )
            print(f"✅ Plan saved: {plan_file_path}")

        except Exception as e:
            QMessageBox.critical(
                self,
                "Save Error",
                f"Failed to save plan: {str(e)}"
            )
            print(f"❌ Error saving plan: {e}")
        
    def on_start_clicked(self):
        """Handle start presentation button click."""
        if self._current_file_path:
            self.presentation_started.emit(self._current_file_path)
        else:
            QMessageBox.warning(
                self,
                "No Presentation",
                "Please load a presentation file before starting."
            )
