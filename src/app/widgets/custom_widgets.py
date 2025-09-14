"""
Custom UI widgets for the AI Presentation Copilot application.
"""
import os
from PySide6.QtWidgets import QFrame, QLabel, QWidget, QVBoxLayout, QHBoxLayout
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QIcon, QPainter, QColor, QPen

class PlanCard(QFrame):
    """A clickable card for displaying a recent plan."""
    clicked = Signal()

    def __init__(self, title, subtitle):
        super().__init__()
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(70)
        self.setObjectName("PlanCard")
        self.setStyleSheet("""
            #PlanCard { background-color: #21222c; border-radius: 10px; }
            #PlanCard:hover { background-color: #44475a; }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 0, 20, 0)
        icon = QLabel()
        icon.setPixmap(QIcon("assets/icons/file.svg").pixmap(24, 24))
        layout.addWidget(icon)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        title_label = QLabel(title)
        title_label.setStyleSheet("color: #f8f8f2; font-size: 16px;")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setStyleSheet("color: #6272a4; font-size: 12px;")
        text_layout.addWidget(title_label)
        text_layout.addWidget(subtitle_label)
        layout.addLayout(text_layout)
        layout.addStretch()

        arrow_icon = QLabel()
        arrow_icon.setPixmap(QIcon("assets/icons/chevron-right.svg").pixmap(24, 24))
        layout.addWidget(arrow_icon)

    def mouseReleaseEvent(self, event):
        self.clicked.emit()
        super().mouseReleaseEvent(event)

class DragDropArea(QLabel):
    """A label that accepts drag and drop files."""
    file_dropped = Signal(str)

    def __init__(self, text):
        super().__init__(text)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #44475a; border-radius: 12px;
                color: #6272a4; font-size: 14px;
            }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(self.styleSheet() + "border-color: #bd93f9;")

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.styleSheet().replace("border-color: #bd93f9;", ""))

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            file_path = urls[0].toLocalFile()
            self.file_dropped.emit(file_path)
            self.setText(f"Loaded: {os.path.basename(file_path)}")
        self.setStyleSheet(self.styleSheet().replace("border-color: #bd93f9;", "") + "border-color: #50fa7b;")

class CircularProgress(QWidget):
    """A circular progress bar widget."""
    def __init__(self):
        super().__init__()
        self.value = 0
        self.color = QColor("#bd93f9")
        self.setFixedSize(60, 60)

    def setValue(self, value):
        self.value = value
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(5, 5, -5, -5)
        
        painter.setPen(QPen(QColor("#44475a"), 6, Qt.SolidLine, Qt.RoundCap))
        painter.drawArc(rect, 90 * 16, 360 * 16)
        
        painter.setPen(QPen(self.color, 6, Qt.SolidLine, Qt.RoundCap))
        span_angle = int(self.value * 3.6 * 16)
        painter.drawArc(rect, 90 * 16, -span_angle)
