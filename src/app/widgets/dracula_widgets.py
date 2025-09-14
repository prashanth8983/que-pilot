"""
Custom Dracula-themed widgets for the AI Presentation Copilot.
This file contains a library of reusable, styled widgets.
"""

from PySide6.QtWidgets import (
    QPushButton, QFrame, QStackedWidget, QWidget, QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
    QHBoxLayout, QVBoxLayout, QLabel
)
from PySide6.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve, Property, QPoint, QSize, Signal
)
from PySide6.QtGui import (
    QPainter, QBrush, QPen, QColor, QIcon
)

# Import theme - handle running file directly
try:
    from ...config.settings import dracula_theme
except (ImportError, ValueError):
    print("Warning: Could not import dracula_theme. Using fallback colors.")
    class DraculaTheme:
        bg_main = "#282a36"
        bg_secondary = "#21222c"
        bg_hover = "#383A47"
        bg_input_border = "#44475a"
        card_bg = "#44475a"
        border_color = "#6272a4"
        text_primary = "#f8f8f2"
        text_secondary = "#6272a4"
        accent_primary = "#bd93f9"
        accent_primary_light = "#d6acff"
        accent_primary_dark = "#a47fd9"
        accent_secondary = "#8be9fd"
        accent_success = "#50fa7b"
        accent_warning = "#ffb86c"
        accent_error = "#ff5555"
        disabled_bg = "#44475A"
        disabled_fg = "#6272A4"
    dracula_theme = DraculaTheme()


class DraculaButton(QPushButton):
    """A styled button with primary and secondary variants."""
    
    def __init__(self, text="", primary=False, icon_path=None, parent=None):
        super().__init__(text, parent)
        self.primary = primary
        self.setCursor(Qt.PointingHandCursor)
        
        if icon_path:
            self.setIcon(QIcon(icon_path))
            self.setIconSize(QSize(16, 16))

        self.update_style()
        
    def update_style(self):
        base_style = """
            QPushButton {{
                border: none;
                border-radius: 8px;
                padding: {padding};
                font-size: 14px;
                font-weight: 600;
            }}
        """
        if self.primary:
            style = base_style.format(padding="10px 20px") + f"""
                QPushButton {{
                    background-color: {dracula_theme.accent_primary};
                    color: {dracula_theme.bg_main};
                }}
                QPushButton:hover {{ background-color: {dracula_theme.accent_primary_light}; }}
                QPushButton:pressed {{ background-color: {dracula_theme.accent_primary_dark}; }}
                QPushButton:disabled {{
                    background-color: {dracula_theme.disabled_bg};
                    color: {dracula_theme.disabled_fg};
                }}
            """
        else:
            style = base_style.format(padding="9px 19px") + f"""
                QPushButton {{
                    background-color: {dracula_theme.card_bg};
                    color: {dracula_theme.text_primary};
                    border: 1px solid {dracula_theme.border_color};
                }}
                QPushButton:hover {{
                    background-color: {dracula_theme.bg_hover};
                    border-color: {dracula_theme.accent_primary};
                }}
                QPushButton:pressed {{ background-color: {dracula_theme.bg_hover}; }}
                QPushButton:disabled {{
                    background-color: {dracula_theme.bg_hover};
                    color: {dracula_theme.disabled_fg};
                    border-color: {dracula_theme.disabled_bg};
                }}
            """
        self.setStyleSheet(style)


class DarkCard(QFrame):
    """A basic card widget with background and border."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {dracula_theme.bg_secondary};
                border: 1px solid {dracula_theme.border_color};
                border-radius: 10px;
                padding: 20px;
            }}
        """)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(25)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 70))
        self.setGraphicsEffect(shadow)


class SidebarItem(QPushButton):
    """Sidebar navigation item with icon support."""
    
    def __init__(self, text="", icon_path=None, parent=None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        if icon_path:
            self.setIcon(QIcon(icon_path))
            self.setIconSize(QSize(18, 18))
        
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {dracula_theme.text_secondary};
                border: none;
                border-radius: 6px;
                padding: 12px 20px;
                text-align: left;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {dracula_theme.card_bg};
                color: {dracula_theme.text_primary};
            }}
            QPushButton:checked {{
                background-color: {dracula_theme.card_bg};
                color: {dracula_theme.accent_primary};
                font-weight: 600;
            }}
        """)


class ToggleSwitch(QWidget):
    """An animated toggle switch widget."""
    
    toggled = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(44, 24)
        self.setCursor(Qt.PointingHandCursor)
        self._checked = False
        self._knob_position = 3
        
        self.knob_animation = QPropertyAnimation(self, b"knob_position", self)
        self.knob_animation.setEasingCurve(QEasingCurve.InOutCubic)
        self.knob_animation.setDuration(200)

    def paintEvent(self, event):
        with QPainter(self) as painter:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            bg_rect = self.rect()
            bg_color = QColor(dracula_theme.accent_primary) if self._checked else QColor(dracula_theme.border_color)
            painter.setBrush(QBrush(bg_color))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(bg_rect, 12, 12)

            knob_size = 18
            knob_y = (self.height() - knob_size) / 2
            knob_x = self.property("knob_position")
            painter.setBrush(QBrush(QColor(dracula_theme.text_primary)))
            painter.drawEllipse(int(knob_x), int(knob_y), knob_size, knob_size)
        
    def setChecked(self, checked):
        if self._checked == checked:
            return
        self._checked = checked
        self.knob_animation.setStartValue(self.knob_position)
        self.knob_animation.setEndValue(self.width() - 21 if self._checked else 3)
        self.knob_animation.start()
        self.toggled.emit(self._checked)
        
    def isChecked(self):
        return self._checked

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.setChecked(not self.isChecked())
        super().mousePressEvent(event)
        
    @Property(int)
    def knob_position(self):
        return self._knob_position

    @knob_position.setter
    def knob_position(self, value):
        self._knob_position = value
        self.update()


class AnimatedStackedWidget(QStackedWidget):
    """A QStackedWidget with a fade transition animation."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.animation_duration = 300  # Duration in milliseconds

    def setCurrentIndex(self, index):
        if self.currentIndex() == index:
            return

        current_widget = self.currentWidget()
        next_widget = self.widget(index)

        if not current_widget or not next_widget:
            super().setCurrentIndex(index)
            return

        opacity_effect = QGraphicsOpacityEffect(opacity=0.0)
        next_widget.setGraphicsEffect(opacity_effect)
        
        super().setCurrentIndex(index)
        
        self.animation = QPropertyAnimation(opacity_effect, b"opacity")
        self.animation.setDuration(self.animation_duration)
        self.animation.setStartValue(0.0)
        self.animation.setEndValue(1.0)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.animation.finished.connect(lambda: next_widget.setGraphicsEffect(None))
        self.animation.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)


class CircularProgress(QWidget):
    """A circular progress indicator widget."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(120, 120)
        self._value = 0
        self._max_value = 100
        self._color = QColor(dracula_theme.accent_primary)
        self._bg_color = QColor(dracula_theme.card_bg)
        self._pen_width = 8
        
    def setValue(self, value):
        self._value = max(0, min(value, self._max_value))
        self.update()
        
    def setMaximum(self, max_value):
        self._max_value = max_value
        self.update()
        
    def setColor(self, color):
        self._color = QColor(color)
        self.update()
        
    def paintEvent(self, event):
        with QPainter(self) as painter:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            rect = self.rect().adjusted(self._pen_width // 2, self._pen_width // 2, -self._pen_width // 2, -self._pen_width // 2)
            
            painter.setPen(QPen(self._bg_color, self._pen_width))
            painter.drawEllipse(rect)
            
            painter.setPen(QPen(self._color, self._pen_width, Qt.SolidLine, Qt.RoundCap))
            progress_angle = (self._value / self._max_value) * 360
            painter.drawArc(rect, 90 * 16, int(-progress_angle * 16))


class PlanCard(QFrame):
    """Card widget for displaying presentation plans."""
    
    clicked = Signal()
    
    def __init__(self, title="", subtitle="", parent=None):
        super().__init__(parent)
        self.setObjectName("planCard")
        self.setFixedHeight(80)
        self.setCursor(Qt.PointingHandCursor)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(15)
        
        icon_label = QLabel("📊")
        icon_label.setStyleSheet("font-size: 24px;")
        
        content_layout = QVBoxLayout()
        content_layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setObjectName("planCardTitle")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("planCardSubtitle")
        content_layout.addWidget(title_label)
        content_layout.addWidget(subtitle_label)
        
        arrow_label = QLabel("→")
        arrow_label.setObjectName("planCardArrow")
        
        layout.addWidget(icon_label)
        layout.addLayout(content_layout)
        layout.addStretch()
        layout.addWidget(arrow_label)
        
        self.setStyleSheet(f"""
            #planCard {{
                background-color: {dracula_theme.bg_secondary};
                border: 1px solid {dracula_theme.border_color};
                border-radius: 8px;
            }}
            #planCard:hover {{
                background-color: {dracula_theme.bg_hover};
                border-color: {dracula_theme.accent_primary};
            }}
            #planCardTitle {{
                color: {dracula_theme.text_primary};
                font-size: 16px;
                font-weight: 600;
            }}
            #planCardSubtitle {{
                color: {dracula_theme.text_secondary};
                font-size: 12px;
            }}
            #planCardArrow {{
                color: {dracula_theme.text_secondary};
                font-size: 18px;
                font-weight: bold;
            }}
        """)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class DragDropArea(QFrame):
    """A styled drag-and-drop area for file uploads."""
    
    file_dropped = Signal(str)
    
    def __init__(self, text="Drag & Drop .pptx file here", icon_path=None, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(200)
        self.setObjectName("dragDropArea")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(15)
        
        icon_label = QLabel()
        if icon_path:
            icon_label.setPixmap(QIcon(icon_path).pixmap(QSize(48, 48)))
        else:
            icon_label.setText("📁") # Fallback emoji
            icon_label.setStyleSheet("font-size: 48px;")
        icon_label.setAlignment(Qt.AlignCenter)
        
        text_label = QLabel(text)
        text_label.setAlignment(Qt.AlignCenter)
        text_label.setWordWrap(True)
        text_label.setObjectName("dragDropText")
        
        subtext_label = QLabel("or")
        subtext_label.setAlignment(Qt.AlignCenter)
        subtext_label.setObjectName("dragDropSubtext")
        
        browse_button = DraculaButton("Browse Files", primary=False)
        
        layout.addWidget(icon_label)
        layout.addWidget(text_label)
        layout.addWidget(subtext_label)
        layout.addWidget(browse_button, 0, Qt.AlignCenter)

        self.setStyleSheet(f"""
            #dragDropArea {{
                background-color: {dracula_theme.bg_secondary};
                border: 2px dashed {dracula_theme.border_color};
                border-radius: 12px;
            }}
            #dragDropArea:hover {{
                border-color: {dracula_theme.accent_primary};
                background-color: {dracula_theme.bg_hover};
            }}
            #dragDropText {{
                color: {dracula_theme.text_primary};
                font-size: 16px;
                font-weight: 600;
            }}
            #dragDropSubtext {{
                color: {dracula_theme.text_secondary};
                font-size: 14px;
            }}
        """)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            # You might want to filter for .pptx here
            self.file_dropped.emit(files[0])


class CustomTitleBar(QWidget):
    """Custom title bar for frameless window."""

    def __init__(self, title="Application", parent=None):
        super().__init__(parent)
        self.setFixedHeight(40)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 5, 0)
        layout.setSpacing(10)

        icon_label = QLabel()
        icon_label.setFixedSize(32, 32)
        icon_pixmap = QIcon("assets/icons/star.svg").pixmap(32, 32)
        icon_label.setPixmap(icon_pixmap)
        icon_label.setStyleSheet("font-size: 20px;")
        
        title_label = QLabel(title)
        title_label.setObjectName("titleBarLabel")

        self.min_btn = self.create_control_button("−", dracula_theme.accent_warning)
        self.max_btn = self.create_control_button("□", dracula_theme.accent_success)
        self.close_btn = self.create_control_button("×", dracula_theme.accent_error)

        layout.addWidget(icon_label)
        layout.addWidget(title_label)
        layout.addStretch()
        layout.addWidget(self.min_btn)
        layout.addWidget(self.max_btn)
        layout.addWidget(self.close_btn)

        self.setObjectName("customTitleBar")
        self.setStyleSheet(f"""
            #customTitleBar {{
                background-color: {dracula_theme.bg_secondary};
                border-bottom: 1px solid {dracula_theme.border_color};
            }}
            #titleBarLabel {{
                color: {dracula_theme.text_primary};
                font-size: 14px;
                font-weight: 600;
            }}
        """)

    def create_control_button(self, text, color):
        btn = QPushButton(text)
        btn.setFixedSize(30, 30)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {dracula_theme.text_secondary};
                border: none;
                border-radius: 15px;
                font-size: 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {color};
                color: {dracula_theme.bg_main};
            }}
        """)
        return btn