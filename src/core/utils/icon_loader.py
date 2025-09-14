"""
Icon loading utility for SVG icons.
"""

import os
from pathlib import Path
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtCore import QSize, Qt


class IconLoader:
    """Utility class for loading SVG icons."""

    def __init__(self):
        # Get the project root and assets directory
        self.project_root = Path(__file__).parent.parent.parent.parent
        self.icons_dir = self.project_root / "assets" / "icons"

    def load_svg_icon(self, icon_name: str, size: QSize = QSize(24, 24), color: str = "#f8f8f2") -> QIcon:
        """
        Load an SVG icon and return as QIcon.

        Args:
            icon_name: Name of the icon file (with or without .svg extension)
            size: Size of the icon (default: 24x24)
            color: Color to use for the icon (default: Dracula theme text primary)

        Returns:
            QIcon object, or empty icon if file not found
        """
        if not icon_name.endswith('.svg'):
            icon_name += '.svg'

        icon_path = self.icons_dir / icon_name

        if not icon_path.exists():
            print(f"Warning: Icon not found: {icon_path}")
            return QIcon()

        try:
            # Read SVG content and replace currentColor with specified color
            with open(icon_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()

            # Replace currentColor with the specified color
            svg_content = svg_content.replace('currentColor', color)

            # Create SVG renderer with the modified content
            renderer = QSvgRenderer()
            renderer.load(svg_content.encode('utf-8'))

            # Create pixmap with the desired size
            pixmap = QPixmap(size)
            pixmap.fill(Qt.transparent)  # Fill with transparent background

            # Render SVG onto pixmap
            painter = QPainter(pixmap)
            renderer.render(painter)
            painter.end()

            return QIcon(pixmap)

        except Exception as e:
            print(f"Error loading icon {icon_name}: {e}")
            return QIcon()

    def get_icon_path(self, icon_name: str) -> str:
        """
        Get the full path to an icon file.

        Args:
            icon_name: Name of the icon file (with or without .svg extension)

        Returns:
            Full path to the icon file as string
        """
        if not icon_name.endswith('.svg'):
            icon_name += '.svg'

        return str(self.icons_dir / icon_name)


# Global icon loader instance
icon_loader = IconLoader()


def load_icon(icon_name: str, size: QSize = QSize(24, 24), color: str = "#f8f8f2") -> QIcon:
    """
    Convenience function to load an icon.

    Args:
        icon_name: Name of the icon file (with or without .svg extension)
        size: Size of the icon (default: 24x24)
        color: Color to use for the icon (default: Dracula theme text primary)

    Returns:
        QIcon object
    """
    return icon_loader.load_svg_icon(icon_name, size, color)