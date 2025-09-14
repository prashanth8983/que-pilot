"""
PowerPoint Format Converter

Handles conversion between .ppt and .pptx formats to improve compatibility.
"""

import logging
import subprocess
import platform
from pathlib import Path
from typing import Optional, Union

class PPTConverter:
    """Service for converting between PowerPoint formats."""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.system = platform.system()

    def convert_ppt_to_pptx(self, ppt_path: Union[str, Path], output_dir: Optional[Path] = None) -> Optional[Path]:
        """Convert .ppt file to .pptx format."""
        ppt_path = Path(ppt_path)

        if not ppt_path.exists():
            self.logger.error(f"Source file not found: {ppt_path}")
            return None

        if ppt_path.suffix.lower() != '.ppt':
            self.logger.error(f"Source file is not .ppt format: {ppt_path}")
            return None

        # Set output directory and filename
        if output_dir is None:
            output_dir = ppt_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)

        output_path = output_dir / f"{ppt_path.stem}.pptx"

        self.logger.info(f"Converting {ppt_path} to {output_path}")

        # Try different conversion methods based on platform
        if self.system == "Darwin":  # macOS
            return self._convert_macos_applescript(ppt_path, output_path)
        elif self.system == "Windows":
            return self._convert_windows_com(ppt_path, output_path)
        else:
            return self._convert_libreoffice(ppt_path, output_path)

    def _convert_macos_applescript(self, ppt_path: Path, output_path: Path) -> Optional[Path]:
        """Convert using AppleScript on macOS."""
        try:
            applescript = f'''
            tell application "Microsoft PowerPoint"
                open POSIX file "{ppt_path}"
                delay 2
                save active presentation in POSIX file "{output_path}" as save as PowerPoint presentation
                close active presentation
            end tell
            '''

            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=60
            )

            if result.returncode == 0 and output_path.exists():
                self.logger.info(f"Successfully converted to {output_path}")
                return output_path
            else:
                self.logger.error(f"AppleScript conversion failed: {result.stderr}")
                return None

        except Exception as e:
            self.logger.error(f"macOS conversion failed: {e}")
            return None

    def _convert_windows_com(self, ppt_path: Path, output_path: Path) -> Optional[Path]:
        """Convert using COM interface on Windows."""
        try:
            import win32com.client

            # Connect to PowerPoint
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = False

            # Open presentation
            presentation = ppt_app.Presentations.Open(str(ppt_path))

            # Save as PPTX (format 24 = ppSaveAsOpenXMLPresentation)
            presentation.SaveAs(str(output_path), 24)

            # Clean up
            presentation.Close()
            ppt_app.Quit()

            if output_path.exists():
                self.logger.info(f"Successfully converted to {output_path}")
                return output_path
            else:
                self.logger.error("Conversion completed but output file not found")
                return None

        except ImportError:
            self.logger.error("pywin32 not available for Windows COM conversion")
            return None
        except Exception as e:
            self.logger.error(f"Windows COM conversion failed: {e}")
            return None

    def _convert_libreoffice(self, ppt_path: Path, output_path: Path) -> Optional[Path]:
        """Convert using LibreOffice (cross-platform fallback)."""
        try:
            # Try different LibreOffice executable names
            libreoffice_commands = ['libreoffice', 'soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice']

            for cmd in libreoffice_commands:
                try:
                    result = subprocess.run([
                        cmd, '--headless', '--convert-to', 'pptx',
                        '--outdir', str(output_path.parent),
                        str(ppt_path)
                    ], capture_output=True, text=True, timeout=60)

                    if result.returncode == 0 and output_path.exists():
                        self.logger.info(f"Successfully converted using LibreOffice: {output_path}")
                        return output_path

                except FileNotFoundError:
                    continue

            self.logger.error("LibreOffice not found for conversion")
            return None

        except Exception as e:
            self.logger.error(f"LibreOffice conversion failed: {e}")
            return None

    def is_conversion_available(self) -> bool:
        """Check if conversion capabilities are available."""
        if self.system == "Darwin":
            # Check if PowerPoint is available on macOS
            try:
                result = subprocess.run(['osascript', '-e', 'tell application "Microsoft PowerPoint" to get version'],
                                      capture_output=True, text=True, timeout=5)
                return result.returncode == 0
            except:
                return False

        elif self.system == "Windows":
            # Check if PowerPoint COM is available
            try:
                import win32com.client
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                ppt_app.Quit()
                return True
            except:
                return False

        else:
            # Check if LibreOffice is available
            libreoffice_commands = ['libreoffice', 'soffice']
            for cmd in libreoffice_commands:
                try:
                    result = subprocess.run([cmd, '--version'], capture_output=True, timeout=5)
                    if result.returncode == 0:
                        return True
                except:
                    continue
            return False

    def cleanup_converted_file(self, file_path: Path) -> bool:
        """Clean up temporary converted file."""
        try:
            if file_path.exists():
                file_path.unlink()
                self.logger.info(f"Cleaned up temporary file: {file_path}")
                return True
            return False
        except Exception as e:
            self.logger.error(f"Failed to cleanup file {file_path}: {e}")
            return False

# Service instance
ppt_converter = PPTConverter()