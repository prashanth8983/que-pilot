#!/usr/bin/env python3
"""
Environment setup script for AI Presentation Copilot.
"""

import os
import sys
import subprocess
import platform
from pathlib import Path


def run_command(command, description):
    """Run a command and handle errors."""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed: {e.stderr}")
        return False


def check_python_version():
    """Check if Python version is compatible."""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 9):
        print(f"❌ Python 3.9+ required, found {version.major}.{version.minor}")
        return False
    print(f"✅ Python {version.major}.{version.minor}.{version.micro} is compatible")
    return True


def setup_virtual_environment():
    """Create and setup virtual environment."""
    venv_path = Path("venv")
    
    if venv_path.exists():
        print("📁 Virtual environment already exists")
        return True
    
    if not run_command("python -m venv venv", "Creating virtual environment"):
        return False
    
    # Activate and install pip
    if platform.system() == "Windows":
        activate_cmd = "venv\\Scripts\\activate"
        pip_cmd = "venv\\Scripts\\pip"
    else:
        activate_cmd = "source venv/bin/activate"
        pip_cmd = "venv/bin/pip"
    
    # Upgrade pip
    if not run_command(f"{pip_cmd} install --upgrade pip", "Upgrading pip"):
        return False
    
    return True


def install_dependencies():
    """Install project dependencies."""
    if platform.system() == "Windows":
        pip_cmd = "venv\\Scripts\\pip"
    else:
        pip_cmd = "venv/bin/pip"
    
    # Install basic requirements
    if not run_command(f"{pip_cmd} install -r requirements.txt", "Installing basic dependencies"):
        return False
    
    # Install development dependencies if requested
    if "--dev" in sys.argv:
        if not run_command(f"{pip_cmd} install -e .[dev]", "Installing development dependencies"):
            return False
    
    # Install AI dependencies if requested
    if "--ai" in sys.argv:
        if not run_command(f"{pip_cmd} install -e .[ai]", "Installing AI dependencies"):
            return False
    
    return True


def check_system_dependencies():
    """Check for system-level dependencies."""
    system = platform.system()
    
    print(f"🖥️ Detected system: {system}")
    
    # Check for Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        print("✅ Tesseract OCR is available")
    except Exception:
        print("⚠️ Tesseract OCR not found")
        if system == "Darwin":  # macOS
            print("📋 Install with: brew install tesseract")
        elif system == "Linux":
            print("📋 Install with: sudo apt install tesseract-ocr")
        elif system == "Windows":
            print("📋 Download from: https://github.com/UB-Mannheim/tesseract/wiki")
    
    # Check for ffmpeg
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
        print("✅ FFmpeg is available")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("⚠️ FFmpeg not found")
        if system == "Darwin":
            print("📋 Install with: brew install ffmpeg")
        elif system == "Linux":
            print("📋 Install with: sudo apt install ffmpeg")
        elif system == "Windows":
            print("📋 Download from: https://ffmpeg.org/download.html")


def create_directories():
    """Create necessary directories."""
    directories = [
        "assets/models",
        "assets/icons", 
        "assets/styles",
        "tests/test_core",
        "tests/test_app",
        "tests/test_services"
    ]
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
    
    print("📁 Created project directories")


def main():
    """Main setup function."""
    print("🚀 Setting up AI Presentation Copilot environment...")
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Create directories
    create_directories()
    
    # Setup virtual environment
    if not setup_virtual_environment():
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        sys.exit(1)
    
    # Check system dependencies
    check_system_dependencies()
    
    print("\n🎉 Environment setup complete!")
    print("\n📋 Next steps:")
    print("1. Activate virtual environment:")
    if platform.system() == "Windows":
        print("   venv\\Scripts\\activate")
    else:
        print("   source venv/bin/activate")
    print("2. Run the application:")
    print("   python main.py")
    print("3. (Optional) Download AI models:")
    print("   python scripts/download_models.py")


if __name__ == "__main__":
    main()
