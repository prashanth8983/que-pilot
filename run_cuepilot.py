#!/usr/bin/env python3
"""
CuePilot Launcher

Easy launcher for the integrated CuePilot presentation assistant.
"""

import sys
import os
import subprocess
import logging
from pathlib import Path

def setup_logging():
    """Setup logging for the launcher."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

def check_python_version():
    """Check if Python version is compatible."""
    if sys.version_info < (3, 8):
        print(f"❌ Python {sys.version_info.major}.{sys.version_info.minor} is not supported.")
        print("Please use Python 3.8 or higher.")
        return False

    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    return True

def check_virtual_environment():
    """Check if virtual environment is activated."""
    venv_path = Path("venv")

    if venv_path.exists():
        # Check if we're in the virtual environment
        if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
            print("✅ Virtual environment is activated")
            return True
        else:
            print("⚠️  Virtual environment exists but not activated")
            print("Please activate it with:")
            if os.name == 'nt':  # Windows
                print("   venv\\Scripts\\activate")
            else:  # Unix/macOS
                print("   source venv/bin/activate")
            return False
    else:
        print("ℹ️  No virtual environment found")
        return True

def check_dependencies():
    """Check if required dependencies are installed."""
    logger = logging.getLogger(__name__)
    missing_deps = []

    # Core dependencies
    try:
        import PySide6
        print("✅ PySide6 (GUI framework)")
    except ImportError:
        missing_deps.append("PySide6>=6.5.0")
        print("❌ PySide6 (GUI framework)")

    try:
        import pptx
        print("✅ python-pptx (PowerPoint processing)")
    except ImportError:
        missing_deps.append("python-pptx>=0.6.21")
        print("❌ python-pptx (PowerPoint processing)")

    try:
        import pyaudio
        print("✅ PyAudio (audio capture)")
    except ImportError:
        missing_deps.append("pyaudio>=0.2.11")
        print("❌ PyAudio (audio capture)")

    try:
        import whisper
        print("✅ OpenAI Whisper (transcription)")
    except ImportError:
        missing_deps.append("openai-whisper")
        print("❌ OpenAI Whisper (transcription)")

    try:
        import yaml
        print("✅ PyYAML (configuration)")
    except ImportError:
        missing_deps.append("PyYAML")
        print("❌ PyYAML (configuration)")

    try:
        import openai
        print("✅ OpenAI (LLM interface)")
    except ImportError:
        missing_deps.append("openai")
        print("❌ OpenAI (LLM interface)")

    # Optional dependencies
    try:
        import cv2
        print("✅ OpenCV (computer vision)")
    except ImportError:
        print("⚠️  OpenCV (optional)")

    try:
        import psutil
        print("✅ psutil (system monitoring)")
    except ImportError:
        print("⚠️  psutil (optional)")

    # macOS specific
    if sys.platform == "darwin":
        try:
            import Cocoa
            import Quartz
            print("✅ PyObjC (macOS integration)")
        except ImportError:
            missing_deps.append("pyobjc-framework-Cocoa")
            missing_deps.append("pyobjc-framework-Quartz")
            print("❌ PyObjC (macOS integration)")

    if missing_deps:
        print(f"\n❌ Missing {len(missing_deps)} dependencies:")
        for dep in missing_deps:
            print(f"   • {dep}")
        print("\nInstall missing dependencies with:")
        print(f"   pip install {' '.join(missing_deps)}")
        return False

    print("\n✅ All dependencies are installed!")
    return True

def check_lm_studio():
    """Check if LM Studio is running."""
    try:
        import requests
        response = requests.get("http://localhost:1234/v1/models", timeout=2)
        if response.status_code == 200:
            models = response.json()
            if models.get('data'):
                print("✅ LM Studio is running with models loaded")
                return True
            else:
                print("⚠️  LM Studio is running but no models loaded")
                return False
        else:
            print("❌ LM Studio is not responding correctly")
            return False
    except Exception:
        print("❌ LM Studio is not running (http://localhost:1234)")
        print("Please start LM Studio and load a model for full functionality")
        return False

def check_config_file():
    """Check if config file exists."""
    config_path = Path("config.yaml")
    if config_path.exists():
        print("✅ Configuration file found")
        return True
    else:
        print("⚠️  Configuration file not found")
        print("Creating default config.yaml...")

        # Create default config
        default_config = '''# CuePilot Configuration File
LM_STUDIO_URL: "http://localhost:1234/v1"
LM_STUDIO_API_KEY: "lm-studio"
MODEL_GEMMA34BE: "hugging-quants/llama-3.2-3b-instruct"
AUDIO_SAMPLE_RATE: 16000
CUEPILOT_TEMPERATURE: 0.7
'''
        try:
            config_path.write_text(default_config)
            print("✅ Default configuration created")
            return True
        except Exception as e:
            print(f"❌ Failed to create config file: {e}")
            return False

def run_tests():
    """Run integration tests."""
    test_script = Path("test_cuepilot_integration.py")
    if test_script.exists():
        print("\n🧪 Running integration tests...")
        try:
            result = subprocess.run([sys.executable, str(test_script)], capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ All integration tests passed!")
                return True
            else:
                print("❌ Some integration tests failed:")
                print(result.stdout)
                return False
        except Exception as e:
            print(f"❌ Failed to run tests: {e}")
            return False
    else:
        print("⚠️  Test script not found, skipping tests")
        return True

def launch_application():
    """Launch the CuePilot application."""
    print("\n🚀 Launching CuePilot...")

    try:
        # Try to launch the integrated version first
        from src.app.cuepilot_main_window import main
        print("Starting CuePilot integrated application...")
        return main()
    except ImportError:
        print("Integrated version not available, trying fallback...")
        try:
            # Fallback to original main
            from main import main
            print("Starting fallback application...")
            return main()
        except ImportError as e:
            print(f"❌ Failed to launch application: {e}")
            return 1

def main():
    """Main launcher function."""
    logger = setup_logging()

    print("🎯 CuePilot Presentation Assistant")
    print("=" * 50)

    # Check system requirements
    checks = [
        ("Python Version", check_python_version),
        ("Virtual Environment", check_virtual_environment),
        ("Dependencies", check_dependencies),
        ("Configuration", check_config_file),
        ("LM Studio", check_lm_studio),
    ]

    all_good = True
    for check_name, check_func in checks:
        print(f"\n🔍 {check_name}:")
        try:
            result = check_func()
            if not result and check_name in ["Python Version", "Dependencies"]:
                all_good = False
        except Exception as e:
            print(f"❌ {check_name} check failed: {e}")
            if check_name in ["Python Version", "Dependencies"]:
                all_good = False

    if not all_good:
        print("\n❌ Critical checks failed. Please fix the issues above before continuing.")
        return 1

    # Run tests
    if not run_tests():
        print("\n⚠️  Tests failed, but continuing anyway...")

    # Launch application
    try:
        return launch_application()
    except KeyboardInterrupt:
        print("\n👋 Goodbye!")
        return 0
    except Exception as e:
        print(f"\n❌ Application crashed: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())