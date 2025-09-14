#!/usr/bin/env python3
"""
Download and setup AI models for offline functionality.
"""

import os
import sys
from pathlib import Path
import subprocess
import requests
from tqdm import tqdm


def download_file(url: str, destination: Path, description: str = ""):
    """Download a file with progress bar."""
    print(f"📥 Downloading {description}...")
    
    response = requests.get(url, stream=True)
    response.raise_for_status()
    
    total_size = int(response.headers.get('content-length', 0))
    
    with open(destination, 'wb') as f:
        with tqdm(total=total_size, unit='B', unit_scale=True, desc=description) as pbar:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                pbar.update(len(chunk))
    
    print(f"✅ Downloaded {description}")


def setup_whisper_models():
    """Download Whisper models."""
    try:
        import whisper
        
        models_dir = Path("assets/models/whisper")
        models_dir.mkdir(parents=True, exist_ok=True)
        
        # Download medium model (good balance of speed/accuracy)
        print("🎤 Setting up Whisper models...")
        model = whisper.load_model("medium", download_root=str(models_dir))
        print("✅ Whisper models ready")
        
    except ImportError:
        print("⚠️ Whisper not installed. Install with: pip install whisper")
    except Exception as e:
        print(f"❌ Failed to setup Whisper: {e}")


def setup_vector_models():
    """Download sentence transformer models."""
    try:
        from sentence_transformers import SentenceTransformer
        
        models_dir = Path("assets/models/sentence_transformers")
        models_dir.mkdir(parents=True, exist_ok=True)
        
        print("🔤 Setting up sentence transformer models...")
        model = SentenceTransformer('all-MiniLM-L6-v2', cache_folder=str(models_dir))
        print("✅ Sentence transformer models ready")
        
    except ImportError:
        print("⚠️ Sentence transformers not installed. Install with: pip install sentence-transformers")
    except Exception as e:
        print(f"❌ Failed to setup sentence transformers: {e}")


def setup_tesseract():
    """Check Tesseract installation."""
    try:
        import pytesseract
        
        # Test Tesseract
        pytesseract.get_tesseract_version()
        print("✅ Tesseract OCR is available")
        
    except Exception as e:
        print(f"❌ Tesseract OCR not found: {e}")
        print("📋 Please install Tesseract OCR:")
        print("  - macOS: brew install tesseract")
        print("  - Ubuntu: sudo apt install tesseract-ocr")
        print("  - Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")


def main():
    """Main setup function."""
    print("🚀 Setting up AI Presentation Copilot models...")
    
    # Create models directory
    models_dir = Path("assets/models")
    models_dir.mkdir(parents=True, exist_ok=True)
    
    # Setup components
    setup_tesseract()
    setup_whisper_models()
    setup_vector_models()
    
    print("\n🎉 Model setup complete!")
    print("📝 Note: Some models may require additional setup based on your platform.")
    print("🔗 See README.md for detailed installation instructions.")


if __name__ == "__main__":
    main()
