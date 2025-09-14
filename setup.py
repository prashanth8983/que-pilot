"""
Setup script for AI Presentation Copilot.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
readme_path = Path(__file__).parent / "README.md"
long_description = readme_path.read_text(encoding="utf-8") if readme_path.exists() else ""

# Read requirements
requirements_path = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_path.exists():
    with open(requirements_path, 'r') as f:
        requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

setup(
    name="ai-presentation-copilot",
    version="1.0.0",
    author="Que Pilot Team",
    author_email="team@quepilot.com",
    description="A real-time, offline-first AI assistant for flawless presentations",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/quepilot/ai-presentation-copilot",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Multimedia :: Presentation",
        "Topic :: Scientific/Engineering :: Artificial Intelligence",
    ],
    python_requires=">=3.9",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.4.0",
            "pytest-qt>=4.2.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
        ],
        "ai": [
            "whisper>=1.1.10",
            "sentence-transformers>=2.2.2",
            "faiss-cpu>=1.7.4",
            "chromadb>=0.4.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "ai-presentation-copilot=main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["assets/icons/*.svg", "assets/styles/*.qss", "assets/models/*"],
    },
    zip_safe=False,
)
