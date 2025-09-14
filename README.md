# AI Presentation Copilot

*A real-time, offline-first AI assistant with a modern Dracula-themed UI to help you deliver flawless presentations. This desktop application, built with Python and PySide6, automatically detects and syncs with your live presentation, listens as you speak, and provides intelligent assistance—all without an internet connection.*

## 🚀 Quick Start

```bash
# Clone the repository
git clone https://github.com/quepilot/ai-presentation-copilot.git
cd ai-presentation-copilot

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

## Table of Contents

* [Introduction](#introduction)
* [Features](#features)
* [Offline First & Privacy](#offline-first--privacy)
* [UI Design and Layout](#ui-design-and-layout)
  * [Core Structure](#core-structure)
  * [Visual Style: Dracula Theme](#visual-style-dracula-theme)
* [How It Works](#how-it-works)
* [Tech Stack](#tech-stack)
* [Setup and Installation](#setup-and-installation)
* [Usage](#usage)
* [Contributing](#contributing)
* [License](#license)

## Introduction

Giving a presentation can be stressful. The **AI Presentation Copilot** is your personal, private presentation coach that runs entirely locally on your machine.

The application features a polished, modern user interface inspired by the popular Dracula theme. It can **automatically detect an open PowerPoint presentation** on your screen, load its content, and **track your slide changes in real-time**. As you present, it uses a local instance of **OpenAI's Whisper** to transcribe your speech. By understanding your context, it can automatically display helpful notes or generate answers to questions, ensuring you never lose your train of thought.

## Features

* **Modern Dracula UI**: A beautiful, custom-styled interface built with PySide6, featuring a custom title bar and a professional layout.
* **Automatic PowerPoint Detection**: Natively detects and connects to running PowerPoint presentations on both **Windows and macOS**.
* **Real-time Slide Synchronization**: Actively tracks the current slide and stays in sync, whether in slideshow or edit mode.
* **Advanced Content Extraction**: Uses **OCR (Optical Character Recognition)** to read text from images and diagrams on your slides.
* **Context-Aware Vector Store**: Creates embeddings from your slide content and stores them in a local vector database for fast, relevant retrieval.
* **Real-time Transcription**: Uses a high-accuracy local Whisper model to transcribe your voice with low latency.
* **Intelligent Assistance**: Detects pauses and provides relevant prompts using a local LLM.
* **Fully Offline & Private**: All processing happens on your local machine. No data ever leaves your computer.

## Offline First & Privacy

This application is built with privacy as a core principle. **It does not require an internet connection to function.**

* **No External API Calls**: All models (speech-to-text, embeddings, and language generation) run locally.
* **No Data Transmission**: Your presentation content and your voice are never sent to the cloud.
* **Complete Confidentiality**: Safe for confidential or proprietary presentations.

## UI Design and Layout

The user interface is designed from the ground up to be modern, intuitive, and visually consistent with the Dracula theme popular among developers.

### Core Structure

The application window is frameless and organized into a dashboard layout with distinct functional zones.
```
+----------------------------------------------------------------------------------------------------------+
| O  AI Presentation Copilot                                                          [ _ ] [ # ] [ X ] |
+----------------------------------------------------------------------------------------------------------+
|[I]|| [ ADVANCED CONTROLS ]        | [ DASHBOARD ]                                                        |
|---||------------------------------|----------------------------------------------------------------------|
|[S]||  Presentation                |  +-[ AI Assistance ]----------------------------------------------+  |
|   ||  [ Open File... ]            |  | > You mentioned "market trends". Would you like to elaborate   |  |
|---||                              |  |   on the Q3 forecast?                                          |  |
|[P]||  Transcription Model         |  +----------------------------------------------------------------+  |
|   ||  [ Whisper (Medium)  v ]     |                                                                      |
|---||                              |  +-[ Live Transcription Feed ]------------------------------------+  |
|[D]||  [x] Real-time Assistance    |  | "...and as we can see from the chart, the quarterly |  |
|   ||  [ ] Auto-answer Questions   |  | "revenue has shown significant growth..."           |  |
|---||                              |  +----------------------------------------------------------------+  |
|[C]||  Confidence Threshold        |                                                                      |
|   ||  <----[ 75% ]----o----->     |  +-[ Presentation Status ]----++-[ Performance Metrics ]-----------+  |
|   ||                              |  |                            ||                                |  |
|---||  LLM Temperature             |  | Slide: 5 / 20              || Speaking Pace: 145 WPM         |  |
|   ||  <----[ 0.7 ]--o------->     |  | Notes: Key takeaway is...  || Confidence: 96%              |  |
+---+------------------------------+  +----------------------------++--------------------------------+  |
| [ STATUS: Listening... | File: my_slides.pptx | v1.0.0 ]                                                 |
+----------------------------------------------------------------------------------------------------------+

```
1.  **Custom Title Bar**: A draggable header containing the application icon, title, and custom window controls (minimize, maximize, close).
2.  **Activity Bar**: A compact, icon-based navigation bar on the far left for switching between major app sections (e.g., Info, Search, Settings).
3.  **Side Bar (Advanced Controls)**: This panel houses all user-configurable settings. It includes file management, model selection dropdowns, feature toggles (checkboxes), and fine-tuning parameters like confidence thresholds and LLM temperature (sliders).
4.  **Main Dashboard**: The primary content area is divided into several distinct widgets:
    * **AI Assistance**: Displays real-time suggestions, answers, and contextual prompts generated by the local LLM.
    * **Live Transcription Feed**: Shows the timestamped output from the Whisper model as you speak.
    * **Presentation Status**: Provides at-a-glance information about the loaded presentation, including the current slide number and speaker notes.
    * **Performance Metrics**: Offers live feedback on presentation delivery, such as speaking pace (WPM) and transcription confidence.
5.  **Status Bar**: A persistent bar at the bottom displaying the current application status, the name of the active file, and the app version.

### Visual Style: Dracula Theme

A global QSS stylesheet ensures a consistent and high-fidelity appearance based on the precise Dracula color palette.

| Element                | Color Code | Description               |
| ---------------------- | ---------- | ------------------------- |
| **Main Background** | `#282a36`  | `bg_main`                 |
| **Secondary Background**| `#2122c`   | `bg_secondary` (Side Bar) |
| **Borders & Inputs** | `#44475a`  | `bg_input_border`         |
| **Primary Text** | `#f8f8f2`  | White                     |
| **Secondary Text** | `#6272a4`  | Muted / Grey              |
| **Primary Accent** | `#bd93f9`  | Purple                    |
| **Secondary Accent** | `#8be9fd`  | Cyan / Link               |
| **Success Accent** | `#50fa7b`  | Green                     |
| **Destructive Accent** | `#ff5555`  | Red                       |

Key styling features include soft, rounded corners on all elements, custom-styled widgets (buttons, sliders, tables), and crisp iconography provided by the `qtawesome` library.

## How It Works

The application operates in two main phases: Ingestion/Sync and Real-time Assistance, all performed locally.

1.  **Ingestion & Sync**: Automatically detects a running presentation or allows manual upload. It extracts all content (including OCR from images) and creates a searchable vector database.
2.  **Real-time Assistance**: Tracks the active slide while transcribing your voice with a local Whisper model. When a pause or question is detected, it uses the vector database and a local LLM to generate and display helpful text.

## Tech Stack

* **Frontend (GUI)**: [PySide6](https://www.qt.io/qt-for-python)
* **Iconography**: [QtAwesome](https://github.com/spyder-ide/qtawesome)
* **Speech-to-Text**: [OpenAI Whisper (local model)](https://github.com/openai/whisper)
* **Presentation Parsing**: [python-pptx](https://python-pptx.readthedocs.io/)
* **Windowing & Automation**: [pywin32](https://pypi.org/project/pywin32/), [pyobjc](https://pyobjc.readthedocs.io/en/latest/)
* **Image Processing & OCR**: [OpenCV](https://opencv.org/), [Pillow](https://python-pillow.org/), [Tesseract (pytesseract)](https://pypi.org/project/pytesseract/)
* **Vector Database**: [FAISS](https://faiss.ai/) / [ChromaDB](https://www.trychroma.com/)
* **Core Language**: [Python 3.9+](https://www.python.org/)

## 📁 Project Structure

```
que-pilot/
├── main.py                          # Application entry point
├── setup.py                         # Package setup
├── requirements.txt                 # Dependencies
├── config/                          # Configuration
│   ├── settings.py                 # App settings & themes
├── src/                            # Main source code
│   ├── app/                        # Application layer
│   │   ├── main_window.py         # Main application window
│   │   ├── views/                  # UI Views
│   │   └── widgets/               # Custom UI widgets
│   ├── core/                       # Core business logic
│   │   ├── presentation/           # Presentation handling
│   │   ├── ai/                     # AI/ML components
│   │   └── utils/                  # Utilities
│   └── services/                   # Service layer
│       ├── presentation_service.py # Presentation management
│       ├── ai_service.py          # AI services coordination
│       └── sync_service.py        # Real-time synchronization
├── assets/                         # Static assets
│   ├── icons/                      # SVG icons
│   ├── styles/                     # QSS stylesheets
│   └── models/                     # AI model files
└── tests/                          # Test suite
```

## 🛠️ Setup and Installation

### Prerequisites

* Python 3.9+
* Git
* Tesseract-OCR Engine (for OCR functionality)
* ffmpeg (for audio processing)

### Installation Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/quepilot/ai-presentation-copilot.git
   cd ai-presentation-copilot
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install the required packages:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install additional AI dependencies (optional):**
   ```bash
   pip install -e ".[ai]"  # For Whisper, vector store, etc.
   ```

5. **Download Local Models (optional):**
   ```bash
   python scripts/download_models.py
   ```

6. **Run the application:**
   ```bash
   python main.py
   ```