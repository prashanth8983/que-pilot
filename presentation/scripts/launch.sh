#!/bin/bash

# Quick launch script for PowerPoint Tracker
echo "🚀 Launching PowerPoint Tracker..."

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found!"
    echo "Please run: python3 -m venv venv"
    echo "Then: source venv/bin/activate"
    echo "Then: pip install -r requirements.txt"
    exit 1
fi

# Activate virtual environment
source venv/bin/activate

# Check if main app exists
if [ ! -f "main_app.py" ]; then
    echo "❌ main_app.py not found!"
    echo "Please make sure you're in the correct directory."
    exit 1
fi

# Launch the application
echo "✅ Starting PowerPoint Tracker GUI..."
python main_app.py
