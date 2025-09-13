#!/bin/bash

# Activate virtual environment
source venv/bin/activate

# Check if an argument was provided
if [ $# -eq 0 ]; then
    echo "PowerPoint Tracking Application - Virtual Environment Activated"
    echo "Available commands:"
    echo "  python main_app.py    - Run the GUI application"
    echo "  python test_app.py    - Run the test suite"
    echo "  python ppt_tracker.py - Use the PowerPoint tracker directly"
    echo ""
    echo "Virtual environment is now active. You can run Python commands."
    exec bash
else
    # Run the provided command
    python "$@"
fi
