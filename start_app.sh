#!/bin/bash

echo "Starting PDF to Markdown Converter..."
echo ""
echo "Make sure you have installed all dependencies:"
echo "pip install -r requirements.txt"
echo "pip install -U \"mineru[core]\""
echo ""
echo "Starting the application..."

# Check if Python is available
if command -v python3 &> /dev/null; then
    python3 app.py
elif command -v python &> /dev/null; then
    python app.py
else
    echo "Error: Python not found. Please install Python 3.10 or higher."
    exit 1
fi
