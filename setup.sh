#!/bin/bash

# setup.sh - Auto-setup script for new clones

echo "Setting up SPS Excel Pipeline..."

# Create necessary directories
echo "Creating required folders..."
mkdir -p data/raw data/processed reports

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo " Python 3 is not installed. Please install Python 3.8+"
    exit 1
fi

# Create virtual environment (optional)
if [ ! -d "venv" ]; then
    echo " Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install dependencies
echo " Installing dependencies..."
pip install -r requirements.txt

echo ""
echo " Setup complete!"
echo ""
echo "To run the application:"
echo "  1. Activate venv: source venv/bin/activate"
echo "  2. Run: streamlit run src/app.py"
echo ""
echo "Or with Docker:"
echo "  docker-compose up -d"