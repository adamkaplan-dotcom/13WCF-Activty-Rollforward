#!/bin/bash

# Navigate to the app directory
cd "$(dirname "$0")"

echo "=========================================="
echo "Balance Updater (Python) - Starting..."
echo "=========================================="
echo ""

# Check if Python3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed!"
    echo "Please install Python 3 from https://www.python.org/"
    read -p "Press Enter to exit..."
    exit 1
fi

# Check if virtual environment exists, create if not
if [ ! -d "python_venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv python_venv
fi

# Activate virtual environment
source python_venv/bin/activate

# Install/upgrade dependencies
echo "Installing dependencies..."
pip install -q --upgrade pip
pip install -q Flask==3.0.0 openpyxl==3.1.2 Werkzeug==3.0.1

# Create necessary directories
mkdir -p uploads outputs

# Start the Flask app
echo ""
echo "=========================================="
echo "✅ Server is starting..."
echo "=========================================="
echo ""
echo "🌐 Open your browser and go to:"
echo ""
echo "    http://localhost:5000"
echo ""
echo "=========================================="
echo "Press CTRL+C to stop the server"
echo "=========================================="
echo ""

python3 "13WCF-Activity Rollforward App v1.0.py"

# Deactivate when done
deactivate
