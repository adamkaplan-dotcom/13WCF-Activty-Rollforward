#!/bin/bash

# Navigate to the app directory
cd "$(dirname "$0")"

echo "=========================================="
echo "Balance Updater React App - Starting..."
echo "=========================================="
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed!"
    echo "Please install Node.js from https://nodejs.org/"
    read -p "Press Enter to exit..."
    exit 1
fi

# Check if node_modules exists, install if not
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
    echo ""
fi

# Start the development server
echo "=========================================="
echo "✅ Server is starting..."
echo "=========================================="
echo ""
echo "🌐 Your browser will open automatically to:"
echo ""
echo "    http://localhost:5173"
echo ""
echo "=========================================="
echo "Press CTRL+C to stop the server"
echo "=========================================="
echo ""

npm run dev
