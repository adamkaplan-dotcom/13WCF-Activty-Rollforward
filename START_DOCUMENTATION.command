#!/bin/bash

# Get the directory where this script is located
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Change to that directory
cd "$DIR"

# Run the documentation server
echo "Starting 13WCF Activity Rollforward Documentation Dashboard..."
python3 documentation_server.py
