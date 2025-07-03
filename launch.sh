#!/bin/bash

# NSNA Mail Merge Tool Launcher Script
# This script sets up the environment and launches the application

echo "🚀 Starting NSNA Mail Merge Tool..."

# Check if we're in the correct directory
if [ ! -f "run_app.py" ]; then
    echo "❌ Error: Please run this script from the mail-merge-tool directory"
    exit 1
fi

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed or not in PATH"
    exit 1
fi

# Check if required dependencies are installed
echo "🔍 Checking dependencies..."
python3 -c "
import sys
required_modules = ['PyQt6', 'pandas', 'reportlab', 'openpyxl']
missing = []
for module in required_modules:
    try:
        __import__(module)
    except ImportError:
        missing.append(module)

if missing:
    print(f'❌ Missing dependencies: {missing}')
    print('📦 Install with: pip install -r requirements.txt')
    sys.exit(1)
else:
    print('✅ All dependencies found')
"

if [ $? -ne 0 ]; then
    echo "💡 Run: pip install -r requirements.txt"
    exit 1
fi

# Launch the application
echo "🎯 Launching application..."
python3 run_app.py

echo "👋 Application closed"
