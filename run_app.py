#!/usr/bin/env python3
"""
NSNA Mail Merge Tool Launcher
Run this script to start the application
"""

import sys
import os
from pathlib import Path

# Add src directory to Python path
project_root = Path(__file__).parent
src_path = project_root / 'src'
sys.path.insert(0, str(src_path))

# Change to project directory
os.chdir(project_root)

# Import and run the application
try:
    from app import main
    main()
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure all dependencies are installed:")
    print("pip install -r requirements.txt")
    sys.exit(1)
except Exception as e:
    print(f"Error starting application: {e}")
    sys.exit(1)
