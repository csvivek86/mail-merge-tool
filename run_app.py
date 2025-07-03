#!/usr/bin/env python3
"""
NSNA Mail Merge Tool Launcher
Run this script to start the application
"""

import sys
import os
import logging
from pathlib import Path

def initialize_logging():
    """Initialize basic logging configuration before importing app"""
    try:
        # Determine the logs directory
        if sys.platform == 'darwin':
            base_dir = Path.home() / "Library" / "Application Support" / "NSNA Mail Merge"
        else:
            base_dir = Path.home() / "NSNA_Mail_Merge_Data"
        
        log_dir = base_dir / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        
        # Set the environment variable for other parts of the app
        os.environ['LOGS_DIR'] = str(log_dir)
        
        # Configure basic logging to file and console
        log_file = log_dir / 'app.log'
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
            handlers=[
                logging.FileHandler(str(log_file)),
                logging.StreamHandler(sys.stdout)
            ]
        )
        logging.info("Initial logging system configured")
        return True
    except Exception as e:
        print(f"Error setting up logging: {e}", file=sys.stderr)
        # Fallback to basic console logging
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
        logging.error("Failed to set up file logging, using console only", exc_info=True)
        return False

# Add src directory to Python path
project_root = Path(__file__).parent
src_path = project_root / 'src'
sys.path.insert(0, str(src_path))

# Change to project directory
os.chdir(project_root)

# Initialize logging before importing the app
initialize_logging()

# Import and run the application
try:
    from src.app import MainWindow
    from PyQt6.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
except ImportError as e:
    logging.error(f"Import error: {e}")
    print("Make sure all dependencies are installed:")
    print("pip install -r requirements.txt")
    sys.exit(1)
except Exception as e:
    logging.error("Error starting application", exc_info=True)
    print(f"Error starting application: {e}")
    sys.exit(1)
