#!/usr/bin/env python3
"""
Bootstrap script for PyInstaller packaged app
This script ensures that the application can find the right modules
when packaged with PyInstaller
"""

import os
import sys
import logging
from pathlib import Path
import mimetypes

def setup_logging():
    """Set up logging with both file and console handlers"""
    try:
        # Determine the logs directory
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # We are running in a PyInstaller bundle
            base_dir = Path.home() / "Library" / "Application Support" / "NSNA Mail Merge"
        else:
            # We are running in a normal Python environment
            base_dir = Path(__file__).parent.resolve()
        
        logs_dir = base_dir / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        log_file = logs_dir / "app.log"
        
        # Configure root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        
        # Remove any existing handlers
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        # Create formatters
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
        )
        
        # Create and configure file handler
        file_handler = logging.FileHandler(str(log_file))
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(file_handler)
        
        # Create and configure console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(console_handler)
        
        # Set environment variable for the logs directory
        os.environ['LOGS_DIR'] = str(logs_dir)
        
        logging.info(f"Logging system initialized. Log file: {log_file}")
        return True
        
    except Exception as e:
        # If we fail to set up logging, print to stderr and set up basic console logging
        print(f"Error setting up logging: {e}", file=sys.stderr)
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
        logging.error("Failed to set up logging with custom configuration", exc_info=True)
        return False

# Initialize logging before anything else
setup_logging()

# Patch mimetypes to avoid reading from system files
# This prevents "Operation not permitted" errors in sandboxed apps
def patch_mimetypes():
    """Fix mimetypes to work in sandboxed environments by completely recreating the types dictionary"""
    # Create a fresh mimetypes database without reading files
    mimetypes.knownfiles = []  # Prevent reading system files
    mimetypes.inited = True    # Prevent further initialization attempts
    
    # Create a new, empty types dictionary
    mimetypes.types_map = {
        '.xls': 'application/vnd.ms-excel',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.pdf': 'application/pdf',
        '.html': 'text/html',
        '.htm': 'text/html',
        '.css': 'text/css',
        '.js': 'application/javascript',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.txt': 'text/plain',
        '.csv': 'text/csv',
        '.xml': 'application/xml',
        '.zip': 'application/zip',
        '.json': 'application/json',
    }
    
    # Also set the other maps that mimetypes uses
    mimetypes.common_types = mimetypes.types_map.copy()
    mimetypes.encodings_map = {}

def copy_pdf_template(bundle_dir, data_dir):
    """Copy PDF template to user data directory if it doesn't exist"""
    pdf_name = "NSNA Atlanta Letterhead Updated.pdf"
    pdf_target = data_dir / pdf_name
    
    if not pdf_target.exists():
        import shutil
        # Try to find the template in various locations
        pdf_sources = [
            Path.cwd() / pdf_name,
            bundle_dir / pdf_name,
            bundle_dir / 'src' / pdf_name,
            bundle_dir / 'data' / pdf_name,
            Path(__file__).parent / pdf_name
        ]
        
        for pdf_source in pdf_sources:
            try:
                if pdf_source.exists():
                    shutil.copy(str(pdf_source), str(pdf_target))
                    logging.info(f"PDF template copied from {pdf_source} to {pdf_target}")
                    break
            except Exception as e:
                logging.warning(f"Error copying template from {pdf_source}: {e}")
        
        if not pdf_target.exists():
            logging.error("Failed to copy PDF template from any source location")
            return None
    else:
        logging.info(f"Using existing PDF template at {pdf_target}")
    
    return pdf_target


def setup_pyinstaller_environment():
    """Set up the environment for PyInstaller bundled application"""
    logging.info("Setting up PyInstaller environment")
    
    # Apply patches for sandboxed environment
    patch_mimetypes()
    logging.info("Applied mimetypes patch")
    
    # Create user data directories
    user_home = Path(os.path.expanduser("~"))
    app_data_dir = user_home / "NSNA_Mail_Merge_Data"
    logs_dir = app_data_dir / "logs"
    config_dir = app_data_dir / "config"
    receipts_dir = app_data_dir / "receipts"
    data_dir = app_data_dir / "data"
    
    # Create directories with proper permissions
    try:
        os.makedirs(logs_dir, exist_ok=True)
        os.makedirs(config_dir, exist_ok=True)
        os.makedirs(receipts_dir, exist_ok=True)
        os.makedirs(data_dir, exist_ok=True)
        logging.info("Created application directories")
    except Exception as e:
        logging.error(f"Failed to create application directories: {e}")
        raise
    
    # Set environment variables for our directories
    os.environ['LOGS_DIR'] = str(logs_dir)
    os.environ['RECEIPTS_DIR'] = str(receipts_dir)
    os.environ['DATA_DIR'] = str(data_dir)
    logging.info("Set environment variables")
    
    # Check if we're running in a PyInstaller bundle
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        logging.info("Running in PyInstaller bundle")
        bundle_dir = Path(sys._MEIPASS)
        
        # Make sure src directory is in the Python path
        src_dir = bundle_dir / 'src'
        if src_dir.exists():
            sys.path.insert(0, str(src_dir))
            logging.info(f"Added {src_dir} to Python path")
        else:
            logging.warning(f"Source directory not found at {src_dir}")
        
        # Set working directory to the user data directory
        os.chdir(app_data_dir)
        logging.info(f"Changed working directory to {app_data_dir}")
        
        # Make sure we can find our config files
        bundle_config_dir = bundle_dir / 'src' / 'config'
        if bundle_config_dir.exists():
            # Copy config files if not exist
            import shutil
            for config_file in bundle_config_dir.glob('*.json'):
                target_file = config_dir / config_file.name
                if not target_file.exists():
                    try:
                        shutil.copy(str(config_file), str(target_file))
                        logging.info(f"Copied config file: {config_file.name}")
                    except Exception as e:
                        logging.error(f"Failed to copy config file {config_file.name}: {e}")
            
            os.environ['CONFIG_DIR'] = str(config_dir)
            logging.info("Set CONFIG_DIR environment variable")
        else:
            logging.warning(f"Config directory not found at {bundle_config_dir}")
        
        # Copy PDF template if needed and set environment variable
        pdf_target = copy_pdf_template(bundle_dir, data_dir)
        if pdf_target:
            os.environ['PDF_TEMPLATE'] = str(pdf_target)
            logging.info(f"Set PDF_TEMPLATE to: {pdf_target}")
        else:
            logging.error("Could not locate PDF template")
    else:
        # Running in development mode
        project_root = Path(__file__).parent
        src_dir = project_root / 'src'
        sys.path.insert(0, str(src_dir))
        os.chdir(project_root)

if __name__ == "__main__":
    try:
        setup_pyinstaller_environment()
        
        # Log environment variables
        if getattr(sys, 'frozen', False):
            for env_var in ['LOGS_DIR', 'RECEIPTS_DIR', 'DATA_DIR', 'CONFIG_DIR', 'PDF_TEMPLATE']:
                if env_var in os.environ:
                    logging.info(f"{env_var}: {os.environ[env_var]}")
                else:
                    logging.warning(f"Environment variable {env_var} not set")
        
        # Add file handler now that logs directory exists
        if 'LOGS_DIR' in os.environ:
            try:
                log_file = Path(os.environ['LOGS_DIR']) / 'app.log'
                file_handler = logging.FileHandler(str(log_file))
                file_handler.setFormatter(
                    logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s')
                )
                root_logger = logging.getLogger()
                root_logger.addHandler(file_handler)
                logging.info(f"Added log file handler: {log_file}")
                
                # Test write to log file
                logging.info("Log file write test successful")
            except Exception as e:
                logging.error(f"Failed to set up log file handler: {e}", exc_info=True)
                print(f"Warning: Could not set up log file: {e}", file=sys.stderr)
        
        logging.info("Importing main application...")
        from src.app import main
        logging.info("Successfully imported src.app.main")
        main()
        
    except ImportError as e:
        error_msg = f"Import error: {e}\nMake sure all dependencies are installed (pip install -r requirements.txt)"
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)
        
    except Exception as e:
        error_msg = f"Error starting application: {e}"
        logging.error(error_msg, exc_info=True)
        print(error_msg)
        import traceback
        traceback.print_exc()
        sys.exit(1)
