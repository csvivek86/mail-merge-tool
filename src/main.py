import sys
import logging
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from app import MainWindow
from config.settings import EMAIL_SETTINGS
from utils.excel_reader import read_excel
from utils.mail_sender import send_email
from utils.pdf_generator import PDFGenerator

def get_project_root():
    """Get the project root directory regardless of OS"""
    return Path(__file__).parent.parent

def initialize_logging():
    """Set up logging to file and console"""
    try:
        # Create logs directory in the application directory
        log_dir = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path.cwd()
        log_dir = log_dir / 'logs'
        log_dir.mkdir(exist_ok=True)
        
        # Setup logging configuration
        log_file = log_dir / 'app.log'
        logging.basicConfig(
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ],
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    except Exception as e:
        # If logging setup fails, write to a basic file
        with open('error.log', 'w') as f:
            f.write(f"Failed to initialize logging: {str(e)}")

def main():
    """Main application entry point with error handling"""
    try:
        initialize_logging()
        app = QApplication(sys.argv)
        
        # Log startup information
        logging.info("Application starting...")
        logging.info(f"Python version: {sys.version}")
        logging.info(f"Qt version: {app.applicationVersion()}")
        
        # Create and show main window
        window = MainWindow()
        window.show()
        
        # Start event loop
        sys.exit(app.exec())
        
    except Exception as e:
        logging.critical(f"Fatal error: {str(e)}", exc_info=True)
        with open('crash.log', 'w') as f:
            f.write(f"Application crashed: {str(e)}")
        if getattr(sys, 'gettrace', None) is not None:
            raise
        sys.exit(1)

if __name__ == "__main__":
    main()
    print('Emails and receipts sent successfully!')