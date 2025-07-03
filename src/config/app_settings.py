import json
import os
from pathlib import Path
import logging

# Create a module-level logger
logger = logging.getLogger(__name__)

class AppSettings:
    def __init__(self):
        # Use environment variables if available
        if 'CONFIG_DIR' in os.environ:
            self.config_dir = Path(os.environ['CONFIG_DIR'])
        else:
            self.config_dir = Path.home() / ".nsna-mail-merge"
            
        try:
            self.config_dir.mkdir(exist_ok=True)
        except PermissionError:
            logger.warning(f"Could not access config directory {self.config_dir}, using temporary directory")
            self.config_dir = Path(os.path.join(os.path.expanduser("~"), "NSNA_Mail_Merge_Data/config"))
            os.makedirs(self.config_dir, exist_ok=True)
            
        self.settings_file = self.config_dir / "settings.json"
        
        # Use environment variables for receipts if available
        if 'RECEIPTS_DIR' in os.environ:
            self.default_receipts_dir = Path(os.environ['RECEIPTS_DIR'])
        else:
            self.default_receipts_dir = Path.home() / "Documents" / "NSNA Receipts"
            
        # Look for template in multiple locations
        template_paths = []
        
        # First check if there's a template path in environment variable
        if 'PDF_TEMPLATE' in os.environ:
            template_paths.append(Path(os.environ['PDF_TEMPLATE']))
        
        # Add other possible template locations
        template_paths.extend([
            Path(__file__).parent.parent.parent / "NSNA Atlanta Letterhead Updated.pdf",
            Path(os.environ.get('DATA_DIR', '.')) / "NSNA Atlanta Letterhead Updated.pdf",
            Path.home() / "NSNA_Mail_Merge_Data" / "data" / "NSNA Atlanta Letterhead Updated.pdf",
            Path.home() / "NSNA_Mail_Merge_Data" / "NSNA Atlanta Letterhead Updated.pdf"
        ])
        
        # Try to find the first template that exists
        existing_template = next((p for p in template_paths if p.exists()), None)
        
        if existing_template:
            self.default_template = existing_template
            logger.info(f"Using PDF template: {self.default_template}")
        else:
            logger.warning("No PDF template found in any of the expected locations")
            self.default_template = None
        self._load_settings()

    def _load_settings(self):
        if self.settings_file.exists():
            with open(self.settings_file) as f:
                self.settings = json.load(f)
        else:
            self.settings = {
                "receipts_dir": str(self.default_receipts_dir),
                "from_email": ""
            }
            self._save_settings()

    def _save_settings(self):
        """Save settings to file"""
        with open(self.settings_file, 'w') as f:
            json.dump(self.settings, f, indent=2)

    def get_receipts_dir(self) -> Path:
        return Path(self.settings.get("receipts_dir", str(self.default_receipts_dir)))

    def set_receipts_dir(self, path: str):
        self.settings["receipts_dir"] = path
        self._save_settings()

    def get_from_email(self) -> str:
        return self.settings.get("from_email", "")

    def set_from_email(self, email: str):
        self.settings["from_email"] = email
        self._save_settings()

    def get_template_path(self) -> Path:
        """Get the PDF template path"""
        return Path(self.settings.get("template_path", str(self.default_template)))

    def set_template_path(self, path: str):
        """Save PDF template path"""
        self.settings["template_path"] = path
        self._save_settings()

class MainWindow:
    def __init__(self, app_settings: AppSettings):
        self.app_settings = app_settings
        self.template_path = self.app_settings.get_template_path()
        if not self.template_path.exists():
            logging.warning(f"Template not found at {self.template_path}")
            self.template_path = None
        else:
            logging.info(f"PDF template loaded from {self.template_path}")