import json
from pathlib import Path
import logging

class AppSettings:
    def __init__(self):
        self.config_dir = Path.home() / ".nsna-mail-merge"
        self.config_dir.mkdir(exist_ok=True)
        self.settings_file = self.config_dir / "settings.json"
        self.default_receipts_dir = Path.home() / "Documents" / "NSNA Receipts"
        self.default_template = Path(__file__).parent.parent.parent / "NSNA Atlanta Letterhead Updated.pdf"
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