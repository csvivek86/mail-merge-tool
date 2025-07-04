import os
import sys
import json
import logging
from dataclasses import dataclass, asdict
from typing import List, Optional
from pathlib import Path

# Create a module-level logger
logger = logging.getLogger(__name__)

def get_template_file_path() -> Path:
    """Get the path to the email template settings file"""
    try:
        # Use CONFIG_DIR from environment if available (set by AppSettings)
        if 'CONFIG_DIR' in os.environ:
            config_dir = Path(os.environ['CONFIG_DIR'])
        else:
            # Fallback to user's application support directory
            if sys.platform == 'darwin':
                config_dir = Path.home() / "Library" / "Application Support" / "NSNA Mail Merge" / "config"
            else:
                config_dir = Path.home() / ".nsna-mail-merge" / "config"
        
        # Ensure directory exists
        config_dir.mkdir(parents=True, exist_ok=True)
        template_path = config_dir / "email_template_settings.json"
        logger.debug(f"Email template path: {template_path}")
        return template_path
        
    except Exception as e:
        logger.error(f"Error determining template path: {e}")
        # Fallback to temporary directory
        fallback_dir = Path.home() / "NSNA_Mail_Merge_Data" / "config"
        fallback_dir.mkdir(parents=True, exist_ok=True)
        template_path = fallback_dir / "email_template_settings.json"
        logger.info(f"Using fallback template path: {template_path}")
        return template_path

@dataclass
class EmailTemplateSettings:
    """Settings for email template"""
    subject: str = "NSNA Donation Receipt"
    greeting: str = "Dear"
    body: Optional[List[str]] = None
    signature: Optional[List[str]] = None
    body_html: Optional[str] = None
    variables: Optional[List[str]] = None
    
    def __post_init__(self):
        if self.body is None:
            self.body = [
                "Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.",
                "",
                "Please note that you may be receiving separate tax acknowledgement letters if you have made multiple contributions to NSNA in 2025.",
                "",
                "Please do not hesitate to contact me if you have any questions on the attached document."
            ]
        if self.signature is None:
            self.signature = [
                "Regards,",
                "",
                "Nagarathar Sangam of North America Treasurer",
                "Email: nsnatreasurer@achi.org"
            ]
        if self.body_html is None:
            self.body_html = "<p>Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.</p><p>Please note that you may be receiving separate tax acknowledgement letters if you have made multiple contributions to NSNA in 2025.</p><p>Please do not hesitate to contact me if you have any questions on the attached document.</p>"
        if self.variables is None:
            self.variables = []
    
    def save(self) -> bool:
        """Save the email template settings to the user's configuration directory"""
        try:
            settings_file = get_template_file_path()
            # Ensure parent directory exists
            settings_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Convert to dict, excluding None values
            data = {k: v for k, v in asdict(self).items() if v is not None}
            
            # Write to a temporary file first
            temp_file = settings_file.with_suffix('.tmp')
            with open(temp_file, 'w') as f:
                json.dump(data, f, indent=2)
            
            # Rename temporary file to final name
            temp_file.replace(settings_file)
            
            logger.info(f"Email template settings saved to {settings_file}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to save email template settings: {e}")
            return False
    
    @classmethod
    def load(cls) -> 'EmailTemplateSettings':
        """Load email template settings from the user's configuration directory"""
        try:
            settings_file = get_template_file_path()
            if settings_file.exists():
                with open(settings_file, 'r') as f:
                    data = json.load(f)
                    logger.info(f"Loaded email template settings from {settings_file}")
                    return cls(**data)
            else:
                logger.info("No existing template found, using defaults")
                return cls()
                
        except Exception as e:
            logger.error(f"Failed to load email template settings: {e}")
            return cls()  # Return default settings