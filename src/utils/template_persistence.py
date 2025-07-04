"""
Template persistence manager for NSNA Mail Merge Tool
Handles saving and loading of email and PDF templates
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional

class TemplatePersistence:
    """Manages saving and loading of email and PDF templates"""
    
    def __init__(self):
        # Create templates directory
        self.templates_dir = Path.home() / "NSNA_Mail_Merge_Data" / "saved_templates"
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        
        # Template file paths
        self.email_template_file = self.templates_dir / "email_template.json"
        self.pdf_template_file = self.templates_dir / "pdf_template.json"
        self.user_settings_file = self.templates_dir / "user_settings.json"
    
    def save_email_template(self, template_data: Dict[str, Any]) -> bool:
        """Save email template to disk"""
        try:
            template_data['last_saved'] = datetime.now().isoformat()
            template_data['version'] = '1.0'
            
            with open(self.email_template_file, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Error saving email template: {e}")
            return False
    
    def load_email_template(self) -> Optional[Dict[str, Any]]:
        """Load email template from disk"""
        try:
            if self.email_template_file.exists():
                with open(self.email_template_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Error loading email template: {e}")
            return None
    
    def save_pdf_template(self, template_data: Dict[str, Any]) -> bool:
        """Save PDF template to disk"""
        try:
            template_data['last_saved'] = datetime.now().isoformat()
            template_data['version'] = '1.0'
            
            with open(self.pdf_template_file, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Error saving PDF template: {e}")
            return False
    
    def load_pdf_template(self) -> Optional[Dict[str, Any]]:
        """Load PDF template from disk"""
        try:
            if self.pdf_template_file.exists():
                with open(self.pdf_template_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Error loading PDF template: {e}")
            return None
    
    def save_user_settings(self, settings_data: Dict[str, Any]) -> bool:
        """Save user settings (like from_email, subject defaults)"""
        try:
            settings_data['last_saved'] = datetime.now().isoformat()
            
            with open(self.user_settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, indent=2, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Error saving user settings: {e}")
            return False
    
    def load_user_settings(self) -> Optional[Dict[str, Any]]:
        """Load user settings"""
        try:
            if self.user_settings_file.exists():
                with open(self.user_settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Error loading user settings: {e}")
            return None
    
    def get_template_info(self) -> Dict[str, Any]:
        """Get information about saved templates"""
        info = {
            'email_template_exists': self.email_template_file.exists(),
            'pdf_template_exists': self.pdf_template_file.exists(),
            'user_settings_exists': self.user_settings_file.exists(),
            'templates_directory': str(self.templates_dir)
        }
        
        # Add last modified dates
        if info['email_template_exists']:
            info['email_template_modified'] = datetime.fromtimestamp(
                self.email_template_file.stat().st_mtime
            ).isoformat()
        
        if info['pdf_template_exists']:
            info['pdf_template_modified'] = datetime.fromtimestamp(
                self.pdf_template_file.stat().st_mtime
            ).isoformat()
        
        return info
    
    def delete_templates(self) -> bool:
        """Delete all saved templates"""
        try:
            for file_path in [self.email_template_file, self.pdf_template_file, self.user_settings_file]:
                if file_path.exists():
                    file_path.unlink()
            return True
        except Exception as e:
            print(f"Error deleting templates: {e}")
            return False
