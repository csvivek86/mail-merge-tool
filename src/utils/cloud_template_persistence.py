"""
Cloud-compatible template persistence using session state and export/import
"""

import streamlit as st
import json
import base64
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, Any, Optional

class CloudTemplatePersistence:
    """
    Template persistence designed for Streamlit Community Cloud deployment.
    Uses session state for in-session storage and provides export/import functionality.
    """
    
    def __init__(self):
        self.init_session_state()
    
    def init_session_state(self):
        """Initialize session state for template storage"""
        if 'cloud_templates' not in st.session_state:
            st.session_state.cloud_templates = {
                'email_template': None,
                'pdf_template': None,
                'user_settings': None,
                'last_modified': {
                    'email_template': None,
                    'pdf_template': None,
                    'user_settings': None
                }
            }
    
    def save_email_template(self, template_data: Dict[str, Any]) -> bool:
        """Save email template to session state"""
        try:
            template_data['last_saved'] = datetime.now().isoformat()
            st.session_state.cloud_templates['email_template'] = template_data.copy()
            st.session_state.cloud_templates['last_modified']['email_template'] = datetime.now().isoformat()
            return True
        except Exception as e:
            logging.error(f"Failed to save email template: {e}")
            return False
    
    def save_pdf_template(self, template_data: Dict[str, Any]) -> bool:
        """Save PDF template to session state"""
        try:
            template_data['last_saved'] = datetime.now().isoformat()
            st.session_state.cloud_templates['pdf_template'] = template_data.copy()
            st.session_state.cloud_templates['last_modified']['pdf_template'] = datetime.now().isoformat()
            return True
        except Exception as e:
            logging.error(f"Failed to save PDF template: {e}")
            return False
    
    def save_user_settings(self, settings_data: Dict[str, Any]) -> bool:
        """Save user settings to session state"""
        try:
            settings_data['last_saved'] = datetime.now().isoformat()
            st.session_state.cloud_templates['user_settings'] = settings_data.copy()
            st.session_state.cloud_templates['last_modified']['user_settings'] = datetime.now().isoformat()
            return True
        except Exception as e:
            logging.error(f"Failed to save user settings: {e}")
            return False
    
    def load_email_template(self) -> Optional[Dict[str, Any]]:
        """Load email template from session state"""
        return st.session_state.cloud_templates.get('email_template')
    
    def load_pdf_template(self) -> Optional[Dict[str, Any]]:
        """Load PDF template from session state"""
        return st.session_state.cloud_templates.get('pdf_template')
    
    def load_user_settings(self) -> Optional[Dict[str, Any]]:
        """Load user settings from session state"""
        return st.session_state.cloud_templates.get('user_settings')
    
    def get_template_info(self) -> Dict[str, Any]:
        """Get template information"""
        templates = st.session_state.cloud_templates
        return {
            'email_template_exists': templates['email_template'] is not None,
            'pdf_template_exists': templates['pdf_template'] is not None,
            'user_settings_exists': templates['user_settings'] is not None,
            'email_template_modified': templates['last_modified']['email_template'],
            'pdf_template_modified': templates['last_modified']['pdf_template'],
            'user_settings_modified': templates['last_modified']['user_settings'],
            'templates_directory': 'Session State (Cloud Compatible)',
            'storage_type': 'Cloud Session State'
        }
    
    def delete_templates(self) -> bool:
        """Delete all templates from session state"""
        try:
            st.session_state.cloud_templates = {
                'email_template': None,
                'pdf_template': None,
                'user_settings': None,
                'last_modified': {
                    'email_template': None,
                    'pdf_template': None,
                    'user_settings': None
                }
            }
            return True
        except Exception as e:
            logging.error(f"Failed to delete templates: {e}")
            return False
    
    def export_templates(self) -> str:
        """Export all templates as JSON string for download"""
        try:
            export_data = {
                'export_version': '1.0',
                'export_date': datetime.now().isoformat(),
                'app_version': 'NSNA Mail Merge v2.0 (Streamlit Cloud)',
                'templates': st.session_state.cloud_templates
            }
            return json.dumps(export_data, indent=2)
        except Exception as e:
            logging.error(f"Failed to export templates: {e}")
            return None
    
    def import_templates(self, json_data: str) -> bool:
        """Import templates from JSON string"""
        try:
            import_data = json.loads(json_data)
            
            # Validate import data
            if 'templates' not in import_data:
                raise ValueError("Invalid template file format")
            
            # Import templates
            templates = import_data['templates']
            st.session_state.cloud_templates['email_template'] = templates.get('email_template')
            st.session_state.cloud_templates['pdf_template'] = templates.get('pdf_template')
            st.session_state.cloud_templates['user_settings'] = templates.get('user_settings')
            
            # Update modification times
            current_time = datetime.now().isoformat()
            if templates.get('email_template'):
                st.session_state.cloud_templates['last_modified']['email_template'] = current_time
            if templates.get('pdf_template'):
                st.session_state.cloud_templates['last_modified']['pdf_template'] = current_time
            if templates.get('user_settings'):
                st.session_state.cloud_templates['last_modified']['user_settings'] = current_time
            
            return True
            
        except Exception as e:
            logging.error(f"Failed to import templates: {e}")
            return False

# Cloud-compatible letterhead path resolver
def get_letterhead_path():
    """Get the path to the letterhead PDF, cloud-compatible"""
    # Priority order for Streamlit Cloud deployment
    letterhead_paths = [
        Path('NSNA Atlanta Letterhead Updated.pdf'),  # Root directory (in repository)
        Path('src/templates/letterhead_template.pdf')  # Templates directory (in repository)
    ]
    
    for path in letterhead_paths:
        if path.exists():
            return path
    
    return None

# Cloud-compatible settings management
def get_cloud_email_settings():
    """Get email settings from Streamlit secrets or fallback to defaults"""
    try:
        # Try to get from Streamlit secrets first
        if hasattr(st, 'secrets') and 'email' in st.secrets:
            return {
                'smtp_server': st.secrets.email.get('smtp_server', 'smtp.gmail.com'),
                'smtp_port': st.secrets.email.get('smtp_port', 587),
                'use_tls': st.secrets.email.get('use_tls', True),
                'default_from': st.secrets.email.get('default_from', ''),
            }
    except:
        pass
    
    # Fallback to default settings
    return {
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'use_tls': True,
        'default_from': '',
    }

def get_oauth_credentials():
    """Get OAuth credentials from Streamlit secrets"""
    try:
        if hasattr(st, 'secrets') and 'oauth' in st.secrets:
            return {
                'client_id': st.secrets.oauth.get('client_id', ''),
                'client_secret': st.secrets.oauth.get('client_secret', ''),
                'redirect_uri': st.secrets.oauth.get('redirect_uri', ''),
            }
    except:
        pass
    
    return None
