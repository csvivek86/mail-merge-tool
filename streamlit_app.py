"""
NSNA Mail Merge Tool - Streamlit Web Application
Convert PyQt6 desktop app to a modern web interface
Docker-ready with desktop persistent storage
"""

import streamlit as st
import pandas as pd
import os
import sys
import logging
import json
import base64
import tempfile
import time
import platform
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import re
from html import unescape

# Additional imports for missing functionality
try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

# Rich text editor import
try:
    from streamlit_quill import st_quill
except ImportError:
    st_quill = None

# Add src directory to Python path
src_path = Path(__file__).parent / 'src'
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

# Enhanced editor no longer needed - using compact streamlit-quill editor

# Rich text editor helper function
def rich_text_editor(value="", placeholder="", key=None, height=300, overwrite=None, available_variables=None, editor_type="standard", **kwargs):
    """
    Rich text editor wrapper that uses streamlit_quill with embedded toolbar,
    falls back to st.text_area if unavailable
    """
    
    # Default variables if not provided
    if available_variables is None:
        available_variables = ["First Name", "Last Name", "Amount", "Date", "Current Year"]
    
    # Use streamlit_quill as primary editor with compact embedded toolbar
    if st_quill is not None:
        # Add custom CSS for ultra-compact, Word-like toolbar
        st.markdown("""
        <style>
        /* Make Quill toolbar more compact and Word-like */
        .ql-toolbar {
            border: 1px solid #ccc !important;
            border-bottom: none !important;
            padding: 4px 8px !important;
            min-height: 36px !important;
            background: #f9f9f9 !important;
        }
        
        .ql-toolbar .ql-formats {
            margin-right: 8px !important;
        }
        
        .ql-toolbar button {
            height: 24px !important;
            width: 24px !important;
            padding: 2px !important;
            margin: 1px !important;
            border-radius: 2px !important;
        }
        
        .ql-toolbar button:hover {
            background: #e6e6e6 !important;
        }
        
        .ql-toolbar button.ql-active {
            background: #d0d0d0 !important;
        }
        
        .ql-toolbar .ql-picker {
            font-size: 12px !important;
        }
        
        .ql-toolbar .ql-picker-label {
            padding: 2px 4px !important;
            height: 24px !important;
            line-height: 20px !important;
        }
        
        .ql-container {
            border: 1px solid #ccc !important;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif !important;
        }
        
        .ql-editor {
            padding: 12px !important;
            line-height: 1.4 !important;
            min-height: """ + str(height - 40) + """px !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Compact, Word-like toolbar configuration - single row with essential tools
        quill_content = st_quill(
            value=value,
            placeholder=placeholder,
            key=key,
            html=True,  # Return HTML content
            toolbar=[
                ['bold', 'italic', 'underline'],  # Basic formatting
                [{'list': 'ordered'}, {'list': 'bullet'}],  # Lists
                [{'header': [1, 2, 3, False]}],  # Headers
                [{'align': []}],  # Alignment
                ['link'],  # Links
                ['clean']  # Remove formatting
            ]
        )
        return quill_content
    else:
        # Fall back to st.text_area with helper text
        st.warning("⚠️ Rich text editor not available. Using plain text editor.")
        st.info("💡 Tip: Use markdown syntax - **bold**, *italic*, • bullets, 1. numbers")
        return st.text_area(
            label="Template Content",
            value=value,
            placeholder=placeholder + "\n\nMarkdown tips:\n**bold text**\n*italic text*\n• bullet point\n1. numbered item",
            key=key,
            height=height,
            help="Use markdown syntax for formatting: **bold**, *italic*, • bullets, 1. numbers"
        )

# Desktop Storage Configuration
def get_user_configured_path():
    """Get user-configured storage path from session state or saved settings"""
    # Check if user has configured a custom path in session state
    if 'custom_storage_path' in st.session_state and st.session_state.custom_storage_path:
        return Path(st.session_state.custom_storage_path)
    
    # Check for saved custom path from previous sessions
    try:
        settings_file = Path.home() / ".nsna_mail_merge_settings.json"
        if settings_file.exists():
            import json
            with open(settings_file, 'r') as f:
                settings = json.load(f)
                custom_path = settings.get('storage_path')
                if custom_path and Path(custom_path).exists():
                    return Path(custom_path)
    except:
        pass
    
    return None

def save_user_storage_path(path):
    """Save user's chosen storage path for future sessions"""
    try:
        import json
        settings_file = Path.home() / ".nsna_mail_merge_settings.json"
        settings = {}
        
        # Load existing settings if any
        if settings_file.exists():
            with open(settings_file, 'r') as f:
                settings = json.load(f)
        
        # Update storage path
        settings['storage_path'] = str(path)
        
        # Save settings
        with open(settings_file, 'w') as f:
            json.dump(settings, f, indent=2)
        
        # Update session state
        st.session_state.custom_storage_path = str(path)
        return True
    except Exception as e:
        st.error(f"Failed to save storage path: {e}")
        return False

def get_desktop_storage_path():
    """Get the desktop storage path for persistent data"""
    # First check if user has configured a custom path
    user_path = get_user_configured_path()
    if user_path:
        nsna_folder = user_path / "NSNA_Mail_Merge"
        nsna_folder.mkdir(exist_ok=True)
        
        # Create subfolders
        (nsna_folder / "templates").mkdir(exist_ok=True)
        (nsna_folder / "data").mkdir(exist_ok=True)
        (nsna_folder / "receipts").mkdir(exist_ok=True)
        (nsna_folder / "exports").mkdir(exist_ok=True)
        
        return nsna_folder
    
    # Check if running in Docker (environment variable set in docker-compose)
    if os.environ.get('DOCKER_DEPLOYMENT', 'false').lower() == 'true':
        # In Docker, use mounted volume that maps to user's desktop
        desktop_path = Path('/app/desktop_storage')
    else:
        # Check if we have a pre-configured path from the launcher
        if 'NSNA_DESKTOP_PATH' in os.environ:
            return Path(os.environ['NSNA_DESKTOP_PATH'])
        
        # Local development - find the correct desktop path
        import platform
        system = platform.system()
        
        if system == "Windows":
            # Try multiple desktop locations for OneDrive scenarios
            possible_desktops = [
                Path.home() / "Desktop",
                Path.home() / "OneDrive" / "Desktop", 
                Path.home() / "OneDrive - HD Supply, Inc" / "Desktop"
            ]
            
            desktop_path = None
            for path in possible_desktops:
                if path.exists():
                    try:
                        # Test write permission
                        test_folder = path / "test_nsna_write"
                        test_folder.mkdir(exist_ok=True)
                        test_folder.rmdir()
                        desktop_path = path
                        break
                    except:
                        continue
            
            # Fallback to Documents if no desktop is writable
            if desktop_path is None:
                desktop_path = Path.home() / "Documents"
                
        elif system == "Darwin":  # macOS
            desktop_path = Path.home() / "Desktop"
        else:  # Linux
            desktop_path = Path.home() / "Desktop"
            # Fallback if Desktop doesn't exist
            if not desktop_path.exists():
                desktop_path = Path.home() / "Documents"
    
    # Create NSNA Mail Merge folder on desktop
    nsna_folder = desktop_path / "NSNA_Mail_Merge"
    nsna_folder.mkdir(exist_ok=True)
    
    # Create subfolders
    (nsna_folder / "templates").mkdir(exist_ok=True)
    (nsna_folder / "data").mkdir(exist_ok=True)
    (nsna_folder / "receipts").mkdir(exist_ok=True)
    (nsna_folder / "exports").mkdir(exist_ok=True)
    
    return nsna_folder

# Global storage path will be set in main function
DESKTOP_STORAGE = None

# Import existing utility functions
try:
    from utils.excel_reader import read_excel, get_contacts_as_list
    from utils.mail_sender import send_email_with_diagnostics
    from utils.pdf_generator import PDFGenerator
    from utils.oauth_manager import get_user_credentials
    
    # Desktop-compatible persistence system
    try:
        from utils.cloud_template_persistence import (
            CloudTemplatePersistence, 
            get_letterhead_path, 
            get_cloud_email_settings,
            get_oauth_credentials
        )
        # Use cloud persistence but with desktop storage
        TemplatePersistence = CloudTemplatePersistence
        CLOUD_DEPLOYMENT = True
    except ImportError:
        # Fallback to local persistence for development
        from utils.template_persistence import TemplatePersistence
        CLOUD_DEPLOYMENT = False
    
    # Old rich text editor utilities no longer needed - using compact streamlit-quill
    from config.settings import EMAIL_SETTINGS, DEFAULT_EMAIL, USE_OAUTH, DYNAMIC_USER_OAUTH
    from config.email_template_settings import EmailTemplateSettings
    from config.template_settings import TemplateSettings
except ImportError as e:
    st.error(f"Import error: {e}")
    st.info("Please ensure all required modules are in the src/ directory")

# Configure Streamlit page
st.set_page_config(
    page_title="NSNA Mail Merge Tool",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="collapsed"
)



def create_desktop_persistence_manager():
    """Create a persistence manager that uses desktop storage"""
    # Get current desktop storage path
    storage_path = get_desktop_storage_path()
    
    class DesktopTemplatePersistence:
        def __init__(self, storage_path):
            self.storage_path = storage_path / "templates"
            self.storage_path.mkdir(exist_ok=True)
        
        def save_email_template(self, template):
            """Save email template to desktop"""
            try:
                template_file = self.storage_path / "email_template.json"
                with open(template_file, 'w', encoding='utf-8') as f:
                    json.dump(template, f, indent=2, ensure_ascii=False)
                return True
            except Exception as e:
                st.error(f"Failed to save email template: {e}")
                return False
        
        def load_email_template(self):
            """Load email template from desktop"""
            try:
                template_file = self.storage_path / "email_template.json"
                if template_file.exists():
                    with open(template_file, 'r', encoding='utf-8') as f:
                        return json.load(f)
                return None
            except Exception as e:
                st.error(f"Failed to load email template: {e}")
                return None
        
        def save_pdf_template(self, template):
            """Save PDF template to desktop"""
            try:
                template_file = self.storage_path / "pdf_template.json"
                with open(template_file, 'w', encoding='utf-8') as f:
                    json.dump(template, f, indent=2, ensure_ascii=False)
                return True
            except Exception as e:
                st.error(f"Failed to save PDF template: {e}")
                return False
        
        def load_pdf_template(self):
            """Load PDF template from desktop"""
            try:
                template_file = self.storage_path / "pdf_template.json"
                if template_file.exists():
                    with open(template_file, 'r', encoding='utf-8') as f:
                        return json.load(f)
                return None
            except Exception as e:
                st.error(f"Failed to load PDF template: {e}")
                return None
        
        def save_user_settings(self, settings):
            """Save user settings to desktop"""
            try:
                settings_file = self.storage_path / "user_settings.json"
                with open(settings_file, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, indent=2, ensure_ascii=False)
                return True
            except Exception as e:
                st.error(f"Failed to save user settings: {e}")
                return False
        
        def load_user_settings(self):
            """Load user settings from desktop"""
            try:
                settings_file = self.storage_path / "user_settings.json"
                if settings_file.exists():
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        return json.load(f)
                return None
            except Exception as e:
                st.error(f"Failed to load user settings: {e}")
                return None
    
    return DesktopTemplatePersistence(storage_path)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'email_settings' not in st.session_state:
    st.session_state.email_settings = None
if 'template_settings' not in st.session_state:
    st.session_state.template_settings = None
if 'sent_emails' not in st.session_state:
    st.session_state.sent_emails = []
if 'email_config' not in st.session_state:
    st.session_state.email_config = None
if 'pdf_template' not in st.session_state:
    st.session_state.pdf_template = None
if 'loaded_templates' not in st.session_state:
    st.session_state.loaded_templates = False

def load_settings():
    """Load application settings"""
    try:
        # Use settings from config module directly
        app_settings = {
            'email_settings': EMAIL_SETTINGS,
            'default_email': DEFAULT_EMAIL,
            'use_oauth': USE_OAUTH,
            'dynamic_user_oauth': DYNAMIC_USER_OAUTH
        }
        
        # Load email settings
        email_settings_path = Path('config/email_template_settings.json')
        if email_settings_path.exists():
            with open(email_settings_path) as f:
                email_data = json.load(f)
                email_settings = EmailTemplateSettings(**email_data)
        else:
            email_settings = EmailTemplateSettings()
        
        # Load template settings
        template_settings_path = Path('config/template_settings.json')
        if template_settings_path.exists():
            with open(template_settings_path) as f:
                template_data = json.load(f)
                template_settings = TemplateSettings(**template_data)
        else:
            template_settings = TemplateSettings()
            
        return app_settings, email_settings, template_settings
    except Exception as e:
        st.error(f"Error loading settings: {e}")
        return None, None, None

def load_saved_templates():
    """Load saved templates from disk"""
    if not st.session_state.loaded_templates:
        persistence = st.session_state.template_persistence
        
        # Load email template
        email_template = persistence.load_email_template()
        if email_template:
            st.session_state.saved_email_template = email_template
        
        # Load PDF template
        pdf_template = persistence.load_pdf_template()
        if pdf_template:
            st.session_state.saved_pdf_template = pdf_template
        
        # Load user settings
        user_settings = persistence.load_user_settings()
        if user_settings:
            st.session_state.saved_user_settings = user_settings
        
        st.session_state.loaded_templates = True

def main():
    """Main Streamlit application - Single page layout utilizing full width"""
    
    # Initialize desktop storage and template persistence
    global DESKTOP_STORAGE
    DESKTOP_STORAGE = get_desktop_storage_path()
    
    # Initialize template persistence manager if not already done
    if 'template_persistence' not in st.session_state:
        st.session_state.template_persistence = create_desktop_persistence_manager()
    
    # Custom CSS to remove padding and use full width
    st.markdown("""
    <style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: none;
    }
    
    .stApp > header {
        background-color: transparent;
    }
    
    .stApp {
        margin: 0;
        padding: 0;
    }
    
    /* Remove default margins */
    .element-container {
        margin: 0 !important;
    }
    
    /* Make containers use full width */
    .stColumn {
        padding: 0 0.5rem;
    }
    
    /* Remove extra spacing */
    .css-1d391kg {
        padding: 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.title("📧 NSNA Mail Merge Tool")
    
    # Load settings and templates
    app_settings, email_settings, template_settings = load_settings()
    if not app_settings:
        st.error("Failed to load application settings")
        return
    
    # Load saved templates on first run
    load_saved_templates()
    
    # Store in session state
    st.session_state.email_settings_config = email_settings
    st.session_state.template_settings = template_settings
    
    # Single page layout with columns
    single_page_layout()

def single_page_layout():
    """Single page layout utilizing full width - no tabs"""
    
    # Top row: Data upload and email settings
    st.header("📊 Setup")
    col1, col2, col3 = st.columns([2, 2, 2])
    
    with col1:
        st.subheader("Upload Data")
        
        # File uploader only - no persistence
        uploaded_file = st.file_uploader(
            "Upload Excel file:",
            type=['xlsx', 'xls']
        )
        
        data_uploaded = False
        
        if uploaded_file is not None:
            try:
                # Use temporary file for processing
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                df = read_excel(tmp_file_path)
                
                if df is not None and not df.empty:
                    st.session_state.excel_data = df
                    st.success(f"✅ {len(df)} records loaded from {uploaded_file.name}")
                    data_uploaded = True
                
                # Clean up temp file
                os.unlink(tmp_file_path)
                
            except Exception as e:
                st.error(f"Error uploading file: {e}")
        else:
            if st.session_state.excel_data is not None:
                data_uploaded = True
                st.info(f"✅ {len(st.session_state.excel_data)} records loaded")
        
        # Show data preview if available
        if st.session_state.excel_data is not None:
            with st.expander("📊 Preview Data"):
                st.dataframe(st.session_state.excel_data.head(3), use_container_width=True)
    
    with col2:
        st.subheader("Email Settings")
        saved_user_settings = getattr(st.session_state, 'saved_user_settings', None)
        default_from = saved_user_settings.get('from_email', '') if saved_user_settings else ""
        default_subject = saved_user_settings.get('subject', '') if saved_user_settings else ""
        
        from_email = st.text_input(
            "From Email",
            value=default_from,
            placeholder="your-email@domain.com"
        )
        
        subject = st.text_input(
            "Subject",
            value=default_subject,
            placeholder="Email subject"
        )
        
        email_settings_valid = bool(from_email and '@' in from_email and subject)
        if email_settings_valid:
            st.session_state.email_settings = {
                'from_email': from_email,
                'subject': subject
            }
    
    with col3:
        st.subheader("Send Emails")
        
        # Show authentication method info
        if st.session_state.email_settings_config:
            from config.settings import USE_OAUTH
            auth_info = "🔐 OAuth 2.0" if USE_OAUTH else "🔑 Password"

        
        if (data_uploaded and email_settings_valid and 
            hasattr(st.session_state, 'current_email_template')):
            
            df = st.session_state.excel_data
            template = st.session_state.current_email_template
            
            if st.button("🧪 Test Email", type="secondary", use_container_width=True):
                # Update templates with current editor content
                update_current_templates()
                template = st.session_state.current_email_template
                execute_mail_merge(df.head(1), template, "Test Mode", 0.1)
            
            # Add some space between buttons
            st.write("")
            
            if st.button("🚀 Send All Emails", type="primary", use_container_width=True):
                st.session_state.show_send_confirmation = True
            
            # Show confirmation dialog
            if st.session_state.get('show_send_confirmation', False):
                if show_send_confirmation_dialog(df, template, "Live Mode"):
                    # Update templates with current editor content
                    update_current_templates()
                    template = st.session_state.current_email_template
                    execute_mail_merge(df, template, "Live Mode", 0.5)
                    st.session_state.show_send_confirmation = False
        else:
            st.info("Complete setup first")
    
    # Persistent Debug Log Section (separate from results so it doesn't get cleared)
    if st.session_state.get('debug_log') and len(st.session_state.debug_log) > 0:
        st.header("🔍 Debug Information")
        with st.expander("📋 **Mail Merge Debug Log** (Persistent)", expanded=True):
            for log_entry in st.session_state.debug_log:
                st.markdown(log_entry)
            
            # Add button to clear debug log manually
            if st.button("🗑️ Clear Debug Log", key="clear_debug_log"):
                st.session_state.debug_log = []
                st.rerun()
    
    # Results section (if any)
    if st.session_state.sent_emails:
        st.header("📈 Results")
        sent_df = pd.DataFrame(st.session_state.sent_emails)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total", len(sent_df))
        with col2:
            successful = len(sent_df[sent_df['status'] == 'Success'])
            st.metric("Success", successful)
        with col3:
            failed = len(sent_df[sent_df['status'] == 'Failed'])
            st.metric("Failed", failed)
        with col4:
            if st.button("🗑️ Clear Results", use_container_width=True):
                st.session_state.sent_emails = []
                st.rerun()
    
    # Template editors row
    st.header("📝 Templates")
    
    # Show variables once for both templates
    if email_settings_valid:
        available_vars = show_available_variables()
        if available_vars:
            create_variable_interface(available_vars, "shared_vars")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("✏️ Email Template")
        if email_settings_valid:
            # Default email template content
            default_email_template = """Dear {First Name} {Last Name},

Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.

Please do not hesitate to contact me if you have any questions on the attached document.

Regards,
Nagarathar Sangam of North America Treasurer,
Email: nsnatreasurer@achi.org"""
            
            # Check if user has a saved template
            saved_email = getattr(st.session_state, 'saved_email_template', None)
            has_saved_template = saved_email and saved_email.get('content')
            
            # Determine template content: use saved template if available, otherwise use default
            if has_saved_template:
                template_content = saved_email.get('content')
                st.info("📁 Using your saved template from user directory")
            else:
                template_content = default_email_template
                st.info("📄 Using default template (will be saved to your user directory when you save)")
                
                # Auto-save default template to user directory on first use
                if 'auto_saved_default_email' not in st.session_state:
                    st.session_state.current_email_template = {
                        'content': default_email_template,
                        'from_email': from_email,
                        'subject': subject
                    }
                    save_email_template()
                    st.session_state.auto_saved_default_email = True
            
            # Get available variables for the editor
            available_vars = []
            if st.session_state.excel_data is not None:
                available_vars.extend(st.session_state.excel_data.columns.tolist())
            # Add system variables
            available_vars.extend(["Current Year", "Current Date"])
            
            # Use streamlit-quill editor with embedded toolbar
            email_content = rich_text_editor(
                value=template_content,
                placeholder="Enter your email template here...",
                key="email_template_quill",
                height=300,
                available_variables=available_vars
            )
            
            # Store the current email content in session state for mail merge access
            # Use a different key to avoid Streamlit widget key conflict
            if email_content is not None:
                st.session_state.current_email_content = email_content
            
            # Always update the current template with editor content, even if empty
            if email_content is not None:  # Check for None instead of strip()
                st.session_state.current_email_template = {
                    'content': email_content,
                    'from_email': from_email,
                    'subject': subject
                }
                
                if st.button("💾 Save Email Template", use_container_width=True):
                    save_email_template()
                
                # Add debug button to check template content
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("🔍 Debug Template Content", use_container_width=True):
                        with st.expander("🔍 **Template Debug Information**", expanded=True):
                            st.write("**Session State Keys:**")
                            relevant_keys = [k for k in st.session_state.keys() if 'template' in k.lower() or 'email' in k.lower()]
                            for key in relevant_keys:
                                value = st.session_state[key]
                                if isinstance(value, str):
                                    st.write(f"- `{key}`: {len(value)} chars - '{value[:50]}{'...' if len(value) > 50 else ''}'")
                                elif isinstance(value, dict):
                                    content_len = len(value.get('content', '')) if value.get('content') else 0
                                    st.write(f"- `{key}`: dict with content {content_len} chars")
                                else:
                                    st.write(f"- `{key}`: {type(value).__name__} = {value}")
                            
                            st.write("**Current Editor Content:**")
                            st.write(f"- **Returned by editor:** '{email_content[:100] if email_content else 'None or Empty'}{'...' if email_content and len(email_content) > 100 else ''}'")
                            st.write(f"- **Length:** {len(email_content) if email_content else 0} characters")
                            st.write(f"- **Stored in session state (current_email_content):** {len(st.session_state.get('current_email_content', '')) if st.session_state.get('current_email_content') else 0} characters")
                            
                            st.write("**Template Sources:**")
                            st.write(f"- **template_content (initial):** {len(template_content)} chars")
                            st.write(f"- **email_content (from editor):** {len(email_content) if email_content else 0} chars")
                            
                            if email_content:
                                st.write("**Raw Editor Content (first 500 chars):**")
                                st.code(email_content[:500] + ('...' if len(email_content) > 500 else ''))
                            else:
                                st.error("❌ Email editor is returning empty content!")
                
                with col2:
                    # Show preview of current template
                    if st.button("👁️ Preview Template", use_container_width=True):
                        pass  # Content moved below
                
                # Show preview of current template
                with st.expander("📧 Preview Email Template"):
                    if st.session_state.excel_data is not None and len(st.session_state.excel_data) > 0:
                        sample_row = st.session_state.excel_data.iloc[0]
                        preview_content = email_content
                        
                        # Replace system variables first
                        preview_content = replace_system_variables(preview_content)
                        
                        # Then replace Excel data variables
                        for col in st.session_state.excel_data.columns:
                            preview_content = preview_content.replace(f"{{{col}}}", str(sample_row[col]))
                        
                        st.markdown("**Preview with sample data:**")
                        # Use clean_html_content to ensure proper HTML formatting
                        clean_preview = clean_html_content(preview_content)
                        st.markdown(clean_preview, unsafe_allow_html=True)
                    else:
                        # Show preview with just system variables
                        preview_content = replace_system_variables(email_content)
                        st.markdown("**Preview with system variables:**")
                        clean_preview = clean_html_content(preview_content)
                        st.markdown(clean_preview, unsafe_allow_html=True)
                        st.info("Upload data to see preview with Excel variables")
        else:
            st.info("Set email settings first")
    
    with col2:
        st.subheader("📄 PDF Template")
        
        # Check for letterhead
        if CLOUD_DEPLOYMENT:
            letterhead_path = get_letterhead_path()
        else:
            letterhead_paths = [
                Path('NSNA Atlanta Letterhead Updated.pdf'),
                Path('src/templates/letterhead_template.pdf')
            ]
            
            letterhead_path = None
            for path in letterhead_paths:
                if path.exists():
                    letterhead_path = path
                    break
        
        if letterhead_path:
            # Check if user has a saved PDF template
            saved_pdf = getattr(st.session_state, 'saved_pdf_template', None)
            has_saved_pdf_template = saved_pdf and saved_pdf.get('content')
            
            
            # Determine PDF template content based on selection
            if has_saved_pdf_template:
                default_pdf_content = saved_pdf.get('content')
                st.info("� Using your saved PDF template from user directory")
            else:
                # Default PDF template content
                default_pdf_template = """<p>Dear {First Name} {Last Name},</p>

<p>On behalf of Nagarathar Sangam of North America, thank you for your recent donation to our organization.</p>

<p>This letter will serve as a tax receipt for your contribution listed below.</p>

<p><strong>Donation Amount: ${Donation Amount}</strong><br>
<strong>Donation(s) Year: {Current Year}</strong></p>

<p>No goods or services were provided in exchange for your contribution.</p>

<p>Nagarathar Sangam of North America is a registered Section 501(c)(3) non-profit organization (EIN #: 22-3974176). For questions regarding this acknowledgment letter, please contact us at treasurer@achi.org.</p>

<p>We truly appreciate your donation and look forward to your continued support of our mission.</p>

<p>Sincerely,<br>
Treasurer, NSNA<br>
(2025-2026 term)</p>"""
                default_pdf_content = default_pdf_template
                st.info("📄 Using default PDF template (will be saved to your user directory when you save)")
                
                # Auto-save default PDF template to user directory on first use
                if 'auto_saved_default_pdf' not in st.session_state:
                    st.session_state.current_pdf_template = {
                        'content': default_pdf_template
                    }
                    save_pdf_template()
                    st.session_state.auto_saved_default_pdf = True
            
            # Initialize session state for PDF editor
            if f'pdf_content_main_content' not in st.session_state:
                st.session_state[f'pdf_content_main_content'] = default_pdf_content
                st.session_state[f'pdf_content_main_overwrite'] = True
            
            # Update content if template changed
            if st.session_state.get(f'pdf_content_main_content') != default_pdf_content:
                st.session_state[f'pdf_content_main_content'] = default_pdf_content
                st.session_state[f'pdf_content_main_overwrite'] = True
            
            # Get available variables for the enhanced editor
            available_vars = []
            if st.session_state.excel_data is not None:
                available_vars.extend(st.session_state.excel_data.columns.tolist())
            # Add system variables
            available_vars.extend(["Current Year", "Current Date"])
            
            # Use compact rich text editor for PDF template (same as email)
            pdf_content = rich_text_editor(
                value=st.session_state[f'pdf_content_main_content'],
                placeholder="Enter your PDF template here...",
                key="pdf_content_main",
                height=350,
                overwrite=st.session_state.get(f'pdf_content_main_overwrite', False),
                available_variables=available_vars
            )
            st.session_state[f'pdf_content_main_overwrite'] = False
            
            # Always update the current template with editor content
            if pdf_content is not None:
                st.session_state.current_pdf_template = {
                    'content': pdf_content,
                    'template_path': str(letterhead_path)
                }
                
                subcol1, subcol2 = st.columns(2)
                with subcol1:
                    if st.button("💾 Save PDF Template", use_container_width=True):
                        save_pdf_template()
                
                with subcol2:
                    if st.button("📄 Download Sample", use_container_width=True):
                        generate_sample_pdf(pdf_content, letterhead_path)
                
                # Show preview of current PDF template
                with st.expander("📄 Preview PDF Template"):
                    if st.session_state.excel_data is not None and len(st.session_state.excel_data) > 0:
                        sample_row = st.session_state.excel_data.iloc[0]
                        preview_content = pdf_content
                        
                        # Replace system variables first
                        preview_content = replace_system_variables(preview_content)
                        
                        # Then replace Excel data variables
                        for col in st.session_state.excel_data.columns:
                            preview_content = preview_content.replace(f"{{{col}}}", str(sample_row[col]))
                        
                        st.markdown("**Preview with sample data (formatted as it will appear in PDF):**")
                        # Show the HTML content with formatting preserved instead of converting to plain text
                        clean_preview = clean_html_content(preview_content)
                        st.markdown(clean_preview, unsafe_allow_html=True)
                        st.info("💡 This preview shows the formatted content. Bold and italic formatting will be preserved in the generated PDF.")
                    else:
                        # Show preview with just system variables
                        preview_content = replace_system_variables(pdf_content)
                        st.markdown("**Preview with system variables (formatted as it will appear in PDF):**")
                        clean_preview = clean_html_content(preview_content)
                        st.markdown(clean_preview, unsafe_allow_html=True)
                        st.info("Upload data to see preview with Excel variables")
        else:
            st.warning("NSNA letterhead template not found")

def data_settings_tab():
    """Data & Settings Tab: Upload data and email settings only"""
    
    # Data Upload Section
    st.header("📊 Upload Data")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xls']
    )
    
    data_uploaded = False
    
    if uploaded_file is not None:
        try:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Read Excel data
            df = read_excel(tmp_file_path)
            
            if df is not None and not df.empty:
                st.session_state.excel_data = df
                st.success(f"✅ Loaded {len(df)} records")
                data_uploaded = True
                
                # Display data preview
                st.dataframe(df.head(3), use_container_width=True)
            
            # Clean up temporary file
            os.unlink(tmp_file_path)
            
        except Exception as e:
            st.error(f"Error reading file: {e}")
    else:
        if st.session_state.excel_data is not None:
            data_uploaded = True
    
    # Email Settings Section
    st.header("📧 Email Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Load saved settings if available
        saved_user_settings = getattr(st.session_state, 'saved_user_settings', None)
        default_from = saved_user_settings.get('from_email', '') if saved_user_settings else ''
            
        from_email = st.text_input(
            "From Email",
            value=default_from,
            placeholder="your-email@domain.com"
        )
        
    with col2:
        default_subject = ""
        if saved_user_settings:
            default_subject = saved_user_settings.get('subject', '')
            
        subject = st.text_input(
            "Subject",
            value=default_subject,
            placeholder="Email subject"
        )
    
    # Store email settings in session state
    email_settings_valid = bool(from_email and '@' in from_email and subject)
    if email_settings_valid:
        st.session_state.email_settings = {
            'from_email': from_email,
            'subject': subject
        }


def templates_sending_tab():
    """Templates & Sending Tab: Email template editor, PDF template editor, send buttons, and results"""
    
    # Check if setup is complete (simplified check)
    if (st.session_state.excel_data is None or 
        not hasattr(st.session_state, 'email_settings') or
        st.session_state.email_settings is None):
        st.info("Complete Data & Settings tab first")
        return
    
    # Send Buttons and Results at the top (not in subtab)
    st.header("📧 Send Emails")
    
    # Check if we have email template before showing send buttons
    if hasattr(st.session_state, 'current_email_template'):
        df = st.session_state.excel_data
        template = st.session_state.current_email_template
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🧪 Test Email", type="secondary"):
                execute_mail_merge(df.head(1), template, "Test Mode", 0.1)
        
        with col2:
            # Confirmation checkbox for sending all emails
            confirm_send = st.checkbox("⚠️ Confirm send all", key="confirm_send_tab")
            if st.button("🚀 Send All Emails", type="primary", disabled=not confirm_send):
                if confirm_send:
                    execute_mail_merge(df, template, "Live Mode", 0.5)
                    # Reset confirmation
                    st.session_state.confirm_send_tab = False
        
        # Results section (simplified)
        if st.session_state.sent_emails:
            sent_df = pd.DataFrame(st.session_state.sent_emails)
            
            # Simple metrics in one row
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total", len(sent_df))
            with col2:
                successful = len(sent_df[sent_df['status'] == 'Success'])
                st.metric("Success", successful)
            with col3:
                failed = len(sent_df[sent_df['status'] == 'Failed'])
                st.metric("Failed", failed)
            with col4:
                if st.button("🗑️ Clear", key="clear_results"):
                    st.session_state.sent_emails = []
                    st.rerun()
    else:
        st.info("Create an email template first using the subtabs below")
    
    # Template Editor Subtabs
    st.header("📝 Template Editors")
    
    subtab1, subtab2 = st.tabs([
        "✏️ Email Template", 
        "📄 PDF Template"
    ])
    
    with subtab1:
        email_template_subtab()
    
    with subtab2:
        pdf_template_subtab()


def email_template_subtab():
    """Email Template subtab content"""
    
    # Check if email settings exist
    if not hasattr(st.session_state, 'email_settings') or st.session_state.email_settings is None:
        st.warning("Please complete the Data & Settings tab first to set your email configuration.")
        return
    
    # Get email settings from session state
    from_email = st.session_state.email_settings.get('from_email', '')
    subject = st.session_state.email_settings.get('subject', '')
    
    # Default email template content
    default_email_template = """<p>Dear {First Name} {Last Name},</p>
<p>Thank you for your generous donation to the Nagarathar Sangam of North America.</p>
<p><strong>Donation Details:</strong><br>
• Amount: ${Amount}<br>
• Date: {Date}</p>
<p>Please find your official receipt attached as a PDF for your tax records.</p>
<p>Best regards,<br>
<strong>NSNA Treasurer</strong></p>"""
    
    # Load saved template if available, otherwise use default
    saved_email = getattr(st.session_state, 'saved_email_template', None)
    if saved_email:
        template_content = saved_email.get('content')
        st.info("📁 Using your saved template from user directory")
    else:
        template_content = default_email_template
        st.info("📄 Using default template (will be saved to your user directory when you save)")
        
        # Auto-save default template to user directory on first use
        if 'auto_saved_default_email_subtab' not in st.session_state:
            st.session_state.current_email_template = {
                'content': default_email_template,
                'from_email': from_email,
                'subject': subject
            }
            save_email_template()
            st.session_state.auto_saved_default_email_subtab = True
    
    # Email template editor
    # Initialize session state for email editor in subtab
    if f'email_template_subtab_content' not in st.session_state:
        st.session_state[f'email_template_subtab_content'] = template_content
        st.session_state[f'email_template_subtab_overwrite'] = True
    
    # Update content if template changed
    if st.session_state.get(f'email_template_subtab_content') != template_content:
        st.session_state[f'email_template_subtab_content'] = template_content
        st.session_state[f'email_template_subtab_overwrite'] = True
    
    email_content = rich_text_editor(
        value=st.session_state[f'email_template_subtab_content'],
        placeholder="Enter your email template here...",
        key="email_template_quill_subtab",
        height=300,
        overwrite=st.session_state.get(f'email_template_subtab_overwrite', False)
    )
    st.session_state[f'email_template_subtab_overwrite'] = False
    
    if email_content and email_content.strip():
        st.session_state.current_email_template = {
            'content': email_content,
            'from_email': from_email,
            'subject': subject
        }
        
        # Save template button
        if st.button("� Save Template", key="save_email_template"):
            save_email_template()

def pdf_template_subtab():
    """PDF Template subtab content"""
    
    # Check for letterhead
    if CLOUD_DEPLOYMENT:
        letterhead_path = get_letterhead_path()
    else:
        letterhead_paths = [
            Path('NSNA Atlanta Letterhead Updated.pdf'),
            Path('src/templates/letterhead_template.pdf')
        ]
        
        letterhead_path = None
        for path in letterhead_paths:
            if path.exists():
                letterhead_path = path
                break
    
    if letterhead_path:
        # Default PDF content
        saved_pdf = getattr(st.session_state, 'saved_pdf_template', None)
        default_pdf_content = saved_pdf.get('content', """Dear {First Name} {Last Name},

Thank you for your generous donation to the Nagarathar Sangam of North America.

**Donation Details:**
• Amount: ${Amount}
• Date: {Date}
• Tax Year: {Year}

This receipt serves as official documentation of your charitable contribution.

Best regards,
NSNA Treasurer""") if saved_pdf else """Dear {First Name} {Last Name},

Thank you for your generous donation to the Nagarathar Sangam of North America.

**Donation Details:**
• Amount: ${Amount}
• Date: {Date}
• Tax Year: {Year}

This receipt serves as official documentation of your charitable contribution.

Best regards,
NSNA Treasurer"""
        
        # PDF content editor - use text area instead of rich editor for plain text
        pdf_content = st.text_area(
            "PDF Receipt Content",
            value=default_pdf_content,
            placeholder="Create your PDF receipt content here...",
            height=300,
            key="pdf_content_text"
        )
        
        if pdf_content and pdf_content.strip():
            st.session_state.current_pdf_template = {
                'content': pdf_content,
                'template_path': str(letterhead_path)
            }
            
            # Save and sample buttons at bottom
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 Save Template", key="save_pdf_template"):
                    save_pdf_template()
            
            with col2:
                if st.button("📄 Download Sample PDF", key="download_sample_pdf"):
                    generate_sample_pdf(pdf_content, letterhead_path)
        else:
            st.warning("NSNA letterhead template not found")

def save_email_template():
    """Save email template"""
    try:
        template = st.session_state.current_email_template
        persistence = st.session_state.template_persistence
        
        # Save the email template
        persistence.save_email_template(template)
        
        # Save user settings
        user_settings = {
            'from_email': template['from_email'],
            'subject': template['subject']
        }
        persistence.save_user_settings(user_settings)
        
        st.success("✅ Email template saved successfully!")
        return True
    except Exception as e:
        st.error(f"❌ Failed to save email template: {e}")
        return False

def save_pdf_template():
    """Save PDF template"""
    try:
        template = st.session_state.current_pdf_template
        persistence = st.session_state.template_persistence
        
        # Save the PDF template
        persistence.save_pdf_template(template)
        
        st.success("✅ PDF template saved successfully!")
        return True
    except Exception as e:
        st.error(f"❌ Failed to save PDF template: {e}")
        return False

def generate_sample_pdf(pdf_content, letterhead_path):
    """Generate a sample PDF for preview"""
    try:
        with st.spinner("Generating sample PDF..."):
            # Generate PDF with sample data
            pdf_generator = PDFGenerator()
            
            # Create sample data
            sample_data = {
                'First Name': 'John',
                'Last Name': 'Smith',
                'Email': 'john.smith@email.com',
                'Amount': '150.00',
                'Date': datetime.now().strftime('%Y-%m-%d'),
                'Year': datetime.now().strftime('%Y'),
            }
            
            # Process PDF content with sample data
            pdf_content_with_sample = pdf_content
            
            # Replace system variables first
            pdf_content_with_sample = replace_system_variables(pdf_content_with_sample)
            
            # Replace sample data variables
            for key, value in sample_data.items():
                pdf_content_with_sample = pdf_content_with_sample.replace(f"{{{key}}}", str(value))
            
            # Pass HTML content directly to PDF generator for proper formatting
            sample_data['content'] = pdf_content_with_sample
            sample_data['template_path'] = str(letterhead_path)
            
            # Generate the PDF
            pdf_path = pdf_generator.generate_receipt(sample_data)
            
            if pdf_path and os.path.exists(pdf_path):
                # Read the PDF file
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_bytes = pdf_file.read()
                
                # Show success message
                st.success("✅ Sample PDF generated successfully!")
                
                # Provide download button
                st.download_button(
                    label="📥 Download Sample PDF Receipt",
                    data=pdf_bytes,
                    file_name=f"NSNA_Sample_Receipt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    help="Download the sample PDF to verify letterhead and formatting",
                    type="primary"
                )
                
                # Clean up temporary file
                try:
                    os.unlink(pdf_path)
                except:
                    pass
            else:
                st.error("❌ Failed to generate sample PDF")
    except Exception as e:
        st.error(f"❌ Error generating sample PDF: {e}")

def preview_templates_section():
    """Preview email and PDF templates"""
    st.subheader("�📧 Template Preview")
    
    if hasattr(st.session_state, 'current_email_template'):
        template = st.session_state.current_email_template
        
        # Create sample data for preview
        sample_data = create_sample_data()
        
        # Generate preview content
        preview_content = template['content']
        for var, value in sample_data.items():
            preview_content = preview_content.replace(f"{{{var}}}", str(value))
        
        # Convert markdown to HTML for preview (same as PDF generator)
        if '**' in preview_content or '*' in preview_content:
            # Convert markdown formatting to HTML for proper preview display
            preview_content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', preview_content)
            preview_content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', preview_content)
        
        # Email Preview Section
        st.markdown("### 📧 Email Preview")
        st.markdown("**Email Subject:** " + template['subject'])
        st.markdown("**From:** " + template['from_email'])
        st.markdown("---")
        st.markdown(preview_content, unsafe_allow_html=True)
        
        # PDF Preview Section
        if hasattr(st.session_state, 'current_pdf_template'):
            st.markdown("---")
            st.markdown("### � PDF Preview")
            
            pdf_template = st.session_state.current_pdf_template
            pdf_content = pdf_template.get('content', 'No PDF content created yet.')
            
            # Replace sample data in PDF content
            pdf_preview = pdf_content
            for var, value in sample_data.items():
                pdf_preview = pdf_preview.replace(f"{{{var}}}", str(value))
            
            # Convert markdown to HTML for PDF preview display
            if '**' in pdf_preview or '*' in pdf_preview:
                # Convert markdown formatting to HTML for proper preview display
                pdf_preview = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', pdf_preview)
                pdf_preview = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', pdf_preview)
            
            st.markdown("**PDF Content Preview:**")
            st.markdown(pdf_preview, unsafe_allow_html=True)
            st.caption("� This preview shows how the PDF will look with sample data.")
    else:
        st.info("Please create your email template first")

def create_sample_data():
    """Create sample data for previews"""
    sample_data = {}
    available_vars = list(st.session_state.excel_data.columns) if st.session_state.excel_data is not None else []
    
    if available_vars:
        for var in available_vars:
            if 'name' in var.lower() and 'first' in var.lower():
                sample_data[var] = "John"
            elif 'name' in var.lower() and 'last' in var.lower():
                sample_data[var] = "Smith"
            elif 'email' in var.lower():
                sample_data[var] = "john.smith@email.com"
            elif 'amount' in var.lower():
                sample_data[var] = "150.00"
            elif 'date' in var.lower():
                sample_data[var] = datetime.now().strftime("%Y-%m-%d")
            elif 'year' in var.lower():
                sample_data[var] = datetime.now().strftime("%Y")
            else:
                sample_data[var] = f"Sample {var}"
    
    # Add standard variables
    sample_data.update({
        'Amount': '150.00',
        'Year': datetime.now().strftime('%Y'),
        'Date': datetime.now().strftime('%Y-%m-%d')
    })
    
    return sample_data

# Helper functions for template management and mail merge operations

def execute_mail_merge(df, template, send_mode, delay_between_emails):
    """Execute the mail merge process - always use current editor content"""
    
    # Initialize debug log in session state
    if 'debug_log' not in st.session_state:
        st.session_state.debug_log = []
    
    # Don't clear the debug log - append to it for persistence
    # Add separator for new mail merge session
    st.session_state.debug_log.append(f"\n{'='*50}")
    st.session_state.debug_log.append(f"🚀 **NEW MAIL MERGE SESSION: {send_mode}**")
    st.session_state.debug_log.append(f"⏰ **Timestamp:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.session_state.debug_log.append(f"{'='*50}")
    
    # Create containers that won't be refreshed
    debug_container = st.container()
    progress_container = st.container()
    
    with progress_container:
        # Show progress at the current location
        progress_bar = st.progress(0)
        status_text = st.empty()
    
    successful_sends = 0
    failed_sends = 0
    
    try:
        # Debug: Show what templates we're working with
        st.session_state.debug_log.append("🔍 **TEMPLATE DEBUG INFO:**")
        
        # Get the latest content from editors instead of using passed template
        latest_email_content = st.session_state.get('current_email_content', template.get('content', ''))
        latest_pdf_content = st.session_state.get('pdf_content_main', '')
        
        st.session_state.debug_log.append(f"- **Template from parameter:** {len(template.get('content', '')) if template.get('content') else 0} characters")
        st.session_state.debug_log.append(f"- **Latest email content from session:** {len(latest_email_content) if latest_email_content else 0} characters")
        st.session_state.debug_log.append(f"- **Latest PDF content from session:** {len(latest_pdf_content) if latest_pdf_content else 0} characters")
        
        # Additional debug: Show all template-related session state keys
        template_keys = {k: v for k, v in st.session_state.items() if 'template' in k.lower() or 'email' in k.lower()}
        st.session_state.debug_log.append("- **Template-related session state keys:**")
        for key, value in template_keys.items():
            if isinstance(value, str):
                st.session_state.debug_log.append(f"  - `{key}`: {len(value)} chars")
            elif isinstance(value, dict):
                content_len = len(value.get('content', '')) if value.get('content') else 0
                st.session_state.debug_log.append(f"  - `{key}`: dict with content {content_len} chars")
            else:
                st.session_state.debug_log.append(f"  - `{key}`: {type(value).__name__}")
        
        # Ensure content is not None
        if latest_email_content is None:
            latest_email_content = ''
            st.session_state.debug_log.append("⚠️ **WARNING:** Email content was None, set to empty string")
        if latest_pdf_content is None:
            latest_pdf_content = ''
            st.session_state.debug_log.append("⚠️ **WARNING:** PDF content was None, set to empty string")
        
        # Debug: Check if email content is actually empty
        if not latest_email_content or latest_email_content.strip() == "":
            st.session_state.debug_log.append("❌ **CRITICAL ISSUE:** Email content is empty or whitespace only!")
            st.session_state.debug_log.append(f"   - Raw content: '{latest_email_content}'")
            st.session_state.debug_log.append(f"   - Session state current_email_content: '{st.session_state.get('current_email_content', 'NOT FOUND')}'")
            
            # Try to get content from current template as fallback
            if hasattr(st.session_state, 'current_email_template') and st.session_state.current_email_template:
                fallback_content = st.session_state.current_email_template.get('content', '')
                if fallback_content:
                    latest_email_content = fallback_content
                    st.session_state.debug_log.append(f"✅ **FALLBACK:** Using content from current_email_template: {len(fallback_content)} characters")
                else:
                    st.session_state.debug_log.append("❌ **FALLBACK FAILED:** current_email_template content is also empty")
        
        # Use email config from current template
        from_email = template['from_email']
        subject_template = template.get('subject', 'NSNA Donation Receipt')
        
        # Ensure subject is not None
        if subject_template is None:
            subject_template = 'NSNA Donation Receipt'
        
        st.session_state.debug_log.append(f"- **From Email:** {from_email}")
        st.session_state.debug_log.append(f"- **Subject Template:** {subject_template}")
        
        # Create updated template with latest content
        current_template = {
            'content': latest_email_content,
            'from_email': from_email,
            'subject': subject_template
        }
        
        # Import OAuth settings from config
        from config.settings import GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET
        
        # Create SMTP settings for OAuth authentication
        smtp_settings = {
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'smtp_user_name': from_email,
            'use_oauth': True,
            'dynamic_user_oauth': True,
            'client_id': GOOGLE_CLIENT_ID,
            'client_secret': GOOGLE_CLIENT_SECRET
        }
        
        # Show authentication method to user
        status_text.text(f"🔐 Authenticating with OAuth 2.0 for {from_email}...")
        st.session_state.debug_log.append(f"🔐 **AUTHENTICATION:** OAuth 2.0 for {from_email}")

        total_emails = len(df)
        
        for index, row in df.iterrows():
            try:
                # Update progress
                progress = (index + 1) / total_emails
                progress_bar.progress(progress)
                status_text.text(f"Processing {index + 1}/{total_emails}: {row['First Name']} {row['Last Name']}")
                
                # Add debug info for this recipient
                recipient_name = f"{row['First Name']} {row['Last Name']}"
                st.session_state.debug_log.append(f"\n📧 **PROCESSING: {recipient_name}**")
                
                # Generate personalized content using current template
                row_dict = row.to_dict()
                personalized_content = current_template['content']
                personalized_subject = subject_template
                
                st.session_state.debug_log.append(f"- **Original content length:** {len(personalized_content) if personalized_content else 0}")
                st.session_state.debug_log.append(f"- **Original subject:** {personalized_subject}")
                
                # Ensure content is not None before processing
                if personalized_content is None:
                    personalized_content = ''
                    st.session_state.debug_log.append("⚠️ **WARNING:** Personalized content was None")
                if personalized_subject is None:
                    personalized_subject = 'NSNA Donation Receipt'
                    st.session_state.debug_log.append("⚠️ **WARNING:** Personalized subject was None")
                
                # First replace system variables
                personalized_content = replace_system_variables(personalized_content)
                personalized_subject = replace_system_variables(personalized_subject)
                
                # Then replace Excel data variables
                for col, value in row_dict.items():
                    if personalized_content:  # Only if not empty
                        personalized_content = personalized_content.replace(f"{{{col}}}", str(value))
                    if personalized_subject:  # Only if not empty
                        personalized_subject = personalized_subject.replace(f"{{{col}}}", str(value))
                
                # Clean HTML content for email - preserve HTML formatting
                email_content = clean_html_content(personalized_content)
                
                st.session_state.debug_log.append(f"- **Final email content length:** {len(email_content) if email_content else 0}")
                
                # Add detailed content debugging
                if email_content:
                    st.session_state.debug_log.append(f"- **Final email content preview:** '{email_content[:200]}{'...' if len(email_content) > 200 else ''}'")
                else:
                    st.session_state.debug_log.append("- **Final email content is empty!**")
                
                # Check for remaining unreplaced variables
                remaining_vars = re.findall(r'\{[^}]+\}', email_content + " " + personalized_subject)
                if remaining_vars:
                    st.session_state.debug_log.append(f"⚠️ **UNREPLACED VARIABLES:** {', '.join(remaining_vars)}")
                
                # Validate email content before sending
                if not email_content or email_content.strip() == "":
                    st.session_state.debug_log.append("❌ **CRITICAL ERROR:** Final email content is empty!")
                    st.session_state.debug_log.append(f"   - Raw personalized content: '{personalized_content}'")
                    st.session_state.debug_log.append(f"   - Cleaned email content: '{email_content}'")
                    # Skip this email
                    failed_sends += 1
                    st.session_state.sent_emails.append({
                        'name': recipient_name,
                        'email': row['Email'] if send_mode == "Live Mode" else from_email,
                        'status': 'Failed - Empty Content',
                        'timestamp': datetime.now()
                    })
                    continue
                
                # Generate PDF receipt if template exists
                pdf_path = None
                if latest_pdf_content and latest_pdf_content.strip():
                    try:
                        pdf_generator = PDFGenerator()
                        pdf_data = row_dict.copy()
                        
                        # Use latest PDF content from editor
                        pdf_content = latest_pdf_content
                        
                        # Ensure PDF content is not None
                        if pdf_content is None:
                            pdf_content = ''
                        
                        # For PDF: First replace variables in HTML, then convert to plain text
                        pdf_content_with_vars = pdf_content
                        
                        # First replace system variables (only if content exists)
                        if pdf_content_with_vars:
                            pdf_content_with_vars = replace_system_variables(pdf_content_with_vars)
                            
                            # Then replace Excel data variables
                            for col, value in row_dict.items():
                                if pdf_content_with_vars:  # Check again after system variable replacement
                                    pdf_content_with_vars = pdf_content_with_vars.replace(f"{{{col}}}", str(value))
                        
                        # Pass HTML content directly to PDF generator for proper formatting
                        pdf_data['content'] = pdf_content_with_vars
                        
                        # Get letterhead path
                        letterhead_path = None
                        if CLOUD_DEPLOYMENT:
                            letterhead_path = get_letterhead_path()
                        else:
                            letterhead_paths = [
                                Path('NSNA Atlanta Letterhead Updated.pdf'),
                                Path('src/templates/letterhead_template.pdf')
                            ]
                            for path in letterhead_paths:
                                if path.exists():
                                    letterhead_path = path
                                    break
                        
                        if letterhead_path:
                            pdf_data['template_path'] = str(letterhead_path)
                            pdf_path = pdf_generator.generate_receipt(pdf_data)
                            st.session_state.debug_log.append(f"✅ **PDF GENERATED:** {pdf_path}")
                        else:
                            st.session_state.debug_log.append("⚠️ **WARNING:** No letterhead found, PDF not generated")
                    except Exception as pdf_error:
                        st.session_state.debug_log.append(f"❌ **PDF ERROR:** {pdf_error}")
                
                # Determine recipient email
                recipient_email = row['Email'] if send_mode == "Live Mode" else from_email
                st.session_state.debug_log.append(f"- **Recipient Email:** {recipient_email}")
                st.session_state.debug_log.append(f"- **Has PDF Attachment:** {'Yes' if pdf_path else 'No'}")
                
                # Send email with enhanced error handling
                st.session_state.debug_log.append(f"🚀 **ATTEMPTING TO SEND EMAIL**")
                st.session_state.debug_log.append(f"   - **From:** {from_email}")
                st.session_state.debug_log.append(f"   - **To:** {recipient_email}")
                st.session_state.debug_log.append(f"   - **Subject:** {personalized_subject}")
                st.session_state.debug_log.append(f"   - **Content Length:** {len(email_content)} chars")
                st.session_state.debug_log.append(f"   - **Has Attachment:** {'Yes' if pdf_path else 'No'}")
                
                try:
                    success = send_email_debug_wrapper(
                        from_email=from_email,
                        to_email=recipient_email,
                        subject=personalized_subject,
                        html_content=email_content,
                        attachment_path=pdf_path,
                        smtp_settings=smtp_settings,
                        is_test=(send_mode == "Test Mode")
                    )
                    
                    # Additional check to ensure success is boolean
                    if success is True:
                        st.session_state.debug_log.append("✅ **EMAIL SENT SUCCESSFULLY**")
                    elif success is False:
                        st.session_state.debug_log.append("❌ **EMAIL SEND FAILED** (send_email returned False)")
                        st.session_state.debug_log.append("   - This usually indicates authentication or SMTP configuration issues")
                    else:
                        st.session_state.debug_log.append(f"⚠️ **UNEXPECTED RETURN VALUE:** send_email returned {type(success).__name__}: {success}")
                        success = False
                        
                except Exception as email_error:
                    st.session_state.debug_log.append(f"❌ **EMAIL ERROR:** {str(email_error)}")
                    st.session_state.debug_log.append(f"   - **Error Type:** {type(email_error).__name__}")
                    
                    # Additional debugging for common issues
                    error_str = str(email_error).lower()
                    if "authentication" in error_str or "oauth" in error_str:
                        st.session_state.debug_log.append("🔐 **DIAGNOSIS:** Authentication Issue - Check OAuth credentials")
                    elif "smtp" in error_str or "connection" in error_str:
                        st.session_state.debug_log.append("📡 **DIAGNOSIS:** SMTP/Connection Issue - Check internet connection")
                    elif "attachment" in error_str:
                        st.session_state.debug_log.append("📎 **DIAGNOSIS:** Attachment Issue - Check PDF generation")
                    elif "recipient" in error_str or "email" in error_str:
                        st.session_state.debug_log.append("📧 **DIAGNOSIS:** Email Address Issue - Check recipient email format")
                    
                    success = False
                
                if success:
                    successful_sends += 1
                    st.session_state.sent_emails.append({
                        'name': recipient_name,
                        'email': recipient_email,
                        'status': 'Success',
                        'timestamp': datetime.now()
                    })
                else:
                    failed_sends += 1
                    st.session_state.sent_emails.append({
                        'name': recipient_name,
                        'email': recipient_email,
                        'status': 'Failed',
                        'timestamp': datetime.now()
                    })
                
                # Add delay between emails
                time.sleep(delay_between_emails)
                
            except Exception as e:
                failed_sends += 1
                st.session_state.debug_log.append(f"❌ **PROCESSING ERROR for {recipient_name}:** {e}")
                st.error(f"Failed to process {row['First Name']} {row['Last Name']}: {e}")
        
        # Final results
        progress_bar.progress(1.0)
        status_text.text("✅ Mail merge complete!")
        
        # Show simple results summary
        if successful_sends > 0:
            st.success(f"🎉 Successfully sent {successful_sends} emails!")
        if failed_sends > 0:
            st.error(f"❌ {failed_sends} emails failed")
            
        # Display persistent debug log
        with debug_container:
            with st.expander("🔍 **Detailed Debug Log**", expanded=True):
                for log_entry in st.session_state.debug_log:
                    st.markdown(log_entry)
            
        # Auto-refresh to show results
        time.sleep(0.5)
        st.rerun()
    
    except Exception as e:
        st.error(f"Mail merge failed: {e}")
        st.session_state.debug_log.append(f"❌ **MAIL MERGE FAILED:** {e}")
        status_text.text("❌ Mail merge failed!")
        progress_bar.empty()
        
        # Display debug log even on failure
        with debug_container:
            with st.expander("🔍 **Detailed Debug Log**", expanded=True):
                for log_entry in st.session_state.debug_log:
                    st.markdown(log_entry)

def convert_markdown_to_html(content):
    """Convert markdown formatting to HTML for better preview display"""
    if not content:
        return ""
    
    # Convert markdown bold to HTML
    content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', content)
    
    # Convert markdown italic to HTML (but not double asterisks)
    content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', content)
    
    # Convert bullet points
    content = re.sub(r'^• (.+)$', r'<li>\1</li>', content, flags=re.MULTILINE)
    content = re.sub(r'(<li>.*</li>)', r'<ul>\1</ul>', content, flags=re.DOTALL)
    
    # Convert numbered lists
    content = re.sub(r'^(\d+)\. (.+)$', r'<li>\2</li>', content, flags=re.MULTILINE)
    
    # Convert line breaks - double newlines become paragraph breaks
    content = re.sub(r'\n\n+', '</p><p>', content)
    content = f'<p>{content}</p>'
    
    # Convert single newlines to <br> tags
    content = content.replace('\n', '<br>')
    
    # Clean up empty paragraphs
    content = re.sub(r'<p>\s*</p>', '', content)
    
    return content

def clean_html_content(html_content):
    """Clean HTML content for email sending while preserving formatting"""
    if not html_content:
        return ""
    
    # Handle Quill editor HTML output (which includes comprehensive formatting)
    if ('<p>' in html_content or '<div>' in html_content or '<br>' in html_content or 
        '<strong>' in html_content or '<em>' in html_content or '<span>' in html_content):
        # Already HTML from Quill editor, clean and standardize it
        cleaned = html_content
        
        # Clean up Quill-specific formatting
        import re
        
        # Remove empty paragraphs
        cleaned = re.sub(r'<p>\s*<br>\s*</p>', '', cleaned)
        cleaned = re.sub(r'<p>\s*</p>', '', cleaned)
        
        # Add paragraph styling to match PDF output (normal line spacing)
        cleaned = re.sub(r'<p>', r'<p style="margin: 4px 0; line-height: 1.3; font-family: Helvetica, Arial, sans-serif; font-size: 11pt;">', cleaned)
        
        # Ensure consistent bold/italic tags
        cleaned = re.sub(r'<(strong|b)([^>]*)>', r'<strong\2>', cleaned)
        cleaned = re.sub(r'</(strong|b)>', r'</strong>', cleaned)
        cleaned = re.sub(r'<(em|i)([^>]*)>', r'<em\2>', cleaned)
        cleaned = re.sub(r'</(em|i)>', r'</em>', cleaned)
        
        # Handle strikethrough
        cleaned = re.sub(r'<(s|strike)([^>]*)>', r'<s\2>', cleaned)
        cleaned = re.sub(r'</(s|strike)>', r'</s>', cleaned)
        
        # Handle underline
        cleaned = re.sub(r'<u([^>]*)>', r'<u\1>', cleaned)
        
        # Handle headers
        for i in range(1, 7):
            cleaned = re.sub(f'<h{i}([^>]*)>', f'<h{i}\\1>', cleaned)
        
        # Handle Quill's list formatting with minimal spacing to match PDF
        cleaned = re.sub(r'<ol>', r'<ol style="padding-left: 20px; margin: 2px 0;">', cleaned)
        cleaned = re.sub(r'<ul>', r'<ul style="padding-left: 20px; margin: 2px 0;">', cleaned)
        
        # Handle blockquotes with minimal spacing
        cleaned = re.sub(r'<blockquote>', r'<blockquote style="border-left: 4px solid #ccc; margin: 2px 0; padding-left: 16px; color: #666;">', cleaned)
        
        # Handle code blocks with styling
        cleaned = re.sub(r'<pre([^>]*)>', r'<pre style="background: #f4f4f4; padding: 10px; border-radius: 4px; font-family: monospace;"\1>', cleaned)
        
        # Handle text alignment
        cleaned = re.sub(r'class="ql-align-center"', r'style="text-align: center;"', cleaned)
        cleaned = re.sub(r'class="ql-align-right"', r'style="text-align: right;"', cleaned)
        cleaned = re.sub(r'class="ql-align-justify"', r'style="text-align: justify;"', cleaned)
        
        return cleaned
    
    # If not HTML, check if it looks like markdown and convert it
    if not ('<p>' in html_content or '<div>' in html_content or '<br>' in html_content):
        # Convert markdown to HTML first
        html_content = convert_markdown_to_html(html_content)
    
    # Convert plain text to HTML as final fallback
    if not ('<p>' in html_content or '<div>' in html_content):
        cleaned = html_content.replace('\n', '<br>\n')
        cleaned = f"<div>{cleaned}</div>"
        return cleaned
    
    return html_content

def html_to_plain_text(html_content):
    """Convert HTML content to plain text for PDF generation with proper formatting preservation"""
    if not html_content:
        return ""
    
    # Use BeautifulSoup for better HTML parsing if available, otherwise use regex
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Handle headers first (before other formatting)
        for i in range(1, 7):
            for tag in soup.find_all(f'h{i}'):
                prefix = '#' * i + ' '
                tag.string = f"{prefix}{tag.get_text()}\n\n"
                tag.unwrap()
        
        # Handle blockquotes
        for tag in soup.find_all('blockquote'):
            lines = tag.get_text().split('\n')
            quoted_lines = [f"> {line}" if line.strip() else ">" for line in lines]
            tag.string = '\n'.join(quoted_lines) + '\n\n'
            tag.unwrap()
        
        # Handle code blocks
        for tag in soup.find_all('pre'):
            tag.string = f"```\n{tag.get_text()}\n```\n\n"
            tag.unwrap()
        
        # Handle inline code
        for tag in soup.find_all('code'):
            tag.string = f"`{tag.get_text()}`"
            tag.unwrap()
        
        # Handle strikethrough
        for tag in soup.find_all(['s', 'strike']):
            tag.string = f"~~{tag.get_text()}~~"
            tag.unwrap()
        
        # Handle underline (convert to markdown emphasis)
        for tag in soup.find_all('u'):
            tag.string = f"__{tag.get_text()}__"
            tag.unwrap()
        
        # Handle subscript and superscript
        for tag in soup.find_all('sub'):
            tag.string = f"_{tag.get_text()}_"
            tag.unwrap()
        
        for tag in soup.find_all('sup'):
            tag.string = f"^{tag.get_text()}^"
            tag.unwrap()
        
        # Handle lists first (before unwrapping other tags)
        for ol in soup.find_all('ol'):
            items = ol.find_all('li')
            for i, item in enumerate(items, 1):
                item.string = f"{i}. {item.get_text()}\n"
            ol.unwrap()
        
        for ul in soup.find_all('ul'):
            items = ul.find_all('li')
            for item in items:
                item.string = f"• {item.get_text()}\n"
            ul.unwrap()
        
        # Handle remaining list items that weren't in ol/ul
        for tag in soup.find_all('li'):
            tag.string = f"• {tag.get_text()}\n"
            tag.unwrap()
        
        # Handle bold/strong tags
        for tag in soup.find_all(['strong', 'b']):
            tag.string = f"**{tag.get_text()}**"
            tag.unwrap()
        
        # Handle italic/emphasis tags  
        for tag in soup.find_all(['em', 'i']):
            tag.string = f"*{tag.get_text()}*"
            tag.unwrap()
        
        # Handle paragraphs
        for tag in soup.find_all('p'):
            text_content = tag.get_text()
            if text_content.strip():  # Only add breaks for non-empty paragraphs
                tag.string = f"{text_content}\n\n"
            else:
                tag.string = ""
            tag.unwrap()
        
        # Handle line breaks
        for tag in soup.find_all('br'):
            tag.replace_with('\n')
        
        # Handle divs
        for tag in soup.find_all('div'):
            text_content = tag.get_text()
            if text_content.strip():
                tag.string = f"{text_content}\n"
            else:
                tag.string = ""
            tag.unwrap()
        
        text = soup.get_text()
        
    except ImportError:
        # Fallback to regex if BeautifulSoup not available
        text = html_content
        import re
        
        # Handle headers first
        for i in range(1, 7):
            prefix = '#' * i + ' '
            text = re.sub(f'<h{i}[^>]*>', prefix, text)
            text = re.sub(f'</h{i}>', '\n\n', text)
        
        # Handle blockquotes
        blockquote_pattern = r'<blockquote[^>]*>(.*?)</blockquote>'
        for match in re.finditer(blockquote_pattern, text, re.DOTALL):
            quote_content = match.group(1)
            quote_content = re.sub(r'<[^>]+>', '', quote_content)
            lines = quote_content.split('\n')
            quoted_lines = [f"> {line}" if line.strip() else ">" for line in lines]
            text = text.replace(match.group(0), '\n'.join(quoted_lines) + '\n\n')
        
        # Handle code blocks
        text = re.sub(r'<pre[^>]*>', '```\n', text)
        text = re.sub(r'</pre>', '\n```\n\n', text)
        
        # Handle inline code
        text = re.sub(r'<code[^>]*>', '`', text)
        text = re.sub(r'</code>', '`', text)
        
        # Handle strikethrough
        text = re.sub(r'<(s|strike)[^>]*>', '~~', text)
        text = re.sub(r'</(s|strike)>', '~~', text)
        
        # Handle underline
        text = re.sub(r'<u[^>]*>', '__', text)
        text = re.sub(r'</u>', '__', text)
        
        # Handle subscript and superscript
        text = re.sub(r'<sub[^>]*>', '_', text)
        text = re.sub(r'</sub>', '_', text)
        text = re.sub(r'<sup[^>]*>', '^', text)
        text = re.sub(r'</sup>', '^', text)
        
        # Handle ordered lists first
        ol_pattern = r'<ol[^>]*>(.*?)</ol>'
        for match in re.finditer(ol_pattern, text, re.DOTALL):
            ol_content = match.group(1)
            li_items = re.findall(r'<li[^>]*>(.*?)</li>', ol_content, re.DOTALL)
            numbered_items = []
            for i, item in enumerate(li_items, 1):
                clean_item = re.sub(r'<[^>]+>', '', item).strip()
                numbered_items.append(f"{i}. {clean_item}")
            text = text.replace(match.group(0), '\n'.join(numbered_items) + '\n')
        
        # Handle unordered lists
        ul_pattern = r'<ul[^>]*>(.*?)</ul>'
        for match in re.finditer(ul_pattern, text, re.DOTALL):
            ul_content = match.group(1)
            li_items = re.findall(r'<li[^>]*>(.*?)</li>', ul_content, re.DOTALL)
            bullet_items = []
            for item in li_items:
                clean_item = re.sub(r'<[^>]+>', '', item).strip()
                bullet_items.append(f"• {clean_item}")
            text = text.replace(match.group(0), '\n'.join(bullet_items) + '\n')
        
        # Handle remaining list items
        text = re.sub(r'<li[^>]*>', '• ', text)
        text = re.sub(r'</li>', '\n', text)
        
        # Convert paragraph tags to double line breaks
        text = re.sub(r'<p[^>]*>', '', text)
        text = re.sub(r'</p>', '\n\n', text)
        
        # Convert break tags to single line breaks
        text = re.sub(r'<br[^>]*/?>', '\n', text)
        
        # Convert div tags to line breaks
        text = re.sub(r'<div[^>]*>', '', text)
        text = re.sub(r'</div>', '\n', text)
        
        # Handle bold and strong tags - preserve formatting
        text = re.sub(r'<(strong|b)[^>]*>', '**', text)
        text = re.sub(r'</(strong|b)>', '**', text)
        
        # Handle italic and emphasis tags - preserve formatting
        text = re.sub(r'<(em|i)[^>]*>', '*', text)
        text = re.sub(r'</(em|i)>', '*', text)
        
        # Remove any remaining HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # Unescape HTML entities
        text = unescape(text)
    
    # Clean up multiple consecutive line breaks (more than 2)
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # Clean up spaces but preserve intentional formatting
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        # Strip trailing spaces but preserve leading spaces for indentation
        cleaned_line = line.rstrip()
        cleaned_lines.append(cleaned_line)
    
    text = '\n'.join(cleaned_lines)
    
    # Remove leading/trailing whitespace from the entire text
    return text.strip()

def show_available_variables():
    """Get available variables from uploaded Excel data and system variables"""
    excel_vars = []
    if st.session_state.excel_data is not None and not st.session_state.excel_data.empty:
        excel_vars = list(st.session_state.excel_data.columns)
    
    # Add system variables
    system_vars = ["Current Date", "Current Year", "Current Month", "Today"]
    
    return excel_vars + system_vars

def create_variable_interface(columns, key_prefix):
    """Create an easy-to-copy variable interface"""
    if not columns:
        return
    
    st.markdown("**📋 Available Variables**")
    
    # Create tabs for different variable types
    excel_vars = [col for col in columns if col not in ["Current Date", "Current Year", "Current Month", "Today"]]
    system_vars = [col for col in columns if col in ["Current Date", "Current Year", "Current Month", "Today"]]
    
    # Show variables in columns
    if excel_vars:
        st.markdown("**Excel Data Variables:**")
        cols = st.columns(min(3, len(excel_vars)))
        for i, var in enumerate(excel_vars):
            with cols[i % len(cols)]:
                st.code(f"{{{var}}}")
    
    if system_vars:
        st.markdown("**System Variables:**")
        cols = st.columns(min(3, len(system_vars)))
        for i, var in enumerate(system_vars):
            with cols[i % len(cols)]:
                st.code(f"{{{var}}}")

def create_storage_directory_selector():
    """Create interface for user to select storage directory"""
    st.markdown("### 📁 Storage Directory Configuration")
    
    current_path = get_user_configured_path()
    if current_path:
        current_display = str(current_path)
    else:
        current_display = "Using default (auto-detected desktop/documents)"
    
    st.info(f"**Current Storage Location:** `{current_display}`")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # Text input for manual path entry
        custom_path = st.text_input(
            "Custom Storage Directory (optional):",
            placeholder="Enter full path to directory (e.g., C:\\MyFolder or /home/user/MyFolder)",
            help="Leave empty to use auto-detected desktop/documents folder"
        )
    
    with col2:
        st.write("")  # Add some spacing to align with text input
        if st.button("📁 Browse", help="Browse for directory"):
            st.info("💡 **Tip:** Copy and paste the full path of your desired folder into the text box above.")
    
    # Validate and set custom path
    if custom_path and custom_path.strip():
        path_obj = Path(custom_path.strip())
        
        if path_obj.exists():
            if path_obj.is_dir():
                # Test if we can write to this directory
                try:
                    test_file = path_obj / "nsna_test_write.tmp"
                    test_file.touch()
                    test_file.unlink()
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("✅ Use This Directory", type="primary"):
                            if save_user_storage_path(path_obj):
                                st.success(f"✅ Storage directory updated to: `{path_obj}`")
                                st.info("🔄 Please refresh the page to apply changes.")
                                time.sleep(1)
                                st.rerun()
                    
                    with col2:
                        if st.button("🔄 Reset to Default"):
                            if reset_storage_path():
                                st.success("✅ Reset to default storage location")
                                st.info("🔄 Please refresh the page to apply changes.")
                                time.sleep(1)
                                st.rerun()
                    
                    st.success(f"✅ Directory is valid and writable: `{path_obj}`")
                except Exception as e:
                    st.error(f"❌ Cannot write to this directory: {e}")
            else:
                st.error("❌ Path exists but is not a directory")
        else:
            st.error("❌ Directory does not exist")
    
    # Reset button for current custom path
    if current_path:
        if st.button("🔄 Reset to Default Location"):
            if reset_storage_path():
                st.success("✅ Reset to default storage location")
                st.info("🔄 Please refresh the page to apply changes.")
                time.sleep(1)
                st.rerun()

def reset_storage_path():
    """Reset storage path to default"""
    try:
        import json
        settings_file = Path.home() / ".nsna_mail_merge_settings.json"
        
        if settings_file.exists():
            # Load existing settings
            with open(settings_file, 'r') as f:
                settings = json.load(f)
            
            # Remove storage path
            if 'storage_path' in settings:
                del settings['storage_path']
            
            # Save updated settings
            with open(settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        
        # Clear session state
        if 'custom_storage_path' in st.session_state:
            del st.session_state.custom_storage_path
        
        return True
    except Exception as e:
        st.error(f"Failed to reset storage path: {e}")
        return False

# Missing helper functions that are referenced but not defined
def show_send_confirmation_dialog(df, template, mode):
    """Show confirmation dialog before sending emails"""
    st.warning(f"⚠️ You are about to send {len(df)} emails in {mode}")
    
    # Show a sample of what will be sent
    if len(df) > 0:
        first_row = df.iloc[0]
        st.markdown("**Preview of first email:**")
        st.write(f"To: {first_row.get('First Name', 'N/A')} {first_row.get('Last Name', 'N/A')} ({first_row.get('Email', 'N/A')})")
        st.write(f"Subject: {template.get('subject', 'No Subject')}")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 Confirm and Send", type="primary", key="confirm_send_btn"):
            return True
    with col2:
        if st.button("❌ Cancel", key="cancel_send_btn"):
            st.session_state.show_send_confirmation = False
            st.rerun()
    
    return False

def replace_system_variables(content):
    """Replace system variables with current values"""
    if not content:
        return content
        
    from datetime import datetime
    now = datetime.now()
    
    replacements = {
        "{Current Date}": now.strftime("%Y-%m-%d"),
        "{Current Year}": now.strftime("%Y"),
        "{Current Month}": now.strftime("%B"),
        "{Today}": now.strftime("%Y-%m-%d")
    }
    
    for var, value in replacements.items():
        content = content.replace(var, value)
    
    return content

def update_current_templates():
    """Update current templates with the latest content from editors"""
    # Get email template content from the separate session state key to avoid widget conflicts
    email_content = st.session_state.get('current_email_content', '')
    
    # Ensure email content is not None
    if email_content is None:
        email_content = ''
    
    if email_content and hasattr(st.session_state, 'email_settings'):
        email_settings = st.session_state.email_settings or {}
        st.session_state.current_email_template = {
            'content': email_content,
            'from_email': email_settings.get('from_email', '') or '',
            'subject': email_settings.get('subject', '') or 'NSNA Donation Receipt'
        }
    
    # Get PDF template content
    pdf_content = st.session_state.get('pdf_content_text', '')
    
    # Ensure PDF content is not None
    if pdf_content is None:
        pdf_content = ''
        
    if pdf_content:
        st.session_state.current_pdf_template = {
            'content': pdf_content,
            'template_path': 'NSNA Atlanta Letterhead Updated.pdf'
        }

def send_email_debug_wrapper(from_email, to_email, subject, html_content, attachment_path, smtp_settings, is_test=False):
    """Local debug wrapper for email sending - calls the actual mail sender function"""
    try:
        # Add basic debug info to the session debug log
        st.session_state.debug_log.append(f"  📤 **ATTEMPTING EMAIL SEND:**")
        st.session_state.debug_log.append(f"    - Content: {len(html_content)} chars")
        st.session_state.debug_log.append(f"    - Subject: {subject[:50]}{'...' if len(subject) > 50 else ''}")
        
        if attachment_path:
            if os.path.exists(attachment_path):
                file_size = os.path.getsize(attachment_path) / 1024  # Size in KB
                st.session_state.debug_log.append(f"    - Attachment: {file_size:.1f} KB")
            else:
                st.session_state.debug_log.append(f"    - ⚠️ Attachment missing: {attachment_path}")
        
        # Import and use the actual mail sender function from utils
        from utils.mail_sender import send_email_with_diagnostics as actual_send_email
        
        # Call the actual send email function
        success = actual_send_email(
            from_email=from_email,
            to_email=to_email,
            subject=subject,
            html_content=html_content,
            attachment_path=attachment_path,
            smtp_settings=smtp_settings,
            is_test=is_test
        )
        
        return success
        
    except Exception as e:
        st.session_state.debug_log.append(f"    - ❌ **SEND FAILED:** {type(e).__name__}: {str(e)}")
        return False

def get_letterhead_path():
    """Get the letterhead path for cloud deployment"""
    # For cloud deployment, check if letterhead exists in expected locations
    possible_paths = [
        'NSNA Atlanta Letterhead Updated.pdf',
        'src/templates/letterhead_template.pdf',
        'letterhead_template.pdf'
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    return None

def load_saved_templates():
    """Load saved email and PDF templates"""
    try:
        persistence = st.session_state.template_persistence
        
        # Load email template
        email_template = persistence.load_email_template()
        if email_template:
            st.session_state.saved_email_template = email_template
        
        # Load PDF template
        pdf_template = persistence.load_pdf_template()
        if pdf_template:
            st.session_state.saved_pdf_template = pdf_template
            
        # Load user settings
        user_settings = persistence.load_user_settings()
        if user_settings:
            st.session_state.email_settings.update(user_settings)
            
    except Exception as e:
        # Don't show error for missing templates - they might not exist yet
        pass

# Call main function at the end
if __name__ == "__main__":
    main()