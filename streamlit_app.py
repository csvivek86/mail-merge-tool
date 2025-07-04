"""
NSNA Mail Merge Tool - Streamlit Web Application
Convert PyQt6 desktop app to a modern web interface
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
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
from streamlit_quill import st_quill
import re
from html import unescape

# Add src directory to Python path
src_path = Path(__file__).parent / 'src'
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

# Import existing utility functions
try:
    from utils.excel_reader import read_excel, get_contacts_as_list
    from utils.mail_sender import send_email_with_diagnostics
    from utils.pdf_generator import PDFGenerator
    from utils.oauth_manager import get_user_credentials
    
    # Use cloud-compatible persistence for Streamlit Cloud deployment
    try:
        from utils.cloud_template_persistence import (
            CloudTemplatePersistence, 
            get_letterhead_path, 
            get_cloud_email_settings,
            get_oauth_credentials
        )
        # Use cloud persistence
        TemplatePersistence = CloudTemplatePersistence
        CLOUD_DEPLOYMENT = True
    except ImportError:
        # Fallback to local persistence for development
        from utils.template_persistence import TemplatePersistence
        CLOUD_DEPLOYMENT = False
    
    try:
        from utils.rich_text_editor import (
            create_rich_text_editor, 
            create_template_formatting_options,
            create_email_template_presets,
            apply_formatting_to_html
        )
    except ImportError:
        # Rich text editor functions not available
        def create_rich_text_editor(*args, **kwargs):
            return ""
        def create_template_formatting_options(*args, **kwargs):
            return ""
        def create_email_template_presets(*args, **kwargs):
            return {}
        def apply_formatting_to_html(*args, **kwargs):
            return ""
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
if 'template_persistence' not in st.session_state:
    st.session_state.template_persistence = TemplatePersistence()
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
        uploaded_file = st.file_uploader(
            "Choose Excel file",
            type=['xlsx', 'xls']
        )
        
        data_uploaded = False
        
        if uploaded_file is not None:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                df = read_excel(tmp_file_path)
                
                if df is not None and not df.empty:
                    st.session_state.excel_data = df
                    st.success(f"✅ {len(df)} records")
                    data_uploaded = True
                
                os.unlink(tmp_file_path)
                
            except Exception as e:
                st.error(f"Error: {e}")
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
                show_send_confirmation_dialog(df, template)
        else:
            st.info("Complete setup first")
    
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
            presets = {
                "Donation Receipt": """Dear {First Name} {Last Name},

Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.

Please do not hesitate to contact me if you have any questions on the attached document.

Regards,
Nagarathar Sangam of North America Treasurer,
Email: nsnatreasurer@achi.org""",
                "Custom": ""
            }
            
            selected_preset = st.selectbox("Template:", list(presets.keys()))
            
            saved_email = getattr(st.session_state, 'saved_email_template', None)
            if saved_email and saved_email.get('content'):
                template_content = saved_email.get('content')
            else:
                template_content = presets[selected_preset]
            
            email_content = st_quill(
                value=template_content,
                placeholder="Type your email template here...",
                key="email_template_quill"
            )
            
            # Always update the current template with editor content, even if empty
            if email_content is not None:  # Check for None instead of strip()
                st.session_state.current_email_template = {
                    'content': email_content,
                    'from_email': from_email,
                    'subject': subject
                }
                
                if st.button("💾 Save Email Template", use_container_width=True):
                    save_email_template()
                
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
                        st.markdown(preview_content, unsafe_allow_html=True)
                    else:
                        # Show preview with just system variables
                        preview_content = replace_system_variables(email_content)
                        st.markdown("**Preview with system variables:**")
                        st.markdown(preview_content, unsafe_allow_html=True)
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
            saved_pdf = getattr(st.session_state, 'saved_pdf_template', None)
            if saved_pdf and saved_pdf.get('content'):
                default_pdf_content = saved_pdf.get('content')
            else:
                default_pdf_content = """<p>Dear {First Name} {Last Name},</p>

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
            
            pdf_content = st_quill(
                value=default_pdf_content,
                placeholder="Create your PDF receipt content here...",
                key="pdf_content_main"
            )
            
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
                        
                        # Convert HTML to formatted plain text for PDF preview
                        formatted_preview = html_to_plain_text(preview_content)
                        
                        st.markdown("**Preview with sample data:**")
                        st.text(formatted_preview)  # Use text to show exact PDF formatting
                    else:
                        # Show preview with just system variables
                        preview_content = replace_system_variables(pdf_content)
                        formatted_preview = html_to_plain_text(preview_content)
                        st.markdown("**Preview with system variables:**")
                        st.text(formatted_preview)
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
    
    # Template presets
    presets = {
        "Donation Receipt": """<p>Dear {First Name} {Last Name},</p>
<p>Thank you for your generous donation to the Nagarathar Sangam of North America.</p>
<p><strong>Donation Details:</strong><br>
• Amount: ${Amount}<br>
• Date: {Date}</p>
<p>Please find your official receipt attached as a PDF for your tax records.</p>
<p>Best regards,<br>
<strong>NSNA Treasurer</strong></p>""",
        
        "Custom": ""
    }
    
    selected_preset = st.selectbox("Template:", list(presets.keys()))
    
    # Load saved template if available
    saved_email = getattr(st.session_state, 'saved_email_template', None)
    if saved_email:
        template_content = saved_email.get('content', presets[selected_preset])
    else:
        template_content = presets[selected_preset]
    
    # Email template editor
    email_content = st_quill(
        value=template_content,
        placeholder="Type your email template here...",
        key="email_template_quill"
    )
    
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
                'content': html_to_plain_text(replace_system_variables(pdf_content)),  # Convert HTML to plain text and replace system vars
                'template_path': str(letterhead_path)
            }
            
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
    
    # Show progress at the current location
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    successful_sends = 0
    failed_sends = 0
    
    try:
        # Get the latest content from editors instead of using passed template
        latest_email_content = st.session_state.get('email_template_quill', template.get('content', ''))
        latest_pdf_content = st.session_state.get('pdf_content_main', '')
        
        # Use email config from current template
        from_email = template['from_email']
        subject_template = template['subject']
        
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
        

        
        total_emails = len(df)
        
        for index, row in df.iterrows():
            try:
                # Update progress
                progress = (index + 1) / total_emails
                progress_bar.progress(progress)
                status_text.text(f"Processing {index + 1}/{total_emails}: {row['First Name']} {row['Last Name']}")
                
                # Generate personalized content using current template
                row_dict = row.to_dict()
                personalized_content = current_template['content']
                personalized_subject = subject_template
                
                # First replace system variables
                personalized_content = replace_system_variables(personalized_content)
                personalized_subject = replace_system_variables(personalized_subject)
                
                # Then replace Excel data variables
                for col, value in row_dict.items():
                    personalized_content = personalized_content.replace(f"{{{col}}}", str(value))
                    personalized_subject = personalized_subject.replace(f"{{{col}}}", str(value))
                
                # Clean HTML content for email
                email_content = clean_html_content(personalized_content)
                
                # Generate PDF receipt if template exists
                pdf_path = None
                if latest_pdf_content and latest_pdf_content.strip():
                    try:
                        pdf_generator = PDFGenerator()
                        pdf_data = row_dict.copy()
                        
                        # Use latest PDF content from editor
                        pdf_content = latest_pdf_content
                        
                        # Convert HTML content to plain text for PDF
                        pdf_content = html_to_plain_text(pdf_content)
                        
                        # First replace system variables
                        pdf_content = replace_system_variables(pdf_content)
                        
                        # Then replace Excel data variables
                        for col, value in row_dict.items():
                            pdf_content = pdf_content.replace(f"{{{col}}}", str(value))
                        
                        pdf_data['content'] = pdf_content
                        
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
                    except Exception as pdf_error:
                        st.warning(f"PDF generation failed for {row['First Name']} {row['Last Name']}: {pdf_error}")
                
                # Determine recipient email
                recipient_email = row['Email'] if send_mode == "Live Mode" else from_email
                
                # Send email with error handling
                try:
                    success = send_email_with_diagnostics(
                        from_email=from_email,
                        to_email=recipient_email,
                        subject=personalized_subject,
                        html_content=email_content,
                        attachment_path=pdf_path,
                        smtp_settings=smtp_settings,
                        is_test=(send_mode == "Test Mode")
                    )
                except Exception as email_error:
                    st.error(f"❌ Email error for {row['First Name']} {row['Last Name']}: {str(email_error)}")
                    success = False
                
                if success:
                    successful_sends += 1
                    st.session_state.sent_emails.append({
                        'name': f"{row['First Name']} {row['Last Name']}",
                        'email': recipient_email,
                        'status': 'Success',
                        'timestamp': datetime.now()
                    })
                else:
                    failed_sends += 1
                    st.session_state.sent_emails.append({
                        'name': f"{row['First Name']} {row['Last Name']}",
                        'email': recipient_email,
                        'status': 'Failed',
                        'timestamp': datetime.now()
                    })
                
                # Add delay between emails
                time.sleep(delay_between_emails)
                
            except Exception as e:
                failed_sends += 1
                st.error(f"Failed to process {row['First Name']} {row['Last Name']}: {e}")
        
        # Final results
        progress_bar.progress(1.0)
        status_text.text("✅ Mail merge complete!")
        
        # Show simple results summary
        if successful_sends > 0:
            st.success(f"🎉 Successfully sent {successful_sends} emails!")
        if failed_sends > 0:
            st.error(f"❌ {failed_sends} emails failed")
            
        # Auto-refresh to show results
        time.sleep(0.5)
        st.rerun()
    
    except Exception as e:
        st.error(f"Mail merge failed: {e}")
        status_text.text("❌ Mail merge failed!")
        progress_bar.empty()

def clean_html_content(html_content):
    """Clean HTML content for email sending"""
    if not html_content:
        return ""
    
    # If it's already proper HTML (from rich editor), use it as-is
    if '<p>' in html_content or '<div>' in html_content or '<br>' in html_content:
        return html_content
    
    # Convert plain text to HTML
    cleaned = html_content
    
    # Convert line breaks to proper HTML
    cleaned = cleaned.replace('\n', '<br>\n')
    
    # Wrap in proper HTML structure
    cleaned = f"<div>{cleaned}</div>"
    
    return cleaned

def html_to_plain_text(html_content):
    """Convert HTML content to plain text for PDF generation with proper formatting"""
    if not html_content:
        return ""
    
    # Replace HTML tags with appropriate text formatting
    text = html_content
    
    # Convert paragraph tags to double line breaks
    text = re.sub(r'<p[^>]*>', '', text)
    text = re.sub(r'</p>', '\n\n', text)
    
    # Convert break tags to single line breaks
    text = re.sub(r'<br[^>]*/?>', '\n', text)
    
    # Convert div tags to line breaks
    text = re.sub(r'<div[^>]*>', '', text)
    text = re.sub(r'</div>', '\n', text)
    
    # Handle bold and strong tags
    text = re.sub(r'<(strong|b)[^>]*>', '**', text)
    text = re.sub(r'</(strong|b)>', '**', text)
    
    # Handle italic and emphasis tags
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
        return ""
    
    st.markdown("**📋 Available Variables**")
    
    # Create tabs for different variable types
    excel_vars = [col for col in columns if col not in ["Current Date", "Current Year", "Current Month", "Today"]]
    system_vars = [col for col in columns if col in ["Current Date", "Current Year", "Current Month", "Today"]]
    
    if excel_vars and system_vars:
        tab1, tab2 = st.tabs(["📊 Excel Data", "⚙️ System"])
        
        with tab1:
            create_variable_display(excel_vars, f"{key_prefix}_excel")
            
        with tab2:
            create_variable_display(system_vars, f"{key_prefix}_system")
    else:
        # Show all variables if we only have one type
        create_variable_display(columns, key_prefix)
    
    return ""

def create_variable_display(variables, key_prefix):
    """Create a compact variable display with easy copying"""
    if not variables:
        return
    
    st.markdown("**📋 Available Variables**")
    
    # Create a grid of variables with copy-friendly display
    cols_per_row = 2
    for i in range(0, len(variables), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, var in enumerate(variables[i:i+cols_per_row]):
            with cols[j]:
                variable_text = f"{{{var}}}"
                # Create a text input that's easy to select and copy
                st.text_input(
                    f"📋 {var}:",
                    value=variable_text,
                    key=f"{key_prefix}_{var}_{i}_{j}",
                    help="Select all (Ctrl+A) and copy (Ctrl+C)",
                    label_visibility="collapsed"
                )
                st.caption(f"📋 {var}")
    
    st.caption("💡 **Tip:** Click in any text box above, select all (Ctrl+A), then copy (Ctrl+C)")

def replace_system_variables(content):
    """Replace system variables with actual values"""
    if not content:
        return content
    
    from datetime import datetime
    now = datetime.now()
    
    replacements = {
        "{Current Date}": now.strftime("%Y-%m-%d"),
        "{Current Year}": now.strftime("%Y"),
        "{Current Month}": now.strftime("%B"),
        "{Today}": now.strftime("%B %d, %Y")
    }
    
    for var, value in replacements.items():
        content = content.replace(var, value)
    
    return content

def update_current_templates():
    """Update current templates with the latest editor content"""
    # This function ensures templates are always up-to-date before sending
    
    # Update email template if we have email settings
    if (hasattr(st.session_state, 'email_settings') and 
        st.session_state.email_settings and
        'email_template_quill' in st.session_state):
        
        email_content = st.session_state.get('email_template_quill', '')
        if email_content:
            st.session_state.current_email_template = {
                'content': email_content,
                'from_email': st.session_state.email_settings.get('from_email', ''),
                'subject': st.session_state.email_settings.get('subject', '')
            }
    
    # Update PDF template if we have content
    if 'pdf_content_main' in st.session_state:
        pdf_content = st.session_state.get('pdf_content_main', '')
        if pdf_content:
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
                st.session_state.current_pdf_template = {
                    'content': pdf_content,
                    'template_path': str(letterhead_path)
                }
    
def show_send_confirmation_dialog(df, template):
    """Show confirmation dialog before sending all emails"""
    
    # Create a container for the confirmation dialog
    with st.container():
        st.markdown("---")
        st.markdown("### ⚠️ Confirm Email Sending")
        st.warning(f"**Send emails to {len(df)} recipients?**")
        
        st.error("⚠️ **This action cannot be undone!**")
        
        # Confirmation buttons
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if st.button("❌ Cancel", use_container_width=True, key="cancel_send"):
                st.session_state.show_send_confirmation = False
                st.rerun()
        
        with col2:
            if st.button("🧪 Send Test First", use_container_width=True, type="secondary", key="test_send"):
                st.session_state.show_send_confirmation = False
                # Update templates with current editor content
                update_current_templates()
                template = st.session_state.current_email_template
                execute_mail_merge(df.head(1), template, "Test Mode", 0.1)
        
        with col3:
            if st.button("🚀 Send All", use_container_width=True, type="primary", key="confirm_send_all"):
                st.session_state.show_send_confirmation = False
                # Update templates with current editor content
                update_current_templates()
                template = st.session_state.current_email_template
                execute_mail_merge(df, template, "Live Mode", 0.5)

# Run the main application
if __name__ == "__main__":
    main()