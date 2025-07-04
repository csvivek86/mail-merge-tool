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
    
    from utils.rich_text_editor import (
        create_rich_text_editor, 
        create_template_formatting_options,
        create_email_template_presets,
        create_variable_helper,
        apply_formatting_to_html
    )
    from config.settings import EMAIL_SETTINGS, DEFAULT_EMAIL, USE_OAUTH, DYNAMIC_USER_OAUTH
    from config.email_template_settings import EmailTemplateSettings
    from config.template_settings import TemplateSettings
except ImportError as e:
    st.error(f"Import error: {e}")
    st.info("Please ensure all required modules are in the src/ directory")    # Configure Streamlit page
    st.set_page_config(
        page_title="NSNA Mail Merge Tool",
        page_icon="📧",
        layout="wide",
        initial_sidebar_state="collapsed"  # Collapse sidebar since we're using tabs
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
    """Main Streamlit application - Simplified 2-tab interface"""
    
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
    
    # Simplified 2-tab navigation
    tab1, tab2 = st.tabs([
        "📊 Email Content", 
        "� PDF Content"
    ])
    
    with tab1:
        email_content_tab()
    
    with tab2:
        pdf_content_tab()

def email_content_tab():
    """Email Content Tab: Upload data, email settings, and email template editor"""
    
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
        default_from = ""
        if saved_user_settings:
            default_from = saved_user_settings.get('from_email', '')
            
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
    
    # Email Template Editor
    st.header("✏️ Email Template")
    
    if data_uploaded and email_settings_valid:
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
            key="email_template_quill",
            height=300
        )
        
        # Variable insertion helper
        create_variable_helper("email")
        
        if email_content and email_content.strip():
            st.session_state.current_email_template = {
                'content': email_content,
                'from_email': from_email,
                'subject': subject
            }
            
            # Save template button
            if st.button("💾 Save Template"):
                save_email_template()
    else:
        st.info("Complete settings above first")


def pdf_content_tab():
    """PDF Content Tab: Send buttons at top, PDF template editor, and sample download"""
    
    # Check if setup is complete (simplified check)
    if (st.session_state.excel_data is None or 
        not hasattr(st.session_state, 'email_settings') or
        not hasattr(st.session_state, 'current_email_template')):
        st.info("Complete Email Content tab first")
        return
    
    # Send Buttons at the top
    st.header("📧 Send Emails")
    
    df = st.session_state.excel_data
    template = st.session_state.current_email_template
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🧪 Test Email", type="secondary"):
            execute_mail_merge(df.head(1), template, "Test Mode", 1)
    
    with col2:
        if st.button("🚀 Send All Emails", type="primary"):
            # Add confirmation
            if st.session_state.get('confirm_send_all', False):
                execute_mail_merge(df, template, "Live Mode", 3)
                st.session_state.confirm_send_all = False
            else:
                st.session_state.confirm_send_all = True
                st.warning("Click again to confirm")
                st.rerun()
    
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
            if st.button("🗑️ Clear"):
                st.session_state.sent_emails = []
                st.rerun()
    
    # PDF Template Editor
    st.header("📄 PDF Template")
    
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
        
        # PDF content editor
        pdf_content = st_quill(
            value=default_pdf_content,
            placeholder="Create your PDF receipt content here...",
            key="pdf_content_quill",
            height=250
        )
        
        # Variable insertion helper
        create_variable_helper("pdf")
        
        if pdf_content and pdf_content.strip():
            st.session_state.current_pdf_template = {
                'content': pdf_content,
                'template_path': str(letterhead_path)
            }
            
            # Save and sample buttons at bottom
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 Save Template"):
                    save_pdf_template()
            
            with col2:
                if st.button("📄 Download Sample PDF"):
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

def create_variable_helper(template_type):
    """Create variable insertion helper"""
    with st.expander("🎯 Insert Variables", expanded=False):
        st.info("💡 Tip: Click any button below to add the variable to your clipboard area, then copy and paste into the editor above.")
        
        # Create a text area to show the inserted variables
        if f'inserted_{template_type}_variables' not in st.session_state:
            st.session_state[f'inserted_{template_type}_variables'] = ""
        
        available_vars = list(st.session_state.excel_data.columns) if st.session_state.excel_data is not None else []
        if available_vars:
            st.write("**From your Excel file:**")
            cols = st.columns(4)
            for i, var in enumerate(available_vars):
                with cols[i % 4]:
                    if st.button(f"Copy {{{var}}}", key=f"{template_type}_var_{i}", help=f"Copy {{{var}}} to clipboard area"):
                        st.session_state[f'inserted_{template_type}_variables'] += f" {{{var}}}"
                        st.success(f"Added {{{var}}} to clipboard area below!")
        
        st.write("**Additional variables:**")
        cols = st.columns(4)
        additional_vars = ["Date", "Year", "Amount"]
        for i, var in enumerate(additional_vars):
            with cols[i % 4]:
                if st.button(f"Copy {{{var}}}", key=f"{template_type}_additional_{i}", help=f"Copy {{{var}}} to clipboard area"):
                    st.session_state[f'inserted_{template_type}_variables'] += f" {{{var}}}"
                    st.success(f"Added {{{var}}} to clipboard area below!")
        
        # Show the accumulated variables for easy copying
        if st.session_state[f'inserted_{template_type}_variables']:
            st.write("**� Variables and Text to Copy:**")
            st.text_area(
                "Copy this content to your template editor above:",
                value=st.session_state[f'inserted_{template_type}_variables'],
                height=100,
                key=f"{template_type}_clipboard",
                help="Select all text (Ctrl+A) and copy (Ctrl+C), then paste into the rich text editor above"
            )
            if st.button("�️ Clear Clipboard", key=f"clear_{template_type}_clipboard"):
                st.session_state[f'inserted_{template_type}_variables'] = ""
                st.success("Clipboard cleared!")
                st.rerun()

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
                'content': pdf_content,
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
    """Execute the mail merge process - simplified version"""
    
    # Show progress at the current location
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    successful_sends = 0
    failed_sends = 0
    
    try:
        # Use email config from current template
        from_email = template['from_email']
        subject_template = template['subject']
        smtp_settings = st.session_state.email_settings_config.copy()
        smtp_settings.update({'smtp_user_name': from_email})
        
        total_emails = len(df)
        
        for index, row in df.iterrows():
            try:
                # Update progress
                progress = (index + 1) / total_emails
                progress_bar.progress(progress)
                status_text.text(f"Processing {index + 1}/{total_emails}: {row['First Name']} {row['Last Name']}")
                
                # Generate personalized content
                row_dict = row.to_dict()
                personalized_content = template['content']
                personalized_subject = subject_template
                
                # Replace variables in content and subject
                for col, value in row_dict.items():
                    personalized_content = personalized_content.replace(f"{{{col}}}", str(value))
                    personalized_subject = personalized_subject.replace(f"{{{col}}}", str(value))
                
                # Generate PDF receipt if template exists
                pdf_path = None
                if hasattr(st.session_state, 'current_pdf_template'):
                    try:
                        pdf_generator = PDFGenerator()
                        pdf_data = row_dict.copy()
                        pdf_data.update(st.session_state.current_pdf_template)
                        pdf_path = pdf_generator.generate_receipt(pdf_data)
                    except Exception as pdf_error:
                        st.warning(f"PDF generation failed for {row['First Name']} {row['Last Name']}: {pdf_error}")
                
                # Determine recipient email
                recipient_email = row['Email'] if send_mode == "Live Mode (Send to Recipients)" else from_email
                
                # Send email
                success = send_email_with_diagnostics(
                    from_email=from_email,
                    to_email=recipient_email,
                    subject=personalized_subject,
                    html_content=personalized_content,
                    attachment_path=pdf_path,
                    smtp_settings=smtp_settings,
                    is_test=(send_mode == "Test Mode (Send to Self)")
                )
                
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
        time.sleep(2)
        st.rerun()
    
    except Exception as e:
        st.error(f"Mail merge failed: {e}")
        status_text.text("❌ Mail merge failed!")
        progress_bar.empty()

   
# Main application entry point
if __name__ == "__main__":
    main()