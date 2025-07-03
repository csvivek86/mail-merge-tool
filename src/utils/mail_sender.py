from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import logging
from pathlib import Path
from config.email_template_settings import EmailTemplateSettings
import json


def send_email(from_email, to_email, subject, html_content, attachment_path=None, smtp_settings=None, is_test=False):
    """Send email with attachments"""
    try:
        # Validate SMTP settings
        if not smtp_settings or not isinstance(smtp_settings, dict):
            raise ValueError("SMTP settings must be provided")
            
        required_settings = ['smtp_server', 'smtp_port', 'smtp_user_name', 'smtp_password']
        missing_settings = [s for s in required_settings if s not in smtp_settings]
        if missing_settings:
            raise ValueError(f"Missing required SMTP settings: {', '.join(missing_settings)}")
        
        msg = MIMEMultipart()
        actual_recipient = from_email if is_test else to_email
        
        # Set email headers with validated SMTP settings
        msg['From'] = f"Nagarathar Sangam of North America <{from_email}>"  # Show treasurer's email as sender
        # msg['From'] = smtp_settings['smtp_user_name']
        msg['To'] = actual_recipient
        msg['Subject'] = f"{'[TEST] ' if is_test else ''}{subject}"
        msg['Reply-To'] = from_email

        # Attach HTML body
        msg.attach(MIMEText(html_content, 'html'))
        
        # Attach PDF if provided
        if attachment_path and Path(attachment_path).exists():
            try:
                with open(str(attachment_path), 'rb') as f:
                    pdf_data = f.read()
                    
                # Create MIMEApplication for PDF
                part = MIMEApplication(pdf_data, _subtype='pdf')
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=Path(attachment_path).name
                )
                msg.attach(part)
                logging.info(f"PDF attachment added: {Path(attachment_path).name}")
            except Exception as e:
                logging.error(f"Failed to attach PDF {attachment_path}: {str(e)}")
                # Continue sending email without attachment
        elif attachment_path:
            logging.warning(f"Attachment path provided but file not found: {attachment_path}")
        
        # Send email
        with smtplib.SMTP(smtp_settings['smtp_server'], smtp_settings['smtp_port']) as server:
            server.starttls()
            server.login(smtp_settings['smtp_user_name'], smtp_settings['smtp_password'])
            server.send_message(msg)
            
        logging.info(f"{'Test email' if is_test else 'Email'} sent successfully to {actual_recipient}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send email: {str(e)}")
        return False

def load_email_template():
    """Load email template settings"""
    try:
        settings_path = Path.cwd() / 'config' / 'email_template_settings.json'
        if settings_path.exists():
            with open(settings_path) as f:
                return EmailTemplateSettings(**json.load(f))
        return EmailTemplateSettings()
    except Exception as e:
        logging.error(f"Failed to load email template: {str(e)}")
        return EmailTemplateSettings()

def format_email_body(template: EmailTemplateSettings, donor_info: dict) -> str:
    """Format email body with donor information"""
    body = [
        f"{template.greeting} {donor_info['First Name']} {donor_info['Last Name']},",
        "",
        *template.body,
        "",
        *template.signature
    ]
    return "<br>".join(body)

def send_email_with_multiple_attachments(from_email, to_email, subject, html_content, attachment_paths=None, smtp_settings=None, is_test=False):
    """Send email with multiple attachments"""
    try:
        # Validate SMTP settings
        if not smtp_settings or not isinstance(smtp_settings, dict):
            raise ValueError("SMTP settings must be provided")
            
        required_settings = ['smtp_server', 'smtp_port', 'smtp_user_name', 'smtp_password']
        missing_settings = [s for s in required_settings if s not in smtp_settings]
        if missing_settings:
            raise ValueError(f"Missing required SMTP settings: {', '.join(missing_settings)}")
        
        msg = MIMEMultipart()
        actual_recipient = from_email if is_test else to_email
        
        # Set email headers
        msg['From'] = f"Nagarathar Sangam of North America <{from_email}>"
        msg['To'] = actual_recipient
        msg['Subject'] = f"{'[TEST] ' if is_test else ''}{subject}"
        msg['Reply-To'] = from_email

        # Attach HTML body
        msg.attach(MIMEText(html_content, 'html'))
        
        # Attach multiple files if provided
        if attachment_paths:
            if isinstance(attachment_paths, (str, Path)):
                attachment_paths = [attachment_paths]  # Convert single path to list
                
            for attachment_path in attachment_paths:
                if attachment_path and Path(attachment_path).exists():
                    try:
                        with open(str(attachment_path), 'rb') as f:
                            file_data = f.read()
                        
                        # Determine file type
                        file_path = Path(attachment_path)
                        if file_path.suffix.lower() == '.pdf':
                            part = MIMEApplication(file_data, _subtype='pdf')
                        else:
                            part = MIMEApplication(file_data, _subtype='octet-stream')
                            
                        part.add_header(
                            'Content-Disposition',
                            'attachment',
                            filename=file_path.name
                        )
                        msg.attach(part)
                        logging.info(f"Attachment added: {file_path.name}")
                    except Exception as e:
                        logging.error(f"Failed to attach file {attachment_path}: {str(e)}")
                        continue
                elif attachment_path:
                    logging.warning(f"Attachment path provided but file not found: {attachment_path}")
        
        # Send email
        with smtplib.SMTP(smtp_settings['smtp_server'], smtp_settings['smtp_port']) as server:
            server.starttls()
            server.login(smtp_settings['smtp_user_name'], smtp_settings['smtp_password'])
            server.send_message(msg)
            
        logging.info(f"{'Test email' if is_test else 'Email'} sent successfully to {actual_recipient}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send email with multiple attachments: {str(e)}")
        return False