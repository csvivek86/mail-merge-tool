from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import logging
import base64
from pathlib import Path
from config.email_template_settings import EmailTemplateSettings
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os
from utils.oauth_manager import get_user_credentials

# OAuth2 SMTP authentication mechanism
def generate_oauth2_string(username, access_token):
    """Generates the OAuth2 string to be used for SMTP authentication."""
    auth_string = f"user={username}\1auth=Bearer {access_token}\1\1"
    return base64.b64encode(auth_string.encode()).decode()

def get_oauth_access_token(user_email, client_id, client_secret, refresh_token=None, dynamic_user_oauth=False):
    """
    Get an access token for SMTP authentication.
    
    If dynamic user OAuth is enabled, it will get credentials for the specific user.
    Otherwise, it will use the provided refresh token.
    """
    try:
        # Check if we're using dynamic user OAuth
        if user_email and dynamic_user_oauth:
            logging.info(f"Using dynamic OAuth for {user_email}")
            
            # Get user-specific credentials - this will trigger browser auth only if needed
            creds = get_user_credentials(user_email, client_id, client_secret)
            if not creds:
                raise ValueError(f"Failed to get OAuth credentials for {user_email}")
            
            logging.info(f"Successfully obtained token for {user_email}")
            return creds.token
        
        # Traditional OAuth with refresh token from settings
        elif refresh_token:
            creds = Credentials.from_authorized_user_info(
                info={
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "refresh_token": refresh_token,
                    "token_uri": "https://oauth2.googleapis.com/token"
                }
            )
            
            # Refresh the token
            creds.refresh(Request())
            return creds.token
            
        else:
            raise ValueError("No refresh token provided and dynamic OAuth is disabled")
            
    except Exception as e:
        logging.error(f"Failed to get OAuth access token: {str(e)}")
        raise


def send_email(from_email, to_email, subject, html_content, attachment_path=None, smtp_settings=None, is_test=False):
    """Send email with attachments"""
    try:
        # Validate SMTP settings
        if not smtp_settings or not isinstance(smtp_settings, dict):
            raise ValueError("SMTP settings must be provided")
        
        # Check if we're using OAuth or password auth
        use_oauth = smtp_settings.get('use_oauth', False)
        dynamic_user_oauth = smtp_settings.get('dynamic_user_oauth', False)
        
        # Determine required settings based on authentication method
        if use_oauth:
            if dynamic_user_oauth:
                # For dynamic user OAuth, we need client ID and secret
                required_settings = ['smtp_server', 'smtp_port', 'smtp_user_name', 'client_id', 'client_secret']
            else:
                # For traditional OAuth, we also need a refresh token
                required_settings = ['smtp_server', 'smtp_port', 'smtp_user_name', 'client_id', 'client_secret', 'refresh_token']
        else:
            # For password authentication
            required_settings = ['smtp_server', 'smtp_port', 'smtp_user_name', 'smtp_password']
            
        # Check for missing settings
        missing_settings = [s for s in required_settings if s not in smtp_settings or not smtp_settings[s]]
        if missing_settings:
            raise ValueError(f"Missing required SMTP settings: {', '.join(missing_settings)}")
        
        msg = MIMEMultipart()
        actual_recipient = from_email if is_test else to_email
        
        # Ensure from_email is clean without any display name parts
        clean_from_email = from_email.strip()
        # Remove any display name part if exists
        if '<' in clean_from_email and '>' in clean_from_email:
            clean_from_email = clean_from_email[clean_from_email.find('<')+1:clean_from_email.find('>')]
        
        # Set email headers with validated SMTP settings
        msg['From'] = f"Nagarathar Sangam of North America <{clean_from_email}>"  # Show treasurer's email as sender
        msg['To'] = actual_recipient
        msg['Subject'] = f"{'[TEST] ' if is_test else ''}{subject}"
        msg['Reply-To'] = clean_from_email

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
        
        # Send email - use different auth methods depending on settings
        smtp_timeout = 30  # 30 second timeout
        max_retries = 3
        retry_delay = 5  # seconds
        
        for attempt in range(max_retries):
            try:
                logging.info(f"Attempting SMTP connection (attempt {attempt + 1}/{max_retries})")
                
                # Create SMTP connection with timeout
                server = smtplib.SMTP(smtp_settings['smtp_server'], smtp_settings['smtp_port'], timeout=smtp_timeout)
                server.set_debuglevel(1 if logging.getLogger().level == logging.DEBUG else 0)
                
                # Start TLS encryption
                logging.info("Starting TLS encryption...")
                server.starttls()
                break  # Connection successful
                
            except (smtplib.SMTPConnectError, OSError, TimeoutError) as conn_error:
                logging.warning(f"SMTP connection attempt {attempt + 1} failed: {str(conn_error)}")
                
                if attempt < max_retries - 1:  # Not the last attempt
                    logging.info(f"Retrying in {retry_delay} seconds...")
                    import time
                    time.sleep(retry_delay)
                else:
                    # Last attempt failed
                    error_msg = f"Failed to connect to {smtp_settings['smtp_server']}:{smtp_settings['smtp_port']} after {max_retries} attempts.\n\n"
                    
                    if "10060" in str(conn_error) or "timeout" in str(conn_error).lower():
                        error_msg += "This appears to be a network timeout issue. Possible solutions:\n"
                        error_msg += "• Check your internet connection\n"
                        error_msg += "• Verify firewall settings (allow port 587)\n"
                        error_msg += "• Contact your ISP if they block SMTP\n"
                        error_msg += "• Try using a VPN if on a corporate network\n"
                        error_msg += "• Check if antivirus is blocking the connection"
                    
                    raise ConnectionError(error_msg) from conn_error
        
        try:
            # Determine which email address to use for SMTP authentication
            auth_email = from_email if from_email else smtp_settings['smtp_user_name']
            
            # Ensure auth_email is clean without any display name parts
            if '<' in auth_email and '>' in auth_email:
                auth_email = auth_email[auth_email.find('<')+1:auth_email.find('>')]
            auth_email = auth_email.strip()
            
            if use_oauth:
                # Use OAuth 2.0 authentication
                try:
                    dynamic_oauth = smtp_settings.get('dynamic_user_oauth', False)
                    logging.info(f"Using {'dynamic user ' if dynamic_oauth else ''}OAuth 2.0 authentication for SMTP")
                    
                    # Get access token - using either dynamic user OAuth or refresh token
                    try:
                        logging.info(f"Getting OAuth access token for {auth_email}")
                        access_token = get_oauth_access_token(
                            auth_email if dynamic_oauth else None,
                            smtp_settings['client_id'],
                            smtp_settings['client_secret'],
                            smtp_settings.get('refresh_token'),
                            dynamic_user_oauth=dynamic_oauth
                        )
                        logging.info(f"Successfully obtained access token for {auth_email}")
                    except Exception as token_error:
                        logging.error(f"Error getting OAuth token for {auth_email}: {str(token_error)}")
                        raise
                    
                    # Generate the authentication string
                    auth_string = generate_oauth2_string(
                        auth_email,  # Use the from_email for authentication
                        access_token
                    )
                    
                    # Using a more robust approach for XOAUTH2 authentication
                    # Close and reopen connection to ensure we start fresh
                    server.close()
                    server = smtplib.SMTP(smtp_settings['smtp_server'], smtp_settings['smtp_port'])
                    server.starttls()
                    server.ehlo()
                    
                    logging.info(f"Attempting OAuth2 login for {auth_email}")
                    
                    # Custom implementation of XOAUTH2 authentication
                    auth_message = f"user={auth_email}\x01auth=Bearer {access_token}\x01\x01"
                    auth_message_base64 = base64.b64encode(auth_message.encode('utf-8')).decode('utf-8')
                    
                    # Initial AUTH command
                    code, resp = server.docmd("AUTH", f"XOAUTH2 {auth_message_base64}")
                    
                    if code == 235:
                        # Authentication successful
                        logging.info(f"OAuth2 authentication successful for {auth_email}")
                    elif code == 334:
                        # Server is expecting client response to a challenge
                        # Send empty response as required by protocol
                        code, resp = server.docmd("", "")
                        if code != 235:
                            error_msg = resp.decode() if isinstance(resp, bytes) else str(resp)
                            logging.error(f"XOAUTH2 challenge-response failed: {code} - {error_msg}")
                            raise smtplib.SMTPAuthenticationError(code, error_msg)
                    else:
                        # Authentication failed
                        error_msg = resp.decode() if isinstance(resp, bytes) else str(resp)
                        logging.error(f"XOAUTH2 authentication failed: {code} - {error_msg}")
                        raise smtplib.SMTPAuthenticationError(code, error_msg)
                        
                    logging.info(f"Successfully authenticated with OAuth as {auth_email}")
                    
                except Exception as oauth_error:
                    logging.error(f"OAuth authentication failed for {auth_email}: {str(oauth_error)}")
                    raise
            else:
                # Use traditional password authentication
                logging.info(f"Using password authentication for SMTP as {smtp_settings['smtp_user_name']}")
                server.login(smtp_settings['smtp_user_name'], smtp_settings['smtp_password'])
                
            server.send_message(msg)
            
        except Exception as e:
            logging.error(f"Failed to send email: {str(e)}")
            raise
        finally:
            # Ensure server connection is closed
            try:
                server.quit()
            except:
                pass
            
        logging.info(f"{'Test email' if is_test else 'Email'} sent successfully to {actual_recipient} using {'OAuth' if use_oauth else 'password'} authentication")
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

def send_email_with_diagnostics(from_email, to_email, subject, html_content, attachment_path=None, smtp_settings=None, is_test=False):
    """
    Send email with enhanced error handling and network diagnostics
    
    This wrapper function provides better error messages and diagnostic information
    when email sending fails, particularly for network-related issues.
    """
    try:
        return send_email(from_email, to_email, subject, html_content, attachment_path, smtp_settings, is_test)
    
    except ConnectionError as e:
        # Handle connection-specific errors with detailed diagnostics
        logging.error(f"Connection error sending email: {str(e)}")
        
        # Run quick diagnostics
        from utils.network_diagnostics import test_network_connectivity, show_connection_help
        connectivity = test_network_connectivity()
        
        error_details = str(e)
        if not connectivity.get('internet'):
            error_details += "\n\nNetwork diagnostics indicate no internet connection."
        elif not connectivity.get('gmail_smtp'):
            error_details += "\n\nNetwork diagnostics indicate Gmail SMTP is not reachable."
        
        error_details += "\n\n" + show_connection_help()
        
        # Re-raise with enhanced error message
        raise ConnectionError(error_details) from e
    
    except Exception as e:
        # Handle other email sending errors
        error_msg = str(e)
        
        # Check if it's a timeout-related error
        if "10060" in error_msg or "timeout" in error_msg.lower() or "timed out" in error_msg.lower():
            from utils.network_diagnostics import show_connection_help
            error_msg += "\n\nThis appears to be a network timeout issue."
            error_msg += show_connection_help()
        
        logging.error(f"Email sending failed: {error_msg}")
        raise Exception(error_msg) from e