"""
OAuth Token Manager

This module provides functions to manage user-specific OAuth tokens
for the NSNA Mail Merge Tool. It handles token storage, retrieval,
and dynamic authentication for users.
"""

import os
import json
import logging
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# Gmail API scopes required for SMTP access
SCOPES = ['https://mail.google.com/']

def get_tokens_dir():
    """Return the directory path for storing user tokens"""
    tokens_dir = Path(__file__).parent.parent / 'user_tokens'
    tokens_dir.mkdir(exist_ok=True)
    return tokens_dir

def get_user_token_path(email):
    """Get the token file path for a specific user email"""
    if not email:
        return None
        
    # Create a filename based on email (sanitized)
    email_filename = ''.join(c if c.isalnum() else '_' for c in email)
    return get_tokens_dir() / f"{email_filename}_token.json"

def get_user_credentials(email, client_id, client_secret):
    """
    Get OAuth credentials for the specified user.
    If no stored credentials exist or they're invalid, prompt for authentication.
    
    Args:
        email: The user's email address
        client_id: The OAuth client ID
        client_secret: The OAuth client secret
        
    Returns:
        Credentials object or None if authentication failed
    """
    token_path = get_user_token_path(email)
    creds = None
    
    # Check if we have a valid token file
    if token_path and token_path.exists():
        try:
            with open(token_path, 'r') as token:
                creds_data = json.load(token)
                creds = Credentials.from_authorized_user_info(
                    creds_data, SCOPES)
        except Exception as e:
            logging.error(f"Error loading credentials: {e}")
            # Continue to authentication flow if loading failed
    
    # If no valid credentials, we need to authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                logging.error(f"Error refreshing token: {e}")
                # Token refresh failed, prompt for re-authentication
                creds = None
        
        if not creds:
            # Need to authenticate from scratch
            try:
                # Prepare client config
                client_config = {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "redirect_uris": ["http://localhost", "urn:ietf:wg:oauth:2.0:oob"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token"
                    }
                }
                
                logging.info(f"Starting OAuth flow for {email}")
                
                flow = InstalledAppFlow.from_client_config(
                    client_config, SCOPES)
                # Use headless mode to prevent UI issues and handle the browser window better
                creds = flow.run_local_server(port=0, open_browser=True, authorization_prompt_message="")
                
                # Save the credentials
                if token_path:
                    with open(token_path, 'w') as token:
                        token.write(creds.to_json())
                        
                logging.info(f"Successfully authenticated {email}")
                        
            except Exception as e:
                logging.error(f"Authentication error: {e}")
                return None
                
    return creds

def clear_user_token(email):
    """
    Clear the stored token for a specific user
    
    Args:
        email: The user's email address
    """
    token_path = get_user_token_path(email)
    if token_path and token_path.exists():
        try:
            os.remove(token_path)
            return True
        except Exception as e:
            logging.error(f"Error clearing token for {email}: {e}")
            return False
    return False
