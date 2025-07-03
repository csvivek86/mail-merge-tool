EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
# Default email is used only if no "From Email" is provided in the app
DEFAULT_EMAIL = 'webcommunications@achi.org'  
# OAuth 2.0 settings - you will need to obtain these from Google Cloud Console
# These are application-level credentials for your Google Cloud project
# Required for dynamic user-specific OAuth
GOOGLE_CLIENT_ID = '707888835132-2lknrcd62uifc978ukvmk1adqa7p5nnj.apps.googleusercontent.com'  # Fill this with your OAuth client ID (REQUIRED for OAuth)
GOOGLE_CLIENT_SECRET = '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'  # Fill this with your OAuth client secret (REQUIRED for OAuth)
# User-specific tokens will be stored in a separate file per user
SUBJECT = 'NSNA Donation Receipt'
TEMPLATE_PATH = 'src/templates/letterhead_template.pdf'
EXCEL_FILE_PATH = 'data/contacts.xlsx'

# Flag to determine which authentication method to use
# True = Use OAuth (recommended)
# False = Use password authentication
USE_OAUTH = True

# Flag to enable user-specific OAuth (each user authenticates with their own Google account)
# When True, the app will prompt for authentication when needed
DYNAMIC_USER_OAUTH = True  # Only used when USE_OAUTH is True

EMAIL_SETTINGS = {
    'smtp_server': EMAIL_HOST,
    'smtp_port': EMAIL_PORT,
    'smtp_user_name': DEFAULT_EMAIL,  # This will be overridden by the app's "From Email"
    # Legacy app password (will be used if USE_OAUTH is False)
    'smtp_password': '',  # Replace with the 16-character App Password from Google
    'subject': SUBJECT,
    # OAuth settings
    'use_oauth': USE_OAUTH,
    'dynamic_user_oauth': DYNAMIC_USER_OAUTH,
    'client_id': GOOGLE_CLIENT_ID,
    'client_secret': GOOGLE_CLIENT_SECRET,
    # refresh_token is not needed here as we'll get it during authentication
}
