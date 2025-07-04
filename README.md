# NSNA Mail Merge Tool

Enhanced mail merge application for NSNA Atlanta with modern UI and rich text template editing.

## Features

- **Rich Text PDF Template Editor**: Create styled PDF templates with formatting toolbar
- **Variable Insertion**: Easy insertion of Excel field variables (First Name, Last Name, etc.)
- **Live Preview**: Real-time preview of templates with sample data
- **Email Configuration**: Customizable email templates and settings
- **Excel Integration**: Import donor data and track email status
- **PDF Generation**: Generate personalized receipts with attachments
- **Modern UI**: Clean, tabbed interface with color-coded buttons and status indicators
- **OAuth Authentication**: Secure Google OAuth 2.0 support for SMTP authentication
- **Dynamic User OAuth**: Each user can send emails using their own Google account

## Installation

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Verify Installation**:
   ```bash
   python3 test_implementation.py
   ```

## How to Run the Application

### Method 1: Using the Launcher Script (Recommended)
```bash
python3 run_app.py
```

### Method 2: Direct Execution
```bash
cd src
python3 app.py
```

### Method 3: From Python
```python
from src.app import main
main()
```

## Usage Guide

### 1. **Email Settings Tab**
- Enter your "From Email" address
- Select or create an Excel file with donor data
- Configure receipt storage location
- Use "Send Test Email" to test configuration

### 2. **Email Template Tab**
- Customize email subject, greeting, body, and signature
- Save template settings for reuse

### 3. **PDF Template Tab**
Choose between two editing modes:

#### Rich Text Editor:
- Use formatting toolbar (Bold, Italic, Underline, Font Size)
- Insert variables using dropdown menu
- View live preview with sample data
- Save rich template settings

#### Form Fields:
- Traditional form-based editing
- Configure all template sections
- Compatible with existing templates

## Required Excel Columns

Your Excel file must contain these columns:
- `First Name`
- `Last Name` 
- `Email`
- `Donation Amount`

Optional columns:
- `Email Status` (tracks sent/failed status)

## Configuration

The application creates configuration files in `src/config/`:
- `email_template_settings.json`: Email template configuration
- `template_settings.json`: PDF template settings  
- `rich_template_settings.json`: Rich text template settings
- `app_settings.json`: Application preferences

## Attachment Handling

The application includes robust attachment processing:

1. **PDF Generation**: Creates personalized receipts for each donor
2. **Email Attachment**: Attaches PDF receipts to outgoing emails
3. **Error Handling**: Continues email sending even if PDF generation fails
4. **Multiple Format Support**: Ready for future attachment type expansion

## Troubleshooting

### Common Issues

1. **Import Errors**: 
   ```bash
   pip install -r requirements.txt
   ```

2. **PDF Template Not Found**:
   - Place your PDF template in the project directory
   - Update template path in settings

3. **Email Sending Fails**:
   - Check SMTP settings in `src/config/settings.py`
   - Verify email credentials and server settings

4. **WeasyPrint Installation Issues** (macOS):
   ```bash
   brew install cairo pango gdk-pixbuf libffi
   pip install weasyprint
   ```

### Logs

Application logs are stored in `logs/app.log` for debugging.

## Dependencies

- **PyQt6**: Modern GUI framework
- **pandas**: Excel file processing
- **reportlab**: PDF generation
- **PyPDF2**: PDF manipulation
- **openpyxl**: Excel file handling
- **weasyprint**: HTML to PDF conversion (for rich templates)
- **Pillow**: Image processing

## Project Structure

```
mail-merge-tool
├── src
│   ├── main.py                # Entry point of the application
│   ├── config
│   │   └── settings.py        # Configuration settings for the application
│   ├── templates
│   │   └── email_template.html # HTML template for the emails
│   └── utils
│       ├── excel_reader.py     # Functions to read data from the Excel file
│       └── mail_sender.py      # Functions to send emails
├── data
│   └── contacts.xlsx          # Excel file containing recipient data
├── requirements.txt           # Project dependencies
└── README.md                  # Project documentation
```

## Requirements

To run this project, you need to install the following dependencies:

- `pandas`
- `openpyxl`
- `smtplib` (part of the Python standard library)

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Usage

1. Update the `src/config/settings.py` file with your email server details and sender email.
2. Populate the `data/contacts.xlsx` file with the email addresses and any other variables you want to include in your emails.
3. Modify the `src/templates/email_template.html` file to customize your email content.
4. Run the application:

```
python src/main.py
```

This will initiate the mail merge process, sending personalized emails to each recipient listed in the Excel file.

## License

This project is open-source and available for use and modification.

# Installation
## Windows
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```
## Building the Application

### Mac
```bash
# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Build the application
python build.py
```

The executable will be created in the `dist/NSNA-Mail-Merge` directory.

### Windows
```bash
# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Build the application
python build.py
```

## Google OAuth 2.0 Setup

The application supports two OAuth authentication methods:

### 1. Dynamic User OAuth (Recommended)

With Dynamic User OAuth, each person using the application can authenticate with their own Google account. This allows emails to be sent from the user's own email address, which provides better email deliverability and tracking.

#### Setting up Dynamic User OAuth:

1. **Create a Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project

2. **Enable the Gmail API**:
   - In your project, go to "APIs & Services" > "Library"
   - Search for "Gmail API" and enable it

3. **Create OAuth Credentials**:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Select "Desktop App" as the application type
   - Name your client and click "Create"
   - Download the JSON file (this is your credentials file)

4. **Run the OAuth Setup Tool**:
   ```bash
   python setup_oauth.py
   ```
   - Follow the prompts and provide the path to your downloaded credentials.json
   - The tool will extract and store the client ID and client secret in settings.py

5. **Using Dynamic User OAuth**:
   - Ensure that both `USE_OAUTH = True` and `DYNAMIC_USER_OAUTH = True` in `src/config/settings.py`
   - When users enter their email in the "From Email" field and send an email, they'll be prompted to authenticate with Google
   - Each user's authentication token is securely stored for future use
   - Users will only need to authenticate once per email address

### 2. Static OAuth (Alternative)

With Static OAuth, the application uses a single OAuth token for all email sending, regardless of the "From Email" field.

#### Setting up Static OAuth:

1. Follow steps 1-4 from Dynamic User OAuth setup above
   
2. After completing the OAuth setup, set:
   - `USE_OAUTH = True`
   - `DYNAMIC_USER_OAUTH = False` in `src/config/settings.py`

3. The application will use the refresh token stored in settings.py for all email sending

### 3. App Password Authentication (Legacy Method)

If you prefer to use app passwords instead of OAuth:

1. Set `USE_OAUTH = False` in `src/config/settings.py`
2. Update `EMAIL_SETTINGS['smtp_password']` with your app password

## Security Notes

- With Dynamic User OAuth, the application never sees or stores user passwords
- Each user's OAuth token is stored in a separate file in the `src/user_tokens` directory
- OAuth tokens can be revoked at any time through the user's Google account settings
- The application requires the "https://mail.google.com/" scope to send emails on behalf of users
