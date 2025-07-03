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
pip install -r [requirements.txt](http://_vscodecontentref_/5)
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