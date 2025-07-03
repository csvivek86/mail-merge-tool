EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_HOST_USER = 'webcommunications@achi.org'
EMAIL_HOST_PASSWORD = 'txbuxzecmvdzhjxw'
SUBJECT = 'NSNA Donation Receipt'
TEMPLATE_PATH = 'src/templates/letterhead_template.pdf'  # Update this to your PDF template path
EXCEL_FILE_PATH = 'data/contacts.xlsx'

EMAIL_SETTINGS = {
    'smtp_server': EMAIL_HOST,
    'smtp_port': EMAIL_PORT,
    'smtp_user_name': EMAIL_HOST_USER,
    'smtp_password': EMAIL_HOST_PASSWORD,
    'subject': SUBJECT
}
