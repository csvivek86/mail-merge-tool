from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

def read_excel(file_path):
    """
    Read contacts from Excel file
    
    Args:
        file_path (Path): Path to Excel file
    
    Returns:
        list: List of dictionaries containing contact information
    """
    df = pd.read_excel(str(file_path))
    return df.to_dict('records')

def get_email_addresses(contacts):
    # Extract email addresses from the contacts
    return [contact['Email'] for contact in contacts]

def get_variables(contacts):
    # Extract other variables from the contacts
    return [{key: contact[key] for key in contact if key != 'Email'} for contact in contacts]