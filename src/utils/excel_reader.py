from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

def read_excel(file_path):
    """
    Read contacts from Excel file
    
    Args:
        file_path (Path): Path to Excel file
    
    Returns:
        DataFrame: Pandas DataFrame containing contact information
    """
    df = pd.read_excel(str(file_path))
    return df

def get_contacts_as_list(df):
    """
    Convert DataFrame to list of dictionaries for compatibility
    
    Args:
        df (DataFrame): Pandas DataFrame
    
    Returns:
        list: List of dictionaries containing contact information
    """
    return df.to_dict('records')

def get_email_addresses(df):
    """Extract email addresses from the DataFrame"""
    if 'Email' in df.columns:
        return df['Email'].tolist()
    else:
        # Try common email column names
        for col in ['email', 'EMAIL', 'Email Address', 'email_address']:
            if col in df.columns:
                return df[col].tolist()
        return []

def get_variables(df):
    """Extract other variables from the DataFrame"""
    return df.to_dict('records')