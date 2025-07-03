from openpyxl import Workbook
from pathlib import Path

def create_excel_template(save_path: Path) -> Path:
    """Create a new Excel file with required headers"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Donation Data"
    
    # Define required headers with corrected names
    headers = [
        "First Name",
        "Last Name",
        "Email",
        "Region",
        "Donation Amount",  # Changed from Amount
        "Value of Item",    # Changed from Date
        "Email Status"
    ]
    
    # Add headers to first row
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
        
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    
    # Save workbook
    wb.save(save_path)
    return save_path