from docx import Document
from bs4 import BeautifulSoup
import html5lib
from pathlib import Path
import logging

def convert_docx_to_html(docx_path):
    """Convert Word document to HTML while preserving formatting"""
    try:
        doc = Document(docx_path)
        html_content = []
        
        # Start with basic HTML structure
        html_content.append('<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>')
        
        # Convert paragraphs to HTML
        for para in doc.paragraphs:
            if para.text.strip():  # Skip empty paragraphs
                style = para.style.name
                if style.startswith('Heading'):
                    level = style[-1]  # Get heading level
                    html_content.append(f'<h{level}>{para.text}</h{level}>')
                else:
                    html_content.append(f'<p>{para.text}</p>')
        
        # Add placeholder for dynamic content
        html_content.append('<!-- DONATION_DETAILS_PLACEHOLDER -->')
        html_content.append('</body></html>')
        
        return '\n'.join(html_content)
    except Exception as e:
        logging.error(f"Failed to convert Word template: {str(e)}")
        raise