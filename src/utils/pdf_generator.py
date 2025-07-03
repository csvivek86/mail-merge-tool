from fpdf import FPDF
# from docx import Document
# from docx.oxml.shared import qn
# from docx.text.paragraph import Paragraph
# from docx.table import _Cell, Table
from datetime import datetime
import logging
from pathlib import Path

class PDFGenerator:
    def __init__(self, template_path=None):
        self.template_path = template_path
        self.pdf = FPDF()
        self.pdf.add_page()
        self.pdf.set_auto_page_break(auto=True, margin=15)
        self.pdf.set_margins(20, 20, 20)
    
    def _add_header_footer(self, doc):
        """Add header and footer content from Word document"""
        try:
            # Get header content
            header = doc.sections[0].header
            if header:
                self.pdf.set_y(10)  # Position at top
                for paragraph in header.paragraphs:
                    if paragraph.text.strip():
                        self.pdf.set_font('Arial', 'B', 12)
                        self.pdf.cell(0, 10, paragraph.text.strip(), ln=True, align='C')
        
            # Get footer content
            footer = doc.sections[0].footer
            if footer:
                footer_text = ' '.join([p.text.strip() for p in footer.paragraphs if p.text.strip()])
                if footer_text:
                    # Save current position
                    current_y = self.pdf.get_y()
                    
                    # Move to bottom of page
                    self.pdf.set_y(-30)
                    self.pdf.set_font('Arial', '', 10)
                    self.pdf.cell(0, 10, footer_text, ln=True, align='C')
                    
                    # Restore position
                    self.pdf.set_y(current_y)
        except Exception as e:
            logging.error(f"Failed to add header/footer: {str(e)}")
            raise
    
    def _add_text_boxes(self, doc):
        """Add text box content from Word document"""
        try:
            # Get shapes that contain text boxes
            for shape in doc.inline_shapes:
                if hasattr(shape, 'text_box'):
                    text = shape.text_box.text.strip()
                    if text:
                        # Save current position
                        current_y = self.pdf.get_y()
                        
                        # Add text box content on left margin
                        self.pdf.set_xy(10, current_y)
                        self.pdf.set_font('Arial', '', 10)
                        self.pdf.multi_cell(50, 5, text)  # Narrow width for margin text
                        
                        # Restore position
                        self.pdf.set_xy(20, current_y)
        except Exception as e:
            logging.warning(f"Could not process text boxes: {str(e)}")

    def _convert_template_content(self):
        """Extract and convert content from Word template"""
        try:
            if not self.template_path:
                return
            
            doc = Document(self.template_path)
            
            # Add header and footer first
            self._add_header_footer(doc)
            
            # Add text boxes
            self._add_text_boxes(doc)
            
            # Start with letterhead content
            self.pdf.set_font('Arial', '', 12)
            
            # Process main document paragraphs
            for para in doc.paragraphs:
                if para.text.strip():
                    if para.style.name.startswith('Heading'):
                        self.pdf.set_font('Arial', 'B', 14)
                    else:
                        self.pdf.set_font('Arial', '', 12)
                    
                    self.pdf.multi_cell(0, 10, para.text.strip())
                    self.pdf.ln(5)
            
            self.pdf.ln(10)
            
        except Exception as e:
            logging.error(f"Error converting template: {str(e)}")
            raise


    def generate_donation_receipt(self, donor_info, output_dir):
        """Generate PDF receipt using template and donor information"""
        try:
            # Reset PDF for new generation
            self.pdf = FPDF()
            self.pdf.add_page()
            self.pdf.set_margins(20, 20, 20)
            
            # First, add the letterhead content from template
            if self.template_path:
                logging.info(f"Using template from: {self.template_path}")
                self._convert_template_content()
            
            # Add receipt header
            self.pdf.set_font('Arial', 'B', 14)
            self.pdf.cell(0, 10, 'Donation Receipt', ln=True, align='C')
            self.pdf.ln(5)
            
            # Add date and donor information
            self.pdf.set_font('Arial', '', 12)
            self.pdf.cell(0, 10, f'Date: {datetime.now().strftime("%B %d, %Y")}', ln=True)
            self.pdf.cell(0, 10, f'Donor: {donor_info["First Name"]} {donor_info["Last Name"]}', ln=True)
            self.pdf.cell(0, 10, f'Region: {donor_info["Region"]}', ln=True)
            
            # Add donation details
            self.pdf.ln(5)
            self.pdf.cell(0, 10, 'Donation Details:', ln=True)
            self.pdf.cell(0, 10, f'Amount: ${float(donor_info["Donation Amount"]):.2f}', ln=True)
            
            if donor_info.get("Value of Item", 0) > 0:
                self.pdf.cell(0, 10, f'Value of Items: ${float(donor_info["Value of Item"]):.2f}', ln=True)
            
            # Add tax receipt notice
            self.pdf.ln(10)
            self.pdf.multi_cell(0, 10, 'This letter serves as your official receipt for tax purposes.')
            
            # Generate unique filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = output_dir / f'receipt_{donor_info["First Name"]}_{donor_info["Last Name"]}_{timestamp}.pdf'
            
            # Save PDF
            self.pdf.output(str(filename))
            logging.info(f"Generated PDF receipt: {filename}")
            
            return filename

        except Exception as e:
            logging.error(f"Failed to generate PDF: {str(e)}")
            raise