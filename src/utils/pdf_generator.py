from fpdf import FPDF
from datetime import datetime
import logging
from pathlib import Path
import PyPDF2
import io
import tempfile
import re
from html import unescape

# Import reportlab for better PDF handling
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
    logging.info("✅ ReportLab successfully imported")
except ImportError:
    REPORTLAB_AVAILABLE = False
    logging.warning("⚠️  ReportLab not available")

# IronPDF disabled due to .NET runtime issues - keeping for future reference
IRONPDF_AVAILABLE = False
logging.info("📋 IronPDF disabled due to runtime issues, using ReportLab as primary method")

class PDFGenerator:
    def __init__(self, template_path=None):
        self.template_path = template_path
        self.pdf = FPDF()
        self.pdf.add_page()
        self.pdf.set_auto_page_break(auto=True, margin=15)
        self.pdf.set_margins(20, 20, 20)
    
    def _convert_template_content(self):
        """Add basic letterhead content for PDF template"""
        try:
            if not self.template_path:
                return
            
            # Add basic NSNA letterhead content
            self.pdf.set_font('Arial', 'B', 16)
            self.pdf.cell(0, 10, 'Nagarathar Sangam of North America', ln=True, align='C')
            self.pdf.set_font('Arial', '', 12)
            self.pdf.cell(0, 10, 'Official Donation Receipt', ln=True, align='C')
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

    def generate_receipt(self, receipt_data):
        """Generate PDF receipt - compatibility method for Streamlit app"""
        try:
            logging.info(f"generate_receipt called with keys: {list(receipt_data.keys())}")
            
            # Create temporary output directory
            import tempfile
            output_dir = Path(tempfile.gettempdir())
            
            # Convert receipt_data to donor_info format expected by generate_donation_receipt
            donor_info = {
                'First Name': receipt_data.get('First Name', 'Sample'),
                'Last Name': receipt_data.get('Last Name', 'User'),
                'Region': receipt_data.get('Region', 'Sample Region'),
                'Donation Amount': receipt_data.get('Amount', receipt_data.get('Donation Amount', '100.00')),
                'Value of Item': receipt_data.get('Value of Item', 0)
            }
            
            logging.info(f"Converted donor_info: {donor_info}")
            
            # If template path is provided in receipt_data, use it
            if 'template_path' in receipt_data:
                self.template_path = receipt_data['template_path']
                logging.info(f"Using template_path: {self.template_path}")
            
            # Generate enhanced receipt with custom content
            return self.generate_enhanced_receipt(donor_info, receipt_data, output_dir)
            
        except Exception as e:
            logging.error(f"Failed to generate receipt: {str(e)}")
            raise

    def generate_enhanced_receipt(self, donor_info, template_data, output_dir):
        """Generate PDF receipt with custom template content overlaid on letterhead PDF"""
        try:
            # Use ReportLab as primary method (IronPDF disabled due to .NET runtime issues)
            if REPORTLAB_AVAILABLE:
                logging.info("🚀 Attempting PDF generation with ReportLab (primary method)")
                return self.generate_enhanced_receipt_reportlab(donor_info, template_data, output_dir)
            else:
                logging.info("📋 Using PyPDF2 fallback method")
                return self.generate_enhanced_receipt_pypdf2(donor_info, template_data, output_dir)
                
        except Exception as e:
            logging.error(f"PDF generation failed: {str(e)}")
            # Ultimate fallback to simple receipt
            return self._generate_simple_receipt(donor_info, template_data, output_dir)

    def generate_enhanced_receipt_pypdf2(self, donor_info, template_data, output_dir):
        """Generate PDF receipt with custom template content overlaid on letterhead PDF"""
        try:
            logging.info(f"generate_enhanced_receipt called with template_data keys: {list(template_data.keys())}")
            logging.info(f"Template data structure: {template_data}")
            
            # Check for letterhead PDF - try both locations
            letterhead_paths = [
                Path('NSNA Atlanta Letterhead Updated.pdf'),  # Root directory
                Path('src/templates/letterhead_template.pdf'),  # Templates directory
                Path(template_data.get('template_path', ''))  # Explicit path from template
            ]
            
            letterhead_path = None
            for path in letterhead_paths:
                if path.exists():
                    letterhead_path = path
                    logging.info(f"Found letterhead at: {letterhead_path}")
                    break
            
            if not letterhead_path:
                logging.warning("No letterhead PDF found, creating receipt without letterhead")
                return self._generate_simple_receipt(donor_info, template_data, output_dir)
            
            logging.info(f"Using letterhead from: {letterhead_path}")
            
            # Create content PDF with FPDF - using transparent background
            content_pdf_buffer = io.BytesIO()
            self.pdf = FPDF()
            self.pdf.add_page()
            
            # Set margins and positioning for letterhead overlay
            # Updated positioning: move content up slightly and increase width
            # Left margin stays at 2 inches, top reduced to ~1.7 inches, right reduced for more content width
            
            left_margin = 144    # 144 points (2 inches) - keep left positioning
            top_margin = 120     # 120 points (~1.67 inches) - moved up from 2 inches
            right_margin = 50    # Reduced right margin to allow more content width
            
            self.pdf.set_margins(left_margin, top_margin, right_margin)
            
            # Position cursor at the start of available content area
            self.pdf.set_xy(left_margin, top_margin)
            
            # Check if we have the new consolidated content structure
            if 'content' in template_data:
                logging.info("✅ Using new consolidated content structure")
                # New consolidated structure - use the complete content
                content = template_data['content']
                logging.info(f"Original content length: {len(content)} characters")
                logging.info(f"Content preview: {content[:100]}...")
                
                # Replace all variables in the content
                for key, value in donor_info.items():
                    content = content.replace(f"{{{key}}}", str(value))
                
                # Replace additional common variables
                content = content.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
                content = content.replace('{Year}', datetime.now().strftime('%Y'))
                content = content.replace('{Date}', datetime.now().strftime('%Y-%m-%d'))
                
                logging.info(f"Content after variable replacement: {len(content)} characters")
                
                # Parse HTML content and preserve formatting for PDF
                import re
                from html import unescape
                
                logging.info(f"Content before HTML processing: {content[:200]}...")
                
                # Split content into segments with formatting information
                # This approach maintains bold, italic, and other formatting
                content_segments = []
                
                # Handle paragraphs and divs first - but preserve bold/italic tags
                content = re.sub(r'<p[^>]*>(.*?)</p>', r'\1\n\n', content, flags=re.DOTALL)
                content = re.sub(r'<div[^>]*>(.*?)</div>', r'\1\n', content, flags=re.DOTALL)
                content = re.sub(r'<br\s*/?>', '\n', content)
                
                # Handle lists
                content = re.sub(r'<li[^>]*>(.*?)</li>', r'- \1\n', content, flags=re.DOTALL)
                content = re.sub(r'<ul[^>]*>(.*?)</ul>', r'\1', content, flags=re.DOTALL)
                content = re.sub(r'<ol[^>]*>(.*?)</ol>', r'\1', content, flags=re.DOTALL)
                
                logging.info(f"Content after HTML structure processing: {content[:200]}...")
                
                # Now split content by lines to process formatting per line
                lines = content.split('\n')
                
                # Clean up any problematic characters for PDF but preserve formatting tags
                for i, line in enumerate(lines):
                    # Replace smart quotes and special characters
                    line = line.replace('"', '"').replace('"', '"')
                    line = line.replace(''', "'").replace(''', "'")
                    line = line.replace('–', '-').replace('—', '-')
                    line = line.replace('…', '...')
                    
                    # Don't remove HTML formatting tags yet - we'll handle them in _process_formatted_line
                    lines[i] = line
                
                # Calculate available width for text with updated margins
                # Standard page width is 595 points (8.5 inches * 72 points/inch)
                # With 2-inch left margin (144pt) and reduced right margin (50pt)
                available_width = 595 - left_margin - right_margin  # About 401 points (~5.57 inches) for content
                
                logging.info(f"Processing {len(lines)} lines of content with optimized margins")
                logging.info(f"Updated positioning: left={left_margin}pt (2.0in), top={top_margin}pt ({top_margin/72:.2f}in)")
                logging.info(f"Available content width: {available_width}pt (~{available_width/72:.2f}in)")
                
                # Ensure we're positioned correctly
                self.pdf.set_xy(left_margin, top_margin)
                
                # Process each line with proper formatting support
                for i, line in enumerate(lines):
                    line = line.strip()
                    if line:
                        # Check for and process formatting in each line
                        self._process_formatted_line(line, available_width)
                        logging.info(f"Processed line {i+1}/{len(lines)}: {line[:50]}{'...' if len(line) > 50 else ''}")
                    else:
                        # Empty line - add minimal spacing
                        self.pdf.ln(3)
                
            else:
                logging.warning("⚠️  Using legacy structure - 'content' field not found")
                logging.info(f"Available template fields: {list(template_data.keys())}")
                # Legacy structure support - fallback to old method
                # Add custom greeting if provided
                if 'greeting' in template_data:
                    self.pdf.set_font('Arial', '', 11)
                    greeting = template_data['greeting']
                    # Replace variables in greeting
                    for key, value in donor_info.items():
                        greeting = greeting.replace(f"{{{key}}}", str(value))
                    self.pdf.multi_cell(0, 5, greeting)
                    self.pdf.ln(6)
                
                # Add custom thank you message if provided
                if 'thank_you_message' in template_data:
                    self.pdf.set_font('Arial', '', 10)
                    thank_you = template_data['thank_you_message']
                    # Replace variables in thank you message
                    for key, value in donor_info.items():
                        thank_you = thank_you.replace(f"{{{key}}}", str(value))
                    self.pdf.multi_cell(0, 4, thank_you)
                    self.pdf.ln(8)
                
                # Add donation details section with better spacing
                self.pdf.set_font('Arial', 'B', 11)
                self.pdf.cell(0, 6, 'Donation Details:', ln=True)
                self.pdf.ln(2)
                
                # Add date and donor information with consistent formatting
                self.pdf.set_font('Arial', '', 10)
                self.pdf.cell(0, 5, f'Date: {datetime.now().strftime("%B %d, %Y")}', ln=True)
                self.pdf.cell(0, 5, f'Donor: {donor_info["First Name"]} {donor_info["Last Name"]}', ln=True)
                
                # Add custom donation statement if provided
                if 'donation_statement' in template_data:
                    statement = template_data['donation_statement']
                    # Replace variables in donation statement
                    for key, value in donor_info.items():
                        statement = statement.replace(f"{{{key}}}", str(value))
                    # Also replace common variables
                    statement = statement.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
                    statement = statement.replace('{Year}', datetime.now().strftime('%Y'))
                    self.pdf.cell(0, 5, statement, ln=True)
                else:
                    # Default donation information
                    amount = float(donor_info.get("Donation Amount", 0))
                    self.pdf.cell(0, 5, f'Donation Amount: ${amount:.2f} - Tax Year: {datetime.now().strftime("%Y")}', ln=True)
                
                # Add value of items if applicable
                if donor_info.get("Value of Item", 0) > 0:
                    value = float(donor_info["Value of Item"])
                    self.pdf.cell(0, 5, f'Value of Items Received: ${value:.2f}', ln=True)
                
                # Add standard tax receipt language (legacy mode only)
                self.pdf.ln(6)
                self.pdf.set_font('Arial', '', 9)
                self.pdf.multi_cell(0, 4, 'No goods or services were provided in exchange for your contribution.')
                self.pdf.ln(2)
                self.pdf.multi_cell(0, 4, 'Nagarathar Sangam of North America is a registered Section 501(c)(3) non-profit organization (EIN #: 22-3974176). For questions regarding this acknowledgment letter, please contact us at treasurer@achi.org.')
                self.pdf.ln(3)
                self.pdf.multi_cell(0, 4, 'We truly appreciate your donation and look forward to your continued support of our mission.')
                
                # Add closing (legacy mode only)
                self.pdf.ln(6)
                self.pdf.set_font('Arial', '', 10)
                self.pdf.cell(0, 5, 'Sincerely,', ln=True)
                self.pdf.ln(2)
                self.pdf.cell(0, 5, 'Treasurer, NSNA', ln=True)
                self.pdf.cell(0, 5, f'({datetime.now().strftime("%Y")}-{datetime.now().year+1} term)', ln=True)
            
            # Save content PDF to buffer
            self.pdf.output(content_pdf_buffer, 'S')
            content_pdf_buffer.seek(0)
            
            # Now merge letterhead and content using PyPDF2
            try:
                # Read letterhead PDF
                with open(letterhead_path, 'rb') as letterhead_file:
                    letterhead_reader = PyPDF2.PdfReader(letterhead_file)
                    
                    if len(letterhead_reader.pages) == 0:
                        raise Exception("Letterhead PDF has no pages")
                    
                    letterhead_page = letterhead_reader.pages[0]
                    
                    # Read content PDF
                    content_reader = PyPDF2.PdfReader(content_pdf_buffer)
                    
                    if len(content_reader.pages) == 0:
                        raise Exception("Content PDF has no pages")
                    
                    content_page = content_reader.pages[0]
                    
                    # Create a new PDF writer
                    writer = PyPDF2.PdfWriter()
                    
                    # Start with the letterhead as the base
                    final_page = letterhead_page
                    
                    # Merge content onto letterhead (content overlays on top)
                    final_page.merge_page(content_page)
                    
                    # Add the merged page to writer
                    writer.add_page(final_page)
                    
                    # Generate unique filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    donor_name = f"{donor_info.get('First Name', 'Unknown')}_{donor_info.get('Last Name', 'Donor')}"
                    filename = output_dir / f'NSNA_Receipt_{donor_name}_{timestamp}.pdf'
                    
                    # Write final PDF
                    with open(filename, 'wb') as output_file:
                        writer.write(output_file)
                    
                    logging.info(f"Successfully generated PDF receipt with letterhead: {filename}")
                    return filename
                    
            except Exception as pdf_error:
                logging.error(f"PDF merging failed: {str(pdf_error)}")
                raise pdf_error

        except Exception as e:
            logging.error(f"Failed to generate enhanced PDF with letterhead: {str(e)}")
            # Fallback to simple PDF generation
            return self._generate_simple_receipt(donor_info, template_data, output_dir)

    def generate_enhanced_receipt_reportlab(self, donor_info, template_data, output_dir):
        """Generate PDF receipt using ReportLab for better letterhead overlay"""
        try:
            if not REPORTLAB_AVAILABLE:
                logging.warning("ReportLab not available, falling back to PyPDF2")
                return self.generate_enhanced_receipt_pypdf2(donor_info, template_data, output_dir)
            
            logging.info("🔧 Using ReportLab for enhanced PDF generation")
            
            # Check for letterhead PDF
            letterhead_path = None
            letterhead_paths = [
                Path('NSNA Atlanta Letterhead Updated.pdf'),
                Path('src/templates/letterhead_template.pdf'),
                Path(template_data.get('template_path', ''))
            ]
            
            for path in letterhead_paths:
                if path.exists():
                    letterhead_path = path
                    logging.info(f"Found letterhead at: {letterhead_path}")
                    break
            
            if not letterhead_path:
                logging.warning("No letterhead PDF found")
                return self._generate_simple_receipt(donor_info, template_data, output_dir)
            
            # Generate output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            donor_name = f"{donor_info.get('First Name', 'Unknown')}_{donor_info.get('Last Name', 'Donor')}"
            temp_content_file = output_dir / f'temp_content_{timestamp}.pdf'
            final_filename = output_dir / f'NSNA_Receipt_ReportLab_{donor_name}_{timestamp}.pdf'
            
            # Create content PDF using ReportLab with optimized margins
            doc = SimpleDocTemplate(str(temp_content_file), pagesize=letter,
                                   rightMargin=50, leftMargin=144, topMargin=120, bottomMargin=72)
            
            story = []
            styles = getSampleStyleSheet()
            
            # Create custom styles
            normal_style = styles['Normal']
            normal_style.fontSize = 11
            normal_style.fontName = 'Helvetica'
            
            bold_style = styles['Normal'].clone('BoldStyle')
            bold_style.fontSize = 11
            bold_style.fontName = 'Helvetica-Bold'
            
            if 'content' in template_data:
                content = template_data['content']
                
                # Replace variables
                for key, value in donor_info.items():
                    content = content.replace(f"{{{key}}}", str(value))
                
                content = content.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
                content = content.replace('{Year}', datetime.now().strftime('%Y'))
                content = content.replace('{Date}', datetime.now().strftime('%Y-%m-%d'))
                
                # Parse HTML content for ReportLab
                # Simple HTML to ReportLab conversion
                lines = content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line:
                        story.append(Spacer(1, 6))
                        continue
                    
                    # Remove HTML tags and detect formatting
                    clean_line = re.sub(r'<[^>]+>', '', line)
                    clean_line = unescape(clean_line)
                    
                    if not clean_line.strip():
                        continue
                    
                    # Detect bold formatting
                    is_bold = (
                        '<strong>' in line.lower() or '<b>' in line.lower() or
                        '**' in line or 'donation details' in line.lower() or
                        any(keyword in line.lower() for keyword in ['dear ', 'sincerely', 'thank you'])
                    )
                    
                    # Create paragraph with appropriate style
                    if is_bold:
                        para = Paragraph(clean_line, bold_style)
                    else:
                        para = Paragraph(clean_line, normal_style)
                    
                    story.append(para)
                    story.append(Spacer(1, 3))
            
            else:
                # Legacy format handling
                if 'greeting' in template_data:
                    greeting = template_data['greeting']
                    for key, value in donor_info.items():
                        greeting = greeting.replace(f"{{{key}}}", str(value))
                    story.append(Paragraph(greeting, bold_style))
                    story.append(Spacer(1, 12))
                
                if 'thank_you_message' in template_data:
                    thank_you = template_data['thank_you_message']
                    for key, value in donor_info.items():
                        thank_you = thank_you.replace(f"{{{key}}}", str(value))
                    story.append(Paragraph(thank_you, normal_style))
                    story.append(Spacer(1, 12))
            
            # Build the PDF
            doc.build(story)
            
            # Now merge with letterhead using PyPDF2
            try:
                with open(letterhead_path, 'rb') as letterhead_file:
                    letterhead_reader = PyPDF2.PdfReader(letterhead_file)
                    letterhead_page = letterhead_reader.pages[0]
                    
                    with open(temp_content_file, 'rb') as content_file:
                        content_reader = PyPDF2.PdfReader(content_file)
                        content_page = content_reader.pages[0]
                        
                        # Create final PDF
                        writer = PyPDF2.PdfWriter()
                        
                        # Merge content onto letterhead
                        letterhead_page.merge_page(content_page)
                        writer.add_page(letterhead_page)
                        
                        # Write final PDF
                        with open(final_filename, 'wb') as output_file:
                            writer.write(output_file)
                
                # Clean up temporary file
                if temp_content_file.exists():
                    temp_content_file.unlink()
                
                logging.info(f"✅ Successfully generated PDF with ReportLab: {final_filename}")
                
                if final_filename.exists():
                    file_size = final_filename.stat().st_size
                    logging.info(f"Generated PDF size: {file_size} bytes")
                
                return final_filename
                
            except Exception as merge_error:
                logging.error(f"PDF merging failed: {str(merge_error)}")
                # Clean up temp file
                if temp_content_file.exists():
                    temp_content_file.unlink()
                raise merge_error
                
        except Exception as e:
            logging.error(f"Failed to generate PDF with ReportLab: {str(e)}")
            return self.generate_enhanced_receipt_pypdf2(donor_info, template_data, output_dir)

    def _generate_simple_receipt(self, donor_info, template_data, output_dir):
        """Fallback method to generate receipt without letterhead overlay"""
        try:
            # Reset PDF for new generation
            self.pdf = FPDF()
            self.pdf.add_page()
            self.pdf.set_margins(20, 20, 20)
            
            # Add basic header
            self.pdf.set_font('Arial', 'B', 16)
            self.pdf.cell(0, 10, 'Nagarathar Sangam of North America', ln=True, align='C')
            self.pdf.set_font('Arial', '', 12)
            self.pdf.cell(0, 10, 'Official Donation Receipt', ln=True, align='C')
            self.pdf.ln(15)
            
            # Add custom greeting if provided
            if 'greeting' in template_data:
                self.pdf.set_font('Arial', '', 12)
                greeting = template_data['greeting']
                for key, value in donor_info.items():
                    greeting = greeting.replace(f"{{{key}}}", str(value))
                self.pdf.multi_cell(0, 8, greeting)
                self.pdf.ln(5)
            
            # Add custom thank you message if provided
            if 'thank_you_message' in template_data:
                self.pdf.set_font('Arial', '', 11)
                thank_you = template_data['thank_you_message']
                for key, value in donor_info.items():
                    thank_you = thank_you.replace(f"{{{key}}}", str(value))
                self.pdf.multi_cell(0, 6, thank_you)
                self.pdf.ln(8)
            
            # Add donation details
            self.pdf.set_font('Arial', 'B', 12)
            self.pdf.cell(0, 10, 'Donation Details:', ln=True)
            self.pdf.set_font('Arial', '', 11)
            
            self.pdf.cell(0, 8, f'Date: {datetime.now().strftime("%B %d, %Y")}', ln=True)
            self.pdf.cell(0, 8, f'Donor: {donor_info["First Name"]} {donor_info["Last Name"]}', ln=True)
            
            # Add custom donation statement
            if 'donation_statement' in template_data:
                statement = template_data['donation_statement']
                for key, value in donor_info.items():
                    statement = statement.replace(f"{{{key}}}", str(value))
                statement = statement.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
                statement = statement.replace('{Year}', datetime.now().strftime('%Y'))
                self.pdf.cell(0, 8, statement, ln=True)
            else:
                self.pdf.cell(0, 8, f'Amount: ${float(donor_info["Donation Amount"]):.2f}', ln=True)
            
            # Add tax notice
            self.pdf.ln(15)
            self.pdf.set_font('Arial', 'I', 9)
            self.pdf.multi_cell(0, 6, 'This letter serves as your official receipt for tax purposes.')
            
            # Generate filename and save
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = output_dir / f'receipt_{donor_info["First Name"]}_{donor_info["Last Name"]}_{timestamp}.pdf'
            
            self.pdf.output(str(filename))
            logging.info(f"Generated simple PDF receipt: {filename}")
            
            return filename
            
        except Exception as e:
            logging.error(f"Failed to generate simple PDF: {str(e)}")
            raise

    def _process_formatted_line(self, line, available_width):
        """Process a line of text with HTML formatting and render to PDF"""
        import re
        from html import unescape
        
        try:
            # Clean up the line first
            clean_line = line.strip()
            if not clean_line:
                return
            
            # Save original line for better formatting detection
            original_line = clean_line
            
            # Detect formatting patterns more precisely - only from actual HTML tags
            # Don't auto-bold based on content patterns, only use explicit formatting
            has_bold = (
                bool(re.search(r'<(b|strong)[^>]*>', original_line, re.IGNORECASE)) or
                '**' in original_line  # Markdown bold syntax
            )
            
            has_italic = (
                bool(re.search(r'<(i|em)[^>]*>', original_line, re.IGNORECASE)) or
                (original_line.startswith('*') and not original_line.startswith('**') and original_line.endswith('*'))
            )
            
            # Remove HTML tags and markdown formatting but preserve content
            clean_line = re.sub(r'<[^>]+>', '', clean_line)
            clean_line = clean_line.replace('**', '')  # Remove markdown bold markers
            clean_line = clean_line.replace('*', '')   # Remove single asterisks
            clean_line = unescape(clean_line)
            
            # Skip empty lines after cleaning
            if not clean_line.strip():
                self.pdf.ln(3)
                return
            
            # Handle bullet points
            if clean_line.startswith('-') or clean_line.startswith('•'):
                self.pdf.set_font('Arial', '', 11)
                # Use specified width for consistent letterhead positioning
                self.pdf.multi_cell(available_width, 5, clean_line, ln=1)
                return
            
            # Apply formatting based on what we detected
            if has_bold:
                self.pdf.set_font('Arial', 'B', 11)
                logging.info(f"✅ BOLD formatting applied to: {clean_line[:50]}...")
            elif has_italic:
                self.pdf.set_font('Arial', 'I', 11)
                logging.info(f"✅ ITALIC formatting applied to: {clean_line[:50]}...")
            else:
                self.pdf.set_font('Arial', '', 11)
            
            # Render the line with specified width for consistent positioning
            self.pdf.multi_cell(available_width, 5, clean_line, ln=1)
            
            # Add some spacing after each line for readability
            self.pdf.ln(1)
            
        except Exception as e:
            logging.error(f"❌ Error processing formatted line: {str(e)}")
            # Fallback: render as plain text
            fallback_text = re.sub(r'<[^>]+>', '', line).replace('**', '').replace('*', '')
            fallback_text = unescape(fallback_text) if fallback_text else line
            if fallback_text.strip():
                self.pdf.set_font('Arial', '', 11)
                self.pdf.multi_cell(available_width, 5, fallback_text.strip(), ln=1)