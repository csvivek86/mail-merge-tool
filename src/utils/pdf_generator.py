from fpdf import FPDF
from datetime import datetime
import logging
from pathlib import Path
import PyPDF2
import io
import tempfile
import re
from html import unescape

# Configure logging to be more verbose for debugging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

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
            # PRIORITY 1: ReportLab HTML method (NO REGEX PARSING!)
            if REPORTLAB_AVAILABLE:
                logging.info("🎨 Using ReportLab HTML method (no regex parsing required)")
                try:
                    result = self.generate_html_pdf_reportlab(donor_info, template_data, output_dir)
                    return result
                except Exception as reportlab_error:
                    logging.error(f"ReportLab method failed: {reportlab_error}")
                    # Fall through to PyPDF2 method
            
            # FALLBACK: PyPDF2 with regex parsing (legacy method)
            logging.info("📋 Falling back to PyPDF2 method with regex parsing")
            try:
                result = self.generate_enhanced_receipt_pypdf2(donor_info, template_data, output_dir)
                return result
            except Exception as pypdf2_error:
                logging.error(f"PyPDF2 method failed: {pypdf2_error}")
                # Fall through to simple receipt
                
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
                Path(__file__).parent.parent.parent / 'NSNA Atlanta Letterhead Updated.pdf',  # Root directory
                Path(__file__).parent.parent / 'templates' / 'letterhead_template.pdf',  # Templates directory
                Path(template_data.get('template_path', '')) if template_data.get('template_path') else None  # Explicit path from template
            ]
            
            letterhead_path = None
            for path in letterhead_paths:
                if path and path.exists():
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
            # OPTIMIZED POSITIONING: Move text slightly left from previous position
            # Based on user feedback: positioning looks better, just move a bit more to the left
            
            left_margin = 160    # 160 points (2.22 inches) - moved slightly left for better positioning
            top_margin = 100     # 100 points (~1.39 inches) - keep vertical position
            right_margin = 50    # Reduced right margin to allow adequate content width
            
            self.pdf.set_margins(left_margin, top_margin, right_margin)
            
            # Position cursor at the start of available content area
            self.pdf.set_xy(left_margin, top_margin)
            
            # Check if we have the new consolidated content structure
            if 'content' in template_data:
                logging.info("✅ Using new consolidated content structure")
                # New consolidated structure - use the complete content
                content = template_data['content']
                logging.info(f"🔍 ORIGINAL CONTENT DEBUG:")
                logging.info(f"   Content type: {type(content)}")
                logging.info(f"   Content length: {len(content)} characters")
                logging.info(f"   First 200 chars: {content[:200]}...")
                logging.info(f"   Last 100 chars: ...{content[-100:]}")
                
                # Check for HTML tags in original content
                import re
                strong_tags = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                em_tags = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                p_tags = re.findall(r'</?p[^>]*>', content, re.IGNORECASE)
                
                # Check for markdown formatting
                markdown_bold = re.findall(r'\*\*[^*]+\*\*', content)
                markdown_italic = re.findall(r'\*[^*]+\*', content)
                
                logging.info(f"🔍 ORIGINAL HTML TAGS FOUND:")
                logging.info(f"   Strong/B tags: {len(strong_tags)} - {strong_tags[:5]}")
                logging.info(f"   Em/I tags: {len(em_tags)} - {em_tags[:5]}")
                logging.info(f"   P tags: {len(p_tags)} - {p_tags[:5]}")
                
                logging.info(f"🔍 MARKDOWN FORMATTING FOUND:")
                logging.info(f"   **bold** patterns: {len(markdown_bold)} - {markdown_bold[:5]}")
                logging.info(f"   *italic* patterns: {len(markdown_italic)} - {markdown_italic[:5]}")
                
                if not strong_tags and not em_tags and (markdown_bold or markdown_italic):
                    logging.info("🔄 CONVERTING MARKDOWN TO HTML...")
                    
                    # Convert markdown formatting to HTML
                    # Handle bold: **text** -> <strong>text</strong>
                    content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', content)
                    
                    # Handle italic: *text* -> <em>text</em> (but not **text**)
                    # Use negative lookbehind and lookahead to avoid interfering with bold
                    content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', content)
                    
                    # Re-check after conversion
                    strong_tags_after = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                    em_tags_after = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                    
                    logging.info(f"✅ MARKDOWN CONVERSION COMPLETE:")
                    logging.info(f"   Converted to {len(strong_tags_after)} strong tags")
                    logging.info(f"   Converted to {len(em_tags_after)} em tags")
                    logging.info(f"   Content after conversion: {content[:200]}...")
                elif not strong_tags and not em_tags:
                    logging.warning("⚠️  NO FORMATTING TAGS OR MARKDOWN FOUND IN ORIGINAL CONTENT!")
                    logging.info("🔍 Checking for alternative formatting indicators...")
                    if '**' in content:
                        logging.info("   Found ** markdown bold indicators")
                    if '*' in content and '**' not in content:
                        logging.info("   Found * markdown italic indicators")
                    if content.isupper():
                        logging.info("   Content appears to be all uppercase")
                else:
                    logging.info(f"✅ Found {len(strong_tags + em_tags)} HTML formatting tags in original content")
                
                # Replace all variables in the content
                original_content = content  # Keep original for comparison
                for key, value in donor_info.items():
                    content = content.replace(f"{{{key}}}", str(value))
                
                # Replace additional common variables
                content = content.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
                content = content.replace('{Year}', datetime.now().strftime('%Y'))
                content = content.replace('{Date}', datetime.now().strftime('%Y-%m-%d'))
                
                logging.info(f"🔄 AFTER VARIABLE REPLACEMENT:")
                logging.info(f"   Content changed: {original_content != content}")
                logging.info(f"   New length: {len(content)} characters")
                if original_content != content:
                    logging.info(f"   First 200 chars after replacement: {content[:200]}...")
                
                # Re-check formatting tags after variable replacement
                strong_tags_after = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                em_tags_after = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                
                logging.info(f"🔍 FORMATTING TAGS AFTER VARIABLE REPLACEMENT:")
                logging.info(f"   Strong/B tags: {len(strong_tags_after)}")
                logging.info(f"   Em/I tags: {len(em_tags_after)}")
                
                if len(strong_tags_after) == 0 and ('**' in content or '*' in content):
                    logging.warning("⚠️  MARKDOWN FORMATTING DETECTED AFTER VARIABLE REPLACEMENT!")
                    logging.info("🔄 Applying markdown conversion after variable replacement...")
                    
                    # Convert any remaining markdown
                    content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', content)
                    content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', content)
                    
                    # Final check
                    strong_tags_final = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                    em_tags_final = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                    logging.info(f"   Final conversion: {len(strong_tags_final)} strong, {len(em_tags_final)} em tags")
                
                # Parse HTML content and preserve formatting for PDF
                from html import unescape
                
                logging.info(f"🔍 STARTING HTML PROCESSING:")
                logging.info(f"   Content before processing: {content[:200]}...")
                
                # Improved HTML processing to handle streamlit-quill output with inline formatting
                # Strategy: Preserve formatting tags within content while handling structure
                
                # First, handle lists properly (convert to bullet points)
                logging.info("🔄 Processing lists...")
                before_lists = content
                content = re.sub(r'<li[^>]*>(.*?)</li>', r'• \1\n', content, flags=re.DOTALL)
                content = re.sub(r'</?(?:ul|ol)[^>]*>', '', content, flags=re.IGNORECASE)
                if before_lists != content:
                    logging.info(f"   Lists processed: {before_lists.count('<li')} list items converted")
                
                # Handle line breaks and paragraph spacing - improved approach
                logging.info("🔄 Processing line breaks and paragraphs...")
                before_breaks = content
                content = re.sub(r'<br\s*/?>', '\n', content)
                if before_breaks != content:
                    logging.info(f"   Line breaks processed: {before_breaks.count('<br')} <br> tags converted")
                
                # Handle paragraphs while preserving inline formatting
                # Split on paragraph boundaries but preserve content within paragraphs 
                before_paragraphs = content
                content = re.sub(r'<p[^>]*>(.*?)</p>', r'\1\n\n', content, flags=re.DOTALL)
                content = re.sub(r'<div[^>]*>(.*?)</div>', r'\1\n', content, flags=re.DOTALL)
                if before_paragraphs != content:
                    logging.info(f"   Paragraphs processed: {before_paragraphs.count('<p')} <p> tags converted")
                
                # Clean up excessive newlines while preserving intentional spacing
                before_cleanup = content
                content = re.sub(r'\n{3,}', '\n\n', content)  # Max 2 consecutive newlines
                content = content.strip()
                if before_cleanup != content:
                    logging.info(f"   Newlines cleaned up: excessive newlines removed")
                
                # CRITICAL: Check formatting tags after HTML processing
                strong_tags_processed = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                em_tags_processed = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                
                logging.info(f"🔍 FORMATTING TAGS AFTER HTML PROCESSING:")
                logging.info(f"   Strong/B tags: {len(strong_tags_processed)} (was {len(strong_tags_after)})")
                logging.info(f"   Em/I tags: {len(em_tags_processed)} (was {len(em_tags_after)})")
                
                if len(strong_tags_processed) == 0 and len(strong_tags_after) > 0:
                    logging.error("❌ CRITICAL: FORMATTING TAGS LOST DURING HTML PROCESSING!")
                    logging.info(f"   Before processing: {strong_tags_after}")
                    logging.info(f"   After processing: {strong_tags_processed}")
                
                # Keep formatting tags intact for the _parse_inline_formatting function
                
                logging.info(f"Content after HTML structure processing: {content[:200]}...")
                logging.info(f"🔍 Checking for formatting tags in processed content...")
                
                # Log formatting tag detection for debugging
                strong_matches = re.findall(r'</?(?:strong|b)[^>]*>', content, re.IGNORECASE)
                em_matches = re.findall(r'</?(?:em|i)[^>]*>', content, re.IGNORECASE)
                
                if strong_matches:
                    logging.info(f"✅ Found STRONG/B tags: {strong_matches}")
                if em_matches:
                    logging.info(f"✅ Found EM/I tags: {em_matches}")
                
                if not strong_matches and not em_matches:
                    logging.warning("⚠️  No formatting tags found in processed content")
                else:
                    logging.info(f"✅ Total formatting tags preserved: {len(strong_matches) + len(em_matches)}")
                
                # Now split content by lines to process formatting per line
                lines = content.split('\n')
                logging.info(f"🔍 SPLIT INTO LINES:")
                logging.info(f"   Total lines: {len(lines)}")
                logging.info(f"   Non-empty lines: {len([l for l in lines if l.strip()])}")
                
                # Log first few lines for debugging
                for i, line in enumerate(lines[:5]):
                    logging.info(f"   Line {i+1}: '{line[:100]}{'...' if len(line) > 100 else ''}'")
                    if '<strong>' in line.lower() or '<b>' in line.lower():
                        logging.info(f"     ✅ Line {i+1} contains bold formatting")
                    if '<em>' in line.lower() or '<i>' in line.lower():
                        logging.info(f"     ✅ Line {i+1} contains italic formatting")
                
                # Clean up any problematic characters for PDF but preserve formatting tags
                for i, line in enumerate(lines):
                    # Replace smart quotes and special characters
                    line = line.replace('"', '"').replace('"', '"')
                    line = line.replace(''', "'").replace(''', "'")
                    line = line.replace('–', '-').replace('—', '-')
                    line = line.replace('…', '...')
                    
                    # Preserve HTML formatting tags for processing
                    lines[i] = line
                
                # Calculate available width for text with optimized margins  
                # Standard page width is 595 points (8.5 inches * 72 points/inch)
                # With 2.22-inch left margin (160pt) and reduced right margin (50pt)
                available_width = 595 - left_margin - right_margin  # About 385 points (~5.35 inches) for content
                
                logging.info(f"Processing {len(lines)} lines of content with optimized margins")
                logging.info(f"OPTIMIZED positioning: left={left_margin}pt (2.22in), top={top_margin}pt ({top_margin/72:.2f}in)")
                logging.info(f"Available content width: {available_width}pt (~{available_width/72:.2f}in)")
                
                # Ensure we're positioned correctly
                self.pdf.set_xy(left_margin, top_margin)
                
                # Process each line with proper formatting support
                for i, line in enumerate(lines):
                    line = line.strip()
                    if line:
                        # Check for and process formatting in each line
                        logging.info(f"📝 Processing line {i+1}/{len(lines)}: '{line[:80]}{'...' if len(line) > 80 else ''}'")
                        if '<strong>' in line.lower() or '<b>' in line.lower():
                            logging.info(f"   ✅ Line {i+1} has BOLD formatting tags")
                        if '<em>' in line.lower() or '<i>' in line.lower():
                            logging.info(f"   ✅ Line {i+1} has ITALIC formatting tags")
                        if not ('<strong>' in line.lower() or '<b>' in line.lower() or '<em>' in line.lower() or '<i>' in line.lower()):
                            logging.info(f"   ❌ Line {i+1} has NO formatting tags")
                        
                        self._process_formatted_line(line, available_width)
                    else:
                        # Empty line - add appropriate spacing based on context
                        logging.info(f"📝 Processing empty line {i+1}/{len(lines)}")
                        self.pdf.ln(8)  # Increased spacing between paragraphs
                
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
                # Read letterhead PDF data into memory to avoid file handle issues
                with open(letterhead_path, 'rb') as letterhead_file:
                    letterhead_data = letterhead_file.read()
                
                # Create PDF readers from memory buffers
                letterhead_reader = PyPDF2.PdfReader(io.BytesIO(letterhead_data))
                content_reader = PyPDF2.PdfReader(content_pdf_buffer)
                
                if len(letterhead_reader.pages) == 0:
                    raise Exception("Letterhead PDF has no pages")
                
                if len(content_reader.pages) == 0:
                    raise Exception("Content PDF has no pages")
                
                letterhead_page = letterhead_reader.pages[0]
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
                Path(__file__).parent.parent.parent / 'NSNA Atlanta Letterhead Updated.pdf',
                Path(__file__).parent.parent / 'templates' / 'letterhead_template.pdf',
                Path(template_data.get('template_path', '')) if template_data.get('template_path') else None
            ]
            
            for path in letterhead_paths:
                if path and path.exists():
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
            
            # Create content PDF using ReportLab with optimized margins (matching FPDF method)
            doc = SimpleDocTemplate(str(temp_content_file), pagesize=letter,
                                   rightMargin=50, leftMargin=160, topMargin=100, bottomMargin=72)
            
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
                        story.append(Spacer(1, 3))  # Small space for empty lines
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
                # Read letterhead data into memory to avoid file handle issues
                with open(letterhead_path, 'rb') as letterhead_file:
                    letterhead_data = letterhead_file.read()
                
                # Read content data into memory
                with open(temp_content_file, 'rb') as content_file:
                    content_data = content_file.read()
                
                # Create PDF readers from memory buffers
                letterhead_reader = PyPDF2.PdfReader(io.BytesIO(letterhead_data))
                content_reader = PyPDF2.PdfReader(io.BytesIO(content_data))
                
                letterhead_page = letterhead_reader.pages[0]
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

    def generate_html_pdf_reportlab(self, donor_info, template_data, output_dir):
        """Generate PDF using ReportLab's HTML support - NO REGEX PARSING NEEDED!"""
        try:
            if not REPORTLAB_AVAILABLE:
                logging.warning("ReportLab not available, falling back to FPDF")
                return self._generate_simple_receipt(donor_info, template_data, output_dir)
            
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus.frames import Frame
            from reportlab.platypus.doctemplate import PageTemplate
            import tempfile
            
            logging.info("🎨 Using ReportLab HTML rendering - no regex parsing required!")
            
            # Create output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = f"{donor_info.get('First Name', 'Unknown')}_{donor_info.get('Last Name', 'Donor')}"
            safe_name = "".join(c for c in safe_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"NSNA_Receipt_{safe_name}_{timestamp}.pdf"
            
            output_path = Path(output_dir) / filename
            
            # Get HTML content
            html_content = template_data.get('content', '<p>No content provided</p>')
            
            logging.info(f"📄 Processing HTML content length: {len(html_content)}")
            logging.info(f"📄 Content preview: {html_content[:200]}...")
            
            # Log the full content for debugging
            if len(html_content) < 1000:
                logging.info(f"📄 Full content: {html_content}")
            else:
                logging.info(f"📄 Content too long, showing first 500 chars: {html_content[:500]}...")
            
            # MARKDOWN TO HTML CONVERSION for ReportLab method
            if '**' in html_content or '*' in html_content:
                logging.info("🔄 Converting markdown to HTML for ReportLab...")
                
                # Convert markdown formatting to HTML
                original_content = html_content
                html_content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', html_content)
                html_content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', html_content)
                
                if original_content != html_content:
                    logging.info(f"✅ Markdown converted to HTML for ReportLab")
                    logging.info(f"   Content after conversion: {html_content[:200]}...")
                else:
                    logging.info("   No markdown conversion needed")
            
            # Replace variables in the content
            logging.info(f"🔄 Replacing variables in content...")
            for key, value in donor_info.items():
                old_content = html_content
                html_content = html_content.replace(f"{{{key}}}", str(value))
                if old_content != html_content:
                    logging.info(f"   Replaced {{{key}}} with {value}")
            
            html_content = html_content.replace('{Amount}', str(donor_info.get('Donation Amount', '0.00')))
            html_content = html_content.replace('{Year}', datetime.now().strftime('%Y'))
            html_content = html_content.replace('{Date}', datetime.now().strftime('%Y-%m-%d'))
            
            logging.info(f"📄 Final content after variable replacement: {html_content[:200]}...")
            
            # Create PDF document with improved positioning to avoid letterhead overlap
            doc = SimpleDocTemplate(
                str(output_path),
                pagesize=letter,
                topMargin=1.5*inch,      # Increased to avoid letterhead overlap
                leftMargin=2.0*inch,     # Optimized left positioning
                rightMargin=0.75*inch,   # Right margin
                bottomMargin=1.0*inch    # Bottom margin
            )
            
            logging.info(f"📏 PDF margins set - Top: 1.5in, Left: 2.0in, Right: 0.75in, Bottom: 1.0in")
            
            # Get styles and create improved custom style
            styles = getSampleStyleSheet()
            
            # Create custom style that handles HTML properly with normal spacing
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=11,
                leading=13,              # Normal line spacing (1.2x font size)
                spaceAfter=4,            # Minimal space after paragraphs
                spaceBefore=0,           # No space before paragraphs
                allowWidows=1,
                allowOrphans=1,
                wordWrap='CJK'           # Better word wrapping
            )
            
            # Build the story (content)
            story = []
            
            # Enhanced paragraph processing for better HTML formatting support
            logging.info(f"🔄 Processing HTML content for ReportLab paragraphs...")
            
            # Import BeautifulSoup for proper HTML parsing if available
            try:
                from bs4 import BeautifulSoup
                USE_BEAUTIFUL_SOUP = True
                logging.info("✅ Using BeautifulSoup for proper HTML parsing")
            except ImportError:
                USE_BEAUTIFUL_SOUP = False
                logging.info("⚠️  BeautifulSoup not available, using basic HTML parsing")
            
            if USE_BEAUTIFUL_SOUP:
                # Use BeautifulSoup for proper HTML parsing
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # Process each element in the HTML
                for element in soup.find_all(['p', 'div', 'br']):
                    if element.name == 'br':
                        story.append(Spacer(1, 3))  # Small space for line breaks
                        continue
                    
                    # Get text content while preserving formatting tags
                    if element.name in ['p', 'div']:
                        # Extract HTML content with formatting preserved
                        element_html = str(element)
                        
                        # Convert to paragraph-friendly HTML
                        element_html = element_html.replace('<p>', '').replace('</p>', '')
                        element_html = element_html.replace('<div>', '').replace('</div>', '')
                        
                        if element_html.strip():
                            logging.info(f"     Processing element: {element_html[:50]}{'...' if len(element_html) > 50 else ''}")
                            
                            # Check for formatting
                            has_formatting = any(tag in element_html for tag in ['<strong>', '<b>', '<em>', '<i>', '<u>'])
                            logging.info(f"       Has HTML formatting: {has_formatting}")
                            
                            # ReportLab Paragraph handles HTML tags directly
                            try:
                                story.append(Paragraph(element_html, normal_style))
                                story.append(Spacer(1, 2))
                            except Exception as para_error:
                                logging.warning(f"       Failed to create paragraph: {para_error}")
                                # Fallback: strip tags and use plain text
                                plain_text = element.get_text()
                                story.append(Paragraph(plain_text, normal_style))
                                story.append(Spacer(1, 2))
                
                # If no block elements found, treat as plain text with line breaks
                if not soup.find_all(['p', 'div']):
                    text_content = soup.get_text()
                    paragraphs = text_content.split('\n\n')
                    for para in paragraphs:
                        para = para.strip()
                        if para:
                            story.append(Paragraph(para, normal_style))
                            story.append(Spacer(1, 2))
            else:
                # Fallback: Enhanced basic HTML parsing
                # Clean up the HTML content first
                content_to_process = html_content.strip()
                
                # Convert line breaks to proper paragraph breaks if no <p> tags exist
                if '<p>' not in content_to_process and '<div>' not in content_to_process:
                    # Split by double newlines to create paragraphs
                    paragraphs = content_to_process.split('\n\n')
                    logging.info(f"   Split into {len(paragraphs)} paragraphs by double newlines")
                    
                    for i, para in enumerate(paragraphs):
                        para = para.strip().replace('\n', '<br/>')  # Convert single newlines to HTML breaks
                        if para:
                            logging.info(f"     Paragraph {i+1}: {para[:50]}{'...' if len(para) > 50 else ''}")
                            
                            # Check if paragraph has formatting
                            has_formatting = any(tag in para for tag in ['<strong>', '<b>', '<em>', '<i>', '<u>', '**'])
                            logging.info(f"       Has HTML formatting: {has_formatting}")
                            
                            # Convert any remaining markdown to HTML
                            para = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', para)
                            para = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', para)
                            
                            # ReportLab Paragraph can handle HTML tags directly
                            try:
                                story.append(Paragraph(para, normal_style))
                                story.append(Spacer(1, 2))  # Minimal spacing between paragraphs
                            except Exception as para_error:
                                logging.warning(f"       Failed to create paragraph: {para_error}")
                                # Strip HTML and use plain text
                                plain_text = re.sub(r'<[^>]+>', '', para)
                                story.append(Paragraph(plain_text, normal_style))
                                story.append(Spacer(1, 2))
                else:
                    # HTML content with block-level tags
                    # Split by paragraph or div tags
                    import re
                    
                    # Extract content between paragraph tags
                    para_pattern = r'<(?:p|div)[^>]*>(.*?)</(?:p|div)>'
                    matches = re.findall(para_pattern, content_to_process, re.DOTALL | re.IGNORECASE)
                    
                    if matches:
                        logging.info(f"   Found {len(matches)} HTML block elements")
                        
                        for i, para in enumerate(matches):
                            para = para.strip()
                            if para:
                                logging.info(f"     Paragraph {i+1}: {para[:50]}{'...' if len(para) > 50 else ''}")
                                
                                # Check if paragraph has formatting
                                has_formatting = any(tag in para for tag in ['<strong>', '<b>', '<em>', '<i>', '<u>'])
                                logging.info(f"       Has HTML formatting: {has_formatting}")
                                
                                # ReportLab Paragraph can handle basic HTML tags!
                                try:
                                    story.append(Paragraph(para, normal_style))
                                    story.append(Spacer(1, 2))
                                except Exception as para_error:
                                    logging.warning(f"       Failed to create paragraph: {para_error}")
                                    # Strip HTML and use plain text
                                    plain_text = re.sub(r'<[^>]+>', '', para)
                                    story.append(Paragraph(plain_text, normal_style))
                                    story.append(Spacer(1, 2))
                    else:
                        # No structured HTML found, treat as plain text
                        logging.info("   No structured HTML found, processing as plain text")
                        paragraphs = content_to_process.split('\n\n')
                        for para in paragraphs:
                            para = para.strip()
                            if para:
                                story.append(Paragraph(para, normal_style))
                                story.append(Spacer(1, 2))
            
            # Log final story summary
            logging.info(f"📊 Final story contains {len(story)} elements")
            paragraph_count = len([item for item in story if hasattr(item, 'text')])
            spacer_count = len(story) - paragraph_count
            logging.info(f"   - {paragraph_count} paragraphs")
            logging.info(f"   - {spacer_count} spacers")
            
            # Build PDF
            logging.info("🔨 Building PDF document...")
            doc.build(story)
            logging.info(f"✅ PDF document built successfully: {output_path}")
            
            # Check file size
            if output_path.exists():
                file_size = output_path.stat().st_size
                logging.info(f"📁 Generated PDF size: {file_size} bytes")
            else:
                logging.error("❌ PDF file was not created!")
                return None
            
            # IMPORTANT: Always overlay letterhead if template_path is provided
            # Get the root directory of the project (go up from src/utils/ to root)
            project_root = Path(__file__).parent.parent.parent
            print(f"🔍 LETTERHEAD DEBUG: Project root: {project_root}")
            
            letterhead_paths = [
                project_root / 'NSNA Atlanta Letterhead Updated.pdf',  # Root directory
                project_root / 'src' / 'templates' / 'letterhead_template.pdf',  # Templates directory
                Path(template_data.get('template_path', '')) if template_data.get('template_path') else None  # Explicit path from template
            ]
            
            print(f"🔍 LETTERHEAD DEBUG: Searching for letterhead...")
            print(f"🔍 LETTERHEAD DEBUG: Search paths:")
            for i, path in enumerate(letterhead_paths):
                if path:
                    abs_path = path.absolute()
                    exists = path.exists()
                    print(f"   {i+1}. {abs_path} (exists: {exists})")
                    if exists:
                        size = path.stat().st_size
                        print(f"       Size: {size} bytes")
                else:
                    print(f"   {i+1}. None")
            
            letterhead_path = None
            for path in letterhead_paths:
                if path and path.exists():
                    letterhead_path = path
                    print(f"✅ LETTERHEAD DEBUG: Found letterhead for overlay: {letterhead_path}")
                    logging.info(f"Found letterhead for overlay: {letterhead_path}")
                    break
            
            if letterhead_path:
                print(f"🎨 LETTERHEAD DEBUG: Overlaying content PDF onto letterhead: {letterhead_path}")
                logging.info(f"🎨 Overlaying content PDF onto letterhead: {letterhead_path}")
                output_path = self._overlay_letterhead_reportlab(output_path, letterhead_path)
                print(f"✅ LETTERHEAD DEBUG: Letterhead overlay complete: {output_path}")
                logging.info(f"✅ Letterhead overlay complete: {output_path}")
            else:
                print(f"⚠️ LETTERHEAD DEBUG: No letterhead found - PDF will be plain content only")
                logging.warning("⚠️ No letterhead found - PDF will be plain content only")
                # List what files ARE in the project root for debugging
                print(f"🔍 LETTERHEAD DEBUG: Files in project root:")
                try:
                    for item in project_root.iterdir():
                        if item.is_file() and item.suffix.lower() == '.pdf':
                            print(f"     PDF: {item.name} ({item.stat().st_size} bytes)")
                except Exception as list_error:
                    print(f"     Error listing files: {list_error}")
            
            logging.info(f"✅ PDF generated successfully using ReportLab HTML: {output_path}")
            return str(output_path)
            
        except Exception as e:
            logging.error(f"❌ ReportLab HTML generation failed: {str(e)}")
            # Fallback to existing method
            return self.generate_enhanced_receipt_pypdf2(donor_info, template_data, output_dir)
    
    def _overlay_letterhead_reportlab(self, content_pdf_path, letterhead_path):
        """Overlay content PDF onto letterhead using PyPDF2"""
        try:
            import PyPDF2
            
            print(f"🔗 OVERLAY DEBUG: Starting letterhead overlay process...")
            print(f"📄 OVERLAY DEBUG: Content PDF: {content_pdf_path}")
            print(f"🏢 OVERLAY DEBUG: Letterhead PDF: {letterhead_path}")
            
            logging.info(f"🔗 Starting letterhead overlay process...")
            logging.info(f"📄 Content PDF: {content_pdf_path}")
            logging.info(f"🏢 Letterhead PDF: {letterhead_path}")
            
            # Verify both files exist
            if not Path(content_pdf_path).exists():
                error_msg = f"Content PDF not found: {content_pdf_path}"
                print(f"❌ OVERLAY DEBUG: {error_msg}")
                raise FileNotFoundError(error_msg)
            if not Path(letterhead_path).exists():
                error_msg = f"Letterhead PDF not found: {letterhead_path}"
                print(f"❌ OVERLAY DEBUG: {error_msg}")
                raise FileNotFoundError(error_msg)
            
            print(f"✅ OVERLAY DEBUG: Both files exist")
            
            # Create output path with letterhead suffix
            content_path = Path(content_pdf_path)
            output_path = content_path.parent / f"{content_path.stem}_with_letterhead{content_path.suffix}"
            
            print(f"📝 OVERLAY DEBUG: Output path: {output_path}")
            
            # Read both PDFs and merge in a single transaction to avoid file handle issues
            letterhead_page = None
            content_page = None
            
            # Read the letterhead PDF first and keep the page in memory
            with open(letterhead_path, 'rb') as letterhead_file:
                letterhead_data = letterhead_file.read()
            
            letterhead_reader = PyPDF2.PdfReader(io.BytesIO(letterhead_data))
            if len(letterhead_reader.pages) == 0:
                raise Exception("Letterhead PDF has no pages")
            letterhead_page = letterhead_reader.pages[0]
            print(f"✅ OVERLAY DEBUG: Letterhead page loaded successfully")
            logging.info(f"✅ Letterhead page loaded successfully")
            
            # Read the content PDF and keep the page in memory
            with open(content_pdf_path, 'rb') as content_file:
                content_data = content_file.read()
            
            content_reader = PyPDF2.PdfReader(io.BytesIO(content_data))
            if len(content_reader.pages) == 0:
                raise Exception("Content PDF has no pages")
            content_page = content_reader.pages[0]
            print(f"✅ OVERLAY DEBUG: Content page loaded successfully")
            logging.info(f"✅ Content page loaded successfully")
            
            # Merge content onto letterhead (letterhead is background, content overlays)
            letterhead_page.merge_page(content_page)
            print(f"✅ OVERLAY DEBUG: Pages merged successfully")
            logging.info(f"✅ Pages merged successfully")
            
            # Write the result
            with open(output_path, 'wb') as output_file:
                writer = PyPDF2.PdfWriter()
                writer.add_page(letterhead_page)
                writer.write(output_file)
            
            print(f"✅ OVERLAY DEBUG: Final PDF written: {output_path}")
            logging.info(f"✅ Final PDF written: {output_path}")
            
            # Verify the output file was created and has content
            if output_path.exists():
                file_size = output_path.stat().st_size
                print(f"✅ OVERLAY DEBUG: Output file size: {file_size} bytes")
                logging.info(f"✅ Output file size: {file_size} bytes")
                
                # Clean up original content PDF (keep only the final version)
                try:
                    Path(content_pdf_path).unlink(missing_ok=True)
                    print(f"🗑️ OVERLAY DEBUG: Cleaned up temporary content PDF")
                    logging.info(f"🗑️ Cleaned up temporary content PDF")
                except:
                    pass  # Don't fail if cleanup fails
                
                return str(output_path)
            else:
                raise Exception("Output file was not created")
            
        except Exception as e:
            print(f"❌ OVERLAY DEBUG: Failed to overlay letterhead: {e}")
            logging.error(f"❌ Failed to overlay letterhead: {e}")
            print(f"📄 OVERLAY DEBUG: Returning original content PDF: {content_pdf_path}")
            logging.info(f"📄 Returning original content PDF: {content_pdf_path}")
            return str(content_pdf_path)  # Return original if overlay fails

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
                self.pdf.ln(5)  # Increased spacing for empty lines
                return
            
            # Save original line for better formatting detection
            original_line = clean_line
            logging.info(f"🔍 _process_formatted_line called with: '{original_line[:100]}{'...' if len(original_line) > 100 else ''}'")
            
            # Check for formatting tags in original line
            has_strong = '<strong>' in original_line.lower() or '<b>' in original_line.lower()
            has_em = '<em>' in original_line.lower() or '<i>' in original_line.lower()
            logging.info(f"   Original line formatting check: BOLD={has_strong}, ITALIC={has_em}")
            
            # NEW APPROACH: Process inline formatting by splitting text into segments
            # This handles cases like: "No goods or <strong>services</strong> were provided"
            
            # Find all formatting segments in the line
            segments = self._parse_inline_formatting(original_line)
            logging.info(f"   Parsed into {len(segments)} segments")
            
            if segments:
                logging.info(f"   ✅ Using segment-based formatting approach")
                # Process each segment with its formatting
                current_y = self.pdf.get_y()
                x_position = self.pdf.get_x()
                
                for j, segment in enumerate(segments):
                    text = segment['text']
                    is_bold = segment['bold']
                    is_italic = segment['italic']
                    
                    logging.info(f"     Segment {j+1}: '{text[:30]}{'...' if len(text) > 30 else ''}' BOLD={is_bold} ITALIC={is_italic}")
                    
                    if not text.strip():
                        continue
                    
                    # Apply formatting based on segment
                    if is_bold and is_italic:
                        self.pdf.set_font('Arial', 'BI', 11)
                        logging.info(f"       ✅ Applied BOLD+ITALIC font to: '{text[:20]}...'")
                    elif is_bold:
                        self.pdf.set_font('Arial', 'B', 11)
                        logging.info(f"       ✅ Applied BOLD font to: '{text[:20]}...'")
                    elif is_italic:
                        self.pdf.set_font('Arial', 'I', 11)
                        logging.info(f"       ✅ Applied ITALIC font to: '{text[:20]}...'")
                    else:
                        self.pdf.set_font('Arial', '', 11)
                        logging.info(f"       📝 Applied REGULAR font to: '{text[:20]}...'")
                    
                    # Verify font was set correctly
                    current_font = getattr(self.pdf, 'current_font', {})
                    if hasattr(self.pdf, 'current_font'):
                        logging.info(f"       🔧 Current PDF font: {current_font.get('family', 'unknown')} {current_font.get('style', 'unknown')}")
                    else:
                        logging.info(f"       ⚠️  Could not verify current PDF font")
                    
                    # Calculate text width and render
                    text_width = self.pdf.get_string_width(text)
                    
                    # Check if we need to wrap to next line
                    if x_position + text_width > (self.pdf.w - self.pdf.r_margin):
                        self.pdf.ln()
                        x_position = self.pdf.l_margin
                    
                    # Position and render text with improved spacing
                    self.pdf.set_xy(x_position, current_y if x_position > self.pdf.l_margin else self.pdf.get_y())
                    self.pdf.cell(text_width, 7, text, ln=0)  # Increased cell height for better spacing
                    
                    # Update position for next segment
                    x_position += text_width
                
                # Move to next line after processing all segments
                self.pdf.ln(8)  # Increased spacing between lines for better readability
                
            else:
                # Fallback to original method if no segments found
                # Remove ALL HTML tags and clean content
                clean_line = re.sub(r'<[^>]+>', '', clean_line)
                clean_line = unescape(clean_line)
                
                # Skip empty lines after cleaning
                if not clean_line.strip():
                    self.pdf.ln(5)  # Consistent spacing for empty lines
                    return
                
                # Handle bullet points
                if clean_line.startswith(('-', '•')):
                    self.pdf.set_font('Arial', '', 11)
                    self.pdf.multi_cell(available_width, 6, clean_line, ln=1)  # Increased line height for bullets
                    self.pdf.ln(2)
                    return
                
                # Apply default formatting
                self.pdf.set_font('Arial', '', 11)
                logging.info(f"📝 Fallback - Regular text: {clean_line[:50]}...")
                
                # Render the line with specified width for consistent positioning
                self.pdf.multi_cell(available_width, 6, clean_line, ln=1)  # Increased line height
                self.pdf.ln(2)  # Consistent spacing
            
        except Exception as e:
            logging.error(f"❌ Error processing formatted line: {str(e)}")
            # Fallback: render as plain text
            fallback_text = re.sub(r'<[^>]+>', '', line)
            fallback_text = unescape(fallback_text) if fallback_text else line
            if fallback_text.strip():
                self.pdf.set_font('Arial', '', 11)
                self.pdf.multi_cell(available_width, 6, fallback_text.strip(), ln=1)  # Consistent line height
                self.pdf.ln(2)  # Consistent spacing

    def _parse_inline_formatting(self, html_text):
        """Parse HTML text into segments with formatting information - ENHANCED VERSION"""
        import re
        from html import unescape
        
        try:
            # Clean up the input but preserve formatting tags
            text = html_text.strip()
            
            if not text:
                return []
            
            logging.info(f"🔍 ENHANCED parsing HTML: {text[:100]}{'...' if len(text) > 100 else ''}")
            
            # Pre-process to normalize formatting tags
            original_text = text
            text = re.sub(r'<strong\b[^>]*>', '<strong>', text, flags=re.IGNORECASE)
            text = re.sub(r'<b\b[^>]*>', '<strong>', text, flags=re.IGNORECASE)
            text = re.sub(r'</b>', '</strong>', text, flags=re.IGNORECASE)
            text = re.sub(r'<em\b[^>]*>', '<em>', text, flags=re.IGNORECASE)
            text = re.sub(r'<i\b[^>]*>', '<em>', text, flags=re.IGNORECASE)
            text = re.sub(r'</i>', '</em>', text, flags=re.IGNORECASE)
            
            if original_text != text:
                logging.info(f"🔧 Normalized HTML: {text[:100]}{'...' if len(text) > 100 else ''}")
            else:
                logging.info(f"🔧 No normalization needed")
            
            # Count formatting tags after normalization
            strong_count = text.lower().count('<strong>') + text.lower().count('</strong>')
            em_count = text.lower().count('<em>') + text.lower().count('</em>')
            logging.info(f"   Formatting tag count: {strong_count} strong tags, {em_count} em tags")
            
            segments = []
            
            # Enhanced pattern to capture all content including nested tags
            pattern = r'(</?(?:strong|em)>)'
            parts = re.split(pattern, text, flags=re.IGNORECASE)
            
            # Track current formatting state
            is_bold = False
            is_italic = False
            
            for part in parts:
                if not part:
                    continue
                
                # Check if this part is a formatting tag
                if re.match(r'</?(?:strong|em)>', part, re.IGNORECASE):
                    if part.lower() == '<strong>':
                        is_bold = True
                        logging.info("🎨 Starting BOLD formatting")
                    elif part.lower() == '</strong>':
                        is_bold = False
                        logging.info("🎨 Ending BOLD formatting")
                    elif part.lower() == '<em>':
                        is_italic = True
                        logging.info("🎨 Starting ITALIC formatting")
                    elif part.lower() == '</em>':
                        is_italic = False
                        logging.info("🎨 Ending ITALIC formatting")
                else:
                    # This is text content
                    clean_text = unescape(part).strip()
                    if clean_text:
                        segments.append({
                            'text': clean_text,
                            'bold': is_bold,
                            'italic': is_italic
                        })
                        logging.info(f"📝 Added segment: '{clean_text[:30]}...' Bold:{is_bold} Italic:{is_italic}")
            
            if not segments:
                # Fallback: treat entire text as unformatted if no tags found
                clean_text = re.sub(r'<[^>]+>', '', text)
                clean_text = unescape(clean_text).strip()
                if clean_text:
                    segments.append({
                        'text': clean_text,
                        'bold': False,
                        'italic': False
                    })
                    logging.info(f"📝 Fallback segment: '{clean_text[:30]}...'")
            
            logging.info(f"✅ Total segments parsed: {len(segments)}")
            return segments
            
        except Exception as e:
            logging.error(f"❌ Error in enhanced HTML parsing: {str(e)}")
            # Ultimate fallback
            clean_text = re.sub(r'<[^>]+>', '', html_text)
            clean_text = unescape(clean_text).strip()
            if clean_text:
                return [{'text': clean_text, 'bold': False, 'italic': False}]
            return []