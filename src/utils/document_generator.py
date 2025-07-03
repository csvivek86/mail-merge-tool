from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PyPDF2 import PdfReader, PdfWriter
from pathlib import Path
import logging
from datetime import datetime
import io
import json
from config.template_settings import TemplateSettings

class DocumentGenerator:
    def __init__(self, template_path, template_settings, contact_info):
        """Initialize DocumentGenerator with required parameters
        
        Args:
            template_path: Path to PDF template
            template_settings: Template settings for formatting
            contact_info: Contact information for the receipt
        """
        self.template_path = template_path
        self.settings = template_settings
        self.contact_info = contact_info
        self.margins = {
            'left': 2 * inch,
            'right': 0.5 * inch,
            'top': 2 * inch,
            'bottom': 0.5 * inch
        }
        self.spacing = {
            'line': 0.25 * inch,      # Space between lines within a section
            'section': 0.5 * inch     # Space between different sections
        }

    def _load_settings(self) -> TemplateSettings:
        """Load or create template settings"""
        try:
            config_dir = Path.cwd() / 'config'
            config_dir.mkdir(exist_ok=True)
            settings_path = config_dir / 'template_settings.json'

            if settings_path.exists():
                with open(settings_path) as f:
                    return TemplateSettings(**json.load(f))
            return TemplateSettings()  # Return default settings if file doesn't exist

        except Exception as e:
            logging.error(f"Settings load error: {e}")
            return TemplateSettings()  # Fallback to defaults

    def _write_content(self, canvas_obj, content, y_pos, font="Helvetica", size=12, spacing=0.25):
        """Write content to PDF with consistent formatting"""
        canvas_obj.setFont(font, size)
        if isinstance(content, list):
            for line in content:
                canvas_obj.drawString(self.margins['left'], y_pos, line)
                y_pos -= spacing * inch
        else:
            canvas_obj.drawString(self.margins['left'], y_pos, content)
            y_pos -= spacing * inch
        return y_pos

    def generate_receipt(self, donor_info, output_dir):
        """Generate donation receipt PDF"""
        try:
            # Format date for filename
            date_str = datetime.now().strftime("%Y-%m-%d")
            # Create filename with First Name, Last Name and date
            base_name = f"{donor_info['First Name']}_{donor_info['Last Name']}_{date_str}"
            pdf_path = output_dir / f"{base_name}.pdf"
            
            # Setup PDF generation
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_path = output_dir / f'receipt_{donor_info["First Name"]}_{donor_info["Last Name"]}_{timestamp}.pdf'
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=letter)
            y_pos = letter[1] - self.margins['top']

            # Write content sections with custom spacing
            sections = [
                # Date (standalone line)
                [(datetime.now().strftime('%B %d, %Y'), "Helvetica", 12)],
                
                # Greeting (standalone line)
                [(f"{self.settings.greeting} {donor_info['First Name']} {donor_info['Last Name']},", "Helvetica-Bold", 12)],
                
                # Thank you section
                [(line, "Helvetica", 12) for line in self.settings.thank_you_lines],
                
                # Receipt statement (standalone line)
                [(self.settings.receipt_statement, "Helvetica", 12)],
                
                # Donation details (grouped together with line spacing)
                [(f"Donation Amount: ${float(donor_info['Donation Amount']):.2f}", "Helvetica-Bold", 12),
                 (f"Donation(s) Year: {self.settings.donation_year}", "Helvetica-Bold", 12)],
                
                # Disclaimer (standalone line)
                [(self.settings.disclaimer, "Helvetica", 12)],
                
                # Organization info section
                [(line, "Helvetica", 12) for line in self.settings.org_info],
                
                # Closing lines section
                [(line, "Helvetica", 12) for line in self.settings.closing_lines],
                
                # Signature section
                [(line, "Helvetica", 12) for line in self.settings.signature]
            ]

            # Process each section
            for section in sections:
                # Write lines within section
                for i, (content, font, size) in enumerate(section):
                    c.setFont(font, size)
                    c.drawString(self.margins['left'], y_pos, content)
                    # Use line spacing within section, section spacing after last line
                    y_pos -= self.spacing['line'] if i < len(section) - 1 else self.spacing['section']

            # Generate final PDF
            c.save()
            packet.seek(0)
            template = PdfReader(open(self.template_path, "rb"))
            new_content = PdfReader(packet)
            
            output = PdfWriter()
            page = template.pages[0]
            page.merge_page(new_content.pages[0])
            output.add_page(page)

            with open(pdf_path, "wb") as f:
                output.write(f)

            logging.info(f"Generated receipt: {pdf_path}")
            return pdf_path

        except Exception as e:
            logging.error(f"Receipt generation failed: {e}")
            raise

    def generate_receipt_from_html(self, donor_info, output_dir, html_template=None):
        """Generate donation receipt PDF from HTML template overlaid on letterhead"""
        try:
            import tempfile
            from PyPDF2 import PdfReader, PdfWriter
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.units import inch
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_LEFT
            import re
            
            # Format date for timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_path = output_dir / f'receipt_{donor_info["First Name"]}_{donor_info["Last Name"]}_{timestamp}.pdf'
            
            if html_template:
                # Use provided HTML template
                processed_html = self._process_html_template_simple(html_template, donor_info)
            else:
                # Generate HTML from current settings
                processed_html = self._generate_html_from_settings(donor_info)
            
            # Create overlay PDF with HTML content
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_overlay:
                overlay_path = temp_overlay.name
                
                # Create overlay PDF with parsed HTML content
                doc = SimpleDocTemplate(
                    overlay_path,
                    pagesize=letter,
                    rightMargin=1*inch,
                    leftMargin=2*inch,
                    topMargin=2*inch,
                    bottomMargin=1*inch
                )
                
                # Parse HTML and convert to reportlab elements (without CSS)
                elements = self._html_to_reportlab_elements(processed_html)
                doc.build(elements)
            
            try:
                # Merge with letterhead template
                if self.template_path and self.template_path.exists():
                    template_reader = PdfReader(str(self.template_path))
                    overlay_reader = PdfReader(overlay_path)
                    writer = PdfWriter()
                    
                    # Get the first page of both PDFs
                    template_page = template_reader.pages[0]
                    overlay_page = overlay_reader.pages[0]
                    
                    # Merge overlay onto template
                    template_page.merge_page(overlay_page)
                    writer.add_page(template_page)
                    
                    # Write final PDF
                    with open(pdf_path, 'wb') as output_file:
                        writer.write(output_file)
                    
                    logging.info(f"Generated HTML-based receipt with letterhead: {pdf_path}")
                else:
                    # If no template, just use the overlay
                    import shutil
                    shutil.move(overlay_path, str(pdf_path))
                    logging.info(f"Generated HTML-based receipt (no letterhead): {pdf_path}")
                
                return pdf_path
                
            finally:
                # Clean up temporary file
                Path(overlay_path).unlink(missing_ok=True)
            
        except Exception as e:
            logging.error(f"HTML receipt generation failed: {e}")
            # Fall back to standard PDF generation
            return self.generate_receipt(donor_info, output_dir)

    def _process_html_template_simple(self, html_template, donor_info):
        """Process HTML template with donor information - simplified version"""
        replacements = {
            'First Name': donor_info.get('First Name', ''),
            'Last Name': donor_info.get('Last Name', ''),
            'Email': donor_info.get('Email', ''),
            'Donation Amount': f"{float(donor_info.get('Donation Amount', 0)):.2f}",
            'Date': datetime.now().strftime('%B %d, %Y'),
            'Current Year': str(datetime.now().year)
        }
        
        processed_html = html_template
        for var, value in replacements.items():
            processed_html = processed_html.replace(f'{{{{{var}}}}}', str(value))
        
        return processed_html

    def _html_to_reportlab_elements(self, html_content):
        """Convert HTML content to reportlab elements, removing CSS styling"""
        from reportlab.platypus import Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.lib.enums import TA_LEFT
        import re
        
        styles = getSampleStyleSheet()
        elements = []
        
        # Custom styles
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=12,
            leftIndent=0
        )
        
        bold_style = ParagraphStyle(
            'CustomBold',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=12,
            leftIndent=0,
            fontName='Helvetica-Bold'
        )
        
        # Remove all CSS styling and HTML structure, keeping only content
        # Remove DOCTYPE, style tags, and unwanted HTML
        clean_content = re.sub(r'<!DOCTYPE[^>]*>', '', html_content, flags=re.IGNORECASE)
        clean_content = re.sub(r'<style[^>]*>.*?</style>', '', clean_content, flags=re.IGNORECASE | re.DOTALL)
        clean_content = re.sub(r'<head>.*?</head>', '', clean_content, flags=re.IGNORECASE | re.DOTALL)
        clean_content = re.sub(r'<html[^>]*>', '', clean_content, flags=re.IGNORECASE)
        clean_content = re.sub(r'</html>', '', clean_content, flags=re.IGNORECASE)
        clean_content = re.sub(r'<body[^>]*>', '', clean_content, flags=re.IGNORECASE)
        clean_content = re.sub(r'</body>', '', clean_content, flags=re.IGNORECASE)
        
        # Remove style attributes from remaining tags
        clean_content = re.sub(r' style="[^"]*"', '', clean_content)
        clean_content = re.sub(r' margin-[^=]*="[^"]*"', '', clean_content)
        clean_content = re.sub(r' -qt-[^=]*="[^"]*"', '', clean_content)
        
        # Extract text content from HTML paragraphs
        paragraphs = re.findall(r'<p[^>]*>(.*?)</p>', clean_content, re.DOTALL | re.IGNORECASE)
        
        if not paragraphs:
            # If no paragraphs found, split by line breaks
            clean_text = re.sub(r'<[^>]+>', '', clean_content)
            lines = clean_text.split('\n')
            for line in lines:
                line = line.strip()
                if line:
                    if any(keyword in line.lower() for keyword in ['dear', 'donation amount', 'donation(s) year']):
                        elements.append(Paragraph(line, bold_style))
                    else:
                        elements.append(Paragraph(line, normal_style))
                else:
                    elements.append(Spacer(1, 6))
        else:
            # Process paragraphs
            for para in paragraphs:
                # Remove any remaining HTML tags from paragraph content
                clean_para = re.sub(r'<[^>]+>', '', para).strip()
                
                if clean_para:
                    # Check if paragraph should be bold
                    if any(keyword in clean_para.lower() for keyword in ['dear', 'donation amount', 'donation(s) year']):
                        elements.append(Paragraph(clean_para, bold_style))
                    else:
                        elements.append(Paragraph(clean_para, normal_style))
                else:
                    elements.append(Spacer(1, 6))
        
        return elements

    def _process_html_template(self, html_template, donor_info):
        """Process HTML template with donor information"""
        replacements = {
            'First Name': donor_info.get('First Name', ''),
            'Last Name': donor_info.get('Last Name', ''),
            'Email': donor_info.get('Email', ''),
            'Donation Amount': f"{float(donor_info.get('Donation Amount', 0)):.2f}",
            'Date': datetime.now().strftime('%B %d, %Y'),
            'Current Year': str(datetime.now().year)
        }
        
        processed_html = html_template
        for var, value in replacements.items():
            processed_html = processed_html.replace(f'{{{{{var}}}}}', str(value))
        
        # Add basic CSS styling
        css_styles = """
        <style>
        body { 
            font-family: Arial, sans-serif; 
            line-height: 1.6; 
            margin: 2in 1in 1in 2in;
            font-size: 12pt;
        }
        .donation-amount { font-weight: bold; }
        .signature { margin-top: 2em; }
        </style>
        """
        
        if '<head>' in processed_html:
            processed_html = processed_html.replace('</head>', f'{css_styles}</head>')
        else:
            processed_html = f"<html><head>{css_styles}</head><body>{processed_html}</body></html>"
        
        return processed_html

    def _generate_html_from_settings(self, donor_info):
        """Generate HTML from current template settings"""
        html_parts = [
            "<html><head>",
            "<style>",
            "body { font-family: Arial, sans-serif; line-height: 1.6; margin: 2in 1in 1in 2in; font-size: 12pt; }",
            ".donation-amount { font-weight: bold; }",
            ".signature { margin-top: 2em; }",
            "</style>",
            "</head><body>",
            f"<p>{datetime.now().strftime('%B %d, %Y')}</p>",
            f"<p><strong>{self.settings.greeting} {donor_info['First Name']} {donor_info['Last Name']},</strong></p>",
            "<br>"
        ]
        
        # Add thank you lines
        for line in self.settings.thank_you_lines:
            html_parts.append(f"<p>{line}</p>")
        
        html_parts.extend([
            "<br>",
            f"<p>{self.settings.receipt_statement}</p>",
            "<br>",
            f"<p class='donation-amount'>Donation Amount: ${float(donor_info['Donation Amount']):.2f}</p>",
            f"<p class='donation-amount'>Donation(s) Year: {self.settings.donation_year}</p>",
            "<br>",
            f"<p>{self.settings.disclaimer}</p>",
            "<br>"
        ])
        
        # Add org info
        for line in self.settings.org_info:
            html_parts.append(f"<p>{line}</p>")
        
        html_parts.append("<br>")
        
        # Add closing lines
        for line in self.settings.closing_lines:
            html_parts.append(f"<p>{line}</p>")
        
        html_parts.append("<br><div class='signature'>")
        
        # Add signature
        for line in self.settings.signature:
            html_parts.append(f"<p>{line}</p>")
        
        html_parts.extend(["</div>", "</body></html>"])
        
        return "".join(html_parts)