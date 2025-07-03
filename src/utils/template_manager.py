from pathlib import Path
from jinja2 import Template
import logging
from .template_converter import convert_docx_to_html

class TemplateManager:
    def __init__(self):
        self.template_dir = Path(__file__).parent.parent / 'templates'
        self.base_template = None
    
    def load_word_template(self, docx_path):
        """Load and convert Word template"""
        try:
            html_content = convert_docx_to_html(docx_path)
            
            # Insert our Jinja2 template content
            donation_details = """
            <p>Donation Details:</p>
            <ul>
                <li>Amount: ${{ donation_amount }}</li>
                {% if value_of_item > 0 %}
                <li>Value of Items: ${{ value_of_item }}</li>
                {% endif %}
                <li>Region: {{ region }}</li>
            </ul>
            """
            
            # Replace placeholder with donation details
            template_content = html_content.replace(
                '<!-- DONATION_DETAILS_PLACEHOLDER -->', 
                donation_details
            )
            
            self.base_template = Template(template_content)
            return True
        except Exception as e:
            logging.error(f"Failed to load template: {str(e)}")
            return False
    
    def render_template(self, contact_data):
        """Render template with contact data"""
        if not self.base_template:
            raise ValueError("No template loaded")
            
        return self.base_template.render(
            first_name=contact_data['First Name'],
            last_name=contact_data['Last Name'],
            donation_amount="{:.2f}".format(contact_data['Donation Amount']),
            value_of_item="{:.2f}".format(contact_data['Value of Item']),
            region=contact_data['Region']
        )