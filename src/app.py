import sys
import os
import subprocess
import logging
import json
import re
import pandas as pd
from PyQt6.QtCore import (
    Qt, 
    QTimer
)
from PyQt6.QtWidgets import (
    QApplication, 
    QMainWindow, 
    QPushButton, 
    QLabel, 
    QVBoxLayout, 
    QHBoxLayout, 
    QWidget, 
    QFileDialog, 
    QProgressBar, 
    QMessageBox,
    QFrame,  
    QTabWidget,
    QLineEdit,
    QComboBox,
    QCheckBox,
    QFormLayout,
    QGroupBox,
    QGridLayout,
    QScrollArea,
    QTabBar,
    QMenuBar,
    QStatusBar,
    QDialog,
    QDialogButtonBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QSpinBox,
    QPlainTextEdit,
    QTextEdit,
    QToolBar,
    QSplitter,
    QColorDialog
)
from PyQt6.QtGui import QFont, QIcon, QTextCharFormat, QColor, QTextCursor, QAction, QTextBlockFormat, QTextListFormat
import sys
import json  # Add this import
import logging
from pathlib import Path
from config.settings import EMAIL_SETTINGS
from datetime import datetime
from config.template_settings import TemplateSettings
from config.email_template_settings import EmailTemplateSettings
from config.app_settings import AppSettings
from utils.excel_reader import read_excel
from utils.mail_sender import send_email
from utils.document_generator import DocumentGenerator  # Only import DocumentGenerator
from utils.excel_template import create_excel_template
from utils.styles import MAIN_STYLE
import pandas as pd  # Import pandas for Excel validation
import subprocess

# Disable CATransaction warning on macOS
if sys.platform == "darwin":  # macOS specific
    import os
    os.environ["QT_MAC_WANTS_LAYER"] = "1"

class RichTextTemplateEditor(QWidget):
    """Rich text editor for PDF template with variable insertion"""
    
    def __init__(self, template_settings, parent=None):
        super().__init__(parent)
        self.template_settings = template_settings
        self.setup_ui()
        self.load_template()
    
    def setup_ui(self):
        """Setup the rich text editor UI"""
        layout = QVBoxLayout(self)
        
        # Formatting toolbar
        toolbar = QToolBar()
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        
        # Font formatting actions
        self.bold_action = QAction("Bold", self)
        self.bold_action.setCheckable(True)
        self.bold_action.triggered.connect(self.toggle_bold)
        toolbar.addAction(self.bold_action)
        
        self.italic_action = QAction("Italic", self)
        self.italic_action.setCheckable(True)
        self.italic_action.triggered.connect(self.toggle_italic)
        toolbar.addAction(self.italic_action)
        
        self.underline_action = QAction("Underline", self)
        self.underline_action.setCheckable(True)
        self.underline_action.triggered.connect(self.toggle_underline)
        toolbar.addAction(self.underline_action)
        
        toolbar.addSeparator()
        
        # Font size
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32'])
        self.font_size_combo.setCurrentText('12')
        self.font_size_combo.currentTextChanged.connect(self.change_font_size)
        toolbar.addWidget(QLabel("Size:"))
        toolbar.addWidget(self.font_size_combo)
        
        toolbar.addSeparator()
        
        # Font color
        self.color_button = QPushButton("Color")
        self.color_button.clicked.connect(self.change_color)
        toolbar.addWidget(self.color_button)
        
        toolbar.addSeparator()
        
        # Line spacing
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(['Single', '1.5', 'Double', '2.5', 'Triple'])
        self.line_spacing_combo.setCurrentText('Single')
        self.line_spacing_combo.currentTextChanged.connect(self.change_line_spacing)
        toolbar.addWidget(QLabel("Line Spacing:"))
        toolbar.addWidget(self.line_spacing_combo)
        
        toolbar.addSeparator()
        
        # List formatting
        self.bullet_action = QAction("• Bullets", self)
        self.bullet_action.triggered.connect(self.toggle_bullet_list)
        toolbar.addAction(self.bullet_action)
        
        self.number_action = QAction("1. Numbers", self)
        self.number_action.triggered.connect(self.toggle_numbered_list)
        toolbar.addAction(self.number_action)
        
        toolbar.addSeparator()
        
        # Variable insertion
        self.variable_combo = QComboBox()
        self.variable_combo.addItems([
            "Select Variable...",
            "{{First Name}}",
            "{{Last Name}}",
            "{{Email}}",
            "{{Donation Amount}}",
            "{{Date}}",
            "{{Current Year}}"
        ])
        self.variable_combo.currentTextChanged.connect(self.insert_variable)
        toolbar.addWidget(QLabel("Variables:"))
        toolbar.addWidget(self.variable_combo)
        
        layout.addWidget(toolbar)
        
        # Main editor layout
        editor_layout = QHBoxLayout()
        
        # Rich text editor
        self.text_editor = QTextEdit()
        self.text_editor.setMinimumHeight(400)
        self.text_editor.cursorPositionChanged.connect(self.update_format_buttons)
        editor_layout.addWidget(self.text_editor, stretch=2)
        
        # Preview pane
        preview_layout = QVBoxLayout()
        preview_label = QLabel("Live Preview:")
        preview_label.setStyleSheet("font-weight: bold;")
        preview_layout.addWidget(preview_label)
        
        self.preview_area = QTextEdit()
        self.preview_area.setReadOnly(True)
        self.preview_area.setMaximumWidth(300)
        self.preview_area.setMinimumHeight(400)
        preview_layout.addWidget(self.preview_area)
        
        # Sample data for preview
        self.sample_data = {
            'First Name': 'John',
            'Last Name': 'Doe',
            'Email': 'john.doe@example.com',
            'Donation Amount': '500.00',
            'Date': datetime.now().strftime('%B %d, %Y'),
            'Current Year': str(datetime.now().year)
        }
        
        preview_layout.addWidget(QLabel("Sample Data:"))
        sample_text = QTextEdit()
        sample_text.setReadOnly(True)
        sample_text.setMaximumHeight(100)
        sample_text.setPlainText("\n".join([f"{k}: {v}" for k, v in self.sample_data.items()]))
        preview_layout.addWidget(sample_text)
        
        editor_layout.addLayout(preview_layout)
        layout.addLayout(editor_layout)
        
        # Connect text change to preview update
        self.text_editor.textChanged.connect(self.update_preview)
    
    def toggle_bold(self):
        """Toggle bold formatting"""
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Weight.Bold if self.bold_action.isChecked() else QFont.Weight.Normal)
        self.merge_format_on_selection(fmt)
    
    def toggle_italic(self):
        """Toggle italic formatting"""
        fmt = QTextCharFormat()
        fmt.setFontItalic(self.italic_action.isChecked())
        self.merge_format_on_selection(fmt)
    
    def toggle_underline(self):
        """Toggle underline formatting"""
        fmt = QTextCharFormat()
        fmt.setFontUnderline(self.underline_action.isChecked())
        self.merge_format_on_selection(fmt)
    
    def change_font_size(self, size):
        """Change font size"""
        try:
            fmt = QTextCharFormat()
            fmt.setFontPointSize(int(size))
            self.merge_format_on_selection(fmt)
        except ValueError:
            pass
    
    def change_color(self):
        """Change text color"""
        color = QColorDialog.getColor(Qt.GlobalColor.black, self, "Select Color")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            self.merge_format_on_selection(fmt)
    
    def change_line_spacing(self, spacing_text):
        """Change line spacing"""
        cursor = self.text_editor.textCursor()
        block_format = QTextBlockFormat()
        
        spacing_map = {
            'Single': 1.0,
            '1.5': 1.5,
            'Double': 2.0,
            '2.5': 2.5,
            'Triple': 3.0
        }
        
        if spacing_text in spacing_map:
            block_format.setLineHeight(spacing_map[spacing_text] * 100, 1)
            cursor.mergeBlockFormat(block_format)
    
    def toggle_bullet_list(self):
        """Toggle bullet list formatting"""
        cursor = self.text_editor.textCursor()
        list_format = QTextListFormat()
        
        current_list = cursor.currentList()
        if current_list and current_list.format().style() == QTextListFormat.Style.ListDisc:
            # Remove list formatting
            block_format = QTextBlockFormat()
            cursor.mergeBlockFormat(block_format)
            current_list.removeItem(cursor.blockNumber())
        else:
            # Add bullet list formatting
            list_format.setStyle(QTextListFormat.Style.ListDisc)
            list_format.setIndent(1)
            cursor.createList(list_format)
    
    def toggle_numbered_list(self):
        """Toggle numbered list formatting"""
        cursor = self.text_editor.textCursor()
        list_format = QTextListFormat()
        
        current_list = cursor.currentList()
        if current_list and current_list.format().style() == QTextListFormat.Style.ListDecimal:
            # Remove list formatting
            block_format = QTextBlockFormat()
            cursor.mergeBlockFormat(block_format)
            current_list.removeItem(cursor.blockNumber())
        else:
            # Add numbered list formatting
            list_format.setStyle(QTextListFormat.Style.ListDecimal)
            list_format.setIndent(1)
            cursor.createList(list_format)

    def merge_format_on_selection(self, fmt):
        """Apply formatting to selected text"""
        cursor = self.text_editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.text_editor.mergeCurrentCharFormat(fmt)
    
    def insert_variable(self, variable):
        """Insert a variable at cursor position"""
        if variable and variable != "Select Variable...":
            cursor = self.text_editor.textCursor()
            cursor.insertText(variable)
            self.variable_combo.setCurrentIndex(0)  # Reset selection
    
    def update_format_buttons(self):
        """Update toolbar buttons based on current cursor position"""
        cursor = self.text_editor.textCursor()
        fmt = cursor.charFormat()
        
        self.bold_action.setChecked(fmt.fontWeight() == QFont.Weight.Bold)
        self.italic_action.setChecked(fmt.fontItalic())
        self.underline_action.setChecked(fmt.fontUnderline())
        
        size = fmt.fontPointSize()
        if size > 0:
            self.font_size_combo.setCurrentText(str(int(size)))
    
    def update_preview(self):
        """Update the preview pane with processed template"""
        try:
            html_content = self.text_editor.toHtml()
            # Replace variables with sample data
            processed_content = html_content
            for var, value in self.sample_data.items():
                processed_content = processed_content.replace(f"{{{{{var}}}}}", str(value))
            
            self.preview_area.setHtml(processed_content)
        except Exception as e:
            self.preview_area.setPlainText(f"Preview error: {str(e)}")
    
    def load_template(self):
        """Load existing template into editor"""
        try:
            # First try to load saved rich template
            config_dir = Path(__file__).parent / 'config'
            rich_settings_path = config_dir / 'rich_template_settings.json'
            
            if rich_settings_path.exists():
                with open(rich_settings_path) as f:
                    rich_settings = json.load(f)
                    if rich_settings.get('html_template'):
                        self.text_editor.setHtml(rich_settings['html_template'])
                        logging.info("Loaded saved rich template")
                        return
            
            # If no saved template, create default template with formatting
            template_parts = [
                f"<p><strong>{{{{Date}}}}</strong></p>",
                f"<p><strong>Dear {{{{First Name}}}} {{{{Last Name}}}},</strong></p>",
                "<p></p>"  # Empty line
            ]
            
            # Add thank you lines
            for line in self.template_settings.thank_you_lines:
                template_parts.append(f"<p>{line}</p>")
            
            template_parts.append("<p></p>")  # Empty line
            template_parts.append(f"<p>{self.template_settings.receipt_statement}</p>")
            template_parts.append("<p></p>")  # Empty line
            
            # Add donation details
            template_parts.append(f"<p><strong>Donation Amount: ${{{{Donation Amount}}}}</strong></p>")
            template_parts.append(f"<p><strong>Donation(s) Year: {{{{Current Year}}}}</strong></p>")
            template_parts.append("<p></p>")  # Empty line
            
            template_parts.append(f"<p>{self.template_settings.disclaimer}</p>")
            template_parts.append("<p></p>")  # Empty line
            
            # Add org info
            for line in self.template_settings.org_info:
                template_parts.append(f"<p>{line}</p>")
            
            template_parts.append("<p></p>")  # Empty line
            
            # Add closing lines
            for line in self.template_settings.closing_lines:
                template_parts.append(f"<p>{line}</p>")
            
            template_parts.append("<p></p>")  # Empty line
            
            # Add signature
            for line in self.template_settings.signature:
                template_parts.append(f"<p>{line}</p>")
            
            html_template = "".join(template_parts)
            self.text_editor.setHtml(html_template)
            
        except Exception as e:
            logging.error(f"Failed to load template: {str(e)}")
            self.text_editor.setPlainText("Error loading template")
    
    def get_template_content(self):
        """Get the current template content as HTML"""
        return self.text_editor.toHtml()
    
    def get_template_variables(self):
        """Extract variables from template content"""
        content = self.text_editor.toPlainText()
        variables = re.findall(r'\{\{([^}]+)\}\}', content)
        return list(set(variables))

class RichTextEmailEditor(QWidget):
    """Rich text editor for Email template with variable insertion and preview"""
    def __init__(self, email_settings, parent=None):
        super().__init__(parent)
        self.email_settings = email_settings
        self.setup_ui()
        self.load_template()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        toolbar = QToolBar()

        # Bold
        self.bold_action = QAction("B", self)
        self.bold_action.setCheckable(True)
        self.bold_action.triggered.connect(self.toggle_bold)
        toolbar.addAction(self.bold_action)
        # Italic
        self.italic_action = QAction("I", self)
        self.italic_action.setCheckable(True)
        self.italic_action.triggered.connect(self.toggle_italic)
        toolbar.addAction(self.italic_action)
        # Underline
        self.underline_action = QAction("U", self)
        self.underline_action.setCheckable(True)
        self.underline_action.triggered.connect(self.toggle_underline)
        toolbar.addAction(self.underline_action)
        toolbar.addSeparator()
        # Font size
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32'])
        self.font_size_combo.setCurrentText('12')
        self.font_size_combo.currentTextChanged.connect(self.change_font_size)
        toolbar.addWidget(QLabel("Size:"))
        toolbar.addWidget(self.font_size_combo)
        toolbar.addSeparator()
        
        # Font color
        self.color_button = QPushButton("Color")
        self.color_button.clicked.connect(self.change_color)
        toolbar.addWidget(self.color_button)
        
        toolbar.addSeparator()
        
        # Line spacing
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(['Single', '1.5', 'Double', '2.5', 'Triple'])
        self.line_spacing_combo.setCurrentText('Single')
        self.line_spacing_combo.currentTextChanged.connect(self.change_line_spacing)
        toolbar.addWidget(QLabel("Line Spacing:"))
        toolbar.addWidget(self.line_spacing_combo)
        
        toolbar.addSeparator()
        
        # List formatting
        self.bullet_action = QAction("• Bullets", self)
        self.bullet_action.triggered.connect(self.toggle_bullet_list)
        toolbar.addAction(self.bullet_action)
        
        self.number_action = QAction("1. Numbers", self)
        self.number_action.triggered.connect(self.toggle_numbered_list)
        toolbar.addAction(self.number_action)
        
        toolbar.addSeparator()
        
        # Variable insertion
        self.variable_combo = QComboBox()
        self.variable_combo.addItems([
            "Select Variable...",
            "{First Name}",
            "{Last Name}",
            "{Email}",
            "{Donation Amount}",
            "{Date}",
            "{Current Year}"
        ])
        self.variable_combo.currentTextChanged.connect(self.insert_variable)
        toolbar.addWidget(QLabel("Variables:"))
        toolbar.addWidget(self.variable_combo)
        layout.addWidget(toolbar)
        # Rich text editor
        self.text_editor = QTextEdit()
        self.text_editor.setMinimumHeight(300)
        self.text_editor.cursorPositionChanged.connect(self.update_format_buttons)
        layout.addWidget(self.text_editor)
        # Preview
        preview_label = QLabel("Live Preview:")
        preview_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(preview_label)
        self.preview_area = QTextEdit()
        self.preview_area.setReadOnly(True)
        self.preview_area.setMaximumHeight(120)
        layout.addWidget(self.preview_area)
        self.text_editor.textChanged.connect(self.update_preview)

    def toggle_bold(self):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Weight.Bold if self.bold_action.isChecked() else QFont.Weight.Normal)
        self.merge_format_on_selection(fmt)
    def toggle_italic(self):
        fmt = QTextCharFormat()
        fmt.setFontItalic(self.italic_action.isChecked())
        self.merge_format_on_selection(fmt)
    def toggle_underline(self):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(self.underline_action.isChecked())
        self.merge_format_on_selection(fmt)
    def change_font_size(self, size):
        try:
            fmt = QTextCharFormat()
            fmt.setFontPointSize(float(size))
            self.merge_format_on_selection(fmt)
        except ValueError:
            pass
    def merge_format_on_selection(self, fmt):
        cursor = self.text_editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.text_editor.mergeCurrentCharFormat(fmt)
    def insert_variable(self, variable):
        if variable and variable != "Select Variable...":
            cursor = self.text_editor.textCursor()
            cursor.insertText(variable)
            self.variable_combo.setCurrentIndex(0)
    def update_format_buttons(self):
        cursor = self.text_editor.textCursor()
        fmt = cursor.charFormat()
        self.bold_action.setChecked(fmt.fontWeight() == QFont.Weight.Bold)
        self.italic_action.setChecked(fmt.fontItalic())
        self.underline_action.setChecked(fmt.fontUnderline())
        size = fmt.fontPointSize()
        if size > 0:
            self.font_size_combo.setCurrentText(str(int(size)))
    def update_preview(self):
        try:
            html = self.text_editor.toHtml()
            # Replace variables with sample data
            sample = {
                'First Name': 'John',
                'Last Name': 'Doe',
                'Email': 'john.doe@example.com',
                'Donation Amount': '500.00',
                'Date': datetime.now().strftime('%B %d, %Y'),
                'Current Year': str(datetime.now().year)
            }
            for k, v in sample.items():
                html = html.replace(f'{{{k}}}', v)
            self.preview_area.setHtml(html)
        except Exception:
            self.preview_area.setPlainText("Preview error")
    def load_template(self):
        try:
            if self.email_settings and hasattr(self.email_settings, 'body_html') and self.email_settings.body_html:
                self.text_editor.setHtml(self.email_settings.body_html)
            elif self.email_settings and self.email_settings.body:
                html = "<br>".join(self.email_settings.body)
                self.text_editor.setHtml(html)
        except Exception:
            self.text_editor.setPlainText("")
    def get_template_content(self):
        return self.text_editor.toHtml()
    def get_template_variables(self):
        content = self.text_editor.toPlainText()
        variables = re.findall(r'\{([^}]+)\}', content)
        return list(set(variables))
    
    def change_color(self):
        """Change text color"""
        color = QColorDialog.getColor(Qt.GlobalColor.black, self, "Select Color")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            self.merge_format_on_selection(fmt)
    
    def change_line_spacing(self, spacing_text):
        """Change line spacing"""
        cursor = self.text_editor.textCursor()
        block_format = QTextBlockFormat()
        
        spacing_map = {
            'Single': 1.0,
            '1.5': 1.5,
            'Double': 2.0,
            '2.5': 2.5,
            'Triple': 3.0
        }
        
        if spacing_text in spacing_map:
            block_format.setLineHeight(spacing_map[spacing_text] * 100, 1)
            cursor.mergeBlockFormat(block_format)
    
    def toggle_bullet_list(self):
        """Toggle bullet list formatting"""
        cursor = self.text_editor.textCursor()
        list_format = QTextListFormat()
        
        current_list = cursor.currentList()
        if current_list and current_list.format().style() == QTextListFormat.Style.ListDisc:
            # Remove list formatting
            block_format = QTextBlockFormat()
            cursor.mergeBlockFormat(block_format)
            current_list.removeItem(cursor.blockNumber())
        else:
            # Add bullet list formatting
            list_format.setStyle(QTextListFormat.Style.ListDisc)
            list_format.setIndent(1)
            cursor.createList(list_format)
    
    def toggle_numbered_list(self):
        """Toggle numbered list formatting"""
        cursor = self.text_editor.textCursor()
        list_format = QTextListFormat()
        
        current_list = cursor.currentList()
        if current_list and current_list.format().style() == QTextListFormat.Style.ListDecimal:
            # Remove list formatting
            block_format = QTextBlockFormat()
            cursor.mergeBlockFormat(block_format)
            current_list.removeItem(cursor.blockNumber())
        else:
            # Add numbered list formatting
            list_format.setStyle(QTextListFormat.Style.ListDecimal)
            list_format.setIndent(1)
            cursor.createList(list_format)

def setup_logging():
    """Set up logging configuration"""
    log_dir = Path.cwd() / 'logs'
    log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / 'app.log'
    logging.basicConfig(
        filename=str(log_file),
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # Also log to console for debugging
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    logging.getLogger().addHandler(console_handler)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        try:
            # Import EMAIL_SETTINGS at initialization
            from config.settings import EMAIL_SETTINGS
            self.smtp_settings = EMAIL_SETTINGS  # Store SMTP settings

            # Set window properties
            self.setWindowTitle("NSNA Mail Merge Tool")
            self.setMinimumSize(900, 700)
            self.setStyleSheet(MAIN_STYLE)
            
            # Initialize attributes
            self.from_email = None
            self.excel_path = None
            self.template_path = None
            
            # Load settings
            self.app_settings = AppSettings()
            self.receipts_dir = self.app_settings.get_receipts_dir()
            self.receipts_dir.mkdir(parents=True, exist_ok=True)
            
            # Load email and template settings
            self.email_settings = self._load_email_settings()
            self.template_settings = self._load_template_settings()
            
            # Set default PDF template path
            self.template_path = self.app_settings.get_template_path()
            if not self.template_path.exists():
                logging.warning(f"Template not found at {self.template_path}")
                self.template_path = None
            else:
                logging.info(f"PDF template loaded: {self.template_path}")
            
            self.init_ui()
            
        except Exception as e:
            logging.critical(f"Failed to initialize UI: {str(e)}")
            raise

    def _load_email_settings(self) -> EmailTemplateSettings:
        """Load email template settings from file"""
        try:
            config_dir = Path(__file__).parent / 'config'
            settings_path = config_dir / 'email_template_settings.json'
            
            if settings_path.exists():
                with open(settings_path) as f:
                    settings = json.load(f)
                    return EmailTemplateSettings(**settings)
            else:
                # Return default settings
                return self._get_default_email_settings()
        except Exception as e:
            logging.error(f"Failed to load email settings: {str(e)}")
            return self._get_default_email_settings()

    def _load_template_settings(self) -> TemplateSettings:
        """Load PDF template settings from file"""
        try:
            config_dir = Path(__file__).parent / 'config'
            settings_path = config_dir / 'template_settings.json'
            
            if settings_path.exists():
                with open(settings_path) as f:
                    settings = json.load(f)
                    return TemplateSettings(**settings)
            else:
                # Return default settings
                return self._get_default_template_settings()
        except Exception as e:
            logging.error(f"Failed to load template settings: {str(e)}")
            return self._get_default_template_settings()

    def _get_default_email_settings(self) -> EmailTemplateSettings:
        """Return default email template settings"""
        return EmailTemplateSettings(
            subject="NSNA Donation Receipt",
            greeting="Dear",
            body=[
                "Thank you for your generous donation to NSNA Atlanta.",
                "Please find attached your donation receipt for tax purposes.",
            ],
            signature=[
                "Best regards,",
                "NSNA Atlanta"
            ]
        )

    def _get_default_template_settings(self) -> TemplateSettings:
        """Return default PDF template settings"""
        return TemplateSettings(
            greeting="Dear",
            thank_you_lines=[
                "On behalf of Nagarathar Sangam of North America, thank you for your recent",
                "donation to our organization."
            ],
            receipt_statement="This letter will serve as a tax receipt for your contribution listed below.",
            donation_year=str(datetime.now().year),
            disclaimer="No goods or services were provided in exchange for your contribution.",
            org_info=[
                "Nagarathar Sangam of North America is a registered Section 501(c)(3) non-",
                "profit organization (EIN #: 22-3974176). For questions regarding this",
                "acknowledgment letter, please contact us at treasurer@achi.org."
            ],
            closing_lines=[
                "We truly appreciate your donation and look forward to your continued support",
                "of our mission."
            ],
            signature=[
                "Sincerely,",
                "",
                "Treasurer, NSNA",
                "(2025-2026 term)"
            ]
        )
        
    def init_ui(self):
        """Initialize the UI"""
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Set up the UI components
        self.setup_ui(layout)
        
        # Show the window
        self.show()

    def setup_ui(self, layout):
        """Set up main UI components"""
        # Create tab widget
        tabs = QTabWidget()
        
        # Create tab layouts
        email_settings_layout = QVBoxLayout()
        template_layout = QVBoxLayout()
        pdf_layout = QVBoxLayout()
        
        # Create tab widgets
        email_settings_widget = QWidget()
        template_widget = QWidget()
        pdf_widget = QWidget()
        
        # Set layouts for widgets
        email_settings_widget.setLayout(email_settings_layout)
        template_widget.setLayout(template_layout)
        pdf_widget.setLayout(pdf_layout)
        
        # Add widgets to tabs
        tabs.addTab(email_settings_widget, "Email Settings")
        tabs.addTab(template_widget, "Email Template")
        tabs.addTab(pdf_widget, "PDF Template")
        
        # Setup individual tabs
        self.setup_email_settings_tab(email_settings_layout)
        self.setup_email_template_tab(template_layout)
        self.setup_pdf_template_tab(pdf_layout)
        
        # Add tabs to main layout
        layout.addWidget(tabs)

    def setup_email_settings_tab(self, layout):
        """Set up email settings tab"""
        self.add_email_section(layout)
        self.add_excel_section(layout)
        self.add_summary_section(layout)
        self.add_receipt_location_section(layout)  # Make sure this line is
        layout.addStretch()

    def setup_email_template_tab(self, layout):
        """Set up email template tab with subject and rich text editor"""
        form = QFormLayout()
        # Subject
        self.subject_edit = QPlainTextEdit()
        self.subject_edit.setPlaceholderText("Enter email subject")
        self.subject_edit.setMaximumHeight(50)
        self.subject_edit.setPlainText(self.email_settings.subject)
        form.addRow("Subject:", self.subject_edit)
        # Rich text editor for email body
        self.rich_email_editor = RichTextEmailEditor(self.email_settings)
        form.addRow("Body:", self.rich_email_editor)
        # Save button
        save_btn = QPushButton("Save Template")
        save_btn.clicked.connect(self.save_email_template)
        layout.addLayout(form)
        layout.addWidget(save_btn)
        layout.addStretch()

    def setup_pdf_template_tab(self, layout):
        """Set up PDF template customization tab with rich text editor only"""
        tab_widget = QTabWidget()
        # Rich Text Editor Tab
        rich_editor_widget = QWidget()
        rich_editor_layout = QVBoxLayout(rich_editor_widget)
        self.rich_template_editor = RichTextTemplateEditor(self.template_settings)
        rich_editor_layout.addWidget(self.rich_template_editor)
        # Save button for rich editor
        rich_save_btn = QPushButton("Save Rich Template")
        rich_save_btn.clicked.connect(self.save_rich_template)
        rich_editor_layout.addWidget(rich_save_btn)
        tab_widget.addTab(rich_editor_widget, "Rich Text Editor")
        layout.addWidget(tab_widget)
    
    def create_traditional_form_layout(self):
        """Create the traditional form layout"""
        main_layout = QVBoxLayout()
        form = QHBoxLayout()
        
        # Left pane
        left_pane = QVBoxLayout()
        
        # Greeting section
        greeting_group = QGroupBox("Greeting")
        greeting_layout = QFormLayout()
        self.greeting_edit = QLineEdit()
        self.greeting_edit.setText(self.template_settings.greeting)
        greeting_layout.addRow("Greeting:", self.greeting_edit)
        greeting_group.setLayout(greeting_layout)
        left_pane.addWidget(greeting_group)
        
        # Thank you section with more height
        thank_you_group = QGroupBox("Thank You Message")
        thank_you_layout = QVBoxLayout()
        self.thank_you_edit = QPlainTextEdit()
        self.thank_you_edit.setPlainText("\n".join(self.template_settings.thank_you_lines))
        self.thank_you_edit.setMinimumHeight(150)
        thank_you_layout.addWidget(self.thank_you_edit)
        thank_you_group.setLayout(thank_you_layout)
        left_pane.addWidget(thank_you_group)
        
        # Receipt statement
        receipt_group = QGroupBox("Receipt Statement")
        receipt_layout = QVBoxLayout()
        self.receipt_edit = QPlainTextEdit()
        self.receipt_edit.setPlainText(self.template_settings.receipt_statement)
        self.receipt_edit.setMinimumHeight(60)
        receipt_layout.addWidget(self.receipt_edit)
        receipt_group.setLayout(receipt_layout)
        left_pane.addWidget(receipt_group)
        
        # Right pane
        right_pane = QVBoxLayout()
        
        # Organization info with more height
        org_group = QGroupBox("Organization Information")
        org_layout = QVBoxLayout()
        self.org_info_edit = QPlainTextEdit()
        self.org_info_edit.setPlainText("\n".join(self.template_settings.org_info))
        self.org_info_edit.setMinimumHeight(100)
        org_layout.addWidget(self.org_info_edit)
        org_group.setLayout(org_layout)
        right_pane.addWidget(org_group)
        
        # Closing lines with more height
        closing_group = QGroupBox("Closing Lines")
        closing_layout = QVBoxLayout()
        self.closing_edit = QPlainTextEdit()
        self.closing_edit.setPlainText("\n".join(self.template_settings.closing_lines))
        self.closing_edit.setMinimumHeight(80)
        closing_layout.addWidget(self.closing_edit)
        closing_group.setLayout(closing_layout)
        right_pane.addWidget(closing_group)
        
        # Signature section
        signature_group = QGroupBox("Signature")
        signature_layout = QVBoxLayout()
        self.signature_edit = QPlainTextEdit()
        self.signature_edit.setPlainText("\n".join(self.template_settings.signature))
        self.signature_edit.setMinimumHeight(80)
        signature_layout.addWidget(self.signature_edit)
        signature_group.setLayout(signature_layout)
        right_pane.addWidget(signature_group)
        
        # Add donation year spinbox to right pane
        year_group = QGroupBox("Donation Year")
        year_layout = QHBoxLayout()
        self.year_edit = QSpinBox()
        self.year_edit.setRange(2000, 2100)
        self.year_edit.setValue(int(self.template_settings.donation_year))
        year_layout.addWidget(self.year_edit)
        year_group.setLayout(year_layout)
        right_pane.addWidget(year_group)
        
        # Donation content fields
        donation_group = QGroupBox("Donation Content")
        donation_layout = QFormLayout()
        
        self.amount_label_edit = QLineEdit()
        self.amount_label_edit.setText(self.template_settings.donation_content["amount_label"])
        donation_layout.addRow("Amount Label:", self.amount_label_edit)
        
        self.year_label_edit = QLineEdit()
        self.year_label_edit.setText(self.template_settings.donation_content["year_label"])
        donation_layout.addRow("Year Label:", self.year_label_edit)
        
        self.disclaimer_edit = QPlainTextEdit()
        self.disclaimer_edit.setPlainText(self.template_settings.donation_content["disclaimer"])
        self.disclaimer_edit.setMaximumHeight(50)
        donation_layout.addRow("Disclaimer:", self.disclaimer_edit)
        
        donation_group.setLayout(donation_layout)
        right_pane.addWidget(donation_group)
        
        # Add both panes to form
        form.addLayout(left_pane, stretch=1)
        form.addLayout(right_pane, stretch=1)
        
        # Save button at bottom
        save_btn = QPushButton("Save Template")
        save_btn.clicked.connect(self.save_template_settings)
        
        main_layout.addLayout(form)
        main_layout.addWidget(save_btn)
        main_layout.addStretch()
        
        return main_layout

    def add_email_section(self, layout):
        """Add email configuration widgets"""
        section = QFrame()
        section.setStyleSheet("QFrame {background-color: white; border-radius: 5px;}")
        section_layout = QVBoxLayout(section)
    
        # Email input
        email_layout = QHBoxLayout()
        email_label = QLabel("From Email:")
        self.email_input = QLineEdit()
        saved_email = self.app_settings.get_from_email()
        self.email_input.setText(saved_email)
        self.email_input.setPlaceholderText("donations@nsna.org")
        email_layout.addWidget(email_label)
        email_layout.addWidget(self.email_input)
        section_layout.addLayout(email_layout)
    
        layout.addWidget(section)

    def add_excel_section(self, layout):
        """Add Excel file selection widgets"""
        section = QFrame()
        section_layout = QVBoxLayout(section)
    
        # Excel file buttons
        file_layout = QHBoxLayout()
    
        self.select_button = QPushButton("Select Excel File")
        self.select_button.setObjectName("select_button")  # Add object name
        self.select_button.clicked.connect(self.select_file)
    
        self.create_template_button = QPushButton("Create New Excel")
        self.create_template_button.setObjectName("create_button")  # Add object name
        self.create_template_button.clicked.connect(self.create_excel_template)
    
        file_layout.addWidget(self.select_button)
        file_layout.addWidget(self.create_template_button)
    
        # Only unsent checkbox
        self.unsent_only_checkbox = QCheckBox("Only Send to Unsent")
        self.unsent_only_checkbox.setChecked(True)
        file_layout.addWidget(self.unsent_only_checkbox)
    
        section_layout.addLayout(file_layout)
    
        # Add test and send buttons
        button_layout = QHBoxLayout()
    
        # Test button - Green
        self.test_button = QPushButton("Send Test Email")
        self.test_button.setObjectName("test_button")  # Verify object name is set
        self.test_button.clicked.connect(self.send_test_email)
        self.test_button.setEnabled(False)
    
        # Send all button - Yellow
        self.send_button = QPushButton("Send All Emails")
        self.send_button.setObjectName("send_button")  # Verify object name is set
        self.send_button.clicked.connect(self.send_all_emails)
        self.send_button.setEnabled(False)
    
        button_layout.addWidget(self.test_button)
        button_layout.addWidget(self.send_button)
        section_layout.addLayout(button_layout)
    
        layout.addWidget(section)

    def add_summary_section(self, layout):
        """Add statistics summary widgets"""
        section = QFrame()
        section.setStyleSheet("QFrame {background-color: white; border-radius: 5px;}")
        section_layout = QVBoxLayout(section)
    
        # Add stats layout
        stats_layout = QHBoxLayout()  # Added this line to define stats_layout
    
        # Stats labels with object names for styling
        self.total_label = QLabel("Total: 0")
        self.total_label.setObjectName("total_label")
        
        self.sent_label = QLabel("✓ Sent: 0")
        self.sent_label.setObjectName("sent_label")
        
        self.failed_label = QLabel("✗ Failed: 0")
        self.failed_label.setObjectName("failed_label")
        
        self.skipped_label = QLabel("⇢ Skipped: 0")
        self.skipped_label.setObjectName("skipped_label")

        for label in [self.total_label, self.sent_label, self.failed_label, self.skipped_label]:
            label.setStyleSheet("""
                QLabel {
                    padding: 10px;
                    border: 2px solid #ccc;
                    border-radius: 8px;
                    background-color: #f8f9fa;
                    font-size: 14px;
                    font-weight: bold;
                    qproperty-alignment: AlignCenter;
                }
            """)
            stats_layout.addWidget(label)

        section_layout.addLayout(stats_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        section_layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel()
        self.status_label.setStyleSheet("color: #666;")
        section_layout.addWidget(self.status_label)

        # Error label
        self.error_label = QLabel()
        self.error_label.setStyleSheet("color: red;")
        self.error_label.setVisible(False)
        section_layout.addWidget(self.error_label)

        layout.addWidget(section)

    def add_receipt_location_section(self, layout):
        """Add receipt location configuration widgets"""
        section = QFrame()
        section.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton {
                padding: 5px 10px;
                border: 2px solid #3b82f6;
                border-radius: 4px;
                background-color: #3b82f6;
                color: white;
            }
            QPushButton:hover {
                background-color: #2563eb;
                border-color: #1d4ed8;
            }
        """)
        section_layout = QVBoxLayout(section)
    
        # Location label
        self.location_label = QLabel(f"Receipt Location: {self.receipts_dir}")
        section_layout.addWidget(self.location_label)
    
        # Buttons
        button_layout = QHBoxLayout()
        self.change_location_btn = QPushButton("Change Location")
        self.reset_location_btn = QPushButton("Reset to Default")
    
        self.change_location_btn.clicked.connect(self.change_receipt_location)
        self.reset_location_btn.clicked.connect(self.reset_receipt_location)
    
        button_layout.addWidget(self.change_location_btn)
        button_layout.addWidget(self.reset_location_btn)
        section_layout.addLayout(button_layout)
    
        layout.addWidget(section)

    def select_file(self):
        """Handle Excel file selection"""
        # Get current email from input field
        from_email = self.email_input.text().strip()
        if not from_email:
            QMessageBox.warning(self, "Input Required", "Please enter a From Email address")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            str(Path.home()),
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            try:
                # Validate selected file
                if not self.validate_excel_file(file_path):
                    return
                
                # Update UI and store path
                self.excel_path = Path(file_path)
                self.from_email = from_email
                self.app_settings.set_from_email(from_email)
                
                # Update UI
                self.select_button.setText("Excel File Selected")
                self.select_button.setStyleSheet("background-color: lightgreen")
                self.test_button.setEnabled(True)
                self.send_button.setEnabled(True)
                
                # Read and display current settings
                self.read_excel_settings(self.excel_path)
                
            except Exception as e:
                logging.error(f"Failed to select Excel file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to select Excel file: {str(e)}")

    def validate_excel_file(self, file_path: Path) -> bool:
        """Validate Excel file format and required columns"""
        try:
            contacts = read_excel(file_path)
            required_columns = {'First Name', 'Last Name', 'Email', 'Donation Amount'}
            
            if not contacts:
                QMessageBox.warning(self, "Invalid File", "Excel file is empty")
                return False
                
            columns = set(contacts[0].keys())
            missing = required_columns - columns
            
            if missing:
                QMessageBox.warning(
                    self,
                    "Invalid Format",
                    f"Excel file is missing required columns: {', '.join(missing)}"
                )
                return False
                
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to validate Excel file: {str(e)}")
            return False

    def send_test_email(self):
        """Send test email using first record"""
        try:
            # Get current email from input field
            from_email = self.email_input.text().strip()
            
            # Validate inputs
            if not hasattr(self, 'excel_path') or not self.excel_path:
                QMessageBox.warning(self, "Missing File", "Please select an Excel file first")
                return
                
            if not from_email:
                QMessageBox.warning(self, "Missing Email", "Please enter a From Email address")
                return
                
            if not hasattr(self, 'template_path') or not self.template_path:
                QMessageBox.warning(self, "Missing Template", "Please select a PDF template")
                return

            # Update stored email and settings
            self.from_email = from_email
            self.app_settings.set_from_email(from_email)

            # Update UI state
            self.test_button.setEnabled(False)
            self.send_button.setEnabled(False)
            self.status_label.setText("Sending test email...")
            self.progress_bar.setVisible(True)
            
            # Process test email
            self.process_emails(test_mode=True)
            
        except Exception as e:
            logging.error(f"Failed to send test email: {str(e)}")
            self.show_error(f"Test email failed: {str(e)}")
            self.test_button.setEnabled(True)
            self.send_button.setEnabled(True)

    def send_all_emails(self):
        """Send emails to all recipients"""
        try:
            # Get current email from input field
            from_email = self.email_input.text().strip()
            
            # Validate inputs
            if not hasattr(self, 'excel_path') or not self.excel_path:
                QMessageBox.warning(self, "Missing File", "Please select an Excel file first")
                return
                
            if not from_email:
                QMessageBox.warning(self, "Missing Email", "Please enter a From Email address")
                return
                
            if not hasattr(self, 'template_path') or not self.template_path:
                QMessageBox.warning(self, "Missing Template", "Please select a PDF template")
                return

            # Update stored email and settings
            self.from_email = from_email
            self.app_settings.set_from_email(from_email)

            # Confirm sending
            reply = QMessageBox.question(
                self,
                "Confirm Send",
                "Are you sure you want to send emails to all recipients?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                return

            # Update UI state
            self.test_button.setEnabled(False)
            self.send_button.setEnabled(False)
            self.status_label.setText("Processing emails...")
            self.progress_bar.setVisible(True)
            
            # Process all emails
            self.process_emails(test_mode=False)
            
        except Exception as e:
            logging.error(f"Failed to start email sending: {str(e)}")
            self.show_error(f"Failed to start email sending: {str(e)}")
            self.test_button.setEnabled(True)
            self.send_button.setEnabled(True)

    def show_success_message(self, message: str):
        """Show success message with green styling"""
        self.status_label.setStyleSheet("color: #059669; font-weight: bold; padding: 8px;")
        self.status_label.setText(f"✓ {message}")

    def show_error(self, error_msg: str):
        """Display error message in UI with red styling"""
        self.error_label.setStyleSheet("""
            color: #dc2626; 
            background-color: #fee2e2;
            padding: 8px;
            border-radius: 4px;
            font-weight: bold;
        """)
        self.error_label.setText(f"⚠️ {error_msg}")
        self.error_label.setVisible(True)

    def update_stats(self, stats):
        """Update statistics display"""
        self.total_label.setText(f"Total: {stats.get('total', 0)}")
        self.sent_label.setText(f"✓ Sent: {stats.get('sent', 0)}")
        self.failed_label.setText(f"✗ Failed: {stats.get('failed', 0)}")
        if 'skipped' in stats:
            self.skipped_label.setText(f"⇢ Skipped: {stats['skipped']}")
        # Add PDF generation stats if available
        if 'pdf_generated' in stats:
            self.status_label.setText(f"PDFs Generated: {stats['pdf_generated']}")
        if 'pdf_failed' in stats and stats['pdf_failed'] > 0:
            logging.warning(f"PDF generation failed for {stats['pdf_failed']} contacts")

    def read_excel_settings(self, excel_path):
        """Display Excel file information"""
        try:
            contacts = read_excel(excel_path)
            if contacts:
                total = len(contacts)
                sent = sum(1 for c in contacts if c.get('Email Status') == 'Sent')
                self.update_stats({
                    'total': total,
                    'sent': sent,
                    'failed': 0,
                    'skipped': 0
                })
            else:
                QMessageBox.information(self, "Excel Info", "The selected Excel file is empty.")
        except Exception as e:
            logging.error(f"Failed to read Excel file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to read Excel file: {str(e)}")

    def change_receipt_location(self):
        """Allow user to change the receipts directory"""
        new_dir = QFileDialog.getExistingDirectory(
            self,
            "Select Receipts Directory",
            str(self.receipts_dir),
            QFileDialog.Option.ShowDirsOnly
        )
        
        if new_dir:
            try:
                # Update path
                self.receipts_dir = Path(new_dir)
                self.receipts_dir.mkdir(parents=True, exist_ok=True)
                
                # Update display
                self.location_label.setText(f"Receipt Location: {self.receipts_dir}")
                
                # Save to settings
                self.app_settings.set_receipts_dir(str(self.receipts_dir))
                
                logging.info(f"Receipt location changed to: {self.receipts_dir}")
                
            except Exception as e:
                logging.error(f"Failed to change receipt location: {str(e)}")
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Failed to change receipt location: {str(e)}"
                )

    def reset_receipt_location(self):
        """Reset receipts directory to default location"""
        try:
            # Set default path
            default_dir = Path.home() / "Documents" / "NSNA Receipts"
            default_dir.mkdir(parents=True, exist_ok=True)
            
            # Update instance and settings
            self.receipts_dir = default_dir
            self.app_settings.set_receipts_dir(str(default_dir))
            
            # Update display
            self.location_label.setText(f"Receipt Location: {self.receipts_dir}")
            
            logging.info("Receipt location reset to default")
            
            QMessageBox.information(
                self,
                "Reset Complete",
                f"Receipt location has been reset to:\n{default_dir}"
            )
            
        except Exception as e:
            logging.error(f"Failed to reset receipt location: {str(e)}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to reset receipt location: {str(e)}"
            )

    def save_email_template(self):
        """Save email template settings from subject and rich editor"""
        try:
            subject = self.subject_edit.toPlainText().strip()
            html_content = self.rich_email_editor.get_template_content()
            variables = self.rich_email_editor.get_template_variables()
            if not subject or not html_content.strip():
                QMessageBox.warning(self, "Missing Input", "Subject and body are required")
                return
            settings = {
                "subject": subject,
                "body_html": html_content,
                "variables": variables
            }
            config_dir = Path(__file__).parent / 'config'
            config_dir.mkdir(exist_ok=True)
            settings_path = config_dir / 'email_template_settings.json'
            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
            self.email_settings.subject = subject
            self.email_settings.body_html = html_content
            QMessageBox.information(self, "Success", "Email template settings saved!")
            logging.info("Email template settings saved")
        except Exception as e:
            error_msg = f"Failed to save email template: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)

    def save_template_settings(self):
        """Save PDF template settings"""
        try:
            greeting = self.greeting_edit.toPlainText().strip()
            thank_you = [line.strip() for line in self.thank_you_edit.toPlainText().split('\n') if line.strip()]
            receipt = self.receipt_edit.toPlainText().strip()
            year = str(self.year_edit.value())
            org_info = [line.strip() for line in self.org_info_edit.toPlainText().split('\n') if line.strip()]
            disclaimer = self.disclaimer_edit.toPlainText().strip()
            signature = [line.strip() for line in self.signature_edit.toPlainText().split('\n') if line.strip()]

            if not all([greeting, thank_you, receipt, year, org_info, signature]):
                QMessageBox.warning(self, "Missing Input", "All fields are required")
                return

            settings = {
                "greeting": greeting,
                "thank_you_lines": thank_you,
                "receipt_statement": receipt,
                "donation_year": year,
                "disclaimer": disclaimer,
                "org_info": org_info,
                "closing_lines": [line.strip() for line in self.closing_edit.toPlainText().split('\n') if line.strip()],
                "signature": signature,
                "donation_content": {
                    "amount_label": self.amount_label_edit.text().strip(),
                    "year_label": self.year_label_edit.text().strip(),
                    "disclaimer": self.disclaimer_edit.toPlainText().strip()
                }
            }
            
            config_dir = Path(__file__).parent / 'config'
            config_dir.mkdir(exist_ok=True)
            settings_path = config_dir / 'template_settings.json'

            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=2)

            self.template_settings = TemplateSettings(**settings)
            QMessageBox.information(self, "Success", "Template settings saved!")
            logging.info("PDF template settings saved")

        except Exception as e:
            error_msg = f"Failed to save template settings: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)

    def save_rich_template(self):
        """Save rich template settings"""
        try:
            # Get HTML content from rich editor
            html_content = self.rich_template_editor.get_template_content()
            
            # Extract variables used in template
            variables = self.rich_template_editor.get_template_variables()
            
            # Save as rich template
            settings = {
                "html_template": html_content,
                "variables": variables,
                "greeting": self.template_settings.greeting,
                "thank_you_lines": self.template_settings.thank_you_lines,
                "receipt_statement": self.template_settings.receipt_statement,
                "donation_year": str(datetime.now().year),
                "disclaimer": self.template_settings.disclaimer,
                "org_info": self.template_settings.org_info,
                "closing_lines": self.template_settings.closing_lines,
                "signature": self.template_settings.signature,
                "donation_content": self.template_settings.donation_content
            }
            
            config_dir = Path(__file__).parent / 'config'
            config_dir.mkdir(exist_ok=True)
            settings_path = config_dir / 'rich_template_settings.json'

            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=2)

            QMessageBox.information(self, "Success", "Rich template settings saved!")
            logging.info("Rich template settings saved")

        except Exception as e:
            error_msg = f"Failed to save rich template: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)

    def _load_rich_template_settings(self):
        """Load rich template settings from file"""
        try:
            config_dir = Path(__file__).parent / 'config'
            settings_path = config_dir / 'rich_template_settings.json'
            
            if settings_path.exists():
                with open(settings_path) as f:
                    return json.load(f)
            return None
        except Exception as e:
            logging.error(f"Failed to load rich template settings: {str(e)}")
            return None

    def process_emails(self, test_mode=False):
        """Process all emails - generate PDFs and send emails"""
        try:
            # Read Excel data
            contacts = read_excel(self.excel_path)
            if not contacts:
                self.show_error("Failed to read Excel file or file is empty")
                return

            total_recipients = len(contacts)
            self.progress_bar.setMaximum(total_recipients)
            
            # Initialize statistics
            stats = {
                'total': total_recipients,
                'sent': 0,
                'failed': 0,
                'pdf_generated': 0,
                'pdf_failed': 0
            }

            # Get SMTP settings
            smtp_settings = EMAIL_SETTINGS.copy()
            
            # Update from email if changed
            if hasattr(self, 'from_email'):
                smtp_settings['from_email'] = self.from_email

            # Process each contact
            for index, contact in enumerate(contacts):
                try:
                    self.status_label.setText(f"Processing {contact['First Name']} {contact['Last Name']}...")
                    self.progress_bar.setValue(index)
                    QApplication.processEvents()  # Update UI

                    # Generate PDF receipt
                    pdf_path = None
                    try:
                        if hasattr(self, 'template_path') and self.template_path:
                            # Initialize document generator
                            doc_generator = DocumentGenerator(
                                self.template_path,
                                self.template_settings,
                                contact
                            )
                            
                            # Check if rich template is available
                            rich_settings = self._load_rich_template_settings()
                            if rich_settings and rich_settings.get('html_template'):
                                # Use HTML template for PDF generation
                                pdf_path = doc_generator.generate_receipt_from_html(
                                    contact,
                                    self.receipts_dir,
                                    rich_settings['html_template']
                                )
                                logging.info(f"Generated PDF from rich HTML template for {contact['First Name']} {contact['Last Name']}")
                            else:
                                # Use standard template
                                pdf_path = doc_generator.generate_receipt(
                                    contact,
                                    self.receipts_dir
                                )
                                logging.info(f"Generated PDF from standard template for {contact['First Name']} {contact['Last Name']}")
                            stats['pdf_generated'] += 1
                            
                    except Exception as pdf_error:
                        logging.error(f"PDF generation failed for {contact['First Name']} {contact['Last Name']}: {str(pdf_error)}")
                        stats['pdf_failed'] += 1
                        # Continue with email even if PDF fails
                    
                    # Generate email content
                    email_subject = self.subject_edit.toPlainText().strip() or self.email_settings.subject
                    email_body = self._format_email_body(contact)
                    
                    # Send email
                    success = send_email(
                        from_email=self.from_email,
                        to_email=contact['Email'],
                        subject=email_subject,
                        html_content=email_body,
                        attachment_path=pdf_path,
                        smtp_settings=smtp_settings,
                        is_test=test_mode
                    )
                    
                    if success:
                        stats['sent'] += 1
                    else:
                        stats['failed'] += 1
                        
                except Exception as e:
                    logging.error(f"Processing failed for {contact['First Name']} {contact['Last Name']}: {str(e)}")
                    stats['failed'] += 1
                    continue

            # Update final status
            self.progress_bar.setValue(total_recipients)
            self.update_stats(stats)
            
            # Show completion message
            if test_mode:
                self.show_success_message(f"Test completed: {stats['sent']} emails sent to your address")
            else:
                self.show_success_message(f"All emails processed: {stats['sent']} sent, {stats['failed']} failed")
                
        except Exception as e:
            logging.error(f"Email processing failed: {str(e)}")
            self.show_error(f"Email processing failed: {str(e)}")
        finally:
            # Re-enable buttons
            self.test_button.setEnabled(True)
            self.send_button.setEnabled(True)
            self.progress_bar.setVisible(False)

    def _format_email_body(self, donor_info: dict) -> str:
        """Format email body with donor information using rich editor"""
        try:
            html_content = self.rich_email_editor.get_template_content()
            # Replace variables with actual data
            for k, v in donor_info.items():
                html_content = html_content.replace(f'{{{k}}}', str(v))
            html_content = html_content.replace('{Date}', datetime.now().strftime('%B %d, %Y'))
            html_content = html_content.replace('{Current Year}', str(datetime.now().year))
            return html_content
        except Exception as e:
            logging.error(f"Error formatting email body: {str(e)}")
            return "<p>Email formatting error.</p>"

    def create_excel_template(self):
        """Create new Excel file with required headers"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Create Excel Template",
                str(Path.home() / "donations.xlsx"),
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                create_excel_template(file_path)
                
                # Set as current Excel file
                self.excel_path = Path(file_path)
                self.update_excel_status()
                
                # Enable buttons if template is valid
                if self.validate_excel_template():
                    self.test_button.setEnabled(True)
                    self.send_button.setEnabled(True)
                
                # Open Excel file for editing
                subprocess.run(['open', file_path], check=True)
                
                QMessageBox.information(
                    self,
                    "Success",
                    f"Excel template created and opened at:\n{file_path}"
                )
                logging.info(f"Created Excel template: {file_path}")
                
        except Exception as e:
            error_msg = f"Failed to create Excel template: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)

    def validate_excel_template(self) -> bool:
        """Validate Excel template has required columns
    
        Returns:
            bool: True if template is valid, False otherwise
        """
        try:
            if not self.excel_path or not self.excel_path.exists():
                return False
                
            required_columns = ['First Name', 'Last Name', 'Email', 'Donation Amount']
            df = pd.read_excel(self.excel_path)
            
            # Check if all required columns exist
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                QMessageBox.warning(
                    self,
                    "Invalid Template",
                    f"Missing required columns: {', '.join(missing_cols)}"
                )
                return False
                
            logging.info("Excel template validation successful")
            return True
            
        except Exception as e:
            error_msg = f"Excel template validation failed: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
            return False

    def update_excel_status(self):
        """Update UI to reflect Excel file status"""
        try:
            if self.excel_path and self.excel_path.exists():
                self.select_button.setText("Excel File Selected")
                self.select_button.setStyleSheet("background-color: lightgreen")
                
                # Read and display current settings
                self.read_excel_settings(self.excel_path)
                
                # Enable test and send buttons if email is set
                if self.from_email:
                    self.test_button.setEnabled(True)
                    self.send_button.setEnabled(True)
                    
                logging.info(f"Excel file status updated: {self.excel_path}")
            else:
                self.select_button.setText("Select Excel File")
                self.select_button.setStyleSheet("")
                self.test_button.setEnabled(False)
                self.send_button.setEnabled(False)
                
        except Exception as e:
            logging.error(f"Failed to update Excel status: {str(e)}")
            self.show_error("Failed to update Excel status")

    # ...existing code...
def main():
    """Main application entry point"""
    try:
        # Set up logging first
        setup_logging()
        
        # Create QApplication
        app = QApplication(sys.argv)
        
        # Set application properties
        app.setApplicationName("NSNA Mail Merge Tool")
        app.setApplicationVersion("2.0")
        app.setOrganizationName("NSNA Atlanta")
        
        # Create and show main window
        window = MainWindow()
        
        # Start the application event loop
        sys.exit(app.exec())
        
    except Exception as e:
        logging.critical(f"Failed to start application: {str(e)}")
        print(f"Critical error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
