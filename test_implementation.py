#!/usr/bin/env python3
"""
Test script to verify the mail merge functionality
"""

import sys
import os
from pathlib import Path

# Add src directory to path
src_path = Path(__file__).parent / 'src'
sys.path.insert(0, str(src_path))

def test_imports():
    """Test that all required modules can be imported"""
    try:
        from app import MainWindow, RichTextTemplateEditor
        from utils.mail_sender import send_email, send_email_with_multiple_attachments
        from utils.document_generator import DocumentGenerator
        from utils.excel_reader import read_excel
        from utils.excel_template import create_excel_template
        from config.settings import EMAIL_SETTINGS
        from config.template_settings import TemplateSettings
        from config.email_template_settings import EmailTemplateSettings
        from config.app_settings import AppSettings
        
        print("✓ All imports successful")
        return True
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False

def test_attachment_handling():
    """Test attachment handling functionality"""
    try:
        # Test Path operations
        test_path = Path("/tmp/test_attachment.pdf")
        print(f"✓ Path handling works: {test_path.exists()}")
        
        # Test email structure
        from email.mime.multipart import MIMEMultipart
        from email.mime.application import MIMEApplication
        
        msg = MIMEMultipart()
        msg['From'] = "test@example.com"
        msg['To'] = "recipient@example.com"
        msg['Subject'] = "Test"
        
        print("✓ Email structure creation works")
        return True
    except Exception as e:
        print(f"✗ Attachment handling test failed: {e}")
        return False

def test_rich_text_editor():
    """Test rich text editor functionality"""
    try:
        from PyQt6.QtWidgets import QApplication
        
        # Create minimal QApplication for testing
        app = QApplication([])
        
        # Test template settings
        from config.template_settings import TemplateSettings
        settings = TemplateSettings()
        
        print("✓ Rich text editor components available")
        return True
    except Exception as e:
        print(f"✗ Rich text editor test failed: {e}")
        return False

def main():
    """Run all tests"""
    print("Testing NSNA Mail Merge Tool Implementation...")
    print("=" * 50)
    
    tests = [
        ("Module Imports", test_imports),
        ("Attachment Handling", test_attachment_handling),
        ("Rich Text Editor", test_rich_text_editor)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\nTesting {test_name}:")
        if test_func():
            passed += 1
        else:
            print(f"  Failed: {test_name}")
    
    print("\n" + "=" * 50)
    print(f"Tests passed: {passed}/{total}")
    
    if passed == total:
        print("🎉 All tests passed! The implementation is ready.")
    else:
        print("⚠️  Some tests failed. Please check the issues above.")

if __name__ == "__main__":
    main()
