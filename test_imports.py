#!/usr/bin/env python3
"""
Test script to verify imports and basic functionality
"""
import sys
import os
from pathlib import Path

# Add src directory to Python path
src_path = Path(__file__).parent / 'src'
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

print("Testing imports...")

try:
    # Test config imports
    from config.settings import EMAIL_SETTINGS, DEFAULT_EMAIL, USE_OAUTH, DYNAMIC_USER_OAUTH
    print("✅ Config settings imported successfully")
    print(f"   - USE_OAUTH: {USE_OAUTH}")
    print(f"   - DYNAMIC_USER_OAUTH: {DYNAMIC_USER_OAUTH}")
except Exception as e:
    print(f"❌ Config settings import failed: {e}")

try:
    # Test mail sender import
    from utils.mail_sender import send_email_with_diagnostics
    print("✅ Mail sender imported successfully")
except Exception as e:
    print(f"❌ Mail sender import failed: {e}")

try:
    # Test OAuth manager import
    from utils.oauth_manager import get_user_credentials
    print("✅ OAuth manager imported successfully")
except Exception as e:
    print(f"❌ OAuth manager import failed: {e}")

try:
    # Test streamlit-quill import
    from streamlit_quill import st_quill
    print("✅ Streamlit-quill imported successfully")
except Exception as e:
    print(f"❌ Streamlit-quill import failed: {e}")

print("\nAll imports tested.")
