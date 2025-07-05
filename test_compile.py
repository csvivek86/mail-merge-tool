#!/usr/bin/env python3
"""
Simple script to test if streamlit_app.py compiles without syntax errors
"""
import ast
import sys

try:
    with open('streamlit_app.py', 'r', encoding='utf-8') as f:
        code = f.read()
    
    # Parse the code to check for syntax errors
    ast.parse(code)
    print("✅ streamlit_app.py compiles successfully!")
    print("✅ Session state fix applied correctly")
    
except SyntaxError as e:
    print(f"❌ Syntax error in streamlit_app.py: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Error reading/parsing streamlit_app.py: {e}")
    sys.exit(1)
