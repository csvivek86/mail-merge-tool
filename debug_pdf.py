#!/usr/bin/env python3
"""
Debug script to check for PDF template availability
"""

import os
import sys
from pathlib import Path

def check_pdf_template():
    # Locations where we should look for the template
    user_home = Path(os.path.expanduser("~"))
    app_data_dir = user_home / "NSNA_Mail_Merge_Data"
    data_dir = app_data_dir / "data"
    
    # Create the directory if it doesn't exist
    os.makedirs(data_dir, exist_ok=True)
    
    # Check various locations where the PDF might be
    pdf_name = "NSNA Atlanta Letterhead Updated.pdf"
    
    pdf_locations = [
        Path.cwd() / pdf_name,  # Current directory
        app_data_dir / pdf_name,  # App data dir
        data_dir / pdf_name,  # Data dir
        Path.cwd() / "src" / pdf_name,  # src dir
        Path.cwd() / "dist" / "NSNA-Mail-Merge" / pdf_name,  # dist dir
        Path.cwd() / "dist" / "NSNA Mail Merge.app" / "Contents" / "MacOS" / pdf_name  # MacOS dir
    ]
    
    print(f"Checking for PDF template: {pdf_name}")
    print(f"Current working directory: {Path.cwd()}")
    
    found = False
    for location in pdf_locations:
        if location.exists():
            print(f"FOUND at: {location}")
            found = True
        else:
            print(f"Not found at: {location}")
    
    # Create directories
    print("\nCreating data directories...")
    logs_dir = app_data_dir / "logs"
    os.makedirs(logs_dir, exist_ok=True)
    print(f"Created logs dir: {logs_dir}")
    
    # Create a log file
    log_file = logs_dir / "debug_pdf_template.log"
    with open(log_file, "w") as f:
        f.write(f"PDF Template check run at {os.path.abspath(__file__)}\n")
        f.write(f"Current working directory: {Path.cwd()}\n\n")
        
        for location in pdf_locations:
            if location.exists():
                f.write(f"FOUND at: {location}\n")
            else:
                f.write(f"Not found at: {location}\n")
    
    print(f"Log file created at: {log_file}")
    
    # If not found, copy the PDF to the data directory
    if not found and (Path.cwd() / pdf_name).exists():
        print("\nCopying PDF template...")
        import shutil
        source = Path.cwd() / pdf_name
        target = data_dir / pdf_name
        shutil.copy(str(source), str(target))
        print(f"Copied from {source} to {target}")

if __name__ == "__main__":
    check_pdf_template()
