#!/usr/bin/env python3
"""
Cleanup script to remove unnecessary files from the NSNA Mail Merge project
This script will safely remove duplicate and temporary files that are no longer needed
"""

import os
import shutil
from pathlib import Path
import sys

# Files to keep (essential files)
essential_files = [
    'bootstrap.py',
    'build_app.sh',
    'build_app.bat',
    'entitlements.plist',
    'README.md',
    'release_mail_merge.spec',
    'requirements.txt',
    'run_app.py',
    'src',
    'config',
    'NSNA Atlanta Letterhead Updated.pdf',
    'NSNA Atlanta Letterhead Updated.docx',
    'create_app_icon.py',
    '.git',
    '.gitignore',
    '.vscode'
]

# Files and directories to remove (redundant or temporary files)
files_to_remove = [
    # Duplicate and old spec files
    'debug_mail_merge.spec',
    'simple_mail_merge.spec',
    'mail_merge.spec',
    'fixed_mail_merge.spec',
    'nsna_mail_merge.spec',
    'NSNA-Mail-Merge.spec',
    'debug_build.spec',
    
    # Patch files that have been applied
    'app_patch.py',
    'app_patch_imports.py',
    'app_settings_patch.py',
    'fix_lineheight.py',
    'fix_lineheight_final.py',
    
    # Debug and temporary files
    'debug_bootstrap.py',
    'crash.log',
    'README.md.new',
    'test_implementation.py',
    'launch.sh',
    'run_debug.py',
    'setup_oauth.py',
    'fix_os_import.py'
]

def cleanup():
    """Clean up the project directory"""
    project_root = Path(__file__).parent
    
    print(f"Cleaning up project directory: {project_root}")
    print("The following files will be removed:")
    
    # Check each file to confirm it exists before deleting
    for file_name in files_to_remove:
        file_path = project_root / file_name
        if file_path.exists():
            print(f"  - {file_name}")
        else:
            print(f"  - {file_name} (not found, skipping)")
    
    # Ask for confirmation before proceeding
    confirm = input("\nDo you want to proceed with removal? (y/n): ")
    if confirm.lower() != 'y':
        print("Cleanup cancelled.")
        return
    
    # Remove the files
    print("\nRemoving files...")
    for file_name in files_to_remove:
        file_path = project_root / file_name
        try:
            if file_path.exists():
                if file_path.is_dir():
                    shutil.rmtree(file_path)
                    print(f"  Removed directory: {file_name}")
                else:
                    os.remove(file_path)
                    print(f"  Removed file: {file_name}")
        except Exception as e:
            print(f"  Error removing {file_name}: {e}")
    
    print("\nCleanup completed.")
    print("The following essential files were preserved:")
    for file_name in essential_files:
        file_path = project_root / file_name
        if file_path.exists():
            print(f"  - {file_name}")

if __name__ == "__main__":
    cleanup()
