"""
Simple launcher for the NSNA Mail Merge Tool
Handles OneDrive desktop paths and environment setup
"""

import os
import sys
import subprocess
from pathlib import Path

def find_desktop_path():
    """Find the correct desktop path, handling OneDrive scenarios"""
    
    # Standard desktop path
    standard_desktop = Path.home() / "Desktop"
    
    # OneDrive desktop paths
    onedrive_desktop1 = Path.home() / "OneDrive" / "Desktop"
    onedrive_desktop2 = Path.home() / "OneDrive - HD Supply, Inc" / "Desktop"
    
    # Check which desktop path exists and is writable
    for desktop_path in [standard_desktop, onedrive_desktop1, onedrive_desktop2]:
        if desktop_path.exists():
            try:
                # Test if we can create a folder
                test_folder = desktop_path / "test_write_permission"
                test_folder.mkdir(exist_ok=True)
                test_folder.rmdir()
                return desktop_path
            except:
                continue
    
    # Fallback to Documents if Desktop isn't accessible
    documents_path = Path.home() / "Documents"
    return documents_path

def setup_environment():
    """Set up the environment and create necessary folders"""
    
    print("🚀 Starting NSNA Mail Merge Tool Setup")
    print("=" * 50)
    
    # Find the best desktop path
    desktop_path = find_desktop_path()
    print(f"📁 Desktop path: {desktop_path}")
    
    # Create NSNA folder structure
    nsna_folder = desktop_path / "NSNA_Mail_Merge"
    nsna_folder.mkdir(exist_ok=True)
    
    subfolders = ["templates", "data", "receipts", "exports"]
    for folder in subfolders:
        (nsna_folder / folder).mkdir(exist_ok=True)
    
    print(f"✅ Created storage folders at: {nsna_folder}")
    
    # Set environment variable for the app
    os.environ['NSNA_DESKTOP_PATH'] = str(nsna_folder)
    
    return nsna_folder

def run_streamlit():
    """Run the Streamlit application"""
    
    print("\n🌐 Starting Streamlit application...")
    print("=" * 50)
    
    try:
        # Run streamlit
        cmd = [sys.executable, "-m", "streamlit", "run", "streamlit_app.py", 
               "--server.port", "8501", "--server.headless", "false"]
        
        print(f"📱 Opening web browser at: http://localhost:8501")
        print("📋 Press Ctrl+C to stop the application")
        print("")
        
        subprocess.run(cmd)
        
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")
    except Exception as e:
        print(f"❌ Error running Streamlit: {e}")
        print("\nTry running manually with:")
        print("streamlit run streamlit_app.py")

if __name__ == "__main__":
    nsna_folder = setup_environment()
    print(f"\n📂 Your files will be saved to: {nsna_folder}")
    print("🔄 Starting application...")
    
    run_streamlit()
