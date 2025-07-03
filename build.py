import PyInstaller.__main__
import shutil
from pathlib import Path

def build_app():
    """Build the executable with required resources"""
    # Define paths
    src_dir = Path('src')
    dist_dir = Path('dist')
    build_dir = Path('build')
    
    # Clean previous builds
    for dir_path in [dist_dir, build_dir]:
        if dir_path.exists():
            shutil.rmtree(dir_path)
    
    # PyInstaller configuration
    PyInstaller.__main__.run([
        'src/main.py',                    # Main script
        '--name=NSNA-Mail-Merge',         # App name
        '--windowed',                     # GUI mode
        '--onedir',                       # One directory bundle
        '--add-data=src/templates:templates',  # Include templates
        '--add-data=src/config:config',   # Include config
        '--clean',                        # Clean cache
        '--noconfirm',                    # Replace existing
        '--hidden-import=json'            # Ensure json module is included
    ])
    
    print("Build completed!")

if __name__ == "__main__":
    build_app()