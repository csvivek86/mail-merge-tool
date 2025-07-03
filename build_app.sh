#!/bin/bash

# NSNA Mail Merge Tool Build Script
# This script builds the executable version of the app

echo "🛠️  Building NSNA Mail Merge Tool executable..."

# Check if PyInstaller is installed
if ! python3 -c "import PyInstaller" &> /dev/null; then
    echo "❌ PyInstaller is not installed. Installing now..."
    pip3 install pyinstaller
fi

# Create directories for build artifacts
mkdir -p build dist

# Check for icon file
if [ ! -f "app_icon.icns" ]; then
    echo "⚠️  Warning: app_icon.icns not found. App will use default icon."
fi

# Run PyInstaller with our spec file
echo "🚀 Running PyInstaller..."
python3 -m PyInstaller nsna_mail_merge.spec

# Check if build succeeded
if [ $? -eq 0 ]; then
    echo "✅ Build successful!"
    
    # For macOS, show the .app location
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "📦 macOS app bundle created at: dist/NSNA Mail Merge.app"
        echo "🚀 To run: open 'dist/NSNA Mail Merge.app'"
    else
        echo "📦 Executable created at: dist/NSNA-Mail-Merge/NSNA-Mail-Merge"
        echo "🚀 To run: ./dist/NSNA-Mail-Merge/NSNA-Mail-Merge"
    fi
    
    echo ""
    echo "💡 Tip: To distribute this application, compress the entire folder:"
    echo "   - macOS: zip -r NSNA-Mail-Merge.zip 'dist/NSNA Mail Merge.app'"
    echo "   - Other: zip -r NSNA-Mail-Merge.zip dist/NSNA-Mail-Merge"
else
    echo "❌ Build failed. Check the error messages above."
fi
