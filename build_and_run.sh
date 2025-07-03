#!/usr/bin/env bash
# Build and run script for NSNA Mail Merge Tool

# Set working directory to script location
cd "$(dirname "$0")"

# Colors for terminal output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}===== NSNA Mail Merge Tool Build Script =====${NC}"

# Check for PyInstaller
echo -e "${BLUE}Checking for PyInstaller...${NC}"
if ! command -v pyinstaller &> /dev/null; then
    echo -e "${YELLOW}PyInstaller not found. Installing...${NC}"
    pip install pyinstaller
fi

# Build the application
echo -e "${BLUE}Building application...${NC}"
pyinstaller release_mail_merge.spec --noconfirm

# Check if build was successful
if [ $? -eq 0 ]; then
    echo -e "${GREEN}Build successful!${NC}"
    
    # Run the application
    echo -e "${BLUE}Running application...${NC}"
    open "./dist/NSNA Mail Merge.app"
    
    echo -e "${GREEN}Done! The app is now running.${NC}"
else
    echo -e "${YELLOW}Build failed. Check the error messages above.${NC}"
    exit 1
fi
