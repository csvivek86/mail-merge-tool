@echo off
REM NSNA Mail Merge Tool Build Script for Windows
REM This script builds the executable version of the app

echo 🛠️  Building NSNA Mail Merge Tool executable...

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo ❌ PyInstaller is not installed. Installing now...
    pip install pyinstaller
)

REM Create directories for build artifacts
if not exist build mkdir build
if not exist dist mkdir dist

REM Check for icon file
if not exist app_icon.ico (
    echo ⚠️  Warning: app_icon.ico not found. App will use default icon.
)

REM Run PyInstaller with our spec file
echo 🚀 Running PyInstaller...
python -m PyInstaller nsna_mail_merge.spec

REM Check if build succeeded
if %ERRORLEVEL% EQU 0 (
    echo ✅ Build successful!
    echo 📦 Executable created at: dist\NSNA-Mail-Merge\NSNA-Mail-Merge.exe
    echo 🚀 To run: double-click dist\NSNA-Mail-Merge\NSNA-Mail-Merge.exe
    echo.
    echo 💡 Tip: To distribute this application, compress the entire folder:
    echo     zip -r NSNA-Mail-Merge.zip dist\NSNA-Mail-Merge
) else (
    echo ❌ Build failed. Check the error messages above.
)

pause
