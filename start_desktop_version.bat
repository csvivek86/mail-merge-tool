@echo off
echo NSNA Mail Merge Tool - Desktop Storage Version
echo =============================================
echo.
echo This will start the NSNA Mail Merge Tool with desktop storage.
echo Your files will be saved to: %USERPROFILE%\Desktop\NSNA_Mail_Merge\
echo.
echo Starting Docker containers...
echo.

docker-compose up --build

echo.
echo Application stopped. Your files remain saved on your desktop.
pause
