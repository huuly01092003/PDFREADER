@echo off
echo ============================================
echo  Create Distribution Package
echo ============================================
echo.

:: Check if exe exists
if not exist "PDF_Extractor_Pro.exe" (
    echo [ERROR] PDF_Extractor_Pro.exe not found!
    echo Please run build.bat first!
    pause
    exit /b 1
)

:: Create distribution folder
set DIST_FOLDER=PDF_Extractor_Pro_Distribution
if exist "%DIST_FOLDER%" rmdir /s /q "%DIST_FOLDER%"
mkdir "%DIST_FOLDER%"

echo [1/4] Copying executable...
copy "PDF_Extractor_Pro.exe" "%DIST_FOLDER%\"

echo [2/4] Creating README...
(
echo Smart PDF Data Extractor Pro v3.0 - Dual Format
echo ================================================
echo.
echo QUICK START:
echo 1. Double-click PDF_Extractor_Pro.exe to run
echo 2. No installation required!
echo.
echo FEATURES:
echo - Format 1: Old PO format (14 columns)
echo - Format 2: New Order format (19 columns)
echo - Local files and Google Drive support
echo - Auto duplicate detection
echo - Debug mode for troubleshooting
echo.
echo GOOGLE DRIVE SETUP (Optional):
echo 1. Place service_account.json in the same folder
echo 2. Share your Drive folders with the service account email
echo 3. Restart the app
echo.
echo OUTPUTS:
echo - output.xlsx (Format 1)
echo - output_format2.xlsx (Format 2)
echo - app_log.txt (processing log)
echo - success_log.txt (successful files)
echo - error_log.txt (failed files)
echo.
echo SYSTEM REQUIREMENTS:
echo - Windows 10/11
echo - 4GB RAM minimum
echo - 500MB free disk space
echo.
echo TROUBLESHOOTING:
echo - If app doesn't start: Install Visual C++ Redistributable
echo - If Google Drive fails: Check service_account.json
echo - If PDF processing fails: Enable Debug Mode
echo.
echo Version: 3.0
echo Build Date: %DATE%
) > "%DIST_FOLDER%\README.txt"

echo [3/4] Creating Quick Start Guide...
(
echo QUICK START GUIDE
echo =================
echo.
echo Step 1: Run the Application
echo   - Double-click PDF_Extractor_Pro.exe
echo.
echo Step 2: Select Format
echo   - Format 1 (Old): For legacy PO documents
echo   - Format 2 (New): For new order documents
echo.
echo Step 3: Add Files
echo   - File: Select individual PDFs
echo   - Folder: Select entire folder
echo   - Drive F: Select files from Google Drive
echo   - Drive D: Select entire Drive folder
echo.
echo Step 4: Process
echo   - Click "Start Processing"
echo   - Wait for completion
echo   - Check output Excel file
echo.
echo TIPS:
echo - Enable Debug Mode to see detailed processing info
echo - Use Refresh button to update results
echo - Check Error Log if files fail
echo - Excel files are created automatically
echo.
echo GOOGLE DRIVE:
echo 1. Get service_account.json from Google Cloud Console
echo 2. Place it in the same folder as the .exe
echo 3. Share folders with service account email
echo 4. Click "Drive F" or "Drive D" to access
) > "%DIST_FOLDER%\QUICK_START.txt"

echo [4/4] Creating service_account template...
(
echo {
echo   "type": "service_account",
echo   "project_id": "YOUR_PROJECT_ID",
echo   "private_key_id": "YOUR_PRIVATE_KEY_ID",
echo   "private_key": "YOUR_PRIVATE_KEY",
echo   "client_email": "YOUR_SERVICE_ACCOUNT_EMAIL",
echo   "client_id": "YOUR_CLIENT_ID",
echo   "auth_uri": "https://accounts.google.com/o/oauth2/auth",
echo   "token_uri": "https://oauth2.googleapis.com/token",
echo   "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
echo   "client_x509_cert_url": "YOUR_CERT_URL"
echo }
) > "%DIST_FOLDER%\service_account.json.template"

echo.
echo Creating ZIP archive...
powershell -command "Compress-Archive -Path '%DIST_FOLDER%\*' -DestinationPath 'PDF_Extractor_Pro_v3.0.zip' -Force"

if exist "PDF_Extractor_Pro_v3.0.zip" (
    echo.
    echo ============================================
    echo  Package Complete!
    echo ============================================
    echo.
    echo Distribution folder: %DIST_FOLDER%
    echo ZIP file: PDF_Extractor_Pro_v3.0.zip
    echo.
    echo File size:
    dir /s "PDF_Extractor_Pro_v3.0.zip" | find "PDF_Extractor_Pro_v3.0.zip"
    echo.
    echo Ready to distribute!
) else (
    echo [ERROR] Failed to create ZIP file!
)

echo.
pause