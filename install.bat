@echo off
echo ============================================
echo  Smart PDF Extractor - Auto Installer
echo ============================================
echo.

:: Check Python
echo [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please install Python from: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python found!
echo.

:: Install requirements
echo [2/4] Installing requirements...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install requirements!
    pause
    exit /b 1
)
echo [OK] Requirements installed!
echo.

:: Check Tesseract (optional)
echo [3/4] Checking Tesseract OCR (optional)...
if exist "tesseract.exe" (
    echo [OK] Tesseract found in project folder
) else (
    echo [SKIP] Tesseract not found - OCR will not work
    echo You can download from: https://github.com/UB-Mannheim/tesseract/wiki
)
echo.

:: Check Google Drive credentials (optional)
echo [4/4] Checking Google Drive setup (optional)...
if exist "credentials.json" (
    echo [OK] credentials.json found - Google Drive ready
) else (
    echo [SKIP] credentials.json not found - Google Drive disabled
    echo To enable: See README.md section "Setup Google Drive"
)
echo.

echo ============================================
echo  Installation Complete!
echo ============================================
echo.
echo To run the app:
echo   python main.py
echo.
echo Or double-click: run.bat
echo.
pause