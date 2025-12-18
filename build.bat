@echo off
echo ============================================
echo  Smart PDF Extractor Pro - Build Script
echo ============================================
echo.

:: Check Python
echo [1/5] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please install Python from: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python found!
echo.

:: Install PyInstaller
echo [2/5] Installing PyInstaller...
pip install pyinstaller
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller!
    pause
    exit /b 1
)
echo [OK] PyInstaller installed!
echo.

:: Clean previous builds
echo [3/5] Cleaning previous builds...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "PDF_Extractor_Pro.exe" del /q "PDF_Extractor_Pro.exe"
echo [OK] Cleaned!
echo.

:: Build with PyInstaller
echo [4/5] Building executable (this may take 2-3 minutes)...
echo.
pyinstaller build.spec --clean
if errorlevel 1 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)
echo.
echo [OK] Build complete!
echo.

:: Move exe to root
echo [5/5] Finalizing...
if exist "dist\PDF_Extractor_Pro.exe" (
    move /y "dist\PDF_Extractor_Pro.exe" "PDF_Extractor_Pro.exe"
    echo [OK] Executable moved to root folder
) else (
    echo [ERROR] Executable not found in dist folder!
    pause
    exit /b 1
)
echo.

:: Clean up build artifacts
echo Cleaning build artifacts...
rmdir /s /q "build"
rmdir /s /q "dist"
del /q "PDF_Extractor_Pro.spec" 2>nul
echo.

echo ============================================
echo  Build Complete!
echo ============================================
echo.
echo Executable: PDF_Extractor_Pro.exe
echo.
echo To create distribution package:
echo   Run: package.bat
echo.
pause