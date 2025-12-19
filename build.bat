@echo off
chcp 65001 >nul
echo ============================================
echo  PDF Extractor Pro - OPTIMIZED BUILD
echo ============================================
echo.

:: Kiểm tra Python
echo [1/6] Kiểm tra Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [LỖI] Chưa cài Python!
    echo Tải tại: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python found
echo.

:: Cài PyInstaller
echo [2/6] Cài PyInstaller...
pip install pyinstaller --upgrade --quiet
if errorlevel 1 (
    echo [LỖI] Không cài được PyInstaller!
    pause
    exit /b 1
)
echo [OK] PyInstaller ready
echo.

:: Dọn dẹp
echo [3/6] Dọn dẹp build cũ...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
if exist "PDF_Extractor_Pro.exe" del /q "PDF_Extractor_Pro.exe"
echo [OK] Cleaned
echo.

:: Build với PyInstaller
echo [4/6] Building executable...
echo      (Có thể mất 3-5 phút)
echo.

pyinstaller --clean ^
    --onefile ^
    --windowed ^
    --optimize=2 ^
    --name=PDF_Extractor_Pro ^
    --icon=icon.ico ^
    --version-file=version_info.txt ^
    --upx-dir=upx ^
    --hidden-import=pdfplumber ^
    --hidden-import=pytesseract ^
    --hidden-import=openpyxl ^
    --hidden-import=google.auth ^
    --hidden-import=google.oauth2.service_account ^
    --hidden-import=googleapiclient.discovery ^
    --exclude-module=matplotlib ^
    --exclude-module=numpy ^
    --exclude-module=pandas ^
    --exclude-module=pytest ^
    --exclude-module=IPython ^
    main.py

if errorlevel 1 (
    echo [LỖI] Build failed!
    pause
    exit /b 1
)
echo.
echo [OK] Build complete
echo.

:: Di chuyển exe
echo [5/6] Hoàn tất...
if exist "dist\PDF_Extractor_Pro.exe" (
    move /y "dist\PDF_Extractor_Pro.exe" "PDF_Extractor_Pro.exe"
    echo [OK] Đã di chuyển exe
) else (
    echo [LỖI] Không tìm thấy exe!
    pause
    exit /b 1
)
echo.

:: Dọn dẹp thư mục build
echo [6/6] Dọn dẹp...
rmdir /s /q "build"
rmdir /s /q "dist"
del /q "*.spec" 2>nul
echo.

:: Hiển thị kết quả
echo ============================================
echo  BUILD HOÀN TẤT!
echo ============================================
echo.
echo File: PDF_Extractor_Pro.exe
for %%A in (PDF_Extractor_Pro.exe) do echo Size: %%~zA bytes (%%~zK KB)
echo.
echo Bước tiếp theo:
echo   1. Test: PDF_Extractor_Pro.exe
echo   2. Package: package_final.bat
echo.
pause