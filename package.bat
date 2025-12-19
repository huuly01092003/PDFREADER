@echo off
chcp 65001 >nul
echo ============================================
echo  TẠO GÓI PHÂN PHỐI
echo ============================================
echo.

:: Kiểm tra exe
if not exist "PDF_Extractor_Pro.exe" (
    echo [LỖI] Không tìm thấy PDF_Extractor_Pro.exe
    echo Chạy build_fast.bat trước!
    pause
    exit /b 1
)

:: Tạo thư mục phân phối
set DIST=PDF_Extractor_Pro_v3.0_Release
if exist "%DIST%" rmdir /s /q "%DIST%"
mkdir "%DIST%"

echo [1/4] Sao chép file thực thi...
copy "PDF_Extractor_Pro.exe" "%DIST%\"
echo [OK] Copied exe
echo.

echo [2/4] Tạo README...
(
echo ========================================
echo  PDF EXTRACTOR PRO v3.0
echo  HƯỚNG DẪN SỬ DỤNG
echo ========================================
echo.
echo KHỞI ĐỘNG:
echo   • Double-click: PDF_Extractor_Pro.exe
echo   • Không cần cài đặt!
echo.
echo TÍNH NĂNG:
echo   ✓ 2 Format PDF (Old PO + New Order^)
echo   ✓ Xử lý từ máy tính hoặc Google Drive
echo   ✓ Tự động loại bỏ trùng lặp
echo   ✓ Debug mode để khắc phục lỗi
echo   ✓ Log chi tiết (success/error^)
echo.
echo CÀI ĐẶT GOOGLE DRIVE (tùy chọn^):
echo   1. Tạo Service Account tại:
echo      console.cloud.google.com
echo   2. Tải file JSON key
echo   3. Đổi tên: service_account.json
echo   4. Copy vào cùng thư mục với .exe
echo   5. Share folder Drive với email trong JSON
echo.
echo OUTPUT:
echo   • output.xlsx (Format 1^)
echo   • output_format2.xlsx (Format 2^)
echo   • app_log.txt
echo   • success_log.txt
echo   • error_log.txt
echo.
echo YÊU CẦU HỆ THỐNG:
echo   • Windows 10/11 (64-bit^)
echo   • 4GB RAM
echo   • 100MB ổ cứng trống
echo.
echo KHẮ C PHỤC SỰ CỐ:
echo   • App không chạy?
echo     → Cài Visual C++ Redistributable
echo     → https://aka.ms/vs/17/release/vc_redist.x64.exe
echo.
echo   • Không đọc được PDF?
echo     → Bật Debug Mode
echo     → Kiểm tra Error Log
echo.
echo   • Google Drive lỗi?
echo     → Kiểm tra service_account.json
echo     → Đảm bảo đã share folder
echo.
echo LIÊN HỆ HỖ TRỢ:
echo   Email: support@example.com
echo.
echo Version: 3.0
echo Build: %DATE% %TIME%
) > "%DIST%\README.txt"
echo [OK] Created README
echo.

echo [3/4] Tạo Quick Start Guide...
(
echo ========================================
echo  QUICK START - 3 BƯỚC
echo ========================================
echo.
echo BƯỚC 1: CHỌN FORMAT
echo   □ Format 1: PO cũ (14 cột^)
echo   □ Format 2: Order mới (19 cột^)
echo.
echo BƯỚC 2: THÊM FILE
echo   □ File: Chọn từng file PDF
echo   □ Folder: Chọn cả thư mục
echo   □ Drive F: Chọn file từ Drive
echo   □ Drive D: Chọn folder từ Drive
echo.
echo BƯỚC 3: XỬ LÝ
echo   □ Click "Bắt đầu xử lý"
echo   □ Đợi hoàn tất
echo   □ Mở Excel xem kết quả
echo.
echo ========================================
echo  MẸO VẶT
echo ========================================
echo.
echo • Debug Mode: Xem chi tiết quá trình
echo • Refresh: Cập nhật kết quả
echo • Clear: Xóa dữ liệu Excel
echo • Log Viewer: Xem lịch sử xử lý
echo.
echo • File trùng sẽ tự động bỏ qua
echo • Excel tự động tạo nếu chưa có
echo • Log lưu vĩnh viễn để tra cứu
echo.
echo ========================================
echo  GOOGLE DRIVE SETUP
echo ========================================
echo.
echo 1. Vào: console.cloud.google.com
echo 2. Tạo project mới
echo 3. Enable Google Drive API
echo 4. Tạo Service Account
echo 5. Download JSON key
echo 6. Đổi tên: service_account.json
echo 7. Share Drive folder với email trong JSON
echo 8. Chạy lại app
echo.
echo Xong! Giờ có thể dùng Drive F và Drive D
) > "%DIST%\QUICK_START.txt"
echo [OK] Created Quick Start
echo.

echo [4/4] Tạo template Service Account...
(
echo {
echo   "type": "service_account",
echo   "project_id": "YOUR_PROJECT_ID",
echo   "private_key_id": "YOUR_KEY_ID",
echo   "private_key": "-----BEGIN PRIVATE KEY-----\nYOUR_PRIVATE_KEY\n-----END PRIVATE KEY-----\n",
echo   "client_email": "your-service-account@project.iam.gserviceaccount.com",
echo   "client_id": "YOUR_CLIENT_ID",
echo   "auth_uri": "https://accounts.google.com/o/oauth2/auth",
echo   "token_uri": "https://oauth2.googleapis.com/token",
echo   "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs"
echo }
) > "%DIST%\service_account.json.TEMPLATE"
echo [OK] Created template
echo.

:: Tạo ZIP
echo Đang tạo file ZIP...
powershell -command "Compress-Archive -Path '%DIST%\*' -DestinationPath 'PDF_Extractor_Pro_v3.0.zip' -Force"

if exist "PDF_Extractor_Pro_v3.0.zip" (
    echo.
    echo ============================================
    echo  HOÀN TẤT!
    echo ============================================
    echo.
    echo Thư mục: %DIST%
    echo File ZIP: PDF_Extractor_Pro_v3.0.zip
    echo.
    for %%A in (PDF_Extractor_Pro_v3.0.zip) do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1024/1024
    )
    echo Kích thước: !sizeMB! MB
    echo.
    echo SẴN SÀNG PHÂN PHỐI!
    echo.
    echo Gửi file ZIP cho người dùng
    echo Họ chỉ cần giải nén và chạy .exe
) else (
    echo [LỖI] Không tạo được ZIP!
)

echo.
pause