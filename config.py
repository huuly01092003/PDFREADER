"""
Optimized config.py - Lazy imports for faster startup
"""
import os

# Đường dẫn
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "output.xlsx")

# SERVICE ACCOUNT
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'service_account.json')

# Log files
LOG_FILE = os.path.join(BASE_DIR, "app_log.txt")
SUCCESS_FILE = os.path.join(BASE_DIR, "success_log.txt")
ERROR_FILE = os.path.join(BASE_DIR, "error_log.txt")

# Tesseract OCR (optional) - LAZY LOAD
TESSERACT_PATH = os.path.join(BASE_DIR, "tesseract.exe")
TESSDATA_PATH = os.path.join(BASE_DIR, "tessdata")

# LAZY IMPORT - Chỉ import khi dùng Google Drive
GOOGLE_DRIVE_AVAILABLE = None

def check_google_drive():
    """Lazy check Google Drive availability"""
    global GOOGLE_DRIVE_AVAILABLE
    if GOOGLE_DRIVE_AVAILABLE is None:
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            GOOGLE_DRIVE_AVAILABLE = True
        except ImportError:
            GOOGLE_DRIVE_AVAILABLE = False
    return GOOGLE_DRIVE_AVAILABLE

def init_tesseract():
    """Lazy init Tesseract - chỉ khi cần OCR"""
    if os.path.exists(TESSERACT_PATH):
        try:
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
            os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH
        except ImportError:
            pass