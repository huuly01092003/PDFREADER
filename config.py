import os

# Đường dẫn
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "output.xlsx")

# SERVICE ACCOUNT - Không cần token.pickle và credentials.json nữa
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'service_account.json')

# Log files
LOG_FILE = os.path.join(BASE_DIR, "app_log.txt")
SUCCESS_FILE = os.path.join(BASE_DIR, "success_log.txt")
ERROR_FILE = os.path.join(BASE_DIR, "error_log.txt")

# Tesseract OCR (optional)
TESSERACT_PATH = os.path.join(BASE_DIR, "tesseract.exe")
TESSDATA_PATH = os.path.join(BASE_DIR, "tessdata")

# Cấu hình OCR
if os.path.exists(TESSERACT_PATH):
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

# Kiểm tra Google Drive API
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    GOOGLE_DRIVE_AVAILABLE = True
except ImportError:
    GOOGLE_DRIVE_AVAILABLE = False