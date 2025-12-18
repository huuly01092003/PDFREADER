import os
from datetime import datetime
from config import LOG_FILE, SUCCESS_FILE, ERROR_FILE

def write_log(message, log_type="info"):
    """Ghi log vào file app_log.txt"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{log_type.upper()}] {message}\n"
    
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Không thể ghi log: {e}")

def write_success(filename):
    """Ghi file thành công vào success_log.txt"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Kiểm tra xem file đã tồn tại trong log chưa
    if os.path.exists(SUCCESS_FILE):
        with open(SUCCESS_FILE, 'r', encoding='utf-8') as f:
            existing_logs = f.read()
            # Kiểm tra nếu file name đã có trong log (tránh duplicate)
            if f"] {filename}\n" in existing_logs:
                return False  # Skip nếu đã tồn tại
    
    log_entry = f"[{timestamp}] {filename}\n"
    
    try:
        with open(SUCCESS_FILE, 'a', encoding='utf-8') as f:
            f.write(log_entry)
        return True
    except Exception as e:
        print(f"Không thể ghi success log: {e}")
        return False

def is_file_processed(filename):
    """Kiểm tra file đã được xử lý thành công chưa"""
    if not os.path.exists(SUCCESS_FILE):
        return False
    
    try:
        with open(SUCCESS_FILE, 'r', encoding='utf-8') as f:
            existing_logs = f.read()
            return f"] {filename}\n" in existing_logs
    except:
        return False

def write_error(filename, error_message=None):
    """
    Ghi file lỗi vào error_log.txt
    Format: [timestamp] filename - error_message
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Kiểm tra xem file đã tồn tại trong log chưa
    if os.path.exists(ERROR_FILE):
        with open(ERROR_FILE, 'r', encoding='utf-8') as f:
            existing_logs = f.read()
            # Kiểm tra nếu file name đã có trong log (tránh duplicate)
            if f"] {filename} -" in existing_logs or f"] {filename}\n" in existing_logs:
                return  # Skip nếu đã tồn tại
    
    # GHI TIMESTAMP, FILENAME VÀ ERROR MESSAGE
    if error_message:
        log_entry = f"[{timestamp}] {filename} - {error_message}\n"
    else:
        log_entry = f"[{timestamp}] {filename} - Lỗi không xác định\n"
    
    try:
        with open(ERROR_FILE, 'a', encoding='utf-8') as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Không thể ghi error log: {e}")

def read_log_file(log_type="app"):
    """Đọc nội dung file log"""
    log_files = {
        "app": LOG_FILE,
        "success": SUCCESS_FILE,
        "error": ERROR_FILE
    }
    
    file_path = log_files.get(log_type, LOG_FILE)
    
    if not os.path.exists(file_path):
        return f"File log chưa tồn tại: {os.path.basename(file_path)}"
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            if not content.strip():
                return "File log trống"
            return content
    except Exception as e:
        return f"Lỗi đọc file: {e}"

def clear_log_file(log_type="app"):
    """Xóa nội dung file log"""
    log_files = {
        "app": LOG_FILE,
        "success": SUCCESS_FILE,
        "error": ERROR_FILE
    }
    
    file_path = log_files.get(log_type, LOG_FILE)
    
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("")
        return True
    except Exception as e:
        print(f"Không thể xóa log: {e}")
        return False

def init_log_files():
    """Khởi tạo các file log nếu chưa tồn tại"""
    for log_file in [LOG_FILE, SUCCESS_FILE, ERROR_FILE]:
        if not os.path.exists(log_file):
            try:
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write(f"# Log file created at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            except Exception as e:
                print(f"Không thể tạo file log: {e}")

def get_error_count():
    """Đếm số file bị lỗi"""
    if not os.path.exists(ERROR_FILE):
        return 0
    
    try:
        with open(ERROR_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            # Đếm các dòng có format [timestamp] filename
            count = sum(1 for line in lines if line.strip() and line.startswith('['))
            return count
    except:
        return 0

def get_success_count():
    """Đếm số file thành công"""
    if not os.path.exists(SUCCESS_FILE):
        return 0
    
    try:
        with open(SUCCESS_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            count = sum(1 for line in lines if line.strip() and line.startswith('['))
            return count
    except:
        return 0