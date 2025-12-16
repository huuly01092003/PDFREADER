import os
from config import GOOGLE_DRIVE_AVAILABLE, SERVICE_ACCOUNT_FILE

if GOOGLE_DRIVE_AVAILABLE:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

class GoogleDriveManager:
    '''Quản lý Google Drive với Service Account'''
    
    def __init__(self):
        self.service = None
        self.authenticated = False
    
    def authenticate(self):
        '''Xác thực với Google Drive qua Service Account'''
        if not GOOGLE_DRIVE_AVAILABLE:
            return False, "Chưa cài đặt thư viện Google Drive"
        
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            return False, f"Không tìm thấy {SERVICE_ACCOUNT_FILE}\n\nHướng dẫn:\n1. Tạo Service Account tại console.cloud.google.com\n2. Download JSON key\n3. Đổi tên thành 'service_account.json'\n4. Copy vào folder project"
        
        try:
            # Đọc credentials từ service account file
            SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
            
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            
            self.service = build('drive', 'v3', credentials=creds)
            self.authenticated = True
            
            # Test connection
            self.service.files().list(pageSize=1).execute()
            
            return True, "✅ Kết nối thành công (Service Account)"
            
        except Exception as e:
            return False, f"Lỗi xác thực: {str(e)}\n\nKiểm tra:\n- File service_account.json có đúng không?\n- Đã enable Google Drive API chưa?\n- Đã share folder với service account chưa?"
    
    def list_folders(self, parent_id='root'):
        '''Liệt kê thư mục'''
        if not self.authenticated:
            return []
        
        try:
            query = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                pageSize=100,
                fields="files(id, name)",
                orderBy="name"
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            print(f"Lỗi list folders: {e}")
            return []
    
    def list_pdf_files(self, folder_id='root'):
        '''Liệt kê PDF'''
        if not self.authenticated:
            return []
        
        try:
            query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                pageSize=1000,
                fields="files(id, name, size)",
                orderBy="name"
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            print(f"Lỗi list files: {e}")
            return []
    
    def get_shared_folders(self):
        '''Lấy danh sách folder được share với service account'''
        if not self.authenticated:
            return []
        
        try:
            # Tìm tất cả folder được share
            query = "mimeType='application/vnd.google-apps.folder' and sharedWithMe=true and trashed=false"
            
            results = self.service.files().list(
                q=query,
                pageSize=100,
                fields="files(id, name, owners)",
                orderBy="name"
            ).execute()
            
            folders = results.get('files', [])
            return folders
        except Exception as e:
            print(f"Lỗi get shared folders: {e}")
            return []
    
    def search_files_in_folder(self, folder_id, query_text=''):
        '''Tìm kiếm PDF trong folder'''
        if not self.authenticated:
            return []
        
        try:
            if query_text:
                query = f"'{folder_id}' in parents and name contains '{query_text}' and mimeType='application/pdf' and trashed=false"
            else:
                query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                pageSize=100,
                fields="files(id, name, size)",
                orderBy="name"
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            print(f"Lỗi search files: {e}")
            return []
    
    def download_file(self, file_id, destination_path):
        '''Download file từ Drive'''
        if not self.authenticated:
            return False
        
        try:
            request = self.service.files().get_media(fileId=file_id)
            
            with open(destination_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
            
            return True
        except Exception as e:
            print(f"Lỗi download: {e}")
            return False
    
    def search_files(self, query_text, folder_id='root'):
        '''Tìm kiếm file (backward compatibility)'''
        return self.search_files_in_folder(folder_id, query_text)
    
    def get_service_account_email(self):
        '''Lấy email của service account'''
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            return None
        
        try:
            import json
            with open(SERVICE_ACCOUNT_FILE, 'r') as f:
                data = json.load(f)
                return data.get('client_email')
        except:
            return None