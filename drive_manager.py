import os
from config import GOOGLE_DRIVE_AVAILABLE, SERVICE_ACCOUNT_FILE

if GOOGLE_DRIVE_AVAILABLE:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

class GoogleDriveManager:
    '''Quản lý Google Drive với Service Account - Hỗ trợ cả Shared Drives'''
    
    def __init__(self):
        self.service = None
        self.authenticated = False
        self.service_email = None
    
    def authenticate(self):
        '''Xác thực với Google Drive qua Service Account'''
        if not GOOGLE_DRIVE_AVAILABLE:
            return False, "Chưa cài đặt thư viện Google Drive"
        
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            return False, f"Không tìm thấy {SERVICE_ACCOUNT_FILE}\n\nHướng dẫn:\n1. Tạo Service Account tại console.cloud.google.com\n2. Download JSON key\n3. Đổi tên thành 'service_account.json'\n4. Copy vào folder project"
        
        try:
            # Đọc credentials từ service account file
            # QUAN TRỌNG: Dùng scope .drive.readonly để có quyền đọc TẤT CẢ drives
            SCOPES = [
                'https://www.googleapis.com/auth/drive.readonly',
                'https://www.googleapis.com/auth/drive.metadata.readonly'
            ]
            
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            
            self.service = build('drive', 'v3', credentials=creds)
            self.authenticated = True
            
            # Lấy email của service account
            self.service_email = self.get_service_account_email()
            
            # Test connection
            self.service.files().list(pageSize=1).execute()
            
            return True, f"✅ Kết nối thành công\n📧 Service Account: {self.service_email}"
            
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
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
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
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            print(f"Lỗi list files: {e}")
            return []
    
    def get_shared_drives(self):
        '''Lấy danh sách Shared Drives mà Service Account có quyền truy cập'''
        if not self.authenticated:
            return []
        
        try:
            # QUAN TRỌNG: Thêm useDomainAdminAccess=False để lấy ĐÚNG drives mà service account có quyền
            results = self.service.drives().list(
                pageSize=100,
                fields="drives(id, name, capabilities)",
                useDomainAdminAccess=False
            ).execute()
            
            drives = results.get('drives', [])
            
            # Debug: In ra chi tiết
            print(f"\n🔍 DEBUG: Tìm thấy {len(drives)} Shared Drives:")
            for drive in drives:
                print(f"  - {drive['name']} (ID: {drive['id']})")
                if 'capabilities' in drive:
                    caps = drive['capabilities']
                    print(f"    Quyền: canAddChildren={caps.get('canAddChildren')}, canListChildren={caps.get('canListChildren')}")
            
            return drives
        except Exception as e:
            print(f"Lỗi get shared drives: {e}")
            return []
    
    def get_shared_folders(self):
        '''Lấy danh sách folder được share - BAO GỒM CẢ SHARED DRIVES'''
        if not self.authenticated:
            return []
        
        all_folders = []
        
        # 1. LẤY SHARED DRIVES (thư mục dùng chung) - ƯU TIÊN TRƯỚC
        try:
            shared_drives = self.get_shared_drives()
            
            # Chuyển Shared Drives thành format folder
            for drive in shared_drives:
                all_folders.append({
                    'id': drive['id'],
                    'name': f"📁 {drive['name']} (Shared Drive)",
                    'source': 'Shared Drive',
                    'driveId': drive['id']  # Lưu driveId để dùng sau
                })
            
            print(f"✓ Tìm thấy {len(shared_drives)} Shared Drives")
            
        except Exception as e:
            print(f"Lỗi get shared drives: {e}")
        
        # 2. LẤY FOLDERS TRONG MY DRIVE (shared trực tiếp)
        try:
            query = "mimeType='application/vnd.google-apps.folder' and sharedWithMe=true and trashed=false"
            
            results = self.service.files().list(
                q=query,
                pageSize=100,
                fields="files(id, name, owners, driveId, capabilities)",
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            
            my_drive_folders = results.get('files', [])
            
            # Debug: In ra chi tiết
            print(f"\n🔍 DEBUG: Tìm thấy {len(my_drive_folders)} shared folders:")
            for folder in my_drive_folders:
                print(f"  - {folder['name']} (ID: {folder['id']})")
                if 'driveId' in folder:
                    print(f"    driveId: {folder['driveId']} (thuộc Shared Drive)")
                if 'capabilities' in folder:
                    caps = folder['capabilities']
                    print(f"    Quyền: canListChildren={caps.get('canListChildren')}")
            
            # Đánh dấu là My Drive folders
            for folder in my_drive_folders:
                folder['source'] = 'My Drive (Shared)'
            
            all_folders.extend(my_drive_folders)
            
            print(f"✓ Tìm thấy {len(my_drive_folders)} folders trong My Drive (shared)")
            
        except Exception as e:
            print(f"Lỗi get my drive folders: {e}")
        
        # 3. TÌM FOLDERS TRONG TẤT CẢ CÁC SHARED DRIVES
        # Điều này sẽ bao gồm cả folders từ Shared Drives mà service account được add vào
        try:
            print(f"\n🔍 DEBUG: Tìm kiếm folders trong tất cả Shared Drives...")
            
            query = "mimeType='application/vnd.google-apps.folder' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                corpora='allDrives',  # TÌM TRONG TẤT CẢ DRIVES
                pageSize=100,
                fields="files(id, name, driveId, owners)",
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            
            all_drive_folders = results.get('files', [])
            
            # Lọc chỉ lấy folders có driveId (tức là từ Shared Drives)
            # và chưa có trong danh sách
            existing_ids = {f['id'] for f in all_folders}
            
            for folder in all_drive_folders:
                if folder['id'] not in existing_ids and 'driveId' in folder:
                    folder['source'] = 'Shared Drive Folder'
                    all_folders.append(folder)
                    print(f"  + Tìm thấy: {folder['name']} (từ driveId: {folder['driveId']})")
            
        except Exception as e:
            print(f"Lỗi search in all drives: {e}")
        
        # 4. THÔNG BÁO NẾU KHÔNG TÌM THẤY GÌ
        if not all_folders:
            print(f"\n⚠️ KHÔNG TÌM THẤY FOLDER/DRIVE NÀO")
            print(f"📧 Service Account Email: {self.service_email}")
            print(f"\n💡 Kiểm tra:")
            print(f"1. Đã share folder My Drive với Service Account chưa?")
            print(f"2. Đã thêm Service Account vào Shared Drive chưa?")
            print(f"   - Mở Shared Drive → ⚙️ Settings → Manage members")
            print(f"   - Add email: {self.service_email}")
            print(f"   - Chọn role: Viewer hoặc Content Manager\n")
        
        return all_folders
    
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
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
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
            request = self.service.files().get_media(
                fileId=file_id,
                supportsAllDrives=True
            )
            
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
    
    def check_folder_permissions(self, folder_id):
        '''Kiểm tra quyền truy cập folder (DEBUG)'''
        if not self.authenticated:
            return None
        
        try:
            file = self.service.files().get(
                fileId=folder_id,
                fields="id, name, permissions, owners, driveId",
                supportsAllDrives=True
            ).execute()
            
            return file
        except Exception as e:
            print(f"Lỗi check permissions: {e}")
            return None
    
    def list_folders_in_shared_drive(self, drive_id):
        '''Liệt kê folders trong Shared Drive cụ thể'''
        if not self.authenticated:
            return []
        
        try:
            # QUAN TRỌNG: Không dùng "in parents" cho root của Shared Drive
            # Thay vào đó, search tất cả folders trong drive đó
            query = f"mimeType='application/vnd.google-apps.folder' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                corpora='drive',  # Chỉ tìm trong Shared Drive này
                driveId=drive_id,
                pageSize=100,
                fields="files(id, name, parents)",
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            
            folders = results.get('files', [])
            
            # Lọc chỉ lấy folders ở root level (không có parents hoặc parents là drive_id)
            root_folders = []
            for folder in folders:
                parents = folder.get('parents', [])
                # Nếu không có parents hoặc parents chứa drive_id thì là root folder
                if not parents or drive_id in parents:
                    root_folders.append(folder)
            
            return root_folders
            
        except Exception as e:
            print(f"Lỗi list folders in shared drive: {e}")
            return []
    
    def list_all_folders_in_shared_drive(self, drive_id):
        '''Liệt kê TẤT CẢ folders trong Shared Drive (bao gồm cả subfolders)'''
        if not self.authenticated:
            return []
        
        try:
            query = f"mimeType='application/vnd.google-apps.folder' and trashed=false"
            
            results = self.service.files().list(
                q=query,
                corpora='drive',
                driveId=drive_id,
                pageSize=1000,
                fields="files(id, name, parents)",
                orderBy="name",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            
            return results.get('files', [])
            
        except Exception as e:
            print(f"Lỗi list all folders: {e}")
            return []