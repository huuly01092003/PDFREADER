import tkinter as tk
from tkinter import messagebox

class DriveFilePicker:
    """Dialog chọn file từ Google Drive"""
    
    def __init__(self, parent, drive_manager, log_callback):
        self.parent = parent
        self.drive_manager = drive_manager
        self.log = log_callback
        self.selected_files = []
        self.current_folder_id = 'root'
    
    def show(self):
        """Hiển thị dialog và trả về danh sách file đã chọn"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Chọn file từ Google Drive")
        self.dialog.geometry("750x600")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        self._create_ui()
        self._load_initial()
        
        self.dialog.wait_window()
        return self.selected_files
    
    def _create_ui(self):
        """Tạo giao diện"""
        # Header
        header = tk.Label(
            self.dialog,
            text="☁️ Chọn file PDF từ Google Drive",
            font=("Segoe UI", 13, "bold"),
            bg="#4285f4",
            fg="white",
            pady=12
        )
        header.pack(fill=tk.X)
        
        # Info about service account
        info_frame = tk.Frame(self.dialog, bg="#fff3cd", pady=8)
        info_frame.pack(fill=tk.X, padx=15, pady=(10, 0))
        
        service_email = self.drive_manager.get_service_account_email()
        if service_email:
            tk.Label(
                info_frame,
                text=f"📧 Service Account: {service_email}",
                font=("Segoe UI", 8),
                bg="#fff3cd",
                fg="#856404"
            ).pack()
            tk.Label(
                info_frame,
                text="💡 Chỉ hiển thị folder/file được share với email này",
                font=("Segoe UI", 8),
                bg="#fff3cd",
                fg="#856404"
            ).pack()
        
        # Folder selector
        folder_frame = tk.Frame(self.dialog, bg="white", pady=10)
        folder_frame.pack(fill=tk.X, padx=15)
        
        tk.Label(
            folder_frame,
            text="📁 Chọn folder:",
            font=("Segoe UI", 10, "bold"),
            bg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        self.folder_combo_var = tk.StringVar()
        self.folder_combo = tk.ttk.Combobox(
            folder_frame,
            textvariable=self.folder_combo_var,
            state="readonly",
            width=50,
            font=("Segoe UI", 9)
        )
        self.folder_combo.pack(side=tk.LEFT, padx=5)
        self.folder_combo.bind("<<ComboboxSelected>>", self._on_folder_change)
        
        tk.Button(
            folder_frame,
            text="🔄",
            command=self._load_initial,
            bg="#95a5a6",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3
        ).pack(side=tk.LEFT, padx=5)
        
        # Search frame
        search_frame = tk.Frame(self.dialog, bg="white", pady=5)
        search_frame.pack(fill=tk.X, padx=15)
        
        tk.Label(
            search_frame,
            text="🔍 Tìm kiếm:",
            font=("Segoe UI", 10),
            bg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        self.search_entry = tk.Entry(search_frame, width=50, font=("Segoe UI", 10))
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind("<Return>", lambda e: self._do_search())
        
        tk.Button(
            search_frame,
            text="Tìm",
            command=self._do_search,
            bg="#4285f4",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        # File listbox
        list_frame = tk.Frame(self.dialog, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 9),
            bg="#f8f9fa",
            selectbackground="#4285f4",
            selectforeground="white"
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Info label
        self.info_label = tk.Label(
            self.dialog,
            text="Đang tải...",
            font=("Segoe UI", 9),
            fg="#7f8c8d",
            bg="white"
        )
        self.info_label.pack(pady=5)
        
        # Buttons
        btn_frame = tk.Frame(self.dialog, bg="white")
        btn_frame.pack(pady=15)
        
        tk.Button(
            btn_frame,
            text="✅ Thêm file đã chọn",
            command=self._add_selected,
            bg="#27ae60",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=25,
            pady=10
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=self.dialog.destroy,
            bg="#95a5a6",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=25,
            pady=10
        ).pack(side=tk.LEFT, padx=5)
    
    def _load_initial(self):
        """Load folders ban đầu"""
        self.log("📂 Đang tải danh sách folder được share...")
        
        # Lấy shared folders
        folders = self.drive_manager.get_shared_folders()
        
        if not folders:
            self.info_label.config(text="⚠️ Không có folder nào được share với service account")
            messagebox.showwarning(
                "Cảnh báo",
                "Không tìm thấy folder nào!\n\n"
                "Hãy share folder Drive với service account:\n"
                f"{self.drive_manager.get_service_account_email()}"
            )
            return
        
        # Populate combo box
        self.folders_data = [('root', '📁 My Drive')] + [(f['id'], f['name']) for f in folders]
        self.folder_combo['values'] = [name for _, name in self.folders_data]
        self.folder_combo.current(0)
        
        self.log(f"✅ Tìm thấy {len(folders)} folder được share")
        
        # Load files from first folder
        self._load_files_from_folder(self.folders_data[0][0])
    
    def _on_folder_change(self, event):
        """Khi chọn folder khác"""
        idx = self.folder_combo.current()
        if idx >= 0 and idx < len(self.folders_data):
            folder_id, folder_name = self.folders_data[idx]
            self.current_folder_id = folder_id
            self._load_files_from_folder(folder_id)
    
    def _load_files_from_folder(self, folder_id):
        """Load PDF files từ folder"""
        self.info_label.config(text="Đang tải...")
        files = self.drive_manager.list_pdf_files(folder_id)
        self._display_files(files)
    
    def _display_files(self, files):
        """Hiển thị danh sách file"""
        self.file_listbox.delete(0, tk.END)
        
        if not files:
            self.info_label.config(text="Không có file PDF nào")
            return
        
        for f in files:
            size_mb = int(f.get('size', 0)) / (1024 * 1024)
            self.file_listbox.insert(
                tk.END,
                f"📄 {f['name']} ({size_mb:.1f} MB)"
            )
        
        self.file_listbox.files = files
        self.info_label.config(text=f"Tìm thấy {len(files)} file PDF")
    
    def _do_search(self):
        """Tìm kiếm file"""
        query = self.search_entry.get().strip()
        if not query:
            self._load_files_from_folder(self.current_folder_id)
            return
        
        self.info_label.config(text=f"Đang tìm: {query}...")
        self.log(f"🔍 Tìm kiếm: {query}")
        
        files = self.drive_manager.search_files_in_folder(self.current_folder_id, query)
        self._display_files(files)
    
    def _add_selected(self):
        """Thêm file đã chọn"""
        selected = self.file_listbox.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Chưa chọn file nào")
            return
        
        files = self.file_listbox.files
        self.selected_files = []
        
        for idx in selected:
            file = files[idx]
            self.selected_files.append((file['id'], file['name']))
        
        self.dialog.destroy()


class DriveFolderPicker:
    """Dialog chọn folder từ Google Drive"""
    
    def __init__(self, parent, drive_manager, log_callback):
        self.parent = parent
        self.drive_manager = drive_manager
        self.log = log_callback
        self.selected_files = []
    
    def show(self):
        """Hiển thị dialog và trả về danh sách file trong folder"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Chọn thư mục từ Google Drive")
        self.dialog.geometry("750x600")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        self.current_folder_id = ['root']  # ID của folder hiện tại
        self.folder_stack = []  # Stack để quay lại
        self.folder_name_stack = []  # Stack tên folder
        self.available_folders = []  # Danh sách folders được share
        
        self._create_ui()
        self._load_initial()
        
        self.dialog.wait_window()
        return self.selected_files
    
    def _create_ui(self):
        """Tạo giao diện"""
        # Header
        header = tk.Label(
            self.dialog,
            text="☁️ Chọn thư mục từ Google Drive",
            font=("Segoe UI", 13, "bold"),
            bg="#0f9d58",
            fg="white",
            pady=12
        )
        header.pack(fill=tk.X)
        
        # Info
        info_frame = tk.Frame(self.dialog, bg="#fff3cd", pady=8)
        info_frame.pack(fill=tk.X, padx=15, pady=(10, 0))
        
        service_email = self.drive_manager.get_service_account_email()
        if service_email:
            tk.Label(
                info_frame,
                text=f"📧 Service Account: {service_email}",
                font=("Segoe UI", 8),
                bg="#fff3cd",
                fg="#856404"
            ).pack()
            tk.Label(
                info_frame,
                text="💡 Double-click vào folder để xem bên trong",
                font=("Segoe UI", 8),
                bg="#fff3cd",
                fg="#856404"
            ).pack()
        
        # Path
        path_frame = tk.Frame(self.dialog, bg="white", pady=8)
        path_frame.pack(fill=tk.X, padx=15)
        
        tk.Label(
            path_frame,
            text="📂 Đường dẫn:",
            font=("Segoe UI", 9, "bold"),
            bg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        self.path_var = tk.StringVar(value="Shared Folders")
        self.path_label = tk.Label(
            path_frame,
            textvariable=self.path_var,
            font=("Segoe UI", 9),
            bg="white",
            fg="#3498db"
        )
        self.path_label.pack(side=tk.LEFT)
        
        # Folder listbox
        list_frame = tk.Frame(self.dialog, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.folder_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
            bg="#f8f9fa",
            selectbackground="#0f9d58",
            selectforeground="white"
        )
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.folder_listbox.bind("<Double-Button-1>", self._on_double_click)
        scrollbar.config(command=self.folder_listbox.yview)
        
        # Info
        self.info_label = tk.Label(
            self.dialog,
            text="Double-click để mở thư mục | Click 'Chọn' để lấy file PDF",
            font=("Segoe UI", 9),
            fg="#7f8c8d",
            bg="white"
        )
        self.info_label.pack(pady=5)
        
        # Buttons
        btn_frame = tk.Frame(self.dialog, bg="white")
        btn_frame.pack(pady=15)
        
        tk.Button(
            btn_frame,
            text="✅ Chọn thư mục này",
            command=self._select_folder,
            bg="#27ae60",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=25,
            pady=10
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=self.dialog.destroy,
            bg="#95a5a6",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=25,
            pady=10
        ).pack(side=tk.LEFT, padx=5)
    
    def _load_initial(self):
        """Load folders ban đầu"""
        self.log("📂 Đang tải folders được share...")
        folders = self.drive_manager.get_shared_folders()
        
        if not folders:
            messagebox.showwarning(
                "Cảnh báo",
                "Không tìm thấy folder nào!\n\n"
                f"Share folder với: {self.drive_manager.get_service_account_email()}"
            )
            return
        
        # Hiển thị folders
        self._display_folders(folders)
        
        # LƯU DANH SÁCH FOLDERS ĐỂ DÙNG SAU
        self.available_folders = folders
        self.log(f"✅ Tìm thấy {len(folders)} folder được share")
    
    def _display_folders(self, folders):
        """Hiển thị folders"""
        self.folder_listbox.delete(0, tk.END)
        
        # Back button
        if self.current_folder_id[0] != 'root':
            self.folder_listbox.insert(tk.END, "📁 .. (Quay lại)")
        
        for folder in folders:
            self.folder_listbox.insert(tk.END, f"📁 {folder['name']}")
        
        self.folder_listbox.folders = folders
        self.info_label.config(text=f"{len(folders)} thư mục")
    
    def _on_double_click(self, event):
        """Double click vào folder"""
        selection = self.folder_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        item = self.folder_listbox.get(idx)
        
        # Back
        if item == "📁 .. (Quay lại)":
            if self.folder_stack:
                self.current_folder_id[0] = self.folder_stack.pop()
                
                # Lấy danh sách subfolders
                if self.current_folder_id[0] == 'root':
                    # Quay về root, hiển thị shared folders
                    folders = self.available_folders
                else:
                    folders = self.drive_manager.list_folders(self.current_folder_id[0])
                
                self._display_folders(folders)
                
                # Update path
                if self.folder_name_stack:
                    self.folder_name_stack.pop()
                    if self.folder_name_stack:
                        self.path_var.set(" / ".join(self.folder_name_stack))
                    else:
                        self.path_var.set("Shared Folders")
            return
        
        # Enter folder
        actual_idx = idx if self.current_folder_id[0] == 'root' else idx - 1
        if 0 <= actual_idx < len(self.folder_listbox.folders):
            folder = self.folder_listbox.folders[actual_idx]
            
            # Lưu state hiện tại
            self.folder_stack.append(self.current_folder_id[0])
            self.folder_name_stack.append(folder['name'])
            
            # Chuyển sang folder mới
            self.current_folder_id[0] = folder['id']
            self.path_var.set(" / ".join(self.folder_name_stack))
            
            self.log(f"📂 Mở folder: {folder['name']}")
            
            # Load subfolders
            subfolders = self.drive_manager.list_folders(folder['id'])
            self._display_folders(subfolders)
    
    def _select_folder(self):
        """Chọn folder hiện tại"""
        folder_id = self.current_folder_id[0]
        
        # Lấy tên folder hiện tại
        if self.folder_name_stack:
            folder_name = self.folder_name_stack[-1]
        else:
            folder_name = "Shared Folders"
        
        self.log(f"📂 Đang tải file từ: {folder_name} (ID: {folder_id})")
        
        # QUAN TRỌNG: Nếu folder_id là 'root', lấy từ TẤT CẢ shared folders
        if folder_id == 'root':
            self.log("⚠️ Đang ở root, sẽ lấy file từ TẤT CẢ shared folders...")
            all_files = []
            
            for folder in self.available_folders:
                self.log(f"  📁 Đang quét: {folder['name']}")
                files = self.drive_manager.list_pdf_files(folder['id'])
                all_files.extend(files)
                self.log(f"    → Tìm thấy {len(files)} file")
            
            files = all_files
        else:
            # Lấy file từ folder cụ thể
            files = self.drive_manager.list_pdf_files(folder_id)
        
        self.selected_files = []
        for file in files:
            self.selected_files.append((file['id'], file['name']))
        
        self.log(f"✅ Tổng cộng: {len(files)} file PDF")
        
        if len(files) == 0:
            messagebox.showwarning(
                "Không có file",
                f"Không tìm thấy file PDF nào trong '{folder_name}'\n\n"
                "Kiểm tra:\n"
                "1. Folder có chứa PDF không?\n"
                "2. PDF có bị trong subfolder không?\n"
                "3. Service account có quyền truy cập không?"
            )
        
        self.dialog.destroy()