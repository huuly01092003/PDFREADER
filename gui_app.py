import os
import threading
import tempfile
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from config import EXCEL_FILE, GOOGLE_DRIVE_AVAILABLE
from excel_handler import init_excel, read_excel_data
from pdf_processor import process_pdf
from drive_manager import GoogleDriveManager
from dialogs import DriveFilePicker, DriveFolderPicker
from logger_handler import (
    write_log, write_success, write_error, 
    read_log_file, clear_log_file, init_log_files, is_file_processed
)

class LogViewerDialog:
    """Dialog xem file log"""
    
    def __init__(self, parent, log_type="app"):
        self.parent = parent
        self.log_type = log_type
        
        log_titles = {
            "app": "📄 App Log - Nhật ký hoạt động",
            "success": "✅ Success Log - File thành công",
            "error": "❌ Error Log - File thất bại"
        }
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(log_titles.get(log_type, "Log Viewer"))
        self.dialog.geometry("800x600")
        self.dialog.transient(parent)
        
        self._create_ui()
        self._load_content()
    
    def _create_ui(self):
        """Tạo giao diện"""
        # Header
        header = tk.Frame(self.dialog, bg="#34495e", height=50)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        log_icons = {"app": "📄", "success": "✅", "error": "❌"}
        icon = log_icons.get(self.log_type, "📄")
        
        tk.Label(
            header,
            text=f"{icon} Xem Log File",
            font=("Segoe UI", 12, "bold"),
            bg="#34495e",
            fg="white"
        ).pack(pady=12)
        
        # Toolbar
        toolbar = tk.Frame(self.dialog, bg="white", pady=8)
        toolbar.pack(fill=tk.X, padx=10)
        
        tk.Button(
            toolbar,
            text="🔄 Refresh",
            command=self._load_content,
            bg="#3498db",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="🗑️ Clear Log",
            command=self._clear_log,
            bg="#e74c3c",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="💾 Save As...",
            command=self._save_as,
            bg="#27ae60",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        # Text area
        text_frame = tk.Frame(self.dialog, bg="white")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.text_widget = scrolledtext.ScrolledText(
            text_frame,
            font=("Consolas", 9),
            wrap=tk.WORD,
            bg="#1e1e1e",
            fg="#00ff00",
            padx=10,
            pady=10
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Status
        self.status_label = tk.Label(
            self.dialog,
            text="",
            font=("Segoe UI", 9),
            bg="white",
            fg="#7f8c8d"
        )
        self.status_label.pack(pady=5)
        
        # Close button
        tk.Button(
            self.dialog,
            text="❌ Đóng",
            command=self.dialog.destroy,
            bg="#95a5a6",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=30,
            pady=8
        ).pack(pady=10)
    
    def _load_content(self):
        """Load nội dung log"""
        content = read_log_file(self.log_type)
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.insert(1.0, content)
        
        # Count lines
        lines = content.count('\n')
        self.status_label.config(text=f"Tổng: {lines} dòng")
    
    def _clear_log(self):
        """Xóa log"""
        result = messagebox.askyesno(
            "Xác nhận",
            f"Xóa toàn bộ nội dung log này?\n\nHành động không thể hoàn tác!"
        )
        
        if result:
            if clear_log_file(self.log_type):
                self._load_content()
                messagebox.showinfo("Thành công", "Đã xóa log!")
            else:
                messagebox.showerror("Lỗi", "Không thể xóa log")
    
    def _save_as(self):
        """Lưu log ra file khác"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                content = self.text_widget.get(1.0, tk.END)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("Thành công", f"Đã lưu: {file_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu file: {e}")


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart PDF Data Extractor Pro v2.1")
        self.root.geometry("1200x700")
        self.root.configure(bg="#f0f0f0")
        
        self.pdf_files = []
        self.drive_files = []
        self.is_processing = False
        self.drive_manager = GoogleDriveManager()
        self.debug_mode = tk.BooleanVar(value=True)
        
        # Khởi tạo log files
        init_log_files()
        write_log("App started", "info")
        
        self.setup_ui()
        init_excel()
    
    def setup_ui(self):
        """Tạo giao diện"""
        self._create_header()
        
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        self._create_left_panel(main_frame)
        self._create_middle_panel(main_frame)
        self._create_right_panel(main_frame)
        
        self._create_footer()
        
        self.refresh_output()
    
    def _create_header(self):
        """Header"""
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(
            header_frame, 
            text="📄 Smart PDF Data Extractor Pro v2.1",
            font=("Segoe UI", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        ).pack(pady=18)
    
    def _create_left_panel(self, parent):
        """Panel quản lý file"""
        left_frame = tk.Frame(parent, bg="white", relief=tk.RAISED, bd=1)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        
        tk.Label(
            left_frame,
            text="📁 Danh sách file PDF",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(pady=8)
        
        # Listbox
        list_frame = tk.Frame(left_frame, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_frame,
            font=("Consolas", 9),
            selectmode=tk.MULTIPLE,
            yscrollcommand=scrollbar.set,
            bg="#f8f9fa",
            selectbackground="#3498db",
            selectforeground="white",
            relief=tk.FLAT
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Buttons
        self._create_file_buttons(left_frame)
        
        # DEBUG CHECKBOX
        debug_frame = tk.Frame(left_frame, bg="white")
        debug_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Checkbutton(
            debug_frame,
            text="🐛 Debug Mode (xem preview text trong log)",
            variable=self.debug_mode,
            font=("Segoe UI", 9),
            bg="white",
            fg="#e74c3c",
            activebackground="white",
            selectcolor="white"
        ).pack(side=tk.LEFT)
    
    def _create_file_buttons(self, parent):
        """Nút quản lý file"""
        # Row 1
        btn_frame1 = tk.Frame(parent, bg="white")
        btn_frame1.pack(fill=tk.X, padx=10, pady=5)
        
        buttons1 = [
            ("📄 File", self.add_files, "#27ae60"),
            ("📁 Folder", self.add_folder, "#16a085"),
            ("☁️ Drive F", self.add_drive_files, "#4285f4"),
        ]
        
        for text, cmd, color in buttons1:
            tk.Button(
                btn_frame1,
                text=text,
                command=cmd,
                font=("Segoe UI", 8, "bold"),
                bg=color,
                fg="white",
                relief=tk.FLAT,
                cursor="hand2",
                padx=8,
                pady=5
            ).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        # Row 2
        btn_frame2 = tk.Frame(parent, bg="white")
        btn_frame2.pack(fill=tk.X, padx=10, pady=5)
        
        buttons2 = [
            ("☁️ Drive D", self.add_drive_folder, "#0f9d58"),
            ("🗑️ Xóa chọn", self.clear_selected, "#e74c3c"),
            ("🗑️ Xóa tất", self.clear_all, "#c0392b"),
        ]
        
        for text, cmd, color in buttons2:
            tk.Button(
                btn_frame2,
                text=text,
                command=cmd,
                font=("Segoe UI", 8, "bold"),
                bg=color,
                fg="white",
                relief=tk.FLAT,
                cursor="hand2",
                padx=8,
                pady=5
            ).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
    
    def _create_middle_panel(self, parent):
        """Panel log và xử lý"""
        middle_frame = tk.Frame(parent, bg="white", relief=tk.RAISED, bd=1)
        middle_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        
        tk.Label(
            middle_frame,
            text="📊 Nhật ký xử lý",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(pady=8)
        
        # Log text
        log_frame = tk.Frame(middle_frame, bg="white")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        log_scrollbar = tk.Scrollbar(log_frame)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 8),
            wrap=tk.WORD,
            yscrollcommand=log_scrollbar.set,
            bg="#1e1e1e",
            fg="#00ff00",
            relief=tk.FLAT,
            padx=8,
            pady=8
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.config(command=self.log_text.yview)
        
        # Progress
        self.progress = ttk.Progressbar(middle_frame, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_label = tk.Label(
            middle_frame,
            text="Sẵn sàng",
            font=("Segoe UI", 9),
            bg="white",
            fg="#7f8c8d"
        )
        self.status_label.pack(pady=3)
        
        # Process button
        self.process_btn = tk.Button(
            middle_frame,
            text="🚀 Bắt đầu xử lý",
            command=self.start_processing,
            font=("Segoe UI", 11, "bold"),
            bg="#3498db",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=25,
            pady=10
        )
        self.process_btn.pack(pady=12)
    
    def _create_right_panel(self, parent):
        """Panel output"""
        right_frame = tk.Frame(parent, bg="white", relief=tk.RAISED, bd=1)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Header
        output_header = tk.Frame(right_frame, bg="white")
        output_header.pack(fill=tk.X, pady=8, padx=10)
        
        tk.Label(
            output_header,
            text="📋 Kết quả",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(side=tk.LEFT)
        
        # Row 1 buttons
        tk.Button(
            output_header,
            text="📄 Log",
            command=lambda: self.view_log("app"),
            font=("Segoe UI", 8),
            bg="#9b59b6",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        tk.Button(
            output_header,
            text="❌ Error",
            command=lambda: self.view_log("error"),
            font=("Segoe UI", 8),
            bg="#e74c3c",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        tk.Button(
            output_header,
            text="✅ Success",
            command=lambda: self.view_log("success"),
            font=("Segoe UI", 8),
            bg="#27ae60",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        # Second row
        output_header2 = tk.Frame(right_frame, bg="white")
        output_header2.pack(fill=tk.X, pady=(0, 8), padx=10)
        
        tk.Button(
            output_header2,
            text="🗑️ Clear",
            command=self.clear_data,
            font=("Segoe UI", 8),
            bg="#e74c3c",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        tk.Button(
            output_header2,
            text="🔄 Refresh",
            command=self.refresh_output,
            font=("Segoe UI", 8),
            bg="#95a5a6",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        tk.Button(
            output_header2,
            text="📊 Excel",
            command=self.open_excel,
            font=("Segoe UI", 8),
            bg="#e67e22",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=4
        ).pack(side=tk.RIGHT, padx=2)
        
        # Treeview
        tree_frame = tk.Frame(right_frame, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        tree_scroll_y = tk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.output_tree = ttk.Treeview(
            tree_frame,
            columns=("Time", "File", "PO", "SKU", "Desc", "Cost", "Qty", "Total"),
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            height=15
        )
        
        tree_scroll_y.config(command=self.output_tree.yview)
        tree_scroll_x.config(command=self.output_tree.xview)
        
        # Columns
        columns = [
            ("Time", 85), ("File", 120), ("PO", 100), ("SKU", 90),
            ("Desc", 150), ("Cost", 80), ("Qty", 60), ("Total", 90)
        ]
        
        for col_id, width in columns:
            self.output_tree.heading(col_id, text=col_id)
            self.output_tree.column(col_id, width=width, anchor=tk.W)
        
        self.output_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Stats
        self.stats_label = tk.Label(
            right_frame,
            text="Tổng: 0 dòng",
            font=("Segoe UI", 9),
            bg="white",
            fg="#7f8c8d"
        )
        self.stats_label.pack(pady=5)
    
    def _create_footer(self):
        """Footer"""
        tk.Label(
            self.root,
            text="💡 Miễn phí 100% | Version 2.1 - Enhanced | With logging & duplicate prevention",
            font=("Segoe UI", 8),
            bg="#ecf0f1",
            fg="#7f8c8d",
            pady=8
        ).pack(side=tk.BOTTOM, fill=tk.X)
    
    # ========== METHODS ==========
    
    def log(self, message):
        """Ghi log vào UI và file"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        # Ghi vào file log
        write_log(message, "info")
    
    def add_files(self):
        """Thêm file từ máy"""
        files = filedialog.askopenfilenames(
            title="Chọn file PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        count = 0
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.file_listbox.insert(tk.END, f"📄 {os.path.basename(file)}")
                count += 1
        
        if count > 0:
            self.log(f"✅ Đã thêm {count} file")
            write_log(f"Added {count} local files", "info")
    
    def add_folder(self):
        """Thêm folder từ máy"""
        folder = filedialog.askdirectory(title="Chọn thư mục")
        
        if not folder:
            return
        
        pdf_files = list(Path(folder).glob("*.pdf"))
        
        count = 0
        for pdf_file in pdf_files:
            file_path = str(pdf_file)
            if file_path not in self.pdf_files:
                self.pdf_files.append(file_path)
                self.file_listbox.insert(tk.END, f"📄 {pdf_file.name}")
                count += 1
        
        if count > 0:
            self.log(f"✅ Đã thêm {count} file từ thư mục")
            write_log(f"Added {count} files from folder: {folder}", "info")
    
    def add_drive_files(self):
        """Chọn file từ Drive"""
        if not GOOGLE_DRIVE_AVAILABLE:
            messagebox.showerror("Lỗi", "Chưa cài đặt thư viện Google Drive")
            return
        
        if not self.drive_manager.authenticated:
            self.log("🔐 Đang xác thực Google Drive...")
            success, message = self.drive_manager.authenticate()
            if not success:
                messagebox.showerror("Lỗi", message)
                return
            self.log(f"✅ {message}")
        
        picker = DriveFilePicker(self.root, self.drive_manager, self.log)
        files = picker.show()
        
        if files:
            for file_id, file_name in files:
                self.drive_files.append((file_id, file_name))
                self.pdf_files.append(f"drive://{file_id}")
                self.file_listbox.insert(tk.END, f"☁️ {file_name}")
            
            self.log(f"✅ Đã thêm {len(files)} file từ Drive")
            write_log(f"Added {len(files)} files from Google Drive", "info")
    
    def add_drive_folder(self):
        """Chọn folder từ Drive"""
        if not GOOGLE_DRIVE_AVAILABLE:
            messagebox.showerror("Lỗi", "Chưa cài đặt thư viện Google Drive")
            return
        
        if not self.drive_manager.authenticated:
            self.log("🔐 Đang xác thực Google Drive...")
            success, message = self.drive_manager.authenticate()
            if not success:
                messagebox.showerror("Lỗi", message)
                return
            self.log(f"✅ {message}")
        
        picker = DriveFolderPicker(self.root, self.drive_manager, self.log)
        files = picker.show()
        
        if files:
            for file_id, file_name in files:
                self.drive_files.append((file_id, file_name))
                self.pdf_files.append(f"drive://{file_id}")
                self.file_listbox.insert(tk.END, f"☁️ {file_name}")
            
            self.log(f"✅ Đã thêm {len(files)} file từ Drive")
            write_log(f"Added {len(files)} files from Google Drive folder", "info")
    
    def clear_selected(self):
        """Xóa file đã chọn"""
        selected = self.file_listbox.curselection()
        if not selected:
            self.log("⚠️ Chưa chọn file")
            return
        
        for index in reversed(selected):
            self.file_listbox.delete(index)
            del self.pdf_files[index]
        
        self.log(f"🗑️ Đã xóa {len(selected)} file")
        write_log(f"Removed {len(selected)} selected files", "info")
    
    def clear_all(self):
        """Xóa tất cả file"""
        if not self.pdf_files:
            self.log("⚠️ Không có file nào")
            return
        
        result = messagebox.askyesno(
            "Xác nhận",
            f"Xóa TẤT CẢ {len(self.pdf_files)} file trong danh sách?"
        )
        
        if result:
            count = len(self.pdf_files)
            self.file_listbox.delete(0, tk.END)
            self.pdf_files.clear()
            self.drive_files.clear()
            
            self.log(f"🗑️ Đã xóa tất cả {count} file")
            write_log(f"Cleared all {count} files from list", "info")
    
    def refresh_output(self):
        """Làm mới output"""
        for item in self.output_tree.get_children():
            self.output_tree.delete(item)
        
        try:
            data = read_excel_data()
            
            for row in data:
                if len(row) >= 9:
                    display_row = (
                        row[0],
                        row[1][:20] + "..." if len(str(row[1])) > 20 else row[1],
                        row[2],
                        row[3],
                        str(row[4])[:30] + "..." if len(str(row[4])) > 30 else row[4],
                        row[5],
                        row[7],
                        row[8]
                    )
                    self.output_tree.insert("", tk.END, values=display_row)
            
            self.stats_label.config(text=f"Tổng: {len(data)} dòng")
            
        except Exception as e:
            self.log(f"⚠️ Lỗi refresh output: {e}")
            self.stats_label.config(text="Tổng: 0 dòng")
    
    def open_excel(self):
        """Mở Excel"""
        if os.path.exists(EXCEL_FILE):
            try:
                os.startfile(EXCEL_FILE)
                self.log("📊 Đã mở Excel")
                write_log("Opened Excel file", "info")
            except:
                messagebox.showinfo("Thông báo", f"File: {EXCEL_FILE}")
        else:
            init_excel()
            self.log("✅ Đã tạo file Excel mới")
            try:
                os.startfile(EXCEL_FILE)
            except:
                messagebox.showinfo("Thông báo", f"Đã tạo file: {EXCEL_FILE}")
    
    def view_log(self, log_type):
        """Xem file log"""
        LogViewerDialog(self.root, log_type)
        write_log(f"Viewed {log_type} log", "info")
    
    def clear_data(self):
        """Xóa dữ liệu Excel"""
        result = messagebox.askyesno(
            "Xác nhận",
            "Xóa TẤT CẢ dữ liệu trong Excel?\n\n(Header sẽ được giữ lại)"
        )
        
        if result:
            from excel_handler import clear_excel_data
            if clear_excel_data():
                self.log("🗑️ Đã xóa dữ liệu Excel")
                write_log("Cleared Excel data", "info")
                self.refresh_output()
                messagebox.showinfo("Thành công", "Đã xóa dữ liệu!")
            else:
                messagebox.showerror("Lỗi", "Không thể xóa dữ liệu")
    
    def start_processing(self):
        """Bắt đầu xử lý"""
        if not self.pdf_files:
            messagebox.showwarning("Cảnh báo", "Chưa có file!")
            return
        
        if self.is_processing:
            messagebox.showinfo("Thông báo", "Đang xử lý...")
            return
        
        self.is_processing = True
        self.process_btn.config(state=tk.DISABLED, text="⏳ Đang xử lý...")
        
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
    
    def process_files(self):
        """Xử lý các file"""
        total = len(self.pdf_files)
        success = 0
        failed = 0
        skipped = 0
        
        self.log("\n" + "="*50)
        self.log(f"🚀 Bắt đầu xử lý {total} files")
        if self.debug_mode.get():
            self.log("🐛 DEBUG MODE: ON - Preview text trong log")
        self.log("="*50 + "\n")
        
        write_log(f"Started processing {total} files", "info")
        
        temp_dir = tempfile.mkdtemp()
        
        for i, pdf_path in enumerate(self.pdf_files, 1):
            try:
                self.status_label.config(text=f"Đang xử lý {i}/{total}...")
                self.progress['value'] = (i / total) * 100
                
                # Xác định tên file
                if pdf_path.startswith("drive://"):
                    file_id = pdf_path.replace("drive://", "")
                    
                    file_name = None
                    for fid, fname in self.drive_files:
                        if fid == file_id:
                            file_name = fname
                            break
                    
                    if not file_name:
                        raise Exception("Không tìm thấy file")
                    
                    filename_for_log = file_name
                else:
                    filename_for_log = os.path.basename(pdf_path)
                
                # KIỂM TRA FILE ĐÃ XỬ LÝ THÀNH CÔNG CHƯA
                if is_file_processed(filename_for_log):
                    skipped += 1
                    self.log(f"⏭️ [{i}/{total}] Bỏ qua (đã xử lý): {filename_for_log}\n")
                    write_log(f"Skipped already processed file: {filename_for_log}", "info")
                    continue
                
                # Xử lý file
                if pdf_path.startswith("drive://"):
                    self.log(f"☁️ Đang tải: {file_name}")
                    
                    temp_path = os.path.join(temp_dir, file_name)
                    self.drive_manager.download_file(file_id, temp_path)
                    
                    items = process_pdf(temp_path, self.log, self.debug_mode.get())
                    os.remove(temp_path)
                else:
                    items = process_pdf(pdf_path, self.log, self.debug_mode.get())
                
                success += 1
                self.log(f"✅ [{i}/{total}] Thành công: {items} items\n")
                
                # Ghi vào success log
                write_success(filename_for_log)
                
            except Exception as e:
                failed += 1
                filename = "Drive file" if pdf_path.startswith("drive://") else os.path.basename(pdf_path)
                self.log(f"❌ [{i}/{total}] Lỗi '{filename}': {e}\n")
                
                # Ghi vào error log
                write_error(filename, str(e))
        
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        self.log("="*50)
        self.log("🎉 HOÀN TẤT")
        self.log("="*50)
        self.log(f"✅ Thành công: {success} files")
        self.log(f"⏭️ Bỏ qua: {skipped} files (đã xử lý)")
        self.log(f"❌ Thất bại: {failed} files\n")
        
        if self.debug_mode.get():
            self.log("💡 TIP: Debug mode ON - xem preview text trong log")
        
        write_log(f"Processing completed: {success} success, {skipped} skipped, {failed} failed", "info")
        
        self.status_label.config(text=f"Hoàn tất: {success}/{total} (skip: {skipped})")
        self.progress['value'] = 100
        
        self.refresh_output()
        
        self.is_processing = False
        self.process_btn.config(state=tk.NORMAL, text="🚀 Bắt đầu xử lý")
        
        messagebox.showinfo(
            "Hoàn tất",
            f"✅ Thành công: {success}\n⏭️ Bỏ qua: {skipped}\n❌ Thất bại: {failed}"
        )