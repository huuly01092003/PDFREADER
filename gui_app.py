import os
import threading
import tempfile
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from config import EXCEL_FILE, GOOGLE_DRIVE_AVAILABLE
from excel_handler import init_excel, read_excel_data
from excel_handler_format2 import init_excel_format2, read_excel_data_format2
from pdf_processor import process_pdf
from pdf_processor_format2 import process_pdf_format2
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
        
        self.status_label = tk.Label(
            self.dialog,
            text="",
            font=("Segoe UI", 9),
            bg="white",
            fg="#7f8c8d"
        )
        self.status_label.pack(pady=5)
        
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


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart PDF Data Extractor Pro v3.0 - Dual Format")
        self.root.geometry("1500x750")
        self.root.configure(bg="#f0f0f0")
        
        self.pdf_files = []
        self.drive_files = []
        self.is_processing = False
        self.drive_manager = GoogleDriveManager()
        self.debug_mode = tk.BooleanVar(value=True)
        self.current_format = tk.StringVar(value="format1")  # format1 hoặc format2
        
        init_log_files()
        write_log("App started - Dual Format Mode", "info")
        
        self.setup_ui()
        init_excel()
        init_excel_format2()
    
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
            text="📄 Smart PDF Data Extractor Pro v3.0 - Dual Format",
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
        
        # Format selector
        format_frame = tk.Frame(left_frame, bg="#fff3cd", pady=8)
        format_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(
            format_frame,
            text="📋 Format:",
            font=("Segoe UI", 9, "bold"),
            bg="#fff3cd",
            fg="#856404"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            format_frame,
            text="Format 1 (Old)",
            variable=self.current_format,
            value="format1",
            font=("Segoe UI", 9),
            bg="#fff3cd",
            activebackground="#fff3cd",
            command=self.on_format_change
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            format_frame,
            text="Format 2 (New Order)",
            variable=self.current_format,
            value="format2",
            font=("Segoe UI", 9),
            bg="#fff3cd",
            activebackground="#fff3cd",
            command=self.on_format_change
        ).pack(side=tk.LEFT, padx=5)
        
        list_frame = tk.Frame(left_frame, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 10))
        
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
        
        self._create_file_buttons(left_frame)
        
        debug_frame = tk.Frame(left_frame, bg="white")
        debug_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Checkbutton(
            debug_frame,
            text="🐛 Debug Mode",
            variable=self.debug_mode,
            font=("Segoe UI", 9),
            bg="white",
            fg="#e74c3c",
            activebackground="white",
            selectcolor="white"
        ).pack(side=tk.LEFT)
    
    def _create_file_buttons(self, parent):
        """Nút quản lý file"""
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
        
        output_header = tk.Frame(right_frame, bg="white")
        output_header.pack(fill=tk.X, pady=8, padx=10)
        
        tk.Label(
            output_header,
            text="📋 Kết quả",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(side=tk.LEFT)
        
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
        
        tree_frame = tk.Frame(right_frame, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        tree_scroll_y = tk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.output_tree = ttk.Treeview(
            tree_frame,
            columns=(),  # Will be set dynamically
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            height=15
        )
        
        tree_scroll_y.config(command=self.output_tree.yview)
        tree_scroll_x.config(command=self.output_tree.xview)
        
        self.output_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
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
            text="💡 Version 3.0 - Dual Format Support | Format 1: 14 cols | Format 2: 19 cols",
            font=("Segoe UI", 8),
            bg="#ecf0f1",
            fg="#7f8c8d",
            pady=8
        ).pack(side=tk.BOTTOM, fill=tk.X)
    
    def on_format_change(self):
        """Khi thay đổi format"""
        format_type = self.current_format.get()
        self.log(f"📋 Chuyển sang {format_type.upper()}")
        self.refresh_output()
    
    def log(self, message):
        """Ghi log vào UI và file"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
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
        """Làm mới output theo format đã chọn"""
        format_type = self.current_format.get()
        
        # Clear existing tree
        for item in self.output_tree.get_children():
            self.output_tree.delete(item)
        
        # Remove old columns
        self.output_tree['columns'] = ()
        
        try:
            if format_type == "format1":
                # Format 1: 14 columns
                columns = ["Time", "File", "PO", "SKU", "Desc", "VendorPart", "SellUM", "BuyUM",
                          "Buy", "Net", "QtyCS", "QtyOrdPcs", "QtyRecPcs", "Extended"]
                col_widths = [65, 100, 90, 80, 120, 115, 60, 60, 75, 75, 60, 70, 70, 85]
                
                data = read_excel_data()
                
            else:  # format2
                # Format 2: 19 columns
                columns = ["Time", "File", "OrderNo", "OrderDate", "Supplier", "Contract",
                          "OrderedBy", "DeliveredTo", "ForStore", "Article", "Desc",
                          "OUType", "LV", "SKU_OU", "OUQty", "FreeQty", "NetPrice",
                          "Unit", "TotalNet"]
                col_widths = [65, 90, 90, 70, 70, 70, 100, 100, 100, 90, 120,
                             60, 40, 60, 60, 60, 85, 50, 90]
                
                data = read_excel_data_format2()
            
            # Setup columns
            self.output_tree['columns'] = columns
            for col_id, width in zip(columns, col_widths):
                self.output_tree.heading(col_id, text=col_id)
                self.output_tree.column(col_id, width=width, anchor=tk.W)
            
            # Insert data
            for row in data:
                if len(row) >= len(columns):
                    display_row = []
                    for i, val in enumerate(row[:len(columns)]):
                        if i == 0:  # Time
                            display_row.append(val[:5] if val else "")
                        elif i == 1:  # Filename
                            s = str(val)
                            display_row.append(s[:15] + "..." if len(s) > 15 else s)
                        elif i in [4, 10]:  # Description fields
                            s = str(val)
                            display_row.append(s[:20] + "..." if len(s) > 20 else s)
                        elif i in [6, 7, 8] and format_type == "format2":  # Address fields
                            s = str(val)
                            display_row.append(s[:15] + "..." if len(s) > 15 else s)
                        else:
                            display_row.append(val)
                    
                    self.output_tree.insert("", tk.END, values=display_row)
            
            self.stats_label.config(text=f"Tổng: {len(data)} dòng | Format: {format_type.upper()}")
            
        except Exception as e:
            self.log(f"⚠️ Lỗi refresh output: {e}")
            self.stats_label.config(text="Tổng: 0 dòng")
    
    def open_excel(self):
        """Mở Excel theo format đã chọn"""
        format_type = self.current_format.get()
        
        if format_type == "format1":
            excel_file = EXCEL_FILE
        else:
            from excel_handler_format2 import EXCEL_FILE_FORMAT2
            excel_file = EXCEL_FILE_FORMAT2
        
        if os.path.exists(excel_file):
            try:
                os.startfile(excel_file)
                self.log(f"📊 Đã mở Excel {format_type}")
                write_log(f"Opened Excel file {format_type}", "info")
            except:
                messagebox.showinfo("Thông báo", f"File: {excel_file}")
        else:
            if format_type == "format1":
                init_excel()
            else:
                init_excel_format2()
            self.log("✅ Đã tạo file Excel mới")
            try:
                os.startfile(excel_file)
            except:
                messagebox.showinfo("Thông báo", f"Đã tạo file: {excel_file}")
    
    def view_log(self, log_type):
        """Xem file log"""
        LogViewerDialog(self.root, log_type)
        write_log(f"Viewed {log_type} log", "info")
    
    def clear_data(self):
        """Xóa dữ liệu Excel"""
        format_type = self.current_format.get()
        
        result = messagebox.askyesno(
            "Xác nhận",
            f"Xóa TẤT CẢ dữ liệu trong Excel {format_type.upper()}?\n\n(Header sẽ được giữ lại)"
        )
        
        if result:
            if format_type == "format1":
                from excel_handler import clear_excel_data
                success = clear_excel_data()
            else:
                from excel_handler_format2 import clear_excel_data_format2
                success = clear_excel_data_format2()
            
            if success:
                self.log(f"🗑️ Đã xóa dữ liệu Excel {format_type}")
                write_log(f"Cleared Excel data {format_type}", "info")
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
        
        format_type = self.current_format.get()
        
        self.log("\n" + "="*50)
        self.log(f"🚀 Bắt đầu xử lý {total} files - {format_type.upper()}")
        if self.debug_mode.get():
            self.log("🐛 DEBUG MODE: ON")
        self.log("="*50 + "\n")
        
        write_log(f"Started processing {total} files - {format_type}", "info")
        
        temp_dir = tempfile.mkdtemp()
        
        for i, pdf_path in enumerate(self.pdf_files, 1):
            filename_for_log = ""
            
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
                        file_name = f"Unknown_Drive_File_{file_id[:8]}"
                    
                    filename_for_log = file_name
                else:
                    filename_for_log = os.path.basename(pdf_path)
                
                # Kiểm tra đã xử lý chưa
                if is_file_processed(filename_for_log):
                    skipped += 1
                    self.log(f"⏭️ [{i}/{total}] Bỏ qua (đã xử lý): {filename_for_log}\n")
                    write_log(f"Skipped already processed file: {filename_for_log}", "info")
                    continue
                
                # Xử lý file
                if pdf_path.startswith("drive://"):
                    self.log(f"☁️ [{i}/{total}] Đang tải: {filename_for_log}")
                    
                    temp_path = os.path.join(temp_dir, filename_for_log)
                    
                    download_success = self.drive_manager.download_file(file_id, temp_path)
                    if not download_success:
                        raise Exception("Không thể tải file từ Drive")
                    
                    # Process theo format
                    if format_type == "format1":
                        items = process_pdf(temp_path, self.log, self.debug_mode.get())
                    else:
                        items = process_pdf_format2(temp_path, self.log, self.debug_mode.get())
                    
                    try:
                        os.remove(temp_path)
                    except:
                        pass
                else:
                    self.log(f"📄 [{i}/{total}] Đang xử lý: {filename_for_log}")
                    
                    if format_type == "format1":
                        items = process_pdf(pdf_path, self.log, self.debug_mode.get())
                    else:
                        items = process_pdf_format2(pdf_path, self.log, self.debug_mode.get())
                
                # Thành công
                success += 1
                self.log(f"✅ [{i}/{total}] Thành công: {items} items\n")
                write_success(filename_for_log)
                
            except Exception as e:
                failed += 1
                
                if not filename_for_log:
                    if pdf_path.startswith("drive://"):
                        filename_for_log = f"Drive_File_{pdf_path[8:16]}"
                    else:
                        filename_for_log = os.path.basename(pdf_path) if pdf_path else "Unknown_File"
                
                error_msg = str(e)
                self.log(f"❌ [{i}/{total}] Lỗi '{filename_for_log}': {error_msg}\n")
                
                write_error(filename_for_log, error_msg)
                write_log(f"Failed to process '{filename_for_log}': {error_msg}", "error")
        
        # Dọn dẹp
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        # Kết quả
        self.log("="*50)
        self.log("🎉 HOÀN TẤT")
        self.log("="*50)
        self.log(f"✅ Thành công: {success} files")
        self.log(f"⏭️ Bỏ qua: {skipped} files (đã xử lý)")
        self.log(f"❌ Thất bại: {failed} files")
        
        if failed > 0:
            self.log(f"\n💡 Xem danh sách file thất bại trong Error Log")
        
        write_log(f"Processing completed: {success} success, {skipped} skipped, {failed} failed - {format_type}", "info")
        
        self.status_label.config(text=f"Hoàn tất: {success}/{total} (skip: {skipped}, fail: {failed})")
        self.progress['value'] = 100
        
        self.refresh_output()
        
        self.is_processing = False
        self.process_btn.config(state=tk.NORMAL, text="🚀 Bắt đầu xử lý")
        
        messagebox.showinfo(
            "Hoàn tất",
            f"Format: {format_type.upper()}\n\n✅ Thành công: {success}\n⏭️ Bỏ qua: {skipped}\n❌ Thất bại: {failed}"
        )