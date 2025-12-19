"""
Optimized main.py - Faster startup with lazy imports
"""
import tkinter as tk

def main():
    # Lazy import - chỉ import khi cần
    from gui_app import PDFExtractorApp
    
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()