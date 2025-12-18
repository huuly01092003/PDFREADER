import tkinter as tk
from gui_app import PDFExtractorApp

def main():
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()