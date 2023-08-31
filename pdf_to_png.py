import os
import fitz
import tkinter as tk
from tkinter import filedialog, messagebox

class PdfToPngConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to PNG Converter")

        self.pdf_paths = []
        self.output_folder = ""

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Select PDF Files:").pack()
        tk.Button(self.root, text="Browse", command=self.select_pdf_files).pack()

        tk.Label(self.root, text="Select Output Folder:").pack()
        tk.Button(self.root, text="Browse", command=self.select_output_folder).pack()

        tk.Button(self.root, text="Convert to PNG", command=self.convert_to_png).pack()

    def select_pdf_files(self):
        self.pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    
    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()

    def convert_to_png(self):
        if not self.pdf_paths or not self.output_folder:
            messagebox.showerror("Error", "Please select PDF files and output folder.")
            return
        
        for pdf_path in self.pdf_paths:
            pdf_document = fitz.open(pdf_path)
            pdf_name = os.path.basename(pdf_path).split(".")[0]

            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]
                image = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                
                image_path = os.path.join(self.output_folder, f"{pdf_name}_page_{page_number + 1}.png")
                image.save(image_path)

            pdf_document.close()
        
        messagebox.showinfo("Conversion Complete", "PDF to PNG conversion completed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PdfToPngConverterApp(root)
    root.mainloop()
