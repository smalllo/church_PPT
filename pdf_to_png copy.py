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

        tk.Button(self.root, text="選擇PDF", command=self.select_pdf_files).pack()


        self.selected_files_label = tk.Label(self.root, text="")
        self.selected_files_label.pack()

        tk.Button(self.root, text="選擇輸出資料夾", command=self.select_output_folder).pack()


        self.output_folder_label = tk.Label(self.root, text="")
        self.output_folder_label.pack()

        tk.Button(self.root, text="轉換成圖片", command=self.convert_to_png).pack()

    def select_pdf_files(self):
        self.pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        self.update_selected_files_label()

    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
        self.update_output_folder_label()

    def update_selected_files_label(self):
        selected_files_text = "\n".join(self.pdf_paths)
        self.selected_files_label.config(text=selected_files_text)

    def update_output_folder_label(self):
        self.output_folder_label.config(text=self.output_folder)

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
