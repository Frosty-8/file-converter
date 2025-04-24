import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from docx2pdf import convert

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class FileConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DOCX to PDF Converter")
        self.geometry("600x400")

        self.file_paths = []

        self.label = ctk.CTkLabel(self, text="DOCX to PDF Converter", font=ctk.CTkFont(size=20, weight="bold"))
        self.label.pack(pady=20)

        self.select_btn = ctk.CTkButton(self, text="Select DOCX Files", command=self.select_files)
        self.select_btn.pack(pady=10)

        self.output_btn = ctk.CTkButton(self, text="Select Output Folder", command=self.select_output_folder)
        self.output_btn.pack(pady=10)

        self.convert_btn = ctk.CTkButton(self, text="Convert to PDF", command=self.convert_files)
        self.convert_btn.pack(pady=20)

        self.status_label = ctk.CTkLabel(self, text="", wraplength=500)
        self.status_label.pack(pady=10)

        self.output_dir = os.getcwd()

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx")])
        if self.file_paths:
            self.status_label.configure(text=f"Selected {len(self.file_paths)} files.")

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = folder
            self.status_label.configure(text=f"Output folder set to:\n{self.output_dir}")

    def convert_files(self):
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please select DOCX files to convert.")
            return

        converted = 0
        try:
            for docx_path in self.file_paths:
                file_name = os.path.basename(docx_path)
                output_path = os.path.join(self.output_dir, file_name.replace(".docx", ".pdf"))
                convert(docx_path, output_path)
                converted += 1

            self.status_label.configure(text=f"Successfully converted {converted} files to PDF.")
        except Exception as e:
            self.status_label.configure(text=f"Error during conversion: {e}")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()
