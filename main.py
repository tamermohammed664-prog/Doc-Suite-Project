import customtkinter as ctk
from tkinter import filedialog, messagebox
import img2pdf
import os
from pdf2image import convert_from_path
import tabula
import pandas as pd

# Appearance Settings
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class DocSuiteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("All-in-One Doc Suite - Turbo Mode")
        self.geometry("800x550")
        self.resizable(False, False)

        # Grid Layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        self.logo = ctk.CTkLabel(self.sidebar, text="DOC SUITE", font=ctk.CTkFont(size=18, weight="bold"))
        self.logo.pack(pady=30)

        btn_font = ctk.CTkFont(size=12)

        self.btn_single = ctk.CTkButton(self.sidebar, text="Images to 1 PDF", corner_radius=20, font=btn_font,
                                        command=self.images_to_single_pdf)
        self.btn_single.pack(pady=10, padx=15)

        self.btn_multi = ctk.CTkButton(self.sidebar, text="Each Image to PDF", corner_radius=20, font=btn_font,
                                       command=self.images_to_multiple_pdfs)
        self.btn_multi.pack(pady=10, padx=15)

        self.btn_pdf_img = ctk.CTkButton(self.sidebar, text="PDF to Images", corner_radius=20, font=btn_font,
                                         command=self.pdf_to_images)
        self.btn_pdf_img.pack(pady=10, padx=15)

        self.btn_pdf_excel = ctk.CTkButton(self.sidebar, text="PDF to Excel", corner_radius=20, font=btn_font,
                                           command=self.pdf_to_excel)
        self.btn_pdf_excel.pack(pady=10, padx=15)

        self.btn_excel_pdf = ctk.CTkButton(self.sidebar, text="Excel to PDF", corner_radius=20, font=btn_font,
                                           command=self.placeholder)
        self.btn_excel_pdf.pack(pady=10, padx=15)

        # --- Main Area ---
        self.main_frame = ctk.CTkFrame(self, fg_color="#E8F0F7", corner_radius=25)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self.label = ctk.CTkLabel(self.main_frame, text="All-in-One Doc Suite",
                                  font=ctk.CTkFont(size=22, weight="bold"))
        self.label.pack(pady=20)

        self.path_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.path_frame.pack(fill="x", padx=40)

        self.path_entry = ctk.CTkEntry(self.path_frame, placeholder_text="Select folder or file path...",
                                       corner_radius=15, height=35)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_btn = ctk.CTkButton(self.path_frame, text="Browse", width=80, corner_radius=15, fg_color="#34495e",
                                        command=self.browse)
        self.browse_btn.pack(side="right")

        self.log_box = ctk.CTkTextbox(self.main_frame, corner_radius=15, border_width=1)
        self.log_box.pack(pady=25, padx=40, fill="both", expand=True)
        self.log_box.insert("0.0",
                            "System Ready...\n- For Folder tasks: Select folder.\n- For Single File tasks: Select the file.")

    def browse(self):
        # يفتح اختيار ملف أو فولدر حسب العملية
        path = filedialog.askdirectory()
        if not path:  # لو مدسش على فولدر جرب يختار ملف
            path = filedialog.askopenfilename()
        if path:
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, path)

    def log(self, text):
        self.log_box.insert("end", f"\n> {text}")
        self.log_box.see("end")

    # 1. Images to 1 PDF
    def images_to_single_pdf(self):
        folder = self.path_entry.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Warning", "Please select a Folder containing images.")
            return
        try:
            imgs = [os.path.join(folder, f) for f in os.listdir(folder) if
                    f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            if not imgs:
                self.log("No images found.")
                return
            output = os.path.join(folder, "Merged_Output.pdf")
            with open(output, "wb") as f:
                f.write(img2pdf.convert(imgs))
            self.log(f"Successfully merged into: {output}")
        except Exception as e:
            self.log(f"Error: {str(e)}")

    # 2. Each Image to PDF
    def images_to_multiple_pdfs(self):
        folder = self.path_entry.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Warning", "Please select a Folder.")
            return
        try:
            for f in os.listdir(folder):
                if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                    img_path = os.path.join(folder, f)
                    output = os.path.join(folder, f"{os.path.splitext(f)[0]}.pdf")
                    with open(output, "wb") as pdf_file:
                        pdf_file.write(img2pdf.convert(img_path))
            self.log("All images converted to individual PDFs.")
        except Exception as e:
            self.log(f"Error: {str(e)}")

    # 3. PDF to Images
    def pdf_to_images(self):
        file_path = self.path_entry.get()
        if not file_path or not file_path.lower().endswith('.pdf'):
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        try:
            images = convert_from_path(file_path)
            folder = os.path.dirname(file_path)
            for i, image in enumerate(images):
                image.save(os.path.join(folder, f'Page_{i + 1}.jpg'), 'JPEG')
            self.log(f"Extracted {len(images)} images to the same folder.")
        except Exception as e:
            self.log(f"Error: {str(e)} (Ensure Poppler is installed)")

    # 4. PDF to Excel (The New Feature)
    def pdf_to_excel(self):
        file_path = self.path_entry.get()
        if not file_path or not file_path.lower().endswith('.pdf'):
            messagebox.showwarning("Warning", "Please select a PDF file.")
            return
        try:
            self.log("Analyzing PDF for tables... please wait.")
            output = os.path.splitext(file_path)[0] + "_Converted.xlsx"
            # استخراج الجداول من كل الصفحات
            tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
            if not tables:
                self.log("No tables found in this PDF.")
                return

            with pd.ExcelWriter(output) as writer:
                for i, df in enumerate(tables):
                    df.to_excel(writer, sheet_name=f'Table_{i + 1}', index=False)

            self.log(f"Excel file created: {output}")
        except Exception as e:
            self.log(f"Error: {str(e)}")

    def placeholder(self):
        self.log("This feature is under development.")


if __name__ == "__main__":
    app = DocSuiteApp()
    app.mainloop()