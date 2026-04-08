import customtkinter as ctk
from tkinter import filedialog, messagebox
import img2pdf
import os
from pdf2image import convert_from_path
import pandas as pd
import pdfplumber
import threading
from PIL import Image
import pytesseract

# إعدادات المظهر الفخم باللون الأزرق
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class DocSuiteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DocSuite Pro - Edition 2026")
        self.geometry("700x600")
        self.resizable(False, False)

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#1a1a1a")
        self.sidebar.pack(side="left", fill="y")

        self.logo = ctk.CTkLabel(self.sidebar, text="DOC SUITE", font=ctk.CTkFont(family="Impact", size=26))
        self.logo.pack(pady=(30, 30))

        # أزرار العمليات (منفصلة تماماً)
        btn_font = ctk.CTkFont(size=12, weight="bold")

        # قسم الصور لـ PDF
        self.add_btn("Images to One PDF", self.images_to_single_pdf, btn_font)
        self.add_btn("Each Image to PDF", self.images_to_multi_pdf, btn_font)

        # قسم الصور لـ Excel
        self.add_btn("Images to One Excel", self.images_to_single_excel, btn_font)
        self.add_btn("Each Image to Excel", self.images_to_multi_excel, btn_font)

        # قسم الـ PDF
        self.add_btn("PDF to Images", self.pdf_to_images, btn_font)
        self.add_btn("PDF to Excel", self.pdf_to_excel, btn_font)

        # --- الإمضاء (محمد بخط مائل) ---
        self.signature = ctk.CTkLabel(self.sidebar, text="Mohamed",
                                      font=ctk.CTkFont(family="Lucida Handwriting", size=22, slant="italic"),
                                      text_color="#3b8ed0")
        self.signature.pack(side="bottom", pady=25)

        # --- Main Area ---
        self.main_frame = ctk.CTkFrame(self, fg_color="#121212", corner_radius=20)
        self.main_frame.pack(side="right", fill="both", expand=True, padx=15, pady=15)

        self.label = ctk.CTkLabel(self.main_frame, text="Master Control Panel",
                                  font=ctk.CTkFont(size=22, weight="bold"))
        self.label.pack(pady=(40, 20))

        self.path_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Select folder or file path...",
                                       height=45, width=400, corner_radius=12,
                                       border_color="#3b8ed0", fg_color="#1e1e1e")
        self.path_entry.pack(pady=10)

        self.browse_btn = ctk.CTkButton(self.main_frame, text="BROWSE", font=btn_font,
                                        command=self.browse, fg_color="#3b8ed0", hover_color="#1f538d",
                                        height=45, width=150, corner_radius=12)
        self.browse_btn.pack(pady=20)

        self.status_label = ctk.CTkLabel(self.main_frame, text="Ready", text_color="#3b8ed0", font=("Arial", 12))
        self.status_label.pack(side="bottom", pady=20)

    def add_btn(self, text, command, font):
        btn = ctk.CTkButton(self.sidebar, text=text, corner_radius=8, font=font, height=40,
                            fg_color="#3b8ed0", hover_color="#1f538d",
                            command=lambda: threading.Thread(target=command, daemon=True).start())
        btn.pack(pady=8, padx=20, fill="x")

    def browse(self):
        path = filedialog.askdirectory() or filedialog.askopenfilename()
        if path:
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, path)

    def update_status(self, text, color="#3b8ed0"):
        self.status_label.configure(text=text, text_color=color)

    # --- ميكانيكا العمليات ---

    def images_to_single_pdf(self):
        path = self.path_entry.get()
        if not path or not os.path.exists(path): return
        self.update_status("Merging images to PDF...")
        try:
            files = [os.path.join(path, f) for f in os.listdir(path) if
                     f.lower().endswith(('.png', '.jpg', '.jpeg'))] if os.path.isdir(path) else [path]
            out = os.path.join(os.path.dirname(files[0]), "Merged_Result.pdf")
            with open(out, "wb") as f:
                f.write(img2pdf.convert(files))
            self.update_status("Done: PDF Created!", "green")
        except:
            self.update_status("Error occurred", "red")

    def images_to_multi_pdf(self):
        path = self.path_entry.get()
        if not path or not os.path.isdir(path): return
        self.update_status("Converting each image...")
        try:
            files = [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            for f in files:
                with open(os.path.splitext(f)[0] + ".pdf", "wb") as pdf: pdf.write(img2pdf.convert(f))
            self.update_status(f"Done: {len(files)} PDFs created!", "green")
        except:
            self.update_status("Error occurred", "red")

    def images_to_single_excel(self):
        path = self.path_entry.get()
        if not path: return
        self.update_status("OCR in progress (One File)...")
        try:
            files = [os.path.join(path, f) for f in os.listdir(path) if
                     f.lower().endswith(('.png', '.jpg', '.jpeg'))] if os.path.isdir(path) else [path]
            out = os.path.join(os.path.dirname(files[0]), "Images_Collection.xlsx")
            with pd.ExcelWriter(out) as writer:
                for i, f in enumerate(files):
                    text = pytesseract.image_to_string(Image.open(f))
                    pd.DataFrame([l.split() for l in text.split('\n') if l.strip()]).to_excel(writer,
                                                                                              sheet_name=f"Page_{i + 1}",
                                                                                              index=False, header=False)
            self.update_status("Excel Merged!", "green")
        except:
            self.update_status("OCR Failed", "red")

    def images_to_multi_excel(self):
        path = self.path_entry.get()
        if not path or not os.path.isdir(path): return
        self.update_status("OCR in progress (Multi Files)...")
        try:
            files = [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            for f in files:
                text = pytesseract.image_to_string(Image.open(f))
                pd.DataFrame([l.split() for l in text.split('\n') if l.strip()]).to_excel(
                    os.path.splitext(f)[0] + ".xlsx", index=False, header=False)
            self.update_status("Multiple Excels created!", "green")
        except:
            self.update_status("OCR Failed", "red")

    def pdf_to_images(self):
        path = self.path_entry.get()
        if path.lower().endswith('.pdf'):
            self.update_status("Extracting images...")
            images = convert_from_path(path)
            for i, img in enumerate(images): img.save(os.path.join(os.path.dirname(path), f'Page_{i + 1}.jpg'), 'JPEG')
            self.update_status("Images extracted!", "green")

    def pdf_to_excel(self):
        path = self.path_entry.get()
        if path.lower().endswith('.pdf'):
            self.update_status("Extracting table...")
            data = []
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    tbl = page.extract_table()
                    if tbl: data.extend(tbl)
            pd.DataFrame(data).to_excel(os.path.splitext(path)[0] + ".xlsx", index=False, header=False)
            self.update_status("Excel Ready!", "green")


if __name__ == "__main__":
    app = DocSuiteApp()
    app.mainloop()