import customtkinter as ctk
from tkinter import filedialog, simpledialog
import img2pdf
import os
import pandas as pd
import pdfplumber
import threading
from PIL import Image
import pytesseract
import pymupdf
import arabic_reshaper
from bidi.algorithm import get_display
from pypdf import PdfReader, PdfWriter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class DocSuiteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DocSuite 2026")
        self.geometry("650x580") # العرض قل جداً لـ 650
        self.resizable(False, False)

        self.input_path = ""
        self.output_dir = ""
        self.stop_execution = False

        # --- Sidebar (Narrower) ---
        self.sidebar = ctk.CTkFrame(self, width=160, corner_radius=0, fg_color="#1a1a1a")
        self.sidebar.pack(side="left", fill="y")

        self.logo = ctk.CTkLabel(self.sidebar, text="DOC SUITE", font=ctk.CTkFont(family="Impact", size=20))
        self.logo.pack(pady=(15, 10))

        btn_font = ctk.CTkFont(size=11, weight="bold")

        # الأزرار (أصغر وأقرب لبعض)
        self.add_btn("IMG to ONE PDF", self.images_to_single_pdf, btn_font)
        self.add_btn("IMGs to PDFs", self.images_to_multi_pdf, btn_font)
        self.add_btn("PDF to JPG", self.pdf_to_images, btn_font)
        self.add_btn("IMG to EXCEL", self.images_to_single_excel, btn_font)
        self.add_btn("PDF to EXCEL", self.pdf_to_excel, btn_font)
        self.add_btn("MERGE PDFs", self.merge_pdfs_action, btn_font)
        self.add_btn("DELETE PAGES", self.delete_pages_action, btn_font)

        self.stop_btn = ctk.CTkButton(self.sidebar, text="STOP", corner_radius=6,
                                      font=btn_font, height=35, fg_color="#c0392b",
                                      hover_color="#e74c3c", command=self.stop_action)
        self.stop_btn.pack(pady=20, padx=15, fill="x")

        self.signature = ctk.CTkLabel(self.sidebar, text="Mohamed",
                                      font=ctk.CTkFont(family="Lucida Handwriting", size=18),
                                      text_color="#3b8ed0")
        self.signature.pack(side="bottom", pady=10)

        # --- Main Area (Compact) ---
        self.main_frame = ctk.CTkFrame(self, fg_color="#121212", corner_radius=15)
        self.main_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.path_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Source Path...", height=32, width=320)
        self.path_entry.pack(pady=(40, 5))
        self.browse_btn = ctk.CTkButton(self.main_frame, text="BROWSE SOURCE", height=32, command=self.browse_source)
        self.browse_btn.pack(pady=5)

        self.dest_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Destination...", height=32, width=320)
        self.dest_entry.pack(pady=5)
        self.dest_btn = ctk.CTkButton(self.main_frame, text="SELECT DESTINATION", height=32, command=self.browse_dest, fg_color="#2c3e50")
        self.dest_btn.pack(pady=5)

        self.status_label = ctk.CTkLabel(self.main_frame, text="Ready", text_color="#3b8ed0", font=("Arial", 11, "bold"))
        self.status_label.pack(side="bottom", pady=15)

    def add_btn(self, text, command, font):
        btn = ctk.CTkButton(self.sidebar, text=text, corner_radius=6, font=font, height=30,
                            command=lambda: self.start_thread(command))
        btn.pack(pady=3, padx=15, fill="x")

    def start_thread(self, command):
        self.stop_execution = False
        threading.Thread(target=command, daemon=True).start()

    def stop_action(self):
        self.stop_execution = True
        self.update_status("STOPPED", "red")

    def browse_source(self):
        path = filedialog.askopenfilename() or filedialog.askdirectory()
        if path:
            self.input_path = path
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, path)

    def browse_dest(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir = path
            self.dest_entry.delete(0, 'end')
            self.dest_entry.insert(0, path)

    def get_save_path(self, default_name):
        if self.output_dir:
            return os.path.join(self.output_dir, default_name)
        if self.input_path:
            base_dir = os.path.dirname(self.input_path) if os.path.isfile(self.input_path) else self.input_path
            return os.path.join(base_dir, default_name)
        return default_name

    @staticmethod
    def fix_arabic_text(text):
        if not text.strip(): return ""
        return get_display(arabic_reshaper.reshape(text))

    def images_to_single_excel(self):
        if not self.input_path: return
        self.update_status("Scanning...")
        try:
            img = Image.open(self.input_path)
            raw_text = pytesseract.image_to_string(img, lang='ara+eng')
            fixed_text = self.fix_arabic_text(raw_text)
            rows = [line.split() for line in fixed_text.split('\n') if line.strip()]
            pd.DataFrame(rows).to_excel(self.get_save_path("OCR.xlsx"), index=False, header=False)
            self.update_status("Excel OK", "green")
        except Exception as e: self.update_status(f"Error: {e}", "red")

    def merge_pdfs_action(self):
        self.update_status("Select Files...")
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if not files: return
        try:
            writer = PdfWriter()
            for f in files:
                if self.stop_execution: break
                reader = PdfReader(f)
                for page in reader.pages:
                    if self.stop_execution: break
                    writer.add_page(page)
            if not self.stop_execution:
                writer.write(self.get_save_path("Merged.pdf"))
                self.update_status("Merged!", "green")
        except Exception as e: self.update_status("Error", "red")

    def delete_pages_action(self):
        if not self.input_path.lower().endswith('.pdf'): return
        pages = simpledialog.askstring("Del", "Pages (1, 3):")
        if not pages: return
        try:
            to_del = [int(p.strip()) - 1 for p in pages.split(',')]
            reader = PdfReader(self.input_path)
            writer = PdfWriter()
            for i in range(len(reader.pages)):
                if self.stop_execution: break
                if i not in to_del: writer.add_page(reader.pages[i])
            if not self.stop_execution:
                writer.write(self.get_save_path("Edited.pdf"))
                self.update_status("Deleted!", "green")
        except Exception as e: self.update_status("Error", "red")

    def pdf_to_images(self):
        if not self.input_path: return
        self.update_status("Exporting...")
        try:
            doc = pymupdf.open(self.input_path)
            for i, page in enumerate(doc):
                if self.stop_execution: break
                page.get_pixmap(dpi=150).save(self.get_save_path(f"P_{i+1}.jpg"))
            self.update_status("Done!", "green")
        except Exception as e: self.update_status("Error", "red")

    def images_to_single_pdf(self):
        if not self.input_path: return
        self.update_status("Converting...")
        try:
            if os.path.isdir(self.input_path):
                imgs = [os.path.join(self.input_path, f) for f in os.listdir(self.input_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            else: imgs = [self.input_path]
            if imgs and not self.stop_execution:
                with open(self.get_save_path("Result.pdf"), "wb") as f: f.write(img2pdf.convert(imgs))
                self.update_status("PDF OK", "green")
        except Exception as e: self.update_status("Error", "red")

    def images_to_multi_pdf(self):
        if not self.input_path or not os.path.isdir(self.input_path): return
        self.update_status("Processing...")
        try:
            for f in os.listdir(self.input_path):
                if self.stop_execution: break
                if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                    full_p = os.path.join(self.input_path, f)
                    with open(self.get_save_path(f"{f}.pdf"), "wb") as out: out.write(img2pdf.convert(full_p))
            self.update_status("Done!", "green")
        except Exception as e: self.update_status("Error", "red")

    def pdf_to_excel(self):
        if not self.input_path: return
        self.update_status("Tables...")
        try:
            data = []
            with pdfplumber.open(self.input_path) as pdf:
                for page in pdf.pages:
                    if self.stop_execution: break
                    tbl = page.extract_table()
                    if tbl: data.extend(tbl)
            if data and not self.stop_execution:
                pd.DataFrame(data).to_excel(self.get_save_path("Tables.xlsx"), index=False, header=False)
                self.update_status("Excel OK", "green")
        except Exception as e: self.update_status("Error", "red")

    def update_status(self, text, color="#3b8ed0"):
        self.status_label.configure(text=f"STATUS: {text}", text_color=color)

if __name__ == "__main__":
    app = DocSuiteApp()
    app.mainloop()