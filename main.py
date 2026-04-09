import customtkinter as ctk
from tkinter import filedialog
import img2pdf
import os
import fitz  # PyMuPDF
from pathlib import Path
import threading
from pypdf import PdfReader, PdfWriter
import pandas as pd
from PIL import Image, ImageOps, ImageEnhance
import pytesseract

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class DocMaster(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DocMaster 2026 - Final Version")
        self.geometry("450x600")  # صغرت الطول شوية عشان الزرار اللي اتشال
        self.resizable(False, False)

        self.stop_execution = False

        # العنوان
        self.label = ctk.CTkLabel(self, text="DOCMASTER PRO", font=("Impact", 32))
        self.label.pack(pady=15)

        # المدخلات
        self.path_entry = ctk.CTkEntry(self, placeholder_text="Select Folder or File...", width=380, height=35)
        self.path_entry.pack(pady=5)

        self.choice_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.choice_frame.pack(pady=5)
        ctk.CTkButton(self.choice_frame, text="FOLDER", command=self.browse_folder, width=110).grid(row=0, column=0,
                                                                                                    padx=5)
        ctk.CTkButton(self.choice_frame, text="FILE", command=self.browse_file, width=110, fg_color="#2c3e50").grid(
            row=0, column=1, padx=5)

        # التبويبات
        self.tabview = ctk.CTkTabview(self, width=400, height=220)
        self.tabview.pack(pady=10, padx=10)

        self.tab_img = self.tabview.add("Images")
        self.tab_pdf = self.tabview.add("PDF")

        self.setup_tabs_ui()

        # التقدم والحالة
        self.progress_label = ctk.CTkLabel(self, text="Progress: 0%", font=("Arial", 11))
        self.progress_label.pack()
        self.progress_bar = ctk.CTkProgressBar(self, width=350)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(0, 10))

        self.stop_btn = ctk.CTkButton(self, text="STOP", command=self.stop_task, fg_color="#c0392b", width=200,
                                      height=35)
        self.stop_btn.pack(pady=5)

        self.status_label = ctk.CTkLabel(self, text="Status: Ready", text_color="#3b8ed0", font=("Arial", 12, "bold"))
        self.status_label.pack(pady=5)

        # التوقيع
        self.signature = ctk.CTkLabel(self, text="By MOHAMED", font=("Lucida Handwriting", 14, "italic"),
                                      text_color="#d35400")
        self.signature.pack(side="bottom", pady=10)

    def setup_tabs_ui(self):
        # تبويب الصور - الميزات المطلوبة فقط
        ctk.CTkButton(self.tab_img, text="Images to One PDF", command=lambda: self.run_task(self.convert_to_one_pdf),
                      width=280, fg_color="#3b8ed0").pack(pady=8)

        ctk.CTkButton(self.tab_img, text="Each Image to PDF", command=lambda: self.run_task(self.convert_to_multi_pdf),
                      width=280, fg_color="#d35400").pack(pady=8)

        ctk.CTkButton(self.tab_img, text="Images Table to Excel Rows",
                      command=lambda: self.run_task(self.table_to_excel_rows),
                      width=280, fg_color="#1f6e43").pack(pady=8)

        # تبويب الـ PDF - الميزات المطلوبة فقط
        ctk.CTkButton(self.tab_pdf, text="Merge PDFs", command=lambda: self.run_task(self.merge_pdfs),
                      width=280, fg_color="#3b8ed0").pack(pady=10)

        ctk.CTkButton(self.tab_pdf, text="PDF to JPG Images", command=lambda: self.run_task(self.pdf_to_images_logic),
                      width=280, fg_color="#d35400").pack(pady=10)

        ctk.CTkButton(self.tab_pdf, text="PDFs List to Excel",
                      command=lambda: self.run_task(self.export_pdf_list),
                      width=280, fg_color="#1f6e43").pack(pady=10)

    def table_to_excel_rows(self):
        """تحويل محتوى الجدول في الصورة لصفوف إكسيل"""
        path = self.path_entry.get().strip()
        if not os.path.exists(path):
            self.update_status("Select target first", "red")
            return
        try:
            self.update_status("Reading Table Rows...", "orange")
            files = [path] if os.path.isfile(path) else [os.path.join(path, f) for f in os.listdir(path) if
                                                         f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            all_data = []
            for i, f_path in enumerate(files):
                if self.stop_execution: break
                img = ImageOps.grayscale(Image.open(f_path))
                img = ImageEnhance.Contrast(img).enhance(2.5)
                content = pytesseract.image_to_string(img, lang='ara+eng', config='--psm 6')
                for line in content.splitlines():
                    if line.strip():
                        all_data.append({"File": os.path.basename(f_path), "Content": line.strip()})
                self.update_progress(i + 1, len(files))
            df = pd.DataFrame(all_data)
            out_dir = self.get_output_path(os.path.dirname(files[0]) if os.path.isfile(path) else path)
            excel_path = os.path.join(out_dir, "Table_Content.xlsx")
            df.to_excel(excel_path, index=False)
            self.update_status("Excel Ready!", "green")
            os.startfile(excel_path)
        except:
            self.update_status("OCR Error", "red")

    def export_pdf_list(self):
        """عمل قائمة بأسماء ملفات الـ PDF فقط"""
        path = self.path_entry.get().strip()
        if not os.path.isdir(path):
            self.update_status("Select Folder first", "red")
            return
        try:
            self.update_status("Generating PDF List...", "orange")
            files = [{"PDF Name": f} for f in os.listdir(path) if f.lower().endswith(".pdf")]
            df = pd.DataFrame(files)
            out_dir = self.get_output_path(path)
            excel_path = os.path.join(out_dir, "PDF_Files_Report.xlsx")
            df.to_excel(excel_path, index=False)
            self.update_progress(1, 1)
            self.update_status("List Saved!", "green")
        except:
            self.update_status("Error", "red")

    # --- الوظائف الأساسية للمشروع ---
    def stop_task(self):
        self.stop_execution = True

    def update_progress(self, c, t):
        self.progress_bar.set(c / t)
        self.progress_label.configure(text=f"Progress: {int((c / t) * 100)}%")
        self.update()

    def update_status(self, text, color="#3b8ed0"):
        self.status_label.configure(text=f"STATUS: {text}", text_color=color)

    @staticmethod
    def get_output_path(base):
        d = os.path.join(base, "Results")
        if not os.path.exists(d): os.makedirs(d)
        return d

    def run_task(self, func):
        self.stop_execution = False
        self.progress_bar.set(0)
        threading.Thread(target=func, daemon=True).start()

    def browse_folder(self):
        p = filedialog.askdirectory()
        if p: self.path_entry.delete(0, 'end'); self.path_entry.insert(0, os.path.normpath(p))

    def browse_file(self):
        p = filedialog.askopenfilename()
        if p: self.path_entry.delete(0, 'end'); self.path_entry.insert(0, os.path.normpath(p))

    def convert_to_one_pdf(self):
        path = self.path_entry.get().strip()
        if not os.path.isdir(path): return
        try:
            valid = ('.png', '.jpg', '.jpeg')
            imgs = sorted([str(Path(path) / f) for f in os.listdir(path) if f.lower().endswith(valid)])
            if imgs:
                out = self.get_output_path(path)
                with open(os.path.join(out, "Full_Document.pdf"), "wb") as f:
                    f.write(img2pdf.convert(imgs))
                self.update_progress(1, 1)
                self.update_status("One PDF Saved!", "green")
        except:
            self.update_status("Error", "red")

    def convert_to_multi_pdf(self):
        path = self.path_entry.get().strip()
        if not os.path.isdir(path): return
        try:
            out = self.get_output_path(path)
            files = [f for f in os.listdir(path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            for i, f in enumerate(files):
                if self.stop_execution: break
                with open(os.path.join(out, f"{Path(f).stem}.pdf"), "wb") as o:
                    o.write(img2pdf.convert(str(Path(path) / f)))
                self.update_progress(i + 1, len(files))
            self.update_status("PDFs Created!", "green")
        except:
            self.update_status("Error", "red")

    def pdf_to_images_logic(self):
        path = self.path_entry.get().strip()
        if not path: return
        try:
            pdfs = [path] if os.path.isfile(path) else [os.path.join(path, f) for f in os.listdir(path) if
                                                        f.lower().endswith('.pdf')]
            for i, p in enumerate(pdfs):
                if self.stop_execution: break
                doc = fitz.open(p)
                out = self.get_output_path(os.path.dirname(p))
                for j, page in enumerate(doc):
                    page.get_pixmap(dpi=150).save(os.path.join(out, f"{Path(p).stem}_p{j + 1}.jpg"))
                doc.close()
                self.update_progress(i + 1, len(pdfs))
            self.update_status("Images Saved!", "green")
        except:
            self.update_status("Failed", "red")

    def merge_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files: return
        try:
            writer = PdfWriter()
            for i, f in enumerate(files):
                if self.stop_execution: break
                reader = PdfReader(f)
                for page in reader.pages: writer.add_page(page)
                self.update_progress(i + 1, len(files))
            out = self.get_output_path(os.path.dirname(files[0]))
            with open(os.path.join(out, "Merged.pdf"), "wb") as f:
                writer.write(f)
            self.update_status("Merged Successfully!", "green")
        except:
            self.update_status("Error", "red")


if __name__ == "__main__":
    app = DocMaster()
    app.mainloop()