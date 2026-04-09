import customtkinter as ctk
from tkinter import filedialog
import img2pdf
import os
import fitz  # PyMuPDF
from pathlib import Path
import threading
from pypdf import PdfReader, PdfWriter
import pandas as pd

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class DocMaster(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DocMaster 2026")
        self.geometry("450x620")
        self.resizable(False, False)

        self.stop_execution = False

        # العنوان
        self.label = ctk.CTkLabel(self, text="DOCMASTER PRO", font=("Impact", 32))
        self.label.pack(pady=15)

        # خانة المسار
        self.path_entry = ctk.CTkEntry(self, placeholder_text="Select Folder or File...", width=380, height=35)
        self.path_entry.pack(pady=5)

        self.choice_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.choice_frame.pack(pady=5)
        ctk.CTkButton(self.choice_frame, text="FOLDER", command=self.browse_folder, width=110).grid(row=0, column=0,
                                                                                                    padx=5)
        ctk.CTkButton(self.choice_frame, text="FILE", command=self.browse_file, width=110, fg_color="#2c3e50").grid(
            row=0, column=1, padx=5)

        # التبويبات
        self.tabview = ctk.CTkTabview(self, width=400, height=250)
        self.tabview.pack(pady=10, padx=10)

        self.tab_img = self.tabview.add("Images")
        self.tab_pdf = self.tabview.add("PDF")

        self.setup_tabs_ui()

        # شريط التقدم
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

        # الإمضاء باللون البرتقالي كما طلبت
        self.signature = ctk.CTkLabel(self, text="By MOHAMED", font=("Lucida Handwriting", 14, "italic"),
                                      text_color="#d35400")
        self.signature.pack(side="bottom", pady=10)

    def setup_tabs_ui(self):
        # --- تبويب الصور (أزرق، برتقالي، أخضر) ---
        ctk.CTkButton(self.tab_img, text="Images to One PDF", command=lambda: self.run_task(self.convert_to_one_pdf),
                      width=280, fg_color="#3b8ed0").pack(pady=8)

        ctk.CTkButton(self.tab_img, text="Each Image to PDF", command=lambda: self.run_task(self.convert_to_multi_pdf),
                      width=280, fg_color="#d35400").pack(pady=8)

        ctk.CTkButton(self.tab_img, text="Images Info to Excel",
                      command=lambda: self.run_task(lambda: self.export_to_excel(".jpg", ".png", ".jpeg")),
                      width=280, fg_color="#1f6e43").pack(pady=8)

        # --- تبويب الـ PDF (أزرق، برتقالي، أخضر) ---
        ctk.CTkButton(self.tab_pdf, text="Merge PDFs", command=lambda: self.run_task(self.merge_pdfs),
                      width=280, fg_color="#3b8ed0").pack(pady=8)

        ctk.CTkButton(self.tab_pdf, text="PDF to JPG Images", command=lambda: self.run_task(self.pdf_to_images_logic),
                      width=280, fg_color="#d35400").pack(pady=8)

        ctk.CTkButton(self.tab_pdf, text="PDFs List to Excel",
                      command=lambda: self.run_task(lambda: self.export_to_excel(".pdf")),
                      width=280, fg_color="#1f6e43").pack(pady=8)

    def export_to_excel(self, *extensions):
        path = self.path_entry.get().strip()
        if not os.path.isdir(path):
            self.update_status("Select a Folder first", "red")
            return
        try:
            self.update_status("Creating Excel...", "orange")
            files_data = []
            for f in os.listdir(path):
                ext = Path(f).suffix.lower()
                if ext in extensions or not extensions:
                    f_path = os.path.join(path, f)
                    if os.path.isfile(f_path):
                        size = os.path.getsize(f_path) / 1024
                        files_data.append({"File Name": str(f), "Type": ext, "Size (KB)": round(size, 2)})

            if not files_data:
                self.update_status("No matching files!", "red")
                return

            df = pd.DataFrame(files_data)
            out_dir = self.get_output_path(path)
            report_name = "Images_Report.xlsx" if ".jpg" in extensions else "PDF_Report.xlsx"
            excel_path = os.path.join(out_dir, report_name)

            # الكتابة بمحرك openpyxl لضمان سلامة اللغة العربية
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            self.update_progress(1, 1)
            self.update_status(f"Done: {report_name}", "green")
        except:
            self.update_status("Excel Error", "red")

    def stop_task(self):
        self.stop_execution = True
        self.update_status("Stopping...", "orange")

    def update_progress(self, current, total):
        percent = (current / total)
        self.progress_bar.set(percent)
        self.progress_label.configure(text=f"Progress: {int(percent * 100)}%")
        self.update()

    @staticmethod
    def get_output_path(base_folder):
        result_dir = os.path.join(base_folder, "Results")
        if not os.path.exists(result_dir): os.makedirs(result_dir)
        return result_dir

    def run_task(self, task_function):
        self.stop_execution = False
        self.progress_bar.set(0)
        threading.Thread(target=task_function, daemon=True).start()

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, os.path.normpath(path))

    def browse_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.path_entry.delete(0, 'end')
            self.path_entry.insert(0, os.path.normpath(path))

    def update_status(self, text, color="#3b8ed0"):
        self.status_label.configure(text=f"STATUS: {text}", text_color=color)

    def convert_to_one_pdf(self):
        path = self.path_entry.get().strip()
        if not os.path.isdir(path): return
        try:
            valid_exts = ('.png', '.jpg', '.jpeg')
            imgs = sorted([str(Path(path) / f) for f in os.listdir(path) if f.lower().endswith(valid_exts)])
            if imgs:
                out_dir = self.get_output_path(path)
                with open(os.path.join(out_dir, "Combined_Images.pdf"), "wb") as f:
                    f.write(img2pdf.convert(imgs))
                self.update_progress(1, 1)
                self.update_status("Saved in Results", "green")
        except:
            self.update_status("Error", "red")

    def convert_to_multi_pdf(self):
        path = self.path_entry.get().strip()
        if not os.path.isdir(path): return
        try:
            out_dir = self.get_output_path(path)
            files = [f for f in os.listdir(path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            total = len(files)
            for i, f in enumerate(files):
                if self.stop_execution: break
                img_p = str(Path(path) / f)
                with open(os.path.join(out_dir, f"{Path(f).stem}.pdf"), "wb") as out:
                    out.write(img2pdf.convert(img_p))
                self.update_progress(i + 1, total)
            self.update_status("Success!", "green")
        except:
            self.update_status("Error", "red")

    def pdf_to_images_logic(self):
        path = self.path_entry.get().strip()
        if os.path.isfile(path) and path.lower().endswith('.pdf'):
            self.process_single_pdf(path)
            self.update_status("Done!", "green")
        elif os.path.isdir(path):
            pdfs = [str(Path(path) / f) for f in os.listdir(path) if f.lower().endswith('.pdf')]
            total = len(pdfs)
            for i, p in enumerate(pdfs):
                if self.stop_execution: break
                self.process_single_pdf(p)
                self.update_progress(i + 1, total)
            self.update_status("Folder Done!", "green")

    def process_single_pdf(self, file_path):
        try:
            doc = fitz.open(file_path)
            base_name = Path(file_path).stem
            out_dir = self.get_output_path(Path(file_path).parent)
            for i, page in enumerate(doc):
                if self.stop_execution: break
                pix = page.get_pixmap(dpi=150)
                suffix = f"_{i + 1}" if len(doc) > 1 else ""
                pix.save(os.path.join(out_dir, f"{base_name}{suffix}.jpg"))
            doc.close()
        except:
            pass

    def merge_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files: return
        try:
            writer = PdfWriter()
            total = len(files)
            for i, f in enumerate(files):
                if self.stop_execution: break
                reader = PdfReader(f)
                for page in reader.pages: writer.add_page(page)
                self.update_progress(i + 1, total)
            out_dir = self.get_output_path(os.path.dirname(files[0]))
            with open(os.path.join(out_dir, "Merged.pdf"), "wb") as f:
                writer.write(f)
            self.update_status("Merged!", "green")
        except:
            self.update_status("Failed", "red")


if __name__ == "__main__":
    app = DocMaster()
    app.mainloop()