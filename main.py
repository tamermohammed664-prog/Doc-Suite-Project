import os
import queue
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk
import arabic_reshaper
from bidi.algorithm import get_display

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None


def fix_ar(text):
    if text is None:
        return ""
    text = str(text)
    if not text:
        return ""

    parts = []
    for line in text.splitlines():
        if any("\u0600" <= ch <= "\u06FF" for ch in line):
            parts.append(get_display(arabic_reshaper.reshape(line)))
        else:
            parts.append(line)
    return "\n".join(parts)


ar = fix_ar


def _apply_arabic_text(value):
    if value is None:
        return ""
    value = str(value)
    if any("\u0600" <= ch <= "\u06FF" for ch in value):
        return fix_ar(value)
    return value


class ArLabel(ctk.CTkLabel):
    def __init__(self, *args, **kwargs):
        if "text" in kwargs:
            kwargs["text"] = _apply_arabic_text(kwargs["text"])
        super().__init__(*args, **kwargs)

    def configure(self, *args, **kwargs):
        if "text" in kwargs:
            kwargs["text"] = _apply_arabic_text(kwargs["text"])
        return super().configure(*args, **kwargs)


class ArButton(ctk.CTkButton):
    def __init__(self, *args, **kwargs):
        if "text" in kwargs:
            kwargs["text"] = _apply_arabic_text(kwargs["text"])
        super().__init__(*args, **kwargs)

    def configure(self, *args, **kwargs):
        if "text" in kwargs:
            kwargs["text"] = _apply_arabic_text(kwargs["text"])
        return super().configure(*args, **kwargs)


class _NoDnDWrapper:
    pass


_DnDWrapperBase = TkinterDnD.DnDWrapper if TkinterDnD is not None else _NoDnDWrapper


class ArEntry(ctk.CTkEntry, _DnDWrapperBase):
    def __init__(self, *args, **kwargs):
        if "placeholder_text" in kwargs:
            kwargs["placeholder_text"] = _apply_arabic_text(kwargs["placeholder_text"])
        super().__init__(*args, **kwargs)

    def configure(self, *args, **kwargs):
        if "placeholder_text" in kwargs:
            kwargs["placeholder_text"] = _apply_arabic_text(kwargs["placeholder_text"])
        return super().configure(*args, **kwargs)


class _CTkDropRoot(ctk.CTk, _DnDWrapperBase):
    pass

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def enable_windows_dpi_awareness():
    if os.name != "nt":
        return

    try:
        import ctypes

        try:
            if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
                return
        except (AttributeError, OSError):
            pass

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def _runtime_roots():
    roots = []
    bundled_root = getattr(sys, "_MEIPASS", None)
    if bundled_root:
        roots.append(Path(bundled_root))
    roots.append(Path(__file__).resolve().parent)
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent)
    return list(dict.fromkeys(roots))


def resolve_portable_binary_paths(roots=None):
    roots = [Path(root) for root in (roots or _runtime_roots())]
    tesseract_candidates = (
        Path("bin/tesseract/tesseract.exe"),
        Path("bin/tesseract/tesseract"),
        Path("bin/Tesseract-OCR/tesseract.exe"),
        Path("bin/tesseract.exe"),
        Path("bin/tesseract"),
    )
    poppler_candidates = (
        Path("bin/poppler/bin"),
        Path("bin/poppler/Library/bin"),
        Path("bin/poppler"),
    )

    tesseract_path = next(
        (root / candidate for root in roots for candidate in tesseract_candidates if (root / candidate).is_file()),
        None,
    )
    poppler_path = next(
        (
            root / candidate
            for root in roots
            for candidate in poppler_candidates
            if (root / candidate).is_dir()
            and any((root / candidate / executable).is_file() for executable in ("pdfinfo.exe", "pdfinfo"))
        ),
        None,
    )
    return tesseract_path, poppler_path


def configure_portable_tesseract(pytesseract):
    tesseract_path, _ = resolve_portable_binary_paths()
    if tesseract_path is None:
        return None

    tesseract_path = tesseract_path.resolve()
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)
    executable_dir = tesseract_path.parent
    path_entries = os.environ.get("PATH", "").split(os.pathsep)
    path_compare = str(executable_dir).casefold() if os.name == "nt" else str(executable_dir)
    if not any(
        (entry.casefold() if os.name == "nt" else entry) == path_compare
        for entry in path_entries
    ):
        os.environ["PATH"] = str(executable_dir) + os.pathsep + os.environ.get("PATH", "")
    tessdata_dir = executable_dir / "tessdata"
    if tessdata_dir.is_dir():
        os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
    return tesseract_path


def portable_poppler_path():
    _, poppler_path = resolve_portable_binary_paths()
    return str(poppler_path.resolve()) if poppler_path else None


IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
OCR_LANGUAGES = "eng+ara"

STATUS_TEXT = {
    "app_title": "DocNexus",
    "panel_title": "Master Control Panel",
    "path_placeholder": "Select folder or file path...",
    "output_placeholder": "Output folder (optional)",
    "browse": "BROWSE",
    "output": "OUTPUT",
    "ready": "Ready",
    "status_select_valid": "Please select a valid image folder or file.",
    "status_select_valid_pdf": "Please select a valid PDF file.",
    "status_select_valid_folder": "Please select a valid image folder.",
    "status_merge": "Merging images to PDF...",
    "status_done_pdf": "Done: PDF created successfully!",
    "status_convert_each": "Converting each image to PDF...",
    "status_done_multi_pdf": "Done: {count} PDF files created!",
    "status_ocr": "OCR in progress...",
    "status_ocr_multi": "OCR in progress (multiple files)...",
    "status_ocr_done": "Excel file generated successfully!",
    "status_ocr_done_multi": "Done: {count} Excel files created!",
    "status_pdf_images": "Extracting PDF pages to images...",
    "status_pdf_images_done": "Images extracted successfully!",
    "status_pdf_table": "Extracting table data...",
    "status_pdf_table_done": "Excel file created successfully!",
    "status_error": "Error: {error}",
    "status_ocr_error": "OCR failed: {error}",
    "status_pdf_error": "Failed to convert PDF: {error}",
    "status_table_error": "Failed to extract tables: {error}",
}

STATUS_ARABIC = {
    "status_select_valid": "يرجى اختيار مجلد صور أو ملف صالح.",
    "status_select_valid_pdf": "يرجى اختيار ملف PDF صالح.",
    "status_select_valid_folder": "يرجى اختيار مجلد صور صالح.",
    "status_merge": "جارٍ دمج الصور في مستند PDF...",
    "status_done_pdf": "تم إنشاء ملف PDF بنجاح.",
    "status_convert_each": "جارٍ تحويل كل صورة إلى ملف PDF...",
    "status_done_multi_pdf": "تم إنشاء {count} ملفات PDF.",
    "status_ocr": "جارٍ التعرف الضوئي على النصوص...",
    "status_ocr_multi": "جارٍ التعرف الضوئي على عدة ملفات...",
    "status_ocr_done": "تم إنشاء ملف Excel بنجاح.",
    "status_ocr_done_multi": "تم إنشاء {count} ملفات Excel.",
    "status_pdf_images": "جارٍ استخراج صفحات PDF إلى صور...",
    "status_pdf_images_done": "تم استخراج الصور بنجاح.",
    "status_pdf_table": "جارٍ استخراج بيانات الجداول...",
    "status_pdf_table_done": "تم إنشاء ملف Excel بنجاح.",
    "status_error": "حدث خطأ: {error}",
    "status_ocr_error": "فشل التعرف الضوئي على النصوص: {error}",
    "status_pdf_error": "تعذر تحويل ملف PDF: {error}",
    "status_table_error": "تعذر استخراج الجداول: {error}",
}


class DocNexusApp(_CTkDropRoot):
    def __init__(self):
        super().__init__()

        self.title("DocNexus")
        self.geometry("1200x760")
        self.minsize(980, 680)
        self._cancel_event = threading.Event()
        self._worker_queue = queue.Queue()
        self._progress_lock = threading.Lock()
        self._pending_progress = None
        self._ui_thread_id = threading.get_ident()
        self._overwrite_decision_event = threading.Event()
        self._overwrite_choice = None
        self._overwrite_dialog = None
        self._output_mode = "default"
        self._output_path_claims = set()
        self._task_running = False
        self._task_started_at = 0.0
        self._source_path = ""
        self._output_path = ""
        self.selected_action = None

        self.sidebar = ctk.CTkFrame(self, width=260, corner_radius=0, fg_color="#141b2d")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.logo = ArLabel(
            self.sidebar,
            text="DocNexus",
            font=ctk.CTkFont(family="Impact", size=28),
            text_color="#dbe8ff",
        )
        self.logo.pack(pady=(24, 20))

        btn_font = ctk.CTkFont(size=12, weight="bold")
        sub_font = ctk.CTkFont(size=10)

        self.sidebar_buttons = {}
        self.add_sidebar_group("PDF TOOLS", "أدوات PDF")
        self.sidebar_buttons["images_single_pdf"] = self.add_bilingual_button("Images to One PDF", "صور إلى مستند واحد", self.images_to_single_pdf, btn_font, sub_font)
        self.sidebar_buttons["images_multi_pdf"] = self.add_bilingual_button("Each Image to PDF", "كل صورة إلى مستند", self.images_to_multi_pdf, btn_font, sub_font)
        self.sidebar_buttons["pdf_images"] = self.add_bilingual_button("PDF to Images", "تحويل المستند إلى صور", self.pdf_to_images, btn_font, sub_font)
        self.add_sidebar_group("EXCEL TOOLS", "أدوات Excel")
        self.sidebar_buttons["images_single_excel"] = self.add_bilingual_button("Images to One Excel", "صور إلى جدول بيانات", self.images_to_single_excel, btn_font, sub_font)
        self.sidebar_buttons["images_multi_excel"] = self.add_bilingual_button("Each Image to Excel", "كل صورة إلى جدول", self.images_to_multi_excel, btn_font, sub_font)
        self.sidebar_buttons["pdf_excel"] = self.add_bilingual_button("PDF to Excel", "تحويل المستند إلى جدول", self.pdf_to_excel, btn_font, sub_font)

        self.signature = ArLabel(
            self.sidebar,
            text="DocNexus",
            font=ctk.CTkFont(family="Lucida Handwriting", size=22, slant="italic"),
            text_color="#60a5fa",
        )
        self.signature.pack(side="bottom", pady=20)

        self.main_frame = ctk.CTkFrame(self, fg_color="#111827", corner_radius=18)
        self.main_frame.pack(side="right", fill="both", expand=True, padx=22, pady=22)

        self.label = ArLabel(
            self.main_frame,
            text="Master Control Panel\nلوحة التحكم الرئيسية",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        self.label.pack(pady=(24, 8))

        self.selection_label = ArLabel(
            self.main_frame,
            text="Choose an action from the sidebar\nاختر إجراءً من القائمة الجانبية",
            text_color="#a8b6cb",
            font=ctk.CTkFont(size=13),
        )
        self.selection_label.pack(pady=(0, 20))

        self.path_label = ArLabel(
            self.main_frame,
            text="Input file or folder\nملف الإدخال أو المجلد",
            anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.path_label.pack(fill="x", padx=42)

        input_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        input_row.pack(fill="x", padx=42, pady=(4, 14))
        input_row.grid_columnconfigure(0, weight=1)

        self.path_entry = ArEntry(
            input_row,
            placeholder_text="Select a file or folder...",
            height=46,
            corner_radius=9,
            border_color="#60a5fa",
            fg_color="#1e293b",
        )
        self.path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._setup_drag_and_drop()

        self.drag_drop_hint = ArLabel(
            input_row,
            text=self._drag_drop_hint_text(),
            text_color="#8fa4c2",
            font=ctk.CTkFont(size=11),
            anchor="w",
        )
        self.drag_drop_hint.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 0))

        self.browse_btn = ArButton(
            input_row,
            text="Browse\nاستعراض",
            font=btn_font,
            command=self.browse,
            fg_color="#3b82f6",
            hover_color="#1d4ed8",
            height=46,
            width=148,
            corner_radius=9,
        )
        self.browse_btn.grid(row=0, column=1)

        self.output_label = ArLabel(
            self.main_frame,
            text="Output folder (optional)\nمجلد الإخراج (اختياري)",
            anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.output_label.pack(fill="x", padx=42)

        output_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        output_row.pack(fill="x", padx=42, pady=(4, 20))
        output_row.grid_columnconfigure(0, weight=1)

        self.output_entry = ArEntry(
            output_row,
            placeholder_text="Choose an output folder...",
            height=46,
            corner_radius=9,
            border_color="#3b82f6",
            fg_color="#1e293b",
        )
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.output_btn = ArButton(
            output_row,
            text="Choose Folder\nاختيار مجلد",
            font=btn_font,
            command=self.choose_output_directory,
            fg_color="#334155",
            hover_color="#475569",
            height=46,
            width=148,
            corner_radius=9,
        )
        self.output_btn.grid(row=0, column=1, padx=(0, 8))

        self.open_output_btn = ArButton(
            output_row,
            text="Open Output\nفتح مجلد الإخراج",
            font=btn_font,
            fg_color="#1f6feb",
            hover_color="#1d4ed8",
            command=self.open_output_folder,
            height=46,
            width=160,
            corner_radius=9,
        )
        self.open_output_btn.grid(row=0, column=2)

        self.progress_frame = ctk.CTkFrame(self.main_frame, fg_color="#182235", corner_radius=12)
        self.progress_frame.pack(fill="x", padx=42, pady=(0, 20))
        self.progress_heading = ArLabel(
            self.progress_frame,
            text="Processing Progress\nتقدم المعالجة",
            anchor="w",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.progress_heading.pack(fill="x", padx=18, pady=(15, 8))

        progress_row = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        progress_row.pack(fill="x", padx=18)
        progress_row.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(progress_row, height=12, corner_radius=6)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 14), pady=5)
        self.progress_percent = ArLabel(
            progress_row,
            text="0%",
            width=54,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.progress_percent.grid(row=0, column=1)

        self.progress_meta = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        self.progress_meta.pack(fill="x", padx=18, pady=(4, 2))
        self.elapsed_label = ArLabel(
            self.progress_meta,
            text="Elapsed: 00:00\nالوقت المنقضي: 00:00",
            anchor="w",
            text_color="#a8b6cb",
        )
        self.elapsed_label.pack(side="left")
        self.current_file_label = ArLabel(
            self.progress_frame,
            text="Current file: None\nالملف الحالي: لا يوجد",
            anchor="w",
            text_color="#a8b6cb",
            wraplength=650,
        )
        self.current_file_label.pack(fill="x", padx=18, pady=(2, 14))

        process_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        process_row.pack(pady=(0, 10))
        self.start_btn = ArButton(
            process_row,
            text="Start Processing\nبدء المعالجة",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.start_processing,
            fg_color="#16a36a",
            hover_color="#128253",
            height=52,
            width=220,
            corner_radius=9,
        )
        self.start_btn.grid(row=0, column=0, padx=6)
        self.cancel_btn = ArButton(
            process_row,
            text="Cancel Process\nإلغاء المعالجة",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.cancel_processing,
            fg_color="#9f3434",
            hover_color="#7f2929",
            height=52,
            width=220,
            corner_radius=9,
            state="disabled",
        )
        self.cancel_btn.grid(row=0, column=1, padx=6)

        self.status_label = ArLabel(
            self.main_frame,
            text="Ready\nجاهز",
            text_color="#60a5fa",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.status_label.pack(side="bottom", pady=(6, 14))
        self.after(50, self._drain_worker_queue)

    def _setup_drag_and_drop(self):
        self._drag_drop_enabled = False
        if TkinterDnD is None or DND_FILES is None:
            return

        try:
            self.TkdndVersion = TkinterDnD._require(self)
            self.path_entry.drop_target_register(DND_FILES)
            self.path_entry.dnd_bind("<<Drop>>", self._handle_path_drop)
            self._drag_drop_enabled = True
        except Exception:
            self._drag_drop_enabled = False

    def _drag_drop_hint_text(self):
        if self._drag_drop_enabled:
            return "Drop one file or folder here, or browse\nأفلت ملفاً أو مجلداً هنا أو استعرض"
        return "Drag-and-drop unavailable; use Browse\nالسحب والإفلات غير متاح؛ استخدم استعراض"

    def _handle_path_drop(self, event):
        try:
            dropped_paths = self.tk.splitlist(event.data)
        except (AttributeError, tk.TclError):
            dropped_paths = ()

        if len(dropped_paths) != 1:
            self.update_status(
                "Drop one file or folder at a time.\nأفلت ملفاً أو مجلداً واحداً في كل مرة.",
                "#f59e0b",
            )
            return "break"

        dropped_path = Path(dropped_paths[0])
        if not dropped_path.exists():
            self.update_status(
                "The dropped path does not exist.\nالمسار المُفلت غير موجود.",
                "red",
            )
            return "break"
        if dropped_path.is_file() and dropped_path.suffix.lower() not in IMAGE_EXTENSIONS + (".pdf",):
            self.update_status(
                "Drop an image, PDF, or folder.\nأفلت صورة أو ملف PDF أو مجلداً.",
                "#f59e0b",
            )
            return "break"

        self.path_entry.delete(0, "end")
        self.path_entry.insert(0, str(dropped_path))
        self.update_status(
            f"Selected: {dropped_path.name}\nتم الاختيار: {dropped_path.name}",
            "#60a5fa",
        )
        return "copy"

    def add_sidebar_group(self, english_text, arabic_text):
        label = ArLabel(
            self.sidebar,
            text=f"{english_text}\n{arabic_text}",
            anchor="w",
            text_color="#8fa4c2",
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        label.pack(fill="x", padx=18, pady=(12, 1))

    def add_bilingual_button(self, english_text, arabic_text, command, font, sub_font):
        frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        button = ArButton(
            frame,
            text=english_text,
            corner_radius=9,
            font=font,
            height=38,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=lambda: self.select_action(command, english_text, arabic_text),
        )
        button.pack(fill="x", pady=(0, 2))

        subtitle = ArLabel(
            frame,
            text=arabic_text,
            font=sub_font,
            text_color="#b5c7ff",
        )
        subtitle.pack(fill="x", pady=(0, 6))

        frame.pack(pady=2, padx=18, fill="x")
        frame.action_button = button
        return frame

    def select_action(self, command, english_text, arabic_text):
        self.selected_action = command
        for frame in self.sidebar_buttons.values():
            frame.action_button.configure(fg_color="#2563eb")
        for frame in self.sidebar_buttons.values():
            if frame.action_button.cget("text") == english_text:
                frame.action_button.configure(fg_color="#0f8a70")
                break
        self.selection_label.configure(text=f"{english_text}\n{arabic_text}")
        self.current_file_label.configure(text=f"Selected action: {english_text}\nالإجراء المحدد")

    def start_processing(self):
        if self._task_running:
            return
        if self.selected_action is None:
            self.update_status("Select an action first.\nاختر إجراءً أولاً.", "#f59e0b")
            return

        action = self.selected_action
        self._source_path = self.path_entry.get().strip()
        self._output_path = self.output_entry.get().strip()
        self._output_path_claims.clear()
        self._cancel_event.clear()
        self._task_running = True
        self._task_started_at = time.monotonic()
        self.progress_bar.set(0)
        self.progress_percent.configure(text="0%")
        self.current_file_label.configure(text="Starting...\nجارٍ البدء...")
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        for frame in self.sidebar_buttons.values():
            frame.action_button.configure(state="disabled")
        self._tick_elapsed()
        try:
            threading.Thread(target=self._run_selected_action, args=(action,), daemon=True).start()
        except RuntimeError as exc:
            self._finish_task()
            self.update_status(f"Could not start processing: {exc}\nتعذر بدء المعالجة", "red")

    def _run_selected_action(self, action):
        try:
            self.update_status(
                "Checking destination files...\nجارٍ التحقق من ملفات الوجهة...",
                "#60a5fa",
            )
            existing_outputs = self.find_existing_outputs(action)
            if self.is_cancelled():
                return
            self._output_mode = "default"
            if existing_outputs:
                self._overwrite_choice = None
                self._overwrite_decision_event.clear()
                self._worker_queue.put(("confirm_overwrite", existing_outputs))
                while not self._overwrite_decision_event.wait(0.1):
                    if self.is_cancelled():
                        return
                if self._overwrite_choice == "cancel":
                    self._cancel_event.set()
                    return
                self._output_mode = self._overwrite_choice

            if self.is_cancelled():
                return
            action()
        except Exception as exc:
            self.update_status(f"Error: {exc}\nحدث خطأ أثناء المعالجة: {exc}", "red")
        finally:
            self._worker_queue.put(("finished", None))

    def _drain_worker_queue(self):
        try:
            for _ in range(100):
                try:
                    event, payload = self._worker_queue.get_nowait()
                except queue.Empty:
                    break
                try:
                    if event == "status":
                        text, color = payload
                        self.status_label.configure(text=text, text_color=color)
                    elif event == "confirm_overwrite":
                        self._show_overwrite_dialog(payload)
                    elif event == "finished":
                        self._finish_task()
                except Exception as exc:
                    if event == "confirm_overwrite":
                        self._resolve_overwrite_choice("cancel")
                    elif event == "finished":
                        self._finish_task()
                    try:
                        self.status_label.configure(
                            text=f"UI update failed: {exc}\nتعذر تحديث الواجهة",
                            text_color="red",
                        )
                    except tk.TclError:
                        pass
        finally:
            with self._progress_lock:
                pending_progress = self._pending_progress
                self._pending_progress = None
            try:
                if pending_progress is not None and self._task_running:
                    progress, current_file = pending_progress
                    self.progress_bar.set(progress)
                    self.progress_percent.configure(text=f"{round(progress * 100)}%")
                    name = Path(current_file).name if current_file else "None"
                    self.current_file_label.configure(
                        text=f"Current file: {name}\nالملف الحالي: {name}"
                    )
            except Exception as exc:
                try:
                    self.status_label.configure(
                        text=f"Progress update failed: {exc}\nتعذر تحديث التقدم",
                        text_color="red",
                    )
                except tk.TclError:
                    pass
            try:
                self.after(50, self._drain_worker_queue)
            except tk.TclError:
                pass

    def _finish_task(self):
        self._task_running = False
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        for frame in self.sidebar_buttons.values():
            frame.action_button.configure(state="normal")
        if self._cancel_event.is_set():
            self.update_status("Canceled safely.\nتم الإلغاء بأمان.", "#f59e0b")
            self.current_file_label.configure(text="Canceled\nتم الإلغاء")
        else:
            self.current_file_label.configure(text="Finished\nاكتملت المعالجة")

    def _tick_elapsed(self):
        if not self._task_running:
            return
        elapsed = int(time.monotonic() - self._task_started_at)
        minutes, seconds = divmod(elapsed, 60)
        self.elapsed_label.configure(
            text=f"Elapsed: {minutes:02d}:{seconds:02d}\nالوقت المنقضي: {minutes:02d}:{seconds:02d}"
        )
        self.after(250, self._tick_elapsed)

    def report_progress(self, completed, total, current_file=""):
        progress = min(completed / total, 1.0) if total else 0.0
        with self._progress_lock:
            self._pending_progress = (progress, str(current_file))

    def _show_overwrite_dialog(self, existing_outputs):
        if self._overwrite_dialog is not None:
            return

        dialog = ctk.CTkToplevel(self)
        self._overwrite_dialog = dialog
        dialog.title(self.normalize_arabic_text("File Already Exists / الملف موجود بالفعل"))
        dialog.geometry("590x330")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        title = ArLabel(
            dialog,
            text="File Already Exists\nالملف موجود بالفعل",
            font=ctk.CTkFont(size=19, weight="bold"),
        )
        title.pack(padx=24, pady=(22, 12))

        message = ArLabel(
            dialog,
            text=(
                "A file with this name already exists in the destination folder. "
                "Do you want to overwrite it or create a new version?\n"
                "يوجد ملف بنفس الاسم في مجلد الإخراج. هل تريد استبداله أم إنشاء نسخة جديدة؟"
            ),
            wraplength=530,
            justify="center",
            font=ctk.CTkFont(size=13),
        )
        message.pack(fill="x", padx=26, pady=(0, 8))

        if existing_outputs is not None:
            count_label = ArLabel(
                dialog,
                text=f"Existing file: {existing_outputs.name}\n"
                f"الملف الموجود: {existing_outputs.name}",
                text_color="#a8b6cb",
                font=ctk.CTkFont(size=11),
            )
            count_label.pack(pady=(0, 8))

        buttons = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons.pack(side="bottom", pady=(10, 22))

        ArButton(
            buttons,
            text="Overwrite\nاستبدال",
            command=lambda: self._resolve_overwrite_choice("overwrite"),
            fg_color="#b45309",
            hover_color="#92400e",
            width=150,
            height=50,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, padx=6)
        ArButton(
            buttons,
            text="Create New Version\nإنشاء نسخة جديدة",
            command=lambda: self._resolve_overwrite_choice("new_version"),
            fg_color="#16805d",
            hover_color="#116849",
            width=190,
            height=50,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=1, padx=6)
        ArButton(
            buttons,
            text="Cancel\nإلغاء",
            command=lambda: self._resolve_overwrite_choice("cancel"),
            fg_color="#475569",
            hover_color="#334155",
            width=130,
            height=50,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=2, padx=6)

        dialog.protocol("WM_DELETE_WINDOW", lambda: self._resolve_overwrite_choice("cancel"))
        dialog.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - dialog.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - dialog.winfo_height()) // 2)
        dialog.geometry(f"+{x}+{y}")
        dialog.focus_force()

    def _resolve_overwrite_choice(self, choice):
        dialog = self._overwrite_dialog
        self._overwrite_dialog = None
        self._overwrite_choice = choice
        if dialog is not None and dialog.winfo_exists():
            dialog.grab_release()
            dialog.destroy()
        self._overwrite_decision_event.set()

    def is_cancelled(self):
        return self._cancel_event.is_set()

    def cancel_processing(self):
        if self._task_running:
            self._cancel_event.set()
            self.cancel_btn.configure(state="disabled")
            self.update_status(
                "Cancel requested; finishing current item...\nتم طلب الإلغاء؛ جارٍ إنهاء العنصر الحالي...",
                "#f59e0b",
            )

    def open_output_folder(self):
        folder_text = self.output_entry.get().strip()
        if folder_text:
            folder = Path(folder_text)
        else:
            source_text = self.path_entry.get().strip()
            if not source_text:
                folder = None
            else:
                source = Path(source_text)
                folder = source if source.is_dir() else source.parent
        if folder is None or not folder.is_dir():
            self.update_status(
                "Choose an existing output folder first.\nاختر مجلد إخراج موجوداً أولاً.",
                "#f59e0b",
            )
            return
        try:
            if os.name == "nt":
                os.startfile(str(folder))
            elif os.name == "posix":
                subprocess.Popen(["xdg-open", str(folder)])
            else:
                self.update_status("Opening folders is not supported.\nفتح المجلدات غير مدعوم.", "red")
        except Exception as exc:
            self.update_status(f"Could not open folder: {exc}\nتعذر فتح المجلد", "red")

    def normalize_arabic_text(self, value):
        if value is None:
            return ""

        text = str(value).strip()
        if not text:
            return text

        if any("\u0600" <= ch <= "\u06FF" for ch in text):
            try:
                return fix_ar(text)
            except Exception:
                return text
        return text

    def t(self, key, **kwargs):
        english = STATUS_TEXT.get(key, key)
        arabic = STATUS_ARABIC.get(key)
        if kwargs:
            english = english.format(**kwargs)
            if arabic:
                arabic = arabic.format(**kwargs)
        return f"{english}\n{arabic}" if arabic else english

    def browse(self):
        file_path = filedialog.askopenfilename(title="Select an image or PDF file")
        if file_path:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, file_path)
            return

        directory = filedialog.askdirectory(title="Select a folder")
        if directory:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, directory)

    def choose_output_directory(self):
        output_dir = filedialog.askdirectory(title="Choose output folder")
        if output_dir:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, output_dir)

    def update_status(self, text, color="#60a5fa"):
        if threading.get_ident() == self._ui_thread_id:
            self.status_label.configure(text=text, text_color=color)
        else:
            self._worker_queue.put(("status", (text, color)))

    def show_error(self, key, error):
        self.update_status(self.t(key, error=error), "red")

    def resolve_output_dir(self, source_path):
        source_path = Path(source_path)
        custom_dir = self._output_path
        if custom_dir:
            output_dir = Path(custom_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            return output_dir

        if source_path.is_dir():
            return source_path

        return source_path.parent if str(source_path.parent) else Path.cwd()

    def unique_output_path(self, output_dir, stem, suffix):
        output_dir = Path(output_dir)
        claimed_paths = getattr(self, "_output_path_claims", set())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        output_path = output_dir / f"{stem}_{timestamp}{suffix}"
        counter = 1
        while output_path.exists() or output_path in claimed_paths:
            output_path = output_dir / f"{stem}_{timestamp}_{counter}{suffix}"
            counter += 1
        claimed_paths.add(output_path)
        return output_path

    def get_output_path(self, output_dir, stem, suffix):
        default_path = Path(output_dir) / f"{stem}{suffix}"
        if self._output_mode == "overwrite":
            if default_path in self._output_path_claims:
                return self.unique_output_path(output_dir, stem, suffix)
            self._output_path_claims.add(default_path)
            return default_path

        if self._output_mode == "new_version" or default_path.exists() or default_path in self._output_path_claims:
            return self.unique_output_path(output_dir, stem, suffix)

        self._output_path_claims.add(default_path)
        return default_path

    def find_existing_outputs(self, action):
        if not self._source_path:
            return None

        source_path = Path(self._source_path)
        action_name = action.__name__

        image_actions = {
            "images_to_single_pdf",
            "images_to_multi_pdf",
            "images_to_single_excel",
            "images_to_multi_excel",
        }
        pdf_actions = {"pdf_to_images", "pdf_to_excel"}
        if action_name in image_actions:
            image_files = self.get_image_files(source_path)
            if not image_files:
                return None
        elif action_name in pdf_actions and (
            not source_path.is_file() or source_path.suffix.lower() != ".pdf"
        ):
            return None
        else:
            image_files = []

        output_dir = self.resolve_output_dir(source_path)

        if action_name == "images_to_single_pdf":
            planned_paths = (output_dir / "Merged_Result.pdf",)
        elif action_name == "images_to_multi_pdf":
            planned_paths = (
                output_dir / f"{image_path.stem}.pdf"
                for image_path in image_files
            )
        elif action_name == "images_to_single_excel":
            planned_paths = (output_dir / "Images_Collection.xlsx",)
        elif action_name == "images_to_multi_excel":
            planned_paths = (
                output_dir / f"{image_path.stem}.xlsx"
                for image_path in image_files
            )
        elif action_name == "pdf_to_images":
            from pdf2image import pdfinfo_from_path

            page_count = int(
                pdfinfo_from_path(
                    str(source_path),
                    poppler_path=portable_poppler_path(),
                ).get("Pages", 0)
            )
            planned_paths = (output_dir / f"Page_{index}.jpg" for index in range(1, page_count + 1))
        elif action_name == "pdf_to_excel":
            planned_paths = (output_dir / f"{source_path.stem}.xlsx",)
        else:
            return None

        return next((path for path in planned_paths if path.exists()), None)

    def get_image_files(self, path):
        path = Path(path) if path else None
        if path is None or not path.exists():
            return []

        if path.is_dir():
            files = [
                entry
                for entry in path.iterdir()
                if entry.is_file() and entry.suffix.lower() in IMAGE_EXTENSIONS
            ]
            return sorted(files, key=lambda file_path: file_path.name.casefold())

        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            return [path]

        return []

    def images_to_single_pdf(self):
        import img2pdf

        source_path = Path(self._source_path)
        files = self.get_image_files(source_path)
        if not files:
            self.update_status(self.t("status_select_valid"), "red")
            return

        self.update_status(self.t("status_merge"))
        try:
            self.report_progress(0, 1, files[0])
            out_dir = self.resolve_output_dir(source_path)
            output_path = self.get_output_path(out_dir, "Merged_Result", ".pdf")
            result = img2pdf.convert([str(image_path) for image_path in files])
            if self.is_cancelled():
                return
            with output_path.open("wb") as file:
                file.write(result)
            self.report_progress(1, 1, files[-1])
            self.update_status(self.t("status_done_pdf"), "green")
        except Exception as exc:
            self.show_error("status_error", exc)

    def images_to_multi_pdf(self):
        import img2pdf

        source_path = Path(self._source_path)
        files = self.get_image_files(source_path)
        if not files:
            self.update_status(self.t("status_select_valid_folder"), "red")
            return

        self.update_status(self.t("status_convert_each"))
        try:
            output_dir = self.resolve_output_dir(source_path)
            for index, image_path in enumerate(files, start=1):
                if self.is_cancelled():
                    break
                self.report_progress(index - 1, len(files), image_path)
                output_path = self.get_output_path(output_dir, image_path.stem, ".pdf")
                with output_path.open("wb") as pdf_file:
                    pdf_file.write(img2pdf.convert(str(image_path)))
                self.report_progress(index, len(files), image_path)
            self.update_status(self.t("status_done_multi_pdf", count=len(files)), "green")
        except Exception as exc:
            self.show_error("status_error", exc)

    def images_to_single_excel(self):
        import pandas as pd
        import openpyxl
        from PIL import Image
        import pytesseract

        configure_portable_tesseract(pytesseract)
        source_path = Path(self._source_path)
        files = self.get_image_files(source_path)
        if not files:
            self.update_status(self.t("status_select_valid"), "red")
            return

        self.update_status(self.t("status_ocr"))
        try:
            output_dir = self.resolve_output_dir(source_path)
            output_path = self.get_output_path(output_dir, "Images_Collection", ".xlsx")
            with pd.ExcelWriter(output_path, engine=openpyxl.__name__) as writer:
                for index, image_path in enumerate(files, start=1):
                    if self.is_cancelled():
                        break
                    self.report_progress(index - 1, len(files), image_path)
                    with Image.open(image_path) as image:
                        text = pytesseract.image_to_string(image, lang=OCR_LANGUAGES)
                    if self.is_cancelled():
                        break

                    rows = []
                    for line in text.splitlines():
                        cleaned = line.strip()
                        if cleaned:
                            normalized = self.normalize_arabic_text(cleaned)
                            rows.append(normalized.split())
                    if not rows:
                        rows = [["No text detected"]]

                    pd.DataFrame(rows).to_excel(writer, sheet_name=f"Page_{index}", index=False, header=False)
                    self.report_progress(index, len(files), image_path)
            self.update_status(self.t("status_ocr_done"), "green")
        except Exception as exc:
            self.show_error("status_ocr_error", exc)

    def images_to_multi_excel(self):
        import pandas as pd
        import openpyxl
        from PIL import Image
        import pytesseract

        configure_portable_tesseract(pytesseract)
        source_path = Path(self._source_path)
        files = self.get_image_files(source_path)
        if not files:
            self.update_status(self.t("status_select_valid_folder"), "red")
            return

        self.update_status(self.t("status_ocr_multi"))
        try:
            output_dir = self.resolve_output_dir(source_path)
            for index, image_path in enumerate(files, start=1):
                if self.is_cancelled():
                    break
                self.report_progress(index - 1, len(files), image_path)
                with Image.open(image_path) as image:
                    text = pytesseract.image_to_string(image, lang=OCR_LANGUAGES)
                if self.is_cancelled():
                    break

                rows = []
                for line in text.splitlines():
                    cleaned = line.strip()
                    if cleaned:
                        normalized = self.normalize_arabic_text(cleaned)
                        rows.append(normalized.split())
                if not rows:
                    rows = [["No text detected"]]

                excel_path = self.get_output_path(output_dir, image_path.stem, ".xlsx")
                with pd.ExcelWriter(excel_path, engine=openpyxl.__name__) as writer:
                    pd.DataFrame(rows).to_excel(writer, index=False, header=False)
                self.report_progress(index, len(files), image_path)
            self.update_status(self.t("status_ocr_done_multi", count=len(files)), "green")
        except Exception as exc:
            self.show_error("status_ocr_error", exc)

    def pdf_to_images(self):
        from pdf2image import convert_from_path, pdfinfo_from_path

        source_path = Path(self._source_path)
        if source_path.suffix.lower() != ".pdf" or not source_path.is_file():
            self.update_status(self.t("status_select_valid_pdf"), "red")
            return

        self.update_status(self.t("status_pdf_images"))
        try:
            output_dir = self.resolve_output_dir(source_path)
            poppler_path = portable_poppler_path()
            total_pages = int(
                pdfinfo_from_path(
                    str(source_path),
                    poppler_path=poppler_path,
                ).get("Pages", 0)
            )
            batch_size = 6
            for first_page in range(1, total_pages + 1, batch_size):
                if self.is_cancelled():
                    break
                last_page = min(first_page + batch_size - 1, total_pages)
                images = convert_from_path(
                    str(source_path),
                    first_page=first_page,
                    last_page=last_page,
                    poppler_path=poppler_path,
                )
                try:
                    for offset, image in enumerate(images):
                        index = first_page + offset
                        if self.is_cancelled():
                            break
                        self.report_progress(index - 1, total_pages, f"{source_path.name} - Page {index}")
                        output_path = self.get_output_path(output_dir, f"Page_{index}", ".jpg")
                        image.save(output_path, "JPEG")
                        self.report_progress(index, total_pages, f"{source_path.name} - Page {index}")
                finally:
                    for image in images:
                        image.close()
                if self.is_cancelled():
                    break
            self.update_status(self.t("status_pdf_images_done"), "green")
        except Exception as exc:
            self.show_error("status_pdf_error", exc)

    def pdf_to_excel(self):
        import pandas as pd
        import openpyxl
        import pdfplumber

        source_path = Path(self._source_path)
        if source_path.suffix.lower() != ".pdf" or not source_path.is_file():
            self.update_status(self.t("status_select_valid_pdf"), "red")
            return

        self.update_status(self.t("status_pdf_table"))
        try:
            output_dir = self.resolve_output_dir(source_path)
            data = []
            convert_from_path = None
            pytesseract = None
            poppler_path = None
            with pdfplumber.open(str(source_path)) as pdf:
                total_pages = len(pdf.pages)
                for index, page in enumerate(pdf.pages, start=1):
                    if self.is_cancelled():
                        break
                    self.report_progress(index - 1, total_pages, f"{source_path.name} - Page {index}")
                    page_text = page.extract_text() or ""
                    has_readable_text = any(
                        character.isprintable() and not character.isspace()
                        for character in page_text
                    )
                    table = page.extract_table()
                    if table:
                        data.extend(table)
                    elif has_readable_text:
                        for line in page_text.splitlines():
                            cleaned = line.strip()
                            if cleaned:
                                data.append(self.normalize_arabic_text(cleaned).split())
                    else:
                        if convert_from_path is None:
                            from pdf2image import convert_from_path as pdf_to_images
                            import pytesseract as tesseract

                            convert_from_path = pdf_to_images
                            pytesseract = tesseract
                            configure_portable_tesseract(pytesseract)
                            poppler_path = portable_poppler_path()

                        page_images = convert_from_path(
                            str(source_path),
                            first_page=index,
                            last_page=index,
                            dpi=250,
                            thread_count=1,
                            poppler_path=poppler_path,
                        )
                        if page_images:
                            try:
                                if not self.is_cancelled():
                                    scanned_text = pytesseract.image_to_string(
                                        page_images[0],
                                        lang=OCR_LANGUAGES,
                                    )
                                    for line in scanned_text.splitlines():
                                        cleaned = line.strip()
                                        if cleaned:
                                            data.append(self.normalize_arabic_text(cleaned).split())
                            finally:
                                for image in page_images:
                                    image.close()
                    self.report_progress(index, total_pages, f"{source_path.name} - Page {index}")

            if self.is_cancelled():
                return
            output_path = self.get_output_path(output_dir, source_path.stem, ".xlsx")
            with pd.ExcelWriter(output_path, engine=openpyxl.__name__) as writer:
                pd.DataFrame(data).to_excel(writer, index=False, header=False)
            self.report_progress(total_pages, total_pages, source_path.name)
            self.update_status(self.t("status_pdf_table_done"), "green")
        except Exception as exc:
            self.show_error("status_table_error", exc)


if __name__ == "__main__":
    enable_windows_dpi_awareness()
    app = DocNexusApp()
    app.mainloop()
