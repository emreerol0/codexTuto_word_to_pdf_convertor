import os
import threading
import queue
from dataclasses import dataclass
from typing import List, Optional

import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


@dataclass
class ConversionTask:
    source_path: str
    target_path: str


class WordToPdfConverter:
    """Converts Word documents to PDF using Microsoft Word COM automation."""

    PDF_FORMAT = 17

    def __init__(self) -> None:
        self.word_app = None

    def __enter__(self):
        import win32com.client  # imported lazily for clearer startup errors

        pythoncom.CoInitialize()
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False
        self.word_app.DisplayAlerts = 0
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.word_app is not None:
            self.word_app.Quit()
        pythoncom.CoUninitialize()

    def convert(self, source_path: str, target_path: str) -> None:
        if not source_path.lower().endswith((".doc", ".docx")):
            raise ValueError(f"Unsupported file type: {source_path}")

        doc = self.word_app.Documents.Open(os.path.abspath(source_path))
        try:
            doc.SaveAs(os.path.abspath(target_path), FileFormat=self.PDF_FORMAT)
        finally:
            doc.Close(False)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Word to PDF Converter")
        self.geometry("900x560")

        self.files: List[str] = []
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.conversion_running = False
        self.worker_thread: Optional[threading.Thread] = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(120, self._drain_log_queue)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        title = ttk.Label(
            self,
            text="Convert .doc/.docx files to PDF",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=14, pady=(14, 8))

        button_bar = ttk.Frame(self)
        button_bar.grid(row=1, column=0, sticky="ew", padx=14)
        button_bar.columnconfigure(6, weight=1)

        ttk.Button(button_bar, text="Add Files", command=self.add_files).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(button_bar, text="Add Folder", command=self.add_folder).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(button_bar, text="Remove Selected", command=self.remove_selected).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(button_bar, text="Clear", command=self.clear_files).grid(row=0, column=3, padx=(0, 8))
        self.convert_btn = ttk.Button(button_bar, text="Convert", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=4, padx=(0, 8))

        self.progress = ttk.Progressbar(button_bar, mode="determinate", length=220)
        self.progress.grid(row=0, column=5, padx=(16, 0), sticky="e")

        list_frame = ttk.LabelFrame(self, text="Queue")
        list_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(10, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, font=("Consolas", 10))
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        log_frame = ttk.LabelFrame(self, text="Activity")
        log_frame.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 14))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

    def add_files(self) -> None:
        picked = filedialog.askopenfilenames(
            title="Select Word documents",
            filetypes=[("Word Documents", "*.doc *.docx"), ("All Files", "*.*")],
        )
        self._add_paths(picked)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder with Word documents")
        if not folder:
            return

        doc_files = []
        for name in os.listdir(folder):
            if name.lower().endswith((".doc", ".docx")):
                doc_files.append(os.path.join(folder, name))
        self._add_paths(doc_files)

    def _add_paths(self, paths) -> None:
        added = 0
        for path in paths:
            path = os.path.normpath(path)
            if path not in self.files and path.lower().endswith((".doc", ".docx")):
                self.files.append(path)
                self.listbox.insert(tk.END, path)
                added += 1

        if added:
            self._log(f"Added {added} file(s) to queue.")
        elif paths:
            self._log("No new valid files were added.")

    def remove_selected(self) -> None:
        selected = list(self.listbox.curselection())
        for idx in reversed(selected):
            del self.files[idx]
            self.listbox.delete(idx)
        if selected:
            self._log(f"Removed {len(selected)} file(s) from queue.")

    def clear_files(self) -> None:
        if not self.files:
            return
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self._log("Queue cleared.")

    def start_conversion(self) -> None:
        if self.conversion_running:
            return

        if not self.files:
            messagebox.showwarning("No files", "Please add at least one .doc or .docx file.")
            return

        tasks = self._prompt_save_paths()
        if not tasks:
            self._log("Conversion cancelled.")
            return

        self.progress.configure(maximum=len(tasks), value=0)
        self.conversion_running = True
        self.convert_btn.configure(state="disabled")

        self.worker_thread = threading.Thread(target=self._convert_worker, args=(tasks,))
        self.worker_thread.start()

    def _on_close(self) -> None:
        if self.conversion_running:
            messagebox.showwarning(
                "Conversion in progress",
                "Please wait for the current conversion to finish before closing the app.",
            )
            return
        self.destroy()

    def _prompt_save_paths(self) -> Optional[List[ConversionTask]]:
        tasks: List[ConversionTask] = []
        for source in self.files:
            default_name = os.path.splitext(os.path.basename(source))[0] + ".pdf"
            target = filedialog.asksaveasfilename(
                title=f"Save PDF As - {os.path.basename(source)}",
                defaultextension=".pdf",
                initialfile=default_name,
                filetypes=[("PDF", "*.pdf")],
            )
            if not target:
                return None
            tasks.append(ConversionTask(source_path=source, target_path=target))
        return tasks

    def _convert_worker(self, tasks: List[ConversionTask]) -> None:
        try:
            with WordToPdfConverter() as converter:
                for idx, task in enumerate(tasks, 1):
                    self.log_queue.put(f"Converting: {task.source_path}")
                    converter.convert(task.source_path, task.target_path)
                    self.log_queue.put(f"Saved PDF: {task.target_path}")
                    self.log_queue.put(f"__PROGRESS__{idx}")
            self.log_queue.put("__DONE__SUCCESS")
        except Exception as exc:  # noqa: BLE001
            self.log_queue.put(f"ERROR: {exc}")
            self.log_queue.put("__DONE__FAIL")

    def _drain_log_queue(self) -> None:
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            if msg.startswith("__PROGRESS__"):
                value = int(msg.replace("__PROGRESS__", ""))
                self.progress.configure(value=value)
                continue

            if msg == "__DONE__SUCCESS":
                self.conversion_running = False
                self.convert_btn.configure(state="normal")
                messagebox.showinfo("Done", "All selected files were converted successfully.")
                continue

            if msg == "__DONE__FAIL":
                self.conversion_running = False
                self.convert_btn.configure(state="normal")
                messagebox.showerror("Conversion failed", "A conversion error occurred. See Activity log.")
                continue

            self._log(msg)

        self.after(120, self._drain_log_queue)

    def _log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
