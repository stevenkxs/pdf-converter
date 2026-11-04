"""Desktop UI for converter.py: pick PDFs, choose formats, convert."""
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from converter import FORMATS, convert, ocr_languages

LABELS = {"docx": "Word (.docx)", "xlsx": "Excel (.xlsx)", "pptx": "PowerPoint (.pptx)"}
OCR_MODES = {"Auto (scanned pages only)": "auto", "Always (every page)": "force", "Off": "off"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Converter")
        self.geometry("760x580")
        self.files: list[str] = []
        self.out_dir = tk.StringVar(value="")
        self.fmt = {f: tk.BooleanVar(value=True) for f in FORMATS}
        self.msgs: queue.Queue[str] = queue.Queue()
        self.ocr = tk.StringVar(value=next(iter(OCR_MODES)))
        self.lang = tk.StringVar(value="eng")

        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="Add PDFs...", command=self.add).pack(side="left")
        ttk.Button(top, text="Clear", command=self.clear).pack(side="left", padx=4)

        self.listbox = tk.Listbox(self, height=8)
        self.listbox.pack(fill="both", expand=True, **pad)

        opts = ttk.LabelFrame(self, text="Convert to")
        opts.pack(fill="x", **pad)
        for f in FORMATS:
            ttk.Checkbutton(opts, text=LABELS[f], variable=self.fmt[f]).pack(side="left", padx=8)

        out = ttk.Frame(self)
        out.pack(fill="x", **pad)
        ttk.Label(out, text="Output folder:").pack(side="left")
        ttk.Entry(out, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(out, text="Browse...", command=self.pick_out).pack(side="left")
        ttk.Label(self, text="Leave the output folder empty to save next to each PDF.").pack(anchor="w", padx=8)

        ocr = ttk.LabelFrame(self, text="OCR for scanned PDFs")
        ocr.pack(fill="x", **pad)
        ttk.Label(ocr, text="Mode:").pack(side="left", padx=4)
        ttk.Combobox(ocr, textvariable=self.ocr, values=list(OCR_MODES), state="readonly",
                     width=24).pack(side="left")
        ttk.Label(ocr, text="Language(s):").pack(side="left", padx=(12, 4))
        ttk.Combobox(ocr, textvariable=self.lang, values=ocr_languages(), width=14).pack(side="left")
        ttk.Label(ocr, text="combine with +, e.g. eng+chi_sim").pack(side="left", padx=4)

        self.go = ttk.Button(self, text="Convert", command=self.start)
        self.go.pack(**pad)
        self.log = tk.Text(self, height=8, state="disabled")
        self.log.pack(fill="both", expand=True, **pad)
        self.after(100, self.drain)

    def add(self):
        for f in filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")]):
            if f not in self.files:
                self.files.append(f)
                self.listbox.insert("end", f)

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, "end")

    def pick_out(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def start(self):
        formats = [f for f in FORMATS if self.fmt[f].get()]
        if not self.files or not formats:
            self.msgs.put("Add at least one PDF and choose at least one format.")
            return
        self.go.state(["disabled"])
        ocr, lang = OCR_MODES[self.ocr.get()], self.lang.get().strip() or "eng"
        threading.Thread(target=self.work, args=(list(self.files), formats, ocr, lang), daemon=True).start()

    def work(self, files, formats, ocr, lang):
        for f in files:
            self.msgs.put(f"Converting {Path(f).name} ...")
            convert(f, formats, self.out_dir.get() or None, log=self.msgs.put, ocr=ocr, lang=lang)
        self.msgs.put("Done.")
        self.after(0, lambda: self.go.state(["!disabled"]))

    def drain(self):  # UI thread: move worker messages into the log box
        while not self.msgs.empty():
            self.log.configure(state="normal")
            self.log.insert("end", self.msgs.get() + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(100, self.drain)


if __name__ == "__main__":
    App().mainloop()
