"""PAFFP Formatter - window for creating AR and STUB files from masterlists."""
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import paffp_formatter as engine


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PAFFP Formatter - AR & STUB")
        self.geometry("760x560")
        self.minsize(620, 460)
        self.files = []
        self.out_dir = tk.StringVar(value=str(engine.OUTPUT_DIR))
        self.make = {f["name"]: tk.BooleanVar(value=True) for f in engine.FORMATS}
        self.msgs = queue.Queue()
        self._build()
        self.after(100, self._poll)

    # ------------------------------------------------------------ layout
    def _build(self):
        pad = {"padx": 10, "pady": 5}
        style = ttk.Style(self)
        style.configure("Go.TButton", font=("Segoe UI", 11, "bold"), padding=8)

        box = ttk.LabelFrame(self, text="1. Masterlist files")
        box.pack(fill="both", expand=True, **pad)
        btns = ttk.Frame(box)
        btns.pack(fill="x", padx=6, pady=4)
        ttk.Button(btns, text="Add files...", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="Remove selected", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="Clear", command=self.clear_files).pack(side="left")
        self.count = ttk.Label(btns, text="No files selected")
        self.count.pack(side="right")
        lst = ttk.Frame(box)
        lst.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.listbox = tk.Listbox(lst, selectmode="extended", height=6, activestyle="none")
        sb = ttk.Scrollbar(lst, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        opts = ttk.LabelFrame(self, text="2. What to create")
        opts.pack(fill="x", **pad)
        for f in engine.FORMATS:
            ttk.Checkbutton(opts, text=f"{f['name']} format", variable=self.make[f["name"]]).pack(
                side="left", padx=10, pady=6)

        out = ttk.LabelFrame(self, text="3. Save to folder")
        out.pack(fill="x", **pad)
        ttk.Entry(out, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=6, pady=6)
        ttk.Button(out, text="Browse...", command=self.pick_out).pack(side="left", padx=(0, 6))
        ttk.Button(out, text="Open folder", command=self.open_out).pack(side="left", padx=(0, 6))

        run = ttk.Frame(self)
        run.pack(fill="x", **pad)
        self.go = ttk.Button(run, text="CREATE", style="Go.TButton", command=self.start)
        self.go.pack(side="left")
        self.bar = ttk.Progressbar(run, mode="determinate")
        self.bar.pack(side="left", fill="x", expand=True, padx=10)

        logf = ttk.LabelFrame(self, text="Status")
        logf.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(logf, height=8, wrap="word", state="disabled", font=("Consolas", 9))
        lsb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=lsb.set)
        self.log.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        lsb.pack(side="right", fill="y", pady=6)

    # ------------------------------------------------------------ actions
    def add_files(self):
        start = engine.INPUT_DIR if engine.INPUT_DIR.exists() else engine.BASE_DIR
        picked = filedialog.askopenfilenames(
            title="Select masterlist files", initialdir=start,
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        for p in picked:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", Path(p).name)
        self._update_count()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self._update_count()

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.files.clear()
        self._update_count()

    def _update_count(self):
        n = len(self.files)
        self.count.config(text=f"{n} file{'s' if n != 1 else ''} selected" if n else "No files selected")

    def pick_out(self):
        d = filedialog.askdirectory(title="Save output files to", initialdir=self.out_dir.get())
        if d:
            self.out_dir.set(d)

    def open_out(self):
        d = Path(self.out_dir.get())
        d.mkdir(parents=True, exist_ok=True)
        os.startfile(d)

    def write(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        formats = [f for f in engine.FORMATS if self.make[f["name"]].get()]
        if not self.files:
            messagebox.showwarning("PAFFP Formatter", "Add at least one masterlist file first.")
            return
        if not formats:
            messagebox.showwarning("PAFFP Formatter", "Tick AR and/or STUB.")
            return
        self.go.state(["disabled"])
        self.bar.configure(maximum=len(self.files), value=0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=self._work, args=(list(self.files), formats, self.out_dir.get()),
                         daemon=True).start()

    def _work(self, files, formats, out_dir):
        ok = failed = 0
        for i, f in enumerate(files, 1):
            self.msgs.put(("log", f"Processing {Path(f).name} ..."))
            try:
                engine.process(f, out_dir=out_dir, formats=formats,
                               log=lambda s: self.msgs.put(("log", s)))
                ok += 1
            except Exception as e:
                failed += 1
                self.msgs.put(("log", f"   ERROR: {e}"))
            self.msgs.put(("progress", i))
        self.msgs.put(("done", (ok, failed, out_dir)))

    def _poll(self):
        try:
            while True:
                kind, data = self.msgs.get_nowait()
                if kind == "log":
                    self.write(data)
                elif kind == "progress":
                    self.bar.configure(value=data)
                elif kind == "done":
                    ok, failed, out_dir = data
                    self.go.state(["!disabled"])
                    self.write(f"\nDone: {ok} masterlist(s) formatted, {failed} failed.")
                    if failed:
                        messagebox.showerror("PAFFP Formatter",
                                             f"{ok} done, {failed} failed. See the Status box for details.")
                    elif messagebox.askyesno("PAFFP Formatter",
                                             f"Finished {ok} masterlist(s).\n\nOpen the output folder?"):
                        os.startfile(out_dir)
        except queue.Empty:
            pass
        self.after(100, self._poll)


if __name__ == "__main__":
    App().mainloop()
