"""PAFFP Formatter - window with two tabs:
  1. From masterlist (Excel) -> System Upload, AR and STUB files
  2. New format (STUB sheets) -> AR (PDF), STUB with QR codes (PDF), System Upload (Excel)
"""
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import paffp_formatter as engine


def run_masterlist(path, out_dir, selected, amount, log):
    formats = [f for f in engine.FORMATS if f["name"] in selected]
    engine.process(path, out_dir=out_dir, formats=formats, default_amount=amount, log=log)


def run_newformat(path, out_dir, selected, amount, log, province=""):
    try:
        import newformat
    except ImportError as e:
        raise RuntimeError(f"{e}. Install it with:  pip install pymupdf qrcode fonttools")
    kinds = [k for k in ("AR", "STUB", "SYSTEM") if k in selected]
    newformat.process(path, out_dir, kinds=kinds, amount=amount, log=log, province=province)


class Panel(ttk.Frame):
    """Pick files, choose outputs, set amount and folder, run in the background."""

    def __init__(self, master, app, files_title, filetypes, options, amount_label, worker, noun,
                 province_field=False):
        super().__init__(master)
        self.app, self.filetypes, self.worker, self.noun = app, filetypes, worker, noun
        self.files = []
        self.make = {key: tk.BooleanVar(value=default) for key, _, default in options}
        self.amount = tk.StringVar(value=str(engine.DEFAULT_AMOUNT or ""))
        self.province = tk.StringVar() if province_field else None
        self.msgs = queue.Queue()
        pad = {"padx": 10, "pady": 5}

        box = ttk.LabelFrame(self, text=f"1. {files_title}")
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
        self.listbox = tk.Listbox(lst, selectmode="extended", height=5, activestyle="none")
        sb = ttk.Scrollbar(lst, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        opts = ttk.LabelFrame(self, text="2. What to create")
        opts.pack(fill="x", **pad)
        for key, label, _ in options:
            ttk.Checkbutton(opts, text=label, variable=self.make[key]).pack(side="left", padx=10, pady=6)
        ttk.Entry(opts, textvariable=self.amount, width=9).pack(side="right", padx=(0, 10), pady=6)
        ttk.Label(opts, text=amount_label).pack(side="right", padx=4)
        if province_field:
            prov = ttk.Frame(self)
            prov.pack(fill="x", padx=10, pady=(0, 5))
            ttk.Label(prov, text="Province:").pack(side="left", padx=(6, 4))
            ttk.Entry(prov, textvariable=self.province, width=22).pack(side="left")
            ttk.Label(prov, text="printed in the AR and STUB heading (e.g. IFUGAO); "
                                 "empty = the Province on each STUB sheet",
                      foreground="#5d6b62").pack(side="left", padx=8)

        out = ttk.LabelFrame(self, text="3. Save to folder")
        out.pack(fill="x", **pad)
        ttk.Entry(out, textvariable=app.out_dir).pack(side="left", fill="x", expand=True, padx=6, pady=6)
        ttk.Button(out, text="Browse...", command=app.pick_out).pack(side="left", padx=(0, 6))
        ttk.Button(out, text="Open folder", command=app.open_out).pack(side="left", padx=(0, 6))

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
        self.after(100, self._poll)

    # ------------------------------------------------------------ files
    def add_files(self):
        start = engine.INPUT_DIR if engine.INPUT_DIR.exists() else engine.BASE_DIR
        picked = filedialog.askopenfilenames(title=f"Select {self.noun} files", initialdir=start,
                                             filetypes=self.filetypes + [("All files", "*.*")])
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

    # ------------------------------------------------------------ run
    def write(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        selected = [k for k, v in self.make.items() if v.get()]
        if not self.files:
            messagebox.showwarning("PAFFP Formatter", f"Add at least one {self.noun} file first.")
            return
        if not selected:
            messagebox.showwarning("PAFFP Formatter", "Tick at least one output.")
            return
        amount = self.amount.get().replace(",", "").strip()
        if amount and not amount.replace(".", "", 1).isdigit():
            messagebox.showwarning("PAFFP Formatter", "Amount must be a number, e.g. 2325.")
            return
        amount = (float(amount) if "." in amount else int(amount)) if amount else None
        self.go.state(["disabled"])
        self.bar.configure(maximum=len(self.files), value=0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        extra = {"province": self.province.get().strip().upper()} if self.province else {}
        threading.Thread(target=self._work, args=(list(self.files), selected, self.app.out_dir.get(), amount, extra),
                         daemon=True).start()

    def _work(self, files, selected, out_dir, amount, extra):
        ok = failed = 0
        for i, f in enumerate(files, 1):
            self.msgs.put(("log", f"Processing {Path(f).name} ..."))
            try:
                self.worker(f, out_dir, selected, amount, lambda s: self.msgs.put(("log", s)), **extra)
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
                    self.write(f"\nDone: {ok} file(s) formatted, {failed} failed.")
                    if failed:
                        messagebox.showerror("PAFFP Formatter",
                                             f"{ok} done, {failed} failed. See the Status box for details.")
                    elif messagebox.askyesno("PAFFP Formatter",
                                             f"Finished {ok} file(s).\n\nOpen the output folder?"):
                        os.startfile(out_dir)
        except queue.Empty:
            pass
        self.after(100, self._poll)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PAFFP Formatter - System Upload, AR & STUB")
        self.geometry("800x640")
        self.minsize(640, 500)
        self.out_dir = tk.StringVar(value=str(engine.OUTPUT_DIR))
        ttk.Style(self).configure("Go.TButton", font=("Segoe UI", 11, "bold"), padding=8)

        tabs = ttk.Notebook(self)
        tabs.pack(fill="both", expand=True, padx=6, pady=6)
        self.master_tab = Panel(
            tabs, self, "Masterlist files", [("Excel or PDF files", "*.xlsx *.xlsm *.pdf")],
            [(f["name"], f["label"], True) for f in engine.FORMATS],
            "Amount if missing:", run_masterlist, "masterlist")
        self.stub_tab = Panel(
            tabs, self, 'Source workbooks ("STUB <MUNICIPALITY>" sheets, or any list with column headings)',
            [("Excel or PDF files", "*.xlsx *.xlsm *.pdf")],
            [("AR", "AR format (PDF)", True), ("STUB", "STUB format with QR (PDF)", True),
             ("SYSTEM", "System upload (Excel)", True)],
            "AR amount:", run_newformat, "source workbook", province_field=True)
        tabs.add(self.master_tab, text="  From masterlist (Excel)  ")
        tabs.add(self.stub_tab, text="  New format (STUB sheets)  ")

    def pick_out(self):
        d = filedialog.askdirectory(title="Save output files to", initialdir=self.out_dir.get())
        if d:
            self.out_dir.set(d)

    def open_out(self):
        d = Path(self.out_dir.get())
        d.mkdir(parents=True, exist_ok=True)
        os.startfile(d)


if __name__ == "__main__":
    App().mainloop()
