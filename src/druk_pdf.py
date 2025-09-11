from gui import ask_delay_seconds, show_completion_dialog

import os
import subprocess
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ---- Platform printer backends -------------------------------------------------

class WindowsPrinterBackend:
    def __init__(self):
        import win32print  # type: ignore
        import win32api    # type: ignore
        self.win32print = win32print
        self.win32api = win32api

    def list_printers(self):
        flags = self.win32print.PRINTER_ENUM_LOCAL | self.win32print.PRINTER_ENUM_CONNECTIONS
        printers = self.win32print.EnumPrinters(flags)
        names = [p[2] for p in printers]
        # Put default first
        try:
            default = self.win32print.GetDefaultPrinter()
            names.sort(key=lambda n: (n != default, n.lower()))
        except Exception:
            names.sort(key=lambda n: n.lower())
        return names

    def apply_printer_prefs(self, printer_name, duplex_mode, orientation):
        """Best-effort set duplex/orientation on the printer defaults for this process.
        duplex_mode: 'simplex' | 'long' | 'short'
        orientation: 'portrait' | 'landscape'
        Returns a context manager that restores settings afterwards.
        """
        import contextlib
        win32print = self.win32print
        DMDUP = {"simplex": 1, "long": 2, "short": 3}
        DMORIENT = {"portrait": 1, "landscape": 2}

        @contextlib.contextmanager
        def _ctx():
            hprinter = None
            old_dev = None
            try:
                hprinter = win32print.OpenPrinter(printer_name)
                level = 2
                info = win32print.GetPrinter(hprinter, level)
                old_dev = info["pDevMode"]
                # Clone a DEVMODE we can edit
                dev = win32print.DocumentProperties(0, hprinter, printer_name, None, None, 0)
                # When called with None it returns size; we need a real structure next
                dev = win32print.DocumentProperties(0, hprinter, printer_name, old_dev, None, 2)
                # Set duplex
                if duplex_mode in DMDUP:
                    dev.Duplex = DMDUP[duplex_mode]
                    dev.Fields |= 0x1000  # DM_DUPLEX
                # Set orientation
                if orientation in DMORIENT:
                    dev.Orientation = DMORIENT[orientation]
                    dev.Fields |= 0x1  # DM_ORIENTATION
                # Apply
                win32print.DocumentProperties(0, hprinter, printer_name, dev, dev, 0x3)  # DM_IN_BUFFER|DM_OUT_BUFFER
                info["pDevMode"] = dev
                win32print.SetPrinter(hprinter, level, info, 0)
                yield
            except Exception:
                # If anything fails, just proceed with defaults
                yield
            finally:
                if hprinter:
                    try:
                        # Attempt to restore old settings
                        if old_dev:
                            info = win32print.GetPrinter(hprinter, 2)
                            info["pDevMode"] = old_dev
                            win32print.SetPrinter(hprinter, 2, info, 0)
                    except Exception:
                        pass
                    win32print.ClosePrinter(hprinter)
        return _ctx()

    def print_pdf(self, printer_name, pdf_path):
        # Use ShellExecute with PrintTo verb so the registered PDF handler does the rendering
        # Note: this call is asynchronous; add a small delay between jobs to avoid overloading the handler
        self.win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
    
    def print_with_adobe(self, printer, pdf_path):
        acro_paths = [
            r"c:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe",
            r"C:\\Program Files\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe",
            r"C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe",
        ]
        exe = next((p for p in acro_paths if os.path.exists(p)), None)
        if not exe:
            raise RuntimeError("Nie znaleziono AcroRd32.exe – sprawdź instalację Adobe Reader")
        cmd = [exe, "/N", "/T", pdf_path, printer]
        subprocess.Popen(cmd, shell=False)


# ---- UI -----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch PDF Printer")
        self.geometry("680x460")
        self.resizable(True, True)

        # Decide backend
        try:
            self.backend = WindowsPrinterBackend()
        except Exception as e:
            messagebox.showerror("Printer backend error", f"Nie udało się zainicjalizować obsługi drukarki.\n{e}\n\nNa Windows zainstaluj: pywin32")
            self.destroy()
            raise

        # State
        self.folder = tk.StringVar()
        self.printer = tk.StringVar()
        self.duplex = tk.StringVar(value="simplex")  # simplex | long | short
        self.orientation = tk.StringVar(value="portrait")  # portrait | landscape
        self.recursive = tk.BooleanVar(value=False)
        self.delay = tk.IntVar(value=5)  # spacing between jobs

        self._build_ui()
        self._load_printers()

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 8}

        frm = ttk.Frame(self)
        frm.pack(fill=tk.X, **pad)
        ttk.Label(frm, text="Folder z PDF-ami:").pack(side=tk.LEFT)
        ent = ttk.Entry(frm, textvariable=self.folder)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(frm, text="Wybierz…", command=self.choose_folder).pack(side=tk.LEFT)

        frm2 = ttk.Frame(self)
        frm2.pack(fill=tk.X, **pad)
        ttk.Label(frm2, text="Drukarka:").grid(row=0, column=0, sticky="w")
        self.printer_combo = ttk.Combobox(frm2, textvariable=self.printer, state="readonly", width=50)
        self.printer_combo.grid(row=0, column=1, sticky="we", padx=8)
        ttk.Button(frm2, text="Odśwież", command=self._load_printers).grid(row=0, column=2, sticky="e")
        frm2.columnconfigure(1, weight=1)

        frm3 = ttk.Frame(self)
        frm3.pack(fill=tk.X, **pad)
        ttk.Label(frm3, text="Dwustronnie:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frm3, text="Nie", value="simplex", variable=self.duplex).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frm3, text="Dł. krawędź", value="long", variable=self.duplex).grid(row=0, column=2, sticky="w")
        ttk.Radiobutton(frm3, text="Krót. kraw.", value="short", variable=self.duplex).grid(row=0, column=3, sticky="w")

        ttk.Label(frm3, text="Orientacja:").grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Radiobutton(frm3, text="Pion", value="portrait", variable=self.orientation).grid(row=1, column=1, sticky="w", pady=(8,0))
        ttk.Radiobutton(frm3, text="Poziom", value="landscape", variable=self.orientation).grid(row=1, column=2, sticky="w", pady=(8,0))

        frm4 = ttk.Frame(self)
        frm4.pack(fill=tk.X, **pad)
        ttk.Checkbutton(frm4, text="Skanuj podfoldery (rekurencyjnie)", variable=self.recursive).pack(side=tk.LEFT)
        ttk.Label(frm4, text="Odstęp między zadaniami [s]:").pack(side=tk.LEFT, padx=(16,4))
        ttk.Spinbox(frm4, from_=0, to=10000, textvariable=self.delay, width=7).pack(side=tk.LEFT)

        frm5 = ttk.Frame(self)
        frm5.pack(fill=tk.BOTH, expand=True, **pad)
        self.log = tk.Text(frm5, height=12)
        self.log.pack(fill=tk.BOTH, expand=True)

        frm6 = ttk.Frame(self)
        frm6.pack(fill=tk.X, **pad)
        self.start_btn = ttk.Button(frm6, text="Drukuj wszystkie PDF-y", command=self.start_print)
        self.start_btn.pack(side=tk.LEFT)
        ttk.Button(frm6, text="Zamknij", command=self.destroy).pack(side=tk.RIGHT)

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder.set(path)

    def _load_printers(self):
        try:
            names = self.backend.list_printers()
        except Exception as e:
            messagebox.showerror("Błąd drukarek", str(e))
            names = []
        self.printer_combo["values"] = names
        if names and not self.printer.get():
            self.printer.set(names[0])

    def _iter_pdfs(self, root, recursive=False):
        if recursive:
            for dirpath, _, filenames in os.walk(root):
                for f in sorted(filenames):
                    if f.lower().endswith('.pdf'):
                        yield os.path.join(dirpath, f)
        else:
            for f in sorted(os.listdir(root)):
                if f.lower().endswith('.pdf'):
                    yield os.path.join(root, f)

    def start_print(self):
        folder = self.folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Brak folderu", "Wybierz poprawny folder z PDF-ami.")
            return
        if not self.printer.get():
            messagebox.showwarning("Brak drukarki", "Wybierz drukarkę.")
            return

        pdfs = list(self._iter_pdfs(folder, self.recursive.get()))
        if not pdfs:
            messagebox.showinfo("Brak plików", "Nie znaleziono żadnych PDF-ów w wybranym folderze.")
            return

        self.start_btn["state"] = "disabled"
        self.log.delete("1.0", tk.END)
        self.log.insert(tk.END, f"Znalezione PDF-y: {len(pdfs)}\n")

        t = threading.Thread(target=self._print_worker, args=(pdfs,))
        t.daemon = True
        t.start()

    def _print_worker(self, pdfs):
        printer_name = self.printer.get()
        duplex = self.duplex.get()
        orient = self.orientation.get()
        delay = max(0, int(self.delay.get()))

        self._log(f"Drukarka: {printer_name}\nDwustronnie: {duplex}\nOrientacja: {orient}\n---\n")

        try:
            with self.backend.apply_printer_prefs(printer_name, duplex, orient):
                for i, pdf in enumerate(pdfs, 1):
                    try:
                        # self.backend.print_pdf(printer_name, pdf) # Używając printto - nie działa za każdym razem
                        self.backend.print_with_adobe(printer_name, pdf)
                        self._log(f"[{i}/{len(pdfs)}] Wysłano: {os.path.basename(pdf)}\n")
                    except Exception as e:
                        self._log(f"BŁĄD przy {os.path.basename(pdf)}: {e}\n")
                    if i < len(pdfs):
                        self._log(f" ... czekam {delay} sekund ...\n")
                        time.sleep(delay)
            self._log("\nGotowe.\n")
        finally:
            self.start_btn["state"] = "normal"

    def _log(self, msg):
        self.log.insert(tk.END, msg)
        self.log.see(tk.END)


if __name__ == "__main__":
    app = App()
    app.mainloop()
