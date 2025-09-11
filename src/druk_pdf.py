import os
import sys
import time
import platform
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---- Platform printer backends -------------------------------------------------

class WindowsPrinterBackend:
    def __init__(self):
        import win32print  # type: ignore
        import win32api    # type: ignore
        import win32gui    # type: ignore
        self.win32print = win32print
        self.win32api = win32api
        self.win32gui = win32gui

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

    def show_printer_properties_dialog(self, printer_name):
        """Opens the system Printing Preferences UI.
        """
        try:
            subprocess.Popen(["rundll32", "printui.dll,PrintUIEntry", "/p", f"/n{printer_name}"])
        except Exception:
            pass
        return False
   
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
        self.recursive = tk.BooleanVar(value=False)
        self.delay = tk.IntVar(value=10)  # spacing between jobs

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
        ttk.Button(frm2, text="Właściwości…", command=self.open_printer_properties).grid(row=0, column=3, sticky="e", padx=(8,0))
        frm2.columnconfigure(1, weight=1)

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

    def open_printer_properties(self):
        name = self.printer.get()
        if not name:
            messagebox.showwarning("Brak drukarki", "Wybierz drukarkę.")
            return
        try:
            self.backend.show_printer_properties_dialog(name)
        except Exception as e:
            messagebox.showerror("Właściwości drukarki", f"Nie udało się otworzyć właściwości: {e}")

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
        delay = max(0, int(self.delay.get()))

        self._log(f"Drukarka: {printer_name}\n---\n")

        try:
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
