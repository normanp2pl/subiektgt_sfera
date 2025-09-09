import builtins
import logging
import os
import sys
import tkinter as tk
from datetime import datetime, time, timedelta
from pathlib import Path
from tkinter import messagebox, ttk

import pythoncom
import pywintypes
import win32com.client as win32
from pywintypes import com_error

# ===================== Stałe =====================
gtaUruchom = 1   # (4 = gtaUruchomWTle jeśli chcesz bez interfejsu)
LOG_PREFIX = "MM_"   # prefiks nazwy pliku logu

# ===================== Logging =====================
def setup_logging(log_dir: str = "logs",
                  level: int = logging.INFO,
                  echo_to_console: bool = True,
                  capture_print: bool = True) -> str:
    """
    Logi do pliku logs/MM_YYYY-MM-DD.log (append) + opcjonalnie na konsolę.
    Przechwytuje print() -> logger.info() (bez ręcznego echo, więc bez duplikatów).
    """
    date_str = datetime.now().strftime(f"{LOG_PREFIX}%Y-%m-%d")
    log_path = Path(log_dir) / f"{date_str}.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)

    handlers = [logging.FileHandler(log_path, mode="a", encoding="utf-8")]
    if echo_to_console:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True,
    )

    if capture_print:
        _orig_print = builtins.print
        def print_to_logger(*args, **kwargs):
            file = kwargs.get("file")
            if file not in (None, sys.stdout, sys.stderr):
                return _orig_print(*args, **kwargs)
            sep = kwargs.get("sep", " ")
            msg = sep.join(str(a) for a in args)
            logging.getLogger().info(msg)
        builtins.print = print_to_logger

    return str(log_path)

# ===================== Util =====================
def to_com_time(dt: datetime):
    return pywintypes.Time(dt)

# ---------- GUI: jedno okno z datą i checkboxem dry-run ----------
def _parse_user_date(s: str) -> datetime.date:
    s = s.strip()
    # 1) ISO: YYYY-MM-DD
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    # 2) DD.MM.YYYY
    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except Exception:
        pass
    # 3) DD/MM/YYYY
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        pass
    raise ValueError("Nieprawidłowy format daty. Użyj YYYY-MM-DD lub DD.MM.YYYY")

def ask_new_date_and_dryrun(default_dayshift: int = 0, default_dryrun: bool = True):
    """
    Pyta użytkownika o nową datę i czy ma być DRY RUN.
    Zwraca tuple: (date: datetime.date, dry_run: bool) albo (None, None) jeśli anulowano.
    """
    result = {"date": None, "dry": None}

    root = tk.Tk()
    root.title("Ustawienia zmiany daty dokumentów")
    root.resizable(False, False)
    root.lift(); root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.grab_set()

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Podaj nową datę dla zaznaczonych dokumentów:",
              font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,8))

    today = datetime.now().date()
    default_date = today + timedelta(days=default_dayshift)
    date_var = tk.StringVar(value=default_date.isoformat())

    ttk.Label(frm, text="Nowa data:").grid(row=1, column=0, sticky="e", padx=(0,6))
    entry = ttk.Entry(frm, textvariable=date_var, width=14, justify="center")
    entry.grid(row=1, column=1, sticky="w")
    ttk.Label(frm, text="(YYYY-MM-DD)").grid(row=1, column=2, sticky="w", padx=(6,0))

    dry_var = tk.BooleanVar(value=bool(default_dryrun))
    ttk.Checkbutton(frm, text="DRY RUN (symulacja, bez zapisu)", variable=dry_var).grid(
        row=2, column=0, columnspan=3, sticky="w", pady=(10,0)
    )

    btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=3, sticky="e", pady=(12, 0))
    def ok():
        try:
            d = _parse_user_date(date_var.get())
        except Exception as e:
            messagebox.showerror("Błędna data", str(e)); return
        result["date"] = d
        result["dry"] = dry_var.get()
        root.destroy()
    def cancel():
        result["date"] = None
        result["dry"] = None
        root.destroy()

    ttk.Button(btns, text="Anuluj", command=cancel).pack(side="right")
    ttk.Button(btns, text="OK", command=ok).pack(side="right", padx=(0,8))
    root.bind("<Return>", lambda e: ok())
    root.bind("<Escape>", lambda e: cancel())
    entry.focus_set(); entry.select_range(0, tk.END)

    root.mainloop()
    return result["date"], result["dry"]

# ---------- GUI: okno "Operacja zakończona" ----------
def show_completion_dialog(logfile: str | None = None, logs_dir: str = "logs") -> None:
    """
    Pokazuje modalne okno 'Operacja zakończona' i czeka na OK.
    Jeśli podasz 'logfile', folder logów zostanie określony na jego katalog.
    """
    if logfile:
        logs_path = Path(logfile).resolve().parent
        last_file = Path(logfile).resolve()
    else:
        logs_path = Path(logs_dir).resolve()
        last_file = None

    root = tk.Tk()
    root.title("Operacja zakończona")
    root.resizable(False, False)
    root.lift()
    root.attributes("-topmost", True)
    root.focus_force()
    root.grab_set()

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    msg = (
        "Operacja zakończona.\n"
        "Przejrzyj logi i naciśnij OK, aby zamknąć program.\n\n"
        f"Logi znajdziesz w podkatalogu:\n{logs_path}"
    )
    if last_file:
        msg += f"\nNajnowszy plik logu:\n{last_file.name}"

    ttk.Label(frm, text=msg, justify="left", wraplength=520).grid(row=0, column=0, columnspan=3, sticky="w")

    btns = ttk.Frame(frm)
    btns.grid(row=1, column=0, columnspan=3, sticky="e", pady=(12, 0))

    def open_logs():
        try:
            os.startfile(str(logs_path))  # Windows
        except Exception:
            import webbrowser
            webbrowser.open(str(logs_path))

    ttk.Button(btns, text="Otwórz folder logów", command=open_logs).pack(side="left")
    ttk.Button(btns, text="OK", command=root.destroy).pack(side="right")

    root.bind("<Return>", lambda e: root.destroy())
    root.bind("<Escape>", lambda e: root.destroy())

    root.mainloop()

# ===================== GŁÓWNY SKRYPT =====================
def main():
    # Pytanie w GUI zamiast argparse:
    user_date, dry_run = ask_new_date_and_dryrun(default_dayshift=0, default_dryrun=True)
    if user_date is None:
        print("Anulowano przez użytkownika.")
        return

    # Ustal noon, aby uniknąć problemów z DST
    d_noon = datetime.combine(user_date, time(12, 0))
    new_date = to_com_time(d_noon)

    pythoncom.CoInitialize()
    sub = None
    try:
        gt = win32.Dispatch("InsERT.GT")
        gt.Produkt = 1 # gtaProduktSubiekt
        gt.Autentykacja = 0 # gtaAutentykacjaSQL
        gt.Serwer = os.getenv("SFERA_SQL_SERVER", "127.0.0.1")
        gt.Uzytkownik = os.getenv("SFERA_SQL_LOGIN", "sa")
        gt.UzytkownikHaslo = os.getenv("SFERA_SQL_PASSWORD", "SqlPassword01!")
        gt.Baza = os.getenv("SFERA_SQL_DB", "sfera_demo")
        gt.Operator = os.getenv("SFERA_OPERATOR", "admin")
        gt.OperatorHaslo = os.getenv("SFERA_OPERATOR_PASSWORD", "admin")

        sub = win32.Dispatch(gt.Uruchom(1, gtaUruchom))
        sub.MagazynId = 1

        dok = sub.Dokumenty.Wybierz()
        dok.FiltrTyp = 9   # MM
        dok.FiltrOkres = 20  # gtaFiltrOkresDowolnyMiesiac

        # Ustawienie miesiąca do wyboru (przykładowa data – podmień na swoje kryterium jeśli chcesz)
        dt = datetime(2025, 9, 4, 12, 0, 0)
        com_time = to_com_time(dt)

        dok.FiltrOkresUstawDowolnyMiesiac(com_time)
        dok.MultiSelekcja = True
        print("Otwieram okno wyboru dokumentów... Zaznacz dokumenty i kliknij OK.")
        dok.Wyswietl()

        selected = list(dok.ZaznaczoneDokumenty())
        if not selected:
            print("Nie wybrano żadnych dokumentów.")
            return

        for p in selected:
            if not p.NumerPelny.startswith("MM"):
                continue
            old_date = p.DataWystawienia.date()
            msg = (f"Zmieniam datę dokumentu {p.NumerPelny} o wartości {p.WartoscNetto} "
                   f"z dnia {old_date.isoformat()} na {user_date.isoformat()}")
            if dry_run:
                print("DRY RUN:", msg)
            else:
                p.DataWystawienia = new_date
                print(msg)
                # if not p.SkutekMagazynowy:
                #     print("  Dokument nie ma skutku magazynowego, pomijam")
                # else:
                p.Zapisz()

    except com_error as e:
        logging.exception("Błąd COM: %s", e)
    except Exception as e:
        logging.exception("Błąd krytyczny: %s", e)
    finally:
        try:
            if sub is not None:
                sub.Zakoncz()
        except Exception:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    logfile = setup_logging()  # <- tu powstaje logs/MM_YYYY-MM-DD.log
    print(f"Start aplikacji. Logi zapisuję do pliku: {logfile}")
    try:
        main()
    finally:
        show_completion_dialog(logfile=logfile, logs_dir="logs")
