# -*- coding: utf-8 -*-
"""
Drukowanie FS z wyborem wzorca per kontrahent (zapamiętywane w CSV) + launcher dialogi.
"""

# ===== Standard library =====
from __future__ import annotations

import argparse
import builtins
import csv
import logging
import os
import sys
import tempfile
import time
import tkinter as tk
from datetime import datetime
from datetime import time as dtime
from datetime import timedelta
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable, Optional

# ===== Third-party =====
import pythoncom
import pywintypes
import win32com.client as win32
import win32print
from pywintypes import com_error

# ============================================================================ #
#                                   KONFIG                                     
# ============================================================================ #

# Sfera / Subiekt
GTA_URUCHOM_UI = 1  # 4 = gtaUruchomWTle (bez interfejsu) – zwykle chcemy 1

# Drukarka: None => domyślna systemowa; albo wpisz nazwę, np. "Microsoft Print to PDF"
DEFAULT_PRINTER = os.getenv("SFERA_PRINTER_NAME") or None
DEFAULT_PRINTER = "Microsoft Print to PDF" if DEFAULT_PRINTER == "None" else DEFAULT_PRINTER

# Ścieżka do CSV z mapowaniem kh_id -> wzw_id (None => %LOCALAPPDATA%\Subiektowe\wzorce_kontrahentow.csv)
STORAGE_PATH: Optional[str] = "wzorce_kontrahentow.csv"

# Nazwa pliku logu: prefiks + data
LOG_PREFIX = "FS_"

# ============================================================================ #
#                                  LOGOWANIE                                   
# ============================================================================ #

def setup_logging(
    log_dir: str = "logs",
    level: int = logging.INFO,
    echo_to_console: bool = True,
    capture_print: bool = True,
) -> str:
    """
    Logi do pliku logs/<PREFIX>YYYY-MM-DD.log (append) + opcjonalnie na konsolę.
    Przechwytuje print() -> logger.info() bez ręcznego echo na stdout (brak duplikatów).
    """
    date_str = datetime.now().strftime(f"{LOG_PREFIX}%Y-%m-%d")
    log_path = Path(log_dir) / f"{date_str}.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)

    handlers: list[logging.Handler] = [
        logging.FileHandler(log_path, mode="a", encoding="utf-8")
    ]
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


logger = logging.getLogger(__name__)

# ============================================================================ #
#                          POMOCNICZE: CZAS / ŚCIEŻKI                          
# ============================================================================ #

def to_com_time(dt: datetime):
    return pywintypes.Time(dt)

def _default_storage_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or os.getcwd()
    folder = Path(base) / "Subiektowe"
    folder.mkdir(parents=True, exist_ok=True)
    return str(folder / "wzorce_kontrahentow.csv")

def resolve_storage_path(path: Optional[str]) -> str:
    return path if path else _default_storage_path()

# ============================================================================ #
#                      CSV: zapamiętany wzorzec per kontrahent                 
# ============================================================================ #

def load_mapping_csv(path: Optional[str] = None) -> dict[int, int]:
    """Wczytuje mapowanie kh_id -> wzw_id z CSV."""
    p = Path(resolve_storage_path(path))
    if not p.exists():
        return {}
    mapping: dict[int, int] = {}
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        for row in r:
            if not row or row[0] == "kh_id":
                continue
            try:
                mapping[int(row[0])] = int(row[1])
            except Exception:
                continue
    return mapping

def save_mapping_csv(mapping: dict[int, int], path: Optional[str] = None) -> None:
    """Zapisuje mapowanie kh_id -> wzw_id do CSV (atomowo)."""
    p = Path(resolve_storage_path(path))
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="wzorce_", suffix=".csv", dir=str(p.parent))
    try:
        with os.fdopen(tmp_fd, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["kh_id", "wzw_id"])
            for kh_id, wzw_id in mapping.items():
                w.writerow([kh_id, wzw_id])
        os.replace(tmp_path, p)
    finally:
        if os.path.exists(tmp_path) and not os.path.samefile(tmp_path, p):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

def get_saved_wzor(kh_id: int, path: Optional[str] = None) -> Optional[int]:
    return load_mapping_csv(path).get(kh_id)

def set_saved_wzor(kh_id: int, wzw_id: int, path: Optional[str] = None) -> None:
    m = load_mapping_csv(path)
    m[int(kh_id)] = int(wzw_id)
    save_mapping_csv(m, path)

# ============================================================================ #
#                        ADO/COM: zapytania pomocnicze                          
# ============================================================================ #

def run_sql(spAplikacja, sql: str) -> list[dict]:
    """
    Wykonuje dowolny SELECT w Subiekcie przez ADO/COM.
    Zwraca listę słowników: [{kolumna: wartość, ...}, ...]
    """
    conn = spAplikacja.Aplikacja.Baza.Polaczenie  # ADODB.Connection
    rs = win32.Dispatch("ADODB.Recordset")
    # adUseClient=3, adOpenStatic=3, adLockReadOnly=1, adCmdText=1
    rs.CursorLocation = 3
    rs.Open(sql, conn, 3, 1, 1)

    results: list[dict] = []
    field_names = [f.Name for f in rs.Fields]
    while not rs.EOF:
        results.append({name: rs.Fields[name].Value for name in field_names})
        rs.MoveNext()
    rs.Close()
    return results

def fetch_wzorce_fs(spAplikacja) -> list[dict]:
    return run_sql(
        spAplikacja,
        """
        SELECT wz.wzw_Id, wz.wzw_Nazwa
          FROM wy_Wzorzec wz
          JOIN wy_Typ wt ON wz.wzw_Typ = wt.wtp_Id
         WHERE wt.wtp_Nazwa = 'Faktura sprzedaży'
         ORDER BY wz.wzw_Nazwa
        """,
    )

def fetch_kontrahenci_basic(spAplikacja) -> dict[int, str]:
    rows = run_sql(
        spAplikacja,
        """
        SELECT k.kh_Id,
               a.adr_Nazwa       AS Nazwa,
               a.adr_Adres       AS Adres,
               a.adr_Miejscowosc AS Miejscowosc
          FROM kh__Kontrahent k
          JOIN adr__Ewid a ON k.kh_Id = a.adr_IdObiektu
         WHERE a.adr_TypAdresu = 1
        """,
    )
    return {
        int(r["kh_Id"]): f"{r['Nazwa']}, {r['Adres']}, {r['Miejscowosc']}"
        for r in rows
    }

# ============================================================================ #
#                               DIALOGI TKINTER                                
# ============================================================================ #

def choose_wzor_wydruku(
    nazwa_kontrahenta: str,
    wzorce: list[dict],
    num: int,
    total: int,
    kh_id: Optional[int] = None,
    preselect_wzw_id: Optional[int] = None,
    remember_default: bool = True,
    on_remember: Optional[Callable[[int], None]] = None,
) -> Optional[dict]:
    """Okno wyboru wzorca wydruku."""
    # normalize
    items: list[dict] = []
    for w in wzorce or []:
        try:
            items.append({"wzw_Id": int(w["wzw_Id"]), "wzw_Nazwa": str(w["wzw_Nazwa"])})
        except Exception:
            pass
    if not items:
        messagebox.showwarning("Brak wzorców", "Nie znaleziono żadnych wzorców wydruku.")
        return None

    selection_holder: dict[str, Optional[dict]] = {"value": None}

    root = tk.Tk()
    root.title(f"Wybierz wzór wydruku ({num}/{total})")
    root.geometry("700x520")
    root.minsize(560, 400)
    root.lift()
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.grab_set()

    frame = ttk.Frame(root, padding=12)
    frame.pack(fill="both", expand=True)

    ttk.Label(
        frame,
        text=f'Wybierz wzór wydruku dla kontrahenta: "{nazwa_kontrahenta}"',
        font=("Segoe UI", 11, "bold"),
        wraplength=650,
        justify="left",
    ).pack(anchor="w", pady=(0, 8))

    # filtr
    filter_frame = ttk.Frame(frame)
    filter_frame.pack(fill="x", pady=(0, 6))
    ttk.Label(filter_frame, text="Filtruj:").pack(side="left")
    filter_var = tk.StringVar()
    filter_entry = ttk.Entry(filter_frame, textvariable=filter_var)
    filter_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

    # lista
    columns = ("id", "nazwa")
    tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
    tree.heading("id", text="ID")
    tree.heading("nazwa", text="Nazwa wzorca")
    tree.column("id", width=90, anchor="center")
    tree.column("nazwa", anchor="w")
    tree.pack(fill="both", expand=True)

    vsb = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")

    filtered = list(items)

    def refresh_tree():
        tree.delete(*tree.get_children())
        for it in filtered:
            tree.insert("", "end", values=(it["wzw_Id"], it["wzw_Nazwa"]))

    def _preselect(wzw_id: Optional[int]):
        try:
            if wzw_id is None:
                first = tree.get_children()
                if first:
                    tree.selection_set(first[0])
                    tree.focus(first[0])
                return
            for iid in tree.get_children():
                vals = tree.item(iid)["values"]
                if int(vals[0]) == int(wzw_id):
                    tree.selection_set(iid)
                    tree.focus(iid)
                    tree.see(iid)
                    return
            # fallback: pierwszy
            first = tree.get_children()
            if first:
                tree.selection_set(first[0])
                tree.focus(first[0])
        except Exception:
            pass

    def on_filter_change(*_):
        q = filter_var.get().strip().lower()
        filtered.clear()
        if not q:
            filtered.extend(items)
        else:
            for it in items:
                if q in it["wzw_Nazwa"].lower() or q in str(it["wzw_Id"]):
                    filtered.append(it)
        refresh_tree()
        _preselect(preselect_wzw_id)

    filter_var.trace_add("write", on_filter_change)

    # checkbox "Zapamiętaj"
    remember_var = tk.BooleanVar(value=bool(remember_default))
    chk_text = f"Zapamiętaj wybór dla kontrahenta (ID: {kh_id})" if kh_id is not None else "Zapamiętaj wybór"
    ttk.Checkbutton(frame, text=chk_text, variable=remember_var).pack(anchor="w", pady=(8, 0))

    # przyciski
    btns = ttk.Frame(frame)
    btns.pack(fill="x", pady=(10, 0))
    btn_ok = ttk.Button(btns, text="OK")
    btn_cancel = ttk.Button(btns, text="Anuluj")
    btn_cancel.pack(side="right")
    btn_ok.pack(side="right", padx=(0, 8))

    def pick_current():
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Wybór", "Zaznacz wzór z listy.")
            return
        values = tree.item(sel[0])["values"]
        wz = {"wzw_Id": int(values[0]), "wzw_Nazwa": str(values[1])}
        selection_holder["value"] = wz
        try:
            if remember_var.get() and callable(on_remember):
                on_remember(wz["wzw_Id"])
        finally:
            root.destroy()

    def on_double_click(_): pick_current()
    def on_return(_):       pick_current()
    def on_escape(_):
        selection_holder["value"] = None
        root.destroy()

    btn_ok.configure(command=pick_current)
    btn_cancel.configure(command=lambda: on_escape(None))
    tree.bind("<Double-1>", on_double_click)
    root.bind("<Return>", on_return)
    root.bind("<Escape>", on_escape)

    refresh_tree()
    _preselect(preselect_wzw_id)
    filter_entry.focus_set()

    root.mainloop()
    return selection_holder["value"]  # dict lub None


def ask_delay_seconds(
    default: int = 5,
    min_seconds: int = 0,
    max_seconds: int = 600,
) -> Optional[int]:
    """Modalny dialog: opóźnienie (sekundy) między wydrukami."""
    result = {"value": None}

    root = tk.Tk()
    root.title("Opóźnienie między wydrukami")
    root.resizable(False, False)
    root.lift()
    root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.grab_set()

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(
        frm,
        text="Ile sekund program ma czekać między drukowaniami?",
        font=("Segoe UI", 10, "bold"),
        wraplength=360,
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

    ttk.Label(frm, text="Sekundy:").grid(row=1, column=0, sticky="e", padx=(0, 6))

    val_var = tk.StringVar(value=str(default))

    def validate(P):
        if P == "":
            return True
        if P.isdigit():
            n = int(P)
            return min_seconds <= n <= max_seconds
        return False

    vcmd = (root.register(validate), "%P")
    entry = ttk.Entry(frm, textvariable=val_var, width=8, justify="center", validate="key", validatecommand=vcmd)
    entry.grid(row=1, column=1, sticky="w")
    ttk.Label(frm, text=f"(min {min_seconds}, max {max_seconds})").grid(row=1, column=2, sticky="w", padx=(6, 0))

    btns = ttk.Frame(frm)
    btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))

    def ok():
        s = val_var.get().strip()
        if not s:
            messagebox.showinfo("Wartość wymagana", "Podaj liczbę sekund.")
            return
        try:
            n = int(s)
        except ValueError:
            messagebox.showerror("Błąd", "Wpisz liczbę całkowitą.")
            return
        if not (min_seconds <= n <= max_seconds):
            messagebox.showerror("Zakres", f"Podaj wartość {min_seconds}–{max_seconds}.")
            return
        result["value"] = n
        root.destroy()

    def cancel():
        result["value"] = None
        root.destroy()

    ttk.Button(btns, text="Anuluj", command=cancel).pack(side="right")
    ttk.Button(btns, text="OK", command=ok).pack(side="right", padx=(0, 8))

    root.bind("<Return>", lambda e: ok())
    root.bind("<Escape>", lambda e: cancel())

    entry.focus_set()
    entry.select_range(0, tk.END)

    root.mainloop()
    return result["value"]

# ============================================================================ #
#                              DRUKOWANIE / SUBIEKT                            
# ============================================================================ #

def ensure_printer_exists(name: str) -> None:
    available = {
        p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    }
    if name not in available:
        raise RuntimeError(f"Nie znaleziono drukarki: {name}")

def drukuj_wg_ustawien(
    su_dokument,
    wzw_id: int,
    printer_name: Optional[str] = DEFAULT_PRINTER,
    ilosc_kopii: int = 1,
    strona_od: Optional[int] = None,
    strona_do: Optional[int] = None,
) -> None:
    """Wywołuje DrukujWgUstawien na obiekcie dokumentu."""
    ust = win32.Dispatch("InsERT.UstawieniaWydruku")
    ust.WzorzecWydruku = int(wzw_id)

    if printer_name:
        ensure_printer_exists(printer_name)
        ust.DrukarkaDomyslSysOp = False
        ust.Drukarka = printer_name
    else:
        ust.DrukarkaDomyslSysOp = True

    ust.IloscKopii = int(ilosc_kopii) if ilosc_kopii else 1
    if strona_od is not None:
        ust.StronaOd = int(strona_od)
    if strona_do is not None:
        ust.StronaDo = int(strona_do)

    su_dokument.DrukujWgUstawien(ust)

def get_subiekt() -> any:
    """Logowanie do Subiekta wg zmiennych środowiskowych."""
    gt = win32.Dispatch("InsERT.GT")
    gt.Produkt = 1                        # gtaProduktSubiekt
    gt.Autentykacja = 0                   # gtaAutentykacjaSQL
    gt.Serwer = os.getenv("SFERA_SQL_SERVER", "127.0.0.1")
    gt.Uzytkownik = os.getenv("SFERA_SQL_LOGIN", "sa")
    gt.UzytkownikHaslo = os.getenv("SFERA_SQL_PASSWORD", "SqlPassword01!")
    gt.Baza = os.getenv("SFERA_SQL_DB", "sfera_demo")
    gt.Operator = os.getenv("SFERA_OPERATOR", "admin")
    gt.OperatorHaslo = os.getenv("SFERA_OPERATOR_PASSWORD", "admin")
    sub = win32.Dispatch(gt.Uruchom(1, GTA_URUCHOM_UI))
    return sub

def select_docs_prev_month(dok_manager) -> list:
    """
    Otwiera okno wyboru dokumentów (FS, poprzedni miesiąc) i zwraca listę zaznaczonych.
    """
    dok = dok_manager.Wybierz()
    dok.FiltrTyp = 2            # wszystkie możliwe Faktury Sprzedaży
    dok.FiltrOkres = 20         # gtaFiltrOkresDowolnyMiesiac

    # poprzedni miesiac: końcówka dnia poprzedzającego 1-szy dzień bieżącego
    today = datetime.now()
    first_this = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_prev = first_this - timedelta(seconds=1)
    dok.FiltrOkresUstawDowolnyMiesiac(to_com_time(last_prev))

    dok.MultiSelekcja = True
    logger.info("Otwieram okno wyboru dokumentów... zaznacz dokumenty do wydruku i kliknij OK.")
    dok.Wyswietl()

    # zmaterializuj iterator
    docs = list(dok.ZaznaczoneDokumenty())
    logger.info("Wybrano %d dokumentów do wydruku.", len(docs))
    return docs


def show_completion_dialog(logfile: str | None = None, logs_dir: str = "logs") -> None:
    """
    Pokazuje modalne okno 'Operacja zakończona' i czeka na OK.
    Jeśli podasz 'logfile', folder logów zostanie określony na jego katalog.
    """
    # Ustal ścieżki do pokazania / otwarcia
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
            # Windows:
            os.startfile(str(logs_path))
        except Exception:
            # Fallback na inne systemy:
            import webbrowser
            webbrowser.open(str(logs_path))

    ttk.Button(btns, text="Otwórz folder logów", command=open_logs).pack(side="left")
    ttk.Button(btns, text="OK", command=root.destroy).pack(side="right")

    root.bind("<Return>", lambda e: root.destroy())
    root.bind("<Escape>", lambda e: root.destroy())

    root.mainloop()


# ============================================================================ #
#                                     MAIN                                     
# ============================================================================ #

def main():
    # (Opcjonalnie) argumenty CLI
    ap = argparse.ArgumentParser()
    ap.add_argument("--storage", help="Ścieżka do CSV z wyborem wzorców (domyślna lokalna/APPDATA).")
    ap.add_argument("--printer", help="Nazwa drukarki (None = domyślna systemowa).")
    args = ap.parse_args()

    storage_path = args.storage if args.storage else STORAGE_PATH
    printer_name = args.printer if args.printer else DEFAULT_PRINTER

    pythoncom.CoInitialize()
    sub = None
    try:
        sub = get_subiekt()
        logger.info("Subiekt GT Sfera %s, baza: %s (%s)", sub.Aplikacja.Wersja, sub.Baza.Nazwa, sub.Baza.Serwer)

        # dane referencyjne
        wzorce = fetch_wzorce_fs(sub)
        wz_by_id = {int(w["wzw_Id"]): str(w["wzw_Nazwa"]) for w in wzorce}
        kontrahenci = fetch_kontrahenci_basic(sub)

        # wybór dokumentów
        docs = select_docs_prev_month(sub.Dokumenty)

        # zbuduj listę unikalnych kontrahentów
        kh_ids: list[int] = []
        for d in docs:
            if d.KontrahentId not in kh_ids:
                kh_ids.append(int(d.KontrahentId))

        # wybór / preselekcja wzorców per kontrahent
        wz_kontr: dict[int, Optional[int]] = {}
        for i, kh_id in enumerate(kh_ids, start=1):
            def _remember(wzw_id: int, _kh=kh_id):
                set_saved_wzor(_kh, wzw_id, storage_path)

            nazwa = kontrahenci.get(kh_id, f"KH {kh_id}")
            logger.info("Wybór wzorca dla %s (ID: %s) (%d/%d)", nazwa, kh_id, i, len(kh_ids))
            pre = get_saved_wzor(kh_id, storage_path)
            wyb = choose_wzor_wydruku(
                nazwa_kontrahenta=nazwa,
                wzorce=wzorce,
                num=i,
                total=len(kh_ids),
                kh_id=kh_id,
                preselect_wzw_id=pre,
                remember_default=True,
                on_remember=_remember,
            )
            logger.info("Wybrano: %s", wyb["wzw_Nazwa"] if wyb else "Anulowano")
            wz_kontr[kh_id] = int(wyb["wzw_Id"]) if wyb else None

        # opóźnienie między drukami
        delay = ask_delay_seconds(default=5) or 0

        # drukowanie
        for i, d in enumerate(docs, start=1):
            wzw_id = wz_kontr.get(int(d.KontrahentId))
            if not wzw_id:
                logger.warning("Pomijam %s – brak wybranego wzorca.", getattr(d, "NumerPelny", "<bez numeru>"))
                continue
            wz_name = wz_by_id.get(wzw_id, f"wzorzec {wzw_id}")
            logger.info("Drukuję (%d/%d) %s wzorem %s", i, len(docs), d.NumerPelny, wz_name)
            drukuj_wg_ustawien(d, wzw_id=wzw_id, printer_name=printer_name, ilosc_kopii=1)
            if delay and i < len(docs):
                logger.info("  ... czekam %d sekund ...", delay)
                time.sleep(delay)

    except com_error as e:
        logger.exception("Błąd COM: %s", e)
    except Exception as e:
        logger.exception("Błąd krytyczny: %s", e)
    finally:
        try:
            if sub is not None:
                sub.Zakoncz()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    logfile = setup_logging()
    logger.info("Start aplikacji. Logi zapisuję do pliku: %s", logfile)
    try:
        main()
    finally:
        show_completion_dialog(logfile=logfile, logs_dir="logs")
