import os, csv, tempfile
from pathlib import Path
import argparse
import pywintypes
from datetime import datetime, timedelta, time
import pythoncom
import win32com.client as win32
from pywintypes import com_error
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Optional

# Stałe
gtaUruchom = 1   # (4 = gtaUruchomWTle jeśli chcesz bez interfejsu)
storage_path = "wzorce_kontrahentow.csv"  # można zmienić na None aby użyć domyślnej ścieżki
def to_com_time(dt: datetime):
    return pywintypes.Time(dt)

def parse_args():
    ap = argparse.ArgumentParser()
    # ap.add_argument("--new-date", required=True)
    ap.add_argument("--dry-run", action="store_true")
    return ap.parse_args()


def _default_storage_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or os.getcwd()
    folder = Path(base) / "Subiektowe"
    folder.mkdir(parents=True, exist_ok=True)
    return str(folder / "wzorce_kontrahentow.csv")



def load_mapping_csv(path: str | None = None) -> dict[int, int]:
    """Wczytuje mapowanie kh_id -> wzw_id z CSV."""
    if path is None:
        path = _default_storage_path()
    mapping: dict[int, int] = {}
    p = Path(path)
    if not p.exists():
        return mapping
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        for row in r:
            if not row or len(row) < 2:
                continue
            try:
                kh_id = int(row[0])
                wzw_id = int(row[1])
                mapping[kh_id] = wzw_id
            except ValueError:
                # pomiń błędne wiersze
                continue
    return mapping


def save_mapping_csv(mapping: dict[int, int], path: str | None = None) -> None:
    """Zapisuje mapowanie kh_id -> wzw_id do CSV (atomowo)."""
    if path is None:
        path = _default_storage_path()
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="wzorce_", suffix=".csv", dir=str(p.parent))
    try:
        with os.fdopen(tmp_fd, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            # nagłówek (opcjonalny) – nie przeszkadza przy wczytywaniu
            w.writerow(["kh_id", "wzw_id"])
            for kh_id, wzw_id in mapping.items():
                w.writerow([kh_id, wzw_id])
        os.replace(tmp_path, p)  # atomowa podmiana
    finally:
        # jeżeli replace się nie udał, posprzątaj tmp:
        if os.path.exists(tmp_path) and not os.path.samefile(tmp_path, p):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def get_saved_wzor(kh_id: int, path: str | None = None) -> int | None:
    return load_mapping_csv(path).get(kh_id)


def set_saved_wzor(kh_id: int, wzw_id: int, path: str | None = None) -> None:
    m = load_mapping_csv(path)
    m[int(kh_id)] = int(wzw_id)
    save_mapping_csv(m, path)


def choose_wzor_wydruku(nazwa_kontrahenta: str,
                        wzorce: list[dict],
                        num: int, 
                        total: int,
                        kh_id: int | None = None,
                        preselect_wzw_id: int | None = None,
                        remember_default: bool = True,
                        on_remember: Callable[[int], None] | None = None) -> dict | None:
    """
    Okno wyboru wzoru wydruku.
    - preselect_wzw_id: wstępnie zaznacz ten wzór (jeśli istnieje na liście)
    - on_remember: callback (wzw_id:int) wywoływany gdy checkbox zaznaczony i user zatwierdzi
    """
    items = []
    for w in wzorce or []:
        try:
            items.append({"wzw_Id": int(w["wzw_Id"]), "wzw_Nazwa": str(w["wzw_Nazwa"])})
        except Exception:
            pass
    if not items:
        messagebox.showwarning("Brak wzorców", "Nie znaleziono żadnych wzorców wydruku.")
        return None

    selection_holder = {"value": None}

    root = tk.Tk()
    root.title(f"Wybierz wzór wydruku ({num}/{total})")
    root.geometry("700x520")
    root.minsize(560, 400)
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))

    frame = ttk.Frame(root, padding=12)
    frame.pack(fill="both", expand=True)

    lbl = ttk.Label(
        frame,
        text=f'Wybierz wzór wydruku dla kontrahenta: "{nazwa_kontrahenta}"',
        font=("Segoe UI", 11, "bold"),
        wraplength=650,
        justify="left",
    )
    lbl.pack(anchor="w", pady=(0, 8))

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
        # reselekcja
        _preselect(preselect_wzw_id)

    filter_var.trace_add("write", on_filter_change)

    # checkbox "Zapamiętaj"
    remember_var = tk.BooleanVar(value=bool(remember_default))
    chk_text = f"Zapamiętaj wybór dla kontrahenta (ID: {kh_id})" if kh_id is not None else "Zapamiętaj wybór"
    chk = ttk.Checkbutton(frame, text=chk_text, variable=remember_var)
    chk.pack(anchor="w", pady=(8, 0))

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

    # wypełnienie i preselekcja
    refresh_tree()

    def _preselect(wzw_id: int | None):
        try:
            if wzw_id is None:
                # zaznacz pierwszy
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
            # jeśli nie znaleziono - pierwszy
            first = tree.get_children()
            if first:
                tree.selection_set(first[0])
                tree.focus(first[0])
        except Exception:
            pass

    _preselect(preselect_wzw_id)
    filter_entry.focus_set()

    root.mainloop()
    return selection_holder["value"]



def ask_delay_seconds(default: int = 5,
                      min_seconds: int = 0,
                      max_seconds: int = 600) -> Optional[int]:
    """
    Pokazuje okno z pytaniem o opóźnienie (sekundy) między wydrukami.
    Zwraca liczbę sekund (int) lub None jeśli anulowano.
    """
    result = {"value": None}

    root = tk.Tk()
    root.title("Opóźnienie między wydrukami")
    root.resizable(False, False)

    # wyeksponuj okno i przejmij focus (modalne)
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
        wraplength=360
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

    ttk.Label(frm, text="Sekundy:").grid(row=1, column=0, sticky="e", padx=(0,6))

    val_var = tk.StringVar(value=str(default))

    def validate(P):
        if P == "":
            return True
        if P.isdigit():
            n = int(P)
            return min_seconds <= n <= max_seconds
        return False

    vcmd = (root.register(validate), "%P")
    entry = ttk.Entry(frm, textvariable=val_var, width=8,
                      justify="center", validate="key", validatecommand=vcmd)
    entry.grid(row=1, column=1, sticky="w")
    ttk.Label(frm, text=f"(min {min_seconds}, max {max_seconds})").grid(row=1, column=2, sticky="w", padx=(6,0))

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


def run_sql(spAplikacja, sql: str):
    """
    Wykonuje dowolny SELECT w Subiekcie przez ADO/COM.
    Zwraca listę słowników: [{kolumna: wartość, ...}, ...]
    """
    conn = spAplikacja.Aplikacja.Baza.Polaczenie  # ADODB.Connection
    rs = win32.Dispatch("ADODB.Recordset")

    # adUseClient=3, adOpenStatic=3, adLockReadOnly=1, adCmdText=1
    rs.CursorLocation = 3
    rs.Open(sql, conn, 3, 1, 1)

    results = []
    field_names = [f.Name for f in rs.Fields]

    while not rs.EOF:
        row = {}
        for name in field_names:
            row[name] = rs.Fields[name].Value
        results.append(row)
        rs.MoveNext()

    rs.Close()
    return results


def main():
    # args = parse_args()
    # d = datetime.fromisoformat(args.new_date)
    # d_noon = datetime.combine(d.date(), time(12, 0))
    # new_date = to_com_time(d_noon)


    pythoncom.CoInitialize()
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

        wzorce_wydrukow = run_sql(sub, """SELECT wzw_Id, wzw_Nazwa
                                            FROM wy_Wzorzec wz
                                            JOIN wy_Typ wt ON wz.wzw_Typ = wt.wtp_Id
                                           WHERE wt.wtp_Nazwa = 'Faktura sprzedaży'""")
        wz_dict = {wz['wzw_Id']: wz['wzw_Nazwa'] for wz in wzorce_wydrukow}
        
        kontrahenci = run_sql(sub, """SELECT k.kh_Id,
                                             a.adr_Nazwa AS Nazwa, 
                                             a.adr_Adres AS Adres,
                                             a.adr_Miejscowosc AS Miejscowość
                                        FROM kh__Kontrahent k 
                                  INNER JOIN adr__Ewid a ON k.kh_Id = a.adr_IdObiektu
                                       WHERE a.adr_TypAdresu=1""")
        kontrahenci_dict = {k['kh_Id']: f"{k['Nazwa']}, {k['Adres']}, {k['Miejscowość']}" for k in kontrahenci}
        sub.MagazynId = 1
       
        dok = sub.Dokumenty.Wybierz()
        dok.FiltrTyp = 2 # Wsztystkie możliwe Faktury Sprzedaży
        dok.FiltrOkres = 20 # gtaFiltrOkresDowolnyMiesiac 

        today = datetime.now()
        first_this = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        last_prev = first_this - timedelta(seconds=1)  # 23:59:59 dnia poprzedzającego
        com_time = to_com_time(last_prev)

        # dt = datetime(2025, 9, 4, 12, 0, 0)
        # com_time = to_com_time(dt)

        dok.FiltrOkresUstawDowolnyMiesiac(com_time) # poczatek miesiąca poprzedniego
        dok.MultiSelekcja = True
        dok.Wyswietl()
        num_of_docs = len(list(dok.ZaznaczoneDokumenty()))
        print(f"Wybrano {num_of_docs} dokumentów do wydruku.")
        ids = []
        for p in dok.ZaznaczoneDokumenty():   # iteracja po zaznaczonych dokumentach
            if p.KontrahentId not in ids:
                ids.append(p.KontrahentId)
        
        num = 0
        wz_kontr_dict = {}
        for id in ids:
            def _remember(wzw_id: int):
                set_saved_wzor(id, wzw_id, storage_path)
            num += 1
            print(f"Wybór wzorca dla {kontrahenci_dict[id]} (ID: {id}) ({num}/{len(ids)})")
            pre = get_saved_wzor(id, storage_path)
            wybrany = choose_wzor_wydruku(
                nazwa_kontrahenta=kontrahenci_dict[id],
                wzorce=wzorce_wydrukow,
                kh_id=id,
                preselect_wzw_id=pre,
                remember_default=True,
                on_remember=_remember,
                num = num,
                total = len(ids)
            )
            print("Wybrano:", wybrany['wzw_Nazwa'] if wybrany else "Anulowano")
            wz_kontr_dict[id] = wybrany['wzw_Id'] if wybrany else None
        
        delay = ask_delay_seconds(default=5)
        num = 0
        for p in dok.ZaznaczoneDokumenty():   # iteracja po zaznaczonych dokumentach
            num += 1
            ust = win32.Dispatch("InsERT.UstawieniaWydruku")
            if wz_kontr_dict.get(p.KontrahentId):
                ust.WzorzecWydruku = wz_kontr_dict.get(p.KontrahentId)
            ust.DrukarkaDomyslSysOp = False
            ust.Drukarka = "Microsoft Print to PDF"  # PDF
          # oUstWyd.DrukarkaDomyslSysOp = False
          # oUstWyd.Drukarka = "NazwaDrukarki"
          # oUstWyd.WzorzecWydruku = 4
            ust.IloscKopii = 1
          # oUstWyd.StronaOd = 1
          # oUstWyd.StronaDo = 2
            print(f"Drukuję ({num}/{num_of_docs}) {p.NumerPelny} wzorem {wz_dict[wz_kontr_dict[p.KontrahentId]]}")
            p.DrukujWgUstawien(ust)
            if delay and delay > 0 and num < num_of_docs:
                import time
                print(f"  ... czekam {delay} sekund ...")
                time.sleep(delay)
    except com_error as e:
        print("Błąd COM:", e)
    finally:
        try:
            sub.Zakoncz()
        except Exception:
            pass

if __name__ == "__main__":
    main()
