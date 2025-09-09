import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable, Optional


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

def choose_output_dir(default_dir: Path) -> Path:
    try:
        from tkinter import Tk, filedialog
        root = Tk(); root.withdraw()
        chosen = filedialog.askdirectory(initialdir=str(default_dir), title="Wybierz folder zapisu")
        root.destroy()
        return Path(chosen) if chosen else default_dir
    except Exception:
        # brak tkinter/GUI — użyj domyślnego
        return default_dir