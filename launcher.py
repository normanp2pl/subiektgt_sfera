import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

# ===== KONFIGURACJA APLIKACJI =====
# Możesz dopisać kolejne pozycje. Ścieżki względne liczone są od folderu tego pliku.
# 'python' (opcjonalny) pozwala wskazać inny interpreter, np. 32-bitowy venv.
APPS = [
    {
        "id": "drukuj_fs",
        "label": "Drukowanie FS",
        "script": "drukuj_fs.py",
    },
    {
        "id": "zmiana_mm",
        "label": "Zmiana dat dokumentów MM",
        "script": "zmiana_mm.py",
    },
]

# ====== LAUNCHER ======

BASE_DIR = Path(__file__).resolve().parent

def resolve_script_path(p: str) -> Path:
    path = Path(p)
    return path if path.is_absolute() else (BASE_DIR / path)

def launch_app(app: dict):
    try:
        script_path = resolve_script_path(app["script"])
        if not script_path.exists():
            raise FileNotFoundError(f"Nie znaleziono skryptu: {script_path}")

        # Interpreter: domyślnie ten sam, ale można nadpisać w configu.
        python_exe = Path(app.get("python") or sys.executable)
        if not python_exe.exists():
            raise FileNotFoundError(f"Nie znaleziono interpretera Pythona: {python_exe}")

        args = list(map(str, app.get("args", [])))
        cwd = Path(app.get("cwd") or script_path.parent)
        env = os.environ.copy()
        env.update(app.get("env", {}))

        creationflags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)  # Windows: nowe okno konsoli

        proc = subprocess.Popen(
            [str(python_exe), str(script_path), *args],
            cwd=str(cwd),
            env=env,
            creationflags=creationflags,
        )
        # messagebox.showinfo("Uruchomiono",
        #                     f"'{app['label']}' wystartowała.\nPID: {proc.pid}\n\n"
        #                     f"Skrypt: {script_path}\nInterpreter: {python_exe}")
    except Exception as e:
        messagebox.showerror("Błąd uruchamiania", str(e))

def build_ui(root: tk.Tk):
    root.title("Sfera apps launcher by DevNorman")
    root.geometry("420x260")
    root.minsize(380, 220)
    root.lift()
    root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))

    container = ttk.Frame(root, padding=12)
    container.pack(fill="both", expand=True)

    ttk.Label(container, text="Wybierz aplikację:", font=("Segoe UI", 11, "bold")).pack(
        anchor="w", pady=(0, 8)
    )

    # siatka przycisków
    grid = ttk.Frame(container)
    grid.pack(fill="both", expand=True)

    max_cols = 2  # ile kolumn z przyciskami
    for i, app in enumerate(APPS):
        r, c = divmod(i, max_cols)
        btn = ttk.Button(grid, text=app["label"], width=28, command=lambda a=app: launch_app(a))
        btn.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")

    # elastyczna siatka
    rows = (len(APPS) + max_cols - 1) // max_cols
    for i in range(rows):
        grid.rowconfigure(i, weight=1)
    for j in range(max_cols):
        grid.columnconfigure(j, weight=1)

    # pasek dolny
    bottom = ttk.Frame(container)
    bottom.pack(fill="x", pady=(8, 0))
    ttk.Button(bottom, text="Zamknij", command=root.destroy).pack(side="right")

def main():
    root = tk.Tk()
    try:
        # motyw (opcjonalnie)
        style = ttk.Style(root)
        # ustaw domyślny, ale jeśli masz 'vista' / 'clam' itd., możesz podmienić:
        # style.theme_use("vista")
    except Exception:
        pass
    build_ui(root)
    root.mainloop()

if __name__ == "__main__":
    main()
