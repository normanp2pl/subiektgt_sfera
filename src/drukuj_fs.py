# -*- coding: utf-8 -*-
"""
Drukowanie/export FS z wyborem wzorca per kontrahent (zapamiętywane w CSV) + launcher dialogi.
"""

# ===== Standard library =====
from __future__ import annotations

import argparse
import csv
import logging
import os
import tempfile
from pathlib import Path
from typing import Optional

# ===== Third-party =====
import pythoncom
import win32com.client as win32
import win32print
from pywintypes import com_error

import logowanie
from gui import choose_wzor_wydruku, show_completion_dialog, choose_output_dir
from utils import get_subiekt, run_sql, select_docs_prev_month, safe_filename

logger = logging.getLogger(__name__)

# ============================================================================ #
#                                   KONFIG                                     
# ============================================================================ #

# Drukarka: None => domyślna systemowa; albo wpisz nazwę, np. "Microsoft Print to PDF"
DEFAULT_PRINTER = os.getenv("SFERA_PRINTER_NAME") or None
DEFAULT_PRINTER = "Microsoft Print to PDF" if DEFAULT_PRINTER == "None" else DEFAULT_PRINTER

# Ścieżka do CSV z mapowaniem kh_id -> wzw_id (None => %LOCALAPPDATA%\Subiektowe\wzorce_kontrahentow.csv)
STORAGE_PATH: Optional[str] = None

# Nazwa pliku logu: prefiks + data
LOG_PREFIX = "FS_"

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
    # printer_name = args.printer if args.printer else DEFAULT_PRINTER

    default_dir = (Path.cwd().parent / "wydruki")  # ..\wydruki
    default_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    sub = None
    try:
        sub = get_subiekt()

        # dane referencyjne
        wzorce = fetch_wzorce_fs(sub)
        wz_by_id = {int(w["wzw_Id"]): str(w["wzw_Nazwa"]) for w in wzorce}
        kontrahenci = fetch_kontrahenci_basic(sub)

        # wybór dokumentów
        docs = select_docs_prev_month(sub.Dokumenty, typ=2) # 2 = FS

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
        # delay = ask_delay_seconds(default=5) or 0

        # drukowanie/export
        out_dir = choose_output_dir(default_dir)
        for i, d in enumerate(docs, start=1):
            wzw_id = wz_kontr.get(int(d.KontrahentId))
            if not wzw_id:
                logger.warning("Pomijam %s – brak wybranego wzorca.", getattr(d, "NumerPelny", "<bez numeru>"))
                continue
            wz_name = wz_by_id.get(wzw_id, f"wzorzec {wzw_id}")
            fname = safe_filename(str(d.NumerPelny))
            fullpath = str(out_dir / fname)
            logger.info("Exportuję (%d/%d) %s wzorem %s do pliku %s", i, len(docs), d.NumerPelny, wz_name, fullpath)
            d.DrukujDoPlikuWgWzorca(wzw_id, fullpath, 0)  # 0 = PDF
            # drukuj_wg_ustawien(d, wzw_id=wzw_id, printer_name=printer_name, ilosc_kopii=1)

            # if delay and i < len(docs):
            #     logger.info("  ... czekam %d sekund ...", delay)
            #     time.sleep(delay)

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
    logfile = logowanie.setup_logging(LOG_PREFIX = LOG_PREFIX)
    print(f"Start aplikacji. Logi zapisuję do pliku: {logfile}")
    try:
        main()
    finally:
        show_completion_dialog(logfile=logfile, logs_dir="logs")
