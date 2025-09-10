import logging
from datetime import datetime, time

import pythoncom
from pywintypes import com_error

import logowanie
from gui import ask_new_date_and_dryrun, show_completion_dialog
from utils import get_subiekt, select_docs_prev_month, to_com_time

logger = logging.getLogger(__name__)

# ===================== Stałe =====================
LOG_PREFIX = "MM_"   # prefiks nazwy pliku logu

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
        sub = get_subiekt()
        selected = select_docs_prev_month(sub.Dokumenty, typ=9)  # 9 = MM
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
    logfile = logowanie.setup_logging(LOG_PREFIX = LOG_PREFIX)  # <- tu powstaje logs/MM_YYYY-MM-DD.log
    print(f"Start aplikacji. Logi zapisuję do pliku: {logfile}")
    try:
        main()
    finally:
        show_completion_dialog(logfile=logfile, logs_dir="logs")
