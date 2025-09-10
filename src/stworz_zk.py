import logging
from pathlib import Path
from gui import choose_output_dir, show_completion_dialog
from utils import get_subiekt_default_login, safe_filename
import logowanie
from pywintypes import com_error
import pythoncom

logger = logging.getLogger(__name__)

# Listę wzorców można znaleźć w SQL:
# SELECT wzw_Id, wzw_Nazwa FROM wy_Wzorzec wz
# join wy_Typ wt on wz.wzw_Typ = wt.wtp_Id
# where wt.wtp_Nazwa = 'Zamówienie od klienta'
id_wzorca = 440  # ID wzorca wydruku zdefiniowanego w Subiekcie

# Nazwa pliku logu: prefiks + data
LOG_PREFIX = "FS_"


def main():
    pythoncom.CoInitialize()
    sub = None
    
    try:
        sub = get_subiekt_default_login()
        nowy_dok = sub.Dokumenty.Dodaj(-8)
        logger.info("Wyświetlam okno do tworzenia nowego dokumentu ZK...")
        nowy_dok.Wyswietl()
        if not nowy_dok.NumerPelny.startswith("ZK"):
            print("Anulowano przez użytkownika. Dokument nie został utworzony.")
            return
        default_dir = (Path.cwd().parent / "wydruki")  # ..\wydruki
        default_dir.mkdir(parents=True, exist_ok=True)
        print("Wyświetlam okno wyboru katalogu docelowego...")
        out_dir = choose_output_dir(default_dir)
        fname = safe_filename(str(nowy_dok.NumerPelny))
        fullpath = str(out_dir / fname)
        logger.info("Exportuję nowy dokument %s do pliku %s", nowy_dok.NumerPelny, fullpath)
        nowy_dok.DrukujDoPlikuWgWzorca(id_wzorca, fullpath, 0)  # 0 = PDF
    
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
    # try:
    main()
    # finally:
    #     show_completion_dialog(logfile=logfile, logs_dir="logs")