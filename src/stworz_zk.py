import logging

import pythoncom
from pywintypes import com_error

import logowanie
from utils import get_subiekt, run_sql

logger = logging.getLogger(__name__)

kategoria = "Magazyn"

LOG_PREFIX = "ZK_"


def get_kategoria_id(sub, nazwa: str) -> int:
    """Pobiera ID kategorii o podanej nazwie."""
    kategoria = run_sql(sub, f"""SELECT kat_Id
                                   FROM sl_Kategoria
                                  WHERE kat_Nazwa = '{nazwa}'""")
    if not kategoria:
        logger.warning("Nie znaleziono kategorii o nazwie '%s', zostanie użyta domyślna.", nazwa)
        return None
    return kategoria[0]['kat_Id']


def main():
    pythoncom.CoInitialize()
    sub = None
    
    try:
        sub = get_subiekt()
        nowy_dok = sub.Dokumenty.Dodaj(-8)
        logger.info("Wyświetlam okno do tworzenia nowego dokumentu ZK...")
        kategoria_id = get_kategoria_id(sub, kategoria)
        if kategoria_id:
            nowy_dok.KategoriaId = kategoria_id
        nowy_dok.Tytul = "Tutaj też możemy wpisać co nam się podoba"
        nowy_dok.Uwagi = "A to są uwagi do dokumentu\r\nMożna tu wpisać coś dłuższego\r\ni wielolinijkowego."
        nowy_dok.Wyswietl()
        if not nowy_dok.NumerPelny.startswith("ZK"):
            logger.warning("Anulowano przez użytkownika. Dokument nie został utworzony.")
            return
    
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
    logger.info(f"Start aplikacji. Logi zapisuję do pliku: {logfile}")
    # try:
    main()
    # finally:
    #     show_completion_dialog(logfile=logfile, logs_dir="logs")