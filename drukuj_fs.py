import argparse
import pywintypes
import os
from datetime import datetime, timedelta, time
import pythoncom
import win32com.client as win32
from pywintypes import com_error

# Stałe
gtaUruchom = 1   # (4 = gtaUruchomWTle jeśli chcesz bez interfejsu)

def to_com_time(dt: datetime):
    return pywintypes.Time(dt)

def parse_args():
    ap = argparse.ArgumentParser()
    # ap.add_argument("--new-date", required=True)
    ap.add_argument("--dry-run", action="store_true")
    return ap.parse_args()

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

        sub.MagazynId = 1
        print(dir(sub.Kontrahenci))
        dok = sub.Dokumenty.Wybierz()
        kontrahenci = sub.Kontrahenci.Wczytaj() 
        for k in kontrahenci.ZaznaczeniKontrahenci():
            print(k.Nazwa)
        print(dir(kontrahenci))
        dok.FiltrTyp = 2 # Wsztystkie możliwe Faktury Sprzedaży
        dok.FiltrOkres = 20 # gtaFiltrOkresDowolnyMiesiac 
        # za ostatni miesiąc czy to co poniżej będzie ok 
        # today = datetime.now()
        # first_this = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        # last_prev = first_this - timedelta(seconds=1)  # 23:59:59 dnia poprzedzającego
        # com_time = to_com_time(last_prev)

        dt = datetime(2025, 9, 4, 12, 0, 0)
        com_time = to_com_time(dt)

        dok.FiltrOkresUstawDowolnyMiesiac(com_time) # poczatek miesiąca poprzedniego
        dok.MultiSelekcja = True
        dok.Wyswietl()
        
        for p in dok.ZaznaczoneDokumenty():   # iteracja po zaznaczonych dokumentach
            print(dir(p))
            print(p.KontrahentId)
        #     if not p.NumerPelny.startswith("MM"):
        #         continue
        #     printout = f"Zmieniam date dokumentu {p.NumerPelny} o wartości {p.WartoscNetto} " \
        #                f"z dnia {p.DataWystawienia.date().isoformat()} na {new_date.date().isoformat()}"
        #     p.DataWystawienia = new_date
        #     if not args.dry_run:    
        #         print(printout)
        #         # if not p.SkutekMagazynowy:
        #         #     #TO-DO: dokumenty bez skutku magazynowego - ustawiać im skutek czy zostawić bez zmiany?
        #         #     print("  Dokument nie ma skutku magazynowego, pomijam")
        #         # else:
        #         p.Zapisz()
        #     else:
        #         print("DRY RUN:", printout)
    except com_error as e:
        print("Błąd COM:", e)
    finally:
        try:
            sub.Zakoncz()
        except Exception:
            pass

if __name__ == "__main__":
    main()
