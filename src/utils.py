import os
from datetime import datetime, time, timedelta

import pywintypes
import win32com.client as win32


def to_com_time(dt: datetime):
    return pywintypes.Time(dt)


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
    sub = win32.Dispatch(gt.Uruchom(1, 1))
    return sub


def select_docs_prev_month(dok_manager, typ: int) -> list:
    """
    Otwiera okno wyboru dokumentów (wybrany typ, poprzedni miesiąc) i zwraca listę zaznaczonych.
    """
    dok = dok_manager.Wybierz()
    # dok.FiltrTyp = 2            # wszystkie możliwe Faktury Sprzedaży
    dok.FiltrTyp = typ         # 2 = FS, 9 MM
    dok.FiltrOkres = 20         # gtaFiltrOkresDowolnyMiesiac

    # poprzedni miesiac: końcówka dnia poprzedzającego 1-szy dzień bieżącego
    today = datetime.now()
    first_this = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_prev = first_this - timedelta(seconds=1)
    dok.FiltrOkresUstawDowolnyMiesiac(to_com_time(last_prev))

    dok.MultiSelekcja = True
    print("Otwieram okno wyboru dokumentów... zaznacz dokumenty do dalszej analizy i kliknij OK.")
    dok.Wyswietl()

    # zmaterializuj iterator
    docs = list(dok.ZaznaczoneDokumenty())
    print(f"Wybrano {len(docs)} dokumentów.")
    return docs
