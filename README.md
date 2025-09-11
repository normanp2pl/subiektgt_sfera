Zestaw skryptów do wykorzystania w Subiekt GT z zainstalowanym dodatkiem Sfera.

Zawiera:
- Skrypt do seryjnej zmiany dat dokumentów MM;
- Skrypt do exportowania wybranych faktur FS, gdzie dla każdego odbiorcy można wybrać oddzielny wzorzec;
- Skrypt do seryjnego drukowania plików pdf z wybranego folderu (w trakcie pracy).

## Instalacja

1. Należy zainstalować Pythona 3.11 w wersji 32bit - https://www.python.org/ftp/python/3.11.0/python-3.11.0.exe
   **WAŻNE:** podczas instalacji zaznacz opcję dodania Pythona do PATH.
2. W pliku `run.ps1` popraw zmienne środowiskowe, aby pasowały do Twojej instalacji. Aktualnie tylko zmienna SFERA_SQL_DB jest używana. Aby użyć reszty należy również zmodyfikować metodę get_subiekt() w src\utils.py.
4. Login i hasło operatora do Subiekta jest pobierany z Menadżera poświadczeń Windows. Aby stworzyć poświadczenia należy wykonać:
```powershell
python -c "import utils; utils.cred_write()"
```
Program zapyta o login i hasło operatora które zostana zapisane w poświadczeniach Windows dla aktualnego użytkownika.
3. Uruchom skrypt `uruchom.bat`, który po pierwszym uruchomieniu zainstaluje wymaganą bibliotekę Pythona. Każde kolejne uruchomienie tego skryptu nie będzie już reinstalowało biblioteki.

## Uruchamianie

1. Uruchom skrypt `uruchom.bat`. Zostanie uruchomione okno które pozwoli wybrać skrypt do użycia.
