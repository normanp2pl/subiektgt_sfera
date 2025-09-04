Skrypt do seryjnej zmiany dat dokumentów MM.

Skrypt uruchamiany w PowerShell'u na tym samym stanowisku, na którym jest Subiekt GT z aktywną Sferą.

## Instalacja

1. Należy zainstalować Pythona 3.11 w wersji 32bit - https://www.python.org/ftp/python/3.11.0/python-3.11.0.exe  
   **WAŻNE:** podczas instalacji zaznacz opcję dodania Pythona do PATH.
2. Uruchom PowerShell i przejdź do katalogu zawierającego to repozytorium.
3. W pliku `uruchom.ps1` popraw zmienne środowiskowe, aby pasowały do Twojej instalacji.
4. Uruchom skrypt `przygotuj.ps1`, który zainstaluje wymaganą bibliotekę Pythona.  
   Aby uruchamiać "niezaufane" skrypty w Windows, najpierw uruchom polecenie:  
   `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
5. Skrypt jest gotowy do użycia. Przed uruchomieniem zamknij Subiekta.

## Uruchamianie

1. Przed uruchomieniem edytuj skrypt `uruchom.ps1`, ustawiając datę, na którą chcesz zmienić wybrane dokumenty.
2. Skrypt uruchamiasz przez wywołanie `uruchom.ps1`. Spowoduje to uruchomienie Subiekta i po zalogowaniu pokaże się lista dokumentów MM z zeszłego miesiąca. Po wybraniu i zatwierdzeniu klawiszem OK rozpocznie się zmiana.
3. Jeśli nie usunąłeś parametru `--dry-run` z pliku `uruchom.ps1`, edycja nie dojdzie do skutku – program pokaże tylko ewentualne zmiany.
4. Skrypt bez parametru `--dry-run` (możesz przed nim dać znak #, aby go tymczasowo wyłączyć) poprawi daty wybranych dokumentów na datę podaną w parametrze `--new-date`.