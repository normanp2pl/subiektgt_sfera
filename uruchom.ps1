.\.venv\Scripts\activate

# Ustaw zmienne środowiskowe dla połączenia z bazą danych i operatora
# UWAGA - Zmiany nie są konieczne, służą tylko do automatycznego logowania. 
# Można ich nie ustawiać aby było bezpieczniej - wtedy po każdym uruchomieniu skryptu trzeba będzie się zalogować
# (o ile nie mamy ustawionego automatycznego logowania w programie serwisowym Insertu).
#
# $env:SFERA_SQL_SERVER = "192.168.1.123\INSERTGT"
# $env:SFERA_SQL_LOGIN = "sa"
# $env:SFERA_SQL_PASSWORD = "TwojeHaslo"
# $env:SFERA_SQL_DB = "SubiektGT"
# $env:SFERA_OPERATOR = "admin"
# $env:SFERA_OPERATOR_PASSWORD = "admin"

python .\zmiana_mm.py --new-date "2025-10-01" # --dry-run