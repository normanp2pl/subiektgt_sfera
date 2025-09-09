.\.venv\Scripts\activate

# Ustaw zmienne środowiskowe dla połączenia z bazą danych i operatora
# UWAGA - Zmiany nie są konieczne, służą tylko do automatycznego logowania. 
# Można ich nie ustawiać aby było bezpieczniej - wtedy po każdym uruchomieniu skryptu trzeba będzie się zalogować
# (o ile nie mamy ustawionego automatycznego logowania w programie serwisowym Insertu).
#
$env:SFERA_SQL_SERVER = "192.168.1.123"
$env:SFERA_SQL_LOGIN = "sa"
$env:SFERA_SQL_PASSWORD = "!Maniutek123"
$env:SFERA_SQL_DB = "DevNorman"
$env:SFERA_OPERATOR = "Szef"
$env:SFERA_OPERATOR_PASSWORD = "szef"

python .\launcher.py