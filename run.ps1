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

$ErrorActionPreference = "Stop"

# Ustaw bazę na folder pliku skryptu
$root = Split-Path -Parent $PSCommandPath
Set-Location -LiteralPath $root

# Konfiguracja
$venvDir   = ".venv"
$pyTag     = "-3.11-32"             # zmień na -3.12 jeśli chcesz
$entry     = "src\launcher.py"      # co uruchomić
$entryArgs = @()                    # np. @("--printer","Microsoft Print to PDF")

$venvPy = Join-Path $venvDir "Scripts\python.exe"

# jeśli venv nie istnieje -> setup
if (-not (Test-Path -LiteralPath $venvPy)) {
  if (Get-Command py -ErrorAction SilentlyContinue) {
    py $pyTag -m venv $venvDir
  } else {
    python -m venv $venvDir
  }

  & $venvPy -m pip install --upgrade pip

  if (Test-Path -LiteralPath "src\requirements.txt") {
    & $venvPy -m pip install -r requirements.txt
  }
}

# uruchom skrypt tym Pythonem z venv
& $venvPy $entry @entryArgs
