$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $Python)) { throw 'Ambiente .venv ausente. Execute .\setup_windows.ps1 na raiz.' }
Set-Location $ProjectRoot
& $Python (Join-Path $PSScriptRoot 'pos_processamento_ml.py') --config (Join-Path $PSScriptRoot 'config_ml.yaml') --limite 20
