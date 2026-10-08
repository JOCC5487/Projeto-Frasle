$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path '.\.venv\Scripts\python.exe')) { throw 'Ambiente ausente. Execute .\setup_windows.ps1 primeiro.' }
& .\.venv\Scripts\python.exe .\main.py --config .\config.yaml