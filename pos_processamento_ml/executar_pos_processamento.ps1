$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path '.\.venv\Scripts\python.exe')) { throw 'Ambiente .venv ausente.' }
& .\.venv\Scripts\python.exe .\pos_processamento_ml.py --config .\config_ml.yaml
