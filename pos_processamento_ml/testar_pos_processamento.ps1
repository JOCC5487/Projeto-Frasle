$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
& .\.venv\Scripts\python.exe .\pos_processamento_ml.py --config .\config_ml.yaml --limite 20
