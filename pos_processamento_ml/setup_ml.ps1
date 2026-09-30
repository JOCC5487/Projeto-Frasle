$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path '.\.venv\Scripts\python.exe')) {
    throw 'Ambiente .venv nao encontrado. Execute primeiro o setup_windows.ps1 do projeto principal.'
}
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r .\requirements_ml.txt
New-Item -ItemType Directory -Force -Path '.\referencias_ml\logos' | Out-Null
Write-Host 'Dependencias de ML instaladas. Adicione logos de referencia e execute .\executar_pos_processamento.ps1'
