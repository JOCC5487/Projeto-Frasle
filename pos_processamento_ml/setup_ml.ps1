$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $ProjectRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $Python)) {
    throw 'Ambiente .venv nao encontrado. Execute primeiro .\setup_windows.ps1 na raiz do projeto.'
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r (Join-Path $PSScriptRoot 'requirements_ml.txt')
New-Item -ItemType Directory -Force -Path (Join-Path $PSScriptRoot 'referencias_ml\logos') | Out-Null
Write-Host 'Dependencias de ML instaladas. Execute: .\pos_processamento_ml\executar_pos_processamento.ps1'
