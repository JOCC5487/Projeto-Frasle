$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$py = Get-Command py -ErrorAction SilentlyContinue
$python = Get-Command python -ErrorAction SilentlyContinue
if ($py) { & py -3 -m venv .venv }
elseif ($python) { & python -m venv .venv }
else { throw 'Python 3 nao encontrado. Instale-o e marque Add Python to PATH.' }
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
Write-Host 'Ambiente pronto. Execute: .\executar.ps1'
