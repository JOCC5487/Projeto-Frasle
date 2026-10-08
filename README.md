# Base historica de imagens de ensaios de dinamometro

Projeto local para VS Code que percorre subpastas, processa somente planilhas cujo nome contenha `DC` ou `DL` (sem diferenciar maiusculas/minusculas), extrai imagens de todas as abas, coleta contexto/cabecalho, salva os binarios no SQLite e opcionalmente em uma pasta estruturada.

## Inicio rapido no Windows

1. Aguarde o OneDrive sincronizar os arquivos. Arquivos somente online serao baixados pelo proprio Windows quando abertos.
2. Abra esta pasta no VS Code.
3. No PowerShell integrado:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup_windows.ps1
.\executar.ps1
```

O script de setup nao exige Python 3.13: usa qualquer Python 3 instalado via `py -3` ou `python`.

## Saidas

- `saida/historico_dinamometro.sqlite`: banco SQLite ja incluido, inicializado e pronto para receber os dados.
- `saida/imagens/<DL-ou-sem_dl>/<arquivo>/...`: copia estruturada das imagens.
- `saida/relatorios/execucao_<id>.json` e `.csv`: resumo e detalhes.
- `saida/logs/processamento.log`: log completo.

## Estrutura do banco

- `arquivos`: uma linha por planilha e estado do ultimo processamento.
- `imagens`: uma linha por imagem, incluindo `imagem_blob`, hash, aba, ancora, contexto, etapa e classificacao.
- `execucoes`: resumo de cada rodada.
- `erros`: falhas isoladas, sem interromper a varredura.
- `categorias`: catalogo das classificacoes permitidas.
- `metadados`: versao do esquema do banco.
- `vw_imagens_consulta`: visao organizada para consulta.
- `vw_resumo_arquivos`: resumo de processamento e pendencias por arquivo.

A chave unica `(arquivo_id, referencia_imagem, hash_imagem)` impede repeticao na reexecucao. Se o arquivo mudar, ele e reprocessado, e imagens que deixaram de existir sao removidas somente apos um processamento bem-sucedido.

## Consultas

```powershell
.\.venv\Scripts\python.exe .\consultar.py --dl DL123
.\.venv\Scripts\python.exe .\consultar.py --aba Fotos
.\.venv\Scripts\python.exe .\consultar.py --classificacao avaria
.\.venv\Scripts\python.exe .\consultar.py --pendentes
```

## Classificacao

A classificacao inicial usa palavras configuraveis encontradas no nome da aba, celulas proximas da imagem e cabecalho. Sem evidencia suficiente, grava `nao_classificada` com status `pendente`. Isso evita inventar categorias. A etapa usa o mesmo contexto e, se nao houver indicio, recebe o nome da aba.

## Formatos

- `.xlsx`, `.xlsm`, `.xltx`, `.xltm`: leitura direta OOXML, sem abrir o Excel.
- `.xls` e `.xlsb`: fallback via automacao COM, requer Microsoft Excel instalado no Windows. O COM exporta imagens como PNG.

Feche planilhas abertas antes da execucao. O programa continua apos erros e registra cada falha.