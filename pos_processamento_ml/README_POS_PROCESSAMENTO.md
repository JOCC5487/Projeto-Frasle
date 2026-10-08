# Pós-processamento ML

Copie estes arquivos para a pasta principal do projeto, no mesmo nível de `main.py`.

## Arquivos

- `pos_processamento_ml.py`: classifica imagens do SQLite.
- `config_ml.yaml`: categorias, prompts e limiares.
- `requirements_ml.txt`: dependências adicionais.
- `setup_ml.ps1`: instala as dependências na `.venv` existente.
- `executar_pos_processamento.ps1`: executa tudo que ainda não foi classificado.
- `testar_pos_processamento.ps1`: processa no máximo 20 imagens.
- `referencias_ml/logos`: coloque exemplos reais de logos nessa pasta.

O pós-processamento não apaga imagens. Logos recebem `ignorar=1` e permanecem auditáveis.