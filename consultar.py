import argparse, sqlite3
from pathlib import Path

p=argparse.ArgumentParser(description='Consulta a base historica de imagens')
p.add_argument('--db', default='./saida/historico_dinamometro.sqlite')
p.add_argument('--dl'); p.add_argument('--arquivo'); p.add_argument('--aba'); p.add_argument('--etapa'); p.add_argument('--classificacao'); p.add_argument('--pendentes', action='store_true')
a=p.parse_args()
where=[]; vals=[]
for col,val in [('numero_ensaio',a.dl),('arquivo_nome',a.arquivo),('aba_origem',a.aba),('etapa_ensaio',a.etapa),('classificacao_imagem',a.classificacao)]:
    if val: where.append(f'{col} LIKE ?'); vals.append(f'%{val}%')
if a.pendentes: where.append("status_classificacao='pendente'")
sql='''SELECT id_imagem, numero_ensaio, arquivo_nome, aba_origem, etapa_ensaio,
classificacao_imagem, status_classificacao, arquivo_imagem FROM vw_imagens_consulta'''
if where: sql += ' WHERE ' + ' AND '.join(where)
sql += ' ORDER BY arquivo_nome, aba_origem, id_imagem'
con=sqlite3.connect(Path(a.db)); con.row_factory=sqlite3.Row
rows=con.execute(sql, vals).fetchall()
for r in rows: print(' | '.join(str(r[k] or '') for k in r.keys()))
print(f'\nTotal: {len(rows)}')