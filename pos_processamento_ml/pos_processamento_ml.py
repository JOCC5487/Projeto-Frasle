from __future__ import annotations
import argparse
import csv
import hashlib
import io
import json
import logging
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

import yaml
from PIL import Image, ImageOps, UnidentifiedImageError
import imagehash

LOG = logging.getLogger("pos_ml")


def now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def init_schema(c: sqlite3.Connection):
    c.executescript("""
    PRAGMA foreign_keys=ON;
    PRAGMA journal_mode=WAL;
    PRAGMA synchronous=NORMAL;
    CREATE TABLE IF NOT EXISTS classificacoes_ml (
        id_classificacao INTEGER PRIMARY KEY,
        id_imagem INTEGER NOT NULL,
        tipo_visual TEXT,
        categoria_peca TEXT,
        categoria_final TEXT NOT NULL,
        confianca REAL,
        segunda_categoria TEXT,
        confianca_segunda REAL,
        status_processamento TEXT NOT NULL,
        ignorar INTEGER NOT NULL DEFAULT 0 CHECK(ignorar IN (0,1)),
        motivo_ignorar TEXT,
        modelo TEXT NOT NULL,
        versao_modelo TEXT NOT NULL,
        hash_config TEXT NOT NULL,
        data_classificacao TEXT NOT NULL,
        revisado_manualmente INTEGER NOT NULL DEFAULT 0 CHECK(revisado_manualmente IN (0,1)),
        categoria_revisada TEXT,
        observacao_revisao TEXT,
        erro TEXT,
        FOREIGN KEY(id_imagem) REFERENCES imagens(id_imagem) ON DELETE CASCADE,
        UNIQUE(id_imagem, modelo, versao_modelo, hash_config)
    );
    CREATE TABLE IF NOT EXISTS exemplos_treinamento (
        id_exemplo INTEGER PRIMARY KEY,
        id_imagem INTEGER NOT NULL UNIQUE,
        categoria_correta TEXT NOT NULL,
        origem_rotulo TEXT NOT NULL DEFAULT 'revisao_manual',
        incluir_treinamento INTEGER NOT NULL DEFAULT 1 CHECK(incluir_treinamento IN (0,1)),
        data_rotulacao TEXT NOT NULL,
        observacao TEXT,
        FOREIGN KEY(id_imagem) REFERENCES imagens(id_imagem) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS execucoes_ml (
        id_execucao_ml INTEGER PRIMARY KEY,
        inicio TEXT NOT NULL,
        fim TEXT,
        status TEXT NOT NULL,
        modelo TEXT,
        hash_config TEXT,
        avaliadas INTEGER DEFAULT 0,
        classificadas INTEGER DEFAULT 0,
        ignoradas INTEGER DEFAULT 0,
        pendentes INTEGER DEFAULT 0,
        erros INTEGER DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ix_ml_imagem ON classificacoes_ml(id_imagem);
    CREATE INDEX IF NOT EXISTS ix_ml_final ON classificacoes_ml(categoria_final, status_processamento);
    CREATE INDEX IF NOT EXISTS ix_ml_ignorar ON classificacoes_ml(ignorar);
    DROP VIEW IF EXISTS vw_imagens_pos_processadas;
    CREATE VIEW vw_imagens_pos_processadas AS
    SELECT i.id_imagem, i.numero_ensaio, a.arquivo_nome, a.caminho_origem,
           i.aba_origem, i.etapa_ensaio, i.arquivo_imagem,
           m.tipo_visual, m.categoria_peca,
           COALESCE(m.categoria_revisada,m.categoria_final) AS categoria_utilizada,
           m.categoria_final AS categoria_automatica, m.confianca,
           m.segunda_categoria, m.confianca_segunda,
           m.status_processamento, m.ignorar, m.motivo_ignorar,
           m.revisado_manualmente, m.modelo, m.versao_modelo,
           m.data_classificacao
      FROM classificacoes_ml m
      JOIN imagens i ON i.id_imagem=m.id_imagem
      JOIN arquivos a ON a.id=i.arquivo_id;
    """)


def config_hash(cfg):
    material=json.dumps(cfg, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(material).hexdigest()[:16]


def load_image(row):
    blob=row["imagem_blob"]
    if blob:
        return Image.open(io.BytesIO(blob))
    path=row["arquivo_imagem"]
    if path and Path(path).exists():
        return Image.open(path)
    raise FileNotFoundError("Imagem ausente no BLOB e no caminho local")


def normalize_image(im):
    im=ImageOps.exif_transpose(im)
    if getattr(im, "is_animated", False):
        im.seek(0)
    return im.convert("RGB")


def load_logo_hashes(folder):
    refs=[]
    folder=Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        try:
            with Image.open(p) as im:
                refs.append((p.name, imagehash.phash(normalize_image(im))))
        except Exception:
            LOG.warning("Referencia de logo invalida: %s", p)
    return refs


def logo_match(im, refs, max_distance):
    if not refs:
        return None
    h=imagehash.phash(im)
    best=min(((h-rh,name) for name,rh in refs), default=None)
    if best and best[0] <= max_distance:
        return {"name":best[1], "distance":int(best[0])}
    return None


def build_classifier(model_name, device):
    from transformers import pipeline
    return pipeline(task="zero-shot-image-classification", model=model_name, device=device)


def rank(classifier, im, labels):
    result=classifier(im, candidate_labels=list(labels.values()))
    reverse={v:k for k,v in labels.items()}
    ranked=[]
    for item in result:
        ranked.append((reverse[item["label"]], float(item["score"])))
    return ranked


def choose(ranked, auto_threshold, review_threshold):
    first=ranked[0]
    second=ranked[1] if len(ranked)>1 else (None,0.0)
    margin=first[1]-second[1]
    if first[1] >= auto_threshold and margin >= 0.08:
        status="automatica"
    elif first[1] >= review_threshold:
        status="revisao_pendente"
    else:
        status="revisao_pendente"
    return first, second, status


def classify_one(im, classifier, cfg, logo_refs):
    logo=logo_match(im, logo_refs, int(cfg["logos"]["distancia_phash_maxima"]))
    if logo:
        return dict(tipo="logo", peca=None, final="ignorar_logo", conf=1.0,
                    second=None, second_conf=None, status="automatica",
                    ignore=1, reason=f"Similar a {logo['name']}; distancia pHash={logo['distance']}")

    visual=rank(classifier, im, cfg["tipos_visuais"])
    v1,v2,vstatus=choose(visual, float(cfg["limiares"]["automatico"]), float(cfg["limiares"]["revisao"]))

    ignore_types=set(cfg.get("tipos_ignorados", []))
    if v1[0] in ignore_types and v1[1] >= float(cfg["limiares"]["automatico"]):
        return dict(tipo=v1[0], peca=None, final="ignorar_"+v1[0], conf=v1[1],
                    second=v2[0], second_conf=v2[1], status=vstatus,
                    ignore=1, reason="Tipo visual configurado para ser ignorado")

    if v1[0] == "fotografia_peca":
        pieces=rank(classifier, im, cfg["categorias_pecas"])
        p1,p2,pstatus=choose(pieces, float(cfg["limiares"]["automatico"]), float(cfg["limiares"]["revisao"]))
        final=p1[0] if pstatus=="automatica" else "peca_revisao_pendente"
        return dict(tipo=v1[0], peca=p1[0], final=final, conf=p1[1],
                    second=p2[0], second_conf=p2[1], status=pstatus,
                    ignore=0, reason=None)

    final=v1[0] if vstatus=="automatica" else "revisao_pendente"
    return dict(tipo=v1[0], peca=None, final=final, conf=v1[1],
                second=v2[0], second_conf=v2[1], status=vstatus,
                ignore=0, reason=None)


def export_review(c, output):
    output=Path(output); output.parent.mkdir(parents=True, exist_ok=True)
    rows=c.execute("""
      SELECT * FROM vw_imagens_pos_processadas
       WHERE status_processamento='revisao_pendente' OR revisado_manualmente=0
       ORDER BY confianca ASC, id_imagem
    """).fetchall()
    if not rows:
        return
    with output.open("w", newline="", encoding="utf-8-sig") as f:
        w=csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader(); w.writerows(dict(r) for r in rows)


def main():
    ap=argparse.ArgumentParser(description="Pos-processamento ML das imagens extraidas")
    ap.add_argument("--config", default="config_ml.yaml")
    ap.add_argument("--reprocessar", action="store_true", help="Reclassifica mesmo se a mesma versao ja existir")
    ap.add_argument("--limite", type=int, default=0, help="Limite opcional para teste; 0 processa tudo")
    args=ap.parse_args()

    cfg=yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    db=Path(cfg["banco"])
    if not db.exists():
        raise SystemExit(f"Banco nao encontrado: {db}")
    log_path=Path(cfg["log"]); log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path,encoding="utf-8"), logging.StreamHandler(sys.stdout)])

    con=sqlite3.connect(db); con.row_factory=sqlite3.Row; init_schema(con)
    model=cfg["modelo"]; version=str(cfg["versao_modelo"]); chash=config_hash(cfg)
    run=con.execute("INSERT INTO execucoes_ml(inicio,status,modelo,hash_config) VALUES(?,?,?,?)",
                    (now(),"executando",model,chash)).lastrowid; con.commit()

    refs=load_logo_hashes(cfg["logos"]["pasta_referencias"])
    LOG.info("Referencias de logos carregadas: %s", len(refs))
    LOG.info("Carregando modelo %s. Na primeira vez pode haver download.", model)
    classifier=build_classifier(model, int(cfg.get("device",-1)))

    sql="""SELECT i.id_imagem,i.imagem_blob,i.arquivo_imagem
             FROM imagens i
            WHERE (i.imagem_blob IS NOT NULL OR i.arquivo_imagem IS NOT NULL)"""
    params=[]
    if not args.reprocessar:
        sql += " AND NOT EXISTS (SELECT 1 FROM classificacoes_ml m WHERE m.id_imagem=i.id_imagem AND m.modelo=? AND m.versao_modelo=? AND m.hash_config=?)"
        params=[model,version,chash]
    sql += " ORDER BY i.id_imagem"
    if args.limite>0:
        sql += " LIMIT ?"; params.append(args.limite)
    rows=con.execute(sql,params).fetchall()
    stats=dict(avaliadas=0,classificadas=0,ignoradas=0,pendentes=0,erros=0)

    for n,row in enumerate(rows,1):
        stats["avaliadas"]+=1
        try:
            with load_image(row) as raw:
                im=normalize_image(raw)
                result=classify_one(im,classifier,cfg,refs)
            con.execute("""INSERT INTO classificacoes_ml(
                id_imagem,tipo_visual,categoria_peca,categoria_final,confianca,
                segunda_categoria,confianca_segunda,status_processamento,ignorar,
                motivo_ignorar,modelo,versao_modelo,hash_config,data_classificacao,erro)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,NULL)
                ON CONFLICT(id_imagem,modelo,versao_modelo,hash_config) DO UPDATE SET
                tipo_visual=excluded.tipo_visual,categoria_peca=excluded.categoria_peca,
                categoria_final=excluded.categoria_final,confianca=excluded.confianca,
                segunda_categoria=excluded.segunda_categoria,confianca_segunda=excluded.confianca_segunda,
                status_processamento=excluded.status_processamento,ignorar=excluded.ignorar,
                motivo_ignorar=excluded.motivo_ignorar,data_classificacao=excluded.data_classificacao,erro=NULL""",
                (row["id_imagem"],result["tipo"],result["peca"],result["final"],result["conf"],
                 result["second"],result["second_conf"],result["status"],result["ignore"],result["reason"],
                 model,version,chash,now()))
            stats["classificadas"]+=1
            stats["ignoradas"]+=int(result["ignore"])
            stats["pendentes"]+=int(result["status"]=="revisao_pendente")
        except Exception as e:
            LOG.exception("Falha na imagem %s",row["id_imagem"]); stats["erros"]+=1
            con.execute("""INSERT INTO classificacoes_ml(id_imagem,categoria_final,status_processamento,
                ignorar,modelo,versao_modelo,hash_config,data_classificacao,erro)
                VALUES(?,?,?,?,?,?,?,?,?)
                ON CONFLICT(id_imagem,modelo,versao_modelo,hash_config) DO UPDATE SET
                status_processamento='erro',erro=excluded.erro,data_classificacao=excluded.data_classificacao""",
                (row["id_imagem"],"erro_processamento","erro",0,model,version,chash,now(),str(e)[:2000]))
        if n % int(cfg.get("commit_a_cada",10)) == 0:
            con.commit(); LOG.info("Progresso: %s/%s",n,len(rows))
    con.commit()
    export_review(con,cfg["relatorio_revisao"])
    status="concluido_com_erros" if stats["erros"] else "concluido"
    con.execute("""UPDATE execucoes_ml SET fim=?,status=?,avaliadas=?,classificadas=?,ignoradas=?,pendentes=?,erros=? WHERE id_execucao_ml=?""",
                (now(),status,stats["avaliadas"],stats["classificadas"],stats["ignoradas"],stats["pendentes"],stats["erros"],run))
    con.commit(); con.execute("PRAGMA optimize"); con.close()
    print(json.dumps({"status":status,**stats},ensure_ascii=False,indent=2))

if __name__=="__main__":
    main()
