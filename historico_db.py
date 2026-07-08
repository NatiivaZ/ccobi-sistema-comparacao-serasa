"""Persistência do histórico do comparador.

Fica tudo aqui: metadados no SQLite e arquivos exportados em pasta separada.
"""

import json
import sqlite3
import uuid
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "historico_comparacoes.db"
PASTA_EXPORTACOES = BASE_DIR / "historico_exportacoes"


def _conectar():
    return sqlite3.connect(str(DB_PATH))


def _serializar_json(payload):
    """Serializa o que vai para o banco sem perder acentuação."""
    return json.dumps(payload, ensure_ascii=False)


def _listar_arquivos_exportados(pasta):
    """Lista os arquivos da execução em uma ordem estável para a tela."""
    if pasta and pasta.is_dir():
        return sorted(pasta.iterdir())
    return []


def _remover_pasta_exportacao(pasta):
    """Apaga a pasta da execução sem travar a exclusão por detalhe menor."""
    if not (pasta and isinstance(pasta, Path) and pasta.is_dir()):
        return
    import shutil

    try:
        shutil.rmtree(pasta)
    except Exception:
        pass


def init_db():
    """Cria a tabela de comparações se não existir."""
    conn = _conectar()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS comparacoes (
                id                TEXT PRIMARY KEY,
                data_hora         TEXT NOT NULL,
                nome_base_serasa  TEXT NOT NULL,
                nome_base_divida  TEXT NOT NULL,
                ano_analise       INTEGER,
                config_json       TEXT,
                total_serasa      INTEGER NOT NULL,
                total_divida      INTEGER NOT NULL,
                autos_em_ambas    INTEGER NOT NULL,
                autos_apenas_serasa INTEGER NOT NULL,
                autos_apenas_divida INTEGER NOT NULL,
                total_cpf_serasa  INTEGER,
                total_cpf_divida  INTEGER,
                cpf_em_ambas      INTEGER,
                resumo_json       TEXT,
                pasta_export      TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()


def save_run(
    resultados,
    nome_base_serasa,
    nome_base_divida,
    ano_analise,
    config,
    excel_dict=None,
):
    """Salva uma execução completa da comparação e devolve o `run_id`."""
    init_db()
    run_id = uuid.uuid4().hex
    data_hora = datetime.now().isoformat(sep=" ", timespec="seconds")

    resumo = {
        "total_registros_serasa": resultados.get("total_registros_serasa", 0),
        "total_registros_divida": resultados.get("total_registros_divida", 0),
        "total_autos_serasa": resultados.get("total_autos_serasa", 0),
        "total_autos_divida": resultados.get("total_autos_divida", 0),
        "autos_em_ambas": resultados.get("autos_em_ambas", 0),
        "autos_em_ambas_unicos": resultados.get("autos_em_ambas_unicos", 0),
        "autos_em_ambas_geral": resultados.get("autos_em_ambas_geral", 0),
        "autos_apenas_serasa": resultados.get("autos_apenas_serasa", 0),
        "autos_apenas_divida": resultados.get("autos_apenas_divida", 0),
        "cpf_em_ambas": resultados.get("cpf_em_ambas", 0),
        "cpf_apenas_serasa": resultados.get("cpf_apenas_serasa", 0),
        "cpf_apenas_divida": resultados.get("cpf_apenas_divida", 0),
        "total_cpf_serasa": resultados.get("total_cpf_serasa", 0),
        "total_cpf_divida": resultados.get("total_cpf_divida", 0),
    }

    pasta_run = PASTA_EXPORTACOES / run_id
    pasta_run.mkdir(parents=True, exist_ok=True)

    if excel_dict:
        for nome_arquivo, bytes_excel in excel_dict.items():
            if bytes_excel:
                (pasta_run / nome_arquivo).write_bytes(bytes_excel)

    config_json = _serializar_json(config)
    resumo_json = _serializar_json(resumo)

    conn = _conectar()
    try:
        conn.execute(
            """
            INSERT INTO comparacoes (
                id, data_hora, nome_base_serasa, nome_base_divida, ano_analise,
                config_json, total_serasa, total_divida,
                autos_em_ambas, autos_apenas_serasa, autos_apenas_divida,
                total_cpf_serasa, total_cpf_divida, cpf_em_ambas,
                resumo_json, pasta_export
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id, data_hora, nome_base_serasa, nome_base_divida, ano_analise,
                config_json,
                resumo["total_autos_serasa"], resumo["total_autos_divida"],
                resumo["autos_em_ambas"], resumo["autos_apenas_serasa"],
                resumo["autos_apenas_divida"],
                resumo["total_cpf_serasa"], resumo["total_cpf_divida"],
                resumo["cpf_em_ambas"],
                resumo_json, str(pasta_run),
            ),
        )
        conn.commit()
    finally:
        conn.close()

    return run_id


def list_runs():
    """Lista as execuções mais recentes primeiro para preencher o histórico."""
    init_db()
    conn = _conectar()
    try:
        cur = conn.execute(
            """
            SELECT id, data_hora, nome_base_serasa, nome_base_divida, ano_analise,
                   total_serasa, total_divida, autos_em_ambas,
                   autos_apenas_serasa, autos_apenas_divida
            FROM comparacoes
            ORDER BY data_hora DESC
            """
        )
        rows = cur.fetchall()
        return [
            {
                "id": r[0],
                "data_hora": r[1],
                "nome_base_serasa": r[2],
                "nome_base_divida": r[3],
                "ano_analise": r[4],
                "total_serasa": r[5],
                "total_divida": r[6],
                "autos_em_ambas": r[7],
                "autos_apenas_serasa": r[8],
                "autos_apenas_divida": r[9],
            }
            for r in rows
        ]
    finally:
        conn.close()


def get_run(run_id):
    """Busca uma execução completa pelo id."""
    conn = _conectar()
    try:
        cur = conn.execute(
            """
            SELECT data_hora, nome_base_serasa, nome_base_divida, ano_analise,
                   config_json, total_serasa, total_divida,
                   autos_em_ambas, autos_apenas_serasa, autos_apenas_divida,
                   total_cpf_serasa, total_cpf_divida, cpf_em_ambas,
                   resumo_json, pasta_export
            FROM comparacoes WHERE id = ?
            """,
            (run_id,),
        )
        row = cur.fetchone()
        if not row:
            return None

        pasta = Path(row[14]) if row[14] else None
        arquivos = _listar_arquivos_exportados(pasta)

        return {
            "id": run_id,
            "data_hora": row[0],
            "nome_base_serasa": row[1],
            "nome_base_divida": row[2],
            "ano_analise": row[3],
            "config": json.loads(row[4]) if row[4] else {},
            "total_serasa": row[5],
            "total_divida": row[6],
            "autos_em_ambas": row[7],
            "autos_apenas_serasa": row[8],
            "autos_apenas_divida": row[9],
            "total_cpf_serasa": row[10],
            "total_cpf_divida": row[11],
            "cpf_em_ambas": row[12],
            "resumo": json.loads(row[13]) if row[13] else {},
            "pasta_export": pasta,
            "arquivos": arquivos,
        }
    finally:
        conn.close()


def excluir_run(run_id):
    """Exclui a execução do banco e limpa os arquivos dela, se existirem."""
    run = get_run(run_id)
    if not run:
        return False
    conn = _conectar()
    try:
        conn.execute("DELETE FROM comparacoes WHERE id = ?", (run_id,))
        conn.commit()
    finally:
        conn.close()
    _remover_pasta_exportacao(run.get("pasta_export"))
    return True
