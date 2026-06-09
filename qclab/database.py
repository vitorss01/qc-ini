"""
database.py
-----------
Camada de dados em SQLite (SQL puro). Cria o schema, popula com os dados reais
do laboratório e expõe funções de consulta/atualização usadas pela aplicação.
"""

import os
import sqlite3
import hashlib
from datetime import datetime

import seed_data as seed

DB_PATH = os.environ.get("QCLAB_DB", os.path.join(os.path.dirname(__file__), "qclab.db"))


# --------------------------------------------------------------------------- #
# Conexão
# --------------------------------------------------------------------------- #
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


# --------------------------------------------------------------------------- #
# Schema (DDL)
# --------------------------------------------------------------------------- #
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    login        TEXT PRIMARY KEY,
    name         TEXT NOT NULL,
    role         TEXT,
    access_level INTEGER DEFAULT 1,
    pwd_hash     TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS config (
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS analytes (
    name     TEXT PRIMARY KEY,
    ord      INTEGER,
    decimals INTEGER DEFAULT 2
);

CREATE TABLE IF NOT EXISTS reference (
    analyte  TEXT,
    level    INTEGER,
    mean     REAL,
    sd       REAL,
    lote     TEXT,
    validade TEXT,
    PRIMARY KEY (analyte, level)
);

CREATE TABLE IF NOT EXISTS specs (
    analyte  TEXT PRIMARY KEY,
    clia_tea REAL,
    fab_cv   REAL,
    cvw      REAL,
    cvg      REAL,
    perf     TEXT
);

CREATE TABLE IF NOT EXISTS results (
    result_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    level      INTEGER,
    lote       TEXT,
    seq        INTEGER,
    run_date   TEXT,
    run_time   TEXT,
    analyte    TEXT,
    value      REAL,
    is_nc      INTEGER DEFAULT 0,
    nc_reason  TEXT,
    nc_by      TEXT,
    nc_at      TEXT,
    created_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_results_al ON results(analyte, level, seq);

CREATE TABLE IF NOT EXISTS external_qc (
    analyte    TEXT,
    level      INTEGER,
    ini_value  REAL,
    peer_value REAL,
    peer_sd    REAL,
    PRIMARY KEY (analyte, level)
);
"""


# --------------------------------------------------------------------------- #
# Inicialização / seeding
# --------------------------------------------------------------------------- #
def init_db(force: bool = False, demo: bool = True):
    """Cria o schema e popula. Se force=True, recria o banco do zero."""
    if force and os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(SCHEMA)

    # já populado?
    cur.execute("SELECT COUNT(*) AS c FROM analytes")
    if cur.fetchone()["c"] > 0 and not force:
        conn.commit()
        return conn

    # analitos
    for i, a in enumerate(seed.ANALYTES):
        cur.execute("INSERT OR REPLACE INTO analytes(name, ord, decimals) VALUES (?,?,?)",
                    (a, i, 4 if a in ("RET#", "BASO#", "NRBC#", "EO#", "MONO#") else 2))

    # usuários
    for u in seed.USERS:
        cur.execute(
            "INSERT OR REPLACE INTO users(login, name, role, access_level, pwd_hash) VALUES (?,?,?,?,?)",
            (u["login"], u["name"], u["role"], u["access_level"], sha256(u["password"])))

    # config
    for k, v in seed.CONFIG.items():
        cur.execute("INSERT OR REPLACE INTO config(key, value) VALUES (?,?)", (k, str(v)))

    # referência (média/DP)
    for (a, lvl), (m, sd) in seed.REFERENCE.items():
        cur.execute(
            "INSERT OR REPLACE INTO reference(analyte, level, mean, sd, lote, validade) VALUES (?,?,?,?,?,?)",
            (a, lvl, m, sd, seed.LOTES.get(lvl, ""), seed.VALIDADE))

    # especificações
    for a, s in seed.SPECS.items():
        cur.execute(
            "INSERT OR REPLACE INTO specs(analyte, clia_tea, fab_cv, cvw, cvg, perf) VALUES (?,?,?,?,?,?)",
            (a, s["clia_tea"], s["fab_cv"], s["cvw"], s["cvg"], s["perf"]))

    # controle externo
    for (a, lvl), (ini, peer, psd) in seed.EXTERNAL_QC.items():
        cur.execute(
            "INSERT OR REPLACE INTO external_qc(analyte, level, ini_value, peer_value, peer_sd) VALUES (?,?,?,?,?)",
            (a, lvl, ini, peer, psd))

    # resultados de demonstração
    if demo:
        now = datetime.now().isoformat(timespec="seconds")
        for r in seed.gerar_dados_demo():
            cur.execute(
                """INSERT INTO results(level, lote, seq, run_date, run_time, analyte, value, created_at)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (r["level"], r["lote"], r["seq"], r["run_date"], r["run_time"],
                 r["analyte"], r["value"], now))

    conn.commit()
    return conn


# --------------------------------------------------------------------------- #
# Autenticação
# --------------------------------------------------------------------------- #
def authenticate(login: str, password: str):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE login = ?", (login,)).fetchone()
    if row and row["pwd_hash"] == sha256(password):
        return dict(row)
    return None


# --------------------------------------------------------------------------- #
# Consultas
# --------------------------------------------------------------------------- #
def get_config():
    conn = get_conn()
    return {r["key"]: r["value"] for r in conn.execute("SELECT * FROM config")}


def get_analytes():
    conn = get_conn()
    return [r["name"] for r in conn.execute("SELECT name FROM analytes ORDER BY ord")]


def get_decimals():
    conn = get_conn()
    return {r["name"]: r["decimals"] for r in conn.execute("SELECT name, decimals FROM analytes")}


def get_reference(analyte=None, level=None):
    conn = get_conn()
    q = "SELECT * FROM reference WHERE 1=1"
    p = []
    if analyte:
        q += " AND analyte = ?"; p.append(analyte)
    if level:
        q += " AND level = ?"; p.append(level)
    return [dict(r) for r in conn.execute(q, p)]


def get_spec(analyte):
    conn = get_conn()
    r = conn.execute("SELECT * FROM specs WHERE analyte = ?", (analyte,)).fetchone()
    return dict(r) if r else {}


def get_all_specs():
    conn = get_conn()
    return {r["analyte"]: dict(r) for r in conn.execute("SELECT * FROM specs")}


def get_results(analyte, level, only_valid=True):
    """Série ordenada por seq de um analito/nível. only_valid exclui pontos marcados como NC."""
    conn = get_conn()
    q = "SELECT * FROM results WHERE analyte = ? AND level = ?"
    if only_valid:
        q += " AND is_nc = 0"
    q += " ORDER BY seq, result_id"
    return [dict(r) for r in conn.execute(q, (analyte, level))]


def get_results_df_rows(level=None):
    conn = get_conn()
    q = "SELECT * FROM results"
    p = []
    if level:
        q += " WHERE level = ?"; p.append(level)
    q += " ORDER BY analyte, level, seq"
    return [dict(r) for r in conn.execute(q, p)]


def get_lotes():
    """Lista de lotes distintos presentes nos resultados (mais recentes primeiro)."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT lote FROM results WHERE lote IS NOT NULL AND lote <> '' "
        "ORDER BY lote DESC"
    ).fetchall()
    return [r["lote"] for r in rows]


def query_results(analyte=None, level=None, lote=None,
                  date_from=None, date_to=None):
    """
    Consulta flexível de resultados com filtros opcionais.
      analyte   : nome do analito (ou None = todos)
      level     : nível 1/2/3 (ou None = todos)
      lote      : lote exato (ou None = todos)
      date_from : data inicial 'YYYY-MM-DD' inclusiva (ou None)
      date_to   : data final  'YYYY-MM-DD' inclusiva (ou None)
    Retorna lista de dicts ordenada por data/seq.
    """
    conn = get_conn()
    q = "SELECT * FROM results WHERE 1=1"
    p = []
    if analyte:
        q += " AND analyte = ?"; p.append(analyte)
    if level:
        q += " AND level = ?"; p.append(level)
    if lote:
        q += " AND lote = ?"; p.append(lote)
    if date_from:
        q += " AND run_date >= ?"; p.append(date_from)
    if date_to:
        q += " AND run_date <= ?"; p.append(date_to)
    q += " ORDER BY run_date, seq, level, analyte"
    return [dict(r) for r in conn.execute(q, p)]


def get_external_bias(analyte, level):
    """Viés (fração) do controle externo: (INI - par)/par. None se não houver."""
    conn = get_conn()
    r = conn.execute("SELECT * FROM external_qc WHERE analyte = ? AND level = ?",
                     (analyte, level)).fetchone()
    if r and r["peer_value"]:
        return (r["ini_value"] - r["peer_value"]) / r["peer_value"]
    return None


# --------------------------------------------------------------------------- #
# Atualizações
# --------------------------------------------------------------------------- #
def insert_result(level, lote, seq, run_date, run_time, analyte, value):
    conn = get_conn()
    conn.execute(
        """INSERT INTO results(level, lote, seq, run_date, run_time, analyte, value, created_at)
           VALUES (?,?,?,?,?,?,?,?)""",
        (level, lote, seq, run_date, run_time, analyte, value,
         datetime.now().isoformat(timespec="seconds")))
    conn.commit()


def delete_result(result_id):
    """Exclui permanentemente um resultado pelo ID."""
    conn = get_conn()
    conn.execute("DELETE FROM results WHERE result_id = ?", (result_id,))
    conn.commit()


def delete_results(result_ids):
    """Exclui permanentemente vários resultados. Retorna a quantidade removida."""
    if not result_ids:
        return 0
    conn = get_conn()
    placeholders = ",".join("?" for _ in result_ids)
    cur = conn.execute(
        f"DELETE FROM results WHERE result_id IN ({placeholders})",
        list(result_ids))
    conn.commit()
    return cur.rowcount


def set_nc(result_id, is_nc, user_login, reason=""):
    conn = get_conn()
    conn.execute(
        "UPDATE results SET is_nc = ?, nc_reason = ?, nc_by = ?, nc_at = ? WHERE result_id = ?",
        (1 if is_nc else 0, reason, user_login if is_nc else None,
         datetime.now().isoformat(timespec="seconds") if is_nc else None, result_id))
    conn.commit()


def update_reference(analyte, level, mean, sd):
    conn = get_conn()
    conn.execute("UPDATE reference SET mean = ?, sd = ? WHERE analyte = ? AND level = ?",
                 (mean, sd, analyte, level))
    conn.commit()


def next_seq(level):
    conn = get_conn()
    r = conn.execute("SELECT MAX(seq) AS m FROM results WHERE level = ?", (level,)).fetchone()
    return (r["m"] or 0) + 1


if __name__ == "__main__":
    # Recria o banco de demonstração
    init_db(force=True, demo=True)
    c = get_conn()
    print("Analitos:", c.execute("SELECT COUNT(*) FROM analytes").fetchone()[0])
    print("Referências:", c.execute("SELECT COUNT(*) FROM reference").fetchone()[0])
    print("Resultados:", c.execute("SELECT COUNT(*) FROM results").fetchone()[0])
    print("Usuários:", [r[0] for r in c.execute("SELECT login FROM users")])
