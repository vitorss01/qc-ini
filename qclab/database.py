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


def _sample_bias(lab_value, peer_mean):
    """Bias% de uma amostra: ((lab-peer)/peer)*100. None se inválido / peer 0."""
    try:
        lab = float(lab_value)
        peer = float(peer_mean)
    except (TypeError, ValueError):
        return None
    if peer == 0:
        return None
    return ((lab - peer) / peer) * 100.0


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
    decimals INTEGER DEFAULT 2,
    area     TEXT DEFAULT 'Hematologia',
    unit     TEXT DEFAULT ''
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
    analyte    TEXT PRIMARY KEY,
    clia_tea   REAL,
    fab_cv     REAL,
    cvw        REAL,
    cvg        REAL,
    perf       TEXT,
    etp_pct    REAL,
    etp_source TEXT
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

-- ===================== CONTROLE EXTERNO (EQC) ===================== --
CREATE TABLE IF NOT EXISTS eqc_rounds (
    round_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    area         TEXT,
    analyte      TEXT,
    year         INTEGER,
    round_number INTEGER,
    round_date   TEXT,
    provider     TEXT,
    unit         TEXT,
    lote         TEXT,
    status       TEXT,
    notes        TEXT,
    report_ref   TEXT,
    created_at   TEXT,
    created_by   TEXT,
    updated_at   TEXT,
    updated_by   TEXT,
    UNIQUE (analyte, year, round_number)
);
CREATE INDEX IF NOT EXISTS idx_eqc_rounds ON eqc_rounds(analyte, year);

CREATE TABLE IF NOT EXISTS eqc_samples (
    sample_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    round_id     INTEGER NOT NULL,
    sample_label TEXT,
    lab_value    REAL,
    peer_mean    REAL,
    peer_sd      REAL,
    lower_limit  REAL,
    upper_limit  REAL,
    bias_pct     REAL,
    FOREIGN KEY (round_id) REFERENCES eqc_rounds(round_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_eqc_samples ON eqc_samples(round_id);

CREATE TABLE IF NOT EXISTS eqc_audit (
    audit_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    entity     TEXT,
    entity_id  INTEGER,
    action     TEXT,
    user_login TEXT,
    at         TEXT,
    detail     TEXT
);
"""


# --------------------------------------------------------------------------- #
# Migração de schema (adiciona colunas faltantes em bancos já existentes)
# --------------------------------------------------------------------------- #
def _migrate(conn):
    """Adiciona colunas novas a tabelas pré-existentes, sem perder dados."""
    needed = [
        ("analytes", "area", "TEXT DEFAULT 'Hematologia'"),
        ("analytes", "unit", "TEXT DEFAULT ''"),
        ("specs", "etp_pct", "REAL"),
        ("specs", "etp_source", "TEXT"),
    ]
    for table, col, decl in needed:
        cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
        if col not in cols:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")
    # garante área preenchida nos analitos antigos
    conn.execute("UPDATE analytes SET area = 'Hematologia' WHERE area IS NULL OR area = ''")
    conn.commit()


def _seed_incremental(conn, demo: bool):
    """
    Seed idempotente do que é novo (áreas/analitos extras, ETp e rodadas EQC de
    demonstração). Roda sempre, mesmo em banco já populado, sem sobrescrever
    dados do usuário.
    """
    cur = conn.cursor()
    # config padrão do Hórus (relatório diário por IA) — só cria se não existir
    for k, v in (("horus_email", "natallisantos1995@gmail.com"),
                 ("horus_hora", "10:00"), ("horus_enabled", "0")):
        cur.execute("INSERT OR IGNORE INTO config(key, value) VALUES (?,?)", (k, v))
    # analitos extras por área (Bioquímica, Imunologia, Coagulação...) — não toca
    # nos que já existem (INSERT OR IGNORE).
    base_ord = 1000
    for j, a in enumerate(getattr(seed, "EXTRA_ANALYTES", [])):
        cur.execute(
            "INSERT OR IGNORE INTO analytes(name, ord, decimals, area, unit) VALUES (?,?,?,?,?)",
            (a["name"], base_ord + j, a.get("decimals", 2), a.get("area", ""), a.get("unit", "")))

    # ETp configurável por ensaio (só preenche onde ainda está vazio)
    for a, (etp_pct, source) in getattr(seed, "ETP", {}).items():
        cur.execute(
            "UPDATE specs SET etp_pct = ?, etp_source = ? "
            "WHERE analyte = ? AND (etp_pct IS NULL)",
            (etp_pct, source, a))
        # cria a linha de specs se o analito não tiver nenhuma ainda
        cur.execute(
            "INSERT OR IGNORE INTO specs(analyte, etp_pct, etp_source) VALUES (?,?,?)",
            (a, etp_pct, source))

    # rodadas EQC de demonstração (só se a tabela estiver vazia)
    if demo:
        cur.execute("SELECT COUNT(*) AS c FROM eqc_rounds")
        if cur.fetchone()["c"] == 0:
            now = datetime.now().isoformat(timespec="seconds")
            for rd in getattr(seed, "EQC_ROUNDS", []):
                samples = rd.get("samples", [])
                cur.execute(
                    """INSERT OR IGNORE INTO eqc_rounds
                       (area, analyte, year, round_number, round_date, provider,
                        unit, lote, status, notes, report_ref, created_at, created_by)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (rd.get("area"), rd["analyte"], rd["year"], rd["round_number"],
                     rd.get("round_date"), rd.get("provider"), rd.get("unit"),
                     rd.get("lote"), rd.get("status"), rd.get("notes"),
                     rd.get("report_ref"), now, "seed"))
                rid = cur.lastrowid
                if not rid:
                    continue
                for s in samples:
                    bias = _sample_bias(s.get("lab_value"), s.get("peer_mean"))
                    cur.execute(
                        """INSERT INTO eqc_samples
                           (round_id, sample_label, lab_value, peer_mean, peer_sd,
                            lower_limit, upper_limit, bias_pct)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (rid, s.get("label"), s.get("lab_value"), s.get("peer_mean"),
                         s.get("peer_sd"), s.get("lower_limit"), s.get("upper_limit"), bias))
    conn.commit()


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
    _migrate(conn)

    # já populado? (seed-núcleo só na 1ª vez; o incremental roda sempre)
    cur.execute("SELECT COUNT(*) AS c FROM analytes")
    if cur.fetchone()["c"] > 0 and not force:
        _seed_incremental(conn, demo)
        conn.commit()
        return conn

    # analitos (hematologia)
    for i, a in enumerate(seed.ANALYTES):
        cur.execute("INSERT OR REPLACE INTO analytes(name, ord, decimals, area, unit) VALUES (?,?,?,?,?)",
                    (a, i, 4 if a in ("RET#", "BASO#", "NRBC#", "EO#", "MONO#") else 2,
                     "Hematologia", ""))

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
    # áreas/analitos extras, ETp e rodadas EQC de demonstração
    _seed_incremental(conn, demo)
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


def set_config(key, value):
    """Salva/atualiza uma chave de configuração."""
    conn = get_conn()
    conn.execute("INSERT OR REPLACE INTO config(key, value) VALUES (?,?)",
                 (key, "" if value is None else str(value)))
    conn.commit()


def get_analytes(area=None, with_reference=False):
    """
    Lista de analitos (nomes), ordenada. Filtros opcionais:
      area           : restringe a uma área técnica
      with_reference : só os que têm média/DP de referência (Controle Interno)
    """
    conn = get_conn()
    q = "SELECT DISTINCT a.name FROM analytes a"
    p = []
    if with_reference:
        q += " JOIN reference r ON r.analyte = a.name"
    if area:
        q += " WHERE a.area = ?"; p.append(area)
    q += " ORDER BY a.ord"
    return [r["name"] for r in conn.execute(q, p)]


def get_areas():
    """Áreas técnicas cadastradas (dos analitos)."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT area FROM analytes WHERE area IS NOT NULL AND area <> '' "
        "ORDER BY area").fetchall()
    return [r["area"] for r in rows]


def get_analyte_info(name):
    """Retorna dict do analito (area, unit, decimals) ou {}."""
    conn = get_conn()
    r = conn.execute("SELECT * FROM analytes WHERE name = ?", (name,)).fetchone()
    return dict(r) if r else {}


def add_analyte(name, area, unit="", decimals=2):
    """Cadastra um novo analito/ensaio (extensibilidade). INSERT OR IGNORE."""
    conn = get_conn()
    r = conn.execute("SELECT MAX(ord) AS m FROM analytes").fetchone()
    nxt = (r["m"] or 0) + 1
    conn.execute(
        "INSERT OR IGNORE INTO analytes(name, ord, decimals, area, unit) VALUES (?,?,?,?,?)",
        (name.strip(), nxt, int(decimals), area.strip(), (unit or "").strip()))
    conn.commit()


def set_etp(analyte, etp_pct, etp_source):
    """Define o Erro Total Permitido (ETp, em %) e a origem para um ensaio."""
    conn = get_conn()
    conn.execute("INSERT OR IGNORE INTO specs(analyte) VALUES (?)", (analyte,))
    conn.execute("UPDATE specs SET etp_pct = ?, etp_source = ? WHERE analyte = ?",
                 (etp_pct, etp_source, analyte))
    conn.commit()


def get_etp(analyte):
    """Retorna (etp_pct, etp_source) configurados, ou (None, None)."""
    conn = get_conn()
    r = conn.execute("SELECT etp_pct, etp_source FROM specs WHERE analyte = ?",
                     (analyte,)).fetchone()
    return (r["etp_pct"], r["etp_source"]) if r else (None, None)


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


# --------------------------------------------------------------------------- #
# CONTROLE EXTERNO (EQC) — rodadas, amostras, indicador e auditoria
# --------------------------------------------------------------------------- #
MAX_ROUNDS_POR_ANO = 6


def eqc_log(entity, entity_id, action, user, detail=""):
    """Registra uma alteração na trilha de auditoria do EQC."""
    conn = get_conn()
    conn.execute(
        "INSERT INTO eqc_audit(entity, entity_id, action, user_login, at, detail) "
        "VALUES (?,?,?,?,?,?)",
        (entity, entity_id, action, user,
         datetime.now().isoformat(timespec="seconds"), detail))
    conn.commit()


def eqc_round_count(analyte, year, exclude_round_id=None):
    """Quantas rodadas já existem para o analito naquele ano."""
    conn = get_conn()
    q = "SELECT COUNT(*) AS c FROM eqc_rounds WHERE analyte = ? AND year = ?"
    p = [analyte, year]
    if exclude_round_id:
        q += " AND round_id <> ?"; p.append(exclude_round_id)
    return conn.execute(q, p).fetchone()["c"]


def eqc_round_number_taken(analyte, year, round_number, exclude_round_id=None):
    """True se o número de rodada já está usado para o analito/ano."""
    conn = get_conn()
    q = ("SELECT COUNT(*) AS c FROM eqc_rounds "
         "WHERE analyte = ? AND year = ? AND round_number = ?")
    p = [analyte, year, round_number]
    if exclude_round_id:
        q += " AND round_id <> ?"; p.append(exclude_round_id)
    return conn.execute(q, p).fetchone()["c"] > 0


def eqc_add_round(data: dict, samples: list, user: str = "sistema"):
    """
    Cria uma rodada de Controle Externo com suas amostras.
    Regras: no máx. 6 rodadas/ano por analito (regra 1/3); número de rodada único.
    Retorna (round_id, None) em sucesso ou (None, mensagem_de_erro).
    """
    analyte = data.get("analyte")
    year = int(data.get("year"))
    rnum = int(data.get("round_number"))

    if rnum < 1 or rnum > MAX_ROUNDS_POR_ANO:
        return None, f"Número da rodada deve estar entre 1 e {MAX_ROUNDS_POR_ANO}."
    if eqc_round_count(analyte, year) >= MAX_ROUNDS_POR_ANO:
        return None, (f"Limite de {MAX_ROUNDS_POR_ANO} rodadas/ano já atingido "
                      f"para {analyte} em {year}.")
    if eqc_round_number_taken(analyte, year, rnum):
        return None, f"A rodada {rnum} de {analyte}/{year} já está cadastrada."

    conn = get_conn()
    now = datetime.now().isoformat(timespec="seconds")
    cur = conn.execute(
        """INSERT INTO eqc_rounds
           (area, analyte, year, round_number, round_date, provider, unit, lote,
            status, notes, report_ref, created_at, created_by, updated_at, updated_by)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (data.get("area"), analyte, year, rnum, data.get("round_date"),
         data.get("provider"), data.get("unit"), data.get("lote"),
         data.get("status"), data.get("notes"), data.get("report_ref"),
         now, user, now, user))
    rid = cur.lastrowid
    for s in samples:
        bias = _sample_bias(s.get("lab_value"), s.get("peer_mean"))
        conn.execute(
            """INSERT INTO eqc_samples
               (round_id, sample_label, lab_value, peer_mean, peer_sd,
                lower_limit, upper_limit, bias_pct)
               VALUES (?,?,?,?,?,?,?,?)""",
            (rid, s.get("sample_label") or s.get("label"), s.get("lab_value"),
             s.get("peer_mean"), s.get("peer_sd"), s.get("lower_limit"),
             s.get("upper_limit"), bias))
    conn.commit()
    eqc_log("round", rid, "INSERT", user,
            f"{analyte} {year} R{rnum} ({len(samples)} amostras)")
    return rid, None


def eqc_update_round(round_id: int, data: dict, samples: list, user: str = "sistema"):
    """Atualiza rodada + amostras (substitui as amostras). Retorna (ok, erro)."""
    conn = get_conn()
    cur = conn.execute("SELECT * FROM eqc_rounds WHERE round_id = ?", (round_id,))
    atual = cur.fetchone()
    if not atual:
        return False, "Rodada não encontrada."

    analyte = data.get("analyte", atual["analyte"])
    year = int(data.get("year", atual["year"]))
    rnum = int(data.get("round_number", atual["round_number"]))
    if rnum < 1 or rnum > MAX_ROUNDS_POR_ANO:
        return False, f"Número da rodada deve estar entre 1 e {MAX_ROUNDS_POR_ANO}."
    if eqc_round_number_taken(analyte, year, rnum, exclude_round_id=round_id):
        return False, f"A rodada {rnum} de {analyte}/{year} já está cadastrada."

    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """UPDATE eqc_rounds SET area=?, analyte=?, year=?, round_number=?,
           round_date=?, provider=?, unit=?, lote=?, status=?, notes=?,
           report_ref=?, updated_at=?, updated_by=? WHERE round_id=?""",
        (data.get("area"), analyte, year, rnum, data.get("round_date"),
         data.get("provider"), data.get("unit"), data.get("lote"),
         data.get("status"), data.get("notes"), data.get("report_ref"),
         now, user, round_id))
    conn.execute("DELETE FROM eqc_samples WHERE round_id = ?", (round_id,))
    for s in samples:
        bias = _sample_bias(s.get("lab_value"), s.get("peer_mean"))
        conn.execute(
            """INSERT INTO eqc_samples
               (round_id, sample_label, lab_value, peer_mean, peer_sd,
                lower_limit, upper_limit, bias_pct)
               VALUES (?,?,?,?,?,?,?,?)""",
            (round_id, s.get("sample_label") or s.get("label"), s.get("lab_value"),
             s.get("peer_mean"), s.get("peer_sd"), s.get("lower_limit"),
             s.get("upper_limit"), bias))
    conn.commit()
    eqc_log("round", round_id, "UPDATE", user, f"{analyte} {year} R{rnum}")
    return True, None


def eqc_delete_round(round_id: int, user: str = "sistema"):
    """Exclui uma rodada (e suas amostras, via cascade)."""
    conn = get_conn()
    conn.execute("DELETE FROM eqc_samples WHERE round_id = ?", (round_id,))
    conn.execute("DELETE FROM eqc_rounds WHERE round_id = ?", (round_id,))
    conn.commit()
    eqc_log("round", round_id, "DELETE", user, "")


def eqc_get_samples(round_id):
    conn = get_conn()
    return [dict(r) for r in conn.execute(
        "SELECT * FROM eqc_samples WHERE round_id = ? ORDER BY sample_id", (round_id,))]


def eqc_get_round(round_id):
    conn = get_conn()
    r = conn.execute("SELECT * FROM eqc_rounds WHERE round_id = ?", (round_id,)).fetchone()
    return dict(r) if r else None


def eqc_list_rounds(year=None, area=None, analyte=None, provider=None,
                    round_number=None, lote=None):
    """Lista rodadas com filtros opcionais (regras 2 e 14). Cada item inclui o
    |bias| da rodada (round_abs_bias) já calculado a partir das amostras."""
    conn = get_conn()
    q = "SELECT * FROM eqc_rounds WHERE 1=1"
    p = []
    if year:
        q += " AND year = ?"; p.append(year)
    if area:
        q += " AND area = ?"; p.append(area)
    if analyte:
        q += " AND analyte = ?"; p.append(analyte)
    if provider:
        q += " AND provider = ?"; p.append(provider)
    if round_number:
        q += " AND round_number = ?"; p.append(round_number)
    if lote:
        q += " AND lote = ?"; p.append(lote)
    q += " ORDER BY year DESC, analyte, round_number"
    rounds = [dict(r) for r in conn.execute(q, p)]
    for rd in rounds:
        biases = [s["bias_pct"] for s in eqc_get_samples(rd["round_id"])]
        abs_vals = [abs(b) for b in biases if b is not None]
        rd["round_abs_bias"] = (sum(abs_vals) / len(abs_vals)) if abs_vals else None
        rd["n_samples"] = len(biases)
    return rounds


def eqc_annual_indicator(analyte, year):
    """
    Indicador de Bias Anual Robusto de um analito/ano:
      - |bias| de cada rodada = média dos |bias| das amostras;
      - indicador = média dos |bias| das rodadas.
    Retorna dict {indicador, por_rodada:[{round_number, abs_bias}], n_rodadas}.
    """
    rounds = eqc_list_rounds(year=year, analyte=analyte)
    por_rodada = [{"round_number": rd["round_number"],
                   "abs_bias": rd["round_abs_bias"],
                   "round_id": rd["round_id"]}
                  for rd in rounds]
    abs_list = [rd["round_abs_bias"] for rd in rounds if rd["round_abs_bias"] is not None]
    indicador = (sum(abs_list) / len(abs_list)) if abs_list else None
    return {"indicador": indicador,
            "por_rodada": sorted(por_rodada, key=lambda x: x["round_number"]),
            "n_rodadas": len(abs_list)}


def eqc_years():
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT year FROM eqc_rounds ORDER BY year DESC").fetchall()
    return [r["year"] for r in rows]


def eqc_providers():
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT provider FROM eqc_rounds "
        "WHERE provider IS NOT NULL AND provider <> '' ORDER BY provider").fetchall()
    return [r["provider"] for r in rows]


def eqc_lotes():
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT lote FROM eqc_rounds "
        "WHERE lote IS NOT NULL AND lote <> '' ORDER BY lote DESC").fetchall()
    return [r["lote"] for r in rows]


def eqc_analytes_with_rounds(year=None):
    """Analitos que possuem rodadas (opcionalmente em um ano)."""
    conn = get_conn()
    q = "SELECT DISTINCT analyte FROM eqc_rounds"
    p = []
    if year:
        q += " WHERE year = ?"; p.append(year)
    q += " ORDER BY analyte"
    return [r["analyte"] for r in conn.execute(q, p)]


if __name__ == "__main__":
    # Recria o banco de demonstração
    init_db(force=True, demo=True)
    c = get_conn()
    print("Analitos:", c.execute("SELECT COUNT(*) FROM analytes").fetchone()[0])
    print("Referências:", c.execute("SELECT COUNT(*) FROM reference").fetchone()[0])
    print("Resultados:", c.execute("SELECT COUNT(*) FROM results").fetchone()[0])
    print("Usuários:", [r[0] for r in c.execute("SELECT login FROM users")])
