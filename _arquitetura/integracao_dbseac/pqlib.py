# -*- coding: utf-8 -*-
"""pqlib.py -- operacoes de Power Query e tabelas via COM, reutilizadas pelo
instalador e pelos testes da integracao DB_SEAC -> DB_CQ_FINAL.

Regras aprendidas neste projeto e aplicadas aqui:
- Consulta carregada em tabela: ListObjects.Add com a conexao Mashup OLE DB e
  Refresh(BackgroundQuery=False). Em segundo plano o script seguiria antes de a
  tabela existir.
- Workbook.Queries.FastCombine = True: as consultas combinam a pasta atual
  (Excel.CurrentWorkbook) com o arquivo DB_SEAC (File.Contents); sem isso o
  Formula.Firewall recusa a combinacao.
- Nome da tabela definido depois da criacao (DisplayName), porque a camada de
  recebimento le a si mesma pelo nome.
"""
import os
import time

XL_SRC_EXTERNAL = 0
XL_CMD_SQL = 2
XL_INSERT_DELETE = 1    # xlInsertDeleteCells

AQUI = os.path.dirname(os.path.abspath(__file__))
PQ_DIR = os.path.join(AQUI, 'pq')


def m_de(nome):
    with open(os.path.join(PQ_DIR, nome + '.m'), encoding='utf-8') as f:
        return f.read()


def query_existe(wb, nome):
    # A colecao Queries (Microsoft.Mashup) ja respondeu "O indice esta fora dos limites" ao ser
    # percorrida logo depois de abrir a pasta (04/10/2026, Bioquimica; a mesma pasta passou antes e
    # depois). Transitorio: percorre por indice e tenta de novo.
    for tentativa in range(5):
        try:
            qs = wb.Queries
            for i in range(1, qs.Count + 1):
                q = qs.Item(i)
                if q.Name == nome:
                    return q
            return None
        except Exception:
            if tentativa == 4:
                raise
            time.sleep(2)


def gravar_query(wb, nome, formula, descricao=''):
    q = query_existe(wb, nome)
    if q is None:
        return wb.Queries.Add(nome, formula, descricao)
    q.Formula = formula
    if descricao:
        q.Description = descricao
    return q


def tabela(wb, nome):
    # Logo depois de um refresh a enumeracao de abas/tabelas ja devolveu "indice invalido" (DISP_E_BADINDEX,
    # 04/10/2026, QA-ETL-001): transitorio -- tenta de novo antes de desistir
    for tentativa in range(5):
        try:
            for ws in wb.Worksheets:
                for lo in ws.ListObjects:
                    if lo.Name == nome:
                        return lo
            return None
        except Exception:
            if tentativa == 4:
                raise
            time.sleep(2)


def carregar_em_tabela(ws, consulta, nome_tabela, destino='A4'):
    """Carrega a consulta numa tabela nova a partir de 'destino'. Devolve o ListObject."""
    src = ('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
           'Location=%s;Extended Properties=""' % consulta)
    lo = ws.ListObjects.Add(XL_SRC_EXTERNAL, src, None, 1, ws.Range(destino))
    qt = lo.QueryTable
    qt.CommandType = XL_CMD_SQL
    qt.CommandText = 'SELECT * FROM [%s]' % consulta
    qt.RowNumbers = False
    qt.FillAdjacentFormulas = False
    qt.PreserveFormatting = True
    qt.RefreshOnFileOpen = False
    qt.BackgroundQuery = False
    qt.RefreshStyle = XL_INSERT_DELETE
    qt.SavePassword = False
    qt.SaveData = True
    qt.AdjustColumnWidth = False
    qt.PreserveColumnInfo = True
    t0 = time.time()
    qt.Refresh(False)
    # nome DEPOIS da 1a carga: a consulta de recebimento le a propria tabela pelo nome,
    # e antes da carga a tabela so tem o texto provisorio "Obtendo dados..."
    lo.DisplayName = nome_tabela
    return lo, time.time() - t0


def atualizar(lo):
    t0 = time.time()
    lo.QueryTable.BackgroundQuery = False
    lo.QueryTable.Refresh(False)
    return time.time() - t0


def criar_tabela_dados(ws, nome_tabela, cabecalho, linhas=None, destino='A4'):
    """Tabela comum (entrada do usuario / configuracao). linhas: lista de listas."""
    linhas = linhas or []
    c0 = ws.Range(destino)
    r0, k0 = c0.Row, c0.Column
    n = len(cabecalho)
    for j, h in enumerate(cabecalho):
        ws.Cells(r0, k0 + j).Value = h
    for i, lin in enumerate(linhas):
        for j, v in enumerate(lin):
            if v is not None:
                ws.Cells(r0 + 1 + i, k0 + j).Value = v
    ultima = r0 + max(1, len(linhas))
    rng = ws.Range(ws.Cells(r0, k0), ws.Cells(ultima, k0 + n - 1))
    lo = ws.ListObjects.Add(1, rng, None, 1)     # xlSrcRange, xlYes
    lo.Name = nome_tabela
    return lo


def ler_tabela(lo):
    """[(dict por linha)] -- leitura em bloco."""
    if lo.DataBodyRange is None:
        return []
    cab = [c.Name for c in lo.ListColumns]
    vals = lo.DataBodyRange.Value
    if vals is None:
        return []
    if not isinstance(vals[0], tuple):
        vals = (vals,)
    return [dict(zip(cab, v)) for v in vals]
