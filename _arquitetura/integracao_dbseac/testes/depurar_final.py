# -*- coding: utf-8 -*-
"""Bancada de depuracao do DB_CQ_FINAL.

  preparar <produto> <copia.xlsm>   cria entradas + RECEBIMENTO + ORGANIZADO e SALVA
  rodar <copia.xlsm> N [etapa]      grava DB_CQ_FINAL limitado as N primeiras linhas recebidas
                                    (N=0 -> todas) e mede; 'etapa' = nome de passo do let para
                                    devolver no lugar do resultado (localiza o passo que falha)
"""
import collections
import os
import re
import sys
import threading
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import config_produtos as cp  # noqa: E402


def aba(wb, nome):
    for ws in wb.Worksheets:
        if ws.Name == nome:
            return ws
    ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nome
    return ws


def preparar(produto, copia):
    xl = xlh.Excel()
    try:
        wb = xl.abrir(copia)
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        wb.Queries.FastCombine = True
        cfg = aba(wb, 'Cfg_Integracao')
        pqlib.criar_tabela_dados(cfg, 'tblConfigIntegracao', ['CHAVE', 'VALOR', 'DESCRICAO'], cp.linhas_cfg(produto), 'A4')
        pqlib.criar_tabela_dados(cfg, 'tblDeParaAnalitos', ['MATRIZ', 'ANALITO_ORIGEM', 'ANALITO_QCINI', 'BLOCO', 'ORDEM', 'OBS'],
                                 cp.linhas_depara(produto), 'E4')
        pqlib.criar_tabela_dados(aba(wb, 'DB_MANUAL'), 'tblResultados_Manuais',
                                 ['ID_REGISTRO', 'DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO',
                                  'RESULTADO', 'UNIDADE', 'MOTIVO', 'USUARIO', 'REGISTRADO_EM'], [], 'A5')
        pqlib.criar_tabela_dados(aba(wb, 'INATIVACAO_NAO_CONFORMES'), 'tblInativacao_NaoConformes',
                                 ['ID_REGISTRO', 'REGISTRAR - LJ', 'DATA_INATIVACAO', 'USUARIO'], [], 'A5')
        pqlib.criar_tabela_dados(aba(wb, 'COMENTARIOS_TECNICOS'), 'tblComentariosTecnicos',
                                 ['ID_REGISTRO', 'COMENTARIO_TECNICO', 'DATA', 'USUARIO'], [], 'A5')
        for q in ('CFG_INTEGRACAO', 'SEAC_ORIGEM', 'DB_RECEBIMENTO', 'DB_ORGANIZADO'):
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
        for q, tb in (('DB_RECEBIMENTO', 'tblDB_Recebimento'), ('DB_ORGANIZADO', 'tblDB_Organizado')):
            lo, t = pqlib.carregar_em_tabela(aba(wb, q), q, tb, 'A4')
            print(q, round(t, 1), 's', lo.ListRows.Count, 'linhas', flush=True)
        # DB_CQ_FINAL nasce pequeno (50 linhas) so para a tabela existir
        pqlib.gravar_query(wb, 'DB_CQ_FINAL', limitar(pqlib.m_de('DB_CQ_FINAL'), 50))
        lo, t = pqlib.carregar_em_tabela(aba(wb, 'DB_CQ_FINAL'), 'DB_CQ_FINAL', 'tblCQ_Final', 'A4')
        print('DB_CQ_FINAL(50)', round(t, 1), 's', lo.ListRows.Count, 'linhas', flush=True)
        wb.Save()
    finally:
        xl.fechar()


def limitar(m, n):
    alvo = 'R0 = Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content],'
    assert alvo in m
    if n:
        m = m.replace(alvo, 'R0 = Table.FirstN(Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content], %d),' % n)
    return m


def devolver_etapa(m, etapa):
    # troca o 'in Ordenado' final pelo passo pedido, convertido em tabela de contagem
    i = m.rfind('\nin\n')
    return m[:i] + '\nin\n    Table.FromRecords({[etapa = "%s", linhas = try Table.RowCount(%s) otherwise try List.Count(%s) otherwise -1]})\n' % (etapa, etapa, etapa)


def m_final():
    # DB_CQ_FINAL_M=<arquivo.m> mede uma versao alternativa (antes x depois)
    alt = os.environ.get('DB_CQ_FINAL_M')
    if alt:
        with open(alt, encoding='utf-8') as f:
            return f.read()
    return pqlib.m_de('DB_CQ_FINAL')


def rodar(copia, n, etapa=None, teto=900):
    xl = xlh.Excel()
    print('EXCEL_PID', xl.pid, flush=True)
    feito = threading.Event()

    def cao():
        if not feito.wait(teto):
            print(f'  TETO de {teto}s estourado -- Excel encerrado', flush=True)
            xl.matar()
    threading.Thread(target=cao, daemon=True).start()
    try:
        wb = xl.abrir(copia)
        m = limitar(m_final(), n)
        if etapa:
            m = devolver_etapa(m, etapa)
        for q in ('CFG_INTEGRACAO', 'SEAC_ORIGEM'):
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
        pqlib.gravar_query(wb, 'DB_CQ_FINAL', m)
        lo = pqlib.tabela(wb, 'tblCQ_Final')
        t0 = time.time()
        try:
            pqlib.atualizar(lo)
            print(f'N={n} etapa={etapa}: {time.time()-t0:.1f}s linhas={lo.ListRows.Count}', flush=True)
            if etapa:
                print('  ', pqlib.ler_tabela(lo), flush=True)
            else:
                fin = pqlib.ler_tabela(lo)
                print('  status:', dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin)), flush=True)
                ch = collections.Counter((r['EQUIPAMENTO'], r['LOTE'], r['NIVEL'], r['ANALITO'], r['RUN'])
                                         for r in fin if r['PARTICIPA_ESTATISTICA'] == 'SIM')
                print('  >1 participante por (eq,lote,nivel,analito,RUN):', sum(1 for v in ch.values() if v > 1), flush=True)
        except Exception as e:
            print(f'N={n} etapa={etapa}: FALHOU em {time.time()-t0:.1f}s -> {e}', flush=True)
    finally:
        feito.set()
        xl.fechar()


if __name__ == '__main__':
    if sys.argv[1] == 'preparar':
        preparar(sys.argv[2], os.path.abspath(sys.argv[3]))
    else:
        rodar(os.path.abspath(sys.argv[2]), int(sys.argv[3]), sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] != '-' else None,
              int(sys.argv[5]) if len(sys.argv) > 5 else 900)
