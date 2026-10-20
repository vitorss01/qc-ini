# -*- coding: utf-8 -*-
"""instalar_adr079.py -- ADR-079: o ATUALIZAR calcula a camada de dados em VBA (segundos em vez de minutos).

Uso:  python instalar_adr079.py <Bioquimica|Hematologia> <arquivo.xlsm> [--motor VBA|PQ] [--atualizar] [--sem-salvar]

Idempotente. So:
  1. tblConfigIntegracao: chave MOTOR_DADOS (VBA = mEtlSync; PQ = refresh das consultas, o caminho antigo). Valor ja
     existente e preservado; --motor grava o pedido;
  2. aba Sync_Log com a tblSyncLog (uma linha por ATUALIZAR: tempos por etapa, contagens, datas, erro, versao --
     sem credencial e sem dado clinico). Ja existindo, so confere/acrescenta as colunas;
  3. instala todo o VBA atual das fontes (codigo_atual.py: mEtlCQ, mEtlReceb, mEtlQA, mEtlSync e o mIntegracao
     novo); EXIGE compilacao; tabela publicada com colunas de uma consulta antiga (Hematologia de producao em
     20/10/2026: 39 colunas na tblCQ_Final) e refeita AQUI pelo Power Query -- senao o 1o ATUALIZAR do usuario pagaria
     esse refresh (~1 min);
  4. --atualizar: roda o ATUALIZAR DADOS e exige OK + a linha nova no Sync_Log;
  5. devolve a protecao de cada aba e a da estrutura como estavam.
As consultas do Power Query NAO mudam (ficam como reserva e conferencia). Qualquer falha para ANTES de salvar.
"""
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import camada_dados as cd  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

ABA_LOG = 'Sync_Log'
TB_LOG = 'tblSyncLog'
COLUNAS_LOG = ['INICIO', 'FIM', 'RESULTADO', 'MOTOR_DADOS', 'MODO_FONTE', 'LIDOS_ORIGEM', 'SEM_ID', 'NOVOS',
               'TOTAL_RECEBIDO', 'LINHAS_CQ', 'GRAVADAS', 'DATA_MIN', 'DATA_MAX', 'QA_ERROS', 'QA_ALERTAS',
               'T_ORIGEM', 'T_RECEBIMENTO', 'T_CQ', 'T_VALIDAR', 'T_GRAVAR', 'T_QA', 'T_MOTOR', 'TOTAL_S', 'TEMPOS',
               'ERRO', 'ESQUEMA', 'USUARIO', 'VERSAO']
FORMATOS = {'INICIO': 'dd/mm/yyyy hh:mm:ss', 'FIM': 'dd/mm/yyyy hh:mm:ss', 'DATA_MIN': 'dd/mm/yyyy hh:mm',
            'DATA_MAX': 'dd/mm/yyyy hh:mm'}
TITULO = ('SYNC_LOG (ADR-079): uma linha por ATUALIZAR DADOS -- resultado, contagens, datas mínima/máxima do recebido '
          'e tempo de cada etapa (s). Sem senha e sem resultado de paciente/controle. Guardadas as 2000 mais recentes.')
MOTORES = ('VBA', 'PQ')


def log(*a):
    print(*a, flush=True)


def gravar_motor(wb, motor):
    lo = pqlib.tabela(wb, 'tblConfigIntegracao')
    ks = [str(r[0] or '').strip().upper() for r in lo.DataBodyRange.Value]
    if 'MOTOR_DADOS' not in ks:
        raise SystemExit('MOTOR_DADOS nao existe em tblConfigIntegracao depois do passo 1')
    cel = lo.DataBodyRange.Cells(ks.index('MOTOR_DADOS') + 1, 2)
    if motor:
        cel.Value = motor
    return str(cel.Value or '').strip().upper()


def montar_log(wb):
    ws = ii.sheet(wb, ABA_LOG)
    if ws is None:
        ultima = wb.Worksheets(wb.Worksheets.Count)
        ws = wb.Worksheets.Add(None, ultima)
        ws.Name = ABA_LOG
        log(f'   aba {ABA_LOG}: criada')
    lo = pqlib.tabela(wb, TB_LOG)
    if lo is None:
        ws.Range('A1').Value = TITULO
        ws.Range('A1').Font.Bold = True
        lo = pqlib.criar_tabela_dados(ws, TB_LOG, COLUNAS_LOG, [], 'A3')
        lo.DataBodyRange.Rows(1).Delete()                     # tabela vazia: a 1a execucao acrescenta a linha
        cd.formatar_cabecalho_tabela(lo)
        log(f'   {TB_LOG}: criada ({len(COLUNAS_LOG)} colunas)')
    else:
        tem = [c.Name for c in lo.ListColumns]
        faltam = [c for c in COLUNAS_LOG if c not in tem]
        for c in faltam:
            lo.ListColumns.Add().Name = c
        log(f'   {TB_LOG}: existia; {len(faltam)} coluna(s) nova(s)')
    lin = lo.HeaderRowRange.Row + 1
    for c, f in FORMATOS.items():
        k = lo.ListColumns(c).Range.Column
        ws.Range(ws.Cells(lin, k), ws.Cells(lin + 4000, k)).NumberFormat = f
    ws.Columns('A:AA').ColumnWidth = 12
    ws.Columns(lo.ListColumns('TEMPOS').Range.Column).ColumnWidth = 60
    ws.Columns(lo.ListColumns('ERRO').Range.Column).ColumnWidth = 60
    ws.Cells.Locked = True
    return ws


def main(produto, caminho, motor=None, atualizar=False, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)

        log('1. tblConfigIntegracao: MOTOR_DADOS')
        cd.montar_cfg(wb, produto, ii.sheet(wb, 'Cfg_Integracao'))
        final = gravar_motor(wb, motor)
        if final not in MOTORES:
            raise SystemExit(f'MOTOR_DADOS invalido: {final!r} (aceitos: {MOTORES})')
        log(f'   MOTOR_DADOS = {final}')

        log(f'2. aba {ABA_LOG}')
        ws_log = montar_log(wb)

        log('3. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.EnableEvents = True
        esq = str(ex.run("'" + wb.Name + "'!mEtlSync.ConferirEsquemas", teto=1500) or '')
        ex.xl.EnableEvents = False
        if esq.startswith('ERRO|'):
            raise SystemExit('conferencia das colunas falhou: ' + esq[5:])
        log('   colunas das tabelas publicadas: ' + (('REFEITAS pelo Power Query -- ' + esq) if esq else 'iguais ao calculo'))

        if atualizar:
            log('4. ATUALIZAR DADOS')
            lo = pqlib.tabela(wb, TB_LOG)
            n0 = lo.ListRows.Count
            ex.xl.EnableEvents = True
            ex.xl.Calculation = -4105
            t0 = time.time()
            r = str(ex.run("'" + wb.Name + "'!mIntegracao.AtualizarDadosAutomatico", teto=1500))
            ex.xl.EnableEvents = False
            if not r.startswith('OK|') or 'concluída' not in r:
                raise SystemExit('ATUALIZAR DADOS falhou: ' + r[:800])
            if lo.ListRows.Count != n0 + 1:
                raise SystemExit(f'Sync_Log: esperada 1 linha nova, vieram {lo.ListRows.Count - n0}')
            log(f'   OK ({time.time() - t0:.1f}s)\n    ' + r[3:].replace('\n', '\n    '))

        foto[ABA_LOG] = foto.get(ABA_LOG) or dict(foto_protecao(wb)[ABA_LOG], conteudo=True, cenarios=True,
                                                    AllowFiltering=True, AllowSorting=False)
        devolver_protecao(wb, foto)
        if not ws_log.ProtectContents:
            raise SystemExit(f'{ABA_LOG} ficou sem protecao')
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    m = a[a.index('--motor') + 1].upper() if '--motor' in a else None
    main(a[0], a[1], m, atualizar='--atualizar' in a, salvar='--sem-salvar' not in a)
