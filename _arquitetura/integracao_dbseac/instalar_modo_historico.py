# -*- coding: utf-8 -*-
"""instalar_modo_historico.py -- ADR-058: MODO_FONTE (SEAC | HISTORICO) num QC_INI ja integrado (ADR-057).

Uso:  python instalar_modo_historico.py <Bioquimica|Hematologia> <arquivo.xlsm> [--modo SEAC|HISTORICO] [--sem-salvar]

Incremental -- nao refaz a instalacao do ADR-057. So:
  1. acrescenta a chave MODO_FONTE em tblConfigIntegracao (valor existente e preservado;
     --modo grava o valor pedido);
  2. regrava as consultas SEAC_ORIGEM, QA_INTEGRACAO e DB_CQ_FINAL a partir de pq/;
  3. troca o modulo mIntegracao pelo de src/<produto>/ e compila;
  4. CONFERE o que fez (formula de cada consulta e texto do modulo identicos a fonte) --
     o pipeline reporta resultado, nao intencao (QUALITY_GATE, nota de 2.1);
  5. se o modo final for HISTORICO, roda o ATUALIZAR DADOS e exige "concluida" + A07;
  6. devolve a protecao das abas (e da estrutura, se estava ligada) e salva.
Qualquer falha para ANTES de salvar.
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

# DB_CQ_FINAL entra pela regra de data futura do manual (QUALITY_GATE 4.4), da mesma rodada
CONSULTAS = ['SEAC_ORIGEM', 'DB_RECEBIMENTO', 'QA_INTEGRACAO', 'DB_CQ_FINAL']   # DB_RECEBIMENTO: ADR-066 (D01/D02/D04)
MODOS = ('SEAC', 'HISTORICO')


def log(*a):
    print(*a, flush=True)


def normalizar(t):
    return '\n'.join(l.rstrip() for l in str(t).replace('\r\n', '\n').replace('\r', '\n').strip().split('\n'))


def gravar_modo(wb, modo):
    lo = pqlib.tabela(wb, 'tblConfigIntegracao')
    vals = lo.DataBodyRange.Value
    lin = [i for i, r in enumerate(vals, 1) if str(r[0]).strip().upper() == 'MODO_FONTE']
    if not lin:
        raise SystemExit('MODO_FONTE nao existe em tblConfigIntegracao depois do passo 1')
    if modo:
        lo.DataBodyRange.Cells(lin[0], 2).Value = modo
    return str(lo.DataBodyRange.Cells(lin[0], 2).Value).strip().upper()


def main(produto, caminho, modo=None, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:          # outro Excel segura o arquivo: o Save 'funcionaria' sem gravar nada
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        ii.desproteger_tudo(wb)

        log('1. tblConfigIntegracao: MODO_FONTE')
        cd.montar_cfg(wb, produto, ii.sheet(wb, 'Cfg_Integracao'))
        final = gravar_modo(wb, modo)
        if final not in MODOS:
            raise SystemExit(f'MODO_FONTE invalido: {final!r} (aceitos: {MODOS})')
        log(f'   MODO_FONTE = {final}')
        if not produto.startswith('Bio'):
            # ADR-066 (D12): sem seletor, a serie da Hematologia e a do EQUIPAMENTO_PADRAO do CFG. O nome
            # selEquipamento passa a ler o CFG (era ""): a guarda de Calc!B (engEquip = selEquipamento), o bloco
            # de repeticoes da Estatistica e o motor (mEstatistica.EquipFiltro) falam do mesmo equipamento
            # a celula VALOR da linha EQUIPAMENTO_PADRAO (sem formula: por esta automacao o RefersTo e lido com o
            # separador local e "INDEX(...,MATCH(...))" era recusado). Names.Add sobre o nome existente: apagar
            # quebraria as formulas que o usam.
            lo = ii.sheet(wb, 'Cfg_Integracao').ListObjects('tblConfigIntegracao')
            ks = [str(r[0] or '').strip().upper() for r in lo.ListColumns('CHAVE').DataBodyRange.Value]
            cel = lo.ListColumns('VALOR').DataBodyRange.Cells(ks.index('EQUIPAMENTO_PADRAO') + 1, 1)
            wb.Names.Add('selEquipamento', "='Cfg_Integracao'!" + cel.Address)
            if str(wb.Names('selEquipamento').RefersToRange.Value or '').strip() == '':
                raise SystemExit('selEquipamento da Hematologia ficou vazio')
            log(f"   selEquipamento = EQUIPAMENTO_PADRAO do CFG ({wb.Names('selEquipamento').RefersTo})")

        log('2. consultas')
        wb.Queries.FastCombine = True
        for q in CONSULTAS:
            pqlib.gravar_query(wb, q, pqlib.m_de(q))
        for q in CONSULTAS:
            if normalizar(pqlib.query_existe(wb, q).Formula) != normalizar(pqlib.m_de(q)):
                raise SystemExit(f'consulta {q}: formula no arquivo difere da fonte depois de gravar')
        log(f'   {", ".join(CONSULTAS)}: identicas a pq/')

        log('3. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        if final == 'HISTORICO':
            log('4. ATUALIZAR DADOS em modo historico')
            ex.xl.EnableEvents = True
            ex.xl.Calculation = -4105
            t0 = time.time()
            r = str(ex.run("'" + wb.Name + "'!mIntegracao.AtualizarDadosAutomatico", teto=1500))
            ex.xl.EnableEvents = False
            if not r.startswith('OK|') or 'concluída' not in r or 'MODO HISTÓRICO' not in r:
                raise SystemExit('ATUALIZAR DADOS em modo historico falhou: ' + r[:800])
            a07 = [x for x in pqlib.ler_tabela(pqlib.tabela(wb, 'tblQA_Integracao')) if x['CODIGO'] == 'A07']
            if len(a07) != 1:
                raise SystemExit(f'QA A07 esperado 1 vez, veio {len(a07)}')
            log(f'   OK ({time.time() - t0:.1f}s)\n    ' + r[3:].replace('\n', '\n    '))

        for ws in wb.Worksheets:
            try:
                ws.Protect(ii.SENHA, False, True, False, True)
            except Exception:
                pass
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    m = a[a.index('--modo') + 1].upper() if '--modo' in a else None
    main(a[0], a[1], m, salvar='--sem-salvar' not in a)
