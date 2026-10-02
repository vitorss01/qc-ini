# -*- coding: utf-8 -*-
"""Chama, numa copia instalada, uma rotina de cada modulo e cada passo da atualizacao,
para localizar falha de compilacao/execucao do VBA. Uso: python depurar_vba.py <copia.xlsm>"""
import os, sys, time
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..')); sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import xlh
xl = xlh.Excel(); print('EXCEL_PID', xl.pid, flush=True)
try:
    wb = xl.abrir(os.path.abspath(sys.argv[1]))
    xl.xl.EnableEvents = True
    def chama(m, *a, teto=900):
        t0 = time.time()
        try:
            r = xl.run("'" + wb.Name + "'!" + m, *a, teto=teto)
            print(f'OK   {m} ({time.time()-t0:.1f}s) -> {str(r)[:200]!r}', flush=True)
        except Exception as e:
            print(f'FAIL {m} ({time.time()-t0:.1f}s) -> {e}', flush=True)
    for m, a in [('mDados.LoteAtivoCore', ()), ('mDados.UltimaDataCQ', ()), ('mIntegracao.Cfg', ('PREFIXO_ID',)),
                 ('mEstatistica.EquipFiltro', ()), ('mLotes.LotePainel', ()), ('mApp.Ocupado', ()),
                 ('mIntegracao.ContarQA', ('ERRO',)), ('mIntegracao.NormalizarId', ('12345',)), ('mIntegracao.NormalizarId', ('man_7',)),
                 ('mIntegracao.PrepararEntradas', ()), ('mIntegracao.Refrescar', ('tblQA_Integracao',)),
                 ('mDados.AtualizarListasAno', ()), ('mLotes.AtualizarListaLiberacao', ()),
                 ('mEstatistica.InvalidarCache', ()), ('mEstatistica.AtualizarEstatistica', ()),
                 ('mIntegracao.AtualizarDadosCore', ())]:
        chama(m, *a)
finally:
    xl.fechar()
