# -*- coding: utf-8 -*-
"""upgrade_v2.py -- instala a camada v2 (ADR-050) numa COPIA de teste ja com Etapa 1 e salva.

Uso: python upgrade_v2.py <origem.xlsm> <destino.xlsm>
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402
import v2   # noqa: E402

PROV = 'IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A14,Analitos!$A$4:$A$43,0)),"CAP")'
MODS = ('mApp', 'mUI', 'mBanco', 'mLotes', 'mEstatistica', 'mOperacao', 'mImportar', 'mCEQ', 'mDados')


def main(src, dst):
    shutil.copyfile(src, dst)
    ex = xlh.Excel()
    wb = ex.abrir(dst)
    ok = False
    try:
        est = wb.Sheets('Estatística')
        prot = {}
        for s in wb.Worksheets:
            prot[s.Name] = bool(s.ProtectContents)
            if prot[s.Name]:
                s.Unprotect('qcini2025')
        for col in ('R', 'S', 'T', 'AC', 'AD'):
            f = est.Range(f'{col}14').Formula
            if 'eqProvedor' in f:
                est.Range(f'{col}14:{col}93').Formula = f.replace('eqProvedor', PROV)
        vbp = wb.VBProject
        for m in MODS:
            v2.substituir_codigo(vbp, m, v2.codigo(m + '.bas'))
        v2.aplicar_planilhas(ex, wb, 2)
        ex.xl.CalculateFull(); ex.esperar()
        ex.run('mEstatistica.AtualizarEstatistica', teto=600)
        for s in wb.Worksheets:
            if prot[s.Name]:
                s.Protect('qcini2025', False, True, True, True)
        wb.Save()
        ok = True
        print('salvo', dst)
    finally:
        ex.fechar()
        if not ok:
            sys.exit(1)


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]))
