# -*- coding: utf-8 -*-
"""perfil_calculo.py -- onde o recalculo gasta tempo, aba por aba.

Uso: python perfil_calculo.py <arquivo.xlsm>

Com o calculo em MANUAL, chama Worksheet.Calculate em cada aba e mede. Depois
mede o recalculo completo e o custo de uma escrita tipica do motor em
Eng_Saida com o calculo em AUTOMATICO (o que o usuario paga a cada clique).
"""
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402


def main(caminho):
    ex = xlh.Excel()
    t0 = time.time()
    wb = ex.abrir(os.path.abspath(caminho))
    print(f'abrir: {time.time() - t0:.1f}s', flush=True)
    xl = ex.xl
    try:
        xl.Calculation = -4135          # manual
        tempos = []
        for ws in wb.Worksheets:
            if ws.ProtectContents:
                ws.Unprotect('qcini2025')
            t0 = time.time()
            ws.Calculate()
            ex.esperar()
            tempos.append((time.time() - t0, ws.Name))
        for t, n in sorted(tempos, reverse=True)[:12]:
            print(f'   {n:28s} {t:7.2f}s', flush=True)
        t0 = time.time(); xl.CalculateFull(); ex.esperar()
        print(f'CalculateFull: {time.time() - t0:.1f}s', flush=True)
        if len(sys.argv) > 2:
            return
        xl.Calculation = -4105          # automatico
        eng = wb.Sheets('Eng_Saida')
        v = eng.Range('Y3').Value
        t0 = time.time(); eng.Range('Y3').Value = v; ex.esperar()
        print(f'1 escrita em Eng_Saida (auto): {time.time() - t0:.2f}s', flush=True)
        t0 = time.time(); ex.run('mEstatistica.RecalcularAnalitoAtual', teto=1800)
        print(f'RecalcularAnalitoAtual (auto): {time.time() - t0:.1f}s', flush=True)
        xl.Calculation = -4135
        t0 = time.time(); ex.run('mEstatistica.RecalcularAnalitoAtual', teto=1800)
        print(f'RecalcularAnalitoAtual (manual, so VBA): {time.time() - t0:.1f}s', flush=True)
        xl.Calculation = -4105
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(sys.argv[1])
