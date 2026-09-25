# -*- coding: utf-8 -*-
"""prova_volatil.py -- o custo de UMA escrita de celula, com e sem OFFSET volatil.

Uso: python prova_volatil.py <arquivo.xlsm>   (copia; nada e salvo)

Escreve um valor numa celula sem dependente (Eng_Saida!Y3) com o calculo em
automatico e mede. Depois troca SO os nomes lstAnosCIQ/lstAnosCEQ de OFFSET
(volatil) para INDEX (nao volatil) e mede de novo.
"""
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402


def escrita(ex, ws, n=3):
    ts = []
    for _ in range(n):
        v = ws.Range('Z200').Value
        t0 = time.time()
        ws.Range('Z200').Value = 1 if v != 1 else 2
        ex.esperar()
        ts.append(time.time() - t0)
    return ts


def main(caminho):
    ex = xlh.Excel()
    wb = ex.abrir(os.path.abspath(caminho))
    try:
        ws = wb.Sheets('Eng_Saida')
        if ws.ProtectContents:
            ws.Unprotect('qcini2025')
        print('lstAnosCEQ =', wb.Names('lstAnosCEQ').RefersTo)
        print('escrita com OFFSET  :', [f'{t:.2f}s' for t in escrita(ex, ws)], flush=True)
        for n, col, fim in (('lstAnosCIQ', 'Z', 50), ('lstAnosCEQ', 'AA', 50)):
            wb.Names(n).Delete()
            wb.Names.Add(n, f'=Configuração!${col}$2:ÍNDICE(Configuração!${col}$2:${col}${fim};'
                            f'MÁXIMO(1;CONT.VALORES(Configuração!${col}$2:${col}${fim})))')
        ex.xl.CalculateFull(); ex.esperar()
        print('lstAnosCEQ =', wb.Names('lstAnosCEQ').RefersTo, '| linhas:', wb.Names('lstAnosCEQ').RefersToRange.Rows.Count,
              '| Estatistica!N4 =', wb.Sheets('Estatística').Range('N4').Value)
        print('escrita com INDEX   :', [f'{t:.2f}s' for t in escrita(ex, ws)], flush=True)
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(sys.argv[1])
