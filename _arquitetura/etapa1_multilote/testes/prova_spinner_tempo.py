# -*- coding: utf-8 -*-
"""prova_spinner_tempo.py -- quanto custa UM clique no spinner de analito, e por que.

Uso: python prova_spinner_tempo.py <arquivo.xlsm>   (copia; nada e salvo)

1. mede o clique como ele e (B3 muda + macro do spinner), 3 vezes;
2. troca, SO nas colunas R,S,T,AC,AD da Estatistica, o provedor global
   (eqProvedor = provedor do analito EM TELA) pelo provedor DA PROPRIA LINHA,
   que e o que a coluna G ja faz;
3. mede o clique de novo e confere que G/R/S/T da Estatistica seguem calculando.
"""
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402

PROV_LINHA = 'IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A14,Analitos!$A$4:$A$43,0)),"CAP")'


def clique(ex, pai, idx):
    t0 = time.time()
    pai.Range('B3').Value = idx
    ex.esperar()
    t1 = time.time()
    ex.run(pai.Shapes('Spinner 1').OnAction.split('!')[-1], teto=600)
    ex.esperar()
    return t1 - t0, time.time() - t1


def main(caminho):
    ex = xlh.Excel()
    wb = ex.abrir(os.path.abspath(caminho))
    try:
        pai, est = wb.Sheets('Painel'), wb.Sheets('Estatística')
        for s in (pai, est):
            if s.ProtectContents:
                s.Unprotect('qcini2025')
        print('G14:', est.Range('G14').Formula[:160])
        for i, idx in enumerate((13, 3, 1)):
            f, m = clique(ex, pai, idx)
            print(f'ANTES  clique {i + 1}: formulas {f:5.1f}s + motor {m:4.1f}s = {f + m:5.1f}s  ({pai.Range("C3").Value})')
        for col in ('R', 'S', 'T', 'AC', 'AD'):
            f = est.Range(f'{col}14').Formula
            assert 'eqProvedor' in f, (col, f)
            est.Range(f'{col}14:{col}93').Formula = f.replace('eqProvedor', PROV_LINHA)
        ex.xl.CalculateFull(); ex.esperar()
        for i, idx in enumerate((13, 3, 1)):
            f, m = clique(ex, pai, idx)
            print(f'DEPOIS clique {i + 1}: formulas {f:5.1f}s + motor {m:4.1f}s = {f + m:5.1f}s  ({pai.Range("C3").Value})')
        print('Estatistica Lactato N1 G/R/S/T:', est.Range('G14:T14').Value[0][0], est.Range('R14:T14').Value[0])
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(sys.argv[1])
