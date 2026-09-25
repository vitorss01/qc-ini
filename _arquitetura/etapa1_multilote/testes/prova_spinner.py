# -*- coding: utf-8 -*-
"""prova_spinner.py -- o que o Painel mostra logo depois do spinner de analito.

Uso: python prova_spinner.py <arquivo.xlsm>   (copia; nada e salvo)

Faz exatamente o que o clique no spinner faz: muda Painel!B3 e roda a macro
associada ao controle (OnAction). Le o status Westgard (M7/M8) e os
veredictos por corrida. Depois roda o motor completo e le de novo.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402


def estado(wb):
    pai, calc = wb.Sheets('Painel'), wb.Sheets('Calc')
    st = [r[0] for r in calc.Range('P3:P182').Value if r[0] not in (None, '')]
    nb = sum(1 for r in calc.Range('B3:B182').Value if r[0] not in (None, ''))
    return (pai.Range('C3').Value, pai.Range('M7').Value, pai.Range('M8').Value,
            f'corridas={nb} veredictos={len(st)} rejeitados={st.count("REJEITADO")}')


def main(caminho):
    ex = xlh.Excel()
    wb = ex.abrir(os.path.abspath(caminho))
    try:
        pai = wb.Sheets('Painel')
        acao = pai.Shapes('Spinner 1').OnAction
        print('macro do spinner:', acao)
        pai.Range('B3').Value = 1
        ex.esperar()
        ex.run('mEstatistica.RecalcularAnalitoAtual', teto=900); ex.esperar()
        print('Lactato, motor em dia      :', estado(wb))
        for idx in (13, 3):
            pai.Range('B3').Value = idx          # o clique
            ex.esperar()
            ex.run(acao.split('!')[-1], teto=900); ex.esperar()
            print('depois do spinner          :', estado(wb))
            ex.run('mEstatistica.RecalcularAnalitoAtual', teto=900); ex.esperar()
            print('  e depois do motor        :', estado(wb))
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(sys.argv[1])
