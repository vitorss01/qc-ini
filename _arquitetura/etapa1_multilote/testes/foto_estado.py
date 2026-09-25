# -*- coding: utf-8 -*-
"""foto_estado.py -- fotografia do que o usuario ve, para comparar antes/depois.

Uso: python foto_estado.py <arquivo.xlsm> <saida.json> [analito ...]

Para cada analito: recalcula o motor (como o spinner faria) e guarda o Calc
(corridas, datas, valores, z, veredictos, linhas de limite), o Painel
(indicadores e status) e a Estatistica do analito. Nada e salvo no arquivo.
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402


def val(v):
    if isinstance(v, float):
        return round(v, 9)
    if hasattr(v, 'isoformat'):
        return v.isoformat()[:10]
    if isinstance(v, int) and v < -2146820000:
        return '#ERR'
    return v


def matriz(rng):
    v = rng.Value
    return [[val(x) for x in linha] for linha in v]


def main(caminho, saida, analitos):
    ex = xlh.Excel()
    wb = ex.abrir(os.path.abspath(caminho))
    foto = {}
    try:
        pai, calc, est = wb.Sheets('Painel'), wb.Sheets('Calc'), wb.Sheets('Estatística')
        nomes = [wb.Sheets('Analitos').Cells(r, 1).Value for r in range(4, 44)]
        for an in analitos:
            idx = nomes.index(an) + 1
            # calculo MANUAL durante o motor e UM recalculo no fim: o resultado e o
            # mesmo (formulas deterministicas) sem pagar um recalculo por escrita
            pai.Range('B3').Value = idx          # em AUTOMATICO: selAnalito tem de mudar antes do motor
            ex.esperar()
            ex.xl.Calculation = -4135
            ex.run('mEstatistica.InvalidarCache')
            ex.run('mEstatistica.RecalcularAnalitoAtual', teto=1800)
            ex.xl.Calculation = -4105
            ex.xl.Calculate()
            ex.esperar()
            f = {
                'calc_B_AB': matriz(calc.Range('B3:AB182')),
                'calc_AC_AW': matriz(calc.Range('AC3:AW182')),
                'calc_AX_BG1': matriz(calc.Range('AX1:BG1')),
                'painel_A6_M8': matriz(pai.Range('A6:M8')),
                'painel_R10_Y11': matriz(pai.Range('R10:Y11')),
                # ADR-054: a faixa de status desceu para A39 na Bioquimica
                'painel_O3': pai.Range('A39').Value or pai.Range('O3').Value,
                'eng1': matriz(wb.Sheets('Eng_Saida').Range('A1:Q1')),
                'eng185': matriz(wb.Sheets('Eng_Saida').Range('A185:U186')),
            }
            linha = 14 + (idx - 1) * 2
            f['estat'] = matriz(est.Range(f'A{linha}:T{linha + 1}'))
            ev = wb.Sheets('Eventos_Westgard')
            f['eventos_total'] = val(ev.Range('J2').Value)
            foto[an] = f
            print(an, 'n N1 =', f['painel_A6_M8'][1][1], 'status', f['painel_A6_M8'][1][12])
    finally:
        ex.fechar()
    with open(saida, 'w', encoding='utf-8') as fh:
        json.dump(foto, fh, ensure_ascii=False, indent=0, default=str)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3:] or ['Lactato', 'Ácido úrico', 'Glicose'])
