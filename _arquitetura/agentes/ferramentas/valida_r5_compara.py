# -*- coding: utf-8 -*-
"""valida_r5_compara.py -- Tester/QA, ciclo 2026-11-15_reteste05: compara os JSON do valida_r5_excel.py (BASE = entrega
41 x NOVA = entrega 42) e resume os vereditos por cenario. So le JSON; nao abre Excel.

Uso: python valida_r5_compara.py <Bioquimica|Hematologia>
"""
import json
import os
import sys

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-15_reteste05', 'tester')


def carregar(o):
    return json.load(open(os.path.join(PASTA, f'valida_{PROD}_{o}.json'), encoding='utf-8'))


def dif_grade(a, b):
    out = []
    for i, (la, lb) in enumerate(zip(a or [], b or []), 14):
        for j, (x, y) in enumerate(zip(la, lb), 1):
            if x != y:
                out.append((i, j, x, y))
    return out


def main():
    B, N = carregar('base'), carregar('nova')
    print('tblCQ_Final', B.get('tblCQ_Final'), N.get('tblCQ_Final'))
    for nome, cn in N['painel'].items():
        if not isinstance(cn, dict) or 'engPainel' not in cn:
            continue
        cb = B['painel'].get(nome, {})
        r = {'engPainel_igual': cb.get('engPainel') == cn['engPainel'],
             'A6Q12_igual': cb.get('painel_A6Q12') == cn['painel_A6Q12'],
             'calc_igual': cb.get('calc') == cn['calc'], 'eng_saida_igual': cb.get('eng_saida') == cn.get('eng_saida'),
             'M1': (cb.get('M1'), cn['M1']),
             'AF1': (cb.get('AF1'), cn['AF1']), 'AF1_oraculo': cn['oraculo']['AF1'],
             'cortadas_oraculo': cn['oraculo']['cortadas'], 'corridas_periodo': cn['oraculo']['corridas_periodo'],
             'so_X_janela': cn['oraculo']['so_X_na_janela'],
             'confere': cn['confere']}
        if BIO:
            d = dif_grade(cb.get('estat_grade'), cn.get('estat_grade'))
            r['estat_grade_dif'] = len(d)
            r['estat_grade_dif_ex'] = d[:5]
            r['hdr_nova'] = {k: (v['txt'], v['v2']) for k, v in cn['estat_hdr'].items() if k in ('B4', 'B5', 'D5', 'D4', 'H5', 'I5', 'J5')}
            r['hdr_base'] = {k: (v['txt'], v['v2']) for k, v in cb.get('estat_hdr', {}).items() if k in ('B5', 'D5', 'D4', 'J5', 'H5')}
            r['1900_nova'] = cn['estat_1900']
            r['yyyy_nova'] = cn['estat_yyyy']
            r['janelas'] = cn['estat_janelas']
        r['A39_nova_fim'] = cn['A39'][-330:]
        r['A39_base_fim'] = cb.get('A39', '')[-330:]
        r['A39_igual'] = cb.get('A39') == cn['A39']
        print('\n==', nome, json.dumps(r, ensure_ascii=False, default=str))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
