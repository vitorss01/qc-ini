# -*- coding: utf-8 -*-
"""valida_r5_amao.py -- Tester/QA, ciclo 2026-11-15_reteste05: conta "a mao" n e media de uma serie direto da
tblCQ_Final gravada no SaveCopyAs (feito pelo valida_r5_excel.py depois do ATUALIZAR), sem Excel (openpyxl, so leitura).

Uso: python valida_r5_amao.py <arquivo_SaveCopyAs.xlsm> <analito> <equip|-> <lote> <de dd/mm/aaaa> <ate dd/mm/aaaa>
Imprime, por nivel: n e media com e sem os manuais (ORIGEM_RESULTADO = MANUAL) e a lista dos manuais.
"""
import datetime as dt
import statistics as st
import sys
import openpyxl

arq, an, eq, lote, de, ate = sys.argv[1:7]
de = dt.datetime.strptime(de, '%d/%m/%Y').date()
ate = dt.datetime.strptime(ate, '%d/%m/%Y').date()
wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
ws = wb['Principal - Resultados']
cab = None
por = {}
man = []
for row in ws.iter_rows(values_only=True):
    if cab is None:
        if 'ID_REGISTRO' in row and 'PARTICIPA_ESTATISTICA' in row:
            cab = {c: i for i, c in enumerate(row) if c}
        continue
    r = {c: row[i] for c, i in cab.items()}
    if r['ID_REGISTRO'] is None:
        continue
    if str(r['ANALITO'] or '').strip().upper() != an.upper():
        continue
    if eq != '-' and str(r['EQUIPAMENTO'] or '').strip().upper() != eq.upper():
        continue
    if str(r['LOTE'] or '').strip() != lote:
        continue
    d = r['DATA_HORA']
    if not isinstance(d, dt.datetime) or not (de <= d.date() <= ate):
        continue
    if str(r['ORIGEM_RESULTADO']).upper() == 'MANUAL':
        man.append({k: r[k] for k in ('ID_REGISTRO', 'NIVEL', 'RESULTADO', 'PARTICIPA_ESTATISTICA', 'STATUS_ANALITICO')})
    if str(r['PARTICIPA_ESTATISTICA']).upper() != 'SIM' or not isinstance(r['RESULTADO'], (int, float)):
        continue
    nv = int(r['NIVEL'] or 0)
    por.setdefault(nv, {'todos': [], 'sem_manual': []})
    por[nv]['todos'].append(float(r['RESULTADO']))
    if str(r['ORIGEM_RESULTADO']).upper() != 'MANUAL':
        por[nv]['sem_manual'].append(float(r['RESULTADO']))
for nv in sorted(por):
    a, b = por[nv]['todos'], por[nv]['sem_manual']
    print(f'N{nv}: com manuais n={len(a)} media={st.fmean(a)!r} | sem manuais n={len(b)} media={st.fmean(b) if b else None!r}')
print('manuais no periodo:', man)
