# -*- coding: utf-8 -*-
"""valida_r6_xml.py -- Tester/QA, ciclo 2026-11-15_reteste06 (VALIDAR ADR-096). Leitura do XML do .xlsm, SEM Excel.

Le, na producao atual e no backup anterior a entrega 43 (so leitura, direto do zip; nada e gravado):
  * validacoes de dados (dataValidation e x14:dataValidation) das abas Registros, Painel e Digitar Resultados:
    sqref, tipo, formula1, showErrorMessage, errorStyle, errorTitle, error (OC-0074);
  * Estatistica H5/J5: formula e numFmt (OC-0075);
  * Painel A39: formula (OC-0076).
Uso: python valida_r6_xml.py <saida.json>
"""
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
      'xm': 'http://schemas.microsoft.com/office/excel/2006/main'}
ARQS = {'Bioquimica': {'nova': 'QC_Bioquimica.xlsm', 'base': '_backup_pre_integracao_2026-11-15_1802/QC_Bioquimica.xlsm'},
        'Hematologia': {'nova': 'QC_Hematologia.xlsm', 'base': '_backup_pre_integracao_2026-11-15_2059/QC_Hematologia.xlsm'}}
ABAS = ('Registros', 'Painel', 'Digitar Resultados', 'Estatística')


def mapa_abas(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    alvo = {r.get('Id'): r.get('Target') for r in rels}
    out = {}
    for s in wb.find('m:sheets', NS):
        t = alvo[s.get('{%s}id' % NS['r'])]
        out[s.get('name')] = 'xl/' + t.lstrip('/').replace('xl/', '') if not t.startswith('/xl/') else t[1:]
    return out


def numfmts(z):
    st = ET.fromstring(z.read('xl/styles.xml'))
    fm = {int(n.get('numFmtId')): n.get('formatCode') for n in st.iter('{%s}numFmt' % NS['m'])}
    xfs = [int(x.get('numFmtId') or 0) for x in st.find('m:cellXfs', NS)]
    return fm, xfs


def celula(root, ref):
    for c in root.iter('{%s}c' % NS['m']):
        if c.get('r') == ref:
            f = c.find('m:f', NS)
            v = c.find('m:v', NS)
            return {'s': int(c.get('s') or 0), 'f': f.text if f is not None else None, 'v': v.text if v is not None else None}
    return None


def validacoes(root):
    out = []
    for dv in root.iter('{%s}dataValidation' % NS['m']):
        f1 = dv.find('m:formula1', NS)
        out.append({'ns': 'main', 'sqref': dv.get('sqref'), 'type': dv.get('type'), 'formula1': f1.text if f1 is not None else None,
                    'showErrorMessage': dv.get('showErrorMessage'), 'errorStyle': dv.get('errorStyle'),
                    'errorTitle': dv.get('errorTitle'), 'error': dv.get('error')})
    for dv in root.iter('{%s}dataValidation' % NS['x14']):
        f1 = dv.find('x14:formula1/xm:f', NS)
        sq = dv.find('xm:sqref', NS)
        out.append({'ns': 'x14', 'sqref': sq.text if sq is not None else None, 'type': dv.get('type'),
                    'formula1': f1.text if f1 is not None else None, 'showErrorMessage': dv.get('showErrorMessage'),
                    'errorStyle': dv.get('errorStyle'), 'errorTitle': dv.get('errorTitle'), 'error': dv.get('error')})
    return out


def ler(caminho):
    with zipfile.ZipFile(caminho) as z:
        ab = mapa_abas(z)
        fm, xfs = numfmts(z)
        R = {}
        for a in ABAS:
            if a not in ab:
                continue
            root = ET.fromstring(z.read(ab[a]))
            d = {'validacoes': [v for v in validacoes(root) if v['type'] == 'list']}
            if a == 'Estatística':
                for ref in ('H5', 'J5', 'B5', 'D5'):
                    c = celula(root, ref)
                    if c:
                        c['numFmt'] = fm.get(xfs[c['s']], f'builtin {xfs[c["s"]]}')
                    d[ref] = c
            if a == 'Painel':
                d['A39'] = celula(root, 'A39')
            R[a] = d
        return R


def main(saida):
    out = {}
    for prod, d in ARQS.items():
        for modo, rel in d.items():
            out[f'{prod}|{modo}'] = ler(os.path.join(REPO, rel))
    json.dump(out, open(saida, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    for k, v in out.items():
        print('==', k)
        for a, d in v.items():
            for x in d['validacoes']:
                print(f'  {a:20s} {x["ns"]:4s} {str(x["sqref"])[:28]:28s} show={x["showErrorMessage"]} style={x["errorStyle"]} '
                      f'title={x["errorTitle"]!r} msg={str(x["error"])[:90]!r} f1={str(x["formula1"])[:40]!r}')
            for ref in ('H5', 'J5'):
                if ref in d:
                    print(f'  {a} {ref}: {d[ref]}')
            if 'A39' in d and d['A39']:
                print('  Painel A39 fim:', (d['A39']['f'] or '')[-260:])


if __name__ == '__main__':
    main(sys.argv[1])
