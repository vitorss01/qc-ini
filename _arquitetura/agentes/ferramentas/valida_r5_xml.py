# -*- coding: utf-8 -*-
"""valida_r5_xml.py -- Tester/QA, ciclo 2026-11-15_reteste05 (modo VALIDAR, ADR-095). Leitura DIRETA do XML.

Escrito pelo Tester/QA; NAO reaproveita o xml_formatos_listas.py do Usuario Simulado nem o qa_adr095 (L03/L04).
So zipfile + ElementTree.iterparse, sem Excel e sem openpyxl. Abre o .xlsm SO PARA LEITURA (deve ser uma copia).

O que mede, por arquivo:
  * OC-0072: numFmt cujo formatCode tem letra de data ESCAPADA (\\y, \\a ...) ou y/a entre aspas -- o Excel mostra a
    letra (ex.: "14/11/yyyy"); quais cellXfs usam esses numFmt; por aba, quantas celulas (<c s=>), colunas
    (<col style=>) e linhas (<row s= customFormat=1>) usam esses estilos (com exemplos de referencia);
  * OC-0068: quantos <x12ac:list> (lista gravada como item unico com virgula) e todas as validacoes de lista
    (aba, sqref, formula1);
  * valores: com --comparar BASE NOVA, o <v> (e o tipo t=) das celulas que na BASE tinham formato de ano literal,
    comparado com a NOVA (OC-0072: "valores inalterados").

Uso:
  python -I valida_r5_xml.py <arquivo.xlsm> [...]                  -> resumo JSON por arquivo
  python -I valida_r5_xml.py --comparar <base.xlsm> <nova.xlsm>     -> resumo dos dois + comparacao de valores
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RX_LIT = re.compile(r'\\[yYaA]|"[^"]*[yYaA][^"]*"')


def nomes_abas(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    alvo = {}
    for r in re.findall(r'<Relationship\b[^>]*>', rels):
        i = re.search(r'\bId="([^"]+)"', r).group(1)
        t = re.search(r'\bTarget="([^"]+)"', r).group(1)
        alvo[i] = t.lstrip('/').replace('xl/', '') if t.startswith('/') else t
    out = {}
    for s in re.findall(r'<sheet\b[^>]*>', wb):
        nome = re.search(r'\bname="([^"]+)"', s).group(1)
        nome = nome.replace('&amp;', '&').replace('&quot;', '"').replace('&apos;', "'")
        rid = re.search(r'\br:id="([^"]+)"', s).group(1)
        est = re.search(r'\bstate="([^"]+)"', s)
        out[nome] = {'xml': 'xl/' + alvo[rid], 'estado': est.group(1) if est else 'visible'}
    return out


def estilos(z):
    root = ET.fromstring(z.read('xl/styles.xml'))
    fmts = {}
    nf = root.find(M + 'numFmts')
    if nf is not None:
        for f in nf.findall(M + 'numFmt'):
            fmts[int(f.get('numFmtId'))] = f.get('formatCode')
    literais = {i: c for i, c in fmts.items() if RX_LIT.search(c or '')}
    xfs = root.find(M + 'cellXfs').findall(M + 'xf')
    xf_lit = {k: literais[int(x.get('numFmtId', 0))] for k, x in enumerate(xfs) if int(x.get('numFmtId', 0)) in literais}
    return literais, xf_lit, len(xfs)


def varrer_aba(z, caminho, xf_lit, alvos=None):
    """Uma passada pelo XML da aba: uso dos estilos literais, x12ac:list, validacoes de lista e (opcional) valores."""
    usos = {'celulas': 0, 'ex_celulas': [], 'colunas': [], 'linhas': 0}
    vals = {}
    n_x12 = 0
    dvs = []
    with z.open(caminho) as f:
        for ev, el in ET.iterparse(f, events=('end',)):
            tag = el.tag
            if tag == M + 'c':
                s = el.get('s')
                ref = el.get('r')
                if s is not None and int(s) in xf_lit:
                    usos['celulas'] += 1
                    if len(usos['ex_celulas']) < 6:
                        usos['ex_celulas'].append(ref)
                if alvos is not None and ref in alvos:
                    v = el.find(M + 'v')
                    vals[ref] = (el.get('t'), v.text if v is not None else None)
                el.clear()
            elif tag == M + 'row':
                s = el.get('s')
                if s is not None and el.get('customFormat') in ('1', 'true') and int(s) in xf_lit:
                    usos['linhas'] += 1
                el.clear()
            elif tag == M + 'col':
                s = el.get('style')
                if s is not None and int(s) in xf_lit:
                    usos['colunas'].append(f"{el.get('min')}-{el.get('max')}")
            elif tag == M + 'dataValidation':
                if el.get('type') == 'list':
                    f1 = el.find(M + 'formula1')
                    dvs.append({'sqref': el.get('sqref'), 'formula1': f1.text if f1 is not None else None})
    # x12ac:list fica dentro de mc:AlternateContent: conta as tags de abertura no texto cru
    n_x12_txt = z.read(caminho).count(b'<x12ac:list')
    return usos, vals, n_x12, n_x12_txt, dvs


def resumo(caminho, alvos_por_aba=None):
    z = zipfile.ZipFile(caminho)
    abas = nomes_abas(z)
    literais, xf_lit, n_xf = estilos(z)
    out = {'arquivo': caminho.split('\\')[-1], 'numFmt_literais': literais, 'xf_literais': len(xf_lit),
           'n_cellXfs': n_xf, 'abas': {}, 'x12ac_list_total': 0, 'validacoes_lista': {}}
    valores = {}
    for nome, info in abas.items():
        alvos = (alvos_por_aba or {}).get(nome)
        usos, vals, n_x12, n_x12_txt, dvs = varrer_aba(z, info['xml'], xf_lit, set(alvos) if alvos else None)
        if usos['celulas'] or usos['colunas'] or usos['linhas']:
            out['abas'][nome] = dict(usos, estado=info['estado'])
        out['x12ac_list_total'] += n_x12_txt
        if dvs:
            out['validacoes_lista'][nome] = dvs
        if alvos:
            valores[nome] = vals
    out['celulas_ano_literal_total'] = sum(a['celulas'] for a in out['abas'].values())
    return out, valores, abas


def celulas_literais(caminho):
    """{aba: [refs]} de TODAS as celulas que usam estilo com ano literal (para a comparacao de valores)."""
    z = zipfile.ZipFile(caminho)
    abas = nomes_abas(z)
    _, xf_lit, _ = estilos(z)
    out = {}
    for nome, info in abas.items():
        refs = []
        with z.open(info['xml']) as f:
            for ev, el in ET.iterparse(f, events=('end',)):
                if el.tag == M + 'c':
                    s = el.get('s')
                    if s is not None and int(s) in xf_lit:
                        refs.append(el.get('r'))
                    el.clear()
                elif el.tag == M + 'row':
                    el.clear()
        if refs:
            out[nome] = refs
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")   # -I ignora o PYTHONIOENCODING
    a = sys.argv[1:]
    if a and a[0] == '--comparar':
        base, nova = a[1], a[2]
        alvos = celulas_literais(base)
        rb, vb, _ = resumo(base, alvos)
        rn, vn, _ = resumo(nova, alvos)
        comp = {}
        for aba, refs in alvos.items():
            dif = [(r, vb.get(aba, {}).get(r), vn.get(aba, {}).get(r)) for r in refs
                   if vb.get(aba, {}).get(r) != vn.get(aba, {}).get(r)]
            comp[aba] = {'celulas_base_literal': len(refs), 'valores_diferentes': len(dif), 'ex_dif': dif[:8],
                         'faltam_na_nova': sum(1 for r in refs if r not in vn.get(aba, {}))}
        for r in (rb, rn):
            r.pop('validacoes_lista_completas', None)
        print(json.dumps({'BASE': rb, 'NOVA': rn, 'valores_base_x_nova': comp}, ensure_ascii=False, indent=1))
        return
    for c in a:
        r, _, _ = resumo(c)
        print(json.dumps(r, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
