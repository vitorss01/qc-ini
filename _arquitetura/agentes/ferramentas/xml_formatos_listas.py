# -*- coding: utf-8 -*-
"""xml_formatos_listas.py -- Usuario Simulado (ciclo 2026-11-15_reteste05_parte2). Leitura DIRETA do XML de um .xlsm
(zipfile, sem Excel, sem openpyxl), oraculo independente das suites qa_adr095 (L03/L04).

Funcoes:
  abas(z)                    -> {nome da aba: caminho do XML}
  formatos_ano_literal(z)    -> formatos (numFmt) com ano/letra de data ESCAPADA (\\y, "y", \\a ...): o Excel mostra a
                                letra literal ("14/11/yyyy"); mais o mapa cellXfs -> formatCode
  uso_formatos(z, ids_xf)    -> por aba: celulas (<c s=>), colunas (<col style=>) e linhas (<row s= customFormat>) que
                                usam esses estilos
  listas_x12ac(z)            -> validacoes de lista gravadas como item unico com virgula (<x12ac:list>)
  validacoes_lista(z)        -> todas as validacoes tipo lista (aba, sqref, formula1)
  valores(z, aba, refs)      -> {ref: (valor bruto <v>/texto, s)} das celulas pedidas (shared strings resolvidas)

Uso: python -I xml_formatos_listas.py <arquivo.xlsm> [<arquivo2.xlsm>]   (imprime JSON do resumo)
So LEITURA do arquivo indicado (que deve ser uma COPIA). Nada e gravado.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
# formatos internos de data (numFmtId < 164) -- nunca tem letra escapada
EMBUTIDOS_DATA = {14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
                  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0'}
RX_LITERAL = re.compile(r'\\[yYaAdDmMhHsS]|"[^"]*[yYaA][^"]*"')


def abas(z):
    wbx = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    alvo = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
            for r in re.findall(r'<Relationship [^>]*/>', rels)}
    out = {}
    for tag in re.findall(r'<sheet [^>]*/>', wbx):
        nome = re.search(r'name="([^"]+)"', tag).group(1)
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        nome = nome.replace('&amp;', '&')
        out[nome] = 'xl/' + alvo[rid].replace('/xl/', '').lstrip('/')
    return out


def formatos(z):
    st = ET.fromstring(z.read('xl/styles.xml'))
    nf = {int(n.get('numFmtId')): n.get('formatCode') for n in st.iter(M + 'numFmt')}
    xfs = st.find('m:cellXfs', NS)
    xf_fmt = [int(x.get('numFmtId') or 0) for x in xfs.findall('m:xf', NS)]
    return nf, xf_fmt


def formatos_ano_literal(z):
    nf, xf_fmt = formatos(z)
    ruins = {i: c for i, c in nf.items() if RX_LITERAL.search(c or '')}
    ids_xf = [k for k, f in enumerate(xf_fmt) if f in ruins]
    return {'numFmt_com_letra_literal': ruins, 'xf_ids': ids_xf}


def _iter_sheet(z, caminho):
    with z.open(caminho) as f:
        for ev, el in ET.iterparse(f, events=('end',)):
            yield el


def uso_formatos(z, ids_xf, so_abas=None):
    ids = {str(i) for i in ids_xf}
    out = {}
    if not ids:
        return out
    for nome, cam in abas(z).items():
        if so_abas and nome not in so_abas:
            continue
        cel, cols, rows = [], [], []
        for el in _iter_sheet(z, cam):
            if el.tag == M + 'c':
                if el.get('s') in ids:
                    cel.append(el.get('r'))
                el.clear()
            elif el.tag == M + 'col' and el.get('style') in ids:
                cols.append(f"{el.get('min')}-{el.get('max')}")
            elif el.tag == M + 'row':
                if el.get('customFormat') == '1' and el.get('s') in ids:
                    rows.append(el.get('r'))
                el.clear()
        if cel or cols or rows:
            out[nome] = {'n_celulas': len(cel), 'celulas_amostra': cel[:12], 'colunas': cols, 'linhas': rows[:12],
                         'n_linhas': len(rows)}
    return out


def listas_x12ac(z):
    out = []
    for nome, cam in abas(z).items():
        x = z.read(cam).decode('utf8')
        for m in re.finditer(r'<x12ac:list>([^<]*)</x12ac:list>', x):
            ini = x.rfind('sqref', 0, m.start())
            sq = re.search(r'sqref="([^"]+)"', x[ini - 2:ini + 200]) if ini >= 0 else None
            out.append((nome, sq.group(1)[:40] if sq else '?', m.group(1)[:160]))
    return out


def validacoes_lista(z):
    out = []
    for nome, cam in abas(z).items():
        x = z.read(cam).decode('utf8')
        for dv in re.findall(r'<dataValidation [^>]*type="list"[^>]*>.*?</dataValidation>', x, re.S):
            sq = re.search(r'sqref="([^"]+)"', dv)
            f1 = re.search(r'<formula1>(.*?)</formula1>', dv, re.S)
            out.append((nome, sq.group(1)[:40] if sq else '?', (f1.group(1) if f1 else '')[:200]))
        for dv in re.findall(r'<x14:dataValidation [^>]*type="list"[^>]*>.*?</x14:dataValidation>', x, re.S):
            sq = re.search(r'<xm:sqref>([^<]+)</xm:sqref>', dv)
            f1 = re.search(r'<xm:f>(.*?)</xm:f>', dv, re.S)
            out.append((nome + ' (x14)', sq.group(1)[:40] if sq else '?', (f1.group(1) if f1 else '')[:200]))
    return out


def shared_strings(z):
    try:
        f = z.open('xl/sharedStrings.xml')
    except KeyError:
        return []
    out = []
    for ev, el in ET.iterparse(f, events=('end',)):
        if el.tag == M + 'si':
            out.append(''.join(t.text or '' for t in el.iter(M + 't')))
            el.clear()
    return out


def valores(z, aba, filtro, ss=None):
    """filtro(ref) -> bool. Devolve {ref: [valor, s, tipo]} (valor = texto do <v> ou a shared string)."""
    cam = abas(z)[aba]
    out = {}
    for el in _iter_sheet(z, cam):
        if el.tag == M + 'c':
            r = el.get('r')
            if filtro(r):
                v = el.find('m:v', NS)
                tv = v.text if v is not None else None
                tp = el.get('t')
                if tp == 's' and tv is not None and ss is not None:
                    tv = ss[int(tv)]
                elif tp == 'inlineStr':
                    tv = ''.join(t.text or '' for t in el.iter(M + 't'))
                out[r] = [tv, el.get('s'), tp]
            el.clear()
        elif el.tag == M + 'row':
            el.clear()
    return out


def resumo(caminho):
    z = zipfile.ZipFile(caminho)
    fa = formatos_ano_literal(z)
    nf, xf_fmt = formatos(z)
    datas = {i: c for i, c in nf.items() if re.search(r'[dDyYaA]', re.sub(r'"[^"]*"|\\.', '', c or ''))
             and re.search(r'[mM]', c or '')}
    return {'numFmt_com_letra_literal': fa['numFmt_com_letra_literal'], 'xf_ids_literais': fa['xf_ids'],
            'uso_literais': uso_formatos(z, fa['xf_ids']), 'formatos_de_data': datas,
            'x12ac_list': listas_x12ac(z), 'validacoes_lista': validacoes_lista(z)}


if __name__ == '__main__':
    for arq in sys.argv[1:]:
        print(json.dumps({arq.split('\\')[-1]: resumo(arq)}, ensure_ascii=False, indent=1))


# ------------------------------------------------------------------------------------------- OC-0072: alvos
def _col(ref):
    return re.match(r'[A-Z]+', ref).group(0)


def _lin(ref):
    return int(re.search(r'\d+', ref).group(0))


ALVOS_BIO = {
    'Configuração': lambda r: _col(r) in ('D', 'AD') and 26 <= _lin(r) <= 125,
    'Eng_Saida': lambda r: _col(r) == 'AC' and 3 <= _lin(r) <= 182,
    'Registros': lambda r: r in ('E4', 'F4'),
    'Estatística': lambda r: r in ('E8', 'F8', 'E9', 'F9', 'C10'),
    'Sync_Log': lambda r: _col(r) in ('A', 'B', 'L', 'M') and _lin(r) >= 2,
    'Westgard_Fotos': lambda r: _col(r) == 'F',
    'LotesStore': lambda r: r == 'U1',
}


def alvos(caminho, mapa=ALVOS_BIO):
    """Por aba-alvo: {ref: [valor bruto, formatCode do estilo da celula]} + formato da coluna (<col style>)."""
    z = zipfile.ZipFile(caminho)
    nf, xf_fmt = formatos(z)
    ss = shared_strings(z)

    def code(s):
        if s is None:
            return 'General(0)'
        f = xf_fmt[int(s)]
        return nf.get(f, EMBUTIDOS_DATA.get(f, f'builtin{f}'))
    out = {}
    for aba, filt in mapa.items():
        if aba not in abas(z):
            continue
        v = valores(z, aba, filt, ss)
        out[aba] = {r: [x[0], code(x[1])] for r, x in v.items()}
    # estilo de coluna da Westgard_Fotos F
    x = z.read(abas(z)['Westgard_Fotos']).decode('utf8')
    cols = re.findall(r'<col [^>]*/>', x)
    out['_cols_Westgard_Fotos'] = [(c, code(re.search(r'style="(\d+)"', c).group(1)) if 'style=' in c else '')
                                   for c in cols if re.search(r'min="[1-6]"', c)]
    return out
