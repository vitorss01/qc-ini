# -*- coding: utf-8 -*-
"""instalar_adr095.py -- ADR-095 (14/11/2026, reteste 04 + revisao independente): ano literal nos formatos de data.

Por esta automacao, Range.NumberFormat = 'dd/mm/yyyy' e lido na lingua do Excel instalado; no pt-BR o 'y' nao e codigo
de ano e o Excel grava o formato 'dd/mm/\\y\\y\\y\\y' -- a celula mostra "14/11/yyyy". Varios instaladores antigos da
cadeia fazem isso (instalar_adr060/073/083/087, instalar_integracao, etapa1 v2) e as duas producoes tem celulas assim
em abas visiveis depois do login (Eventos_Westgard da Hema, Registros De/Ate, Configuracao, Eng_Saida, Westgard_Fotos).

Este passo roda no FIM da cadeia (antes da blindagem) e conserta TODA celula ou coluna cujo formato tem o ano
literal, pelo XML do arquivo salvo pelo passo anterior: o formato vira o mesmo com o codigo de ano do Excel
(camada_dados.fmt_data, NumberFormatLocal). SO APRESENTACAO: nenhum valor, formula, regra ou criterio muda.
O qa_adr095 L04 le o XML da copia instalada e reprova qualquer formato com ano literal.

Uso:  python instalar_adr095.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]
Idempotente (a 2a execucao registra 0 mudancas).
"""
import os
import re
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import camada_dados as cd  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
BS = chr(92)
ANO_LITERAL = BS + 'y'                       # o '\y' do XML: ano gravado como texto


def log(m):
    print(m, flush=True)


def _col_num(letras):
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def corrigido(codigo):
    """'dd/mm/\\y\\y\\y\\y\\ hh:mm' -> 'dd/mm/yyyy hh:mm:ss' (o formato pretendido, em ingles). Hora sem segundos ganha
    os segundos: ADR-094 (toda data-hora mostrada com segundos; a Westgard_Fotos F ja e assim nas linhas)."""
    c = codigo.replace(ANO_LITERAL, 'y').replace(BS + ' ', ' ').replace('&quot;', '"')
    if 'hh:mm' in c and 'hh:mm:ss' not in c:
        c = c.replace('hh:mm', 'hh:mm:ss')
    return c


def ano_literal(caminho):
    """{aba: {'celulas': {(col, linha): codigo}, 'colunas': {col_num: codigo}}} das celulas e colunas cujo formato tem
    o ano literal -- lido do XML do arquivo (sem Excel)."""
    z = zipfile.ZipFile(caminho)
    st = z.read('xl/styles.xml').decode('utf8')
    fm = {m.group(1): m.group(2) for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', st)}
    ruins = {k for k, v in fm.items() if ANO_LITERAL in v}
    xfs = re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', st, re.S).group(1)
    xf = re.findall(r'<xf [^>]*?numFmtId="(\d+)"', xfs)
    estilo = {str(i): fm[n] for i, n in enumerate(xf) if n in ruins}
    wbx = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    alvo = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
            for r in re.findall(r'<Relationship [^>]*/>', rels)}
    out = {}
    if not estilo:
        return out
    for nome, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        x = z.read('xl/' + alvo[rid].replace('/xl/', '')).decode('utf8')
        cel, cols = {}, {}
        for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)/?>', x):
            s = re.search(r'\ss="(\d+)"', m.group(3))
            if s and s.group(1) in estilo:
                cel[(m.group(1), int(m.group(2)))] = estilo[s.group(1)]
        for m in re.finditer(r'<col ([^>]*?)/>', x):
            a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
            if a.get('style') in estilo:
                for n in range(int(a['min']), min(int(a['max']), 16384) + 1):
                    cols[n] = estilo[a['style']]
        if cel or cols:
            out[nome.replace('&amp;', '&')] = {'celulas': cel, 'colunas': cols}
    return out


def faixas(celulas):
    """Agrupa {(col, linha): codigo} em faixas continuas por coluna e codigo: [(col, l0, l1, codigo)]."""
    por = {}
    for (c, r), cod in celulas.items():
        por.setdefault((c, cod), []).append(r)
    out = []
    for (c, cod), rs in sorted(por.items()):
        rs.sort()
        ini = ant = rs[0]
        for r in rs[1:] + [None]:
            if r is not None and r == ant + 1:
                ant = r
                continue
            out.append((c, ini, ant, cod))
            if r is not None:
                ini = ant = r
    return out


def consertar(wb, achados):
    app = wb.Application
    ano = str(app.International[19 - 1] or 'y')
    ruins = []
    for aba, d in achados.items():
        ws = wb.Worksheets(aba)
        for n, cod in sorted(d['colunas'].items()):
            rng = ws.Columns(n)
            cd.fmt_data(rng, corrigido(cod))
            MUDANCAS.append(f'{aba} coluna {n}: {cod!r} -> {corrigido(cod)!r}')
        for c, l0, l1, cod in faixas(d['celulas']):
            rng = ws.Range(f'{c}{l0}:{c}{l1}')
            cd.fmt_data(rng, corrigido(cod))
            esperado = corrigido(cod).replace('y', ano)
            if str(rng.Cells(1, 1).NumberFormatLocal) != esperado:
                ruins.append(f'{aba}!{c}{l0}: {rng.Cells(1, 1).NumberFormatLocal!r} != {esperado!r}')
            MUDANCAS.append(f'{aba}!{c}{l0}:{c}{l1} {cod!r} -> {corrigido(cod)!r}')
    if ruins:
        raise SystemExit('formato nao ficou como o esperado: ' + '; '.join(ruins[:6]) + ' -- nada foi salvo')


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    achados = ano_literal(caminho)
    n = sum(len(d['celulas']) + len(d['colunas']) for d in achados.values())
    log(f'1. XML: {n} celula(s)/coluna(s) com ano literal em {len(achados)} aba(s): '
        + ', '.join(f'{a} ({len(d["celulas"])} cel, {len(d["colunas"])} col)' for a, d in achados.items()))
    if not achados:
        log('MUDANCAS 0')
        return
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.EnableEvents = False
        log('2. formatos com o codigo de ano do Excel (camada_dados.fmt_data)')
        consertar(wb, achados)
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS[:40]) if MUDANCAS else ''))
        if salvar:
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()
    if salvar:
        resto = ano_literal(caminho)
        if resto:
            raise SystemExit('ainda ha formato com ano literal depois de salvar: '
                             + ', '.join(f'{a} ({len(d["celulas"])} cel, {len(d["colunas"])} col)' for a, d in resto.items()))
        log('3. XML conferido depois de salvar: nenhum formato com ano literal')


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
