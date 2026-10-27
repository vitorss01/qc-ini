# -*- coding: utf-8 -*-
"""Equivalencia do mEtlQA (conferencia em VBA) com a tblQA_Integracao do Power Query. Roda numa COPIA; nada salvo.

Uso: python equivalencia_qa.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--fixtures]

O PQ refaz a tblCQ_Final e a tblQA_Integracao contra o DB_SEAC atual (com os casos de teste, se --fixtures); logo
depois o VBA calcula a mesma coisa com a mesma fonte. Comparacao: o CONJUNTO de achados (as 8 colunas), e a ordem
das partes (severidade, codigo) -- dentro de um codigo o PQ desempata com um introsort instavel.

--fixtures: os MESMOS casos do equivalencia_cq.py (fixtures), inclusive os do ADR-082 (SUBSTITUI_ID: substituicao
valida, chave diferente, ID inexistente, ID de manual, dois manuais para o mesmo ID, interfaceamento INATIVADO,
SEM_VALOR/SEM_DATA_HORA/retransmissao indicados, sem MOTIVO). Com eles, 'adr082' no JSON exige E12 e I04 presentes
nos DOIS motores (e a mesma contagem) e confere os textos contra o ESPERADO (revisao do ADR-082): nenhum E09 de
resultado substituido; o A01 de um manual em conflito com um automatico ja substituido diz que vale o outro manual (e
nao tem A09); o A06 de um manual que substitui nao lista o substituido nem manda usar o SUBSTITUI_ID; o I04 de quem
continua inativado manda reativar; o E12 do 2o manual com o 1o INATIVADO manda apagar o SUBSTITUI_ID do 1o (sem A09)
-- a copia precisa da coluna SUBSTITUI_ID (instalar_adr082.py).
"""
import collections
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, AQUI)
from qa_final import QA  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import lint_vba  # noqa: E402
from equivalencia_cq import importar, fixtures  # noqa: E402

MODS = [os.path.join(AQUI, '..', 'src', 'comum', m) for m in ('mEtlCQ.bas', 'mEtlReceb.bas', 'mEtlQA.bas')]
TESTE = r'''
Option Explicit
Public Function TesteQA(ByVal caminho As String, ByVal tabela As String, ByVal historico As Boolean) As String
    Dim F As Variant, lo As ListObject, rec As Variant, cab As Object, h As Variant, j As Long, nomes As Variant, dados As Variant
    Dim n As Long, msg As String, q As Variant, t As Double, tQA As Double, ws As Worksheet, cols As Variant
    On Error GoTo falha
    F = mEtlCQ.CalcularCQFinal()
    Set lo = ThisWorkbook.Worksheets("DB").ListObjects("tblDB_Recebimento")
    rec = lo.DataBodyRange.Value2
    Set cab = CreateObject("Scripting.Dictionary")
    h = lo.HeaderRowRange.Value
    For j = 1 To UBound(h, 2): cab(CStr(h(1, j))) = j: Next j
    If Not historico Then
        msg = mEtlReceb.LerOrigem(caminho, tabela, nomes, dados, n)
        If Len(msg) > 0 Then TesteQA = "ERRO|" & msg: Exit Function
    End If
    t = Timer
    q = mEtlQA.CalcularQA(F, rec, UBound(rec, 1), cab, dados, nomes, n, historico)
    tQA = Timer - t
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "QA_TESTE"
    cols = mEtlQA.NomesColunasQA()
    ws.Range("A1").Resize(1, 8).Value = cols
    If Not IsEmpty(q) Then
        ws.Range("A2").Resize(UBound(q, 1), 8).NumberFormat = "@"
        ws.Range("F2").Resize(UBound(q, 1), 1).NumberFormat = "0"
        ws.Range("G2").Resize(UBound(q, 1), 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ws.Range("A2").Resize(UBound(q, 1), 8).Value2 = q
        TesteQA = "OK|" & UBound(q, 1) & "|" & Format$(tQA, "0.000") & "|" & mEtlQA.TemposUltimoQA()
    Else
        TesteQA = "OK|0|" & Format$(tQA, "0.000")
    End If
    Exit Function
falha:
    TesteQA = "ERRO|" & Err.Number & "|" & Err.Description
End Function
'''


def chave(row):
    out = []
    for i, v in enumerate(row):
        if v in ('', None):
            v = None
        elif i == 5:
            v = int(float(v))
        elif i == 6 and hasattr(v, 'year'):
            v = v.replace(tzinfo=None, microsecond=0).isoformat()
        out.append(v)
    return tuple(out)


def main(produto, arq, saida, fix=False):
    for m in MODS:
        erros = lint_vba.checar(m)
        if erros:
            raise SystemExit('lint VBA falhou, Excel NAO aberto: ' + ' | '.join(erros))
    res = {'produto': produto}
    q = QA(produto, arq)
    try:
        caminho, tabela = q.cfg('CAMINHO_DB_SEAC'), q.cfg('TABELA_ORIGEM')
        historico = str(q.cfg('MODO_FONTE') or '').strip().upper().startswith('HIST')
        for m in MODS:
            importar_mod(q.wb, m)
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mTesteQA'
        comp.CodeModule.AddFromString(TESTE)
        if ii.compilar_vba(q.ex, q.wb, teto=20) is False:
            raise SystemExit('VBA nao compilou')
        if fix:
            fixtures(q)                                   # ja refaz a tblCQ_Final pelo PQ
        else:
            q.run('mIntegracao.Refrescar', 'tblCQ_Final', teto=900)
        q.run('mIntegracao.Refrescar', 'tblQA_Integracao', teto=900)
        r = str(q.run('mTesteQA.TesteQA', caminho, tabela, historico, teto=900))
        print('VBA:', r, flush=True)
        if not r.startswith('OK|'):
            raise SystemExit(r)
        lo = q.wb.Worksheets('QA_INTEGRACAO').ListObjects('tblQA_Integracao')
        pq = [chave(x) for x in lo.DataBodyRange.Value] if lo.ListRows.Count else []
        n_vba = int(r.split('|')[1])
        ws = q.wb.Worksheets('QA_TESTE')
        vba = [chave(x) for x in ws.Range(ws.Cells(2, 1), ws.Cells(n_vba + 1, 8)).Value] if n_vba else []
        res['linhas'] = {'pq': len(pq), 'vba': len(vba)}
        res['por_codigo_pq'] = dict(collections.Counter(x[1] for x in pq))
        res['por_codigo_vba'] = dict(collections.Counter(x[1] for x in vba))
        cp, cv = collections.Counter(pq), collections.Counter(vba)
        res['so_no_pq'] = [list(map(str, k)) for k in (cp - cv).keys()][:8]
        res['so_no_vba'] = [list(map(str, k)) for k in (cv - cp).keys()][:8]
        partes = lambda xs: [k for k, _ in __import__('itertools').groupby(x[1] for x in xs)]
        res['ordem_das_partes_igual'] = partes(pq) == partes(vba)
        res['tempo_qa_vba_s'] = r.split('|')[2]
        if fix:
            # ADR-082: os casos do SUBSTITUI_ID tem de aparecer (e iguais) nos dois motores
            cods = ('E04', 'E09', 'E12', 'A01', 'A06', 'A09', 'I04')
            res['adr082'] = {'por_codigo_pq': {c: res['por_codigo_pq'].get(c, 0) for c in cods},
                             'por_codigo_vba': {c: res['por_codigo_vba'].get(c, 0) for c in cods},
                             'E12_amostra': [x[7] for x in pq if x[1] == 'E12'][:4],
                             'I04_amostra': [x[7] for x in pq if x[1] == 'I04'][:4]}
            res['adr082']['textos'] = conferir_textos_adr082(q, pq, vba)
            res['adr082']['ok'] = (res['por_codigo_pq'].get('E12', 0) >= 1 and res['por_codigo_pq'].get('I04', 0) >= 1
                                   and all(res['por_codigo_pq'].get(c, 0) == res['por_codigo_vba'].get(c, 0)
                                           for c in cods)
                                   and not res['adr082']['textos'])
            print('ADR-082', json.dumps(res['adr082'], ensure_ascii=False, default=str), flush=True)
        print(json.dumps({k: res[k] for k in ('linhas', 'por_codigo_pq', 'por_codigo_vba', 'ordem_das_partes_igual',
                                              'tempo_qa_vba_s')}, ensure_ascii=False), flush=True)
        for k in ('so_no_pq', 'so_no_vba'):
            for x in res[k]:
                print(' ', k, x, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)


def listados_a06(det):
    """Os IDs que o A06 lista (depois de 'equipamento/lote/nivel/analito...: ' ate '. Os listados')."""
    s = str(det or '')
    p = s.find('analito no dia: ')
    p = s.find('analito: ') if p < 0 else p
    if p < 0:
        return []
    s = s[s.find(': ', p) + 2:]
    return [x.strip() for x in s.split('. Os listados')[0].split(',') if x.strip()]


def conferir_textos_adr082(q, pq, vba):
    """Revisao do ADR-082: os textos do QA contra o esperado, nos DOIS motores, a partir da tblCQ_Final do PQ. Devolve a
    lista do que nao bateu (vazia = ok). Achado = (sev, cod, teste, id, analito, nivel, data_hora, detalhe)."""
    lo = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
    cab = [str(c.Name) for c in lo.ListColumns]
    fin = {}
    for r in lo.DataBodyRange.Value:
        d = dict(zip(cab, r))
        fin[str(d['ID_REGISTRO'])] = d
    rel = lambda d: str(d.get('ID_RELACIONADO') or '')                          # noqa: E731
    alvos = {i: rel(d) for i, d in fin.items() if d['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and rel(d).startswith('MAN_')}
    falhas = []
    for nome, ach in (('pq', pq), ('vba', vba)):
        por = collections.defaultdict(list)
        for x in ach:
            por[(x[1], str(x[3] or ''))].append(str(x[7] or ''))
        for i, man in alvos.items():
            if por.get(('E09', i)):
                falhas.append(f'{nome}: E09 de {i}, que o manual {man} substitui')
            dets = por.get(('I04', i), [])
            fim = ('Para voltar ao automatico, apague o SUBSTITUI_ID (ou a linha) do manual.'
                   if fin[i]['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' else
                   'reative-o (aba Inativar) e apague o SUBSTITUI_ID (ou a linha) do manual.')
            if len(dets) != 1 or not dets[0].endswith(fim):
                falhas.append(f'{nome}: I04 de {i} ({fin[i]["STATUS_ANALITICO"]}) {dets[:1]}')
        for i, d in fin.items():
            if d['ORIGEM_RESULTADO'] != 'MANUAL':
                continue
            if d['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and str(d.get('MOTIVO_EXCLUSAO_AUTOMATICA') or '').startswith(
                    'coincide') and rel(d) in alvos:
                a01 = por.get(('A01', i), [])
                if len(a01) != 1 or f'vale o manual {alvos[rel(d)]}, que ja substitui o automatico {rel(d)}' not in a01[0]:
                    falhas.append(f'{nome}: A01 de {i} (conflito com {rel(d)}, ja substituido) {a01[:1]}')
                if por.get(('A09', i)):
                    falhas.append(f'{nome}: A09 de {i} (o automatico {rel(d)} ja esta substituido)')
            # 2o manual para o mesmo ID com o 1o INATIVADO: o E12 diz que o 1o continua substituindo; sem A09
            if d['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and 'SUBSTITUI_ID invalido: outro manual (' in str(
                    d.get('MOTIVO_EXCLUSAO_AUTOMATICA') or '') and fin.get(rel(d), {}).get('STATUS_ANALITICO') == 'INATIVADO':
                e12 = por.get(('E12', i), [])
                if len(e12) != 1 or f'O manual {rel(d)} esta INATIVADO, mas continua substituindo' not in e12[0] \
                        or 'Corrija o SUBSTITUI_ID' in e12[0]:
                    falhas.append(f'{nome}: E12 de {i} (o 1o manual {rel(d)} esta inativado) {e12[:1]}')
                if por.get(('A09', i)):
                    falhas.append(f'{nome}: A09 de {i} (quem esta inativado e o manual {rel(d)})')
            if d['STATUS_ANALITICO'] == 'ATIVO' and rel(d):
                for det in por.get(('A06', i), []):
                    if not det.startswith(f'este manual substitui {rel(d)};') or rel(d) in listados_a06(det) \
                            or 'SUBSTITUI_ID' in det:
                        falhas.append(f'{nome}: A06 de {i} (substitui {rel(d)}) {det[:160]}')
    return falhas[:20]


def importar_mod(wb, caminho):
    nome = os.path.splitext(os.path.basename(caminho))[0]
    for c in list(wb.VBProject.VBComponents):
        if c.Name == nome:
            wb.VBProject.VBComponents.Remove(c)
    importar(wb, caminho)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], '--fixtures' in a)
