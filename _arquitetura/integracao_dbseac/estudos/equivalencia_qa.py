# -*- coding: utf-8 -*-
"""Equivalencia do mEtlQA (conferencia em VBA) com a tblQA_Integracao do Power Query. Roda numa COPIA; nada salvo.

Uso: python equivalencia_qa.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--fixtures]

O PQ refaz a tblCQ_Final e a tblQA_Integracao contra o DB_SEAC atual (com os casos de teste, se --fixtures); logo
depois o VBA calcula a mesma coisa com a mesma fonte. Comparacao: o CONJUNTO de achados (as 8 colunas), e a ordem
das partes (severidade, codigo) -- dentro de um codigo o PQ desempata com um introsort instavel.
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
        print(json.dumps({k: res[k] for k in ('linhas', 'por_codigo_pq', 'por_codigo_vba', 'ordem_das_partes_igual',
                                              'tempo_qa_vba_s')}, ensure_ascii=False), flush=True)
        for k in ('so_no_pq', 'so_no_vba'):
            for x in res[k]:
                print(' ', k, x, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)


def importar_mod(wb, caminho):
    nome = os.path.splitext(os.path.basename(caminho))[0]
    for c in list(wb.VBProject.VBComponents):
        if c.Name == nome:
            wb.VBProject.VBComponents.Remove(c)
    importar(wb, caminho)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], '--fixtures' in a)
