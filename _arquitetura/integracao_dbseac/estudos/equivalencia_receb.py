# -*- coding: utf-8 -*-
"""Equivalencia do mEtlReceb (recebimento em VBA, DB_SEAC pelo ACE) com o recebimento feito pelo Power Query.
Roda numa COPIA; nada e salvo.

Uso: python equivalencia_receb.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--corte AAAA-MM-DD] [--db <DB_SEAC>]

Simula a chegada: os resultados recebidos a partir de --corte (RECEBIDO_EM) sao tirados do recebimento; o VBA le o
DB_SEAC (por padrao o caminho configurado em Cfg_Integracao) e tem de trazer EXATAMENTE esses de volta, campo a campo
iguais aos que o Power Query tinha gravado, e a tabela final na mesma ordem (DATA_HORA, ID_ORIGEM).
"""
import datetime as dt
import json
import os
import sys
import tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, AQUI)
from qa_final import QA  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import lint_vba  # noqa: E402
from equivalencia_cq import importar  # noqa: E402

SRC = os.path.join(AQUI, '..', 'src', 'comum', 'mEtlReceb.bas')
TESTE = r'''
Option Explicit
Public Function TesteReceb(ByVal caminho As String, ByVal tabela As String, ByVal prefixo As String, ByVal corte As Double, ByVal colCorte As String) As String
    Dim lo As ListObject, atual As Variant, cab As Object, i As Long, j As Long, h As Variant, kept() As Variant, nK As Long
    Dim ids As Object, nomes As Variant, dados As Variant, n As Long, msg As String, novos As Variant, nN As Long, nS As Long
    Dim res As Variant, t As Double, tLer As Double, tNovos As Double, tJuntar As Double, ws As Worksheet, cRe As Long
    On Error GoTo falha
    Set lo = ThisWorkbook.Worksheets("DB").ListObjects("tblDB_Recebimento")
    atual = lo.DataBodyRange.Value2
    Set cab = CreateObject("Scripting.Dictionary")
    h = lo.HeaderRowRange.Value
    For j = 1 To UBound(h, 2): cab(CStr(h(1, j))) = j: Next j
    cRe = cab(colCorte)
    ReDim kept(1 To UBound(atual, 1), 1 To UBound(atual, 2))
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(atual, 1)
        If atual(i, cRe) < corte Then
            nK = nK + 1
            For j = 1 To UBound(atual, 2): kept(nK, j) = atual(i, j): Next j
            ids(CStr(atual(i, cab("ID_REGISTRO")))) = 1
        End If
    Next i
    t = Timer
    msg = mEtlReceb.LerOrigem(caminho, tabela, nomes, dados, n)
    tLer = Timer - t
    If Len(msg) > 0 Then TesteReceb = "ERRO|" & msg: Exit Function
    t = Timer
    novos = mEtlReceb.NovosParaReceber(nomes, dados, n, cab, ids, prefixo, mEtlReceb.Carimbo(), nN, nS)
    tNovos = Timer - t
    t = Timer
    res = mEtlReceb.JuntarOrdenado(kept, nK, novos, nN, cab)
    tJuntar = Timer - t
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "RECEB_TESTE"
    Dim c As Variant
    For Each c In cab.Keys
        Select Case CStr(c)
            Case "NIVEL", "ID_ORIGEM", "ITEM_ID", "VALOR", "DATA", "DATA_HORA", "RECEBIDO_EM"
            Case Else: ws.Cells(1, cab(c)).Resize(UBound(res, 1), 1).NumberFormat = "@"
        End Select
    Next c
    ws.Range("A1").Resize(UBound(res, 1), UBound(res, 2)).Value2 = res
    TesteReceb = "OK|" & n & "|" & nK & "|" & nN & "|" & nS & "|" & Format$(tLer, "0.000") & "|" & Format$(tNovos, "0.000") & _
                 "|" & Format$(tJuntar, "0.000")
    Exit Function
falha:
    TesteReceb = "ERRO|" & Err.Number & "|" & Err.Description
End Function
'''


def main(produto, arq, saida, corte, db=None, col_corte='RECEBIDO_EM'):
    erros = lint_vba.checar(SRC)
    if erros:
        raise SystemExit('lint VBA falhou, Excel NAO aberto: ' + ' | '.join(erros))
    res = {'produto': produto, 'corte': corte, 'coluna_corte': col_corte}
    q = QA(produto, arq)
    try:
        caminho = db or q.cfg('CAMINHO_DB_SEAC')
        tabela, prefixo = q.cfg('TABELA_ORIGEM'), q.cfg('PREFIXO_ID')
        res['fonte'] = 'local' if db else 'rede (CAMINHO_DB_SEAC)'
        for c in list(q.wb.VBProject.VBComponents):
            if c.Name == 'mEtlReceb':
                q.wb.VBProject.VBComponents.Remove(c)
        importar(q.wb, SRC)
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mTesteReceb'
        comp.CodeModule.AddFromString(TESTE)
        if ii.compilar_vba(q.ex, q.wb, teto=20) is False:
            raise SystemExit('VBA nao compilou')
        lo = q.wb.Worksheets('DB').ListObjects('tblDB_Recebimento')
        cab = [str(c.Name) for c in lo.ListColumns]
        orig = lo.DataBodyRange.Value2
        c_corte = (dt.datetime.fromisoformat(corte) - dt.datetime(1899, 12, 30)).total_seconds() / 86400
        r = str(q.run('mTesteReceb.TesteReceb', caminho, tabela, prefixo, c_corte, col_corte, teto=900))
        print('VBA:', r, flush=True)
        if not r.startswith('OK|'):
            raise SystemExit(r)
        _, n_origem, n_mantidos, n_novos, n_sem_id, t_ler, t_novos, t_juntar = r.split('|')
        res.update({'linhas_origem': int(n_origem), 'mantidos': int(n_mantidos), 'novos_vba': int(n_novos),
                    'sem_id': int(n_sem_id), 'tempos_s': {'ler_db_seac': t_ler, 'novos': t_novos, 'juntar_ordenar': t_juntar}})
        ws = q.wb.Worksheets('RECEB_TESTE')
        n_res = int(n_mantidos) + int(n_novos)
        novo = ws.Range(ws.Cells(1, 1), ws.Cells(n_res, len(cab))).Value2
        icor = cab.index(col_corte)
        esperados = sum(1 for row in orig if row[icor] >= c_corte)
        res['esperados'] = esperados
        por_id = {str(b[0]): b for b in novo}
        ids_orig = {str(a[0]) for a in orig}
        res['ids_originais_ausentes'] = [x for x in ids_orig if x not in por_id][:10]
        res['novos_alem_do_original'] = sum(1 for x in por_id if x not in ids_orig)
        # ordem: o resultado tem de estar ordenado por (DATA_HORA, ID_ORIGEM)
        idh, iid = cab.index('DATA_HORA'), cab.index('ID_ORIGEM')
        chaves = [((b[idh] if b[idh] not in ('', None) else -1), b[iid]) for b in novo]
        res['ordenado'] = all(chaves[k] <= chaves[k + 1] for k in range(len(chaves) - 1))
        div = {}
        ex = {}
        pares = [(a, por_id[str(a[0])]) for a in orig if str(a[0]) in por_id]
        for i, (a, b) in enumerate(pares):
            for j, c in enumerate(cab):
                if c == 'RECEBIDO_EM':
                    continue
                va = None if a[j] in ('', None) else a[j]
                vb = None if b[j] in ('', None) else b[j]
                if isinstance(va, float) and isinstance(vb, float):
                    igual = abs(va - vb) < 1e-9 * max(1, abs(va))
                else:
                    igual = va == vb
                if not igual:
                    div[c] = div.get(c, 0) + 1
                    ex.setdefault(c, (i + 2, str(a[0]), str(va), str(vb)))
        res['divergencias'], res['exemplos'] = div, ex
        print(json.dumps({k: res[k] for k in ('linhas_origem', 'mantidos', 'novos_vba', 'esperados', 'novos_alem_do_original',
                                              'ids_originais_ausentes', 'ordenado', 'sem_id', 'tempos_s', 'divergencias')},
                         ensure_ascii=False), flush=True)
        for c, e in ex.items():
            print('  ex', c, e, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], a[a.index('--corte') + 1] if '--corte' in a else '2026-10-16',
         a[a.index('--db') + 1] if '--db' in a else None, 'DATA_HORA' if '--por-data' in a else 'RECEBIDO_EM')
