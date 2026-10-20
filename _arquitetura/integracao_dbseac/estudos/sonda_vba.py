# -*- coding: utf-8 -*-
"""Sondas de desempenho do VBA nesta maquina (reengenharia da camada de dados, 20/10/2026). Roda numa COPIA; nada salvo.

Uso: python sonda_vba.py <copia_QC_Bioquimica.xlsm> <copia_local_DB_SEAC.xlsm> <saida.json>

Injeta um modulo temporario na copia e mede, de dentro do Excel (Timer do VBA):
  S1 ler a tblCQ_Final inteira (33 mil x 40) para uma matriz;
  S2 gravar essa matriz numa aba nova (bloco unico) e de volta na propria tabela;
  S3 indexar por ID num Scripting.Dictionary e percorrer as linhas (custo de um ETL em memoria);
  S4 baixar a cauda do CSV do SIPEC (WinHttp, cabecalho Range) e quebrar em linhas e campos;
  S5 ler o DB_SEAC (copia local) pelo driver ACE/ADO, sem abrir o arquivo no Excel: abrir a conexao, contar, e ler
     so as linhas com ID_ORIGEM acima de um corte (o incremento).
"""
import json
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
from qa_final import QA  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

URL_BIO = 'https://sipec.ini.fiocruz.br/relatorios/relatorio/lab/Hematologia/Indicadores%20Controle%20EXL200/relat209.csv'

MODULO = r'''
Option Explicit

Public Function S1_Ler() As String
    Dim t As Double, a As Variant, lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Principal - Resultados").ListObjects("tblCQ_Final")
    t = Timer
    a = lo.DataBodyRange.Value
    S1_Ler = "OK|" & UBound(a, 1) & "|" & UBound(a, 2) & "|" & Format$(Timer - t, "0.000")
End Function

Public Function S2_Gravar(ByVal senha As String) As String
    Dim t As Double, a As Variant, lo As ListObject, ws As Worksheet, t1 As Double, t2 As Double
    Set lo = ThisWorkbook.Worksheets("Principal - Resultados").ListObjects("tblCQ_Final")
    a = lo.DataBodyRange.Value
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets.Add
    t = Timer
    ws.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    t1 = Timer - t
    ThisWorkbook.Worksheets("Principal - Resultados").Unprotect senha
    t = Timer
    lo.DataBodyRange.Value = a
    t2 = Timer - t
    S2_Gravar = "OK|" & Format$(t1, "0.000") & "|" & Format$(t2, "0.000")
End Function

Public Function S3_Memoria() As String
    Dim t As Double, a As Variant, lo As ListObject, d As Object, i As Long, n As Long, k As String, s As Double
    Set lo = ThisWorkbook.Worksheets("Principal - Resultados").ListObjects("tblCQ_Final")
    a = lo.DataBodyRange.Value
    t = Timer
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(a, 1)
        k = CStr(a(i, 1))
        If Not d.Exists(k) Then d.Add k, i
        If IsNumeric(a(i, 10)) Then s = s + 1
    Next i
    n = d.Count
    S3_Memoria = "OK|" & n & "|" & Format$(Timer - t, "0.000")
End Function

Public Function S4_Http(ByVal url As String, ByVal bytes As Long) As String
    Dim t As Double, h As Object, txt As String, linhas() As String, i As Long, f() As String, nCampos As Long
    Dim tBaixar As Double, tQuebrar As Double, sep As String
    t = Timer
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "GET", url, False
    h.SetRequestHeader "Range", "bytes=-" & bytes
    h.Send
    txt = h.ResponseText
    tBaixar = Timer - t
    t = Timer
    linhas = Split(txt, vbLf)
    If InStr(linhas(1), ";") > 0 Then sep = ";" Else sep = ","
    For i = 1 To UBound(linhas)
        f = Split(linhas(i), sep)
        If UBound(f) + 1 > nCampos Then nCampos = UBound(f) + 1
    Next i
    tQuebrar = Timer - t
    S4_Http = "OK|" & h.Status & "|" & Len(txt) & "|" & (UBound(linhas) + 1) & "|" & nCampos & "|" & _
              Format$(tBaixar, "0.000") & "|" & Format$(tQuebrar, "0.000") & "|" & sep
End Function

Public Function S5_Ace(ByVal arq As String, ByVal aba As String, ByVal corte As Double) As String
    Dim t As Double, cn As Object, rs As Object, tAbrir As Double, tContar As Double, tDelta As Double
    Dim n As Long, nd As Long, v As Variant
    On Error GoTo falha
    t = Timer
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & arq & ";Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";"
    tAbrir = Timer - t
    t = Timer
    Set rs = cn.Execute("SELECT COUNT(*) FROM [" & aba & "$]")
    n = rs.Fields(0).Value
    tContar = Timer - t
    t = Timer
    Set rs = cn.Execute("SELECT * FROM [" & aba & "$] WHERE ID_ORIGEM > " & Str$(corte))
    If Not rs.EOF Then v = rs.GetRows: nd = UBound(v, 2) + 1
    tDelta = Timer - t
    cn.Close
    S5_Ace = "OK|" & n & "|" & nd & "|" & Format$(tAbrir, "0.000") & "|" & Format$(tContar, "0.000") & "|" & Format$(tDelta, "0.000")
    Exit Function
falha:
    S5_Ace = "ERRO|" & Err.Number & "|" & Err.Description
End Function
'''


def main(arq, dbseac, saida):
    res = {}
    q = QA('Bioquimica', arq)
    try:
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mSonda'
        comp.CodeModule.AddFromString(MODULO)
        for nome, args in (('S1_Ler', ()), ('S3_Memoria', ()), ('S4_Http', (URL_BIO, 2700000)),
                           ('S4_Http', (URL_BIO, 300000)),
                           ('S5_Ace', (dbseac, 'DADOS_BIOQUIMICA', 366000)),
                           ('S5_Ace', (dbseac, 'DADOS_BIOQUIMICA', 366000)), ('S2_Gravar', (ii.SENHA,))):
            t = time.perf_counter()
            r = str(q.run('mSonda.' + nome, *args, teto=900))
            res.setdefault(nome, []).append({'retorno': r, 'parede_s': round(time.perf_counter() - t, 3)})
            print(nome, r, flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), sys.argv[3])
