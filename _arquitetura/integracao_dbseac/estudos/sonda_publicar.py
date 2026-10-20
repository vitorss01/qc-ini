# -*- coding: utf-8 -*-
"""Sonda: quanto custa PUBLICAR (gravar) a tblCQ_Final e a tblDB_Recebimento pelo VBA, com calculo manual e eventos
desligados -- inteira, so a cauda alterada, e crescendo. Tambem le as propriedades das conexoes (refresh ao abrir,
periodico). Roda numa COPIA; nada salvo.  Uso: python sonda_publicar.py <copia.xlsm>"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA  # noqa: E402
import instalar_integracao as ii  # noqa: E402

MOD = r'''
Option Explicit
Private Function Tab_(ByVal nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nome Then Set Tab_ = lo: Exit Function
        Next lo
    Next ws
End Function
Public Function Sonda(ByVal senha As String, ByVal nome As String) As String
    Dim lo As ListObject, n As Long, nc As Long, a As Variant, s As String, t As Double, c As Object, cauda As Variant
    Dim i As Long, j As Long, k As Long
    On Error GoTo falha
    Set lo = Tab_(nome)
    lo.Parent.Unprotect senha
    n = lo.ListRows.Count: nc = lo.ListColumns.Count
    a = lo.DataBodyRange.Value2
    For Each c In ThisWorkbook.Connections
        If InStr(1, c.Name, "Consulta", vbTextCompare) > 0 Then
            s = s & c.Name & ":open=" & c.RefreshWithRefreshAll & "/" & c.OLEDBConnection.RefreshOnFileOpen & "/per=" & c.OLEDBConnection.RefreshPeriod & ";"
        End If
    Next c
    s = s & "|n=" & n & "|calc_auto="
    t = Timer
    lo.DataBodyRange.Value2 = a
    s = s & Format$(Timer - t, "0.00")
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    t = Timer
    lo.DataBodyRange.Value2 = a
    s = s & "|inteira_manual=" & Format$(Timer - t, "0.00")
    k = 1000: If k > n Then k = n
    ReDim cauda(1 To k, 1 To nc)
    For i = 1 To k
        For j = 1 To nc: cauda(i, j) = a(n - k + i, j): Next j
    Next i
    t = Timer
    lo.DataBodyRange.Rows(n - k + 1).Resize(k, nc).Value2 = cauda
    s = s & "|cauda1000=" & Format$(Timer - t, "0.00")
    t = Timer
    lo.Resize lo.Range.Resize(n + 1 + 1200, nc)
    s = s & "|resize+1200=" & Format$(Timer - t, "0.00")
    t = Timer
    lo.DataBodyRange.Rows(n + 1).Resize(1200, nc).Value2 = lo.DataBodyRange.Rows(1).Resize(1200, nc).Value2
    s = s & "|grava1200=" & Format$(Timer - t, "0.00")
    t = Timer
    lo.DataBodyRange.Rows(n + 1).Resize(1200, nc).ClearContents
    lo.Resize lo.Range.Resize(n + 1, nc)
    s = s & "|encolhe=" & Format$(Timer - t, "0.00")
    t = Timer
    Application.Calculate
    s = s & "|calculate=" & Format$(Timer - t, "0.00")
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Sonda = s
    Exit Function
falha:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Sonda = s & "|ERRO " & Err.Number & " " & Err.Description
End Function
'''


def main(arq):
    q = QA('Bioquimica', arq)
    try:
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mSondaPub'
        comp.CodeModule.AddFromString(MOD)
        for t in ('tblCQ_Final', 'tblDB_Recebimento', 'tblCQ_Final'):
            print(t, q.run('mSondaPub.Sonda', ii.SENHA, t, teto=600), flush=True)
    finally:
        q.fechar(salvar=False)


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]))
