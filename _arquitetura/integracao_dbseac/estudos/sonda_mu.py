# -*- coding: utf-8 -*-
"""Perfil da incerteza/Sigma no ATUALIZAR (mIncerteza.RecalcularIncerteza ~3 s): quantas formulas, quanto custa a
1a agregacao (passada na tblCQ_Final) e quanto custa cada chamada ja agregada. Roda numa COPIA; nada salvo.
Uso: python sonda_mu.py <Bioquimica|Hematologia> <copia.xlsm>"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
from qa_final import QA  # noqa: E402

MOD = r'''
Option Explicit
Private Function Conta(ByVal nome As String, ByRef nUdf As Long) As Long
    Dim c As Range
    On Error GoTo sai
    For Each c In ThisWorkbook.Names(nome).RefersToRange.Cells
        If c.HasFormula Then
            Conta = Conta + 1
            If InStr(1, c.Formula, "IncertezaCIQ", vbTextCompare) > 0 Then nUdf = nUdf + 1
        End If
    Next c
sai:
End Function
Public Function PerfilMU() As String
    Dim s As String, t As Double, v As Variant, a As String, ini As Variant, fim As Variant, ex As Range, i As Long, n As Long, nU As Long
    Dim nm As Variant
    For Each nm In Array("MU_Faixa", "Sigma_Faixa", "Sigma_CVp")
        nU = 0
        n = Conta(CStr(nm), nU)
        s = s & nm & ":formulas=" & n & ",com_IncertezaCIQ=" & nU & ";"
    Next nm
    a = CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value)
    ini = ThisWorkbook.Names("MU_Ini").RefersToRange.Value
    fim = ThisWorkbook.Names("MU_Fim").RefersToRange.Value
    Set ex = ThisWorkbook.Names("Estat_Exclusoes").RefersToRange
    mDados.IniciarSnapshot
    v = mDados.CarregarDB()
    mIncerteza.InvalidarIncerteza
    t = Timer: v = mIncerteza.IncertezaCIQ(a, 1, "CV", ini, fim, ex): s = s & "1a_chamada(agrega)=" & Format$(Timer - t, "0.000") & ";"
    t = Timer
    For i = 1 To 500: v = mIncerteza.IncertezaCIQ(a, 1, "CV", ini, fim, ex): Next i
    s = s & "500_chamadas_quentes=" & Format$(Timer - t, "0.000") & ";"
    t = Timer
    For i = 1 To 500: v = mIncerteza.IncertezaCIQ(a, 1, "CV", ini, fim): Next i
    s = s & "500_sem_exclusoes=" & Format$(Timer - t, "0.000") & ";"
    Application.Calculation = xlCalculationManual
    mIncerteza.InvalidarIncerteza
    For Each nm In Array("MU_Faixa", "Sigma_Faixa", "Sigma_CVp")
        t = Timer
        ThisWorkbook.Names(nm).RefersToRange.Dirty
        Application.Calculate
        s = s & "calc_" & nm & "=" & Format$(Timer - t, "0.000") & ";"
    Next nm
    ' por coluna da MU_Faixa: formula da 1a linha e tempo de recalculo da coluna (dados ja agregados)
    Dim col As Range, f As String
    For Each col In ThisWorkbook.Names("MU_Faixa").RefersToRange.Columns
        f = col.Cells(1, 1).Formula
        t = Timer
        col.Dirty
        col.Calculate
        s = s & vbLf & col.Column & " " & Format$(Timer - t, "0.000") & "s " & Left$(f, 110)
    Next col
    mDados.EncerrarSnapshot
    Application.Calculation = xlCalculationAutomatic
    PerfilMU = s
End Function
'''


def main(produto, arq):
    q = QA(produto, arq)
    try:
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mSondaMU'
        comp.CodeModule.AddFromString(MOD)
        print(q.run('mSondaMU.PerfilMU', teto=600), flush=True)
    finally:
        q.fechar(salvar=False)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]))
