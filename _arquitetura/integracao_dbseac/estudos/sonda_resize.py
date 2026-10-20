# -*- coding: utf-8 -*-
"""Sonda: uma tabela ligada a consulta (QueryTable do Power Query) aceita ser redimensionada e regravada pelo VBA?
Roda numa COPIA; nada salvo.  Uso: python sonda_resize.py <copia.xlsm>"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA  # noqa: E402
import instalar_integracao as ii  # noqa: E402

MOD = r'''
Option Explicit
Public Function Sonda(ByVal senha As String) As String
    Dim lo As ListObject, n As Long, nc As Long, a As Variant, s As String, t As Double
    Set lo = ThisWorkbook.Worksheets("DB").ListObjects("tblDB_Recebimento")
    lo.Parent.Unprotect senha
    n = lo.ListRows.Count: nc = lo.ListColumns.Count
    a = lo.DataBodyRange.Value2
    s = "tem_querytable=" & (Not lo.QueryTable Is Nothing)
    On Error Resume Next
    t = Timer
    lo.Resize lo.Range.Resize(n + 1 + 10, nc)        ' cabecalho + n + 10 linhas
    s = s & "|crescer=" & IIf(Err.Number = 0, "OK", "ERRO " & Err.Number & " " & Err.Description) & "(" & lo.ListRows.Count & ")"
    Err.Clear
    lo.Resize lo.Range.Resize(n + 1 - 5, nc)
    s = s & "|encolher=" & IIf(Err.Number = 0, "OK", "ERRO " & Err.Number & " " & Err.Description) & "(" & lo.ListRows.Count & ")"
    Err.Clear
    lo.Resize lo.Range.Resize(n + 1, nc)
    lo.DataBodyRange.Value2 = a
    s = s & "|regravar=" & IIf(Err.Number = 0, "OK", "ERRO " & Err.Number & " " & Err.Description) & "|" & Format$(Timer - t, "0.00") & "s"
    Err.Clear
    lo.QueryTable.Refresh False
    s = s & "|refresh_pq_depois=" & IIf(Err.Number = 0, "OK", "ERRO " & Err.Number & " " & Err.Description) & "(" & lo.ListRows.Count & ")"
    Sonda = s
End Function
'''


def main(arq):
    q = QA('Bioquimica', arq)
    try:
        comp = q.wb.VBProject.VBComponents.Add(1)
        comp.Name = 'mSondaResize'
        comp.CodeModule.AddFromString(MOD)
        print(q.run('mSondaResize.Sonda', ii.SENHA, teto=600), flush=True)
    finally:
        q.fechar(salvar=False)


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]))
