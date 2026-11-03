' Aba Estatistica (Bioquimica). Visao (B4), Ano (D4), Lote (H4) e o periodo personalizado (B5/D5) refazem a
' Estatistica. Codigo que ja estava no arquivo, versionado no ADR-088 (sem mudanca de comportamento).
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("B4,D4,H4,B5,D5")) Is Nothing Then
        Application.EnableEvents = False
        mEstatistica.AtualizarEstatistica
        Application.EnableEvents = True
    End If
End Sub
