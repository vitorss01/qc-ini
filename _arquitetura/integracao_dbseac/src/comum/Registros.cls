' Registros -- ADR-084: o ID do resultado digitado na coluna K registra a calibracao (mCalibracao). As demais colunas
' continuam de digitacao livre, como antes. A aba e a vista do lote EM USO (o mLotes guarda cada lote na RegistrosStore).
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Intersect(Target, Me.Range("K4:K203")) Is Nothing Then Exit Sub
    mCalibracao.IdDigitado Target
End Sub
