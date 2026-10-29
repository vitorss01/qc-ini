' Registros -- ADR-084: o ID do resultado digitado na coluna K registra a calibracao (mCalibracao). As demais colunas
' continuam de digitacao livre, como antes. A aba e a vista do lote EM USO (o mLotes guarda cada lote na RegistrosStore).
' ADR-085: duplo clique na coluna I (Assinatura) assina a linha com login e senha -- a mesma assinatura da Liberacao,
' presa ao ID da coluna K (mCalibracao.AssinarRegistro).
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Intersect(Target, Me.Range("K4:K203")) Is Nothing Then Exit Sub
    mCalibracao.IdDigitado Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    If Intersect(Target, Me.Range("I4:I203")) Is Nothing Then Exit Sub
    Cancel = True                                  ' nada de editar a celula: a assinatura vem do login conferido
    mCalibracao.AssinarRegistro Target.Cells(1, 1)
End Sub
