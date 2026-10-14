Option Explicit
' Aba Analitos -- a central de referencia, especificacoes e LIMITES POR LOTE.
' E4:J43 (Media/DP do lote em edicao) -> gravados NA HORA na LotesStore, com trilha (ADR-049)
' e com o retrato da especificacao vigente (ADR-073).
' loteEdicao (seletor ao lado do titulo) -> escolher aqui = escolher no Painel (H3) (ADR-073).
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim sel As Range
    On Error Resume Next
    Set sel = ThisWorkbook.Names("loteEdicao").RefersToRange
    On Error GoTo 0
    If Not sel Is Nothing Then
        If Not Intersect(Target, sel) Is Nothing Then
            mLotes.LoteEdicaoMudou
            Exit Sub
        End If
    End If
    If Intersect(Target, Me.Range("E4:J43")) Is Nothing Then Exit Sub
    mLotes.ParametroEditado Target
End Sub
