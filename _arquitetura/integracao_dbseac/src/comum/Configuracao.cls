Option Explicit
' Aba Configuracao. C20 (loteAtivo) -> troca o lote EM USO (ADR-049).
' D26:D125 (validade do cadastro de lotes) -> confere e registra na trilha (ADR-060).
' C26:C125 (codigo do lote) -> o lote novo ganha o seu bloco na aba Analitos (ADR-074).
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Dim sel As Range
    Set sel = ThisWorkbook.Names("loteAtivo").RefersToRange
    If Not sel Is Nothing Then
        If Not Intersect(Target, sel) Is Nothing Then
            Application.EnableEvents = False
            TrocarLote
            Application.EnableEvents = True
        End If
    End If
    If Not Intersect(Target, Me.Range("D26:D125")) Is Nothing Then
        mLotes.ValidadeEditada Target
    End If
    ' ADR-074: lote digitado direto no cadastro ganha o seu bloco na aba Analitos
    If Not Intersect(Target, Me.Range("C26:C125")) Is Nothing Then
        mLotesBlocos.SincronizarBlocos
    End If
End Sub
