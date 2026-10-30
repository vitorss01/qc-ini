' Registros -- ADR-084: o ID do resultado digitado na coluna K registra a calibracao (mCalibracao). As demais colunas
' continuam de digitacao livre, como antes.
' ADR-085: duplo clique na coluna I (Assinatura) assina a linha com login e senha -- a mesma assinatura da Liberacao,
' presa ao ID da coluna K (mCalibracao.AssinarRegistro).
' ADR-087: a aba e a Tabela tblRegistros (todos os lotes, coluna Lote), que cresce sozinha (mCalibracao.RegistrosMudou
' deixa sempre uma linha vazia no fim); a linha 4 tem os filtros da vista (Lote, Analito, De/Ate, Equipamento,
' Calibrado?, Status), aplicados ao mudar.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = mCalibracao.TabelaRegistros()
    If lo Is Nothing Then
        If Not Intersect(Target, Me.Range("K4:K203")) Is Nothing Then mCalibracao.IdDigitado Target
        Exit Sub
    End If
    If Not Intersect(Target, lo.ListColumns(11).DataBodyRange) Is Nothing Then mCalibracao.IdDigitado Target
    mCalibracao.RegistrosMudou Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim lo As ListObject, col As Range
    On Error Resume Next
    Set lo = mCalibracao.TabelaRegistros()
    If lo Is Nothing Then
        Set col = Me.Range("I4:I203")
    Else
        Set col = lo.ListColumns(9).DataBodyRange
    End If
    If Intersect(Target, col) Is Nothing Then Exit Sub
    Cancel = True                                  ' nada de editar a celula: a assinatura vem do login conferido
    mCalibracao.AssinarRegistro Target.Cells(1, 1)
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    mCalibracao.AtualizarListasFiltro               ' ADR-087: lotes, analitos, equipamentos e status dos dados reais
End Sub
