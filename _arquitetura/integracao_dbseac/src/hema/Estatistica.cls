' Aba Estatistica (Hematologia) -- ADR-088 (OC-0003, auditoria ciclo 01 / G13).
'   B3 (Ano De) e D3 (Ano Ate) sao filtros que o VBA le para gravar n/Media/DP/CV (C:F, pelo Eng_Saida engEstat).
'   Mudar um deles refaz o bloco na hora: antes so o ATUALIZAR refazia, e a aba mostrava os anos novos com os
'   numeros dos anos anteriores. B4 (lote) e formula (= lote em analise): quem refaz a troca de lote e o
'   mLotes.TrocarLoteAnalise.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim nE As Long, sE As String
    If Intersect(Target, Me.Range("B3,D3")) Is Nothing Then Exit Sub
    If mApp.Ocupado() Then Exit Sub
    mApp.InicioUsuario "Recalculando a Estatística..."
    On Error GoTo fim
    mEstatistica.AtualizarEstatisticaAba
fim:
    nE = Err.Number: sE = Err.Description
    mApp.fim
    If nE <> 0 Then MsgBox "Não foi possível recalcular a Estatística: " & sE, vbExclamation, "Estatística"
End Sub
