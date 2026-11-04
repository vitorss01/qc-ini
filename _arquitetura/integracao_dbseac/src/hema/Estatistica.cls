' Aba Estatistica (Hematologia) -- ADR-088 (OC-0003, auditoria ciclo 01 / G13).
'   B3 (Ano De) e D3 (Ano Ate) sao filtros que o VBA le para gravar n/Media/DP/CV (C:F, pelo Eng_Saida engEstat).
'   Mudar um deles refaz o bloco na hora: antes so o ATUALIZAR refazia, e a aba mostrava os anos novos com os
'   numeros dos anos anteriores. B4 (lote) e formula (= lote em analise): quem refaz a troca de lote e o
'   mLotes.TrocarLoteAnalise.
'   ADR-089 (OC-0032): R4 (eqProvedor, o provedor do CEQ de TODOS os analitos) trocado -> PROVEDOR_ALTERADO no
'   Audit_Log com o antes (guardado quando a celula e selecionada) e o depois.
Option Explicit

Private mProvAntes As String

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("R4")) Is Nothing Then mProvAntes = CStr(Me.Range("R4").Value)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim nE As Long, sE As String
    If Not Intersect(Target, Me.Range("R4")) Is Nothing Then
        On Error Resume Next
        mCEQForm.ProvedorEQAlterado mProvAntes, CStr(Me.Range("R4").Value), "Estatística R4 (vale para todos os analitos)"
        mProvAntes = CStr(Me.Range("R4").Value)
        On Error GoTo 0
    End If
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
