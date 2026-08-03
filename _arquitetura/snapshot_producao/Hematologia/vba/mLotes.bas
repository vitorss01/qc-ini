Attribute VB_Name = "mLotes"
Option Explicit
' CAMADA: Lotes — troca do lote ativo e persistencia das views por lote
' (Analitos/LotesStore, Liberacao/LiberStore, Registros/RegistrosStore).
Private Const CAP_ANALITOS As Long = 40
Private Const NLB_LIBER As Long = 200


Private Function BlocoDoLote(ByVal lote As String) As Long
    ' posicao (1-based) do lote no registro; 0 se nao encontrado. Compara como texto (robusto).
    Dim rng As Range, c As Range, k As Long
    BlocoDoLote = 0
    If Trim$(lote) = "" Then Exit Function
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    If rng Is Nothing Then Exit Function
    k = 0
    For Each c In rng
        k = k + 1
        If Trim$(CStr(c.Value)) = Trim$(lote) Then
            BlocoDoLote = k
            Exit Function
        End If
    Next c
End Function

Private Sub SalvarViewNoBloco(ByVal iBloco As Long, ByVal lote As String)
    ' grava a view atual (Analitos specs + Liberacao assinaturas) no bloco do lote.
    If iBloco < 1 Then Exit Sub
    On Error Resume Next
    Dim stA As Worksheet, r0 As Long, srcA As Range, i As Long
    Set stA = ThisWorkbook.Sheets("LotesStore")
    r0 = 2 + (iBloco - 1) * CAP_ANALITOS
    Set srcA = ThisWorkbook.Names("aInput").RefersToRange           ' E4:P43 (40 x 12)
    stA.Range(stA.Cells(r0, 3), stA.Cells(r0 + CAP_ANALITOS - 1, 14)).Value = srcA.Value
    For i = 0 To CAP_ANALITOS - 1
        stA.Cells(r0 + i, 1).Value = lote
        stA.Cells(r0 + i, 2).Value = i + 1
    Next i
    Dim stL As Worksheet, rb0 As Long, srcL As Range
    Set stL = ThisWorkbook.Sheets("LiberStore")
    rb0 = 2 + (iBloco - 1) * NLB_LIBER
    Set srcL = ThisWorkbook.Names("libView").RefersToRange          ' C4:F203 (200 x 4)
    stL.Range(stL.Cells(rb0, 1), stL.Cells(rb0 + NLB_LIBER - 1, 4)).Value = srcL.Value
    Dim stR As Worksheet, rr0 As Long, srcR As Range
    Set stR = ThisWorkbook.Sheets("RegistrosStore")
    rr0 = 2 + (iBloco - 1) * NLB_LIBER
    Set srcR = ThisWorkbook.Names("regView").RefersToRange          ' B4:M203 (200 x 12)
    stR.Range(stR.Cells(rr0, 1), stR.Cells(rr0 + NLB_LIBER - 1, 12)).Value = srcR.Value
End Sub

Private Sub CarregarBlocoNaView(ByVal iBloco As Long)
    ' carrega o bloco do lote na view. Specs: se o bloco estiver vazio, HERDA (mantem a view atual).
    ' Assinaturas de Liberacao: sempre carrega (limpa se o bloco estiver vazio) — nunca herda assinatura.
    If iBloco < 1 Then Exit Sub
    On Error Resume Next
    Dim stA As Worksheet, r0 As Long, dstA As Range
    Set stA = ThisWorkbook.Sheets("LotesStore")
    r0 = 2 + (iBloco - 1) * CAP_ANALITOS
    Set dstA = ThisWorkbook.Names("aInput").RefersToRange
    If Trim$(CStr(stA.Cells(r0, 1).Value)) <> "" Then
        dstA.Value = stA.Range(stA.Cells(r0, 3), stA.Cells(r0 + CAP_ANALITOS - 1, 14)).Value
    End If
    Dim stL As Worksheet, rb0 As Long, dstL As Range
    Set stL = ThisWorkbook.Sheets("LiberStore")
    rb0 = 2 + (iBloco - 1) * NLB_LIBER
    Set dstL = ThisWorkbook.Names("libView").RefersToRange
    dstL.Value = stL.Range(stL.Cells(rb0, 1), stL.Cells(rb0 + NLB_LIBER - 1, 4)).Value
    Dim stR As Worksheet, rr0 As Long, dstR As Range
    Set stR = ThisWorkbook.Sheets("RegistrosStore")
    rr0 = 2 + (iBloco - 1) * NLB_LIBER
    Set dstR = ThisWorkbook.Names("regView").RefersToRange
    dstR.Value = stR.Range(stR.Cells(rr0, 1), stR.Cells(rr0 + NLB_LIBER - 1, 12)).Value
End Sub

Public Sub TrocarLote()
    ' chamado pelo Worksheet_Change da Configuracao quando o seletor loteAtivo muda.
    Dim novo As String, atual As String, iAtual As Long, iNovo As Long
    On Error GoTo fim
    novo = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
    atual = Trim$(CStr(ThisWorkbook.Names("loteCarregado").RefersToRange.Value))
    If novo = atual Then GoTo fim
    If novo = "" Then GoTo fim
    iNovo = BlocoDoLote(novo)
    If iNovo = 0 Then
        MsgBox "O lote '" & novo & "' nao esta no registro (Configuracao). Cadastre-o antes de usar.", vbExclamation
        ThisWorkbook.Names("loteAtivo").RefersToRange.Value = atual
        GoTo fim
    End If
    Application.ScreenUpdating = False
    iAtual = BlocoDoLote(atual)
    If iAtual > 0 Then SalvarViewNoBloco iAtual, atual   ' persiste o lote que estava em uso
    CarregarBlocoNaView iNovo                            ' carrega o novo (herda specs se vazio)
    ThisWorkbook.Names("loteCarregado").RefersToRange.Value = novo
    SalvarViewNoBloco iNovo, novo                        ' carimba/persiste o novo bloco (cobre a heranca)
    Application.CalculateFull
    AtualizarEixos
    HookCharts
    Application.ScreenUpdating = True
fim:
End Sub

Public Sub FlushLoteAtual()
    ' persiste a view do lote em uso nos armazens (chamado no BeforeSave).
    On Error Resume Next
    Dim lote As String, i As Long
    lote = Trim$(CStr(ThisWorkbook.Names("loteCarregado").RefersToRange.Value))
    i = BlocoDoLote(lote)
    If i > 0 Then SalvarViewNoBloco i, lote
End Sub
