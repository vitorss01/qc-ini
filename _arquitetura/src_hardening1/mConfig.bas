Attribute VB_Name = "mConfig"
Option Explicit
' ===== VERSIONAMENTO DA TABELA DE ELEGIBILIDADE (item 2.5) =====
'
' O RISCO. `Cfg_Status` decide o que entra em media, DP, CV, Bias, Sigma e
' Westgard (ADR-006). Mudar UMA celula redefine RETROATIVAMENTE o que compos a
' estatistica de todo o historico -- cinco anos de dados passam a significar
' outra coisa, sem erro, sem aviso e sem rastro. Numa auditoria, a pergunta
' "com que criterio este Sigma foi calculado em marco?" nao tinha resposta.
'
' A CORRECAO, em tres partes:
'
'   1. VERSAO. `Cfg_Status!F1` guarda um inteiro que sobe a cada alteracao
'      efetiva. Nunca desce. E a identidade da configuracao vigente.
'
'   2. SOMBRA. As colunas H e I guardam a copia do par Status/Elegivel como
'      estava na ultima versao. Sem ela nao da para saber o valor ANTERIOR:
'      Worksheet_Change dispara DEPOIS da alteracao, quando o valor antigo ja
'      se perdeu.
'
'   3. LOG. Cada diferenca vira uma linha no Audit_Log, com valor antes e
'      depois, usuario e carimbo -- e a cadeia de hash cobre isso como cobre
'      qualquer outro evento.
'
' Ao final, `InvalidarCache` derruba o cache de elegibilidade do motor: sem
' isso a sessao corrente seguiria calculando com a regra antiga.
'
' O que isto NAO faz: nao impede a mudanca. Impedir seria errado -- a tabela
' existe justamente para ser extensivel sem tocar em codigo (ADR-006). O que
' se exige e que toda mudanca fique registrada e datada.

Public Const CFG As String = "Cfg_Status"
Public Const CFG_R0 As Long = 4            ' primeira linha de estado
Public Const CFG_RN As Long = 40           ' ultima linha varrida (espaco para novos estados)
Public Const CFG_C_STATUS As Long = 2      ' B
Public Const CFG_C_ELEG As Long = 3        ' C
Public Const CFG_C_SOMBRA_ST As Long = 8   ' H
Public Const CFG_C_SOMBRA_EL As Long = 9   ' I
Public Const CFG_L_VERSAO As Long = 1
Public Const CFG_C_VERSAO As Long = 6      ' F1

Public Function VersaoCfg() As Long
    Dim ws As Worksheet, v As Variant
    Set ws = ThisWorkbook.Sheets(CFG)
    v = ws.Cells(CFG_L_VERSAO, CFG_C_VERSAO).Value
    If IsNumeric(v) Then VersaoCfg = CLng(v) Else VersaoCfg = 0
End Function

' Grava a sombra a partir do estado atual. Nao registra nada: e o marco zero.
Public Sub SincronizarSombraCfg()
    Dim ws As Worksheet, i As Long, prot As Boolean
    Set ws = ThisWorkbook.Sheets(CFG)
    ' ADR-046: a sombra e o que detecta mudanca retroativa de elegibilidade.
    ' Uma excecao no meio deixando a Configuracao destrancada abriria a porta
    ' para editar justamente o que a sombra existe para vigiar.
    On Error GoTo restauraCfg
    prot = LiberarEscrita(ws)
    For i = CFG_R0 To CFG_RN
        ws.Cells(i, CFG_C_SOMBRA_ST).Value = ws.Cells(i, CFG_C_STATUS).Value
        ws.Cells(i, CFG_C_SOMBRA_EL).Value = ws.Cells(i, CFG_C_ELEG).Value
    Next i
    If VersaoCfg() = 0 Then ws.Cells(CFG_L_VERSAO, CFG_C_VERSAO).Value = 1
    ws.Columns(CFG_C_SOMBRA_ST).Hidden = True
    ws.Columns(CFG_C_SOMBRA_EL).Hidden = True

restauraCfg:
    Dim nErrC As Long, sErrC As String
    nErrC = Err.Number: sErrC = Err.Description
    RestaurarProtecao ws, prot
    On Error GoTo 0
    If nErrC <> 0 Then Err.Raise nErrC, "mConfig.SincronizarSombraCfg", sErrC
End Sub

' Compara o estado atual com a sombra, registra cada diferenca e sobe a versao.
' Devolve o numero de alteracoes registradas.
Public Function AuditarMudancaCfg() As Long
    Dim ws As Worksheet, i As Long, n As Long, prot As Boolean
    Dim stAtual As String, elAtual As String, stAnt As String, elAnt As String
    Dim versaoNova As Long

    Set ws = ThisWorkbook.Sheets(CFG)
    versaoNova = VersaoCfg() + 1

    For i = CFG_R0 To CFG_RN
        stAtual = Trim$(CStr(ws.Cells(i, CFG_C_STATUS).Value))
        elAtual = Trim$(CStr(ws.Cells(i, CFG_C_ELEG).Value))
        stAnt = Trim$(CStr(ws.Cells(i, CFG_C_SOMBRA_ST).Value))
        elAnt = Trim$(CStr(ws.Cells(i, CFG_C_SOMBRA_EL).Value))

        If stAtual <> stAnt Or elAtual <> elAnt Then
            Dim acao As String
            If stAnt = "" And stAtual <> "" Then
                acao = "CFG_ESTADO_INCLUIDO"
            ElseIf stAtual = "" And stAnt <> "" Then
                acao = "CFG_ESTADO_REMOVIDO"
            Else
                acao = "CFG_ELEGIBILIDADE_ALTERADA"
            End If

            Auditar CAT_CONFIG, acao, "mConfig", _
                    versaoNova, Date, "", "", 0, "", _
                    elAnt, elAtual, _
                    stAnt & "=" & elAnt, stAtual & "=" & elAtual, _
                    "Alteracao de elegibilidade", _
                    "Estado """ & IIf(stAtual <> "", stAtual, stAnt) & """ na linha " & i & _
                    ". Muda RETROATIVAMENTE o que compoe media, DP, CV, Bias, Sigma e Westgard de todo o historico."
            n = n + 1
        End If
    Next i

    If n > 0 Then
        On Error GoTo restauraAud
        prot = LiberarEscrita(ws)
        ws.Cells(CFG_L_VERSAO, CFG_C_VERSAO).Value = versaoNova
        For i = CFG_R0 To CFG_RN
            ws.Cells(i, CFG_C_SOMBRA_ST).Value = ws.Cells(i, CFG_C_STATUS).Value
            ws.Cells(i, CFG_C_SOMBRA_EL).Value = ws.Cells(i, CFG_C_ELEG).Value
        Next i
        RestaurarProtecao ws, prot
        On Error GoTo 0

        ' A sessao corrente ainda tem a regra antiga em memoria.
        InvalidarCache
    End If

    AuditarMudancaCfg = n
    Exit Function

restauraAud:
    ' A sombra fica INCOMPLETA se o erro caiu no meio do laco, e isso e
    ' proposital: sombra parcial faz a proxima execucao acusar diferenca e
    ' auditar de novo. Reprotege e devolve o erro a quem chamou -- o que nao
    ' pode acontecer e a aba ficar aberta e o erro sumir.
    Dim nErrA As Long, sErrA As String
    nErrA = Err.Number: sErrA = Err.Description
    RestaurarProtecao ws, prot
    On Error GoTo 0
    Err.Raise nErrA, "mConfig.AuditarMudancaCfg", sErrA
End Function
