Attribute VB_Name = "mImportar"
Option Explicit
' ===== IMPORTACAO POR ABA (substitui o frmMassa) =====
'
' No lugar do formulario de colagem, uma ABA de entrada: cabecalho
'   Data | Nivel | Lote | <analito 1> | <analito 2> | ...
' horizontal, como o analista pensa. Ele cola os dados (mesmo desnormalizados,
' e so um registro inicial) e clica em Registrar. Os dados MIGRAM para o
' DB_Resultados pela mesma logica do frmMassa e somem da aba.
'
' REUSO, NAO REESCRITA. A validacao, o mapeamento de RUN e o Upsert sao os
' mesmos do frmMassa: ObterOuCriarRUN por (Data + nucleo do lote),
' UpsertResultados, AtualizarOperacao. So muda a ORIGEM -- de um TextBox para
' uma faixa de celulas.
'
' TUDO-OU-NADA (decisao do gestor). Se qualquer linha tiver erro, NADA e
' gravado: a aba lista as linhas com problema e o usuario corrige. O
' DB_Resultados nunca recebe importacao pela metade.
'
' Layout da aba (constantes abaixo): cabecalho na linha IMP_CAB, dados a partir
' de IMP_R0, area de erro a direita dos analitos.

'' "Imp" e PALAVRA RESERVADA do VBA (operador logico, como And/Or/Xor/Eqv).
' Uma constante chamada IMP impede o modulo de ser parseado e derruba o
' PROJETO INTEIRO -- mesma classe do identificador aS. Dai o prefixo.
Public Const ABA_IMP As String = "Importar"
Public Const IMP_CAB As Long = 3           ' linha do cabecalho
Public Const IMP_R0 As Long = 4            ' primeira linha de dado
Public Const IMP_RN As Long = 203          ' ultima linha de dado (200 corridas)
Public Const IMP_C_DATA As Long = 2        ' B
Public Const IMP_C_NIVEL As Long = 3       ' C
Public Const IMP_C_LOTE As Long = 4        ' D
Public Const IMP_C_AN0 As Long = 5         ' E: primeiro analito

' Numero de analitos cadastrados (largura da area de dados).
Private Function NAnalitos() As Long
    NAnalitos = ListaAnalitos().Count
End Function

' Ultima linha COM algum dado na faixa de importacao.
Public Function UltimaLinhaImport() As Long
    Dim ws As Worksheet, i As Long, ult As Long, c As Long, nA As Long
    Set ws = ThisWorkbook.Sheets(ABA_IMP)
    nA = NAnalitos()
    For i = IMP_R0 To IMP_RN
        Dim temAlgo As Boolean: temAlgo = False
        For c = IMP_C_DATA To IMP_C_AN0 + nA - 1
            If Trim$(CStr(ws.Cells(i, c).Value)) <> "" Then temAlgo = True: Exit For
        Next c
        If temAlgo Then ult = i
    Next i
    If ult < IMP_R0 Then ult = IMP_R0 - 1
    UltimaLinhaImport = ult
End Function

' Ponto de entrada do botao Registrar. So interface: chama o nucleo e traduz o
' resultado em MsgBox.
Public Sub RegistrarImportacao()
    Dim r As String, p() As String
    r = ExecutarImportacao(False)
    p = Split(r, "|")
    Select Case p(0)
        Case "VAZIO"
            MsgBox "Nada para registrar - a area de importacao esta vazia.", vbInformation, "Importar"
        Case "SEM_ANALITO"
            MsgBox "Nenhum analito cadastrado.", vbExclamation, "Importar"
        Case "ERRO"
            MsgBox p(1) & " inconsistencia(s). NADA foi gravado." & vbLf & vbLf & _
                   "Corrija as linhas apontadas a direita e clique em Registrar de novo.", _
                   vbExclamation, "Importar"
        Case "CANCELADO"
            ' o usuario desistiu na confirmacao; nada a dizer
        Case "OK"
            MsgBox "Importacao concluida." & vbLf & vbLf & _
                   "Novos: " & p(2) & "   Atualizados: " & p(3), vbInformation, "Importar"
    End Select
End Sub

' Nucleo da importacao. TODA a regra vive aqui; RegistrarImportacao so mostra
' mensagem. Assim o teste automatizado exercita o MESMO codigo que o botao
' executa -- e nao uma copia que poderia divergir dele.
'
' silencioso = True pula a confirmacao (unico MsgBox do caminho de gravacao),
' o que permite rodar sem interface. Nao muda mais nada: valida igual, grava
' igual, limpa igual.
'
' Retorno: "VAZIO" | "SEM_ANALITO" | "CANCELADO"
'          "ERRO|<qtd>"
'          "OK|<gravados>|<novos>|<atualizados>"
Public Function ExecutarImportacao(ByVal silencioso As Boolean) As String
    Dim ws As Worksheet, nA As Long, ult As Long, i As Long, j As Long
    Dim erros As Collection, regs() As Variant, nReg As Long, mx As Long
    Dim analNomes() As String, cA As Collection

    Set ws = ThisWorkbook.Sheets(ABA_IMP)
    Set cA = ListaAnalitos()
    nA = cA.Count
    If nA = 0 Then ExecutarImportacao = "SEM_ANALITO": Exit Function
    ReDim analNomes(1 To nA)
    For j = 1 To nA: analNomes(j) = cA(j): Next j

    ult = UltimaLinhaImport()
    If ult < IMP_R0 Then ExecutarImportacao = "VAZIO": Exit Function

    ' -------- validacao TUDO-OU-NADA, acumulando erros --------
    Set erros = New Collection
    mx = (ult - IMP_R0 + 1) * nA
    ReDim regs(1 To mx, 1 To 7)
    nReg = 0

    ' Chave (data|nivel|nucleo do lote) ja vista nesta colagem. Duas linhas para
    ' a MESMA corrida e nivel nao sao um upsert legitimo: sao erro de digitacao.
    ' O Upsert obedeceria e a segunda linha sobrescreveria a primeira em
    ' silencio -- perda de dado sem aviso. Aqui isso vira inconsistencia.
    Dim vistas As Object
    Set vistas = CreateObject("Scripting.Dictionary")

    For i = IMP_R0 To ult
        Dim dt As Date, ok As Boolean, lvl As Long, lote As String, temLinha As Boolean
        temLinha = (Trim$(CStr(ws.Cells(i, IMP_C_DATA).Value)) <> "" Or _
                    Trim$(CStr(ws.Cells(i, IMP_C_LOTE).Value)) <> "")
        If Not temLinha Then GoTo proxima

        ' Data
        dt = ParseData(CStr(ws.Cells(i, IMP_C_DATA).Value), ok)
        If Not ok Then erros.Add "Linha " & i & " - Data invalida: """ & CStr(ws.Cells(i, IMP_C_DATA).Value) & """"

        ' Nivel
        Dim sN As String: sN = Trim$(CStr(ws.Cells(i, IMP_C_NIVEL).Value))
        If Not IsNumeric(sN) Then
            erros.Add "Linha " & i & " - Nivel invalido: """ & sN & """"
            lvl = 0
        Else
            lvl = CLng(sN)
            If lvl < 1 Or lvl > NLV Then erros.Add "Linha " & i & " - Nivel fora de 1.." & NLV & ": " & lvl
        End If

        ' Lote
        lote = Trim$(CStr(ws.Cells(i, IMP_C_LOTE).Value))
        If lote = "" Then
            erros.Add "Linha " & i & " - Lote nao preenchido"
        ElseIf Not LoteExisteImp(lote) Then
            erros.Add "Linha " & i & " - Lote nao cadastrado: """ & lote & """"
        End If

        ' Corrida/nivel repetido na mesma colagem
        Dim cabOK As Boolean, chave As String
        cabOK = (ok And lvl >= 1 And lvl <= NLV And lote <> "" And LoteExisteImp(lote))
        If cabOK Then
            chave = CStr(CLng(dt)) & "|" & lvl & "|" & NucleoLote(CodigoLote(lote, lvl))
            If vistas.Exists(chave) Then
                erros.Add "Linha " & i & " - mesma corrida e nivel ja aparecem na linha " & _
                          vistas(chave) & ". Junte as duas numa linha so."
                cabOK = False
            Else
                vistas.Add chave, i
            End If
        End If

        ' Resultados dos analitos.
        ' O laco roda SEMPRE, mesmo com o cabecalho invalido: se so gravasse
        ' quando data/nivel/lote estao certos, um valor nao numerico ficaria
        ' escondido ate o usuario consertar a data e clicar de novo -- e o
        ' tudo-ou-nada deixaria de mostrar tudo de uma vez.
        Dim run As Long
        If cabOK Then run = PreverRUN(dt, lote)
        For j = 1 To nA
            Dim s As String: s = Trim$(CStr(ws.Cells(i, IMP_C_AN0 + j - 1).Value))
            If s <> "" Then
                Dim vOK As Boolean, v As Double
                v = ParseNum(s, vOK)
                If Not vOK Then
                    erros.Add "Linha " & i & " - " & analNomes(j) & " nao numerico: """ & s & """"
                ElseIf cabOK Then
                    nReg = nReg + 1
                    regs(nReg, 1) = run
                    regs(nReg, 2) = dt
                    regs(nReg, 3) = lvl
                    regs(nReg, 4) = CodigoLote(lote, lvl)
                    regs(nReg, 5) = analNomes(j)
                    regs(nReg, 6) = v
                    regs(nReg, 7) = ST_ATIVO
                End If
            End If
        Next j
proxima:
    Next i

    ' -------- se houve QUALQUER erro, nada e gravado --------
    MostrarErros ws, nA, erros
    If erros.Count > 0 Then
        ExecutarImportacao = "ERRO|" & erros.Count
        Exit Function
    End If
    If nReg = 0 Then ExecutarImportacao = "VAZIO": Exit Function

    If Not silencioso Then
        If MsgBox("Registrar " & nReg & " resultado(s) no banco?" & vbLf & vbLf & _
                  "Eles migram para o DB_Resultados e somem desta aba. Registros da mesma " & _
                  "corrida e analito sao atualizados, nao duplicados.", _
                  vbYesNo + vbQuestion, "Confirmar") <> vbYes Then
            ExecutarImportacao = "CANCELADO"
            Exit Function
        End If
    End If

    ' -------- RUN definitivo por (Data + nucleo do lote) --------
    ' Igual ao frmMassa: durante a validacao o RUN e provisorio (nenhuma data
    ' nova existe no banco ainda); aqui a corrida passa a existir e recebe o
    ' numero definitivo.
    Dim mapaRun As Object, kRun As String, iR As Long
    Set mapaRun = CreateObject("Scripting.Dictionary")
    For iR = 1 To nReg
        kRun = CStr(CLng(CDate(regs(iR, 2)))) & "|" & NucleoLote(CStr(regs(iR, 4)))
        If Not mapaRun.Exists(kRun) Then
            mapaRun.Add kRun, ObterOuCriarRUN(CDate(regs(iR, 2)), NucleoLote(CStr(regs(iR, 4))))
        End If
        regs(iR, 1) = mapaRun(kRun)
    Next iR

    Dim final_() As Variant, c As Long
    ReDim final_(1 To nReg, 1 To 7)
    For i = 1 To nReg
        For c = 1 To 7: final_(i, c) = regs(i, c): Next c
    Next i

    Dim res As String
    res = UpsertResultados(final_)
    RegistrarLog "IMPORTACAO_ABA", nReg & " resultados"

    ' -------- limpa a aba e volta ao Resultados --------
    LimparAreaImport ws, nA
    AtualizarOperacao
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets("Resultados")
    If wsR.Visible = xlSheetVisible Then wsR.Activate

    ExecutarImportacao = "OK|" & nReg & "|" & Split(res, "|")(0) & "|" & Split(res, "|")(1)
End Function

Private Function LoteExisteImp(ByVal lote As String) As Boolean
    Dim c As Collection, i As Long
    Set c = ListaLotes()
    For i = 1 To c.Count
        If Trim$(CStr(c(i))) = lote Then LoteExisteImp = True: Exit Function
    Next i
End Function

' Lista de erros numa coluna a direita dos analitos.
Private Sub MostrarErros(ByVal ws As Worksheet, ByVal nA As Long, ByVal erros As Collection)
    Dim cErr As Long, i As Long, prot As Boolean
    cErr = IMP_C_AN0 + nA + 1
    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:="qcini2025"
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    ws.Cells(IMP_CAB, cErr).Value = "Inconsistencias"
    ws.Cells(IMP_CAB, cErr).Font.Bold = True
    If erros.Count = 0 Then
        ws.Cells(IMP_R0, cErr).Value = "(nenhuma)"
    Else
        For i = 1 To erros.Count
            ws.Cells(IMP_R0 + i - 1, cErr).Value = erros(i)
        Next i
    End If
    ws.Columns(cErr).ColumnWidth = 52
    If prot Then ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

Private Sub LimparAreaImport(ByVal ws As Worksheet, ByVal nA As Long)
    Dim prot As Boolean, cErr As Long
    cErr = IMP_C_AN0 + nA + 1
    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:="qcini2025"
    ws.Range(ws.Cells(IMP_R0, IMP_C_DATA), ws.Cells(IMP_RN, IMP_C_AN0 + nA - 1)).ClearContents
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    If prot Then ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

' Navegacao: o botao da aba Resultados leva para ca.
Public Sub IrParaImportar()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ABA_IMP)
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
    ws.Cells(IMP_R0, IMP_C_DATA).Select
End Sub
