Attribute VB_Name = "mApp"
Option Explicit
' ============================================================================
'  CASCA DA APLICACAO (ADR-050)
'
'  Toda operacao que o usuario dispara -- trocar analito, trocar lote, gravar
'  parametro, importar corridas -- passa por Inicio/Fim:
'
'    tela congelada ........ nada pisca enquanto o motor escreve
'    calculo em manual ..... cada bloco gravado NAO dispara um recalculo; um so
'                            no fim, e so do que ficou sujo
'    eventos desligados .... escrita do motor nao aciona Worksheet_Change
'    cursor de espera e mensagem na barra de status
'
'  Aninhavel: so a chamada mais externa restaura o estado anterior. Quem chama
'  Inicio TEM de chamar Fim tambem no caminho de erro (padrao abaixo); Reset
'  existe para Workbook_Open e para o botao de socorro, caso algo tenha ficado
'  preso por uma excecao.
'
'      mApp.Inicio "Carregando..."
'      On Error GoTo falha
'      ...
'  falha:
'      mApp.Fim
'      If Err.Number <> 0 Then ...
' ============================================================================
Private mNivel As Long
Private mCalc As Long
Private mTela As Boolean
Private mEvt As Boolean

Public Sub Inicio(Optional ByVal msg As String = "")
    If mNivel = 0 Then
        mCalc = Application.Calculation
        mTela = Application.ScreenUpdating
        mEvt = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        If mCalc <> xlCalculationManual Then Application.Calculation = xlCalculationManual
        On Error Resume Next
        Application.Cursor = xlWait
        On Error GoTo 0
    End If
    mNivel = mNivel + 1
    If Len(msg) > 0 Then Application.StatusBar = msg
End Sub

' Entrada de operacao disparada pelo USUARIO (clique, digitacao). Se o contador
' estiver preso de uma operacao anterior que morreu no meio -- por exemplo, o
' usuario apertou "Finalizar" numa caixa de erro do VBA --, o Excel ficaria em
' calculo manual, sem eventos e com a tela congelada, e o clique seguinte pareceria
' "dar erro" e mostrar a tela errada. Clique so acontece com o VBA parado: se ha
' contador aberto aqui, ele e resto. Zera e comeca limpo (ADR-053).
Public Sub InicioUsuario(Optional ByVal msg As String = "")
    If mNivel > 0 Then Reset
    Inicio msg
End Sub

Public Sub Fim()
    If mNivel <= 0 Then mNivel = 0: Exit Sub
    mNivel = mNivel - 1
    If mNivel > 0 Then Exit Sub
    On Error Resume Next
    ' voltar ao automatico recalcula o que ficou sujo -- uma vez so
    If Application.Calculation <> mCalc Then Application.Calculation = mCalc
    Application.EnableEvents = mEvt
    Application.ScreenUpdating = mTela
    Application.Cursor = xlDefault
    Application.StatusBar = False
End Sub

' Estado limpo: calculo automatico, eventos e tela ligados. Para Workbook_Open
' e para o caso de uma excecao ter pulado o Fim.
Public Sub Reset()
    mNivel = 0
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
End Sub

Public Function Ocupado() As Boolean
    Ocupado = (mNivel > 0)
End Function

' ============================================================================
'  NAVEGACAO E ATALHOS DE TELA (ADR-051)
' ============================================================================

' Painel!C3: o usuario escolheu o analito pelo NOME na lista suspensa. A
' celula volta a ser a formula (nome da posicao B3) e B3 recebe a posicao --
' o spinner continua valendo. Uma escolha substitui ate 39 cliques no spinner.
Public Sub AnalitoEscolhido(ByVal cel As Range)
    Dim nome As String, idx As Variant, ws As Worksheet, prot As Boolean
    Set ws = cel.Worksheet
    nome = Trim$(CStr(cel.Value))
    Application.EnableEvents = False
    On Error Resume Next
    prot = LiberarEscrita(ws)
    idx = Application.Match(nome, ThisWorkbook.Sheets("Analitos").Range("A4:A43"), 0)
    cel.Formula = "=IFERROR(INDEX(Analitos!$A$4:$A$43,$B$3),"""")"
    If Not IsError(idx) And Len(nome) > 0 Then ws.Range("B3").Value = CLng(idx)
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
    On Error GoTo 0
    mEstatistica.PainelMudou
End Sub

' ---------------------------------------------------------------------------
' Barra de navegacao: cada botao chama uma destas. O botao e desenhado pelo
' instalador (ux.py) em todas as telas; o nome do botao e "nav_<destino>".
' ---------------------------------------------------------------------------
Public Sub IrInicio(): Ir "Início", "A1": End Sub
Public Sub IrPainel(): Ir "Painel", "A1": End Sub
Public Sub IrAnalitos(): Ir "Analitos", "A1": End Sub
Public Sub IrLotes(): Ir "Configuração", "B18": End Sub
Public Sub IrEstatistica(): Ir "Estatística", "A1": End Sub
Public Sub IrLiberacao(): Ir "Liberação", "A1": End Sub
Public Sub IrRegistros(): Ir "Registros", "A1": End Sub
Public Sub IrEventos(): Ir "Eventos_Westgard", "A1": End Sub
Public Sub IrResultados(): Ir "Resultados", "A1": End Sub
Public Sub IrEQA(): Ir "EQA.CAP_Dados", "A1": End Sub

' Lancar resultados: Bioquimica tem a aba Importar; Hematologia, o formulario
' da corrida (frmCorrida) sobre a aba Resultados.
Public Sub IrLancar()
    If ExisteAba("Importar") Then
        Ir "Importar", "A1"
    Else
        Ir "Resultados", "A1"
        On Error Resume Next
        Application.Run "AbrirFormCorrida"
    End If
End Sub

Public Sub Ir(ByVal aba As String, Optional ByVal cel As String = "A1")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(aba)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then
        MsgBox "A tela '" & aba & "' nao esta disponivel para o seu perfil.", vbInformation, "QC"
        Exit Sub
    End If
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range(cel).Select
End Sub

Private Function ExisteAba(ByVal aba As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(aba)
    ExisteAba = Not ws Is Nothing
End Function

' Painel -> Analitos na linha do analito em tela: editar media/DP DESTE lote.
Public Sub IrParametrosDoAnalito()
    Dim an As String, idx As Variant
    an = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    idx = Application.Match(an, ThisWorkbook.Sheets("Analitos").Range("A4:A43"), 0)
    If IsError(idx) Then Ir "Analitos", "E4": Exit Sub
    Ir "Analitos", "E" & (3 + CLng(idx))
    Application.StatusBar = "Editando media/DP do lote " & mLotes.LotePainel() & _
                            " -- a gravacao e imediata e fica no historico de auditoria."
End Sub

' Analitos -> Painel, na celula do lote em analise (a lista de lotes).
Public Sub IrLotePainel()
    Dim c As Range
    On Error Resume Next
    Set c = ThisWorkbook.Names("loteSel").RefersToRange
    Ir "Painel", "A1"
    If Not c Is Nothing Then c.Select
    Application.StatusBar = "Escolha o lote na lista da celula selecionada."
End Sub

' ---------------------------------------------------------------------------
' NOVO LOTE -- cadastro guiado: codigo, validade, usar agora?, media/DP.
' ---------------------------------------------------------------------------
Public Sub NovoLote()
    Dim cod As String, v As String, r As String, resp As VbMsgBoxResult
    cod = Trim$(InputBox("Codigo do NOVO lote de controle" & vbLf & vbLf & _
                         "Somente o nucleo do lote -- sem o prefixo 'QC-' e sem os 2 digitos " & _
                         "finais de nivel (ex.: 8975).", "Novo lote -- 1 de 3"))
    If cod = "" Then Exit Sub
    v = Trim$(InputBox("Validade do lote " & cod & " (dd/mm/aaaa) -- pode deixar em branco.", _
                       "Novo lote -- 2 de 3"))
    r = CadastrarLote(cod, v)
    If Left$(r, 2) <> "OK" Then
        MsgBox Mid$(r, InStr(r, "|") + 1), vbExclamation, "Novo lote"
        Exit Sub
    End If
    resp = MsgBox("Lote " & cod & " cadastrado." & vbLf & vbLf & _
                  "Passar a LANCAR os resultados novos neste lote agora?" & vbLf & _
                  "(Sim = lote em uso passa a ser " & cod & ")", vbYesNo + vbQuestion, "Novo lote -- 3 de 3")
    If resp = vbYes Then UsarLote cod
    If MsgBox("Cadastrar agora a MEDIA e o DP do lote " & cod & " na aba Analitos?" & vbLf & vbLf & _
              "Lote novo nasce sem media/DP: o grafico e o Westgard so passam a avaliar " & _
              "depois que eles forem digitados.", vbYesNo + vbQuestion, "Novo lote") = vbYes Then
        AnalisarLote cod
        Ir "Analitos", "E4"
        Application.StatusBar = "Digite Media/DP de cada nivel do lote " & cod & " -- grava na hora."
    End If
End Sub

' Nucleo do cadastro, sem dialogo (testavel). Devolve "OK|linha" ou "ERRO|motivo".
Public Function CadastrarLote(ByVal cod As String, ByVal validade As String) As String
    Dim ws As Worksheet, r As Long, livre As Long, dtv As Variant, prot As Boolean
    cod = Trim$(cod)
    If cod = "" Then CadastrarLote = "ERRO|Codigo vazio.": Exit Function
    Set ws = ThisWorkbook.Sheets("Configuração")
    For r = 26 To 125
        If StrComp(Trim$(CStr(ws.Cells(r, 3).Value)), cod, vbTextCompare) = 0 Then
            CadastrarLote = "ERRO|O lote " & cod & " ja esta cadastrado (linha " & r & ").": Exit Function
        End If
        If livre = 0 And Trim$(CStr(ws.Cells(r, 3).Value)) = "" Then livre = r
    Next r
    If livre = 0 Then CadastrarLote = "ERRO|O cadastro ja tem 100 lotes.": Exit Function
    If Trim$(validade) <> "" Then
        If Not IsDate(validade) Then CadastrarLote = "ERRO|Data invalida: " & validade & ". O lote NAO foi cadastrado.": Exit Function
        dtv = CDate(validade)
    End If
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    ws.Cells(livre, 3).NumberFormat = "@"
    ws.Cells(livre, 3).Value = cod
    If Not IsEmpty(dtv) Then ws.Cells(livre, 4).Value = dtv
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
    On Error Resume Next
    mAuditoria.RegistrarLog "LOTE_CADASTRADO", "Lote " & cod & IIf(IsEmpty(dtv), "", " validade " & Format(dtv, "dd/mm/yyyy"))
    CadastrarLote = "OK|" & livre
End Function

' Lote EM USO (lancamentos) passa a ser este -- o mesmo caminho da Configuracao!C20.
Public Sub UsarLote(ByVal cod As String)
    Dim ws As Worksheet, prot As Boolean
    Set ws = ThisWorkbook.Sheets("Configuração")
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    ThisWorkbook.Names("loteAtivo").RefersToRange.NumberFormat = "@"
    ThisWorkbook.Names("loteAtivo").RefersToRange.Value = cod
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
    mLotes.TrocarLote
End Sub

' Lote EM ANALISE (Painel) passa a ser este -- o mesmo caminho da lista do Painel.
Public Sub AnalisarLote(ByVal cod As String)
    Dim ws As Worksheet, prot As Boolean
    Set ws = ThisWorkbook.Names("loteSel").RefersToRange.Worksheet
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    ThisWorkbook.Names("loteSel").RefersToRange.Value = "'" & cod
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
    mLotes.TrocarLoteAnalise
End Sub

' Esconde/mostra as guias de planilha junto com o "visual de sistema".
Public Sub GuiasVisiveis(ByVal mostrar As Boolean)
    Dim w As Window
    On Error Resume Next
    For Each w In ThisWorkbook.Windows
        w.DisplayWorkbookTabs = mostrar
    Next w
End Sub

Public Sub AlternarGuias()
    Dim w As Window, m As Boolean
    m = Not ActiveWindow.DisplayWorkbookTabs
    For Each w In ThisWorkbook.Windows
        w.DisplayWorkbookTabs = m
    Next w
End Sub

' ---------------------------------------------------------------------------
' TELA ATUAL -- guardar e devolver
'
' Com as guias ocultas (visual de aplicativo), terminar uma operacao noutra aba
' deixa o usuario perdido. Toda rotina de tela guarda a aba antes e devolve
' depois, inclusive no caminho de erro.
Public Function TelaAtual() As Object
    On Error Resume Next
    Set TelaAtual = ActiveSheet
End Function

Public Sub VoltarPara(ByVal ws As Object)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If ws.Parent Is ThisWorkbook Then
        If ws.Visible = xlSheetVisible Then
            If Not ActiveSheet Is ws Then ws.Activate
        End If
    End If
End Sub
' ---------------------------------------------------------------------------
' PERIODO DO PAINEL -- Ano + Periodo sao ATALHO para preencher De/Ate (ADR-054)
'
' O Calc filtra por filtroDe/filtroAte, e so por eles. Havia um segundo
' caminho -- quatro caixas de trimestre (qsel1..4) lidas pelo motor -- e com
' dois caminhos o usuario conseguia pedir coisas que se anulam (2o trimestre
' marcado na caixa e T1 escolhido na lista): o Painel ficava vazio sem dizer
' por que. Agora ha um caminho so.
'
' ANO AUSENTE NAO ESCREVE PERIODO NENHUM.
'
' Antes, ano vazio virava 0 e DateSerial(0, 1, 1) devolvia 01/01/2000 -- a
' regra de ano de dois digitos do VBA le 0 como 2000. O Painel filtrava um
' periodo sem dado e ficava vazio sem dizer por que. Numero plausivel e errado
' e pior que erro visivel: ninguem confere o que parece certo.
Public Function AplicarFiltroAnoMes() As String
    Dim a As Long, s As String, ev As Boolean, v As Variant
    Dim visao As String, parte As Long
    ev = Application.EnableEvents
    On Error GoTo falha
    Application.EnableEvents = False
    v = ThisWorkbook.Names("filtroAno").RefersToRange.Value
    s = Trim$(CStr(ThisWorkbook.Names("filtroPeriodo").RefersToRange.Value))
    ' IsNumeric(Empty) e VERDADEIRO em VBA: celula vazia passava por "ano valido",
    ' virava 0 e a rotina saia calada -- o usuario escolhia T1 e nada acontecia.
    a = 0
    If IsNumeric(v) Then a = CLng(Val(CStr(v)))
    If a < 1900 Or a > 2200 Then
        ' Sem ano nao da para dizer "T1" de coisa nenhuma. Assume o ano da ULTIMA
        ' corrida e ESCREVE na tela, para ficar visivel de que ano se fala.
        a = AnoDosDados()
        If a < 1900 Then AplicarFiltroAnoMes = "ERRO|0|sem ano e sem corrida no banco": GoTo fim
        ThisWorkbook.Names("filtroAno").RefersToRange.Value = a
    End If
    mPeriodo.VisaoDoRotulo s, visao, parte
    If visao = "TUDO" Then
        ' sem janela: o Painel mostra o historico inteiro, de todos os anos
        ThisWorkbook.Names("filtroDe").RefersToRange.ClearContents
        ThisWorkbook.Names("filtroAte").RefersToRange.ClearContents
        AplicarFiltroAnoMes = "OK|" & a & "|TUDO|0"
        GoTo fim
    End If
    ThisWorkbook.Names("filtroDe").RefersToRange.Value = mPeriodo.PeriodoInicio(visao, a, parte)
    ThisWorkbook.Names("filtroAte").RefersToRange.Value = mPeriodo.PeriodoFim(visao, a, parte)
    AplicarFiltroAnoMes = "OK|" & a & "|" & visao & "|" & parte
fim:
    Application.EnableEvents = ev
    Exit Function
falha:
    ' ADR-053: falha vira mensagem, nao silencio.
    AplicarFiltroAnoMes = "ERRO|" & Err.Number & "|" & Err.Description
    Application.EnableEvents = ev
End Function

Private Function AnoDosDados() As Long
    Dim mx As Variant
    On Error Resume Next
    mx = Application.WorksheetFunction.Max(ThisWorkbook.Names("rData").RefersToRange)
    If IsNumeric(mx) Then
        If mx > 0 Then AnoDosDados = Year(CDate(mx))
    End If
    If AnoDosDados = 0 Then AnoDosDados = Year(Date)
End Function
