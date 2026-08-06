Attribute VB_Name = "mImportar"
Option Explicit
' ===== IMPORTACAO POR ABA (substitui o frmMassa) =====
'
' Versao para o arquivo de PRODUCAO (Fase 2). Usa a API que existe nele:
' NovoRUN, UpsertResultados, CodigoLote, ParseData, ParseNum, ListaAnalitos,
' ListaLotes, RegistrarLog, AtualizarOperacao. Nao depende de ObterOuCriarRUN
' nem de NLV, que so existem na linha do hardening.
'
' O analista cola os dados na horizontal -- uma corrida por linha, os analitos
' lado a lado -- e clica em Registrar. Os dados MIGRAM para o DB_Resultados e
' somem da aba.
'
' A SIGLA QUE O USUARIO VE NAO E O NOME QUE O BANCO GUARDA.
'
' O cabecalho visivel traz siglas de bancada (GLI, URE, TGO...). O banco guarda
' o nome cadastrado na aba Analitos ("Glicose", "Ureia", "AST (TGO)"), porque e
' por esse nome que o Calc, o Painel, a Estatistica e os graficos encontram o
' resultado. Gravar a sigla quebraria todos eles em silencio.
'
' Por isso existe a LINHA DE MAPEAMENTO (IMP_MAP), oculta: cada coluna guarda
' ali o nome cadastrado para onde ela vai. O mapeamento e DADO, escrito pelo
' script de montagem a partir da propria aba Analitos -- nao e codigo. Assim
' nao ha nome de analito acentuado dentro do modulo (fonte .bas e cp1252, e
' acento em fonte VBA e origem classica de corrupcao silenciosa), e trocar o
' cabecalho nao exige recompilar nada.
'
' Coluna sem nome mapeado e coluna de analito NAO CADASTRADO no produto. Ela
' aparece no cabecalho -- a ordem pedida e respeitada -- mas dado colado nela e
' recusado, em vez de virar linha orfa no banco.
'
' TUDO-OU-NADA. Se qualquer linha tiver erro, NADA e gravado.

Public Const ABA_IMP As String = "Importar"
Public Const IMP_MAP As Long = 3           ' linha OCULTA: nome cadastrado por coluna
Public Const IMP_CAB As Long = 4           ' linha do cabecalho visivel (siglas)
Public Const IMP_R0 As Long = 5            ' primeira linha de dado
Public Const IMP_RN As Long = 204          ' ultima linha de dado (200 corridas)
Public Const IMP_C_DATA As Long = 2        ' B
Public Const IMP_C_NIVEL As Long = 3       ' C
Public Const IMP_C_LOTE As Long = 4        ' D
Public Const IMP_C_AN0 As Long = 5         ' E: primeira coluna de analito

' Niveis de controle do produto. A producao da Bioquimica fixa 1..2 (o frmMassa
' fazia o mesmo); nao existe constante NLV neste arquivo.
Private Const IMP_NLV As Long = 2

Private Const IMP_SENHA As String = "qcini2025"

' Quantas colunas de analito o cabecalho tem (varre ate achar celula vazia).
Private Function NColunasAnalito(ByVal ws As Worksheet) As Long
    Dim c As Long, n As Long
    c = IMP_C_AN0
    Do While Trim$(CStr(ws.Cells(IMP_CAB, c).Value)) <> ""
        n = n + 1
        c = c + 1
        If n > 200 Then Exit Do
    Loop
    NColunasAnalito = n
End Function

' Le a area de colagem INTEIRA de uma vez.
'
' Ler celula a celula custa uma travessia COM por celula: 200 linhas x 34
' colunas sao 6.800 travessias so para descobrir onde os dados terminam. Em
' bloco e UMA. Todo o resto da rotina passa a trabalhar sobre esta matriz.
Private Function LerArea(ByVal ws As Worksheet, ByVal nC As Long) As Variant
    LerArea = ws.Range(ws.Cells(IMP_R0, IMP_C_DATA), _
                       ws.Cells(IMP_RN, IMP_C_AN0 + nC - 1)).Value
End Function

' Texto de uma celula da matriz da area (linha e coluna em coordenadas da aba).
Private Function Cel(ByRef m As Variant, ByVal linha As Long, ByVal col As Long) As String
    Cel = Trim$(CStr(m(linha - IMP_R0 + 1, col - IMP_C_DATA + 1)))
End Function

' Ultima linha COM algum dado na faixa de importacao.
Public Function UltimaLinhaImport() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ABA_IMP)
    Dim nC As Long: nC = NColunasAnalito(ws)
    UltimaLinhaImport = UltimaLinhaDaArea(LerArea(ws, nC), nC)
End Function

Private Function UltimaLinhaDaArea(ByRef m As Variant, ByVal nC As Long) As Long
    Dim i As Long, c As Long, ult As Long
    For i = IMP_R0 To IMP_RN
        For c = IMP_C_DATA To IMP_C_AN0 + nC - 1
            If Cel(m, i, c) <> "" Then ult = i: Exit For
        Next c
    Next i
    If ult < IMP_R0 Then ult = IMP_R0 - 1
    UltimaLinhaDaArea = ult
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
            MsgBox "Nenhum analito cadastrado na aba Analitos.", vbExclamation, "Importar"
        Case "ERRO"
            MsgBox p(1) & " inconsistencia(s). NADA foi gravado." & vbLf & vbLf & _
                   "Corrija as linhas apontadas a direita e clique em Registrar de novo.", _
                   vbExclamation, "Importar"
        Case "CANCELADO"
            ' o usuario desistiu na confirmacao
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
' o que permite rodar sem interface.
'
' Retorno: "VAZIO" | "SEM_ANALITO" | "CANCELADO"
'          "ERRO|<qtd>"
'          "OK|<gravados>|<novos>|<atualizados>"
Public Function ExecutarImportacao(ByVal silencioso As Boolean) As String
    Dim ws As Worksheet, nC As Long, ult As Long, i As Long, j As Long
    Dim erros As Collection, regs() As Variant, nReg As Long, mx As Long
    Dim destino() As String, rotulo() As String

    Set ws = ThisWorkbook.Sheets(ABA_IMP)
    nC = NColunasAnalito(ws)
    If nC = 0 Then ExecutarImportacao = "SEM_ANALITO": Exit Function

    ' Le o cabecalho visivel (sigla) e a linha oculta de mapeamento (nome
    ' cadastrado). Coluna com destino vazio nao existe no produto.
    ReDim destino(1 To nC)
    ReDim rotulo(1 To nC)
    For j = 1 To nC
        rotulo(j) = Trim$(CStr(ws.Cells(IMP_CAB, IMP_C_AN0 + j - 1).Value))
        destino(j) = Trim$(CStr(ws.Cells(IMP_MAP, IMP_C_AN0 + j - 1).Value))
    Next j

    Dim area As Variant
    area = LerArea(ws, nC)
    ult = UltimaLinhaDaArea(area, nC)
    If ult < IMP_R0 Then ExecutarImportacao = "VAZIO": Exit Function

    ' Lotes cadastrados em um dicionario, UMA vez. Antes, LoteExisteImp
    ' percorria a lista de lotes a cada linha -- e era chamado duas vezes por
    ' linha, relendo as mesmas celulas da Configuracao o tempo todo.
    Dim lotesOK As Object, cl As Collection, iL As Long
    Set lotesOK = CreateObject("Scripting.Dictionary")
    Set cl = ListaLotes()
    For iL = 1 To cl.Count
        If Not lotesOK.Exists(Trim$(CStr(cl(iL)))) Then lotesOK.Add Trim$(CStr(cl(iL))), True
    Next iL

    ' -------- validacao TUDO-OU-NADA, acumulando erros --------
    Set erros = New Collection
    mx = (ult - IMP_R0 + 1) * nC
    ReDim regs(1 To mx, 1 To 7)
    nReg = 0

    ' Chave (data|nivel) ja vista nesta colagem. Duas linhas para a MESMA
    ' corrida e nivel nao sao upsert legitimo: sao erro de digitacao. O Upsert
    ' obedeceria e a segunda sobrescreveria a primeira em silencio.
    Dim vistas As Object
    Set vistas = CreateObject("Scripting.Dictionary")

    For i = IMP_R0 To ult
        Dim dt As Date, ok As Boolean, lvl As Long, lote As String, temLinha As Boolean
        temLinha = (Cel(area, i, IMP_C_DATA) <> "" Or Cel(area, i, IMP_C_LOTE) <> "")
        If Not temLinha Then GoTo proxima

        ' Data
        dt = ParseData(Cel(area, i, IMP_C_DATA), ok)
        If Not ok Then erros.Add "Linha " & i & " - Data invalida: """ & Cel(area, i, IMP_C_DATA) & """"

        ' Nivel
        Dim sN As String: sN = Cel(area, i, IMP_C_NIVEL)
        If Not IsNumeric(sN) Then
            erros.Add "Linha " & i & " - Nivel invalido: """ & sN & """"
            lvl = 0
        Else
            lvl = CLng(sN)
            If lvl < 1 Or lvl > IMP_NLV Then erros.Add "Linha " & i & " - Nivel fora de 1.." & IMP_NLV & ": " & lvl
        End If

        ' Lote
        lote = Cel(area, i, IMP_C_LOTE)
        If lote = "" Then
            erros.Add "Linha " & i & " - Lote nao preenchido"
        ElseIf Not lotesOK.Exists(lote) Then
            erros.Add "Linha " & i & " - Lote nao cadastrado: """ & lote & """"
        End If

        ' Corrida/nivel repetido na mesma colagem
        Dim cabOK As Boolean, chave As String
        cabOK = (ok And lvl >= 1 And lvl <= IMP_NLV And lote <> "" And lotesOK.Exists(lote))
        If cabOK Then
            chave = CStr(CLng(dt)) & "|" & lvl & "|" & lote
            If vistas.Exists(chave) Then
                erros.Add "Linha " & i & " - mesma data, nivel e lote ja aparecem na linha " & _
                          vistas(chave) & ". Junte as duas numa linha so."
                cabOK = False
            Else
                vistas.Add chave, i
            End If
        End If

        ' Resultados dos analitos.
        ' O laco roda SEMPRE, mesmo com o cabecalho da linha invalido: se so
        ' rodasse quando data/nivel/lote estao certos, um valor nao numerico
        ' ficaria escondido ate o usuario consertar a data e clicar de novo --
        ' e o tudo-ou-nada deixaria de mostrar tudo de uma vez.
        ' RUN provisorio: 0. O numero definitivo so e atribuido depois da
        ' confirmacao, em AtribuirRUNs.
        '
        ' NAO chamar NovoRUN aqui. Na producao ele devolve "maior RUN do banco
        ' + 1" quando a data ainda nao existe -- e durante a validacao NADA foi
        ' gravado, entao TODAS as datas novas recebem o MESMO numero. Colar
        ' cinco dias de uma vez fundia os cinco numa corrida so. Defeito
        ' herdado do frmMassa, invisivel enquanto se importava um dia por vez.
        Dim run As Long
        run = 0
        For j = 1 To nC
            Dim s As String: s = Cel(area, i, IMP_C_AN0 + j - 1)
            If s <> "" Then
                If destino(j) = "" Then
                    ' Coluna existe no cabecalho mas o analito nao esta na aba
                    ' Analitos. Gravar isso criaria resultado orfao, que nenhum
                    ' calculo encontraria.
                    erros.Add "Linha " & i & " - " & rotulo(j) & " nao esta cadastrado na aba Analitos"
                Else
                    Dim vOK As Boolean, v As Double
                    v = ParseNum(s, vOK)
                    If Not vOK Then
                        erros.Add "Linha " & i & " - " & rotulo(j) & " nao numerico: """ & s & """"
                    ElseIf cabOK Then
                        nReg = nReg + 1
                        regs(nReg, 1) = run
                        regs(nReg, 2) = dt
                        regs(nReg, 3) = lvl
                        regs(nReg, 4) = CodigoLote(lote, lvl)
                        regs(nReg, 5) = destino(j)
                        regs(nReg, 6) = v
                        regs(nReg, 7) = ST_ATIVO
                    End If
                End If
            End If
        Next j
proxima:
    Next i

    ' -------- se houve QUALQUER erro, nada e gravado --------
    MostrarErros ws, nC, erros
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

    AtribuirRUNs regs, nReg

    Dim final_() As Variant, c As Long
    ReDim final_(1 To nReg, 1 To 7)
    For i = 1 To nReg
        For c = 1 To 7: final_(i, c) = regs(i, c): Next c
    Next i

    Dim res As String
    res = UpsertResultados(final_)
    RegistrarLog "IMPORTACAO_ABA", nReg & " resultados"

    ' -------- limpa a aba e volta ao Resultados --------
    LimparAreaImport ws, nC
    AtualizarOperacao

    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets("Resultados")
    If wsR.Visible = xlSheetVisible Then wsR.Activate

    ExecutarImportacao = "OK|" & nReg & "|" & Split(res, "|")(0) & "|" & Split(res, "|")(1)
End Function

' Da a cada CORRIDA distinta (data + nucleo do lote) o seu proprio RUN.
'
' Le o banco UMA vez: corrida que ja existe reaproveita o numero, corrida nova
' recebe o proximo, e o proximo avanca a cada corrida nova do lote. E o que
' impede cinco datas coladas juntas de virarem uma corrida so.
Private Sub AtribuirRUNs(ByRef regs As Variant, ByVal nReg As Long)
    Dim dados As Variant, i As Long, mx As Long, k As String
    Dim mapa As Object, novo As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    Set novo = CreateObject("Scripting.Dictionary")

    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 And IsNumeric(dados(i, COL_RUN)) Then
                If CLng(dados(i, COL_RUN)) > mx Then mx = CLng(dados(i, COL_RUN))
                If IsDate(dados(i, COL_DATA)) Then
                    k = CStr(CLng(CDate(dados(i, COL_DATA)))) & "|" & NucleoLote(CStr(dados(i, COL_LOTE)))
                    If Not mapa.Exists(k) Then mapa.Add k, CLng(dados(i, COL_RUN))
                End If
            End If
        Next i
    End If

    For i = 1 To nReg
        k = CStr(CLng(CDate(regs(i, 2)))) & "|" & NucleoLote(CStr(regs(i, 4)))
        If mapa.Exists(k) Then
            regs(i, 1) = mapa(k)
        ElseIf novo.Exists(k) Then
            regs(i, 1) = novo(k)
        Else
            mx = mx + 1
            novo.Add k, mx
            regs(i, 1) = mx
        End If
    Next i
End Sub

Private Function LoteExisteImp(ByVal lote As String) As Boolean
    Dim c As Collection, i As Long
    Set c = ListaLotes()
    For i = 1 To c.Count
        If Trim$(CStr(c(i))) = lote Then LoteExisteImp = True: Exit Function
    Next i
End Function

' Lista de erros numa coluna a direita dos analitos.
Private Sub MostrarErros(ByVal ws As Worksheet, ByVal nC As Long, ByVal erros As Collection)
    Dim cErr As Long, i As Long, prot As Boolean
    cErr = IMP_C_AN0 + nC + 1
    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:=IMP_SENHA
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
    ws.Columns(cErr).ColumnWidth = 58
    If prot Then ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

Private Sub LimparAreaImport(ByVal ws As Worksheet, ByVal nC As Long)
    Dim prot As Boolean, cErr As Long
    cErr = IMP_C_AN0 + nC + 1
    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:=IMP_SENHA
    ws.Range(ws.Cells(IMP_R0, IMP_C_DATA), ws.Cells(IMP_RN, IMP_C_AN0 + nC - 1)).ClearContents
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    If prot Then ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
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
