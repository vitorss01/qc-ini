Attribute VB_Name = "mLotes"
Option Explicit
' CAMADA: Lotes
'
' ============================================================================
'  ETAPA 1 -- MULTI-LOTE (ADR-049)
'
'  TRES PAPEIS DE LOTE, TRES NOMES. Tratar dois deles como um so foi a origem
'  dos defeitos que esta versao fecha.
'
'    loteAtivo  (Configuracao!C20) .. lote EM USO. Decide so o lote do
'                                    resultado NOVO (Importar, frmCorrida) e
'                                    as views operacionais Registros/Liberacao.
'    loteSel    (Painel!H3) ......... lote EM ANALISE, escolhido no Painel.
'                                    Chave de tudo que INTERPRETA: Calc,
'                                    grafico, motor, Eventos_Westgard,
'                                    Estatistica. O nome loteAnalise e o
'                                    loteSel efetivo (vazio -> loteAtivo).
'    loteParam  (Configuracao!G2) ... lote cujos parametros estao na tela
'                                    Analitos!E:J (o editor). Segue o loteSel.
'
'  Escolher um lote antigo para ANALISE nao pode mudar o lote em que o proximo
'  resultado sera gravado. Por isso loteSel nao e loteAtivo.
'
'  FONTE UNICA DOS PARAMETROS: LotesStore, uma linha por (lote, analito),
'  achada pela CHAVE da coluna P ("LOTE|ANALITO"). Nunca pela posicao do lote
'  no cadastro, nunca pela ordem dos analitos. Motor, Estatistica, BI e as
'  formulas do Calc leem todos por esta chave.
'
'  LOTE SEM PARAMETRO NAO HERDA. A versao anterior copiava a media/DP do lote
'  que estava na tela para o lote novo "para nao redigitar 40 analitos". Media
'  e DP sao propriedades do MATERIAL: herdadas, viram um grafico plausivel com
'  limites de outro lote. Agora o lote novo nasce vazio, e quem pergunta por
'  ele recebe False e nao interpreta.
' ============================================================================
Private Const CAP_ANALITOS As Long = 40
Private Const NLB_LIBER As Long = 200
Private Const LS_ABA As String = "LotesStore"
Private Const LS_R0 As Long = 2
Private Const LS_CAPLINHAS As Long = 4000     ' 100 lotes x 40 analitos (nomes ls*)
Private Const LS_C_LOTE As Long = 1           ' A
Private Const LS_C_IDX As Long = 2            ' B  posicao do analito (legado, informativo)
Private Const LS_C_MED1 As Long = 3           ' C..H = Med/DP N1, N2, N3
Private Const LS_C_NOME As Long = 15          ' O  analito
Private Const LS_C_CHAVE As Long = 16         ' P  LOTE|ANALITO
Private Const AN_ABA As String = "Analitos"
Private Const AN_R0 As Long = 4
Private Const AN_C_MED1 As Long = 5           ' Analitos!E

Private mLS As Variant                        ' snapshot LotesStore A..P
Private mLSIdx As Object                      ' chave -> linha do snapshot
Private mLSok As Boolean

' ============================ CHAVE E LOTE EM ANALISE ============================
Public Function ChaveLote(ByVal lote As String, ByVal analito As String) As String
    ChaveLote = UCase$(Trim$(lote)) & "|" & UCase$(Trim$(analito))
End Function

' Lote em analise no Painel. Vazio cai no lote em uso -- a mesma regra do nome
' loteAnalise, para formula e VBA nunca discordarem.
Public Function LotePainel() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(ThisWorkbook.Names("loteSel").RefersToRange.Value))
    On Error GoTo 0
    If s = "" Then s = LoteAtivoCore()
    LotePainel = LoteCanonico(s)
End Function

' Codigo do lote COMO CADASTRADO. Numa celula que nao e texto, o Excel faz de
' "010" o numero 10 -- e o nucleo gravado no banco continua "010". Casar pelo
' texto e, na falta, pelo valor devolve o codigo verdadeiro.
Public Function LoteCanonico(ByVal v As Variant) As String
    Dim s As String, rng As Range, c As Range, t As String, cand As String
    s = Trim$(CStr(v))
    LoteCanonico = s
    If s = "" Then Exit Function
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    For Each c In rng.Cells
        t = Trim$(CStr(c.Value))
        If t <> "" Then
            If t = s Then LoteCanonico = t: Exit Function
            If cand = "" And IsNumeric(s) And IsNumeric(t) Then
                If CDbl(s) = CDbl(t) Then cand = t
            End If
        End If
    Next c
    If cand <> "" Then LoteCanonico = cand
End Function

' Grava o lote no seletor do Painel como TEXTO (o apostrofo impede o Excel de
' converter "010" em 10). O formato da celula acrescenta "Lote " na tela.
Private Sub GravarLoteSel(ByVal lote As String)
    ThisWorkbook.Names("loteSel").RefersToRange.Value = "'" & lote
End Sub

Private Function LoteParamTela() As String
    On Error Resume Next
    LoteParamTela = Trim$(CStr(ThisWorkbook.Names("loteParam").RefersToRange.Value))
End Function

' ============================ LEITURA (fonte unica) ============================
Public Sub InvalidarLotes()
    mLSok = False
    mLS = Empty
    Set mLSIdx = Nothing
End Sub

Private Sub GarantirLotes()
    Dim ws As Worksheet, ult As Long, i As Long, k As String
    If mLSok Then Exit Sub
    Set ws = ThisWorkbook.Sheets(LS_ABA)
    Set mLSIdx = CreateObject("Scripting.Dictionary")
    mLSIdx.CompareMode = 1
    ult = ws.Cells(ws.rows.Count, LS_C_CHAVE).End(xlUp).Row
    If ult >= LS_R0 Then
        ' +1 linha: com uma so linha o .Value viria escalar, nao matriz
        mLS = ws.Range(ws.Cells(LS_R0, 1), ws.Cells(ult + 1, LS_C_CHAVE)).Value
        For i = 1 To UBound(mLS, 1)
            k = UCase$(Trim$(CStr(mLS(i, LS_C_CHAVE))))
            If k <> "" Then
                ' duplicata nao sobrescreve: vale a primeira, e ConferirLotesStore acusa
                If Not mLSIdx.Exists(k) Then mLSIdx.Add k, i
            End If
        Next i
    Else
        mLS = Empty
    End If
    mLSok = True
End Sub

Private Function NumeroValido(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbString, vbBoolean: Exit Function   ' "5" digitado como texto nao e parametro
    End Select
    NumeroValido = IsNumeric(v)
End Function

' Media e DP de UM lote para (analito, nivel). False = nao ha parametro
' utilizavel (ausente, texto, DP <= 0): quem recebe False NAO interpreta.
Public Function ParametrosLote(ByVal lote As String, ByVal analito As String, _
                               ByVal nivel As Long, ByRef media As Double, _
                               ByRef dp As Double) As Boolean
    Dim k As String, r As Long, c As Long, vM As Variant, vD As Variant
    media = 0: dp = 0
    If Trim$(lote) = "" Or Trim$(analito) = "" Then Exit Function
    If nivel < 1 Or nivel > 3 Then Exit Function
    GarantirLotes
    k = ChaveLote(lote, analito)
    If Not mLSIdx.Exists(k) Then Exit Function
    r = mLSIdx(k)
    c = LS_C_MED1 + (nivel - 1) * 2
    vM = mLS(r, c): vD = mLS(r, c + 1)
    If Not NumeroValido(vM) Or Not NumeroValido(vD) Then Exit Function
    If CDbl(vD) <= 0 Then Exit Function
    media = CDbl(vM): dp = CDbl(vD)
    ParametrosLote = True
End Function

' Atalhos para teste e planilha: devolvem "" quando nao ha parametro valido.
Public Function MediaDoLote(ByVal lote As String, ByVal analito As String, ByVal nivel As Long) As Variant
    Dim m As Double, s As Double
    If ParametrosLote(lote, analito, nivel, m, s) Then MediaDoLote = m Else MediaDoLote = ""
End Function

Public Function DPDoLote(ByVal lote As String, ByVal analito As String, ByVal nivel As Long) As Variant
    Dim m As Double, s As Double
    If ParametrosLote(lote, analito, nivel, m, s) Then DPDoLote = s Else DPDoLote = ""
End Function

' ============================ ESCRITA NA STORE ============================
' Indice chave -> linha da PLANILHA, e a ultima linha usada (pela coluna A).
Private Function IndiceStore(ByVal ws As Worksheet, ByRef ult As Long) As Object
    Dim d As Object, v As Variant, i As Long, k As String, ultP As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ult = ws.Cells(ws.rows.Count, LS_C_LOTE).End(xlUp).Row
    ultP = ws.Cells(ws.rows.Count, LS_C_CHAVE).End(xlUp).Row
    If ultP > ult Then ult = ultP
    If ult < LS_R0 Then ult = LS_R0 - 1
    If ult >= LS_R0 Then
        v = ws.Range(ws.Cells(LS_R0, LS_C_CHAVE), ws.Cells(ult + 1, LS_C_CHAVE)).Value
        For i = 1 To UBound(v, 1)
            k = UCase$(Trim$(CStr(v(i, 1))))
            If k <> "" Then If Not d.Exists(k) Then d.Add k, LS_R0 + i - 1
        Next i
    End If
    Set IndiceStore = d
End Function

' Grava a tela Analitos!E:J nas linhas do lote, analito a analito, pela chave.
' Cria a linha so quando ha algum valor: lote vazio nao ocupa a store.
Public Sub SalvarParametrosView(ByVal lote As String)
    Dim wa As Worksheet, ws As Worksheet, nomes As Variant, vw As Variant
    Dim idx As Object, ult As Long, i As Long, j As Long, r As Long
    Dim nm As String, k As String, temValor As Boolean, linha(1 To 1, 1 To 6) As Variant
    Dim prot As Boolean
    lote = Trim$(lote)
    If lote = "" Then Exit Sub
    Set wa = ThisWorkbook.Sheets(AN_ABA)
    Set ws = ThisWorkbook.Sheets(LS_ABA)
    nomes = wa.Range(wa.Cells(AN_R0, 1), wa.Cells(AN_R0 + CAP_ANALITOS - 1, 1)).Value
    vw = wa.Range(wa.Cells(AN_R0, AN_C_MED1), wa.Cells(AN_R0 + CAP_ANALITOS - 1, AN_C_MED1 + 5)).Value
    Set idx = IndiceStore(ws, ult)

    On Error GoTo restaura
    prot = LiberarEscrita(ws)
    For i = 1 To CAP_ANALITOS
        nm = Trim$(CStr(nomes(i, 1)))
        If nm <> "" Then
            k = ChaveLote(lote, nm)
            temValor = False
            For j = 1 To 6
                If Len(Trim$(CStr(vw(i, j)))) > 0 Then temValor = True
            Next j
            r = 0
            If idx.Exists(k) Then r = idx(k)
            If r = 0 And temValor Then
                ult = ult + 1
                If ult > LS_R0 + LS_CAPLINHAS - 1 Then
                    Err.Raise vbObjectError + 491, "mLotes.SalvarParametrosView", _
                        "LotesStore cheia (" & LS_CAPLINHAS & " linhas). Ampliar LS_CAPLINHAS e os nomes ls*."
                End If
                r = ult
                idx.Add k, r
            End If
            If r > 0 Then
                For j = 1 To 6
                    linha(1, j) = vw(i, j)
                Next j
                ws.Cells(r, LS_C_LOTE).NumberFormat = "@"   ' "010" fica "010"
                ws.Cells(r, LS_C_LOTE).Value = lote
                ws.Cells(r, LS_C_IDX).Value = i
                ws.Range(ws.Cells(r, LS_C_MED1), ws.Cells(r, LS_C_MED1 + 5)).Value = linha
                ws.Cells(r, LS_C_NOME).Value = nm
                ws.Cells(r, LS_C_CHAVE).Value = k
            End If
        End If
    Next i
restaura:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao ws, prot
    InvalidarLotes
    If nE <> 0 Then Err.Raise nE, "mLotes.SalvarParametrosView", sE
End Sub

' Carrega na tela os parametros do lote. Sem parametro, a celula fica VAZIA --
' nunca a do lote anterior.
Public Sub CarregarParametrosView(ByVal lote As String)
    Dim wa As Worksheet, nomes As Variant, outp() As Variant
    Dim i As Long, j As Long, r As Long, nm As String, k As String, prot As Boolean
    lote = Trim$(lote)
    Set wa = ThisWorkbook.Sheets(AN_ABA)
    GarantirLotes
    nomes = wa.Range(wa.Cells(AN_R0, 1), wa.Cells(AN_R0 + CAP_ANALITOS - 1, 1)).Value
    ReDim outp(1 To CAP_ANALITOS, 1 To 6)
    For i = 1 To CAP_ANALITOS
        nm = Trim$(CStr(nomes(i, 1)))
        If nm <> "" And lote <> "" Then
            k = ChaveLote(lote, nm)
            If mLSIdx.Exists(k) Then
                r = mLSIdx(k)
                For j = 1 To 6
                    If IsError(mLS(r, LS_C_MED1 + j - 1)) Then
                        outp(i, j) = Empty
                    Else
                        outp(i, j) = mLS(r, LS_C_MED1 + j - 1)
                    End If
                Next j
            End If
        End If
    Next i
    On Error GoTo restaura
    prot = LiberarEscrita(wa)
    wa.Range(wa.Cells(AN_R0, AN_C_MED1), wa.Cells(AN_R0 + CAP_ANALITOS - 1, AN_C_MED1 + 5)).Value = outp
    ThisWorkbook.Names("loteParam").RefersToRange.NumberFormat = "@"
    ThisWorkbook.Names("loteParam").RefersToRange.Value = lote
restaura:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao wa, prot
    If nE <> 0 Then Err.Raise nE, "mLotes.CarregarParametrosView", sE
End Sub

' ============================ TROCA DO LOTE EM ANALISE ============================
' Chamado quando Painel!H3 (loteSel) muda. Persiste a tela do lote anterior,
' carrega a do novo e refaz o motor: grafico, Westgard e eventos passam a
' falar do lote novo, e so dele.
Public Sub TrocarLoteAnalise()
    Dim novo As String, atual As String
    ' ADR-050: uma operacao so -- tela congelada, calculo manual, um recalculo
    mApp.InicioUsuario "Carregando o lote " & LotePainel() & "..."
    On Error GoTo falha
    novo = LotePainel()
    If Len(novo) > 0 Then GravarLoteSel novo       ' canoniza o que o usuario escolheu
    atual = LoteParamTela()
    If novo <> atual Then
        If atual <> "" Then SalvarParametrosView atual
        CarregarParametrosView novo
    End If
    InvalidarLotes
    mEstatistica.InvalidarParametros      ' ADR-050: o banco nao mudou; o indice fica
    mEstatistica.RecalcularAnalitoAtual
    mApp.Fim
    Exit Sub
falha:
    Dim s As String
    s = Err.Description
    mApp.Fim
    Avisar "Nao foi possivel trocar o lote em analise: " & s, vbExclamation
End Sub

' Workbook_Open: garante que Painel, tela de parametros e motor contam a
' mesma historia antes do primeiro clique.
Public Sub SincronizarLotesAoAbrir()
    On Error Resume Next
    If Trim$(CStr(ThisWorkbook.Names("loteSel").RefersToRange.Value)) = "" Then
        GravarLoteSel LoteCanonico(LoteAtivoCore())
    End If
    If LoteParamTela() <> LotePainel() Then
        If LoteParamTela() <> "" Then SalvarParametrosView LoteParamTela()
        CarregarParametrosView LotePainel()
    End If
    InvalidarLotes
End Sub

' ============================ EDICAO DE PARAMETRO ============================
' Worksheet_Change da aba Analitos, zona E4:J43. Grava NA HORA na linha do
' lote em tela e deixa rastro de quem mudou o que (ISO 15189 8.4).
Public Sub ParametroEditado(ByVal Target As Range)
    Dim wa As Worksheet, zona As Range, c As Range, lote As String
    Dim an As String, nv As Long, campo As String, antes As Object, a As Variant
    Dim ev As Boolean, invalidos As String
    Set wa = ThisWorkbook.Sheets(AN_ABA)
    Set zona = Intersect(Target, wa.Range("E4:J43"))
    If zona Is Nothing Then Exit Sub
    lote = LoteParamTela()
    If lote = "" Then
        Avisar "Nenhum lote em tela. Selecione o lote no Painel (celula H3) antes de editar Media/DP.", vbExclamation
        Exit Sub
    End If

    mApp.InicioUsuario "Gravando parametros do lote " & lote & "..."
    On Error GoTo falha

    ' "antes" vem da STORE (o que valia), nao da tela (que ja mudou)
    Set antes = CreateObject("Scripting.Dictionary")
    For Each c In zona.Cells
        an = Trim$(CStr(wa.Cells(c.Row, 1).Value))
        If an <> "" Then antes(c.Address) = ValorNaStore(lote, an, c.Column - AN_C_MED1 + 1)
    Next c

    SalvarParametrosView lote

    For Each c In zona.Cells
        an = Trim$(CStr(wa.Cells(c.Row, 1).Value))
        If an <> "" Then
            nv = (c.Column - AN_C_MED1) \ 2 + 1
            If ((c.Column - AN_C_MED1) Mod 2) = 0 Then campo = "Media" Else campo = "DP"
            a = antes(c.Address)
            If CStr(a) <> CStr(c.Value) Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_ALTERADO", "mLotes", _
                    0, Empty, "", lote, nv, an, a, c.Value, "", "", _
                    campo & " N" & nv, "Parametro " & campo & " N" & nv & " do lote " & lote & _
                    " alterado na aba Analitos."
            End If
            If Len(Trim$(CStr(c.Value))) > 0 Then
                If Not NumeroValido(c.Value) Then
                    invalidos = invalidos & vbCrLf & "  " & an & " " & campo & " N" & nv & ": nao numerico"
                ElseIf campo = "DP" And CDbl(c.Value) <= 0 Then
                    invalidos = invalidos & vbCrLf & "  " & an & " DP N" & nv & ": precisa ser maior que zero"
                End If
            End If
        End If
    Next c

    mEstatistica.InvalidarParametros      ' ADR-050: parametro mudou, banco nao
    mEstatistica.RecalcularAnalitoAtual
    mApp.Fim
    If Len(invalidos) > 0 Then
        Avisar "Parametro gravado, mas NAO sera usado ate ser corrigido " & _
               "(o lote fica sem media/DP e o Westgard nao e avaliado):" & invalidos, vbExclamation
    End If
    Exit Sub
falha:
    Dim s As String
    s = Err.Description
    mApp.Fim
    Avisar "Falha ao gravar parametro do lote " & lote & ": " & s, vbCritical
End Sub

Private Function ValorNaStore(ByVal lote As String, ByVal analito As String, ByVal j As Long) As Variant
    Dim k As String
    GarantirLotes
    k = ChaveLote(lote, analito)
    ValorNaStore = Empty
    If mLSIdx.Exists(k) Then
        If Not IsError(mLS(mLSIdx(k), LS_C_MED1 + j - 1)) Then ValorNaStore = mLS(mLSIdx(k), LS_C_MED1 + j - 1)
    End If
End Function

' MsgBox so com gente olhando: sob automacao (Interactive = False) um modal
' invisivel trava o Excel inteiro.
Private Sub Avisar(ByVal msg As String, ByVal estilo As VbMsgBoxStyle)
    If Application.Interactive Then MsgBox msg, estilo, "Lotes"
End Sub

' ============================ CONFERENCIA ============================
' Integridade da store: chave duplicada, chave que nao bate com lote+nome,
' linha com valor e sem chave. Devolve "ok" ou a lista de problemas.
Public Function ConferirLotesStore() As String
    Dim ws As Worksheet, ult As Long, v As Variant, i As Long, d As Object
    Dim k As String, esperado As String, prob As String, n As Long
    Set ws = ThisWorkbook.Sheets(LS_ABA)
    ult = ws.Cells(ws.rows.Count, LS_C_LOTE).End(xlUp).Row
    If ult < LS_R0 Then ConferirLotesStore = "ok|0": Exit Function
    v = ws.Range(ws.Cells(LS_R0, 1), ws.Cells(ult + 1, LS_C_CHAVE)).Value
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To UBound(v, 1) - 1
        k = UCase$(Trim$(CStr(v(i, LS_C_CHAVE))))
        If k <> "" Then
            n = n + 1
            esperado = ChaveLote(CStr(v(i, LS_C_LOTE)), CStr(v(i, LS_C_NOME)))
            If k <> esperado Then prob = prob & "; L" & (i + LS_R0 - 1) & " chave " & k & " <> " & esperado
            If d.Exists(k) Then prob = prob & "; duplicada " & k Else d.Add k, i
        ElseIf Len(Trim$(CStr(v(i, LS_C_MED1)) & CStr(v(i, LS_C_MED1 + 1)))) > 0 Then
            prob = prob & "; L" & (i + LS_R0 - 1) & " tem media/DP e nao tem chave"
        End If
    Next i
    If Len(prob) = 0 Then ConferirLotesStore = "ok|" & n Else ConferirLotesStore = "ERRO|" & n & prob
End Function

' ============================ LOTE EM USO (Configuracao) ============================
Private Function BlocoDoLote(ByVal lote As String) As Long
    ' posicao (1-based) do lote no registro; 0 se nao encontrado.
    ' So para as views OPERACIONAIS (Liberacao, Registros). Parametro nao usa bloco.
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

Private Sub SalvarViewNoBloco(ByVal iBloco As Long)
    ' grava as views operacionais (Liberacao assinaturas, Registros) no bloco do lote.
    If iBloco < 1 Then Exit Sub
    On Error Resume Next
    Dim stL As Worksheet, rb0 As Long, srcL As Range
    Set stL = ThisWorkbook.Sheets("LiberStore")
    rb0 = 2 + (iBloco - 1) * NLB_LIBER
    Set srcL = ThisWorkbook.Names("libView").RefersToRange          ' C4:F203 (200 x 4)
    stL.Range(stL.Cells(rb0, 1), stL.Cells(rb0 + NLB_LIBER - 1, 4)).Value = srcL.Value
    Dim wsReg As Worksheet, rr0 As Long, srcR As Range
    Set wsReg = ThisWorkbook.Sheets("RegistrosStore")
    rr0 = 2 + (iBloco - 1) * NLB_LIBER
    Set srcR = ThisWorkbook.Names("regView").RefersToRange          ' B4:M203 (200 x 12)
    wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 12)).Value = srcR.Value
End Sub

Private Sub CarregarBlocoNaView(ByVal iBloco As Long)
    ' Assinaturas de Liberacao: sempre carrega (limpa se o bloco estiver vazio) -- nunca herda.
    If iBloco < 1 Then Exit Sub
    On Error Resume Next
    Dim stL As Worksheet, rb0 As Long, dstL As Range
    Set stL = ThisWorkbook.Sheets("LiberStore")
    rb0 = 2 + (iBloco - 1) * NLB_LIBER
    Set dstL = ThisWorkbook.Names("libView").RefersToRange
    dstL.Value = stL.Range(stL.Cells(rb0, 1), stL.Cells(rb0 + NLB_LIBER - 1, 4)).Value
    Dim wsReg As Worksheet, rr0 As Long, dstR As Range
    Set wsReg = ThisWorkbook.Sheets("RegistrosStore")
    rr0 = 2 + (iBloco - 1) * NLB_LIBER
    Set dstR = ThisWorkbook.Names("regView").RefersToRange
    dstR.Value = wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 12)).Value
End Sub

' Chamado pelo Worksheet_Change da Configuracao quando o lote EM USO muda.
' Troca as views operacionais e leva o Painel junto: quem troca o lote de
' trabalho quer ver o lote novo. O caminho inverso (Painel -> lote em uso)
' nao existe de proposito.
Public Sub TrocarLote()
    Dim novo As String, atual As String, iAtual As Long, iNovo As Long, emOp As Boolean
    On Error GoTo fim
    novo = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
    atual = Trim$(CStr(ThisWorkbook.Names("loteCarregado").RefersToRange.Value))
    If novo = atual Then GoTo fim
    If novo = "" Then GoTo fim
    iNovo = BlocoDoLote(novo)
    If iNovo = 0 Then
        Avisar "O lote '" & novo & "' nao esta no registro (Configuracao). Cadastre-o antes de usar.", vbExclamation
        ThisWorkbook.Names("loteAtivo").RefersToRange.Value = atual
        GoTo fim
    End If
    mApp.InicioUsuario "Trocando o lote em uso para " & novo & "..."
    emOp = True
    iAtual = BlocoDoLote(atual)
    If iAtual > 0 Then SalvarViewNoBloco iAtual           ' persiste o lote que estava em uso
    CarregarBlocoNaView iNovo
    ThisWorkbook.Names("loteCarregado").RefersToRange.Value = novo
    SalvarViewNoBloco iNovo
    AtualizarListaLiberacao
    mEstatistica.InvalidarCache           ' Registros (repeticoes/calibracao) mudou de lote
    GravarLoteSel novo
    TrocarLoteAnalise
    ' ADR-050: Calculate, e nao CalculateFull. O CalculateFull refazia a pasta
    ' INTEIRA (~15 s com cinco anos de banco) para atualizar o que depende do
    ' lote em uso -- e isso o Calculate ja faz, porque so o que mudou fica sujo.
    Application.Calculate
    AtualizarEixos
    HookCharts
fim:
    If emOp Then mApp.Fim
End Sub

' ============================ LIBERACAO: LISTA DE CORRIDAS ============================
' Liberacao!A:B -- as corridas do lote EM USO, RUN crescente, e a data de cada uma.
'
' ADR-050: eram 600 formulas (AGGREGATE e MINIFS sobre o banco inteiro) que o
' Excel refazia a CADA gravacao no banco -- o grosso dos 25 s de uma
' importacao com cinco anos de dados. A regra e a mesma, numa varredura:
'   A = RUN com ao menos uma linha Ativa do lote em uso (rRunUnico=1, rLote)
'   B = a menor data Ativa desse RUN no lote (MINIFS)
' A ORDEM (RUN crescente) e a das formulas, e nao pode mudar: as assinaturas
' de libView (C:F) estao amarradas a POSICAO da linha.
' Chamada por mBanco.AtualizarFlagsBanco (toda gravacao/exclusao no banco),
' por TrocarLote e na instalacao.
Public Sub AtualizarListaLiberacao(Optional ByVal dados As Variant)
    Dim ws As Worksheet, lote As String, i As Long, n As Long, cod As String
    Dim d As Object, r As Long, dt As Variant, k As Variant, runs() As Long
    Dim a As Long, b As Long, tmp As Long, out() As Variant, prot As Boolean
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Liberação")
    lote = LoteAtivoCore()
    If IsMissing(dados) Then dados = CarregarDB()
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(dados) And Len(lote) > 0 Then
        For i = 1 To UBound(dados, 1)
            If Trim$(CStr(dados(i, COL_STATUS))) = ST_ATIVO Then
                If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
                    cod = Trim$(CStr(dados(i, COL_LOTE)))
                    If Len(cod) > 5 Then
                        If NucleoLote(cod) = lote Then
                            r = CLng(Val(dados(i, COL_RUN)))
                            dt = dados(i, COL_DATA)
                            If Not d.Exists(r) Then
                                d.Add r, dt
                            ElseIf IsDate(dt) Then
                                If Not IsDate(d(r)) Then
                                    d(r) = dt
                                ElseIf CDate(dt) < CDate(d(r)) Then
                                    d(r) = dt
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If
    n = d.Count
    If n > 0 Then
        ReDim runs(1 To n)
        For Each k In d.Keys
            a = a + 1: runs(a) = CLng(k)
        Next k
        For a = 2 To n                                  ' insercao: RUN crescente
            tmp = runs(a): b = a - 1
            Do While b >= 1
                If runs(b) <= tmp Then Exit Do
                runs(b + 1) = runs(b): b = b - 1
            Loop
            runs(b + 1) = tmp
        Next a
        Dim nTodasLib As Long
        nTodasLib = n
        If n > NLB_LIBER Then n = NLB_LIBER            ' as formulas tambem mostravam as 200 primeiras
        ReDim out(1 To n, 1 To 2)
        For a = 1 To n
            out(a, 1) = runs(a)
            If IsDate(d(runs(a))) Then out(a, 2) = CDate(d(runs(a))) Else out(a, 2) = ""
        Next a
    End If
    prot = LiberarEscrita(ws)
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + NLB_LIBER, 2)).ClearContents
    If n > 0 Then ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 2)).Value = out
    ' ADR-053: a lista cabe 200 corridas. Passando disso, DIZER -- antes o lote
    ' longo simplesmente nao mostrava as corridas novas para assinar.
    If nTodasLib > NLB_LIBER Then
        ws.Range("H2").Value = ChrW(9888) & " O lote " & lote & " tem " & nTodasLib & _
            " corridas; esta lista mostra as " & NLB_LIBER & " primeiras. Assine e arquive " & _
            "antes de continuar, ou troque de lote."
    Else
        ws.Range("H2").Value = ""
    End If
    RestaurarProtecao ws, prot
    Exit Sub
fim:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then RestaurarProtecao ws, prot
    On Error GoTo 0
    If nE <> 0 Then Err.Raise nE, "mLotes.AtualizarListaLiberacao", sE
End Sub

Public Sub FlushLoteAtual()
    ' persiste as views nos armazens (chamado no BeforeSave).
    On Error Resume Next
    Dim lote As String, i As Long
    lote = Trim$(CStr(ThisWorkbook.Names("loteCarregado").RefersToRange.Value))
    i = BlocoDoLote(lote)
    If i > 0 Then SalvarViewNoBloco i
    If LoteParamTela() <> "" Then SalvarParametrosView LoteParamTela()
End Sub
