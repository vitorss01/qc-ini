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

' TrocarLote -> TrocarLoteAnalise: a chamada interna entra na operacao que ja esta aberta
Private mDentroTrocaLote As Boolean
Private mLS As Variant                        ' snapshot LotesStore A..P
Private mLSIdx As Object                      ' chave -> linha do snapshot
Private mLSok As Boolean

' ---- cadastro de lotes (Configuracao) -- ADR-060 ----
' Linha do lote = POSICAO no cadastro: BlocoDoLote/BlocoDoLoteBI enderecam por ela
' os blocos de LiberStore/RegistrosStore. Lote novo e sempre ANEXADO depois da
' ultima linha usada -- reaproveitar um buraco herdaria as assinaturas de outro lote.
Private Const LCAD_ABA As String = "Configuração"
Private Const LCAD_R0 As Long = 26
Private Const LCAD_R1 As Long = 125
Private Const LCAD_C_LOTE As Long = 3          ' C  codigo (regLoteCol)
Private Const LCAD_C_VAL As Long = 4           ' D  validade (regValidadeCol)
Private Const LCAD_C_ORIG As Long = 5          ' E  origem do cadastro (texto para o usuario)
Private Const LCAD_C_SOMBRA As Long = 30       ' AD validade registrada (oculta) -- o "antes" da trilha

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
    Dim s As String, rng As Range, t As String, cand As String, vals As Variant, i As Long
    s = Trim$(CStr(v))
    LoteCanonico = s
    If s = "" Then Exit Function
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    ' ADR-062: UMA leitura em bloco (era celula a celula por COM, ~6 vezes por clique)
    vals = rng.Value
    For i = 1 To UBound(vals, 1)
        t = Trim$(CStr(vals(i, 1)))
        If t <> "" Then
            If t = s Then LoteCanonico = t: Exit Function
            If cand = "" And IsNumeric(s) And IsNumeric(t) Then
                If CDbl(s) = CDbl(t) Then cand = t
            End If
        End If
    Next i
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
    ' ADR-050: uma operacao so -- tela congelada, calculo manual, um recalculo. Chamada de DENTRO
    ' do TrocarLote, entra na operacao dele: o InicioUsuario zerava o contador e o resto do
    ' TrocarLote rodava com tela, eventos e calculo ligados (auditoria 04/10/2026).
    If mDentroTrocaLote Then
        mApp.Inicio "Carregando o lote " & LotePainel() & "..."
    Else
        mApp.InicioUsuario "Carregando o lote " & LotePainel() & "..."
    End If
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
    mApp.fim
    Exit Sub
falha:
    Dim s As String
    s = Err.Description
    mApp.fim
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
    mApp.fim
    If Len(invalidos) > 0 Then
        Avisar "Parametro gravado, mas NAO sera usado ate ser corrigido " & _
               "(o lote fica sem media/DP e o Westgard nao e avaliado):" & invalidos, vbExclamation
    End If
    Exit Sub
falha:
    Dim s As String
    s = Err.Description
    mApp.fim
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
    Set srcR = ThisWorkbook.Names("regView").RefersToRange          ' B4:J203 (200 x 9) -- ADR-057: sem Rep 1-3
    wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 9)).Value = srcR.Value
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
    dstR.Value = wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 9)).Value
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
    mDentroTrocaLote = True
    TrocarLoteAnalise
    mDentroTrocaLote = False
    ' ADR-050: Calculate, e nao CalculateFull. O CalculateFull refazia a pasta
    ' INTEIRA (~15 s com cinco anos de banco) para atualizar o que depende do
    ' lote em uso -- e isso o Calculate ja faz, porque so o que mudou fica sujo.
    Application.Calculate
    AtualizarEixos
    HookCharts
fim:
    mDentroTrocaLote = False
    If emOp Then mApp.fim
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
    Dim eqF As String
    eqF = mEstatistica.EquipFiltro()        ' ADR-057: corridas do equipamento em analise
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(dados) And Len(lote) > 0 Then
        For i = 1 To UBound(dados, 1)
            If Trim$(CStr(dados(i, COL_STATUS))) = ST_ATIVO And _
               (Len(eqF) = 0 Or UCase$(Trim$(CStr(dados(i, COL_EQUIP)))) = eqF) Then
                If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
                    cod = Trim$(CStr(dados(i, COL_LOTE)))
                    If Len(cod) > 0 Then
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

' ============================================================================
'  LOTES AUTOMATICOS E VALIDADE (ADR-060, 03/10/2026)
'
'  Pedido do usuario: "nao precisar cadastrar lotes". Todo lote que chega pelo
'  INTERFACEAMENTO com analito cadastrado entra sozinho no cadastro (Configuracao,
'  coluna C), anexado no fim, com a origem na coluna E e a validade VAZIA. Falta de
'  validade nao e aprovacao: depois de cada login e de cada ATUALIZAR DADOS o
'  sistema avisa e oferece levar o usuario ate a celula. Se ele adiar, o aviso volta
'  na proxima abertura -- a pendencia e a propria celula vazia, nao um flag.
'
'  O que NAO entra sozinho (mesmo filtro do QA I03): resultado MANUAL (lote digitado
'  errado viraria lote fantasma) e analito sem cadastro (TNIH, urina, FERR...).
'  Media/DP continuam por lote e nao sao criadas aqui (ADR-049: lote novo nasce sem
'  parametro e o Painel mostra SEM MEDIA/DP ate alguem digitar).
' ============================================================================

' Coluna inteira da tblCQ_Final como matriz 2D (1 linha tambem vira 2D).
Private Function ColunaCQ(ByVal lo As ListObject, ByVal nome As String) As Variant
    Dim v As Variant, u(1 To 1, 1 To 1) As Variant
    v = lo.ListColumns(nome).DataBodyRange.Value
    If IsArray(v) Then
        ColunaCQ = v
    Else
        u(1, 1) = v
        ColunaCQ = u
    End If
End Function

' Mesma igualdade do LoteCanonico: texto (sem caixa) ou, na falta, valor numerico.
Private Function LoteNoCadastro(ByVal lote As String, ByRef cad As Variant) As Boolean
    Dim i As Long, t As String
    For i = 1 To UBound(cad, 1)
        t = Trim$(CStr(cad(i, 1)))
        If Len(t) > 0 Then
            If StrComp(t, lote, vbTextCompare) = 0 Then LoteNoCadastro = True: Exit Function
            If IsNumeric(t) And IsNumeric(lote) Then
                If CDbl(t) = CDbl(lote) Then LoteNoCadastro = True: Exit Function
            End If
        End If
    Next i
End Function

' Ultima linha do cadastro com codigo (LCAD_R0 - 1 se vazio).
Private Function UltimaLinhaCadastro(ByRef cad As Variant) As Long
    Dim i As Long
    UltimaLinhaCadastro = LCAD_R0 - 1
    For i = UBound(cad, 1) To 1 Step -1
        If Len(Trim$(CStr(cad(i, 1)))) > 0 Then UltimaLinhaCadastro = LCAD_R0 + i - 1: Exit Function
    Next i
End Function

' Validade de verdade: data (ou numero de serie de data) positiva. Texto que
' "parece data" nao vale -- as formulas do Inicio fariam conta com ele.
Public Function ValidadeOk(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDate: ValidadeOk = (CDbl(v) > 0)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency: ValidadeOk = (CDbl(v) > 0)
    End Select
End Function

Private Function DataTexto(ByVal v As Variant) As String
    If ValidadeOk(v) Then DataTexto = Format$(CDate(CDbl(v)), "dd/mm/yyyy") Else DataTexto = Trim$(CStr(v))
End Function

' Le a tblCQ_Final e ANEXA ao cadastro os lotes que ainda nao estao nele.
' Sem dialogo e sem levantar erro (roda dentro do ATUALIZAR DADOS).
' Devolve "OK|<n>|<lote1;lote2>|<os que nao couberam>" ou "ERRO|<motivo>".
Public Function RegistrarLotesRecebidos() As String
    Dim lo As ListObject, ws As Worksheet, n As Long, i As Long, j As Long
    Dim vL As Variant, vE As Variant, vA As Variant, vO As Variant, vD As Variant
    Dim novos As Object, k As Variant, lote As String, eq As String, info As Variant
    Dim cad As Variant, ult As Long, r As Long, chaves() As String, nn As Long, tmp As String
    Dim gravados As String, fora As String, nGrav As Long, prot As Boolean, ev As Boolean
    Dim dIni As Double, dFim As Double, origem As String, nE As Long, sE As String, abriu As Boolean

    On Error GoTo falha
    Set lo = mDados.TabelaCQ()
    If lo Is Nothing Then RegistrarLotesRecebidos = "OK|0||": Exit Function
    n = lo.ListRows.Count
    If n = 0 Then RegistrarLotesRecebidos = "OK|0||": Exit Function
    vL = ColunaCQ(lo, "LOTE"): vE = ColunaCQ(lo, "EQUIPAMENTO")
    vA = ColunaCQ(lo, "ANALITO_CADASTRADO"): vO = ColunaCQ(lo, "ORIGEM_RESULTADO")
    vD = ColunaCQ(lo, "DATA_HORA")

    ' lote -> Array(1a data, ultima data, n, equipamentos)
    Set novos = CreateObject("Scripting.Dictionary")
    novos.CompareMode = 1
    For i = 1 To n
        If UCase$(Trim$(CStr(vO(i, 1)))) = "INTERFACEAMENTO" Then
            If UCase$(Trim$(CStr(vA(i, 1)))) = "SIM" Then
                lote = Trim$(CStr(vL(i, 1)))
                If Len(lote) > 0 And IsDate(vD(i, 1)) Then
                    eq = Trim$(CStr(vE(i, 1)))
                    If Not novos.Exists(lote) Then
                        novos.Add lote, Array(CDbl(CDate(vD(i, 1))), CDbl(CDate(vD(i, 1))), 0&, "")
                    End If
                    info = novos(lote)
                    If CDbl(CDate(vD(i, 1))) < info(0) Then info(0) = CDbl(CDate(vD(i, 1)))
                    If CDbl(CDate(vD(i, 1))) > info(1) Then info(1) = CDbl(CDate(vD(i, 1)))
                    info(2) = info(2) + 1
                    If Len(eq) > 0 And InStr(1, "|" & info(3) & "|", "|" & eq & "|", vbTextCompare) = 0 Then
                        info(3) = info(3) & IIf(Len(info(3)) > 0, "|", "") & eq
                    End If
                    novos(lote) = info
                End If
            End If
        End If
    Next i

    Set ws = ThisWorkbook.Sheets(LCAD_ABA)
    cad = ws.Range(ws.Cells(LCAD_R0, LCAD_C_LOTE), ws.Cells(LCAD_R1, LCAD_C_LOTE)).Value
    ' so os que nao estao no cadastro, em ordem do 1o resultado
    ReDim chaves(1 To novos.Count + 1)
    For Each k In novos.Keys
        If Not LoteNoCadastro(CStr(k), cad) Then nn = nn + 1: chaves(nn) = CStr(k)
    Next k
    If nn = 0 Then RegistrarLotesRecebidos = "OK|0||": Exit Function
    For i = 2 To nn
        tmp = chaves(i): j = i - 1
        Do While j >= 1
            If novos(chaves(j))(0) <= novos(tmp)(0) Then Exit Do
            chaves(j + 1) = chaves(j): j = j - 1
        Loop
        chaves(j + 1) = tmp
    Next i

    ult = UltimaLinhaCadastro(cad)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = LiberarEscrita(ws): abriu = True
    For i = 1 To nn
        info = novos(chaves(i))
        r = ult + 1
        If r > LCAD_R1 Then
            fora = fora & IIf(Len(fora) > 0, ";", "") & chaves(i)
        Else
            dIni = info(0): dFim = info(1)
            origem = "automático em " & Format$(Date, "dd/mm/yyyy") & " · " & info(2) & " resultado(s) de " & _
                     Format$(CDate(dIni), "dd/mm/yyyy") & " a " & Format$(CDate(dFim), "dd/mm/yyyy") & _
                     IIf(Len(info(3)) > 0, " · " & Replace(info(3), "|", ", "), "")
            ws.Cells(r, LCAD_C_LOTE).NumberFormat = "@"
            ws.Cells(r, LCAD_C_LOTE).Value = chaves(i)
            ws.Cells(r, LCAD_C_VAL).ClearContents
            ws.Cells(r, LCAD_C_ORIG).Value = origem
            ws.Cells(r, LCAD_C_SOMBRA).ClearContents
            ult = r
            nGrav = nGrav + 1
            gravados = gravados & IIf(Len(gravados) > 0, ";", "") & chaves(i)
        End If
    Next i
    RestaurarProtecao ws, prot: abriu = False
    Application.EnableEvents = ev

    ' trilha: um evento por lote (fora do bloco destravado -- Auditar cuida da propria aba)
    On Error Resume Next
    For i = 1 To nn
        If InStr(1, ";" & gravados & ";", ";" & chaves(i) & ";", vbTextCompare) > 0 Then
            info = novos(chaves(i))
            mAuditoria.Auditar mAuditoria.CAT_CONFIG, "LOTE_CADASTRADO", "mLotes", 0, CDate(info(0)), _
                Replace(info(3), "|", ", "), chaves(i), 0, "", Empty, Empty, "", "SEM_VALIDADE", "AUTOMATICO", _
                "Lote recebido do interfaceamento (" & info(2) & " resultado(s), de " & Format$(CDate(info(0)), "dd/mm/yyyy") & _
                " a " & Format$(CDate(info(1)), "dd/mm/yyyy") & ") registrado automaticamente; validade pendente."
        End If
    Next i
    If Len(fora) > 0 Then mAuditoria.RegistrarLog "LOTE_NAO_CADASTRADO", "Cadastro cheio (100 lotes): " & fora
    On Error GoTo 0
    InvalidarLotes
    RegistrarLotesRecebidos = "OK|" & nGrav & "|" & gravados & "|" & fora
    Exit Function
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If abriu Then RestaurarProtecao ws, prot
    Application.EnableEvents = True
    RegistrarLotesRecebidos = "ERRO|" & nE & ": " & sE
End Function

' Lotes do cadastro sem validade valida. "<n>|<lote1;lote2>" (so leitura).
Public Function LotesSemValidade() As String
    Dim ws As Worksheet, v As Variant, i As Long, n As Long, lista As String
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets(LCAD_ABA)
    v = ws.Range(ws.Cells(LCAD_R0, LCAD_C_LOTE), ws.Cells(LCAD_R1, LCAD_C_VAL)).Value
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, 1)))) > 0 Then
            If Not ValidadeOk(v(i, 2)) Then
                n = n + 1
                lista = lista & IIf(Len(lista) > 0, ";", "") & Trim$(CStr(v(i, 1)))
            End If
        End If
    Next i
fim:
    LotesSemValidade = n & "|" & lista
End Function

Private Function PrimeiraLinhaSemValidade() As Long
    Dim ws As Worksheet, v As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(LCAD_ABA)
    v = ws.Range(ws.Cells(LCAD_R0, LCAD_C_LOTE), ws.Cells(LCAD_R1, LCAD_C_VAL)).Value
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, 1)))) > 0 And Not ValidadeOk(v(i, 2)) Then
            PrimeiraLinhaSemValidade = LCAD_R0 + i - 1: Exit Function
        End If
    Next i
End Function

' Texto do aviso (o mesmo no login e depois do ATUALIZAR DADOS).
Public Function TextoAvisoValidade() As String
    Dim r As String, n As Long, lista As String
    r = LotesSemValidade()
    n = CLng(Val(Split(r, "|")(0)))
    If n = 0 Then Exit Function
    lista = Replace(Split(r, "|")(1), ";", ", ")
    TextoAvisoValidade = IIf(n = 1, "Detectei um lote de controle SEM DATA DE VALIDADE:", _
                                    "Detectei " & n & " lotes de controle SEM DATA DE VALIDADE:") & vbCrLf & vbCrLf & _
        "     " & lista & vbCrLf & vbCrLf & _
        "O cadastro do lote é automático, mas a validade precisa ser informada por você (bula/rótulo do controle)." & vbCrLf & _
        "NÃO é recomendado — não devemos trabalhar sem registrar a data de validade." & vbCrLf & vbCrLf & _
        "RECOMENDADO: vamos inserir agora?"
End Function

' Depois do login e do ATUALIZAR DADOS. Nunca em automacao (modal invisivel trava o
' Excel) e nunca no meio de uma operacao.
Public Sub AvisarLotesSemValidade()
    Dim t As String, resp As VbMsgBoxResult
    If Not Application.Interactive Then Exit Sub
    If mApp.Ocupado() Then Exit Sub
    t = TextoAvisoValidade()
    If Len(t) = 0 Then Exit Sub
    resp = MsgBox(t, vbYesNo + vbExclamation + vbDefaultButton1, "Lotes — validade pendente")
    ResponderAvisoValidade (resp = vbYes)
End Sub

' Nucleo da resposta, sem a pergunta (o QA chama este).
Public Sub ResponderAvisoValidade(ByVal inserirAgora As Boolean)
    Dim lista As String
    If inserirAgora Then
        IrValidadeLotes
        Exit Sub
    End If
    lista = Replace(Split(LotesSemValidade(), "|")(1), ";", ", ")
    On Error Resume Next
    mAuditoria.Auditar mAuditoria.CAT_CONFIG, "VALIDADE_LOTE_ADIADA", "mLotes", 0, Empty, "", "", 0, "", _
                       Empty, Empty, "", "SEM_VALIDADE", "", "Usuario adiou o registro da validade: " & lista
    On Error GoTo 0
    If Application.Interactive Then
        MsgBox "NÃO é recomendado trabalhar sem a data de validade do lote." & vbCrLf & vbCrLf & _
               "Lote(s) sem validade: " & lista & vbCrLf & vbCrLf & _
               "Este aviso vai aparecer de novo toda vez que o arquivo for aberto, até a validade ser registrada " & _
               "(Configuração > cadastro de lotes, coluna Validade).", vbCritical, "Lotes — validade pendente"
    End If
End Sub

' Leva o usuario a celula da validade do primeiro lote pendente.
Public Sub IrValidadeLotes()
    Dim r As Long
    r = PrimeiraLinhaSemValidade()
    If r = 0 Then r = LCAD_R0
    mApp.Ir LCAD_ABA, "D" & r
    Application.StatusBar = "Digite a VALIDADE (dd/mm/aaaa) na coluna D de cada lote sem validade — " & _
                            "fica registrado na trilha de auditoria."
End Sub

' Worksheet_Change da Configuracao, coluna D do cadastro: confere e registra.
' O valor ANTERIOR vem da coluna-sombra AD (oculta), que sobrevive a reset do VBA.
Public Sub ValidadeEditada(ByVal Target As Range)
    Dim ws As Worksheet, alvo As Range, c As Range, r As Long, lote As String
    Dim antes As Variant, depois As Variant, prot As Boolean, ev As Boolean
    Set ws = ThisWorkbook.Sheets(LCAD_ABA)
    Set alvo = Intersect(Target, ws.Range(ws.Cells(LCAD_R0, LCAD_C_VAL), ws.Cells(LCAD_R1, LCAD_C_VAL)))
    If alvo Is Nothing Then Exit Sub
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = LiberarEscrita(ws)
    For Each c In alvo.Cells
        r = c.Row
        lote = Trim$(CStr(ws.Cells(r, LCAD_C_LOTE).Value))
        antes = ws.Cells(r, LCAD_C_SOMBRA).Value
        depois = c.Value
        If Len(Trim$(CStr(depois))) > 0 And Not ValidadeOk(depois) Then
            ' texto ou data invalida (colagem passa por cima da validacao de dados): volta
            c.Value = antes
            Avisar "Validade inválida para o lote " & lote & ": '" & CStr(depois) & "'." & vbCrLf & _
                   "Digite uma data (dd/mm/aaaa).", vbExclamation
        ElseIf DataTexto(antes) <> DataTexto(depois) Then
            If ValidadeOk(depois) Then c.NumberFormat = "dd/mm/yyyy"
            ws.Cells(r, LCAD_C_SOMBRA).Value = depois
            If Len(lote) > 0 Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "VALIDADE_LOTE_ALTERADA", "mLotes", 0, Empty, "", lote, 0, "", _
                    DataTexto(antes), DataTexto(depois), "", "", "", _
                    "Validade do lote " & lote & ": " & IIf(Len(DataTexto(antes)) = 0, "(vazia)", DataTexto(antes)) & _
                    " -> " & IIf(Len(DataTexto(depois)) = 0, "(vazia)", DataTexto(depois))
            End If
        End If
    Next c
fim:
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub
