Attribute VB_Name = "mLotesBlocos"
Option Explicit
' CAMADA: Lotes -- ADR-074: aba Analitos com UM BLOCO POR LOTE
'
' ============================================================================
'  Pedido do usuario (14/10/2026): cada lote cadastrado na Configuracao tem o seu
'  bloco na aba Analitos, com lote e validade herdados do cadastro, Media/DP de
'  cada nivel, os limites de controle (Media +-1, 2, 3 DP), o ano de referencia
'  do CLIA e da Variacao Biologica daquele lote e as especificacoes, o ET e o
'  Sigma calculados com ESSA referencia. Botao para abrir/recolher cada bloco.
'
'  Duas abas, dois papeis:
'    Especificacoes  a antiga aba Analitos, renomeada: catalogo de analitos,
'                    escolha da fonte (CLIA/VB/FAB), biblioteca anual das
'                    especificacoes (linhas 46-89) e o MOTOR (lote em analise).
'                    O Excel ajustou sozinho as ~3.300 formulas que a liam.
'    Analitos        os blocos (este modulo). Bloco k comeca na linha
'                    2 + 54*(k-1): cabecalho LOTE/VALIDADE, titulo, nota,
'                    cabecalho das colunas, 40 analitos, linha vazia, ANO DE
'                    REFERENCIA (CLIA em D:E, VB em J:K), nomes e as 6 linhas
'                    de especificacao, linha vazia.
'
'  FONTE UNICA: Media/DP continuam na LotesStore (chave LOTE|ANALITO). As celulas
'  E:J do bloco sao VALORES digitados pelo usuario (ADR-076, sem formula): digitar
'  grava na store na hora (com trilha e retrato da especificacao do bloco), e toda
'  gravacao na store por outro caminho (tela do motor, CIQ, preenchimento) reescreve
'  o bloco do lote. Os anos de referencia ficam no proprio bloco.
'
'  Blocos sao ANEXADOS (nunca apagados nem reordenados): lote retirado do cadastro
'  continua com o bloco -- e historico.
' ============================================================================
Private Const BL_ABA As String = "Analitos"
Private Const ESP_ABA As String = "Especificações"
Private Const BL_R0 As Long = 2
Private Const BL_PASSO As Long = 54
Private Const BL_NAN As Long = 40
' deslocamentos a partir da linha do cabecalho (LOTE / VALIDADE)
Private Const O_TIT As Long = 1
Private Const O_NOTA As Long = 2
Private Const O_CAB As Long = 3
Private Const O_AN As Long = 4
Private Const O_ANO As Long = 45
Private Const O_NOMES As Long = 46
Private Const O_CLIA As Long = 47
Private Const O_CLIA3 As Long = 48
Private Const O_CVI As Long = 49
Private Const O_CVG As Long = 50
Private Const O_FABE As Long = 51
Private Const O_FABC As Long = 52
Private Const O_FIM As Long = 53
' colunas
Private Const C_MED1 As Long = 5             ' E..J = Media/DP N1..N3 (como na Especificacoes e na LotesStore C..H)
Private Const C_ANOCLIA As Long = 4          ' D (D:E mescladas)
Private Const C_ANOVB As Long = 10           ' J (J:K mescladas)
Private Const C_TAB_FIM As Long = 23         ' W  fim da tabela de especificacoes
Private Const C_LIM1 As Long = 25            ' Y  -3DP N1 .. +3DP N1 (7 colunas por nivel)
Private Const C_BIAS As Long = 47            ' AU Bias do CEQ (abs)
Private Const C_DES1 As Long = 48            ' AV CV / ET / Sigma N1 (3 colunas por nivel)
Private Const C_ULT As Long = 56             ' BD
Private Const C_AUX As Long = 58             ' BF deslocamento CLIA na biblioteca (linha do ano)
Private Const C_AUX2 As Long = 59            ' BG deslocamento VB na biblioteca
Private Const C_SOMB1 As Long = 60           ' BH ano CLIA registrado (o "antes" da trilha)
Private Const C_SOMB2 As Long = 61           ' BI ano VB registrado
Private Const N_CIQ As Long = 20

Private mDentro As Boolean

' ============================ PRESENCA ============================
' Os blocos so existem depois da migracao (ADR-074): a aba do motor chama Especificacoes.
Public Function BlocosAtivos() As Boolean
    BlocosAtivos = ExisteAba(ESP_ABA) And ExisteAba(BL_ABA)
End Function

Private Function ExisteAba(ByVal n As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(n)
    ExisteAba = Not ws Is Nothing
End Function

Private Function EspRef() As String
    EspRef = "'" & ESP_ABA & "'!"
End Function

Private Function Col(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + ((c - 1) Mod 26)) & s
        c = (c - 1) \ 26
    Loop
    Col = s
End Function

' Linha do cabecalho do bloco que contem a linha r (0 = acima do 1o bloco).
Private Function CabecalhoDe(ByVal r As Long) As Long
    If r < BL_R0 Then Exit Function
    CabecalhoDe = r - ((r - BL_R0) Mod BL_PASSO)
End Function

Private Function EhCabecalho(ByVal ws As Worksheet, ByVal h As Long) As Boolean
    If h < BL_R0 Then Exit Function
    EhCabecalho = (UCase$(Trim$(CStr(ws.Cells(h, 1).Value))) = "LOTE")
End Function

' Blocos existentes, na ordem: lote -> linha do cabecalho. proximo = linha do proximo bloco.
Private Function BlocosExistentes(ByVal ws As Worksheet, ByRef proximo As Long) As Object
    Dim d As Object, h As Long, lote As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    h = BL_R0
    Do While EhCabecalho(ws, h)
        lote = Trim$(CStr(ws.Cells(h, 2).Value))
        If Len(lote) > 0 Then
            If Not d.Exists(lote) Then d.Add lote, h
        End If
        h = h + BL_PASSO
    Loop
    proximo = h
    Set BlocosExistentes = d
End Function

' Lotes do cadastro (Configuracao!C26:C125), na ordem, sem repetir.
Private Function LotesDoCadastro() As Collection
    Dim c As New Collection, v As Variant, i As Long, s As String, vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = 1
    On Error Resume Next
    v = ThisWorkbook.Names("regLoteCol").RefersToRange.Value
    On Error GoTo 0
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            s = Trim$(CStr(v(i, 1)))
            If Len(s) > 0 Then
                If Not vistos.Exists(s) Then vistos.Add s, 1: c.Add s
            End If
        Next i
    End If
    Set LotesDoCadastro = c
End Function

' Linha do cabecalho do bloco do lote (0 = sem bloco).
Public Function LinhaDoBloco(ByVal lote As String) As Long
    Dim ws As Worksheet, d As Object, prox As Long
    If Not BlocosAtivos() Then Exit Function
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    Set d = BlocosExistentes(ws, prox)
    lote = Trim$(lote)
    If d.Exists(lote) Then LinhaDoBloco = d(lote)
End Function

' Ano padrao do lote novo: Especificacoes!S2 (espAno); sem ele, o ano corrente.
Private Function AnoPadrao() As Variant
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names("espAno").RefersToRange.Value
    On Error GoTo 0
    If IsNumeric(v) And Not IsEmpty(v) Then AnoPadrao = CLng(v) Else AnoPadrao = Year(Date)
End Function

' Anos que a biblioteca anual tem (lista da validacao: lstAnosEspec).
Private Function AnoNaBiblioteca(ByVal v As Variant) As Boolean
    Dim r As Range, c As Range
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    On Error Resume Next
    Set r = ThisWorkbook.Names("lstAnosEspec").RefersToRange
    On Error GoTo 0
    If r Is Nothing Then AnoNaBiblioteca = IsNumeric(v): Exit Function
    For Each c In r.Cells
        If CStr(c.Value) = Trim$(CStr(v)) Then AnoNaBiblioteca = True: Exit Function
    Next c
End Function

Private Function Niveis() As Long
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names("blNiveis").RefersToRange.Value
    If IsError(v) Or IsEmpty(v) Then v = Application.Evaluate(ThisWorkbook.Names("blNiveis").RefersTo)
    On Error GoTo 0
    If IsNumeric(v) And Not IsEmpty(v) Then Niveis = CLng(v)
    If Niveis < 1 Or Niveis > 3 Then Niveis = 3
End Function

' ============================ MONTAGEM ============================
' Instalador (e reparo): refaz TODOS os blocos com o formato e as formulas atuais, na ordem em que ja
' estavam, mais os lotes do cadastro que ainda nao tinham bloco. Preserva o que o usuario decidiu nos
' blocos (anos de referencia e o "antes" da trilha). Media/DP nao se perdem: moram na LotesStore.
' A linha 1 (titulo e botoes) e do instalador e nao e tocada. Devolve "OK|<blocos>" ou "ERRO|<motivo>".
Public Function MontarBlocos(ByVal nlv As Long) As String
    Dim ws As Worksheet, d As Object, prox As Long, ordem As Collection, k As Variant, h As Long
    Dim anos As Object, v As Variant, s As Shape, nomes As Collection, i As Long, n As Long
    Dim ev As Boolean, prot As Boolean
    If Not BlocosAtivos() Then MontarBlocos = "ERRO|sem as abas Analitos e Especificações": Exit Function
    On Error GoTo falha
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    mDentro = True
    prot = LiberarEscrita(ws)
    ThisWorkbook.Names.Add Name:="blNiveis", RefersTo:="=" & nlv

    ' 1. o que ja existe: ordem e anos de cada lote
    Set d = BlocosExistentes(ws, prox)
    Set ordem = New Collection
    Set anos = CreateObject("Scripting.Dictionary")
    anos.CompareMode = 1
    For Each k In d.Keys
        h = d(k)
        ordem.Add CStr(k)
        anos.Add CStr(k), Array(ws.Cells(h + O_ANO, C_ANOCLIA).Value, ws.Cells(h + O_ANO, C_ANOVB).Value, _
                                ws.Cells(h + O_ANO, C_SOMB1).Value, ws.Cells(h + O_ANO, C_SOMB2).Value)
    Next k
    For Each k In LotesDoCadastro()
        If Not d.Exists(CStr(k)) Then ordem.Add CStr(k): d.Add CStr(k), 0
    Next k

    ' 2. limpa da linha 2 para baixo (a linha 1 e do instalador) e os botoes dos blocos
    Set nomes = New Collection
    For Each s In ws.Shapes
        If s.Name Like "blk*" Then nomes.Add s.Name
    Next s
    For i = 1 To nomes.Count
        ws.Shapes(nomes(i)).Delete
    Next i
    i = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    If i < prox Then i = prox
    With ws.Range(ws.Rows(BL_R0), ws.Rows(i + 1))
        .Hidden = False
        .UnMerge
        .Validation.Delete
        .FormatConditions.Delete
        .Clear
        .RowHeight = ws.StandardHeight
    End With

    ' 3. colunas
    LargurasColunas ws, nlv

    ' 4. blocos
    h = BL_R0
    For Each k In ordem
        If anos.Exists(CStr(k)) Then
            v = anos(CStr(k))
            CriarBloco ws, h, CStr(k), v(0), v(1), False
            If Len(CStr(v(2))) > 0 Then ws.Cells(h + O_ANO, C_SOMB1).Value = v(2)
            If Len(CStr(v(3))) > 0 Then ws.Cells(h + O_ANO, C_SOMB2).Value = v(3)
        Else
            CriarBloco ws, h, CStr(k), AnoPadrao(), AnoPadrao(), False
        End If
        h = h + BL_PASSO
        n = n + 1
    Next k
    RestaurarProtecao ws, prot
    mDentro = False
    Application.EnableEvents = ev
    MontarBlocos = "OK|" & n
    Exit Function
falha:
    MontarBlocos = "ERRO|" & Err.Description & " (linha " & h & ")"
    On Error Resume Next
    RestaurarProtecao ws, prot
    mDentro = False
    Application.EnableEvents = True
End Function

' Cria o bloco de cada lote do cadastro que ainda nao tem bloco (anexado no fim, recolhido).
' Barato quando nao ha nada a fazer: chamado no login, ao ativar a aba, ao cadastrar lote
' e depois do ATUALIZAR DADOS. Devolve o numero de blocos criados.
Public Function SincronizarBlocos() As Long
    Dim ws As Worksheet, d As Object, prox As Long, k As Variant, faltam As Collection
    Dim prot As Boolean, ev As Boolean, n As Long, nE As Long, sE As String, lista As String
    If mDentro Then Exit Function
    If Not BlocosAtivos() Then Exit Function
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    Set d = BlocosExistentes(ws, prox)
    Set faltam = New Collection
    For Each k In LotesDoCadastro()
        If Not d.Exists(CStr(k)) Then faltam.Add CStr(k)
    Next k
    If faltam.Count = 0 Then Exit Function
    mDentro = True
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = LiberarEscrita(ws)
    If n = 0 And prox = BL_R0 Then LarguraSeVazia ws
    For Each k In faltam
        CriarBloco ws, prox, CStr(k), AnoPadrao(), AnoPadrao(), False
        prox = prox + BL_PASSO
        n = n + 1
        lista = lista & IIf(Len(lista) > 0, ", ", "") & CStr(k)
    Next k
fim:
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mDentro = False
    On Error Resume Next
    If n > 0 Then mAuditoria.RegistrarLog "BLOCO_LOTE_CRIADO", "Aba Analitos: bloco criado para o(s) lote(s) " & lista & _
                                                       " (anos de referencia CLIA/VB = " & AnoPadrao() & ", ajustaveis no bloco)."
    If nE <> 0 Then mAuditoria.RegistrarLog "BLOCO_LOTE_FALHOU", "Aba Analitos: " & sE
    SincronizarBlocos = n
End Function

Private Sub LarguraSeVazia(ByVal ws As Worksheet)
    LargurasColunas ws, Niveis()
End Sub

Private Sub LargurasColunas(ByVal ws As Worksheet, ByVal nlv As Long)
    Dim c As Long
    ws.Columns(1).ColumnWidth = 24
    For c = 2 To C_SOMB2
        ws.Columns(c).ColumnWidth = 9.5
    Next c
    ws.Columns(2).ColumnWidth = 11
    ws.Columns(4).ColumnWidth = 7
    ws.Columns(C_TAB_FIM + 1).ColumnWidth = ws.Columns(C_TAB_FIM).ColumnWidth   ' X = W (pedido do usuario)
    ws.Columns(C_BIAS - 1).ColumnWidth = 2             ' AT respiro
    ws.Columns(C_ULT + 1).ColumnWidth = 2              ' BE
    ws.Range(ws.Columns(C_AUX), ws.Columns(C_SOMB2)).Hidden = True
    ' Bioquimica trabalha com N1 e N2: as colunas do N3 (limites e desempenho) ficam ocultas
    ws.Range(ws.Columns(C_LIM1 + 14), ws.Columns(C_LIM1 + 20)).Hidden = (nlv < 3)
    ws.Range(ws.Columns(C_DES1 + 6), ws.Columns(C_DES1 + 8)).Hidden = (nlv < 3)
End Sub

Private Sub CriarBloco(ByVal ws As Worksheet, ByVal h As Long, ByVal lote As String, _
                       ByVal anoClia As Variant, ByVal anoVb As Variant, ByVal aberto As Boolean)
    Dim y As Long
    y = h + O_ANO
    With ws.Range(ws.Cells(h, 1), ws.Cells(h + O_FIM, C_SOMB2))
        .UnMerge
        .Validation.Delete
        .Clear
    End With
    ws.Range(ws.Rows(h), ws.Rows(h + O_FIM)).Hidden = False
    FormatarBloco ws, h
    ws.Cells(h, 1).Value = "LOTE"
    ws.Cells(h, 2).NumberFormat = "@"
    ws.Cells(h, 2).Value = lote
    ws.Cells(h, 5).Value = "VALIDADE"
    FormulasBloco ws, h
    ws.Cells(y, C_ANOCLIA).Value = anoClia
    ws.Cells(y, C_ANOVB).Value = anoVb
    ws.Cells(y, C_SOMB1).Value = anoClia
    ws.Cells(y, C_SOMB2).Value = anoVb
    BotoesBloco ws, h
    MostrarBloco ws, h, aberto
End Sub

' ---------------------------------------------------------------- formulas
Private Function Tx(ByVal s As String, ByVal h As Long) As String
    s = Replace(s, "~", Chr$(34))
    s = Replace(s, "{E}", EspRef())
    s = Replace(s, "{h}", CStr(h))
    s = Replace(s, "{r}", CStr(h + O_AN))
    s = Replace(s, "{a1}", CStr(h + O_AN + BL_NAN - 1))
    s = Replace(s, "{y}", CStr(h + O_ANO))
    s = Replace(s, "{sn}", CStr(h + O_NOMES))
    s = Replace(s, "{cl}", CStr(h + O_CLIA))
    s = Replace(s, "{cl3}", CStr(h + O_CLIA3))
    s = Replace(s, "{cvi}", CStr(h + O_CVI))
    s = Replace(s, "{cvg}", CStr(h + O_CVG))
    s = Replace(s, "{fe}", CStr(h + O_FABE))
    s = Replace(s, "{fc}", CStr(h + O_FABC))
    Tx = s
End Function

' ADR-076: Media/DP do bloco (E:J) sao VALORES digitados pelo usuario, sem formula. A LotesStore continua sendo
' onde o sistema guarda e le (motor, Calc, Westgard, BI); o bloco e sempre reescrito a partir dela.
Public Sub EscreverValoresDoLote(ByVal lote As String)
    Dim ws As Worksheet, h As Long, prot As Boolean, ev As Boolean, dentro As Boolean
    If Not BlocosAtivos() Then Exit Sub
    h = LinhaDoBloco(lote)
    If h = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    dentro = mDentro
    mDentro = True
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = LiberarEscrita(ws)
    EscreverValoresBloco ws, h
fim:
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mDentro = dentro
End Sub

Private Sub EscreverValoresBloco(ByVal ws As Worksheet, ByVal h As Long)
    Dim e As Worksheet, nomes As Variant
    Set e = ThisWorkbook.Worksheets(ESP_ABA)
    nomes = e.Range("A4:A43").Value
    ws.Range(ws.Cells(h + O_AN, C_MED1), ws.Cells(h + O_AN + BL_NAN - 1, C_MED1 + 5)).Value = _
        mLotes.ValoresDoLote(Trim$(CStr(ws.Cells(h, 2).Value)), nomes)
End Sub

' Coluna (de ate 40 linhas) do bloco a partir da formula da 1a linha de analito.
Private Sub ColunaAn(ByVal ws As Worksheet, ByVal h As Long, ByVal c As Long, ByVal f As String)
    ws.Range(ws.Cells(h + O_AN, c), ws.Cells(h + O_AN + BL_NAN - 1, c)).Formula2 = Tx(f, h)
End Sub

Private Sub FormulasBloco(ByVal ws As Worksheet, ByVal h As Long)
    Dim c As Long, nv As Long, k As Long, m As String, s As String, f As String
    Dim cv As String, sn As Long, y As Long, j As Long
    y = h + O_ANO: sn = h + O_NOMES

    ' cabecalho: validade herdada do cadastro, titulo e resumo
    ws.Range(ws.Cells(h, 6), ws.Cells(h, 6)).Formula2 = Tx("=LET(lt,$B${h},lin,IFERROR(MATCH(lt,'Configuração'!$C$26:$C$125,0)," & _
        "IFERROR(MATCH(--lt,'Configuração'!$C$26:$C$125,0),0)),IF(lin=0,~fora do cadastro~,IF(INDEX('Configuração'!$D$26:$D$125,lin)=~~," & _
        "~não informada~,INDEX('Configuração'!$D$26:$D$125,lin))))", h)
    ' ADR-078: conta so os analitos com algum par VALIDO (Media numerica e DP > 0) e avisa valores de TESTE (ADR-076)
    ws.Cells(h, 15).Formula2 = Tx("=LET(nn,SUMPRODUCT(--((ISNUMBER($E{r}:$E{a1})*ISNUMBER($F{r}:$F{a1})*($F{r}:$F{a1}>0)+" & _
        "ISNUMBER($G{r}:$G{a1})*ISNUMBER($H{r}:$H{a1})*($H{r}:$H{a1}>0)+ISNUMBER($I{r}:$I{a1})*ISNUMBER($J{r}:$J{a1})*($J{r}:$J{a1}>0))>0))," & _
        "ts,SUMPRODUCT((LotesStore!$A$2:$A$4001=~~&$B${h})*(ISNUMBER(LotesStore!$X$2:$X$4001)*(LotesStore!$X$2:$X$4001>0)+" & _
        "ISNUMBER(LotesStore!$Y$2:$Y$4001)*(LotesStore!$Y$2:$Y$4001>0)+ISNUMBER(LotesStore!$Z$2:$Z$4001)*(LotesStore!$Z$2:$Z$4001>0)))," & _
        "tot,COUNTIF($A{r}:$A{a1},~?*~),ult,IFERROR(MAXIFS(LotesStore!$U:$U,LotesStore!$A:$A,$B${h}),0)," & _
        "IF(nn=0,~SEM MÉDIA/DP: o gráfico deste lote fica sem limites e o Westgard não avalia. Digite a Média e o DP de cada nível (E:J) ou use o botão CIQ.~," & _
        "~Média/DP em ~&nn&~ de ~&tot&~ analitos~&IF(ult>0,~ · último registro ~&TEXT(DAY(ult),~00~)&~/~&TEXT(MONTH(ult),~00~)&~/~&YEAR(ult),~~))" & _
        "&~ · referência: CLIA ~&$D${y}&~ · VB ~&$J${y}&IF(ts>0,~ · contém valores de TESTE (ADR-076) em ~&ts&~ par(es) analito/nível: substitua pela Média/DP do lote~,~~))", h)
    ws.Cells(h + O_TIT, 1).Formula2 = Tx("=~ANALITOS · MÉDIA/DP, LIMITES E ESPECIFICAÇÕES DO LOTE ~&$B${h}", h)

    ' catalogo (Especificacoes, mesma ordem)
    ColunaAn ws, h, 1, "=IF({E}$A4=~~,~~,{E}$A4)"
    ColunaAn ws, h, 2, "=IF($A{r}=~~,~~,{E}$B4&~~)"
    ColunaAn ws, h, 3, "=IF($A{r}=~~,~~,{E}$C4&~~)"
    ColunaAn ws, h, 4, "=IF($A{r}=~~,~~,{E}$D4)"
    ' Media/DP (E:J): VALORES (ADR-076), escritos a partir da LotesStore
    EscreverValoresBloco ws, h
    ' especificacoes deste bloco (linhas de especificacao abaixo, pelo nome do analito)
    ColunaAn ws, h, 11, "=IFERROR(LET(val,INDEX($B${cl}:$AO${cl},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 12, "=IFERROR(LET(val,INDEX($B${cl3}:$AO${cl3},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 13, "=IFERROR(LET(val,INDEX($B${fe}:$AO${fe},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 14, "=IFERROR(LET(val,INDEX($B${fc}:$AO${fc},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 15, "=IFERROR(LET(val,INDEX($B${cvi}:$AO${cvi},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 16, "=IFERROR(LET(val,INDEX($B${cvg}:$AO${cvg},MATCH($A{r},$B${sn}:$AO${sn},0)),IF(val=~~,~~,val)),~~)"
    ColunaAn ws, h, 17, "=IF($A{r}=~~,~~,IFERROR(INDEX({E}$Q$4:$Q$43,MATCH($A{r},{E}$A$4:$A$43,0))&~~,~~))"
    ' ETp VB (a mesma formula da Especificacoes!R: exige CVi E CVg)
    ColunaAn ws, h, 18, "=IF(AND(N($O{r})>0,N($P{r})>0),IFERROR(IF($Q{r}=~ÓTI~,1.65*(0.25*$O{r})+0.125*SQRT($O{r}^2+$P{r}^2)," & _
        "IF($Q{r}=~DES~,1.65*(0.5*$O{r})+0.25*SQRT($O{r}^2+$P{r}^2),IF($Q{r}=~MIN~,1.65*(0.75*$O{r})+0.375*SQRT($O{r}^2+$P{r}^2),~~))),~~),~~)"
    ColunaAn ws, h, 19, "=IF($A{r}=~~,~~,IFERROR(INDEX({E}$S$4:$S$43,MATCH($A{r},{E}$A$4:$A$43,0))&~~,~~))"
    ' ETp em uso e CVTp: a fonte escolhida (S) decide (ADR-069); sem especificacao, "-" (ADR-068)
    ColunaAn ws, h, 20, "=IF($A{r}=~~,~~,IF($S{r}=~CLIA~,IF(N($K{r})>0,$K{r},~-~),IF($S{r}=~VB~,IF(N($R{r})>0,$R{r},~-~)," & _
        "IF($S{r}=~FAB~,IF(N($M{r})>0,$M{r},~-~),~-~))))"
    ColunaAn ws, h, 21, "=IF($A{r}=~~,~~,IF($S{r}=~CLIA~,IF(N($V{r})>0,$V{r},~-~),IF($S{r}=~VB~,IF(N($W{r})>0,$W{r},~-~)," & _
        "IF($S{r}=~FAB~,IF(N($N{r})>0,$N{r},~-~),~-~))))"
    ColunaAn ws, h, 22, "=IF($A{r}=~~,~~,IF(N($K{r})>0,$K{r}/3,~~))"
    ColunaAn ws, h, 23, "=IF($A{r}=~~,~~,IF(N($O{r})>0,IFERROR(INDEX({E}$Y$4:$Y$6,MATCH($Q{r},{E}$X$4:$X$6,0)),0.75)*$O{r},~~))"

    ' limites de controle: Media +- 1, 2, 3 DP de cada nivel (sem DP valido, vazio -- nunca de outro lote)
    For nv = 1 To 3
        m = "$" & Col(C_MED1 + (nv - 1) * 2)
        s = "$" & Col(C_MED1 + (nv - 1) * 2 + 1)
        For k = -3 To 3
            c = C_LIM1 + (nv - 1) * 7 + (k + 3)
            If k = 0 Then
                f = "=IF(OR(N(" & s & "{r})<=0,NOT(ISNUMBER(" & m & "{r}))),~~," & m & "{r})"
            Else
                f = "=IF(OR(N(" & s & "{r})<=0,NOT(ISNUMBER(" & m & "{r}))),~~," & m & "{r}" & IIf(k > 0, "+", "-") & Abs(k) & "*" & s & "{r})"
            End If
            ColunaAn ws, h, c, f
        Next k
    Next nv

    ' desempenho do lote: Bias do CEQ (selecao da aba Estatistica), CV do lote, ET e Sigma com a referencia do bloco
    ColunaAn ws, h, C_BIAS, "=IF($A{r}=~~,~~,IFERROR(LET(bias,INDEX('Estatística'!$G$14:$G$400,MATCH($A{r},'Estatística'!$A$14:$A$400,0))," & _
        "IF(ISNUMBER(bias),ABS(bias),~~)),~~))"
    For nv = 1 To 3
        m = "$" & Col(C_MED1 + (nv - 1) * 2)
        s = "$" & Col(C_MED1 + (nv - 1) * 2 + 1)
        c = C_DES1 + (nv - 1) * 3
        cv = "$" & Col(c)
        ColunaAn ws, h, c, "=IF(OR(N(" & s & "{r})<=0,N(" & m & "{r})=0),~~,100*" & s & "{r}/ABS(" & m & "{r}))"
        ColunaAn ws, h, c + 1, "=IF(OR(NOT(ISNUMBER(" & cv & "{r})),NOT(ISNUMBER($AU{r}))),~~,1.65*" & cv & "{r}+$AU{r})"
        ColunaAn ws, h, c + 2, "=IF(OR(NOT(ISNUMBER(" & cv & "{r})),NOT(ISNUMBER($AU{r}))),~~,IF(NOT(ISNUMBER($T{r})),~-~,($T{r}-$AU{r})/" & cv & "{r}))"
    Next nv

    ' ano de referencia -> deslocamento do bloco do ano na biblioteca (Especificacoes!B46:AO89, blocos de 9 linhas)
    ws.Cells(y, C_AUX).Formula2 = Tx("=IFERROR((MATCH($D${y},CHOOSE({1;2;3;4;5},{E}$B$46,{E}$B$55,{E}$B$64,{E}$B$73,{E}$B$82),0)-1)*9,~~)", h)
    ws.Cells(y, C_AUX2).Formula2 = Tx("=IFERROR((MATCH($J${y},CHOOSE({1;2;3;4;5},{E}$B$46,{E}$B$55,{E}$B$64,{E}$B$73,{E}$B$82),0)-1)*9,~~)", h)
    ' nomes dos analitos (B..AO) e as especificacoes do ano escolhido
    ws.Range(ws.Cells(sn, 2), ws.Cells(sn, 41)).Formula2 = Tx("=IF(INDEX({E}$A$4:$A$43,COLUMN()-1)=~~,~~,INDEX({E}$A$4:$A$43,COLUMN()-1))", h)
    LinhaEsp ws, h, O_CLIA, "$BF${y}", 3
    ws.Range(ws.Cells(h + O_CLIA3, 2), ws.Cells(h + O_CLIA3, 41)).Formula2 = Tx("=IF(N(B${cl})>0,B${cl}/3,~~)", h)
    LinhaEsp ws, h, O_CVI, "$BG${y}", 5
    LinhaEsp ws, h, O_CVG, "$BG${y}", 6
    LinhaEsp ws, h, O_FABE, "$BF${y}", 7
    LinhaEsp ws, h, O_FABC, "$BF${y}", 8
End Sub

' Linha de especificacao: valor do analito (pelo NOME) na linha `deslocamento + idx` do bloco do ano.
Private Sub LinhaEsp(ByVal ws As Worksheet, ByVal h As Long, ByVal o As Long, ByVal desl As String, ByVal idx As Long)
    ws.Range(ws.Cells(h + o, 2), ws.Cells(h + o, 41)).Formula2 = Tx("=IF(OR(B${sn}=~~," & desl & "=~~),~~,IFERROR(LET(val,INDEX({E}$B$46:$AO$89," & _
        desl & "+" & idx & ",MATCH(B${sn},INDEX({E}$B$46:$AO$89," & desl & "+2,0),0)),IF(val=~~,~~,val)),~~))", h)
End Sub

' ---------------------------------------------------------------- formato
Private Sub FormatarBloco(ByVal ws As Worksheet, ByVal h As Long)
    Dim a0 As Long, a1 As Long, y As Long, sn As Long, nv As Long, c As Long, k As Long
    Dim cab As Variant, rot As Variant, cores As Variant
    a0 = h + O_AN: a1 = h + O_AN + BL_NAN - 1: y = h + O_ANO: sn = h + O_NOMES
    ws.Range(ws.Cells(h, 1), ws.Cells(h + O_FIM, C_ULT)).Font.Name = "Segoe UI"
    ws.Range(ws.Cells(h, 1), ws.Cells(h + O_FIM, C_ULT)).Font.Size = 9

    ' cabecalho do lote
    With ws.Range(ws.Cells(h, 1), ws.Cells(h, C_ULT))
        .Interior.Color = RGB(15, 42, 46)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(h).RowHeight = 26
    ws.Cells(h, 1).HorizontalAlignment = xlRight
    ws.Cells(h, 1).Font.Size = 11
    With ws.Range(ws.Cells(h, 2), ws.Cells(h, 4))
        .Merge
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 244, 241)
        .Font.Color = RGB(15, 42, 46)
        .Font.Size = 14
    End With
    ws.Cells(h, 5).HorizontalAlignment = xlRight
    With ws.Range(ws.Cells(h, 6), ws.Cells(h, 8))
        .Merge
        .HorizontalAlignment = xlCenter
        .NumberFormat = "dd/mm/yyyy"
        .Interior.Color = RGB(230, 244, 241)
        .Font.Color = RGB(15, 42, 46)
        .Font.Size = 11
    End With
    With ws.Range(ws.Cells(h, 15), ws.Cells(h, C_LIM1 + 6))
        .Merge
        .HorizontalAlignment = xlLeft
        .Font.Bold = False
        .Font.Italic = True
        .Font.Color = RGB(95, 227, 194)
        .ShrinkToFit = True
    End With

    ' titulo e nota
    With ws.Range(ws.Cells(h + O_TIT, 1), ws.Cells(h + O_TIT, C_TAB_FIM))
        .Merge
        .Interior.Color = RGB(27, 59, 64)
        .Font.Color = RGB(95, 227, 194)
        .Font.Bold = True
        .Font.Size = 12
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(h + O_TIT).RowHeight = 22
    With ws.Range(ws.Cells(h + O_NOTA, 1), ws.Cells(h + O_NOTA, C_TAB_FIM))
        .Merge
        .Value = "Média/DP valem só para este lote: digite em E:J (grava na hora na LotesStore, com trilha). " & _
                 "Limites de controle = Média ± 1, 2 e 3 DP. As especificações, o ETp, o ET e o Sigma usam os ANOS DE " & _
                 "REFERÊNCIA deste bloco (CLIA e Variação Biológica, abaixo da tabela). Fonte (CLIA/VB/FAB) e desempenho: aba Especificações."
        .WrapText = True
        .Font.Italic = True
        .Font.Color = RGB(64, 64, 64)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(h + O_NOTA).RowHeight = 30
    cores = Array(RGB(222, 235, 247), RGB(226, 239, 218), RGB(252, 228, 214))
    For nv = 1 To 3
        With ws.Range(ws.Cells(h + O_NOTA, C_LIM1 + (nv - 1) * 7), ws.Cells(h + O_NOTA, C_LIM1 + (nv - 1) * 7 + 6))
            .Merge
            .Value = "LIMITES DE CONTROLE N" & nv & " (Levey-Jennings)"
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = cores(nv - 1)
        End With
    Next nv
    With ws.Range(ws.Cells(h + O_NOTA, C_BIAS), ws.Cells(h + O_NOTA, C_ULT))
        .Merge
        .Value = "DESEMPENHO DO LOTE · Bias = CEQ (seleção da aba Estatística) · CV = DP/Média · ET = 1,65·CV + |Bias| · " & _
                 "Sigma = (ETp - |Bias|) / CV"
        .WrapText = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ' cabecalho das colunas
    cab = Array("Analito", "Área", "Unid.", "Casas", "Média N1", "DP N1", "Média N2", "DP N2", "Média N3", "DP N3", _
                "ETp CLIA %", "CVTp CLIA %", "ETp FAB %", "CVTp FAB %", "CVi %", "CVg %", "Desemp.", "ETp VB", _
                "ESPQ FONTE", "ETp em uso %", "CVTp %", "Lim CV CLIA %", "Lim CV VB %")
    For c = 0 To UBound(cab)
        ws.Cells(h + O_CAB, c + 1).Value = cab(c)
    Next c
    rot = Array("-3DP", "-2DP", "-1DP", "Média", "+1DP", "+2DP", "+3DP")
    For nv = 1 To 3
        For k = 0 To 6
            ws.Cells(h + O_CAB, C_LIM1 + (nv - 1) * 7 + k).Value = rot(k) & " N" & nv
        Next k
        ws.Cells(h + O_CAB, C_DES1 + (nv - 1) * 3).Value = "CV % N" & nv
        ws.Cells(h + O_CAB, C_DES1 + (nv - 1) * 3 + 1).Value = "ET % N" & nv
        ws.Cells(h + O_CAB, C_DES1 + (nv - 1) * 3 + 2).Value = "Sigma N" & nv
    Next nv
    ws.Cells(h + O_CAB, C_BIAS).Value = "Bias CEQ %"
    With ws.Range(ws.Cells(h + O_CAB, 1), ws.Cells(h + O_CAB, C_ULT))
        .Interior.Color = RGB(27, 59, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(h + O_CAB, C_TAB_FIM + 1), ws.Cells(h + O_CAB, C_TAB_FIM + 1)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(h + O_CAB, C_BIAS - 1), ws.Cells(h + O_CAB, C_BIAS - 1)).Interior.ColorIndex = xlNone
    ws.Rows(h + O_CAB).RowHeight = 30

    ' corpo da tabela
    With ws.Range(ws.Cells(a0, 1), ws.Cells(a1, C_TAB_FIM))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(a0, 1), ws.Cells(a1, 1)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(a0, 1), ws.Cells(a1, 1)).Font.Bold = True
    ws.Range(ws.Cells(a0, 1), ws.Cells(a1, 1)).ShrinkToFit = True
    With ws.Range(ws.Cells(a0, C_MED1), ws.Cells(a1, C_MED1 + 5))
        .Interior.Color = RGB(255, 255, 204)                    ' entrada
        .Font.Color = RGB(0, 0, 255)
        .NumberFormat = "0.0000"
        .Locked = False
        .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                        Formula1:="-1000000000", Formula2:="1000000000000"
        .Validation.ErrorTitle = "Média/DP"
        .Validation.ErrorMessage = "Digite um número (Média ou DP deste lote). Para apagar, use Delete."
        .Validation.InputTitle = "Média/DP deste lote"
        .Validation.InputMessage = "Grava na hora na LotesStore, com trilha de auditoria."
    End With
    ws.Range(ws.Cells(a0, 11), ws.Cells(a1, C_TAB_FIM)).NumberFormat = "0.00"
    ws.Range(ws.Cells(a0, 20), ws.Cells(a1, 21)).Font.Bold = True
    For nv = 1 To 3
        With ws.Range(ws.Cells(a0, C_LIM1 + (nv - 1) * 7), ws.Cells(a1, C_LIM1 + (nv - 1) * 7 + 6))
            .Interior.Color = cores(nv - 1)
            .NumberFormat = "0.000"
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(a0, C_LIM1 + (nv - 1) * 7 + 3), ws.Cells(a1, C_LIM1 + (nv - 1) * 7 + 3)).Font.Bold = True
    Next nv
    With ws.Range(ws.Cells(a0, C_BIAS), ws.Cells(a1, C_ULT))
        .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
    End With
    For nv = 1 To 3
        CoresSigma ws.Range(ws.Cells(a0, C_DES1 + (nv - 1) * 3 + 2), ws.Cells(a1, C_DES1 + (nv - 1) * 3 + 2))
    Next nv

    ' ano de referencia
    ws.Rows(y).RowHeight = 26
    With ws.Range(ws.Cells(y, 1), ws.Cells(y, 3))
        .Merge
        .Value = "CLIA - ANO DE REFERÊNCIA"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(y, 7), ws.Cells(y, 9))
        .Merge
        .Value = "VARIAÇÃO BIOLÓGICA - ANO DE REFERÊNCIA"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    AnoEntrada ws.Range(ws.Cells(y, C_ANOCLIA), ws.Cells(y, C_ANOCLIA + 1)), "Ano da referência CLIA deste lote"
    AnoEntrada ws.Range(ws.Cells(y, C_ANOVB), ws.Cells(y, C_ANOVB + 1)), "Ano da referência de Variação Biológica deste lote"
    With ws.Range(ws.Cells(y, 13), ws.Cells(y, C_TAB_FIM))
        .Merge
        .Value = "Os anos escolhem as especificações abaixo (biblioteca anual: aba Especificações, linhas 46-89) e com elas o " & _
                 "ETp, o ET e o Sigma deste lote. FAB segue o ano do CLIA."
        .WrapText = True
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .VerticalAlignment = xlCenter
    End With

    ' especificacoes de qualidade do bloco
    ws.Rows(sn).RowHeight = 48
    ws.Cells(sn, 1).Value = "ESPECIFICAÇÕES DE QUALIDADE"
    With ws.Range(ws.Cells(sn, 1), ws.Cells(sn, 41))
        .Interior.Color = RGB(27, 59, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(sn, 1).Font.Size = 8
    ws.Range(ws.Cells(sn, 2), ws.Cells(sn, 41)).Font.Size = 7
    rot = Array("CLIA ETp %", "CLIA 1/3 ETp % (CVTp)", "VB - CVi %", "VB - CVg %", "FAB ETp %", "FAB CVTp %")
    For k = 0 To 5
        ws.Cells(h + O_CLIA + k, 1).Value = rot(k)
    Next k
    With ws.Range(ws.Cells(h + O_CLIA, 1), ws.Cells(h + O_FABC, 41))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.0#"
    End With
    With ws.Range(ws.Cells(h + O_CLIA, 1), ws.Cells(h + O_FABC, 1))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Range(ws.Cells(h + O_CLIA, 2), ws.Cells(h + O_CLIA3, 41)).Interior.Color = RGB(222, 235, 247)
    ws.Range(ws.Cells(h + O_CVI, 2), ws.Cells(h + O_CVG, 41)).Interior.Color = RGB(226, 239, 218)
    ws.Range(ws.Cells(h + O_FABE, 2), ws.Cells(h + O_FABC, 41)).Interior.Color = RGB(252, 228, 214)
End Sub

Private Sub AnoEntrada(ByVal r As Range, ByVal titulo As String)
    r.Merge
    r.HorizontalAlignment = xlCenter
    r.VerticalAlignment = xlCenter
    r.Font.Bold = True
    r.Font.Size = 11
    r.Font.Color = RGB(0, 0, 255)
    r.Interior.Color = RGB(255, 255, 204)
    r.Borders.LineStyle = xlContinuous
    r.Borders.Color = RGB(166, 166, 166)
    r.NumberFormat = "0"
    r.Locked = False
    r.Validation.Delete
    On Error Resume Next
    r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lstAnosEspec"
    r.Validation.ErrorTitle = "Ano de referência"
    r.Validation.ErrorMessage = "Escolha um ano da lista: são os anos que a biblioteca de especificações (aba Especificações) tem."
    r.Validation.InputTitle = titulo
    r.Validation.InputMessage = "Escolha o ano. Muda as especificações, o ETp, o ET e o Sigma deste lote (fica na trilha de auditoria)."
    On Error GoTo 0
End Sub

' Sigma: >= 6 verde, 3 a 6 amarelo, < 3 vermelho. Por VALOR da celula (sem formula: formula de formatacao
' condicional e lida no idioma do Excel). Texto ("-") e vazio ficam fora: texto e "maior" que qualquer numero.
Private Sub CoresSigma(ByVal r As Range)
    Dim fc As Object
    r.FormatConditions.Delete
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="6", Formula2:="1000000")
    fc.SetLastPriority
    fc.Interior.Color = RGB(198, 239, 206): fc.Font.Color = RGB(0, 97, 0): fc.StopIfTrue = True
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="3", Formula2:="6")
    fc.SetLastPriority
    fc.Interior.Color = RGB(255, 235, 156): fc.Font.Color = RGB(156, 87, 0): fc.StopIfTrue = True
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="-1000000", Formula2:="3")
    fc.SetLastPriority
    fc.Interior.Color = RGB(255, 199, 206): fc.Font.Color = RGB(156, 0, 6): fc.StopIfTrue = True
End Sub

' ---------------------------------------------------------------- botoes
Private Sub BotoesBloco(ByVal ws As Worksheet, ByVal h As Long)
    Dim s As Shape, c As Range, nm As Variant
    For Each nm In Array("blkTgl_" & h, "blkCIQ_" & h)
        On Error Resume Next
        ws.Shapes(CStr(nm)).Delete
        On Error GoTo 0
    Next nm
    Set c = ws.Range(ws.Cells(h, 9), ws.Cells(h, 10))
    Set s = ws.Shapes.AddShape(5, c.Left + 3, c.Top + 3, c.Width - 6, ws.Rows(h).RowHeight - 6)
    s.Name = "blkTgl_" & h
    EstiloBotao s, RGB(31, 111, 107)
    s.OnAction = "mLotesBlocos.AlternarBloco"
    Set c = ws.Range(ws.Cells(h, 11), ws.Cells(h, 14))
    Set s = ws.Shapes.AddShape(5, c.Left + 3, c.Top + 3, c.Width - 6, ws.Rows(h).RowHeight - 6)
    s.Name = "blkCIQ_" & h
    EstiloBotao s, RGB(31, 90, 122)
    s.TextFrame2.TextRange.Text = "Preencher vazios pelo CIQ (20 primeiros)"
    s.OnAction = "mLotesBlocos.PreencherCIQDoBloco"
End Sub

Private Sub EstiloBotao(ByVal s As Shape, ByVal cor As Long)
    s.Fill.ForeColor.RGB = cor
    s.Line.Visible = msoFalse
    s.Placement = xlMove
    With s.TextFrame2
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

' ============================ ABRIR / RECOLHER ============================
Private Sub MostrarBloco(ByVal ws As Worksheet, ByVal h As Long, ByVal aberto As Boolean)
    Dim s As Shape
    ws.Range(ws.Rows(h + 1), ws.Rows(h + O_FIM)).Hidden = Not aberto
    On Error Resume Next
    Set s = ws.Shapes("blkTgl_" & h)
    On Error GoTo 0
    If Not s Is Nothing Then
        If aberto Then
            s.TextFrame2.TextRange.Text = ChrW(9660) & " recolher"
        Else
            s.TextFrame2.TextRange.Text = ChrW(9654) & " abrir"
        End If
    End If
End Sub

Private Function BlocoDoBotao(ByVal ws As Worksheet) As Long
    Dim nm As String
    On Error Resume Next
    nm = CStr(Application.Caller)
    If Len(nm) = 0 Then Exit Function
    BlocoDoBotao = CabecalhoDe(ws.Shapes(nm).TopLeftCell.Row)
    If Not EhCabecalho(ws, BlocoDoBotao) Then BlocoDoBotao = 0
End Function

' Botao de cada bloco: abre se estava recolhido, recolhe se estava aberto. So as linhas DESTE bloco.
Public Sub AlternarBloco()
    Dim ws As Worksheet, h As Long, prot As Boolean
    If Not BlocosAtivos() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    h = BlocoDoBotao(ws)
    If h = 0 Then Exit Sub
    prot = LiberarEscritaRapida(ws)
    AbrirBloco h, ws.Rows(h + 1).Hidden
    RestaurarProtecao ws, prot
End Sub

' Nucleo (testavel): abre (True) ou recolhe (False) o bloco que comeca na linha h.
Public Sub AbrirBloco(ByVal h As Long, ByVal aberto As Boolean)
    Dim ws As Worksheet, ev As Boolean, tela As Boolean
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    If Not EhCabecalho(ws, h) Then Exit Sub
    tela = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MostrarBloco ws, h, aberto
    Application.ScreenUpdating = tela
End Sub

Public Sub ExpandirTodos()
    TodosOsBlocos True
End Sub

Public Sub RecolherTodos()
    TodosOsBlocos False
End Sub

Private Sub TodosOsBlocos(ByVal aberto As Boolean)
    Dim ws As Worksheet, h As Long, prot As Boolean, tela As Boolean
    If Not BlocosAtivos() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    tela = Application.ScreenUpdating
    Application.ScreenUpdating = False
    prot = LiberarEscritaRapida(ws)
    h = BL_R0
    Do While EhCabecalho(ws, h)
        MostrarBloco ws, h, aberto
        h = h + BL_PASSO
    Loop
    RestaurarProtecao ws, prot
    Application.ScreenUpdating = tela
End Sub

' Abre o bloco do lote e leva o usuario ate ele (na linha do analito, se dado).
Public Sub IrBlocoDoLote(ByVal lote As String, Optional ByVal analito As String = "")
    Dim ws As Worksheet, h As Long, r As Long, i As Long, prot As Boolean
    If Not BlocosAtivos() Then mApp.Ir ESP_ABA, "A1": Exit Sub
    SincronizarBlocos
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    h = LinhaDoBloco(lote)
    If h = 0 Then mApp.Ir BL_ABA, "A1": Exit Sub
    prot = LiberarEscritaRapida(ws)
    MostrarBloco ws, h, True
    RestaurarProtecao ws, prot
    r = h + O_AN
    If Len(analito) > 0 Then
        For i = 0 To BL_NAN - 1
            If StrComp(Trim$(CStr(ws.Cells(h + O_AN + i, 1).Value)), Trim$(analito), vbTextCompare) = 0 Then r = h + O_AN + i: Exit For
        Next i
    End If
    mApp.Ir BL_ABA, "A1"
    On Error Resume Next
    ActiveWindow.ScrollRow = h
    ws.Cells(r, C_MED1).Select
End Sub

Public Sub IrLoteEmAnalise()
    IrBlocoDoLote mLotes.LotePainel()
End Sub

Public Sub IrEspecificacoes()
    mApp.Ir mLotes.AbaEspec(), "A1"
End Sub

' ============================ EDICAO NO BLOCO ============================
' Worksheet_Change da aba Analitos: Media/DP (E:J das linhas de analito) e anos de referencia.
Public Sub BlocoEditado(ByVal Target As Range)
    Dim ws As Worksheet, c As Range, h As Long, off As Long, params As Collection, anos As Collection
    If mDentro Then Exit Sub
    If Not BlocosAtivos() Then Exit Sub
    Set ws = Target.Worksheet
    If Target.Cells.CountLarge > 4000 Then Exit Sub
    Set params = New Collection
    Set anos = New Collection
    For Each c In Target.Cells
        h = CabecalhoDe(c.Row)
        If h > 0 Then
            If EhCabecalho(ws, h) Then
                off = c.Row - h
                If off >= O_AN And off < O_AN + BL_NAN And c.Column >= C_MED1 And c.Column <= C_MED1 + 5 Then
                    params.Add c
                ElseIf off = O_ANO And (c.Column = C_ANOCLIA Or c.Column = C_ANOVB) Then
                    anos.Add c
                End If
            End If
        End If
    Next c
    If params.Count > 0 Then GravarEditados ws, params
    If anos.Count > 0 Then AnosEditados ws, anos
End Sub

' Retrato da especificacao do analito NO BLOCO do lote: (fonte, "CLIA aaaa · VB aaaa", ETp, CVTp).
' Empty se o lote nao tem bloco ou o analito nao esta nele.
Public Function EspecDoLote(ByVal lote As String, ByVal analito As String) As Variant
    Dim ws As Worksheet, h As Long, i As Long
    h = LinhaDoBloco(lote)
    If h = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    For i = 0 To BL_NAN - 1
        If StrComp(Trim$(CStr(ws.Cells(h + O_AN + i, 1).Value)), Trim$(analito), vbTextCompare) = 0 Then
            EspecDoLote = EspecDaLinhaBloco(ws, h, h + O_AN + i)
            Exit Function
        End If
    Next i
End Function

Private Function EspecDaLinhaBloco(ByVal ws As Worksheet, ByVal h As Long, ByVal r As Long) As Variant
    Dim fonte As String, etp As Variant, cv As Variant
    fonte = Trim$(CStr(ws.Cells(r, 19).Value))
    etp = ws.Cells(r, 20).Value
    cv = ws.Cells(r, 21).Value
    If IsError(etp) Then etp = "-"
    If IsError(cv) Then cv = "-"
    If Not IsNumeric(etp) Or VarType(etp) = vbString Then etp = "-"
    If Not IsNumeric(cv) Or VarType(cv) = vbString Then cv = "-"
    EspecDaLinhaBloco = Array(IIf(fonte = "", "-", fonte), RotuloAnos(ws, h), etp, cv)
End Function

Private Function RotuloAnos(ByVal ws As Worksheet, ByVal h As Long) As String
    RotuloAnos = "CLIA " & CStr(ws.Cells(h + O_ANO, C_ANOCLIA).Value) & " · VB " & CStr(ws.Cells(h + O_ANO, C_ANOVB).Value)
End Function

Private Function NumeroOk(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbString, vbBoolean: Exit Function
    End Select
    NumeroOk = IsNumeric(v)
End Function

Private Sub GravarEditados(ByVal ws As Worksheet, ByVal cels As Collection)
    Dim c As Variant, h As Long, lote As String, an As String, j As Long, nv As Long, campo As String
    Dim v As Variant, antes As Variant, invalidos As String, prot As Boolean, lotes As Object, loteP As String
    Dim nE As Long, sE As String
    mApp.InicioUsuario "Gravando Média/DP..."
    mDentro = True
    Set lotes = CreateObject("Scripting.Dictionary")
    lotes.CompareMode = 1
    On Error GoTo falha
    prot = LiberarEscrita(ws)
    For Each c In cels
        h = CabecalhoDe(c.Row)
        lote = Trim$(CStr(ws.Cells(h, 2).Value))
        an = Trim$(CStr(ws.Cells(c.Row, 1).Value))
        j = c.Column - C_MED1 + 1                                   ' 1..6
        nv = (j - 1) \ 2 + 1
        If (j Mod 2) = 1 Then campo = "Media" Else campo = "DP"
        If lote = "" Or an = "" Then c.ClearContents: GoTo proximo       ' linha sem analito: nada a gravar
        If c.HasFormula Then GoTo volta                                    ' formula colada: volta o valor gravado
        v = c.Value
        If Len(Trim$(CStr(v))) = 0 Then
            v = Empty
        ElseIf Not NumeroOk(v) Then
            invalidos = invalidos & vbCrLf & "  " & an & " " & campo & " N" & nv & ": não numérico (não gravado)"
            GoTo volta
        ElseIf campo = "DP" And CDbl(v) <= 0 Then
            invalidos = invalidos & vbCrLf & "  " & an & " DP N" & nv & ": precisa ser maior que zero (não gravado)"
            GoTo volta
        End If
        antes = mLotes.GravarParametroStore(lote, an, j, v, EspecDaLinhaBloco(ws, h, c.Row), _
                                            "digitado na aba Analitos (bloco do lote)")
        If CStr(antes) <> CStr(v) Then
            mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_ALTERADO", "mLotesBlocos", _
                0, Empty, "", lote, nv, an, antes, v, "", "", campo & " N" & nv, _
                "Parametro " & campo & " N" & nv & " do lote " & lote & " alterado na aba Analitos (bloco do lote)."
            lotes(lote) = 1
        End If
        GoTo proximo
volta:
        c.Value = mLotes.ValorDoLote(lote, an, j)                         ' invalido: volta o que estava gravado
proximo:
    Next c
    RestaurarProtecao ws, prot
    mLotes.InvalidarLotes
    loteP = mLotes.LoteEmTela()
    If lotes.Exists(loteP) Then mLotes.CarregarParametrosView loteP      ' a tela do motor acompanha a store
    If lotes.Count > 0 Then
        mEstatistica.InvalidarParametros
        mEstatistica.RecalcularAnalitoAtual
    End If
    mDentro = False
    mApp.fim
    If Len(invalidos) > 0 Then Avisar "Alguns valores NÃO foram gravados:" & invalidos, vbExclamation
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    RestaurarProtecao ws, prot
    mDentro = False
    mApp.fim
    Avisar "Falha ao gravar Média/DP do lote " & lote & ": " & sE, vbCritical
End Sub

Private Sub AnosEditados(ByVal ws As Worksheet, ByVal cels As Collection)
    Dim c As Variant, h As Long, y As Long, lote As String, campo As String, cs As Long
    Dim v As Variant, antes As Variant, prot As Boolean, lotes As Object, k As Variant, ruins As String
    mApp.InicioUsuario "Gravando ano de referência..."
    mDentro = True
    Set lotes = CreateObject("Scripting.Dictionary")
    lotes.CompareMode = 1
    On Error GoTo falha
    prot = LiberarEscrita(ws)
    For Each c In cels
        h = CabecalhoDe(c.Row): y = h + O_ANO
        lote = Trim$(CStr(ws.Cells(h, 2).Value))
        If c.Column = C_ANOCLIA Then campo = "CLIA": cs = C_SOMB1 Else campo = "VB": cs = C_SOMB2
        v = ws.Cells(y, c.Column).Value
        antes = ws.Cells(y, cs).Value
        If Not AnoNaBiblioteca(v) Then
            ws.Cells(y, c.Column).Value = antes
            ruins = ruins & vbCrLf & "  lote " & lote & " " & campo & ": '" & CStr(v) & "' não está na biblioteca (aba Especificações)"
        ElseIf CStr(v) <> CStr(antes) Then
            ws.Cells(y, c.Column).Value = CLng(v)
            ws.Cells(y, cs).Value = CLng(v)
            mAuditoria.Auditar mAuditoria.CAT_CONFIG, "ANO_REFERENCIA_LOTE_ALTERADO", "mLotesBlocos", _
                0, Empty, "", lote, 0, "", antes, CLng(v), "", "", "Ano " & campo, _
                "Ano de referência " & campo & " do lote " & lote & ": " & CStr(antes) & " -> " & CStr(v) & _
                " (especificações, ETp, ET e Sigma deste lote passam a usar esta referência)."
            lotes(lote) = 1
        End If
    Next c
    RestaurarProtecao ws, prot
    If lotes.Count > 0 Then
        Application.Calculate
        For Each k In lotes.Keys
            RetratoDoLote ws, CStr(k)
        Next k
        mEstatistica.InvalidarParametros
        mEstatistica.RecalcularAnalitoAtual
    End If
    mDentro = False
    mApp.fim
    If Len(ruins) > 0 Then Avisar "Ano não aceito:" & ruins, vbExclamation
    Exit Sub
falha:
    Dim sE As String
    sE = Err.Description
    On Error Resume Next
    RestaurarProtecao ws, prot
    mDentro = False
    mApp.fim
    Avisar "Falha ao gravar o ano de referência: " & sE, vbCritical
End Sub

' A referencia do lote mudou: o retrato (fonte, anos, ETp, CVTp) das linhas da store que tem Media/DP
' passa a ser o do bloco. Quem e quando dos Media/DP nao mudam; a mudanca de referencia esta na trilha.
Private Sub RetratoDoLote(ByVal ws As Worksheet, ByVal lote As String)
    Dim h As Long, i As Long, an As String
    h = LinhaDoBloco(lote)
    If h = 0 Then Exit Sub
    For i = 0 To BL_NAN - 1
        an = Trim$(CStr(ws.Cells(h + O_AN + i, 1).Value))
        If Len(an) > 0 Then mLotes.AtualizarRetratoStore lote, an, EspecDaLinhaBloco(ws, h, h + O_AN + i)
    Next i
End Sub

' ============================ CIQ POR BLOCO ============================
' ADR-078: Media/DP DIGITADA no bloco que nao chegou a LotesStore (eventos desligados na hora). Antes de reescrever
' E:J a partir da store (MontarBlocos), cada CONSTANTE do bloco que difere da store e tratada assim:
'   store VAZIA no campo ............ grava (com trilha) -- RECUPERADO;
'   store com OUTRO numero .......... NAO grava (pode ser um valor mais novo, digitado em outro lugar) -- CONFLITO;
'   texto, DP <= 0 .................. NAO grava -- IGNORADO.
' Layout do ADR-074 (E:J com formula): a constante e digitacao. Layout do ADR-076 (E:J valores): toda gravacao
' reescreve o bloco, entao a diferenca tambem so vem de digitacao que nao chegou a store.
' Devolve "OK|<recuperados>|<conflitos>|<ignorados>" + uma linha "lote;analito;campo;na_store;no_bloco;situacao" por
' caso, ou "ERRO|<motivo>".
Public Function ResgatarDigitados() As String
    Dim ws As Worksheet, d As Object, prox As Long, k As Variant, h As Long, rg As Range, f As Variant, v As Variant
    Dim i As Long, j As Long, lote As String, an As String, antes As Variant, n As Long, nc As Long, ni As Long
    Dim det As String, igual As Boolean, campo As String, nv As Long, sit As String, tocados As Object
    If Not BlocosAtivos() Then ResgatarDigitados = "OK|0|0|0": Exit Function
    On Error GoTo falha
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    Set d = BlocosExistentes(ws, prox)
    Set tocados = CreateObject("Scripting.Dictionary")
    tocados.CompareMode = 1
    For Each k In d.Keys
        h = d(k)
        lote = Trim$(CStr(ws.Cells(h, 2).Value))
        If lote <> "" Then
            Set rg = ws.Range(ws.Cells(h + O_AN, C_MED1), ws.Cells(h + O_AN + BL_NAN - 1, C_MED1 + 5))
            f = rg.Formula
            v = rg.Value
            For i = 1 To BL_NAN
                an = Trim$(CStr(ws.Cells(h + O_AN + i - 1, 1).Value))
                If an <> "" Then
                    For j = 1 To 6
                        If Left$(CStr(f(i, j)), 1) <> "=" And Len(Trim$(CStr(f(i, j)))) > 0 Then
                            nv = (j - 1) \ 2 + 1
                            If (j Mod 2) = 1 Then campo = "Media" Else campo = "DP"
                            antes = mLotes.ValorDoLote(lote, an, j)
                            sit = ""
                            If Not NumeroOk(v(i, j)) Then
                                sit = "IGNORADO: nao numerico"
                            ElseIf (j Mod 2) = 0 And CDbl(v(i, j)) <= 0 Then
                                sit = "IGNORADO: DP <= 0"
                            Else
                                igual = False
                                If NumeroOk(antes) Then igual = (Abs(CDbl(antes) - CDbl(v(i, j))) <= 0.000000001 * IIf(Abs(CDbl(antes)) > 1, Abs(CDbl(antes)), 1))
                                If Not igual Then
                                    If NumeroOk(antes) Then
                                        sit = "CONFLITO: a store tem outro valor (nao gravado; conferir)"
                                    Else
                                        sit = "RECUPERADO"
                                        mLotes.GravarParametroStore lote, an, j, CDbl(v(i, j)), EspecDaLinhaBloco(ws, h, h + O_AN + i - 1), _
                                            "recuperado na instalação (ADR-078): digitado no bloco com os eventos desligados"
                                        mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_RECUPERADO", "mLotesBlocos", _
                                            0, Empty, "", lote, nv, an, antes, CDbl(v(i, j)), "", "", campo & " N" & nv, _
                                            "Valor digitado no bloco do lote " & lote & " que nao tinha chegado a LotesStore (eventos " & _
                                            "desligados na hora); gravado na instalacao do ADR-078."
                                        tocados(lote) = 1
                                    End If
                                End If
                            End If
                            If Len(sit) > 0 Then
                                Select Case Left$(sit, 3)
                                    Case "REC": n = n + 1
                                    Case "CON": nc = nc + 1
                                    Case Else: ni = ni + 1
                                End Select
                                det = det & vbLf & lote & ";" & an & ";" & campo & " N" & nv & ";" & CStr(antes) & ";" & _
                                      CStr(v(i, j)) & ";" & sit
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    Next k
    mLotes.InvalidarLotes
    If tocados.Exists(mLotes.LoteEmTela()) Then mLotes.CarregarParametrosView mLotes.LoteEmTela()   ' a tela do motor acompanha
    ResgatarDigitados = "OK|" & n & "|" & nc & "|" & ni & det
    Exit Function
falha:
    ResgatarDigitados = "ERRO|" & Err.Description & " (bloco do lote " & lote & ")"
End Function

Public Sub PreencherCIQDoBloco()
    Dim ws As Worksheet, h As Long, lote As String, r As String, p() As String, prot As Boolean
    If Not BlocosAtivos() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_ABA)
    h = BlocoDoBotao(ws)
    If h = 0 Then Exit Sub
    lote = Trim$(CStr(ws.Cells(h, 2).Value))
    If MsgBox("Preencher os Média/DP VAZIOS do lote " & lote & " com a média e o DP dos " & N_CIQ & _
              " primeiros resultados dele no CIQ?" & vbCrLf & vbCrLf & _
              "O que já foi digitado NÃO é alterado. Revise os valores antes de usar.", _
              vbQuestion + vbYesNo, "Média/DP pelo CIQ") <> vbYes Then Exit Sub
    mApp.InicioUsuario "Calculando Média/DP do lote " & lote & "..."
    r = mLotes.PreencherVaziosCIQLote(lote)
    mApp.fim
    prot = LiberarEscritaRapida(ws)
    MostrarBloco ws, h, True
    RestaurarProtecao ws, prot
    If Left$(r, 5) = "ERRO|" Then
        Avisar "Falha ao calcular Média/DP do lote " & lote & ": " & Mid$(r, 6), vbCritical
        Exit Sub
    End If
    p = Split(r, "|")
    Avisar p(1) & " par(es) analito/nível preenchido(s) no lote " & lote & "." & _
           IIf(CLng(p(2)) > 0, vbCrLf & p(2) & " par(es) ficaram vazios: menos de " & N_CIQ & " resultados.", "") & _
           vbCrLf & "Revise os valores no bloco do lote.", vbInformation
End Sub

Private Sub Avisar(ByVal msg As String, ByVal estilo As VbMsgBoxStyle)
    If Application.Interactive Then MsgBox msg, estilo, "Lotes"
End Sub
