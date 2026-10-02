Public Sub AtualizarCalc()
    Dim ws As Worksheet, analito As String, lote As String, equip As String
    Dim i As Long, t As Long, nRun As Long, r As Long, e As Long
    Dim runs() As Long, dts() As Double, valor() As Double, temDado() As Boolean
    Dim r13 As Variant, r2sMulti As Variant, rR4 As Variant, r1sMulti As Variant, rSeq As Variant, a12 As Variant
    Dim alvoM() As Double, alvoS() As Double, etp As Double
    Dim seen As Object, ordem As Object
    Dim protEstava As Boolean

    Set ws = ThisWorkbook.Sheets("Eng_Saida")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    GarantirIndice
    CarregarFiltro
    ' ADR-049: o lote e o do PAINEL (loteSel), escolhido pelo usuario.
    lote = mLotes.LotePainel()
    ' ADR-057: na Bioquimica, o equipamento tambem (cada um e uma serie propria)
    equip = mEquipIdx

    ' Eng_Saida e aba tecnica e protegida. Ver o cabecalho de LiberarEscrita.
    On Error GoTo restaura
    protEstava = LiberarEscrita(ws)

    ' limpa a area de saida (colunas B em diante; a coluna A guarda os slots
    ' fixos). Vai ate o bloco do X vermelho (ADR-057): sem isso, o X do analito
    ' anterior ficaria na tela.
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + NK - 1, COL_X0 + NX - 1)).ClearContents
    ws.Range("C1").Value = analito
    ws.Range("E1").NumberFormat = "@"       ' lote "010" nao pode virar 10
    ws.Range("E1").Value = lote
    ws.Range("G1").Value = Now
    ws.Range("I1").Value = 0
    ws.Range("K1").Value = 0
    ws.Range("M1").Value = 0
    ws.Range("O1").Value = ""
    ws.Range("Q1").Value = ""
    ws.Range("R1").Value = "equip:"
    ws.Range("S1").Value = equip
    ws.Range("T1").Value = "X além de " & NXN & ":"
    ws.Range("U1").Value = 0
    If analito = "" Or IsEmpty(mDB) Then GoTo restaura
    ws.Range("O1").Value = LotesNoPeriodo(analito)

    ' ---- descobrir as corridas do analito no lote ----
    ' DOIS conjuntos de linhas (ADR-057):
    '   linhas  = resultados ELEGIVEIS (PARTICIPA_ESTATISTICA = SIM): ponto,
    '             z, Westgard, media/DP/CV;
    '   linhasX = resultados INATIVADOS com REGISTRAR - LJ marcado
    '             (TIPO_PLOTAGEM_LJ = X_VERMELHO): so o X no grafico.
    ' O eixo de PLOTAGEM e a uniao das corridas dos dois (uma corrida em que o
    ' unico resultado do analito foi inativado ainda ganha posicao, com o X).
    ' O eixo do WESTGARD e so o das corridas com resultado elegivel -- o mesmo
    ' de antes da inativacao existir: o X nunca entra em regra, media ou DP.
    ' ADR-049: sem teto na descoberta. ADR-050: pelo indice.
    Dim linhas As Collection, linhasX As Collection, x As Variant
    Set linhas = LinhasDe(analito, lote)
    Set linhasX = LinhasXDe(analito, lote)    ' LinhasXDe, nao LinhasX: VBA nao diferencia maiusculas e a variavel esconderia a funcao
    Set seen = CreateObject("Scripting.Dictionary")
    Dim elegivel As Object: Set elegivel = CreateObject("Scripting.Dictionary")
    Dim nTodas As Long, runsT() As Long, dtsT() As Double, dd As Double, mxDe As Object, passada As Long
    ReDim runsT(1 To linhas.Count + linhasX.Count + 1): ReDim dtsT(1 To linhas.Count + linhasX.Count + 1)
    Set mxDe = CreateObject("Scripting.Dictionary")     ' RUN -> maior data/hora
    Dim fonte As Collection
    For passada = 1 To 2
        If passada = 1 Then Set fonte = linhas Else Set fonte = linhasX
        For Each x In fonte
            i = x
            dd = 0
            If IsDate(mDB(i, COL_DATA)) Then dd = CDbl(CDate(mDB(i, COL_DATA)))
            If passada = 1 Then elegivel(CStr(mDB(i, COL_RUN))) = True
            If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                nTodas = nTodas + 1
                runsT(nTodas) = CLng(Val(mDB(i, COL_RUN)))
                dtsT(nTodas) = dd
                seen.Add CStr(mDB(i, COL_RUN)), nTodas
                mxDe(CStr(mDB(i, COL_RUN))) = dd
            ElseIf dd > mxDe(CStr(mDB(i, COL_RUN))) Then
                ' data EXIBIDA = a maior data/hora da corrida no analito; a ORDEM
                ' continua pela data da 1a linha (ADR-049)
                mxDe(CStr(mDB(i, COL_RUN))) = dd
            End If
        Next x
    Next passada
    ws.Range("K1").Value = nTodas
    If nTodas = 0 Then GoTo restaura
    OrdenarCorridas runsT, dtsT, nTodas

    ' ---- janela: as NK corridas mais recentes ate o fim do periodo ----
    Dim fimJan As Long, iniJan As Long, nCort As Long, vAte As Variant
    fimJan = nTodas
    On Error Resume Next
    vAte = ThisWorkbook.Names("filtroAte").RefersToRange.Value
    On Error GoTo restaura
    If IsDate(vAte) Then
        Do While fimJan >= 1
            If Int(dtsT(fimJan)) <= Int(CDbl(CDate(vAte))) Then Exit Do
            fimJan = fimJan - 1
        Loop
    End If
    If fimJan < 1 Then GoTo restaura
    iniJan = fimJan - NK + 1
    If iniJan < 1 Then iniJan = 1
    For i = 1 To iniJan - 1
        If PassaFiltro(dtsT(i)) Then nCort = nCort + 1
    Next i
    ws.Range("M1").Value = nCort
    nRun = fimJan - iniJan + 1
    ReDim runs(1 To NK): ReDim dts(1 To NK)
    For i = 1 To nRun
        runs(i) = runsT(iniJan + i - 1)
        dts(i) = dtsT(iniJan + i - 1)
    Next i

    Set ordem = CreateObject("Scripting.Dictionary")
    For i = 1 To nRun
        ordem(CStr(runs(i))) = i
    Next i

    ' ---- coletar valores ELEGIVEIS por nivel (eixo de plotagem) ----
    ReDim valor(0 To NLV - 1, 1 To nRun)
    ReDim temDado(0 To NLV - 1, 1 To nRun)
    For Each x In linhas
        i = x
        If IsNumeric(mDB(i, COL_RESULT)) Then
            If ordem.Exists(CStr(mDB(i, COL_RUN))) Then
                t = CLng(Val(mDB(i, COL_NIVEL))) - 1
                If t >= 0 And t <= NLV - 1 Then
                    r = ordem(CStr(mDB(i, COL_RUN)))
                    valor(t, r) = CDbl(mDB(i, COL_RESULT))
                    temDado(t, r) = True
                End If
            End If
        End If
    Next x

    ' ---- coletar os X vermelhos (ate NXN por corrida e nivel) ----
    Dim outX() As Variant, s As Long, nExc As Long
    ReDim outX(1 To nRun, 1 To NX)
    For Each x In linhasX
        i = x
        If IsNumeric(mDB(i, COL_RESULT)) Then
            If ordem.Exists(CStr(mDB(i, COL_RUN))) Then
                t = CLng(Val(mDB(i, COL_NIVEL))) - 1
                If t >= 0 And t <= NLV - 1 Then
                    r = ordem(CStr(mDB(i, COL_RUN)))
                    For s = 1 To NXN
                        If IsEmpty(outX(r, t * NXN + s)) Then outX(r, t * NXN + s) = CDbl(mDB(i, COL_RESULT)): Exit For
                    Next s
                    If s > NXN Then nExc = nExc + 1
                End If
            End If
        End If
    Next x
    ws.Range("U1").Value = nExc

    ' ---- eixo do Westgard: so as corridas com resultado elegivel ----
    ' A sequencia avaliada e identica a de antes da inativacao existir: as
    ' corridas so-X ficam de fora, e o resultado inativado nao entra em z.
    Dim pDeE() As Long, nE As Long
    ReDim pDeE(1 To nRun)
    For i = 1 To nRun
        If elegivel.Exists(CStr(runs(i))) Then nE = nE + 1: pDeE(nE) = i
    Next i

    ' ---- alvos e z-scores (no eixo compactado) ----
    ' Nivel sem media/DP do lote NAO entra no Westgard: z = 0 seria lido como
    ' "no alvo" e o grafico sairia verde sobre um lote nunca configurado.
    Dim parOK() As Boolean, temAval() As Boolean, stPar As String
    Dim zc() As Double, tac() As Boolean
    ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1): ReDim parOK(0 To NLV - 1)
    ReDim temAval(0 To NLV - 1, 1 To nRun)
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r2sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r1sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rSeq(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)
    For t = 0 To NLV - 1
        parOK(t) = AlvoAnalito(analito, t + 1, alvoM(t), alvoS(t), etp, lote)
        stPar = stPar & IIf(t > 0, ";", "") & "N" & (t + 1) & "=" & IIf(parOK(t), "OK", "SEM")
    Next t
    ws.Range("Q1").Value = stPar
    If nE > 0 Then
        ReDim zc(0 To NLV - 1, 1 To nE): ReDim tac(0 To NLV - 1, 1 To nE)
        For t = 0 To NLV - 1
            For e = 1 To nE
                i = pDeE(e)
                If temDado(t, i) And parOK(t) Then
                    zc(t, e) = CalcularZ(valor(t, i), alvoM(t), alvoS(t))
                    tac(t, e) = True
                    temAval(t, i) = True
                End If
            Next e
        Next t

        ' ---- Westgard ----
        Dim c13 As Variant, c2s As Variant, cR4 As Variant, c1s As Variant, cSeq As Variant, c12 As Variant
        ReDim c13(0 To NLV - 1, 1 To nE): ReDim c2s(0 To NLV - 1, 1 To nE)
        ReDim cR4(0 To NLV - 1, 1 To nE): ReDim c1s(0 To NLV - 1, 1 To nE)
        ReDim cSeq(0 To NLV - 1, 1 To nE): ReDim c12(0 To NLV - 1, 1 To nE)
        AvaliarWestgard zc, tac, nE, c13, c2s, cR4, c1s, cSeq, c12
        ' devolve as marcas ao eixo de plotagem
        For t = 0 To NLV - 1
            For e = 1 To nE
                i = pDeE(e)
                r13(t, i) = c13(t, e): r2sMulti(t, i) = c2s(t, e): rR4(t, i) = cR4(t, e)
                r1sMulti(t, i) = c1s(t, e): rSeq(t, i) = cSeq(t, e): a12(t, i) = c12(t, e)
            Next e
        Next t
    End If

    ' ---- publicar em Eng_Saida ----
    ' Sem filtro de data de proposito: as regras dependem da sequencia completa
    ' de corridas. Quem recorta o periodo exibido e o Calc, na hora de plotar.
    Dim outRun() As Variant, outLvl() As Variant, rej As Boolean
    ReDim outRun(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outRun(i, 1) = runs(i)
    Next i
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + nRun - 1, 2)).Value = outRun

    Dim outFil() As Variant, outVal() As Variant, outChv() As Variant, outDt() As Variant
    ReDim outFil(1 To nRun, 1 To 1)
    ReDim outVal(1 To nRun, 1 To NLV)
    ReDim outChv(1 To nRun, 1 To 1)
    ReDim outDt(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outFil(i, 1) = IIf(PassaFiltro(dts(i)), 1, 0)
        Dim dMx As Double
        dMx = mxDe(CStr(runs(i)))
        If dMx > 0 Then outDt(i, 1) = CDate(dMx) Else outDt(i, 1) = ""
        ' Chave logica da linha: ANALITO|RUN (o RUN e compartilhado entre analitos)
        outChv(i, 1) = analito & "|" & runs(i)
        For t = 0 To NLV - 1
            If temDado(t, i) Then outVal(i, t + 1) = valor(t, i) Else outVal(i, t + 1) = ""
        Next t
    Next i
    ws.Range(ws.Cells(KC0, COL_FILTRO), ws.Cells(KC0 + nRun - 1, COL_FILTRO)).Value = outFil
    ws.Range(ws.Cells(KC0, COL_VALOR0), ws.Cells(KC0 + nRun - 1, COL_VALOR0 + NLV - 1)).Value = outVal
    ws.Range(ws.Cells(KC0, COL_CHAVE), ws.Cells(KC0 + nRun - 1, COL_CHAVE)).Value = outChv
    ws.Range(ws.Cells(KC0, COL_DATA_ENG), ws.Cells(KC0 + nRun - 1, COL_DATA_ENG)).Value = outDt
    ' X vermelho: bloco proprio, FORA das colunas de valor (que o Painel soma em n/media/DP)
    ws.Range(ws.Cells(KC0, COL_X0), ws.Cells(KC0 + nRun - 1, COL_X0 + NX - 1)).Value = outX

    For t = 0 To NLV - 1
        ReDim outLvl(1 To nRun, 1 To NEF)
        For i = 1 To nRun
            If temDado(t, i) And Not temAval(t, i) Then
                ' resultado existe, parametro do lote nao: sem veredicto
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = "SEM PARAMETROS"
            ElseIf temDado(t, i) Then
                outLvl(i, 1) = IIf(r13(t, i) = 1, 1, 0)
                outLvl(i, 2) = IIf(r2sMulti(t, i) = 1, 1, 0)
                outLvl(i, 3) = IIf(rR4(t, i) = 1, 1, 0)
                outLvl(i, 4) = IIf(r1sMulti(t, i) = 1, 1, 0)
                outLvl(i, 5) = IIf(rSeq(t, i) = 1, 1, 0)
                outLvl(i, 6) = IIf(a12(t, i) = 1, 1, 0)
                rej = (r13(t, i) = 1 Or r2sMulti(t, i) = 1 Or rR4(t, i) = 1 Or r1sMulti(t, i) = 1 Or rSeq(t, i) = 1)
                ' So REJEITADO/OK. O alerta 12s vai na coluna propria.
                outLvl(i, 7) = IIf(rej, "REJEITADO", "OK")
            Else
                ' corrida so com X (ou nivel sem resultado): sem ponto, sem veredicto
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = ""
            End If
        Next i
        ws.Range(ws.Cells(KC0, EF0 + t * NEF), ws.Cells(KC0 + nRun - 1, EF0 + t * NEF + NEF - 1)).Value = outLvl
    Next t

    ws.Range("I1").Value = nRun
restaura:
    ' Restauro garantido: a saida antecipada e o erro passam por aqui.
    Dim nErrP As Long, sErrP As String
    nErrP = Err.Number: sErrP = Err.Description
    RestaurarProtecao ws, protEstava
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mEstatistica.AtualizarCalc", sErrP
End Sub

