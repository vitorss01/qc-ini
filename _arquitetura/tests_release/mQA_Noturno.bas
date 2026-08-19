Attribute VB_Name = "mQA_Noturno"
Option Explicit

Private Function PrimeiroRegistroNumerico(ByRef lin As Long) As Boolean
    Dim ws As Worksheet, ult As Long, i As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    ult = UltimaLinhaBanco()
    For i = BANCO_R0 To ult
        If IsNumeric(ws.Cells(i, COL_RESULT).Value) And _
           Len(Trim$(CStr(ws.Cells(i, COL_ANALITO).Value))) > 0 Then
            lin = i
            PrimeiroRegistroNumerico = True
            Exit Function
        End If
    Next i
End Function

' Pre-flight deve recusar chave repetida sem aplicar a primeira alteracao.
Public Function QA_T03_DuplicataAtomica() As String
    Dim ws As Worksheet, lin As Long, regs(1 To 2, 1 To 7) As Variant
    Dim original As Variant, i As Long, c As Long, houveErro As Boolean
    On Error GoTo falha
    Set ws = ThisWorkbook.Sheets(BANCO)
    If Not PrimeiroRegistroNumerico(lin) Then QA_T03_DuplicataAtomica = "SKIP|sem registro numerico": Exit Function
    original = ws.Cells(lin, COL_RESULT).Value
    For i = 1 To 2
        For c = 1 To 7: regs(i, c) = ws.Cells(lin, c).Value: Next c
        regs(i, COL_RESULT) = CDbl(original) + i
    Next i
    On Error Resume Next
    UpsertResultados regs
    houveErro = (Err.Number <> 0)
    Err.Clear
    On Error GoTo falha
    If houveErro And Mesmo(original, ws.Cells(lin, COL_RESULT).Value) Then
        QA_T03_DuplicataAtomica = "PASS|duplicata recusada; valor original preservado"
    Else
        QA_T03_DuplicataAtomica = "FAIL|houveErro=" & houveErro & "; antes=" & original & "; depois=" & ws.Cells(lin, COL_RESULT).Value
    End If
    Exit Function
falha:
    QA_T03_DuplicataAtomica = "FAIL|" & Err.Description
End Function

' Registro excluido logicamente nao pode ser reativado por upsert.
Public Function QA_T03_ReativacaoBloqueada() As String
    Dim an As String, core As String, codigo As String, run As Long
    Dim regs(1 To 1, 1 To 7) As Variant, alvos As Object, r As String
    Dim ws As Worksheet, ult As Long
    On Error GoTo falha
    an = CStr(ThisWorkbook.Sheets("Analitos").Cells(4, 1).Value)
    core = "QAR001"
    codigo = CodigoLote(core, 1)
    run = ObterOuCriarRUN(DateSerial(2098, 1, 1), core, "QA", 1)
    regs(1, 1) = run: regs(1, 2) = DateSerial(2098, 1, 1): regs(1, 3) = 1
    regs(1, 4) = codigo: regs(1, 5) = an: regs(1, 6) = 1.2345: regs(1, 7) = ST_ATIVO
    r = UpsertResultados(regs)
    Set alvos = CreateObject("Scripting.Dictionary")
    alvos.Add UCase$(an), True
    If ExcluirLogico(run, 1, alvos, "Teste QA de exclusao logica controlada", ST_EXCLUIDO, "QA", "QA") <> 1 Then
        QA_T03_ReativacaoBloqueada = "FAIL|exclusao nao aplicada": Exit Function
    End If
    regs(1, 6) = 9.8765
    r = UpsertResultados(regs)
    Set ws = ThisWorkbook.Sheets(BANCO)
    ult = UltimaLinhaBanco()
    If r = "0|0|1" And CStr(ws.Cells(ult, COL_STATUS).Value) = ST_EXCLUIDO Then
        QA_T03_ReativacaoBloqueada = "PASS|" & r
    Else
        QA_T03_ReativacaoBloqueada = "FAIL|retorno=" & r & "; status=" & ws.Cells(ult, COL_STATUS).Value
    End If
    Exit Function
falha:
    QA_T03_ReativacaoBloqueada = "FAIL|" & Err.Description
End Function

Public Function QA_T04_Capacidade() As String
    Dim antes As Long, depois As Long, r As String
    antes = UltimaLinhaBanco()
    r = TestarCapacidade(LinhasLivres() + 1)
    depois = UltimaLinhaBanco()
    If Left$(r, 7) = "RECUSOU" And antes = depois Then
        QA_T04_Capacidade = "PASS|barreira recusou sem alterar o banco"
    Else
        QA_T04_Capacidade = "FAIL|" & r & "; antes=" & antes & "; depois=" & depois
    End If
End Function

' Duas corridas no mesmo dia/lote nao colidem; niveis diferentes da segunda
' corrida continuam agrupados no mesmo RUN.
Public Function QA_T08_RunMesmoDia() As String
    Dim an As String, core As String, r1 As Long, r2 As Long, r2n2 As Long
    Dim regs(1 To 1, 1 To 7) As Variant
    On Error GoTo falha
    an = CStr(ThisWorkbook.Sheets("Analitos").Cells(4, 1).Value)
    core = "QAR002"
    r1 = ObterOuCriarRUN(DateSerial(2098, 2, 2), core, "QA", 1)
    regs(1, 1) = r1: regs(1, 2) = DateSerial(2098, 2, 2): regs(1, 3) = 1
    regs(1, 4) = CodigoLote(core, 1): regs(1, 5) = an: regs(1, 6) = 2.3456: regs(1, 7) = ST_ATIVO
    UpsertResultados regs
    r2 = ObterOuCriarRUN(DateSerial(2098, 2, 2), core, "QA", 1)
    r2n2 = ObterOuCriarRUN(DateSerial(2098, 2, 2), core, "QA", 2)
    If r2 <> r1 And r2n2 = r2 Then
        QA_T08_RunMesmoDia = "PASS|RUN1=" & r1 & "; RUN2=" & r2 & "; N2=" & r2n2
    Else
        QA_T08_RunMesmoDia = "FAIL|RUN1=" & r1 & "; RUN2=" & r2 & "; N2=" & r2n2
    End If
    Exit Function
falha:
    QA_T08_RunMesmoDia = "FAIL|" & Err.Description
End Function

' Prepara um RUN exclusivamente sintetico para provar a finalizacao positiva.
' O alvo vem do contrato BI ja materializado no clone; nenhum RUN existente e
' selecionado ou assinado. O chamador deve descartar o clone ao fim da suite.
Public Function QA_T07_PrepararRunPositivo() As String
    Const QA_DATA As Date = #7/7/2098#
    Const QA_TURNO As String = "QA_T07_POSITIVO"
    Dim wsBI As Worksheet, lo As ListObject, lr As ListRow
    Dim cRun As Long, cAtivo As Long, cAnalito As Long, cLote As Long
    Dim cNivel As Long, cMedia As Long, cDP As Long, cRes As Long, cVer As Long
    Dim analito As String, core As String, nivel As Long, media As Double
    Dim run As Long, regs(1 To 1, 1 To 7) As Variant, up As String
    Dim dados As Variant, i As Long, nRunAntes As Long, nAtivos As Long, nOK As Long
    On Error GoTo falha

    AtualizarBIData
    Set wsBI = ThisWorkbook.Sheets("BI_Data")
    Set lo = wsBI.ListObjects("tblBI_Fato")
    cRun = lo.ListColumns("RUN").Index
    cAtivo = lo.ListColumns("Ativo").Index
    cAnalito = lo.ListColumns("Analito").Index
    cLote = lo.ListColumns("ID_Lote").Index
    cNivel = lo.ListColumns("Nivel").Index
    cMedia = lo.ListColumns("Media_Alvo").Index
    cDP = lo.ListColumns("DP_Alvo").Index
    cRes = lo.ListColumns("Resultado").Index
    cVer = lo.ListColumns("Veredito").Index

    For Each lr In lo.ListRows
        If CLng(Val(CStr(lr.Range.Cells(1, cAtivo).Value))) = 1 And _
           IsNumeric(lr.Range.Cells(1, cMedia).Value) And _
           IsNumeric(lr.Range.Cells(1, cDP).Value) Then
            If CDbl(lr.Range.Cells(1, cDP).Value) > 0 Then
                analito = Trim$(CStr(lr.Range.Cells(1, cAnalito).Value))
                core = Trim$(CStr(lr.Range.Cells(1, cLote).Value))
                nivel = CLng(Val(CStr(lr.Range.Cells(1, cNivel).Value)))
                media = CDbl(lr.Range.Cells(1, cMedia).Value)
                If analito <> "" And core <> "" And nivel >= 1 And nivel <= mEstatistica.NLV Then Exit For
            End If
        End If
    Next lr
    If analito = "" Or core = "" Or nivel = 0 Then
        QA_T07_PrepararRunPositivo = "FAIL|nenhum alvo com DP positivo disponivel no BI"
        Exit Function
    End If

    run = ObterOuCriarRUN(QA_DATA, core, QA_TURNO, nivel)
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If CLng(Val(CStr(dados(i, COL_RUN)))) = run Then nRunAntes = nRunAntes + 1
        Next i
    End If
    If nRunAntes <> 0 Then
        QA_T07_PrepararRunPositivo = "FAIL|RUN sentinela ja continha dados; execucao recusada"
        Exit Function
    End If

    regs(1, COL_RUN) = run
    regs(1, COL_DATA) = QA_DATA
    regs(1, COL_NIVEL) = nivel
    regs(1, COL_LOTE) = CodigoLote(core, nivel)
    regs(1, COL_ANALITO) = analito
    regs(1, COL_RESULT) = media
    regs(1, COL_STATUS) = ST_ATIVO
    up = UpsertResultados(regs)
    If Left$(up, 3) <> "1|0" Then
        QA_T07_PrepararRunPositivo = "FAIL|upsert inesperado=" & up
        Exit Function
    End If

    mEstatistica.InvalidarCache
    AtualizarBIData
    Set lo = wsBI.ListObjects("tblBI_Fato")
    For Each lr In lo.ListRows
        If CLng(Val(CStr(lr.Range.Cells(1, cRun).Value))) = run And _
           CLng(Val(CStr(lr.Range.Cells(1, cAtivo).Value))) = 1 And _
           IsNumeric(lr.Range.Cells(1, cRes).Value) Then
            nAtivos = nAtivos + 1
            If UCase$(Trim$(CStr(lr.Range.Cells(1, cVer).Value))) = "OK" Then nOK = nOK + 1
        End If
    Next lr
    If nAtivos <> 1 Or nOK <> 1 Then
        QA_T07_PrepararRunPositivo = "FAIL|RUN=" & run & "; ativos=" & nAtivos & "; ok=" & nOK
        Exit Function
    End If

    QA_T07_PrepararRunPositivo = "PASS|RUN=" & run & ";ANALITO=" & analito & _
        ";NIVEL=" & nivel & ";LOTE=" & core & ";RESULTADO=" & CStr(media) & _
        ";VEREDITO=OK;LINHAS=1"
    Exit Function
falha:
    QA_T07_PrepararRunPositivo = "FAIL|" & Err.Description
End Function

Public Function QA_T05_Reconciliacao() As String
    Dim r As String, p As Variant
    On Error GoTo falha
    mEstatistica.InvalidarCache
    AtualizarBIData
    r = ReconciliarBancoBI()
    p = Split(r, "|")
    If UBound(p) >= 1 And CLng(Val(CStr(p(1)))) = 0 Then
        QA_T05_Reconciliacao = "PASS|" & r
    Else
        QA_T05_Reconciliacao = "FAIL|" & r
    End If
    Exit Function
falha:
    QA_T05_Reconciliacao = "FAIL|" & Err.Description
End Function

Private Function Mesmo(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNumeric(a) And IsNumeric(b) Then
        Mesmo = (Abs(CDbl(a) - CDbl(b)) < 0.0000001)
    Else
        Mesmo = (CStr(a) = CStr(b))
    End If
End Function
