Attribute VB_Name = "mEtlCQ"
' ============================================================================
'  mEtlCQ -- a Principal - Resultados (tblCQ_Final) calculada em VBA (ADR-079)
'
'  Porte fiel de pq/DB_CQ_FINAL.m. A referencia executavel (estudos/ref_cq_final.py) reproduz a tabela gravada
'  pelo Power Query celula a celula; este modulo segue a mesma ordem de etapas e os mesmos nomes.
'  Por que: o Power Query levava 36 a 168 s nesta tabela (33 mil linhas); em memoria, segundos.
'
'  Regras que importam para a equivalencia (provadas nos dados reais):
'    - texto comparado BINARIO (ordinal); nulo (Empty) diferente de "";
'    - horario comparado em MILISSEGUNDOS inteiros (o PQ arredonda o serial do Excel ao ms);
'    - toda ordenacao tem desempate completo terminando no ID_REGISTRO -- a estabilidade nao importa;
'    - corte de corrida: intervalo ESTRITAMENTE maior que GAP_CORRIDA_MIN.
'  Saida: matriz (n, 40) nas colunas e na ordem da tblCQ_Final (NomesColunas). Datas e horas saem como SERIAL
'  (Double, ms exatos); textos como String -- gravar com .Value2 e as colunas de texto formatadas como texto.
'  ADR-082: SUBSTITUI_ID da tblResultados_Manuais -- o manual valido substitui o resultado do interfaceamento
'  indicado (SUBSTITUIDO_MANUAL; inativado pelo usuario continua INATIVADO). O indicado nunca forma corrida; o que
'  continua com status de corrida real fica com a CORRIDA_NO_DIA/RUN do manual. Mesmas regras e textos do PQ.
' ============================================================================
Option Explicit
Option Compare Binary

Private Const SIM As String = "SIM"
Private Const NAO As String = "NÃO"
Private Const NC As Long = 40

' campos internos de cada linha (U)
Private Const F_ID As Long = 1, F_IDO As Long = 2, F_EQ As Long = 3, F_MZ As Long = 4, F_BL As Long = 5
Private Const F_LT As Long = 6, F_NV As Long = 7, F_ND As Long = 8, F_DT As Long = 9, F_DH As Long = 10
Private Const F_AO As Long = 11, F_PD As Long = 12, F_VL As Long = 13, F_FL As Long = 14, F_UN As Long = 15
Private Const F_AM As Long = 16, F_IT As Long = 17, F_RE As Long = 18, F_AQ As Long = 19, F_AN As Long = 20
Private Const F_OR As Long = 21, F_LM As Long = 22, F_DUP As Long = 23, F_DUPDE As Long = 24, F_FALTA As Long = 25
Private Const F_CONFM As Long = 26, F_CONFID As Long = 27, F_MOTM As Long = 28, F_USRM As Long = 29
Private Const F_DHMS As Long = 30, F_PRIMID As Long = 31, F_KMAN As Long = 32, F_PRIMK As Long = 33, F_IDPRIMK As Long = 34
Private Const F_ST As Long = 35, F_GK As Long = 36, F_BLK As Long = 37, F_POS As Long = 38, F_COR As Long = 39
Private Const F_RUN As Long = 40, F_INAT As Long = 41, F_COMENT As Long = 42, F_LJ As Long = 43, F_IDDIG As Long = 44
' ADR-082: no manual, SUBSTITUI_ID (NormId), por que nao vale (Empty = valido), se vale (o 1o para o ID), o manual que ja
' substitui (repetido) e a nota; no interfaceamento, o manual que o substitui, o MOTIVO dele e a linha dele em U
Private Const F_SUBDIG As Long = 45, F_SUBR As Long = 46, F_SUBOK As Long = 47, F_SUBREP As Long = 48
Private Const F_SUBPOR As Long = 49, F_SUBMOT As Long = 50, F_NOTA As Long = 51, F_SUBIDX As Long = 52
Private Const NF As Long = 52
Private Const SUF_SUB As String = " (este manual nao substitui nada)"

Private mPrefixo As String
Private mRx As Object
Private mTempos As String                ' fases do ultimo calculo (diagnostico do desempenho)
Private mT0 As Double
Private mTemRec As Boolean               ' recebimento em memoria (staging do mEtlSync), em vez da tabela
Private mRecDados As Variant, mRecCab As Object, mRecN As Long

' O proximo CalcularCQFinal usa este recebimento (ainda NAO publicado) em vez de ler a tblDB_Recebimento.
' Datas como Date (o mesmo que a leitura .Value da tabela). Vale para UMA chamada.
Public Sub UsarRecebimento(ByRef dados As Variant, ByVal cab As Object, ByVal n As Long)
    mRecDados = dados: Set mRecCab = cab: mRecN = n: mTemRec = True
End Sub

Public Function TemposUltimo() As String
    TemposUltimo = mTempos
End Function

Private Sub Marca(ByVal fase As String)
    mTempos = mTempos & fase & "=" & Format$(Timer - mT0, "0.00") & ";"
    mT0 = Timer
End Sub

' ---------------------------------------------------------------- colunas de saida (tblCQ_Final)
Public Function NomesColunas() As Variant
    NomesColunas = Array("ID_REGISTRO", "ORIGEM_RESULTADO", "SETOR", "EQUIPAMENTO", "MATRIZ", "BLOCO", "DATA", "HORA", _
        "DATA_HORA", "CORRIDA_NO_DIA", "RUN", "NIVEL", "NIVEL_DESC", "LOTE", "ANALITO", "ANALITO_ORIGEM", _
        "ANALITO_CADASTRADO", "RESULTADO", "UNIDADE", "FLAG", "STATUS_ANALITICO", "PARTICIPA_ESTATISTICA", _
        "REGISTRAR_RESULTADO_NO_LJ", "TIPO_PLOTAGEM_LJ", "INATIVACAO_REGISTRADA", "COMENTARIO_TECNICO", _
        "TEM_JUSTIFICATIVA", "GOVERNANCA", "DATA_INATIVACAO", "USUARIO_INATIVACAO", "MOTIVO_INATIVACAO", _
        "MOTIVO_EXCLUSAO_AUTOMATICA", "ID_RELACIONADO", "ID_ORIGEM", "ID_AMOSTRA", "ITEM_ID", "PRIMEIRO_DO_DIA", _
        "MOTIVO_MANUAL", "USUARIO_MANUAL", "RECEBIDO_EM")
End Function

' ================================================================ utilitarios de tipo (M -> VBA)
Private Function EhVazio(ByVal v As Variant) As Boolean
    EhVazio = IsEmpty(v) Or IsNull(v)
End Function

Private Function EhNum(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, 20: EhNum = True
    End Select
End Function

' Text.Trim do M: tira espaco Unicode dos dois lados
Private Function TrimU(ByVal s As String) As String
    Dim i As Long, j As Long
    i = 1: j = Len(s)
    Do While i <= j
        If Not EhEspaco(AscW(Mid$(s, i, 1))) Then Exit Do
        i = i + 1
    Loop
    Do While j >= i
        If Not EhEspaco(AscW(Mid$(s, j, 1))) Then Exit Do
        j = j - 1
    Loop
    If j >= i Then TrimU = Mid$(s, i, j - i + 1) Else TrimU = ""
End Function

Private Function EhEspaco(ByVal c As Long) As Boolean
    If c < 0 Then c = c + 65536
    Select Case c
        Case 9 To 13, 32, 133, 160, 5760, 8192 To 8202, 8232, 8233, 8239, 8287, 12288: EhEspaco = True
    End Select
End Function

' Text.From (pt-BR) nos casos que aparecem aqui; Empty continua Empty
Private Function TextFrom(ByVal x As Variant) As Variant
    If EhVazio(x) Then TextFrom = Empty: Exit Function
    Select Case VarType(x)
        Case vbString: TextFrom = x
        Case vbBoolean: TextFrom = IIf(x, "true", "false")
        Case vbDate
            If Int(CDbl(x)) = CDbl(x) Then
                TextFrom = Format$(x, "dd/mm/yyyy")
            ElseIf CDbl(x) < 1 Then
                TextFrom = Format$(x, "hh:nn:ss")
            Else
                TextFrom = Format$(x, "dd/mm/yyyy hh:nn:ss")
            End If
        Case Else
            If EhNum(x) Then
                If CDbl(x) = Int(CDbl(x)) And Abs(CDbl(x)) < 1E+15 Then
                    TextFrom = Format$(CDbl(x), "0")
                Else
                    TextFrom = CStr(CDbl(x))
                End If
            Else
                TextFrom = CStr(x)
            End If
    End Select
End Function

Private Function Txt(ByVal x As Variant) As Variant
    Dim t As String
    If EhVazio(x) Then Txt = Empty: Exit Function
    t = TrimU(CStr(TextFrom(x)))
    If Len(t) = 0 Then Txt = Empty Else Txt = t
End Function

Private Function SoDigitos(ByVal t As String) As Boolean
    Dim i As Long, c As String
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    SoDigitos = True
End Function

Private Function SemZeros(ByVal d As String) As String
    Dim i As Long
    i = 1
    Do While i <= Len(d)
        If Mid$(d, i, 1) <> "0" Then Exit Do
        i = i + 1
    Loop
    If i > Len(d) Then SemZeros = "0" Else SemZeros = Mid$(d, i)
End Function

Private Function NormId(ByVal x As Variant) As Variant
    Dim t As Variant, u As String, pre As String
    t = Txt(x)
    If IsEmpty(t) Then NormId = Empty: Exit Function
    u = UCase$(Replace(Replace(CStr(t), " ", ""), ChrW$(160), ""))
    pre = mPrefixo & "-"
    If SoDigitos(u) Then
        NormId = pre & SemZeros(u)
    ElseIf Left$(u, Len(pre)) = pre And SoDigitos(Mid$(u, Len(pre) + 1)) Then
        NormId = pre & SemZeros(Mid$(u, Len(pre) + 1))
    ElseIf Left$(u, 4) = "MAN_" And SoDigitos(Mid$(u, 5)) Then
        NormId = "MAN_" & Right$(String$(4, "0") & Mid$(u, 5), IIf(Len(Mid$(u, 5)) > 4, Len(Mid$(u, 5)), 4))
    Else
        NormId = u
    End If
End Function

Private Function Logico(ByVal v As Variant, ByVal padrao As Boolean) As Boolean
    Dim t As String
    If EhVazio(v) Then Logico = padrao: Exit Function
    If VarType(v) = vbBoolean Then Logico = v: Exit Function
    If EhNum(v) Then Logico = (CDbl(v) <> 0): Exit Function
    t = UCase$(TrimU(CStr(TextFrom(v))))
    If Len(t) = 0 Then Logico = padrao: Exit Function
    Select Case t
        Case "SIM", "S", "TRUE", "VERDADEIRO", "V", "X", "1": Logico = True
    End Select
End Function

' Int64.From: numero arredondado ao par (Round do VBA tambem e ao par); texto pt-BR; erro levanta
Private Function Int64(ByVal v As Variant) As Variant
    If EhVazio(v) Then Int64 = Empty: Exit Function
    If VarType(v) = vbBoolean Then Err.Raise 13
    If EhNum(v) Then Int64 = CDbl(Round(CDbl(v), 0)): Exit Function
    If VarType(v) = vbString Then Int64 = CDbl(Round(Val(Replace(Replace(TrimU(CStr(v)), ".", ""), ",", ".")), 0)): Exit Function
    Err.Raise 13
End Function

' milissegundos inteiros de um serial de data/hora (o PQ arredonda ao ms)
Private Function Ms(ByVal d As Variant) As Double
    Ms = Int(CDbl(d) * 86400000# + 0.5)
End Function

Private Function DeMs(ByVal m As Double) As Date
    DeMs = CDate(m / 86400000#)
End Function

Private Function ParaData(ByVal v As Variant) As Variant
    If EhVazio(v) Then ParaData = Empty: Exit Function
    If VarType(v) = vbDate Then ParaData = CDate(Int(CDbl(v))): Exit Function
    If EhNum(v) Then ParaData = CDate(Int(CDbl(v))): Exit Function
    If VarType(v) = vbString Then
        Dim p() As String
        p = Split(TrimU(CStr(v)), "/")
        If UBound(p) <> 2 Then Err.Raise 13
        ParaData = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
        Exit Function
    End If
    Err.Raise 13
End Function

Private Function ParaDataHora(ByVal v As Variant) As Variant
    If EhVazio(v) Then ParaDataHora = Empty: Exit Function
    If VarType(v) = vbDate Or EhNum(v) Then ParaDataHora = DeMs(Ms(CDbl(v))): Exit Function
    If VarType(v) = vbString Then
        Dim s As String, dpart As String, hpart As String, p() As String, h() As String
        s = TrimU(CStr(v))
        If InStr(s, " ") > 0 Then dpart = Left$(s, InStr(s, " ") - 1): hpart = Mid$(s, InStr(s, " ") + 1) Else dpart = s
        p = Split(dpart, "/")
        If UBound(p) <> 2 Then Err.Raise 13
        ParaDataHora = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
        If Len(hpart) > 0 Then
            h = Split(hpart, ":")
            If UBound(h) < 1 Then Err.Raise 13
            ParaDataHora = ParaDataHora + TimeSerial(CInt(h(0)), CInt(h(1)), IIf(UBound(h) >= 2, CInt(h(2)), 0))
        End If
        Exit Function
    End If
    Err.Raise 13
End Function

' hora (fracao do dia) em ms inteiros, ou Empty/erro
Private Function ParaHoraMs(ByVal v As Variant) As Variant
    If VarType(v) = vbBoolean Then Err.Raise 13
    If VarType(v) = vbDate Then
        If CDbl(v) >= 1 Then
            ParaHoraMs = Ms(CDbl(v)) - Int(Ms(CDbl(v)) / 86400000#) * 86400000#
        Else
            ParaHoraMs = Int((CDbl(v) - Fix(CDbl(v))) * 86400000# + 0.5)
        End If
        Exit Function
    End If
    If EhNum(v) Then ParaHoraMs = Int((CDbl(v) - Fix(CDbl(v))) * 86400000# + 0.5): Exit Function
    If VarType(v) = vbString Then
        Dim h() As String
        h = Split(TrimU(CStr(v)), ":")
        If UBound(h) < 1 Or UBound(h) > 2 Then Err.Raise 13
        ParaHoraMs = (CDbl(CInt(h(0))) * 3600# + CInt(h(1)) * 60# + IIf(UBound(h) = 2, CInt(h(2)), 0)) * 1000#
        Exit Function
    End If
    Err.Raise 13
End Function

' Double.Parse(pt-BR): '.' de milhar, ',' decimal
Private Function NumPt(ByVal v As Variant) As Double
    Dim s As String, partes() As String, mant As String
    s = TrimU(CStr(TextFrom(v)))
    If mRx Is Nothing Then
        Set mRx = CreateObject("VBScript.RegExp")
        mRx.Pattern = "^[+-]?(\d[\d.]*)?(,\d*)?([eE][+-]?\d+)?$"
    End If
    If Not mRx.Test(s) Then Err.Raise 13
    partes = Split(s, ",")
    mant = partes(0)
    If UBound(partes) >= 1 Then mant = mant & partes(1)
    If Not (mant Like "*#*") Then Err.Raise 13
    NumPt = Val(Replace(Replace(s, ".", ""), ",", "."))
End Function

' ================================================================ chaves de ordenacao (binarias)
Private Function KS(ByVal v As Variant) As String          ' texto; nulo primeiro
    If EhVazio(v) Then KS = Chr$(1) Else KS = Chr$(2) & CStr(v)
End Function

Private Function KN(ByVal v As Variant) As String          ' inteiro >= 0; nulo primeiro
    If EhVazio(v) Then KN = Chr$(1) Else KN = Chr$(2) & P15(v)
End Function

' inteiro >= 0 com 15 digitos (Format$ e lento demais no laco)
Private Function P15(ByVal v As Variant) As String
    P15 = CStr(CDbl(v))
    P15 = String$(15 - Len(P15), "0") & P15
End Function

Private Function KDup(ByVal v As Variant) As String        ' chave de agrupamento (nulo distinto de "")
    If EhVazio(v) Then
        KDup = Chr$(1)
    ElseIf VarType(v) = vbString Then
        KDup = "s" & v
    Else
        KDup = "n" & CStr(v)
    End If
End Function

' ordena (indice) por chave texto, binario, mesclagem estavel; devolve os indices ordenados
Private Function OrdenarPorChave(ByRef chave() As String, ByVal n As Long) As Long()
    Dim a() As Long, b() As Long, w As Long, i As Long, lo As Long, mid_ As Long, hi As Long
    Dim l As Long, r As Long, k As Long, t() As Long
    ReDim a(1 To IIf(n > 0, n, 1)): ReDim b(1 To IIf(n > 0, n, 1))
    For i = 1 To n: a(i) = i: Next i
    w = 1
    Do While w < n
        lo = 1
        Do While lo <= n
            mid_ = lo + w - 1: If mid_ > n Then mid_ = n
            hi = lo + 2 * w - 1: If hi > n Then hi = n
            l = lo: r = mid_ + 1: k = lo
            Do While l <= mid_ And r <= hi
                If chave(a(r)) < chave(a(l)) Then b(k) = a(r): r = r + 1 Else b(k) = a(l): l = l + 1
                k = k + 1
            Loop
            Do While l <= mid_: b(k) = a(l): l = l + 1: k = k + 1: Loop
            Do While r <= hi: b(k) = a(r): r = r + 1: k = k + 1: Loop
            lo = lo + 2 * w
        Loop
        t = a: a = b: b = t
        w = w * 2
    Loop
    OrdenarPorChave = a
End Function

' ================================================================ leitura das tabelas da pasta
Private Function Tabela(ByVal nome As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set Tabela = ws.ListObjects(nome)
        On Error GoTo 0
        If Not Tabela Is Nothing Then Exit Function
    Next ws
End Function

' cabecalho -> posicao, e os dados (matriz 1-based; n = 0 quando vazia)
Private Function LerTabela(ByVal nome As String, ByRef cab As Object, ByRef n As Long, _
                           Optional ByVal obrigatoria As Boolean = True) As Variant
    Dim lo As ListObject, i As Long, h As Variant
    Set lo = Tabela(nome)
    If lo Is Nothing Then
        If obrigatoria Then Err.Raise vbObjectError + 581, "mEtlCQ", "Tabela obrigatoria ausente ou renomeada: " & nome & _
                                       " (a atualizacao parou; nada foi alterado)"
        Set cab = CreateObject("Scripting.Dictionary"): n = 0: Exit Function
    End If
    Set cab = CreateObject("Scripting.Dictionary")
    h = lo.HeaderRowRange.Value
    For i = 1 To UBound(h, 2)
        If Not cab.Exists(CStr(h(1, i))) Then cab.Add CStr(h(1, i)), i
    Next i
    n = lo.ListRows.Count
    If n > 0 Then
        If n = 1 And lo.ListColumns.Count = 1 Then
            Dim um(1 To 1, 1 To 1) As Variant
            um(1, 1) = lo.DataBodyRange.Value
            LerTabela = um
        Else
            LerTabela = lo.DataBodyRange.Value
        End If
    End If
End Function

Private Function ColIdx(ByVal cab As Object, ByVal nome As String) As Long
    If cab.Exists(nome) Then ColIdx = cab(nome)
End Function

' celula da matriz lida (coluna 0 = ausente); erro de celula e Null viram Empty
Private Function Cel(ByRef dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then Exit Function
    Cel = dados(r, c)
    If IsError(Cel) Or IsNull(Cel) Then Cel = Empty
End Function

Private Function Campo(ByRef dados As Variant, ByVal cab As Object, ByVal r As Long, ByVal nome As String) As Variant
    If cab.Exists(nome) Then
        Campo = dados(r, cab(nome))
        If IsError(Campo) Then Campo = Empty
    Else
        Campo = Empty
    End If
End Function

' ================================================================ o calculo
' Calcula a tblCQ_Final a partir do recebimento e das tabelas de entrada da propria pasta. agora = Now (manuais).
Public Function CalcularCQFinal(Optional ByVal agora As Variant) As Variant
    Dim cCfg As Object, dCfg As Variant, nCfg As Long, C As Object, i As Long, j As Long, k As Variant
    Dim GAP As Double, TOL As Double, NIVEIS As Double, MATRIZ_PADRAO As Variant, EQUIP_PADRAO As Variant, SETOR As Variant
    If IsMissing(agora) Then agora = Now
    mTempos = "": mT0 = Timer

    ' ---- CFG
    dCfg = LerTabela("tblConfigIntegracao", cCfg, nCfg)
    Set C = CreateObject("Scripting.Dictionary")
    For i = 1 To nCfg
        k = Txt(Campo(dCfg, cCfg, i, "CHAVE"))
        If Not IsEmpty(k) Then
            If VarType(Campo(dCfg, cCfg, i, "VALOR")) = vbString Then
                C(UCase$(k)) = TrimU(Campo(dCfg, cCfg, i, "VALOR"))
            Else
                C(UCase$(k)) = Campo(dCfg, cCfg, i, "VALOR")
            End If
        End If
    Next i
    GAP = CDbl(C("GAP_CORRIDA_MIN")): TOL = CDbl(C("TOLERANCIA_CONFLITO_MIN"))
    mPrefixo = UCase$(TrimU(CStr(TextFrom(C("PREFIXO_ID")))))
    NIVEIS = CDbl(C("NIVEIS"))
    MATRIZ_PADRAO = C("MATRIZ_PADRAO"): EQUIP_PADRAO = C("EQUIPAMENTO_PADRAO"): SETOR = C("SETOR")

    ' ---- de/para (so linhas com ANALITO_ORIGEM)
    Dim cDp As Object, dDp As Variant, nDp As Long, DP() As Variant, nD As Long, cadastrados As Object
    dDp = LerTabela("tblDeParaAnalitos", cDp, nDp)
    ReDim DP(1 To IIf(nDp > 0, nDp, 1), 1 To 4)
    Set cadastrados = CreateObject("Scripting.Dictionary")
    For i = 1 To nDp
        If Not IsEmpty(Txt(Campo(dDp, cDp, i, "ANALITO_ORIGEM"))) Then
            nD = nD + 1
            DP(nD, 1) = Txt(Campo(dDp, cDp, i, "MATRIZ"))
            DP(nD, 2) = Txt(Campo(dDp, cDp, i, "ANALITO_ORIGEM"))
            DP(nD, 3) = Txt(Campo(dDp, cDp, i, "ANALITO_QCINI"))
            DP(nD, 4) = Txt(Campo(dDp, cDp, i, "BLOCO"))
            If Not IsEmpty(DP(nD, 3)) Then cadastrados(DP(nD, 3)) = 1
        End If
    Next i

    ' ---- interfaceamento (R1): 1 linha por ID, na ordem da tabela
    Dim cR As Object, dR As Variant, nR As Long, U() As Variant, nU As Long, vistos As Object, idr As Variant, mz As Variant
    Dim nMan As Long, cM As Object, dM As Variant
    If mTemRec Then
        dR = mRecDados: Set cR = mRecCab: nR = mRecN
        mTemRec = False: mRecDados = Empty
    Else
        dR = LerTabela("tblDB_Recebimento", cR, nR)
    End If
    dM = LerTabela("tblResultados_Manuais", cM, nMan)
    ReDim U(1 To nR + nMan + 1, 1 To NF)
    Set vistos = CreateObject("Scripting.Dictionary")
    Dim mapa As Object, nivelDesc As Object, kk As String
    Set mapa = CreateObject("Scripting.Dictionary")
    Set nivelDesc = CreateObject("Scripting.Dictionary")
    ' colunas resolvidas UMA vez (0 = ausente, vira nulo -- MissingField.UseNull)
    Dim cId As Long, cIdo As Long, cEq As Long, cMz As Long, cBl As Long, cLt As Long, cNv As Long, cNd As Long
    Dim cDt As Long, cDh As Long, cAn As Long, cPd As Long, cVl As Long, cFl As Long, cUn As Long, cAm As Long
    Dim cIt As Long, cRe As Long, x As Variant
    cId = ColIdx(cR, "ID_REGISTRO"): cIdo = ColIdx(cR, "ID_ORIGEM"): cEq = ColIdx(cR, "EQUIPAMENTO")
    cMz = ColIdx(cR, "MATRIZ"): cBl = ColIdx(cR, "BLOCO"): cLt = ColIdx(cR, "LOTE"): cNv = ColIdx(cR, "NIVEL")
    cNd = ColIdx(cR, "NIVEL_DESC"): cDt = ColIdx(cR, "DATA"): cDh = ColIdx(cR, "DATA_HORA"): cAn = ColIdx(cR, "ANALITO")
    cPd = ColIdx(cR, "PRIMEIRO_DO_DIA"): cVl = ColIdx(cR, "VALOR"): cFl = ColIdx(cR, "FLAG"): cUn = ColIdx(cR, "UNIDADE")
    cAm = ColIdx(cR, "ID_AMOSTRA"): cIt = ColIdx(cR, "ITEM_ID"): cRe = ColIdx(cR, "RECEBIDO_EM")
    For i = 1 To nR
        idr = Cel(dR, i, cId)
        If Not IsEmpty(idr) Then
            idr = TextFrom(idr)
            If Not vistos.Exists(idr) Then
                vistos.Add idr, 1
                nU = nU + 1
                mz = Cel(dR, i, cMz): If IsEmpty(mz) Then mz = MATRIZ_PADRAO
                U(nU, F_ID) = idr
                U(nU, F_IDO) = Int64(Cel(dR, i, cIdo))
                U(nU, F_EQ) = TextFrom(Cel(dR, i, cEq))
                U(nU, F_MZ) = TextFrom(mz)
                U(nU, F_BL) = TextFrom(Cel(dR, i, cBl))
                U(nU, F_LT) = TextFrom(Cel(dR, i, cLt))
                U(nU, F_NV) = Int64(Cel(dR, i, cNv))
                U(nU, F_ND) = TextFrom(Cel(dR, i, cNd))
                U(nU, F_DT) = ParaData(Cel(dR, i, cDt))
                x = Cel(dR, i, cDh)
                If IsEmpty(x) Then
                    U(nU, F_DH) = Empty
                ElseIf VarType(x) = vbDate Or EhNum(x) Then
                    U(nU, F_DHMS) = Ms(CDbl(x)): U(nU, F_DH) = U(nU, F_DHMS) / 86400000#
                Else
                    U(nU, F_DH) = ParaDataHora(x): U(nU, F_DHMS) = Ms(U(nU, F_DH))
                End If
                U(nU, F_AO) = TextFrom(Cel(dR, i, cAn))
                U(nU, F_PD) = TextFrom(Cel(dR, i, cPd))
                x = Cel(dR, i, cVl): If IsEmpty(x) Then U(nU, F_VL) = Empty Else U(nU, F_VL) = CDbl(x)
                U(nU, F_FL) = TextFrom(Cel(dR, i, cFl))
                U(nU, F_UN) = TextFrom(Cel(dR, i, cUn))
                U(nU, F_AM) = TextFrom(Cel(dR, i, cAm))
                U(nU, F_IT) = Int64(Cel(dR, i, cIt))
                U(nU, F_RE) = ParaDataHora(Cel(dR, i, cRe))
                kk = KDup(U(nU, F_NV))
                If Not nivelDesc.Exists(kk) Then nivelDesc.Add kk, U(nU, F_ND)
                kk = KDup(U(nU, F_MZ)) & Chr$(0) & KDup(U(nU, F_AO))
                If Not mapa.Exists(kk) Then mapa.Add kk, ParaQcini(DP, nD, U(nU, F_MZ), U(nU, F_AO))
                U(nU, F_AQ) = mapa(kk)
                If IsEmpty(U(nU, F_AQ)) Then U(nU, F_AN) = U(nU, F_AO) Else U(nU, F_AN) = U(nU, F_AQ)
                U(nU, F_OR) = "INTERFACEAMENTO"
                U(nU, F_FALTA) = ""
            End If
        End If
    Next i
    Marca "ler+interface"
    Dim nInterface As Long
    nInterface = nU

    ' ---- retransmissao (DUPLICIDADE_ORIGEM): grupo de 10 colunas; o menor ID_ORIGEM e o canonico
    Dim gN As Object, gMin As Object, gkeys() As String, canon As Variant
    Set gN = CreateObject("Scripting.Dictionary"): Set gMin = CreateObject("Scripting.Dictionary")
    ReDim gkeys(1 To IIf(nInterface > 0, nInterface, 1))
    For i = 1 To nInterface
        gkeys(i) = KDup(U(i, F_EQ)) & Chr$(0) & KDup(U(i, F_MZ)) & Chr$(0) & KDup(U(i, F_LT)) & Chr$(0) & KDup(U(i, F_NV)) & _
                   Chr$(0) & KDup(U(i, F_AO)) & Chr$(0) & KDup(U(i, F_DHMS)) & Chr$(0) & KDup(U(i, F_AM)) & Chr$(0) & _
                   KDup(U(i, F_IT)) & Chr$(0) & KDup(U(i, F_VL)) & Chr$(0) & KDup(U(i, F_FL))
        gN(gkeys(i)) = gN(gkeys(i)) + 1
        If Not IsEmpty(U(i, F_IDO)) Then
            If Not gMin.Exists(gkeys(i)) Then
                gMin.Add gkeys(i), U(i, F_IDO)
            ElseIf U(i, F_IDO) < gMin(gkeys(i)) Then
                gMin(gkeys(i)) = U(i, F_IDO)
            End If
        End If
    Next i
    For i = 1 To nInterface
        U(i, F_DUP) = False
        If gN(gkeys(i)) > 1 Then
            If gMin.Exists(gkeys(i)) Then canon = gMin(gkeys(i)) Else canon = Empty
            If Not MesmoValor(U(i, F_IDO), canon) Then
                U(i, F_DUP) = True
                U(i, F_DUPDE) = mPrefixo & "-" & TextFrom(canon)
            End If
        End If
    Next i
    Marca "duplicidade"
    ' ---- manuais
    CalcularManuais U, nU, dM, cM, nMan, DP, nD, cadastrados, nivelDesc, C, NIVEIS, TOL, agora, nInterface

    Marca "manuais"
    ' ---- inativacao e comentarios
    Dim cI As Object, dI As Variant, nI As Long, inat As Object, legado As Boolean, idn As Variant, rec As Variant
    dI = LerTabela("tblInativacao_NaoConformes", cI, nI)
    legado = Not cI.Exists("MOTIVO")
    Set inat = CreateObject("Scripting.Dictionary")
    For i = 1 To nI
        idn = NormId(Campo(dI, cI, i, "ID_REGISTRO"))
        If Not IsEmpty(idn) Then
            If Not inat.Exists(idn) Then
                Dim dti As Variant
                dti = Empty
                On Error Resume Next
                dti = ParaDataHora(Campo(dI, cI, i, "DATA_INATIVACAO"))
                If Err.Number <> 0 Then dti = Empty
                Err.Clear
                On Error GoTo 0
                rec = Array(Logico(Campo(dI, cI, i, "REGISTRAR - LJ"), IsEmpty(dti)), dti, Txt(Campo(dI, cI, i, "USUARIO")), _
                            IIf(legado, Empty, AnKey(Campo(dI, cI, i, "ANALITO"))), IIf(legado, Empty, Txt(Campo(dI, cI, i, "MOTIVO"))))
                inat.Add idn, rec
            End If
        End If
    Next i
    Dim cK As Object, dK As Variant, nK As Long, coment As Object, tx As Variant
    dK = LerTabela("tblComentariosTecnicos", cK, nK)
    Set coment = CreateObject("Scripting.Dictionary")
    For i = 1 To nK
        idn = NormId(Campo(dK, cK, i, "ID_REGISTRO"))
        tx = Txt(Campo(dK, cK, i, "COMENTARIO_TECNICO"))
        If Not IsEmpty(idn) And Not IsEmpty(tx) Then
            If coment.Exists(idn) Then coment(idn) = coment(idn) & " | " & tx Else coment.Add idn, tx
        End If
    Next i

    ' ---- status
    Dim st As String, inatOk As Boolean, lj As Boolean, out() As Variant
    ReDim out(1 To IIf(nU > 0, nU, 1), 1 To NC)
    For i = 1 To nU
        inatOk = False
        U(i, F_LJ) = Empty
        If inat.Exists(U(i, F_ID)) Then
            rec = inat(U(i, F_ID))
            U(i, F_LJ) = rec(0)
            If legado Then
                inatOk = True
            ElseIf Not IsEmpty(rec(3)) Then
                inatOk = (rec(3) = AnKey(U(i, F_AN)))
            End If
        End If
        If coment.Exists(U(i, F_ID)) Then U(i, F_COMENT) = coment(U(i, F_ID)) Else U(i, F_COMENT) = Empty
        If U(i, F_FALTA) <> "" Then
            st = "MANUAL_INCOMPLETO"
        ElseIf U(i, F_DUP) Then
            st = "DUPLICIDADE_ORIGEM"
        ElseIf Not IsEmpty(U(i, F_CONFM)) Then
            st = "CONFLITO_MANUAL"
        ElseIf Not IsEmpty(U(i, F_SUBPOR)) And Not inatOk Then
            st = "SUBSTITUIDO_MANUAL"                    ' ADR-082 (inativado pelo usuario continua INATIVADO)
        ElseIf IsEmpty(U(i, F_DH)) Then
            st = "SEM_DATA_HORA"
        ElseIf IsEmpty(U(i, F_VL)) Then
            st = "SEM_VALOR"
        ElseIf inatOk Then
            st = "INATIVADO"
        Else
            st = "ATIVO"
        End If
        U(i, F_ST) = st
        U(i, F_INAT) = inatOk
        lj = (st = "ATIVO")
        If st = "INATIVADO" Then lj = (U(i, F_LJ) = True)
        out(i, 21) = st
        out(i, 22) = IIf(st = "ATIVO", SIM, NAO)
        out(i, 23) = IIf(lj, SIM, NAO)
        out(i, 24) = IIf(st = "ATIVO", "NORMAL", IIf(lj, "X_VERMELHO", "NAO_PLOTAR"))
        Select Case st
            Case "MANUAL_INCOMPLETO": out(i, 32) = "campos obrigatorios ausentes/invalidos: " & U(i, F_FALTA)
            Case "DUPLICIDADE_ORIGEM": out(i, 32) = "retransmissao do interfaceamento"
            Case "CONFLITO_MANUAL": out(i, 32) = U(i, F_CONFM)
            Case "SUBSTITUIDO_MANUAL": out(i, 32) = TextoSubstituido(U, i)
            Case "SEM_DATA_HORA": out(i, 32) = "resultado do interfaceamento sem DATA_HORA (fora da corrida)"
            Case "SEM_VALOR": out(i, 32) = "resultado sem valor numerico (flag " & IIf(IsEmpty(U(i, F_FL)), "vazia", U(i, F_FL)) & ")"
            Case Else
                ' ADR-082: inativado pelo usuario e substituido continua INATIVADO, com o vinculo; manual: a nota
                If Not IsEmpty(U(i, F_SUBPOR)) Then out(i, 32) = TextoSubstituido(U, i) Else out(i, 32) = U(i, F_NOTA)
        End Select
        If st = "DUPLICIDADE_ORIGEM" Then
            out(i, 33) = U(i, F_DUPDE)
        ElseIf st = "CONFLITO_MANUAL" Then
            out(i, 33) = U(i, F_CONFID)
        ElseIf Not IsEmpty(U(i, F_SUBPOR)) Then
            out(i, 33) = U(i, F_SUBPOR)                 ' ADR-082: vinculo nos dois sentidos
        ElseIf U(i, F_SUBOK) = True Then
            out(i, 33) = U(i, F_SUBDIG)
        End If
        out(i, 25) = IIf(inatOk, SIM, NAO)
        If inatOk Then
            rec = inat(U(i, F_ID))
            out(i, 31) = rec(4): out(i, 29) = rec(1): out(i, 30) = rec(2)
        End If
        out(i, 27) = IIf(Not IsEmpty(U(i, F_COMENT)) Or Not IsEmpty(out(i, 31)), SIM, NAO)
        out(i, 28) = IIf(st = "INATIVADO" And IsEmpty(U(i, F_COMENT)) And IsEmpty(out(i, 31)), _
                         "ERRO DE GOVERNANÇA: INATIVADO SEM JUSTIFICATIVA TÉCNICA", "OK")
        out(i, 26) = U(i, F_COMENT)
    Next i

    Marca "status"
    ' ---- corrida e RUN
    CalcularCorridas U, nU, GAP
    Marca "corridas"

    ' ---- saida, na ordem final (DATA_HORA, ANALITO, NIVEL, ID_REGISTRO; nulos primeiro)
    Dim chave() As String, ordem() As Long, res() As Variant, o As Long
    If nU = 0 Then CalcularCQFinal = Empty: Exit Function
    ReDim chave(1 To nU)
    For i = 1 To nU
        If IsEmpty(U(i, F_DH)) Then chave(i) = Chr$(1) Else chave(i) = Chr$(2) & P15(U(i, F_DHMS))
        chave(i) = chave(i) & Chr$(0) & KS(U(i, F_AN)) & Chr$(0) & KN(U(i, F_NV)) & Chr$(0) & KS(U(i, F_ID))
    Next i
    ordem = OrdenarPorChave(chave, nU)
    Marca "ordem final"
    ReDim res(1 To nU, 1 To NC)
    For o = 1 To nU
        i = ordem(o)
        res(o, 1) = U(i, F_ID): res(o, 2) = U(i, F_OR): res(o, 3) = SETOR: res(o, 4) = U(i, F_EQ)
        res(o, 5) = U(i, F_MZ): res(o, 6) = U(i, F_BL): res(o, 7) = Serial(U(i, F_DT))
        If IsEmpty(U(i, F_DH)) Then
            res(o, 8) = Empty: res(o, 9) = Empty
        Else
            res(o, 8) = (U(i, F_DHMS) - Int(U(i, F_DHMS) / 86400000#) * 86400000#) / 86400000#
            res(o, 9) = U(i, F_DHMS) / 86400000#
        End If
        res(o, 10) = U(i, F_COR): res(o, 11) = U(i, F_RUN): res(o, 12) = U(i, F_NV): res(o, 13) = U(i, F_ND)
        res(o, 14) = U(i, F_LT): res(o, 15) = U(i, F_AN): res(o, 16) = U(i, F_AO)
        res(o, 17) = IIf(IsEmpty(U(i, F_AQ)), NAO, SIM)
        res(o, 18) = U(i, F_VL): res(o, 19) = U(i, F_UN): res(o, 20) = U(i, F_FL)
        For j = 21 To 33: res(o, j) = out(i, j): Next j
        res(o, 29) = Serial(res(o, 29))
        res(o, 34) = U(i, F_IDO): res(o, 35) = U(i, F_AM): res(o, 36) = U(i, F_IT): res(o, 37) = U(i, F_PD)
        res(o, 38) = U(i, F_MOTM): res(o, 39) = U(i, F_USRM): res(o, 40) = Serial(U(i, F_RE))
    Next o
    Marca "saida"
    CalcularCQFinal = res
End Function

' data/hora como serial (Double): gravar Date pelo .Value arredonda ao segundo; a publicacao grava pelo .Value2
Private Function Serial(ByVal d As Variant) As Variant
    If IsEmpty(d) Then Serial = Empty Else Serial = Ms(CDbl(d)) / 86400000#
End Function

Private Function MesmoValor(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) Or IsEmpty(b) Then MesmoValor = (IsEmpty(a) And IsEmpty(b)) Else MesmoValor = (a = b)
End Function

' ADR-082: MOTIVO_EXCLUSAO_AUTOMATICA do resultado do interfaceamento substituido por um manual
Private Function TextoSubstituido(ByRef U() As Variant, ByVal i As Long) As String
    Dim m As String, mt As String
    If IsEmpty(U(i, F_SUBPOR)) Then m = "?" Else m = CStr(U(i, F_SUBPOR))
    If IsEmpty(U(i, F_SUBMOT)) Then mt = "?" Else mt = CStr(U(i, F_SUBMOT))
    TextoSubstituido = "substituido pelo manual " & m & " (motivo: " & mt & ")"
End Function

' ADR-082: por que o SUBSTITUI_ID do manual i NAO vale (Empty = vale). idxInt = ID do interfaceamento -> linha de U.
' Mesma ordem e mesmos textos do PQ (MS2): inexistente / manual, retransmissao, chave diferente.
Private Function RazaoSubstituicao(ByRef U() As Variant, ByVal i As Long, ByVal idxInt As Object) As Variant
    Dim s As String, x As Long
    RazaoSubstituicao = Empty
    s = CStr(U(i, F_SUBDIG))
    If Not idxInt.Exists(s) Then
        If Left$(s, 4) = "MAN_" Then
            RazaoSubstituicao = s & " e resultado manual (SUBSTITUI_ID aceita so resultado do interfaceamento)"
        Else
            RazaoSubstituicao = "o ID " & s & " nao existe nos resultados do interfaceamento"
        End If
        Exit Function
    End If
    x = idxInt(s)
    If U(x, F_DUP) Then
        RazaoSubstituicao = s & " e retransmissao de " & CStr(U(x, F_DUPDE)) & " (informe o ID original)"
    ElseIf KDup(U(x, F_EQ)) <> KDup(U(i, F_EQ)) Or KDup(U(x, F_MZ)) <> KDup(U(i, F_MZ)) Or _
           KDup(U(x, F_LT)) <> KDup(U(i, F_LT)) Or KDup(U(x, F_NV)) <> KDup(U(i, F_NV)) Or _
           KDup(U(x, F_AN)) <> KDup(U(i, F_AN)) Then
        RazaoSubstituicao = s & " tem outra chave: " & ChaveTxt(U, x) & " (este manual: " & ChaveTxt(U, i) & ")"
    End If
End Function

' "EQUIPAMENTO | MATRIZ | LOTE | NIVEL | ANALITO" (vazio = "?"), como o ChaveTxt do PQ
Private Function ChaveTxt(ByRef U() As Variant, ByVal r As Long) As String
    Dim c As Variant, s As String, k As Long
    For Each c In Array(F_EQ, F_MZ, F_LT, F_NV, F_AN)
        If k > 0 Then s = s & " | "
        If IsEmpty(U(r, c)) Then s = s & "?" Else s = s & CStr(TextFrom(U(r, c)))
        k = k + 1
    Next c
    ChaveTxt = s
End Function

Private Function AnKey(ByVal x As Variant) As Variant
    Dim t As Variant
    t = Txt(x)
    If IsEmpty(t) Then AnKey = Empty Else AnKey = UCase$(t)
End Function

Private Function ParaQcini(ByRef DP() As Variant, ByVal nD As Long, ByVal mz As Variant, ByVal ao As Variant) As Variant
    Dim i As Long
    ParaQcini = Empty
    For i = 1 To nD
        If MesmoValor(DP(i, 2), ao) Then
            If IsEmpty(DP(i, 1)) Then ParaQcini = DP(i, 3): Exit Function
            If MesmoValor(DP(i, 1), mz) Then ParaQcini = DP(i, 3): Exit Function
        End If
    Next i
End Function

Private Function ParaOrigem(ByRef DP() As Variant, ByVal nD As Long, ByVal mz As Variant, ByVal q As Variant) As Long
    Dim i As Long
    For i = 1 To nD
        If MesmoValor(DP(i, 3), q) Then
            If IsEmpty(DP(i, 1)) Then ParaOrigem = i: Exit Function
            If MesmoValor(DP(i, 1), mz) Then ParaOrigem = i: Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------- resultados manuais (Manuais0, conflitos)
Private Sub CalcularManuais(ByRef U() As Variant, ByRef nU As Long, ByRef dM As Variant, ByVal cM As Object, ByVal nMan As Long, _
                            ByRef DP() As Variant, ByVal nD As Long, ByVal cadastrados As Object, ByVal nivelDesc As Object, _
                            ByVal C As Object, ByVal NIVEIS As Double, ByVal TOL As Double, ByVal agora As Variant, _
                            ByVal nInterface As Long)
    Dim i As Long, L As Long, conta As Long, cc As Variant, idn As Variant, dtv As Variant, hr As Variant
    Dim eq As Variant, mz As Variant, v As Variant, lt As Variant, nv As Variant, an As Variant, res As Variant
    Dim isnum As Boolean, rtx As Variant, amb As Boolean, vl As Variant, orig As Long, ndesc As Variant, dh As Variant
    Dim reg As Variant, regc As Variant, valido As Boolean, falt As String, ini As Long, subd As Variant
    ini = nU + 1
    For L = 1 To nMan
        conta = 0
        For Each cc In Array("ID_REGISTRO", "DATA", "HORA", "LOTE", "NIVEL", "ANALITO", "RESULTADO")
            If Not IsEmpty(Txt(Campo(dM, cM, L, CStr(cc)))) Then conta = conta + 1
        Next cc
        If conta > 0 Then
            idn = NormId(Campo(dM, cM, L, "ID_REGISTRO"))
            dtv = Empty: hr = Empty: nv = Empty: reg = Empty: vl = Empty
            On Error Resume Next
            dtv = ParaData(Campo(dM, cM, L, "DATA")): If Err.Number <> 0 Then dtv = Empty
            Err.Clear
            If Not EhVazio(Campo(dM, cM, L, "HORA")) Then
                hr = ParaHoraMs(Campo(dM, cM, L, "HORA")): If Err.Number <> 0 Then hr = Empty
                Err.Clear
            End If
            nv = Int64(Campo(dM, cM, L, "NIVEL")): If Err.Number <> 0 Then nv = Empty
            Err.Clear
            reg = ParaDataHora(Campo(dM, cM, L, "REGISTRADO_EM")): If Err.Number <> 0 Then reg = Empty
            Err.Clear
            On Error GoTo 0
            eq = Txt(Campo(dM, cM, L, "EQUIPAMENTO")): If IsEmpty(eq) Then eq = C("EQUIPAMENTO_PADRAO")
            mz = Txt(Campo(dM, cM, L, "MATRIZ")): If IsEmpty(mz) Then mz = C("MATRIZ_PADRAO")
            v = Campo(dM, cM, L, "LOTE")
            If EhNum(v) Then
                lt = Format$(Int(Abs(CDbl(v)) + 0.5) * IIf(CDbl(v) >= 0, 1, -1), "0")
            Else
                lt = Txt(v)
            End If
            an = Txt(Campo(dM, cM, L, "ANALITO"))
            res = Campo(dM, cM, L, "RESULTADO")
            isnum = EhNum(res)
            If isnum Then rtx = Empty Else rtx = Txt(res)
            amb = False
            If Not IsEmpty(rtx) Then amb = (InStr(rtx, ".") > 0 And InStr(rtx, ",") = 0)
            If Not amb Then
                On Error Resume Next
                If isnum Then vl = CDbl(res) Else vl = NumPt(res)
                If Err.Number <> 0 Then vl = Empty
                Err.Clear
                On Error GoTo 0
            End If
            orig = ParaOrigem(DP, nD, mz, an)
            If IsEmpty(nv) Then
                ndesc = Empty
            ElseIf nivelDesc.Exists(KDup(nv)) Then
                ndesc = nivelDesc(KDup(nv))
            Else
                ndesc = Format$(nv, "0")
            End If
            If IsEmpty(dtv) Or IsEmpty(hr) Then dh = Empty Else dh = DeMs(Ms(CDbl(dtv)) + hr)
            If IsEmpty(reg) Then regc = agora Else regc = reg
            valido = False
            If Not IsEmpty(idn) Then valido = (Left$(idn, 4) = "MAN_" And SoDigitos(Mid$(idn, 5)))
            falt = ""
            If Not valido Then falt = falt & ", ID_REGISTRO (padrao MAN_0001)"
            If IsEmpty(dtv) Then falt = falt & ", DATA"
            If IsEmpty(hr) Then falt = falt & ", HORA"
            If IsEmpty(lt) Then falt = falt & ", LOTE"
            If Not IsEmpty(dh) Then If Ms(dh) > Ms(regc) Then falt = falt & ", DATA/HORA posterior ao lancamento"
            If IsEmpty(nv) Then
                falt = falt & ", NIVEL"
            ElseIf nv < 1 Or nv > NIVEIS Then
                falt = falt & ", NIVEL"
            End If
            If IsEmpty(an) Then falt = falt & ", ANALITO"
            If amb Then falt = falt & ", RESULTADO ambiguo (use virgula decimal: 1,5)"
            If IsEmpty(vl) And Not amb Then falt = falt & ", RESULTADO"
            ' ADR-082: SUBSTITUI_ID preenchido exige MOTIVO (sem ele nada e substituido)
            subd = NormId(Campo(dM, cM, L, "SUBSTITUI_ID"))
            If Not IsEmpty(subd) Then
                If IsEmpty(Txt(Campo(dM, cM, L, "MOTIVO"))) Then falt = falt & ", MOTIVO (obrigatorio para substituir um resultado)"
            End If
            If Len(falt) > 0 Then falt = Mid$(falt, 3)
            nU = nU + 1
            U(nU, F_ID) = IIf(valido, idn, "MAN_INVALIDO_L" & L)
            U(nU, F_IDO) = Empty: U(nU, F_IDDIG) = idn
            U(nU, F_EQ) = eq: U(nU, F_MZ) = mz
            If orig > 0 Then U(nU, F_BL) = DP(orig, 4) Else U(nU, F_BL) = Empty
            U(nU, F_LT) = lt: U(nU, F_NV) = nv: U(nU, F_ND) = ndesc: U(nU, F_DT) = dtv: U(nU, F_DH) = dh
            If Not IsEmpty(dh) Then U(nU, F_DHMS) = Ms(dh)
            If orig > 0 Then U(nU, F_AO) = DP(orig, 2) Else U(nU, F_AO) = an
            U(nU, F_PD) = Empty: U(nU, F_VL) = vl: U(nU, F_FL) = Empty: U(nU, F_UN) = Txt(Campo(dM, cM, L, "UNIDADE"))
            U(nU, F_AM) = Empty: U(nU, F_IT) = Empty: U(nU, F_RE) = reg
            If Not IsEmpty(an) Then If cadastrados.Exists(an) Then U(nU, F_AQ) = an
            U(nU, F_AN) = an: U(nU, F_OR) = "MANUAL": U(nU, F_LM) = L: U(nU, F_FALTA) = falt
            U(nU, F_MOTM) = Txt(Campo(dM, cM, L, "MOTIVO")): U(nU, F_USRM) = Txt(Campo(dM, cM, L, "USUARIO"))
            U(nU, F_DUP) = False: U(nU, F_DUPDE) = Empty
            U(nU, F_SUBDIG) = subd: U(nU, F_SUBOK) = False
        End If
    Next L
    If nU < ini Then Exit Sub

    ' 1a linha valida por ID (senao a 1a); as outras ganham ~L<linha>
    Dim porId As Object, k As Variant, prim As Long, primOk As Long
    Set porId = CreateObject("Scripting.Dictionary")
    For i = ini To nU
        If Not porId.Exists(U(i, F_ID)) Then porId.Add U(i, F_ID), New Collection
        porId(U(i, F_ID)).Add i
    Next i
    For Each k In porId.Keys
        prim = 0: primOk = 0
        For Each cc In porId(k)
            If prim = 0 Or U(cc, F_LM) < prim Then prim = U(cc, F_LM)
            If U(cc, F_FALTA) = "" Then If primOk = 0 Or U(cc, F_LM) < primOk Then primOk = U(cc, F_LM)
        Next cc
        If primOk > 0 Then prim = primOk
        For Each cc In porId(k)
            U(cc, F_PRIMID) = prim
        Next cc
    Next k
    For i = ini To nU
        If U(i, F_PRIMID) <> U(i, F_LM) Then U(i, F_ID) = U(i, F_ID) & "~L" & U(i, F_LM)
    Next i

    ' chave+instante entre os manuais validos
    Dim porK As Object, km As String
    Set porK = CreateObject("Scripting.Dictionary")
    For i = ini To nU
        If U(i, F_FALTA) = "" Then
            km = ""
            For Each cc In Array(U(i, F_EQ), U(i, F_MZ), U(i, F_LT), IIf(IsEmpty(U(i, F_NV)), Empty, Format$(U(i, F_NV), "0")), _
                                U(i, F_AN), Format$(U(i, F_DH), "yyyymmddhhnnss"))
                If Not IsEmpty(cc) Then km = km & IIf(Len(km) > 0, "|", "") & cc
            Next cc
            U(i, F_KMAN) = km
            If Not porK.Exists(km) Then porK.Add km, i Else If U(i, F_LM) < U(porK(km), F_LM) Then porK(km) = i
        End If
    Next i
    Dim temSub As Boolean
    For i = ini To nU
        If U(i, F_FALTA) = "" Then
            U(i, F_PRIMK) = U(porK(U(i, F_KMAN)), F_LM)
            U(i, F_IDPRIMK) = U(porK(U(i, F_KMAN)), F_ID)
            If Not IsEmpty(U(i, F_SUBDIG)) Then temSub = True
        End If
    Next i
    ' ADR-082: substituicao escolhida pelo usuario (SUBSTITUI_ID) -- manual completo que nao repete outro manual (mesmo ID
    ' ou mesma chave e instante). Valido: ID do INTERFACEAMENTO (qualquer status), nao retransmissao, mesma chave; sem
    ' limite de horario. Entre os validos com o mesmo ID vale o de menor _LINHA_MANUAL (os manuais estao em U na ordem
    ' da aba); o seguinte fica com F_SUBREP (CONFLITO_MANUAL abaixo). Ifs aninhados: o VBA nao faz curto-circuito.
    Dim idxInt As Object, venc As Object, alvo As Long, razao As Variant
    If temSub Then
        Set idxInt = CreateObject("Scripting.Dictionary")
        For i = 1 To nInterface
            If Not idxInt.Exists(U(i, F_ID)) Then idxInt.Add U(i, F_ID), i
        Next i
        Set venc = CreateObject("Scripting.Dictionary")
        For i = ini To nU
            If U(i, F_FALTA) = "" Then
                If Not IsEmpty(U(i, F_SUBDIG)) Then
                    If U(i, F_PRIMID) = U(i, F_LM) Then
                        If U(i, F_PRIMK) = U(i, F_LM) Then
                            razao = RazaoSubstituicao(U, i, idxInt)
                            If Not IsEmpty(razao) Then
                                U(i, F_SUBR) = razao
                            ElseIf venc.Exists(U(i, F_SUBDIG)) Then
                                U(i, F_SUBREP) = U(venc(U(i, F_SUBDIG)), F_ID)
                            Else
                                venc.Add U(i, F_SUBDIG), i
                                U(i, F_SUBOK) = True
                                alvo = idxInt(U(i, F_SUBDIG))
                                U(alvo, F_SUBPOR) = U(i, F_ID): U(alvo, F_SUBMOT) = U(i, F_MOTM): U(alvo, F_SUBIDX) = i
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If
    ' interfaceamento valido por (chave + DATA), na ordem da Interface
    Dim iv As Object, ki As String
    Set iv = CreateObject("Scripting.Dictionary")
    For i = 1 To nInterface
        If Not U(i, F_DUP) And Not IsEmpty(U(i, F_VL)) And Not IsEmpty(U(i, F_DH)) Then
            ki = KDup(U(i, F_EQ)) & Chr$(0) & KDup(U(i, F_MZ)) & Chr$(0) & KDup(U(i, F_LT)) & Chr$(0) & KDup(U(i, F_NV)) & _
                 Chr$(0) & KDup(U(i, F_AN)) & Chr$(0) & KDup(CDbl(U(i, F_DT)))
            If Not iv.Exists(ki) Then iv.Add ki, New Collection
            iv(ki).Add i
        End If
    Next i
    Dim dd As Variant, melhor As Double, idMelhor As Variant, dist As Double, base As String, achou As Boolean
    For i = ini To nU
        If U(i, F_FALTA) = "" Then
            base = KDup(U(i, F_EQ)) & Chr$(0) & KDup(U(i, F_MZ)) & Chr$(0) & KDup(U(i, F_LT)) & Chr$(0) & KDup(U(i, F_NV)) & _
                   Chr$(0) & KDup(U(i, F_AN)) & Chr$(0)
            achou = False
            For Each dd In Array(0, -1, 1)
                ki = base & KDup(CDbl(U(i, F_DT)) + dd)
                If iv.Exists(ki) Then
                    For Each cc In iv(ki)
                        dist = Abs(U(cc, F_DHMS) - U(i, F_DHMS)) / 60000#
                        If Not achou Or dist < melhor Then melhor = dist: idMelhor = U(cc, F_ID): achou = True
                    Next cc
                End If
            Next dd
            If U(i, F_PRIMID) <> U(i, F_LM) Then
                U(i, F_CONFM) = "ID_REGISTRO manual repetido (vale a 1a linha)": U(i, F_CONFID) = Empty
            ElseIf U(i, F_PRIMK) <> U(i, F_LM) Then
                U(i, F_CONFM) = "resultado manual repetido (mesma chave e instante)": U(i, F_CONFID) = U(i, F_IDPRIMK)
            ElseIf Not IsEmpty(U(i, F_SUBREP)) Then
                ' ADR-082: dois manuais validos para o mesmo resultado -- vale o 1o; este nao duplica em silencio
                U(i, F_CONFM) = "SUBSTITUI_ID invalido: outro manual (" & U(i, F_SUBREP) & ") ja substitui " & U(i, F_SUBDIG) & SUF_SUB
                U(i, F_CONFID) = U(i, F_SUBREP)
            ElseIf U(i, F_SUBOK) = True Then
                U(i, F_CONFM) = Empty                     ' ADR-082: escolha explicita, isento do conflito por horario
            ElseIf achou And melhor <= TOL Then
                U(i, F_CONFM) = "coincide com resultado do interfaceamento (" & Format$(Int(melhor + 0.5), "0") & " min)"
                If Not IsEmpty(U(i, F_SUBR)) Then U(i, F_CONFM) = U(i, F_CONFM) & "; SUBSTITUI_ID invalido: " & U(i, F_SUBR) & SUF_SUB
                U(i, F_CONFID) = idMelhor
            Else
                U(i, F_CONFM) = Empty
            End If
            ' ADR-082: nota do manual ATIVO/INATIVADO (MOTIVO_EXCLUSAO_AUTOMATICA; a coluna DETALHE da aba a mostra)
            If U(i, F_SUBOK) = True Then
                U(i, F_NOTA) = "substitui o resultado do interfaceamento " & U(i, F_SUBDIG)
            ElseIf Not IsEmpty(U(i, F_SUBR)) And IsEmpty(U(i, F_CONFM)) Then
                U(i, F_NOTA) = "SUBSTITUI_ID invalido: " & U(i, F_SUBR) & SUF_SUB
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------- corrida (CORRIDA_NO_DIA) e RUN
Private Sub CalcularCorridas(ByRef U() As Variant, ByVal nU As Long, ByVal GAP As Double)
    Dim reais() As Long, nRe As Long, i As Long, j As Long, gk As String, c As Variant, chave() As String, o1() As Long
    Dim ant As Long, bl As Long, u_ As Long
    ReDim reais(1 To IIf(nU > 0, nU, 1))
    For i = 1 To nU
        U(i, F_COR) = Empty: U(i, F_RUN) = Empty
        ' ADR-082: o resultado que um manual valido substitui nunca forma corrida, inativado ou nao (o manual entra pela
        ' propria DATA_HORA); o que continua com status de corrida real recebe o RUN do manual no fim (HerdarCorrida)
        If U(i, F_FALTA) = "" And Not U(i, F_DUP) And IsEmpty(U(i, F_CONFM)) And Not IsEmpty(U(i, F_DH)) And _
           IsEmpty(U(i, F_SUBPOR)) Then
            nRe = nRe + 1: reais(nRe) = i
            gk = ""
            For Each c In Array(U(i, F_EQ), U(i, F_MZ), U(i, F_LT), U(i, F_AN), IIf(IsEmpty(U(i, F_DT)), Empty, CStr(Year(U(i, F_DT)) * 10000& + Month(U(i, F_DT)) * 100& + Day(U(i, F_DT)))))
                If Not IsEmpty(c) Then gk = gk & IIf(Len(gk) > 0, "|", "") & c
            Next c
            U(i, F_GK) = gk
        End If
    Next i
    If nRe = 0 Then Exit Sub
    ' o1: (_GK, DATA_HORA, NIVEL, ITEM_ID, ID_REGISTRO) -> blocos separados por intervalo > GAP
    ReDim chave(1 To nRe)
    For j = 1 To nRe
        i = reais(j)
        chave(j) = KS(U(i, F_GK)) & Chr$(0) & P15(U(i, F_DHMS)) & Chr$(0) & KN(U(i, F_NV)) & Chr$(0) & _
                   KN(U(i, F_IT)) & Chr$(0) & KS(U(i, F_ID))
    Next j
    o1 = OrdenarPorChave(chave, nRe)
    ant = 0
    For j = 1 To nRe
        u_ = reais(o1(j))
        If ant = 0 Then
            bl = j
        ElseIf U(u_, F_GK) <> U(ant, F_GK) Or (U(u_, F_DHMS) - U(ant, F_DHMS)) > GAP * 60000# Then
            bl = j
        End If
        U(u_, F_BLK) = bl: ant = u_
    Next j
    ' Posicao e corrida em UMA passada pela ordem o1 (forma fechada da spec): dentro de um bloco, a ordem do o1 para um
    ' mesmo nivel e (DATA_HORA, ITEM_ID, ID_REGISTRO) -- a mesma do o2 --, entao a posicao e a contagem corrente do
    ' nivel no bloco. CORRIDA_NO_DIA = soma, nos blocos anteriores do mesmo dia (_GK), da maior contagem por nivel,
    ' mais a posicao. Equivale a ordenar os segmentos (_GK, _BL, _POS) distintos e numera-los dentro do _GK.
    Dim cont As Object, kNiv As String, blAtual As Long, maxPos() As Long, offset() As Long, acc As Long, gkAnt As String
    Dim blAnt As Long
    Set cont = CreateObject("Scripting.Dictionary")
    ReDim maxPos(1 To nRe): ReDim offset(1 To nRe)
    blAtual = 0
    For j = 1 To nRe
        u_ = reais(o1(j))
        If U(u_, F_BLK) <> blAtual Then
            blAtual = U(u_, F_BLK)
            cont.RemoveAll
        End If
        kNiv = KDup(U(u_, F_NV))
        cont(kNiv) = cont(kNiv) + 1
        U(u_, F_POS) = cont(kNiv)
        If cont(kNiv) > maxPos(blAtual) Then maxPos(blAtual) = cont(kNiv)
    Next j
    blAtual = 0: blAnt = 0: acc = 0: gkAnt = Chr$(1)
    For j = 1 To nRe
        u_ = reais(o1(j))
        If U(u_, F_BLK) <> blAtual Then
            If U(u_, F_GK) <> gkAnt Then acc = 0 Else acc = acc + maxPos(blAnt)
            blAnt = U(u_, F_BLK): blAtual = blAnt
            offset(blAtual) = acc: gkAnt = U(u_, F_GK)
        End If
        U(u_, F_COR) = CDbl(offset(blAtual) + U(u_, F_POS))
        If U(u_, F_COR) > 99 Or IsEmpty(U(u_, F_DT)) Then
            U(u_, F_RUN) = Empty
        Else
            U(u_, F_RUN) = CDbl(Year(U(u_, F_DT)) Mod 100) * 1000000# + Month(U(u_, F_DT)) * 10000# + Day(U(u_, F_DT)) * 100# + U(u_, F_COR)
        End If
    Next j
    HerdarCorrida U, nU
End Sub

' ADR-082 (= o Junto do PQ): o resultado indicado que continua com status de corrida real (INATIVADO, ou SEM_VALOR
' inativado -- a inativacao prevalece) fica com a CORRIDA_NO_DIA e o RUN do manual que o substitui: o X vermelho
' aparece na corrida do manual e nenhuma outra corrida do dia muda. SUBSTITUIDO_MANUAL e SEM_DATA_HORA: sem RUN.
Private Sub HerdarCorrida(ByRef U() As Variant, ByVal nU As Long)
    Dim i As Long, m As Long
    For i = 1 To nU
        If Not IsEmpty(U(i, F_SUBPOR)) Then
            If U(i, F_ST) = "ATIVO" Or U(i, F_ST) = "INATIVADO" Or U(i, F_ST) = "SEM_VALOR" Then
                m = U(i, F_SUBIDX)
                If m > 0 Then U(i, F_COR) = U(m, F_COR): U(i, F_RUN) = U(m, F_RUN)
            End If
        End If
    Next i
End Sub
