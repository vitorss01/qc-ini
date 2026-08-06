Attribute VB_Name = "mEstatistica"
Option Explicit
' ============================================================================
'  MOTOR ESTATISTICO (Fase 3)
'  Toda a estatistica do sistema vive aqui. Nenhum UserForm e nenhuma planilha
'  calcula nada — o Excel guarda apenas o RESULTADO produzido por este modulo.
'
'  Elegibilidade (CLSI EP05 / C24): apenas resultados ELEGIVEIS compoem
'  media, DP, CV, Bias, Sigma e as regras de Westgard. Os estados vivem na aba
'  Cfg_Status; acrescentar um novo estado NAO exige alterar nenhuma rotina.
' ============================================================================
Public Const NLV As Long = 3          ' niveis deste setor
Private Const KC0 As Long = 3          ' Calc: 1a linha
Private Const NK  As Long = 180        ' Calc: nº de corridas exibiveis
Private Const CF0 As Long = 6          ' Calc: 1a coluna de nivel (F)
Private Const NFD As Long = 22         ' Calc: campos por nivel
Private Const EF0 As Long = 3          ' Eng_Saida: 1a coluna de bloco de nivel
Private Const NEF As Long = 7          ' Eng_Saida: campos por nivel
Private Const COL_FILTRO As Long = 24  ' Eng_Saida: filtro de data por corrida
Private Const COL_VALOR0 As Long = 25  ' Eng_Saida: 1a coluna de valor por nivel
Private Const COL_CHAVE As Long = 28   ' Eng_Saida: chave logica ANALITO|RUN
Private Const LINHA_STAT As Long = 185 ' Eng_Saida: 1a linha do bloco de estatistica
Private Const LINHA_EST As Long = 190  ' Eng_Saida: 1a linha da tabela de parametros
Private Const AR0 As Long = 4          ' Analitos: 1a linha
Private Const ARN As Long = 43
Private Const E0  As Long = 7          ' Estatistica: 1a linha

' ---- cache ----
Private mElig As Object                ' status -> True/False
Private mCache As Object               ' chave -> Array(n, media, dp)
Private mDB As Variant                 ' snapshot do banco
Private mDBok As Boolean
Private mReg As Variant                ' snapshot de Registros (repeticoes/calibracao)
Private mAnal As Variant               ' snapshot de Analitos (alvos)
Private mIdxAnal As Object             ' analito -> linha em mAnal
Private mAgg As Object                 ' (analito|nivel) -> agregados de Westgard

' ============================ ELEGIBILIDADE ============================
Public Sub CarregarElegibilidade()
    Dim ws As Worksheet, i As Long, s As String
    Set mElig = CreateObject("Scripting.Dictionary")
    mElig.CompareMode = 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Status")
    On Error GoTo 0
    If ws Is Nothing Then
        mElig("Ativo") = True                     ' fallback seguro
        Exit Sub
    End If
    For i = 4 To 60
        s = Trim$(CStr(ws.Cells(i, 2).Value))
        If s <> "" Then mElig(s) = (UCase$(Trim$(CStr(ws.Cells(i, 3).Value))) = "SIM")
    Next i
End Sub

' Unico ponto de verdade sobre elegibilidade estatistica.
Public Function EhElegivel(ByVal status As Variant) As Boolean
    Dim s As String
    If mElig Is Nothing Then CarregarElegibilidade
    s = Trim$(CStr(status))
    If s = "" Then EhElegivel = False: Exit Function
    If mElig.Exists(s) Then
        EhElegivel = mElig(s)
    Else
        EhElegivel = False                        ' estado desconhecido nunca entra no calculo
    End If
End Function

' ============================ CACHE ============================
Public Sub InvalidarCache()
    Set mCache = Nothing
    mDBok = False
    Set mElig = Nothing
    Set mIdxAnal = Nothing
    Set mAgg = Nothing
    mReg = Empty
    mAnal = Empty
End Sub

Private Sub GarantirDB()
    Dim wsG As Worksheet, ig As Long, nmG As String
    If Not mDBok Then
        mDB = CarregarDB()
        ' snapshots em bloco — uma leitura por aba, nunca celula a celula
        Set wsG = ThisWorkbook.Sheets("Registros")
        mReg = wsG.Range(wsG.Cells(4, 2), wsG.Cells(203, 9)).Value      ' Data..Calibrado
        Set wsG = ThisWorkbook.Sheets("Analitos")
        mAnal = wsG.Range(wsG.Cells(AR0, 1), wsG.Cells(ARN, 20)).Value  ' A..T
        Set mIdxAnal = CreateObject("Scripting.Dictionary")
        mIdxAnal.CompareMode = 1
        For ig = 1 To UBound(mAnal, 1)
            nmG = Trim$(CStr(mAnal(ig, 1)))
            If nmG <> "" Then
                If Not mIdxAnal.Exists(nmG) Then mIdxAnal.Add nmG, ig
            End If
        Next ig
        mDBok = True
    End If
    If mCache Is Nothing Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 1
    End If
End Sub

' ============================ PRIMITIVAS ============================
Public Function CalcularMedia(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double
    If n < 1 Then Exit Function
    For i = 1 To n
        s = s + v(i)
    Next i
    CalcularMedia = s / n
End Function

' DP amostral (n-1) — CLSI. Com n<2 nao ha dispersao estimavel: devolve 0.
Public Function CalcularDP(ByRef v() As Double, ByVal n As Long, ByVal media As Double) As Double
    Dim i As Long, s As Double
    If n < 2 Then Exit Function
    For i = 1 To n
        s = s + (v(i) - media) ^ 2
    Next i
    If s <= 0 Then Exit Function
    CalcularDP = Sqr(s / (n - 1))
End Function

Public Function CalcularCV(ByVal dp As Double, ByVal media As Double) As Double
    If media = 0 Then Exit Function
    CalcularCV = dp / media * 100#
End Function

Public Function CalcularBias(ByVal mediaObs As Double, ByVal alvo As Double) As Double
    If alvo = 0 Then Exit Function
    CalcularBias = (mediaObs - alvo) / alvo * 100#
End Function

Public Function CalcularErroTotal(ByVal cv As Double, ByVal bias As Double) As Double
    CalcularErroTotal = Abs(bias) + 1.65 * cv
End Function

Public Function CalcularSigma(ByVal etp As Double, ByVal bias As Double, ByVal cv As Double) As Double
    If cv = 0 Then Exit Function
    CalcularSigma = (etp - Abs(bias)) / cv
End Function

Public Function CalcularZ(ByVal valor As Double, ByVal media As Double, ByVal dp As Double) As Double
    If dp = 0 Then Exit Function
    CalcularZ = (valor - media) / dp
End Function

Public Function ClassificarSigma(ByVal s As Double) As String
    If s >= 6 Then
        ClassificarSigma = "Excelente"
    ElseIf s >= 4 Then
        ClassificarSigma = "Bom"
    ElseIf s >= 3 Then
        ClassificarSigma = "Marginal"
    Else
        ClassificarSigma = "Inaceitável"
    End If
End Function

' ============================ ALVOS (Analitos) ============================
Private Function LinhaAnalito(ByVal analito As String) As Long
    GarantirDB
    If mIdxAnal.Exists(Trim$(analito)) Then LinhaAnalito = mIdxAnal(Trim$(analito))
End Function

' Media/DP configurados (valores-alvo do lote) para o nivel.
Public Sub AlvoAnalito(ByVal analito As String, ByVal nivel As Long, _
                       ByRef alvoMedia As Double, ByRef alvoDP As Double, ByRef etp As Double)
    Dim r As Long, cm As Long, cs As Long
    alvoMedia = 0: alvoDP = 0: etp = 0
    r = LinhaAnalito(analito)
    If r = 0 Then Exit Sub
    cm = 5 + (nivel - 1) * 2          ' E,G,I
    cs = cm + 1                        ' F,H,J
    If IsNumeric(mAnal(r, cm)) Then alvoMedia = CDbl(mAnal(r, cm))
    If IsNumeric(mAnal(r, cs)) Then alvoDP = CDbl(mAnal(r, cs))
    If IsNumeric(mAnal(r, 18)) Then etp = CDbl(mAnal(r, 18))   ' R = ETp final
End Sub

' ============================ ESTATISTICA BASICA ============================
' n/media/DP dos resultados ELEGIVEIS de (analito, nivel) no lote, com filtro de ano.
Public Function EstatBasica(ByVal analito As String, ByVal nivel As Long, ByVal loteCore As String, _
                            Optional ByVal anoDe As Long = 0, Optional ByVal anoAte As Long = 0) As Variant
    Dim k As String, i As Long, n As Long, v() As Double, media As Double, dp As Double
    GarantirDB
    k = UCase$(analito) & "|" & nivel & "|" & loteCore & "|" & anoDe & "|" & anoAte
    If mCache.Exists(k) Then EstatBasica = mCache(k): Exit Function
    If IsEmpty(mDB) Then
        EstatBasica = Array(0, 0, 0): Exit Function
    End If
    ReDim v(1 To UBound(mDB, 1))
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If CLng(Val(mDB(i, COL_NIVEL))) = nivel Then
                If loteCore = "" Or NucleoLote(CStr(mDB(i, COL_LOTE))) = loteCore Then
                    If EhElegivel(mDB(i, COL_STATUS)) Then
                        If IsNumeric(mDB(i, COL_RESULT)) Then
                            If anoDe = 0 Or (IsDate(mDB(i, COL_DATA)) And _
                               Year(mDB(i, COL_DATA)) >= anoDe And Year(mDB(i, COL_DATA)) <= anoAte) Then
                                n = n + 1: v(n) = CDbl(mDB(i, COL_RESULT))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    media = CalcularMedia(v, n)
    dp = CalcularDP(v, n, media)
    EstatBasica = Array(n, media, dp)
    mCache(k) = EstatBasica
End Function

' ============================ WESTGARD ============================
' z(t, i) = z-score do nivel t na i-esima corrida (ordenada). ok(t,i) = tem dado.
' Regras padrao: 12s (alerta) e 13s/22s/R4s/41s/10x (rejeicao).
Public Sub AvaliarWestgard(ByRef z() As Double, ByRef temDado() As Boolean, ByVal nRun As Long, _
                           ByRef r13 As Variant, ByRef r22 As Variant, ByRef rR4 As Variant, _
                           ByRef r41 As Variant, ByRef r10 As Variant, ByRef a12 As Variant)
    Dim t As Long, i As Long, j As Long, cnt As Long, mn As Double, mx As Double, k As Long
    For t = 0 To NLV - 1
        For i = 1 To nRun
            If temDado(t, i) Then
                ' --- 12s (alerta) e 13s (rejeicao) ---
                If Abs(z(t, i)) > 2 Then a12(t, i) = 1
                If Abs(z(t, i)) > 3 Then r13(t, i) = 1

                ' --- 22s: 2 consecutivos no MESMO nivel, mesmo lado, |z|>2 ---
                If i >= 2 Then
                    If temDado(t, i - 1) Then
                        If (z(t, i) > 2 And z(t, i - 1) > 2) Or (z(t, i) < -2 And z(t, i - 1) < -2) Then
                            r22(t, i) = 1
                        End If
                    End If
                End If
                ' --- 22s: 2 NIVEIS na mesma corrida, mesmo lado ---
                For j = 0 To NLV - 1
                    If j <> t Then
                        If temDado(j, i) Then
                            If (z(t, i) > 2 And z(j, i) > 2) Or (z(t, i) < -2 And z(j, i) < -2) Then
                                r22(t, i) = 1
                            End If
                        End If
                    End If
                Next j

                ' --- 41s: 4 consecutivos no mesmo nivel, mesmo lado, |z|>1 ---
                If i >= 4 Then
                    cnt = 0
                    For k = 0 To 3
                        If temDado(t, i - k) Then
                            If z(t, i - k) > 1 Then cnt = cnt + 1
                        End If
                    Next k
                    If cnt = 4 Then r41(t, i) = 1
                    cnt = 0
                    For k = 0 To 3
                        If temDado(t, i - k) Then
                            If z(t, i - k) < -1 Then cnt = cnt + 1
                        End If
                    Next k
                    If cnt = 4 Then r41(t, i) = 1
                End If

                ' --- 10x: 10 consecutivos no mesmo nivel, mesmo lado da media ---
                If i >= 10 Then
                    cnt = 0
                    For k = 0 To 9
                        If temDado(t, i - k) Then
                            If z(t, i - k) > 0 Then cnt = cnt + 1
                        End If
                    Next k
                    If cnt = 10 Then r10(t, i) = 1
                    cnt = 0
                    For k = 0 To 9
                        If temDado(t, i - k) Then
                            If z(t, i - k) < 0 Then cnt = cnt + 1
                        End If
                    Next k
                    If cnt = 10 Then r10(t, i) = 1
                End If
            End If
        Next i
    Next t

    ' --- R4s: amplitude entre niveis DENTRO da mesma corrida > 4 DP ---
    For i = 1 To nRun
        mn = 1E+30: mx = -1E+30: cnt = 0
        For t = 0 To NLV - 1
            If temDado(t, i) Then
                cnt = cnt + 1
                If z(t, i) < mn Then mn = z(t, i)
                If z(t, i) > mx Then mx = z(t, i)
            End If
        Next t
        If cnt >= 2 And (mx - mn) > 4 Then
            For t = 0 To NLV - 1
                If temDado(t, i) Then rR4(t, i) = 1
            Next t
        End If
    Next i
End Sub

' ============================ MOTOR: ESCREVER Eng_Saida ============================
' Marco 2 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: esta rotina dava ClearContents na area inteira do Calc e reescrevia a
' matriz toda com valores, destruindo as 12.614 formulas da aba.
'
' AGORA: nao toca no Calc. Calcula z e Westgard e publica SO o veredicto em
' Eng_Saida. O Calc continua fazendo por formula o que a planilha faz bem --
' buscar, filtrar, ordenar e plotar -- e le as regras daqui.
'
' Motivo de existir: as formulas do Calc implementavam Westgard com definicoes
' diferentes das do motor (41s com 3 pontos em vez de 4, 10x com 6 em vez de 10,
' R4s como max>2 E min<-2 em vez de amplitude>4, e 22s so entre niveis). Duas
' camadas emitiam veredictos diferentes sobre o mesmo dado. Agora ha uma so.
'
' O nome da rotina foi mantido para nao quebrar as chamadas existentes em
' AtualizarEstatistica, RecalcularAnalitoAtual e mOperacao.
Public Sub AtualizarCalc()
    Dim ws As Worksheet, analito As String, lote As String
    Dim i As Long, t As Long, nRun As Long, r As Long
    Dim runs() As Long, dts() As Double, valor() As Double, temDado() As Boolean, z() As Double
    Dim r13 As Variant, r22 As Variant, rR4 As Variant, r41 As Variant, r10 As Variant, a12 As Variant
    Dim alvoM() As Double, alvoS() As Double, etp As Double
    Dim seen As Object, ordem As Object

    Set ws = ThisWorkbook.Sheets("Eng_Saida")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    lote = LoteAtivoCore()
    GarantirDB

    ' limpa a area de saida (colunas B em diante; a coluna A guarda os slots fixos)
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + NK - 1, COL_CHAVE)).ClearContents
    ws.Range("C1").Value = analito
    ws.Range("E1").Value = lote
    ws.Range("G1").Value = Now
    ws.Range("I1").Value = 0
    If analito = "" Or IsEmpty(mDB) Then Exit Sub

    ' ---- descobrir corridas (RUN) elegiveis do analito no lote ----
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim runs(1 To NK): ReDim dts(1 To NK)
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If NucleoLote(CStr(mDB(i, COL_LOTE))) = lote Then
                If EhElegivel(mDB(i, COL_STATUS)) Then
                    If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                        If nRun < NK Then
                            nRun = nRun + 1
                            runs(nRun) = CLng(Val(mDB(i, COL_RUN)))
                            If IsDate(mDB(i, COL_DATA)) Then dts(nRun) = CDbl(CDate(mDB(i, COL_DATA)))
                            seen.Add CStr(mDB(i, COL_RUN)), nRun
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If nRun = 0 Then Exit Sub

    ' ordenar por RUN (insercao: nRun e pequeno)
    Dim a As Long, b As Long, tmpR As Long, tmpD As Double
    For a = 2 To nRun
        tmpR = runs(a): tmpD = dts(a): b = a - 1
        Do While b >= 1
            If runs(b) <= tmpR Then Exit Do
            runs(b + 1) = runs(b): dts(b + 1) = dts(b): b = b - 1
        Loop
        runs(b + 1) = tmpR: dts(b + 1) = tmpD
    Next a
    Set ordem = CreateObject("Scripting.Dictionary")
    For i = 1 To nRun
        ordem(CStr(runs(i))) = i
    Next i

    ' ---- coletar valores por nivel ----
    ReDim valor(0 To NLV - 1, 1 To nRun)
    ReDim temDado(0 To NLV - 1, 1 To nRun)
    ReDim z(0 To NLV - 1, 1 To nRun)
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If NucleoLote(CStr(mDB(i, COL_LOTE))) = lote Then
                If EhElegivel(mDB(i, COL_STATUS)) And IsNumeric(mDB(i, COL_RESULT)) Then
                    If ordem.Exists(CStr(mDB(i, COL_RUN))) Then
                        t = CLng(Val(mDB(i, COL_NIVEL))) - 1
                        If t >= 0 And t <= NLV - 1 Then
                            r = ordem(CStr(mDB(i, COL_RUN)))
                            valor(t, r) = CDbl(mDB(i, COL_RESULT))
                            temDado(t, r) = True
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' ---- alvos e z-scores ----
    ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1)
    For t = 0 To NLV - 1
        AlvoAnalito analito, t + 1, alvoM(t), alvoS(t), etp
        For i = 1 To nRun
            If temDado(t, i) Then z(t, i) = CalcularZ(valor(t, i), alvoM(t), alvoS(t))
        Next i
    Next t

    ' ---- Westgard ----
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r22(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r41(0 To NLV - 1, 1 To nRun)
    ReDim r10(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)
    AvaliarWestgard z, temDado, nRun, r13, r22, rR4, r41, r10, a12

    ' ---- publicar em Eng_Saida ----
    ' Sem filtro de data de proposito: as regras dependem da sequencia completa
    ' de corridas. Quem recorta o periodo exibido e o Calc, na hora de plotar.
    Dim outRun() As Variant, outLvl() As Variant, rej As Boolean
    ReDim outRun(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outRun(i, 1) = runs(i)
    Next i
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + nRun - 1, 2)).Value = outRun

    ' Filtro de data e valores por nivel: e o que AtualizarPainelEng consome.
    ' Publicados aqui para que o Painel nunca precise ler o Calc -- se lesse,
    ' dependeria de o Excel ter recalculado as formulas que apontam para ca.
    Dim outFil() As Variant, outVal() As Variant, outChv() As Variant
    ReDim outFil(1 To nRun, 1 To 1)
    ReDim outVal(1 To nRun, 1 To NLV)
    ReDim outChv(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outFil(i, 1) = IIf(PassaFiltro(dts(i)), 1, 0)
        ' Chave logica da linha: ANALITO|RUN. O RUN sozinho identifica a corrida,
        ' que e compartilhada entre analitos -- casar so por ele deixava o Calc
        ' ler a linha certa do analito errado.
        outChv(i, 1) = analito & "|" & runs(i)
        For t = 0 To NLV - 1
            If temDado(t, i) Then outVal(i, t + 1) = valor(t, i) Else outVal(i, t + 1) = ""
        Next t
    Next i
    ws.Range(ws.Cells(KC0, COL_FILTRO), ws.Cells(KC0 + nRun - 1, COL_FILTRO)).Value = outFil
    ws.Range(ws.Cells(KC0, COL_VALOR0), ws.Cells(KC0 + nRun - 1, COL_VALOR0 + NLV - 1)).Value = outVal
    ws.Range(ws.Cells(KC0, COL_CHAVE), ws.Cells(KC0 + nRun - 1, COL_CHAVE)).Value = outChv

    For t = 0 To NLV - 1
        ReDim outLvl(1 To nRun, 1 To NEF)
        For i = 1 To nRun
            If temDado(t, i) Then
                outLvl(i, 1) = IIf(r13(t, i) = 1, 1, 0)
                outLvl(i, 2) = IIf(r22(t, i) = 1, 1, 0)
                outLvl(i, 3) = IIf(rR4(t, i) = 1, 1, 0)
                outLvl(i, 4) = IIf(r41(t, i) = 1, 1, 0)
                outLvl(i, 5) = IIf(r10(t, i) = 1, 1, 0)
                outLvl(i, 6) = IIf(a12(t, i) = 1, 1, 0)
                rej = (r13(t, i) = 1 Or r22(t, i) = 1 Or rR4(t, i) = 1 Or r41(t, i) = 1 Or r10(t, i) = 1)
                ' So REJEITADO/OK. O alerta 12s vai na coluna propria: o Calc usa
                ' este campo para separar as series de pontos conformes e
                ' rejeitados, e publicar "ALERTA" aqui mudaria o grafico.
                outLvl(i, 7) = IIf(rej, "REJEITADO", "OK")
            Else
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = ""
            End If
        Next i
        ws.Range(ws.Cells(KC0, EF0 + t * NEF), ws.Cells(KC0 + nRun - 1, EF0 + t * NEF + NEF - 1)).Value = outLvl
    Next t

    ws.Range("I1").Value = nRun
End Sub

' Filtro de data/trimestre do Painel.
Private Function PassaFiltro(ByVal dserial As Double) As Boolean
    Dim de As Variant, ate As Variant, d As Date, q As Long
    Dim q1 As Boolean, q2 As Boolean, q3 As Boolean, q4 As Boolean
    PassaFiltro = True
    If dserial <= 0 Then Exit Function
    d = CDate(dserial)
    On Error Resume Next
    de = ThisWorkbook.Names("filtroDe").RefersToRange.Value
    ate = ThisWorkbook.Names("filtroAte").RefersToRange.Value
    q1 = CBool(ThisWorkbook.Names("qsel1").RefersToRange.Value)
    q2 = CBool(ThisWorkbook.Names("qsel2").RefersToRange.Value)
    q3 = CBool(ThisWorkbook.Names("qsel3").RefersToRange.Value)
    q4 = CBool(ThisWorkbook.Names("qsel4").RefersToRange.Value)
    On Error GoTo 0
    If IsDate(de) Then If d < CDate(de) Then PassaFiltro = False: Exit Function
    If IsDate(ate) Then If d > CDate(ate) Then PassaFiltro = False: Exit Function
    If q1 Or q2 Or q3 Or q4 Then
        q = Int((Month(d) - 1) / 3) + 1
        If Not ((q = 1 And q1) Or (q = 2 And q2) Or (q = 3 And q3) Or (q = 4 And q4)) Then PassaFiltro = False
    End If
End Function

' Repeticoes/calibracao vindas da aba Registros (marcadores do grafico).
Private Function MarcadoresRegistro(ByVal analito As String, ByVal nivel As Long, ByVal dserial As Double) As Variant
    ' mReg: 1=Data 2=Analito 3=Nivel 4=Seq 5=Rep1 6=Rep2 7=Rep3 8=Calibrado
    Dim i As Long, res(0 To 3) As Variant, calib As Boolean
    res(0) = CVErr(xlErrNA): res(1) = CVErr(xlErrNA): res(2) = CVErr(xlErrNA): res(3) = False
    If dserial <= 0 Then MarcadoresRegistro = res: Exit Function
    GarantirDB
    For i = 1 To UBound(mReg, 1)
        If IsDate(mReg(i, 1)) Then
            If CDbl(CDate(mReg(i, 1))) = dserial Then
                If StrComp(Trim$(CStr(mReg(i, 2))), analito, 1) = 0 Then
                    If Trim$(CStr(mReg(i, 3))) = "" Or CLng(Val(mReg(i, 3))) = nivel Then
                        If IsNumeric(mReg(i, 5)) Then res(0) = CDbl(mReg(i, 5))
                        If IsNumeric(mReg(i, 6)) Then res(1) = CDbl(mReg(i, 6))
                        If IsNumeric(mReg(i, 7)) Then res(2) = CDbl(mReg(i, 7))
                    End If
                    If UCase$(Trim$(CStr(mReg(i, 8)))) = "SIM" Then calib = True
                End If
            End If
        End If
    Next i
    res(3) = calib
    MarcadoresRegistro = res
End Function


' ============================================================================
'  VIOLACOES DE WESTGARD — estrutura completa e explicavel
'  Escopo revisado: sem Trend/Slope/Deriva/Shift (pertencem a um modulo futuro
'  de Analytics). O comportamento sistematico e capturado por 22s, 41s e 10x.
'
'  Devolve, para (analito, nivel): a ULTIMA violacao da serie elegivel com
'  Regra | Classificacao | RUN | Analito | Nivel | Resultado | Media | DP | Z | Data
'  Referencia: Westgard JO, Barry PL, Hunt MR, Groth T. Clin Chem 1981;27:493-501.
'  Elegibilidade conforme CLSI C24 e EP05 (ver mEstatistica.EhElegivel).
' ============================================================================
Public Function AnalisarViolacoes(ByVal analito As String, ByVal nivel As Long, ByVal loteCore As String) As Variant
    Dim i As Long, j As Long, nS As Long
    Dim ys() As Double, rs() As Long, ds() As Double
    Dim media As Double, dp As Double, alvoM As Double, alvoS As Double, etp As Double
    Dim seen As Object, regra As String, runV As Long, valV As Double, zV As Double, dtV As Double

    GarantirDB
    ' n, media, dp, regra, RUN, valor, z, dataDeteccao, totalViolacoes
    AnalisarViolacoes = Array(0, 0, 0, "", 0, 0, 0, 0, 0)
    If IsEmpty(mDB) Then Exit Function

    ReDim ys(1 To UBound(mDB, 1)): ReDim rs(1 To UBound(mDB, 1)): ReDim ds(1 To UBound(mDB, 1))
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If CLng(Val(mDB(i, COL_NIVEL))) = nivel Then
                If loteCore = "" Or NucleoLote(CStr(mDB(i, COL_LOTE))) = loteCore Then
                    If EhElegivel(mDB(i, COL_STATUS)) And IsNumeric(mDB(i, COL_RESULT)) Then
                        If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                            seen.Add CStr(mDB(i, COL_RUN)), 1
                            nS = nS + 1
                            rs(nS) = CLng(Val(mDB(i, COL_RUN)))
                            ys(nS) = CDbl(mDB(i, COL_RESULT))
                            If IsDate(mDB(i, COL_DATA)) Then ds(nS) = CDbl(CDate(mDB(i, COL_DATA)))
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If nS = 0 Then Exit Function

    Dim tr As Long, ty As Double, td2 As Double
    For i = 2 To nS
        tr = rs(i): ty = ys(i): td2 = ds(i): j = i - 1
        Do While j >= 1
            If rs(j) <= tr Then Exit Do
            rs(j + 1) = rs(j): ys(j + 1) = ys(j): ds(j + 1) = ds(j): j = j - 1
        Loop
        rs(j + 1) = tr: ys(j + 1) = ty: ds(j + 1) = td2
    Next i

    Dim vv() As Double
    ReDim vv(1 To nS)
    For i = 1 To nS
        vv(i) = ys(i)
    Next i
    media = CalcularMedia(vv, nS)
    dp = CalcularDP(vv, nS, media)
    AlvoAnalito analito, nivel, alvoM, alvoS, etp

    ' --- avaliar Westgard sobre a serie ---
    Dim zz() As Double, tdd() As Boolean, tot As Long
    Dim q13 As Variant, q22 As Variant, qR4 As Variant, q41 As Variant, q10 As Variant, q12 As Variant
    ReDim zz(0 To NLV - 1, 1 To nS): ReDim tdd(0 To NLV - 1, 1 To nS)
    For i = 1 To nS
        If alvoS > 0 Then zz(nivel - 1, i) = CalcularZ(ys(i), alvoM, alvoS)
        tdd(nivel - 1, i) = True
    Next i
    ReDim q13(0 To NLV - 1, 1 To nS): ReDim q22(0 To NLV - 1, 1 To nS)
    ReDim qR4(0 To NLV - 1, 1 To nS): ReDim q41(0 To NLV - 1, 1 To nS)
    ReDim q10(0 To NLV - 1, 1 To nS): ReDim q12(0 To NLV - 1, 1 To nS)
    AvaliarWestgard zz, tdd, nS, q13, q22, qR4, q41, q10, q12

    For i = 1 To nS
        Dim rg As String
        rg = ""
        If Val(q13(nivel - 1, i)) = 1 Then rg = "13s"
        If Val(q22(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "22s", rg & "+22s")
        If Val(q41(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "41s", rg & "+41s")
        If Val(q10(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "10x", rg & "+10x")
        If rg <> "" Then
            tot = tot + 1
            regra = rg: runV = rs(i): valV = ys(i): zV = zz(nivel - 1, i): dtV = ds(i)
        End If
    Next i

    AnalisarViolacoes = Array(nS, media, dp, regra, runV, valV, zV, dtV, tot)
End Function

' Bloco descritivo completo da ultima violacao — consulta o modulo de conhecimento.
Public Function DetalheViolacao(ByVal analito As String, ByVal nivel As Long) As String
    Dim a As Variant, r As String, primeira As String
    a = AnalisarViolacoes(analito, nivel, LoteAtivoCore())
    If CStr(a(3)) = "" Then
        DetalheViolacao = "Sem violação de Westgard na série elegível."
        Exit Function
    End If
    primeira = Split(CStr(a(3)), "+")(0)
    r = "? Regra: " & a(3) & vbCrLf
    r = r & "Classificação: " & RegraClassificacao(primeira) & vbCrLf
    r = r & "RUN: " & a(4) & vbCrLf
    r = r & "Analito: " & analito & vbCrLf
    r = r & "Nível: " & nivel & vbCrLf
    r = r & "Resultado: " & Format(a(5), "0.0000") & vbCrLf
    r = r & "Média: " & Format(a(1), "0.0000") & vbCrLf
    r = r & "DP: " & Format(a(2), "0.0000") & vbCrLf
    r = r & "Z-Score: " & Format(a(6), "+0.00;-0.00") & vbCrLf
    If a(7) > 0 Then r = r & "Data/Hora da detecção: " & Format(CDate(a(7)), "dd/mm/yyyy") & vbCrLf
    r = r & vbCrLf & "Interpretação:" & vbCrLf & RegraInterpretacao(primeira) & vbCrLf
    r = r & vbCrLf & "Prováveis causas:" & vbCrLf & RegraCausas(primeira) & vbCrLf
    r = r & vbCrLf & "Sugestões:" & vbCrLf & RegraSugestoes(primeira)
    DetalheViolacao = r
End Function

' ============================ PAINEL ============================
' ======================= MOTOR: ESTATISTICA POR NIVEL =======================
' Marco 3 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: lia o Calc, calculava n/media/DP/CV/bias/ET/Sigma e gravava direto nas
' celulas do Painel, destruindo 45 das 48 formulas da aba.
'
' AGORA: le Eng_Saida (nunca o Calc) e publica em Eng_Saida. O Painel virou
' camada de apresentacao pura: le por formula e nao calcula mais nada.
'
' Por que parou de ler o Calc: depois do Marco 2 as colunas de regra do Calc sao
' formulas que apontam para Eng_Saida. Se o motor as lesse, passaria a depender
' de o Excel ter recalculado a planilha antes -- com Calculation manual, ou com
' a rotina chamada em sequencia dentro de outra, leria valor velho sem avisar.
' Lendo Eng_Saida, le o que o proprio motor acabou de escrever.
'
' Correcao de conta que vem junto (as formulas do Painel estavam erradas):
'   Erro Total  antes 1,65*CV + bias        agora Abs(bias) + 1,65*CV
'   Sigma       antes (ETp - bias)/CV       agora (ETp - Abs(bias))/CV
' Sem o Abs, bias negativo -- media abaixo do alvo, situacao corriqueira --
' subestimava o Erro Total e superestimava o Sigma. CLSI/Westgard: TE = |bias| +
' 1,65*CV. O motor ja estava certo; a interface e que discordava dele.
Public Sub AtualizarPainelEng()
    Dim eng As Worksheet, analito As String
    Dim t As Long, i As Long, n As Long, v() As Double, media As Double, dp As Double
    Dim cv As Double, bias As Double, et As Double, sg As Double, etp As Double
    ' alvoM/alvoS, nunca alvoM/alvoS: VBA e insensivel a maiusculas, entao "alvoS" e o
    ' mesmo token que a palavra reservada "As" e o modulo nao compila.
    Dim alvoM As Double, alvoS As Double, rejTot As Long
    Dim cnt(1 To 5) As Long, k2 As Long
    Dim flags As Variant, vals As Variant, filtros As Variant
    Dim outStat() As Variant

    ' O motor nao referencia mais a aba Painel, nem para leitura.
    Set eng = ThisWorkbook.Sheets("Eng_Saida")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))

    ' uma leitura em bloco de tudo que veio do motor
    filtros = eng.Range(eng.Cells(KC0, COL_FILTRO), eng.Cells(KC0 + NK - 1, COL_FILTRO)).Value
    vals = eng.Range(eng.Cells(KC0, COL_VALOR0), eng.Cells(KC0 + NK - 1, COL_VALOR0 + NLV - 1)).Value
    flags = eng.Range(eng.Cells(KC0, EF0), eng.Cells(KC0 + NK - 1, EF0 + NLV * NEF - 1)).Value

    ReDim outStat(1 To NLV, 1 To 21)

    For t = 0 To NLV - 1
        n = 0: ReDim v(1 To NK)
        For i = 1 To 5
            cnt(i) = 0
        Next i
        rejTot = 0

        For i = 1 To NK
            If Val(filtros(i, 1)) = 1 Then
                If IsNumeric(vals(i, t + 1)) And Trim$(CStr(vals(i, t + 1))) <> "" Then
                    n = n + 1: v(n) = CDbl(vals(i, t + 1))
                End If
                ' campos 1..5 do bloco do nivel = R13s, R22s, RR4s, R41s, R10x
                For k2 = 1 To 5
                    If Val(flags(i, t * NEF + k2)) = 1 Then cnt(k2) = cnt(k2) + 1
                Next k2
                ' campo 7 = Veredicto
                If CStr(flags(i, t * NEF + 7)) = "REJEITADO" Then rejTot = rejTot + 1
            End If
        Next i

        media = CalcularMedia(v, n)
        dp = CalcularDP(v, n, media)
        cv = CalcularCV(dp, media)
        AlvoAnalito analito, t + 1, alvoM, alvoS, etp
        bias = CalcularBias(media, alvoM)
        et = CalcularErroTotal(cv, bias)
        sg = CalcularSigma(etp, bias, cv)

        outStat(t + 1, 2) = n
        outStat(t + 1, 3) = IIf(n = 0, "", media)
        outStat(t + 1, 4) = IIf(n < 2, "", dp)
        outStat(t + 1, 5) = IIf(n < 2 Or media = 0, "", cv)
        outStat(t + 1, 6) = IIf(etp = 0, "", etp)
        outStat(t + 1, 7) = IIf(n = 0 Or alvoM = 0, "", bias)
        outStat(t + 1, 8) = IIf(n < 2 Or media = 0, "", et)
        outStat(t + 1, 9) = IIf(n < 2 Or cv = 0 Or etp = 0, "", sg)
        outStat(t + 1, 10) = IIf(n = 0, "", IIf(rejTot > 0, "REJEITADO", "OK"))
        outStat(t + 1, 11) = ""
        outStat(t + 1, 12) = ""
        For i = 1 To 5
            outStat(t + 1, 12 + i) = cnt(i)
        Next i
        outStat(t + 1, 18) = cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5)

        ' resumo: o historico completo fica na aba Eventos_Westgard
        Dim aa As Variant
        aa = AgregadoWestgard(analito, t + 1)
        If CLng(aa(0)) = 0 Then
            outStat(t + 1, 19) = Chr$(151)
            outStat(t + 1, 20) = ""
            outStat(t + 1, 21) = ""
        Else
            outStat(t + 1, 19) = aa(3) & " " & Chr$(183) & " RUN " & aa(4)
            outStat(t + 1, 20) = RegraClassificacao(Split(CStr(aa(3)), "+")(0))
            outStat(t + 1, 21) = aa(0) & "x  |  maior Z " & Format(aa(5), "+0.00;-0.00")
        End If
    Next t

    ' escrita unica do bloco A..U; a coluna A repoe o proprio numero do nivel
    For t = 1 To NLV
        outStat(t, 1) = t
    Next t
    eng.Range(eng.Cells(LINHA_STAT, 1), eng.Cells(LINHA_STAT + NLV - 1, 21)).Value = outStat
End Sub

' ============================ ABA ESTATISTICA ============================
' ========================= MOTOR: TABELA DE PARAMETROS =========================
' Marco 4 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: gravava direto em Estatistica!C7:M126, destruindo as 1.320 formulas
' daquele bloco. Nao rodava na producao porque o procedimento nao compilava
' (identificador alvoS colidindo com a palavra reservada As) -- e era exatamente
' isso que mantinha as formulas vivas la. Com o alvoS corrigido, passaria a rodar
' e a destruir a aba; por isso o desvio para Eng_Saida vem junto da correcao.
'
' AGORA: publica em Eng_Saida (linhas 190..309, colunas C..M) e a aba
' Estatistica le por INDEX(engEstat, linha, coluna).
'
' Continua lendo B3/D3/B4 da propria aba Estatistica: sao filtros que o usuario
' digita (ano inicial, ano final, lote), entrada de interface e nao calculo.
Public Sub AtualizarEstatisticaAba()
    Dim ws As Worksheet, wa As Worksheet, eng As Worksheet, i As Long, t As Long, er As Long
    Dim analito As String, anoDe As Long, anoAte As Long, loteF As String
    Dim st As Variant, n As Long, media As Double, dp As Double, cv As Double
    ' alvoM/alvoS, nunca alvoM/alvoS: VBA e insensivel a maiusculas, entao "alvoS" e o
    ' mesmo token que a palavra reservada "As" e o modulo nao compila.
    Dim alvoM As Double, alvoS As Double, etp As Double, bias As Double, et As Double, sg As Double
    Dim outp() As Variant, linhas As Long

    Set ws = ThisWorkbook.Sheets("Estatística")
    Set wa = ThisWorkbook.Sheets("Analitos")
    Set eng = ThisWorkbook.Sheets("Eng_Saida")

    anoDe = CLng(Val(ws.Range("B3").Value))
    anoAte = CLng(Val(ws.Range("D3").Value))
    If anoAte < anoDe Then anoAte = anoDe
    loteF = Trim$(CStr(ws.Range("B4").Value))

    linhas = 40 * NLV
    ReDim outp(1 To linhas, 1 To 11)     ' C..M
    er = 0
    For i = AR0 To ARN
        analito = Trim$(CStr(wa.Cells(i, 1).Value))
        For t = 1 To NLV
            er = er + 1
            If analito = "" Then
                Dim c As Long
                For c = 1 To 11
                    outp(er, c) = ""
                Next c
            Else
                st = EstatBasica(analito, t, loteF, anoDe, anoAte)
                n = st(0): media = st(1): dp = st(2)
                cv = CalcularCV(dp, media)
                AlvoAnalito analito, t, alvoM, alvoS, etp
                bias = CalcularBias(media, alvoM)
                et = CalcularErroTotal(cv, bias)
                sg = CalcularSigma(etp, bias, cv)
                outp(er, 1) = n
                outp(er, 2) = IIf(n = 0, "", media)
                outp(er, 3) = IIf(n < 2, "", dp)
                outp(er, 4) = IIf(n < 2 Or media = 0, "", cv)
                outp(er, 5) = IIf(etp = 0, "", etp)
                outp(er, 6) = IIf(IsNumeric(wa.Cells(i, 19).Value), wa.Cells(i, 19).Value, "")
                outp(er, 7) = IIf(IsNumeric(wa.Cells(i, 20).Value), wa.Cells(i, 20).Value, "")
                outp(er, 8) = IIf(n = 0 Or alvoM = 0, "", bias)
                outp(er, 9) = IIf(n < 2 Or media = 0, "", et)
                outp(er, 10) = IIf(n < 2 Or cv = 0 Or etp = 0, "", sg)
                outp(er, 11) = IIf(n < 2 Or cv = 0 Or etp = 0, "", ClassificarSigma(sg))
            End If
        Next t
    Next i

    ' publica em Eng_Saida; a aba Estatistica le por formula
    eng.Range(eng.Cells(LINHA_EST, 3), eng.Cells(LINHA_EST + linhas - 1, 13)).Value = outp

    RegistrarEventosWestgard
End Sub



' ============================================================================
'  EVENTOS_WESTGARD — historico auditavel de violacoes
'  UMA passagem no banco produz todos os eventos de todos os analitos/niveis.
'  A aba Estatística permanece dedicada a PARAMETROS; eventos vivem aqui.
'  Colunas: Data | RUN | Analito | Nível | Regra | Classificação | Resultado | Z-Score
' ============================================================================
Public Sub RegistrarEventosWestgard()
    Dim ws As Worksheet, i As Long, t As Long, j As Long, k As Long
    Dim lote As String, chave As String, nEv As Long
    Dim grupos As Object, listaCh As Object
    Dim ev() As Variant
    Dim agg As Object

    GarantirDB
    Set ws = ThisWorkbook.Sheets("Eventos_Westgard")
    ws.Range("A4:H5003").ClearContents
    If mAgg Is Nothing Then Set mAgg = CreateObject("Scripting.Dictionary")
    If IsEmpty(mDB) Then Exit Sub
    lote = LoteAtivoCore()

    ' ---- agrupar a serie elegivel por (analito|nivel) numa unica varredura ----
    Set grupos = CreateObject("Scripting.Dictionary")
    Set listaCh = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(mDB, 1)
        If Len(Trim$(CStr(mDB(i, COL_ANALITO)))) > 0 Then
            If NucleoLote(CStr(mDB(i, COL_LOTE))) = lote Then
                If EhElegivel(mDB(i, COL_STATUS)) And IsNumeric(mDB(i, COL_RESULT)) Then
                    chave = UCase$(Trim$(CStr(mDB(i, COL_ANALITO)))) & "|" & CLng(Val(mDB(i, COL_NIVEL)))
                    If Not grupos.Exists(chave) Then
                        grupos.Add chave, New Collection
                        listaCh.Add chave, Trim$(CStr(mDB(i, COL_ANALITO))) & "|" & CLng(Val(mDB(i, COL_NIVEL)))
                    End If
                    ' RUN | valor | data
                    grupos(chave).Add Array(CLng(Val(mDB(i, COL_RUN))), CDbl(mDB(i, COL_RESULT)), _
                                            IIf(IsDate(mDB(i, COL_DATA)), CDbl(CDate(mDB(i, COL_DATA))), 0))
                End If
            End If
        End If
    Next i
    If grupos.Count = 0 Then Set mAgg = CreateObject("Scripting.Dictionary"): Exit Sub

    Set agg = CreateObject("Scripting.Dictionary")
    Set mAgg = agg          ' definido JA — AgregadoWestgard nunca re-dispara esta rotina
    ReDim ev(1 To 5000, 1 To 8)
    nEv = 0

    Dim ch As Variant
    For Each ch In grupos.Keys
        Dim col As Collection, nS As Long
        Set col = grupos(ch)
        nS = col.Count
        If nS > 0 Then
            Dim rs() As Long, ys() As Double, ds() As Double, seenR As Object
            ReDim rs(1 To nS): ReDim ys(1 To nS): ReDim ds(1 To nS)
            Set seenR = CreateObject("Scripting.Dictionary")
            Dim m As Long, item As Variant
            m = 0
            For j = 1 To nS
                item = col(j)                      ' resolve o membro padrao ANTES de indexar
                If Not seenR.Exists(CStr(item(0))) Then
                    seenR.Add CStr(item(0)), 1
                    m = m + 1
                    rs(m) = item(0): ys(m) = item(1): ds(m) = item(2)
                End If
            Next j
            nS = m
            If nS > 0 Then
                ' ordenar por RUN
                Dim tr As Long, ty As Double, tdd As Double
                For j = 2 To nS
                    tr = rs(j): ty = ys(j): tdd = ds(j): k = j - 1
                    Do While k >= 1
                        If rs(k) <= tr Then Exit Do
                        rs(k + 1) = rs(k): ys(k + 1) = ys(k): ds(k + 1) = ds(k): k = k - 1
                    Loop
                    rs(k + 1) = tr: ys(k + 1) = ty: ds(k + 1) = tdd
                Next j

                Dim partes As Variant, analitoN As String, nivelN As Long
                partes = Split(CStr(listaCh(ch)), "|")
                analitoN = partes(0): nivelN = CLng(partes(1))

                Dim alvoM As Double, alvoS As Double, etp As Double
                AlvoAnalito analitoN, nivelN, alvoM, alvoS, etp

                Dim zz() As Double, td() As Boolean
                Dim q13 As Variant, q22 As Variant, qR4 As Variant, q41 As Variant, q10 As Variant, q12 As Variant
                ReDim zz(0 To 0, 1 To nS): ReDim td(0 To 0, 1 To nS)
                For j = 1 To nS
                    If alvoS > 0 Then zz(0, j) = CalcularZ(ys(j), alvoM, alvoS)
                    td(0, j) = True
                Next j
                ReDim q13(0 To 0, 1 To nS): ReDim q22(0 To 0, 1 To nS)
                ReDim qR4(0 To 0, 1 To nS): ReDim q41(0 To 0, 1 To nS)
                ReDim q10(0 To 0, 1 To nS): ReDim q12(0 To 0, 1 To nS)
                AvaliarWestgard1N zz, td, nS, q13, q22, qR4, q41, q10, q12

                Dim nv As Long, priR As String, priRun As Long, ultR As String, ultRun As Long, maxZ As Double
                nv = 0: maxZ = 0
                For j = 1 To nS
                    Dim regs As String
                    regs = ""
                    If Val(q13(0, j)) = 1 Then regs = "13s"
                    If Val(q22(0, j)) = 1 Then regs = IIf(regs = "", "22s", regs & "+22s")
                    If Val(q41(0, j)) = 1 Then regs = IIf(regs = "", "41s", regs & "+41s")
                    If Val(q10(0, j)) = 1 Then regs = IIf(regs = "", "10x", regs & "+10x")
                    If regs <> "" And nEv < 5000 Then
                        nEv = nEv + 1
                        ev(nEv, 1) = IIf(ds(j) > 0, CDate(ds(j)), "")
                        ev(nEv, 2) = rs(j)
                        ev(nEv, 3) = analitoN
                        ev(nEv, 4) = nivelN
                        ev(nEv, 5) = regs
                        ev(nEv, 6) = RegraClassificacao(Split(regs, "+")(0))
                        ev(nEv, 7) = ys(j)
                        ev(nEv, 8) = zz(0, j)
                        nv = nv + 1
                        If priR = "" Then priR = regs: priRun = rs(j)
                        ultR = regs: ultRun = rs(j)
                        If Abs(zz(0, j)) > Abs(maxZ) Then maxZ = zz(0, j)
                    End If
                Next j
                agg(UCase$(analitoN) & "|" & nivelN) = Array(nv, priR, priRun, ultR, ultRun, maxZ)
            End If
        End If
    Next ch

    If nEv > 0 Then
        Dim outp() As Variant
        ReDim outp(1 To nEv, 1 To 8)
        For i = 1 To nEv
            For j = 1 To 8
                outp(i, j) = ev(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + nEv, 8)).Value = outp
    End If
    ws.Range("J2").Value = nEv
    Set mAgg = agg
End Sub

' Agregados por analito/nivel a partir dos eventos ja calculados (sem re-varredura).
' Devolve Array(nViolacoes, primeiraRegra, primeiroRUN, ultimaRegra, ultimoRUN, maiorZ)
Public Function AgregadoWestgard(ByVal analito As String, ByVal nivel As Long) As Variant
    Dim k As String
    If mAgg Is Nothing Then RegistrarEventosWestgard
    k = UCase$(Trim$(analito)) & "|" & nivel
    If mAgg.Exists(k) Then
        AgregadoWestgard = mAgg(k)
    Else
        AgregadoWestgard = Array(0, "", 0, "", 0, 0)
    End If
End Function

' Avaliacao de Westgard restrita a UM nivel (series independentes por analito/nivel).
' 22s intra-corrida entre niveis e R4s sao avaliados no motor do Calc, que tem a
' visao multi-nivel; aqui tratamos a serie temporal do proprio nivel.
Public Sub AvaliarWestgard1N(ByRef z() As Double, ByRef temDado() As Boolean, ByVal nRun As Long, _
                             ByRef r13 As Variant, ByRef r22 As Variant, ByRef rR4 As Variant, _
                             ByRef r41 As Variant, ByRef r10 As Variant, ByRef a12 As Variant)
    Dim i As Long, k As Long, cnt As Long
    For i = 1 To nRun
        If temDado(0, i) Then
            If Abs(z(0, i)) > 2 Then a12(0, i) = 1
            If Abs(z(0, i)) > 3 Then r13(0, i) = 1
            If i >= 2 Then
                If temDado(0, i - 1) Then
                    If (z(0, i) > 2 And z(0, i - 1) > 2) Or (z(0, i) < -2 And z(0, i - 1) < -2) Then r22(0, i) = 1
                End If
            End If
            If i >= 4 Then
                cnt = 0
                For k = 0 To 3
                    If z(0, i - k) > 1 Then cnt = cnt + 1
                Next k
                If cnt = 4 Then r41(0, i) = 1
                cnt = 0
                For k = 0 To 3
                    If z(0, i - k) < -1 Then cnt = cnt + 1
                Next k
                If cnt = 4 Then r41(0, i) = 1
            End If
            If i >= 10 Then
                cnt = 0
                For k = 0 To 9
                    If z(0, i - k) > 0 Then cnt = cnt + 1
                Next k
                If cnt = 10 Then r10(0, i) = 1
                cnt = 0
                For k = 0 To 9
                    If z(0, i - k) < 0 Then cnt = cnt + 1
                Next k
                If cnt = 10 Then r10(0, i) = 1
            End If
        End If
    Next i
End Sub

Public Sub AbrirEventosWestgard()
    On Error Resume Next
    ThisWorkbook.Sheets("Eventos_Westgard").Activate
End Sub

' ============================ ORQUESTRACAO ============================
Public Sub AtualizarEstatistica()
    Application.ScreenUpdating = False
    InvalidarCache
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEstatisticaAba
    AtualizarEixos
    Application.ScreenUpdating = True
End Sub

' Recalculo incremental: so o analito atualmente exibido (nao varre a planilha toda).
Public Sub RecalcularAnalitoAtual()
    Application.ScreenUpdating = False
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEixos
    Application.ScreenUpdating = True
End Sub


