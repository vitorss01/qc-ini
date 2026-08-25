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

' ---- Westgard por detector (ADR-042) ----
' Declaracao de MODULO: em VBA tudo isto tem de vir antes da primeira
' procedure. Ficaram no meio do arquivo quando o motor foi substituido, e
' o resultado foi 'variavel nao definida' em gNaoAval.
Private Const DETECTORES As String = _
    "BIOQUIMICA|2|1_3s|INDIVIDUAL|1|1|1;" & _
    "BIOQUIMICA|2|2_2s|WITHIN_RUN|1|2|1;" & _
    "BIOQUIMICA|2|2_2s|ACROSS_RUN_SAME_LEVEL|1|1|2;" & _
    "BIOQUIMICA|2|R_4s|WITHIN_RUN|1|2|1;" & _
    "BIOQUIMICA|2|4_1s|N2_R2|1|2|2;" & _
    "BIOQUIMICA|2|4_1s|SAME_LEVEL_R4|0|1|4;" & _
    "BIOQUIMICA|2|8x|N2_R4|1|2|4;" & _
    "BIOQUIMICA|2|8x|SAME_LEVEL_R8|0|1|8;" & _
    "HEMATOLOGIA|3|1_3s|INDIVIDUAL|1|1|1;" & _
    "HEMATOLOGIA|3|2of3_2s|WITHIN_RUN|1|3|1;" & _
    "HEMATOLOGIA|3|2of3_2s|SAME_LEVEL_R3|0|1|3;" & _
    "HEMATOLOGIA|3|R_4s|WITHIN_RUN|1|3|1;" & _
    "HEMATOLOGIA|3|3_1s|N3_R1|1|3|1;" & _
    "HEMATOLOGIA|3|3_1s|SAME_LEVEL_R3|0|1|3;" & _
    "HEMATOLOGIA|3|6x|N3_R2|1|3|2;" & _
    "HEMATOLOGIA|3|6x|SAME_LEVEL_R6|0|1|6"

Private gTrace As Object            ' evidencias da ultima avaliacao
Private gNaoAval As Object          ' regra|detector -> janelas nao avaliaveis

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

' A CLASSIFICACAO DE SIGMA MORA EM mQualidade (ADR-043)
'
' Existia aqui uma segunda escada, com TRES defeitos: >=6 devolvia
' "Excelente" em vez de "Classe mundial"; a faixa 5 a <6 nao existia, entao
' Sigma 5,5 caia em "Bom"; e abaixo de 3 dizia "Inaceitavel" enquanto o resto
' do projeto diz "Desempenho inadequado".
'
' Pior que os defeitos: ela vencia a canonica. AtualizarEstatisticaAba
' chamava ClassificarSigma SEM QUALIFICAR, e chamada nao qualificada resolve
' para a funcao do PROPRIO modulo -- entao a coluna de classificacao dos dois
' produtos vinha daqui, e nao de mQualidade, sem ninguem perceber.
'
' Removida. Quem precisa da classificacao chama mQualidade.ClassificarSigma,
' que le as faixas de Cfg_PlanoQC e cai numa escada de reserva identica.

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
                If loteCore = "" Or Mid$(CStr(mDB(i, COL_LOTE)), 4, 6) = loteCore Then
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
' ===========================================================================
' MOTOR WESTGARD POR DETECTOR E ESCOPO (ADR-042)
' ===========================================================================
'
' O QUE MUDOU, E POR QUE
'
' A versao anterior avaliava "a regra 8x" como oito CORRIDAS consecutivas do
' mesmo lado da media, num unico nivel. Isso e o detector LONGITUDINAL. A
' matriz Sigma da Bioquimica pede outra coisa: N=2, R=4 -- dois niveis por
' corrida, quatro corridas, oito OBSERVACOES. Sao janelas diferentes, com
' sensibilidade diferente, e a versao anterior chamava uma pelo nome da outra.
' O mesmo valia para 6x: a matriz de tres niveis pede N=3, R=2.
'
' ESCOPOS
'
'   WITHIN_RUN ......... so a corrida atual
'   ACROSS_RUN ......... usa corridas diferentes
'   WITHIN_MATERIAL .... o mesmo nivel ao longo do tempo
'   ACROSS_MATERIALS ... combina niveis diferentes
'
' R_4s e WITHIN_RUN e ponto. Amplitude de 4 DP entre a corrida de ontem e a
' de hoje NAO e R_4s -- e deriva, que outras regras pegam. Aplicar R_4s
' atraves de corridas inventa rejeicao que o metodo nao cometeu.
'
' OFICIAL x COMPLEMENTAR
'
' Detectores adicionais NAO entram automaticamente na decisao. Somar
' longitudinal com N/R multiplica as oportunidades de rejeicao e sobe a
' probabilidade de falsa rejeicao sem ninguem decidir isso. DetectorAtivo()
' e a unica porta: o que nao esta la e calculado, registrado e NAO consolidado.
'
' NAO_AVALIAVEL x FALSE
'
' FALSE quer dizer "avaliei e nao violou". Janela sem os dados que a regra
' exige nao e FALSE -- e ausencia de avaliacao, e some do QA se for tratada
' como aprovacao. 6x N3/R2 exige tres niveis validos nas DUAS corridas; com
' N3 faltando numa delas, a regra nao foi avaliada.

' A tabela de detectores. UMA fonte: o motor consulta, e Cfg_Westgard_Escopo
' e materializada a partir daqui -- planilha e codigo nao podem divergir.
'   Area|Niveis|Regra|Detector|Ativo|N|R



Public Function AreaDoProduto() As String
    If NLV >= 3 Then AreaDoProduto = "HEMATOLOGIA" Else AreaDoProduto = "BIOQUIMICA"
End Function


' A tabela inteira, para a planilha de configuracao e para o QA.
Public Function DetectoresWestgard() As String
    DetectoresWestgard = DETECTORES
End Function


' Este detector participa da DECISAO oficial deste produto?
Public Function DetectorAtivo(ByVal regra As String, ByVal detector As String) As Boolean
    Dim p As Variant, c As Variant, x As Variant
    p = Split(DETECTORES, ";")
    For Each x In p
        c = Split(CStr(x), "|")
        If UCase$(CStr(c(0))) = AreaDoProduto() Then
            If UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And _
               UCase$(CStr(c(3))) = UCase$(Trim$(detector)) Then
                DetectorAtivo = (CStr(c(4)) = "1")
                Exit Function
            End If
        End If
    Next x
End Function


' N e R declarados para um detector, no formato "N=2 R=4".
Public Function EscalaDoDetector(ByVal regra As String, ByVal detector As String) As String
    Dim c As Variant, x As Variant
    For Each x In Split(DETECTORES, ";")
        c = Split(CStr(x), "|")
        If UCase$(CStr(c(0))) = AreaDoProduto() Then
            If UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And _
               UCase$(CStr(c(3))) = UCase$(Trim$(detector)) Then
                EscalaDoDetector = "N=" & CStr(c(5)) & " R=" & CStr(c(6))
                Exit Function
            End If
        End If
    Next x
End Function


Private Sub Evid(ByVal regra As String, ByVal detector As String, ByVal escopo As String, _
                 ByVal run As Long, ByVal detalhe As String)
    If gTrace Is Nothing Then Set gTrace = New Collection
    Dim oficial As String
    If DetectorAtivo(regra, detector) Then oficial = "OFICIAL" Else oficial = "COMPLEMENTAR"
    gTrace.Add regra & "|" & detector & "|" & escopo & "|" & oficial & _
               "|run=" & CStr(run) & "|" & EscalaDoDetector(regra, detector) & _
               "|" & detalhe
End Sub


Private Sub NaoAvaliavel(ByVal regra As String, ByVal detector As String)
    If gNaoAval Is Nothing Then
        Set gNaoAval = CreateObject("Scripting.Dictionary")
        gNaoAval.CompareMode = 1
    End If
    Dim k As String
    k = regra & "|" & detector
    If gNaoAval.Exists(k) Then
        gNaoAval(k) = CLng(gNaoAval(k)) + 1
    Else
        gNaoAval.Add k, 1
    End If
End Sub


' As evidencias da ultima avaliacao, uma por linha.
Public Function TraceWestgard() As String
    Dim s As String, x As Variant
    If gTrace Is Nothing Then Exit Function
    For Each x In gTrace
        If Len(s) > 0 Then s = s & vbLf
        s = s & CStr(x)
    Next x
    TraceWestgard = s
End Function


' Quantas janelas ficaram sem avaliacao por falta de dado, por detector.
Public Function NaoAvaliaveisWestgard() As String
    Dim s As String, k As Variant
    If gNaoAval Is Nothing Then Exit Function
    For Each k In gNaoAval.Keys
        If Len(s) > 0 Then s = s & vbLf
        s = s & CStr(k) & "=" & CStr(gNaoAval(k))
    Next k
    NaoAvaliaveisWestgard = s
End Function


' Todos os niveis da corrida tem dado?
Private Function CorridaCompleta(ByRef td() As Boolean, ByVal i As Long) As Boolean
    Dim t As Long
    For t = 0 To NLV - 1
        If Not td(t, i) Then Exit Function
    Next t
    CorridaCompleta = True
End Function


Public Sub AvaliarWestgard(ByRef z() As Double, ByRef temDado() As Boolean, ByVal nRun As Long, _
                           ByRef r13 As Variant, ByRef r22 As Variant, ByRef rR4 As Variant, _
                           ByRef r41 As Variant, ByRef r10 As Variant, ByRef a12 As Variant)
    Set gTrace = New Collection
    Set gNaoAval = CreateObject("Scripting.Dictionary")
    gNaoAval.CompareMode = 1

    Det_1_3s z, temDado, nRun, r13, a12
    Det_R_4s z, temDado, nRun, rR4

    If NLV >= 3 Then
        Det_2of3_2s_WithinRun z, temDado, nRun, r22
        Det_3_1s_N3R1 z, temDado, nRun, r41
        Det_6x_N3R2 z, temDado, nRun, r10
        ' Complementares: calculados e registrados, NAO consolidados.
        Det_MesmoNivel_Sequencia z, temDado, nRun, "2of3_2s", "SAME_LEVEL_R3", 3, 2#, r22
        Det_MesmoNivel_Sequencia z, temDado, nRun, "3_1s", "SAME_LEVEL_R3", 3, 1#, r41
        Det_MesmoNivel_Lado z, temDado, nRun, "6x", "SAME_LEVEL_R6", 6, r10
    Else
        Det_2_2s_WithinRun z, temDado, nRun, r22
        Det_2_2s_AcrossRun z, temDado, nRun, r22
        Det_4_1s_N2R2 z, temDado, nRun, r41
        Det_8x_N2R4 z, temDado, nRun, r10
        Det_MesmoNivel_Sequencia z, temDado, nRun, "4_1s", "SAME_LEVEL_R4", 4, 1#, r41
        Det_MesmoNivel_Lado z, temDado, nRun, "8x", "SAME_LEVEL_R8", 8, r10
    End If
End Sub


' --- 1_3s: INDIVIDUAL. Nao depende de outro nivel nem da corrida anterior.
Private Sub Det_1_3s(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                     ByRef r13 As Variant, ByRef a12 As Variant)
    Dim t As Long, i As Long
    For t = 0 To NLV - 1
        For i = 1 To nRun
            If td(t, i) Then
                If Abs(z(t, i)) > 2 Then a12(t, i) = 1
                If Abs(z(t, i)) > 3 Then
                    r13(t, i) = 1
                    Evid "1_3s", "INDIVIDUAL", "WITHIN_RUN", i, _
                         "N" & CStr(t + 1) & " z=" & Format$(z(t, i), "0.00")
                End If
            End If
        Next i
    Next t
End Sub


' --- R_4s: WITHIN_RUN ONLY, todos os pares de niveis da MESMA corrida.
' Com tres niveis sao N1xN2, N1xN3 e N2xN3. O par fica registrado com o menor
' indice primeiro, para N1xN3 e N3xN1 nao virarem duas violacoes.
Private Sub Det_R_4s(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                     ByRef rR4 As Variant)
    Dim i As Long, t As Long, j As Long, comDado As Long
    For i = 1 To nRun
        comDado = 0
        For t = 0 To NLV - 1
            If td(t, i) Then comDado = comDado + 1
        Next t
        If comDado < 2 Then
            NaoAvaliavel "R_4s", "WITHIN_RUN"
            GoTo proximaRun
        End If
        For t = 0 To NLV - 2
            For j = t + 1 To NLV - 1
                If td(t, i) And td(j, i) Then
                    If Abs(z(t, i) - z(j, i)) > 4 Then
                        rR4(t, i) = 1
                        rR4(j, i) = 1
                        Evid "R_4s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, _
                             "N" & CStr(t + 1) & "xN" & CStr(j + 1) & _
                             " z=" & Format$(z(t, i), "0.00") & _
                             "/" & Format$(z(j, i), "0.00")
                    End If
                End If
            Next j
        Next t
proximaRun:
    Next i
End Sub


' --- 2_2s WITHIN_RUN (N=2): os dois niveis da corrida, mesmo lado, alem de 2s.
Private Sub Det_2_2s_WithinRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                               ByRef r22 As Variant)
    Dim i As Long
    For i = 1 To nRun
        If Not CorridaCompleta(td, i) Then
            NaoAvaliavel "2_2s", "WITHIN_RUN"
        ElseIf (z(0, i) > 2 And z(1, i) > 2) Or (z(0, i) < -2 And z(1, i) < -2) Then
            r22(0, i) = 1: r22(1, i) = 1
            Evid "2_2s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, _
                 "N1=" & Format$(z(0, i), "0.00") & " N2=" & Format$(z(1, i), "0.00")
        End If
    Next i
End Sub


' --- 2_2s ACROSS_RUN, MESMO nivel: duas corridas seguidas alem do mesmo 2s.
Private Sub Det_2_2s_AcrossRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                               ByRef r22 As Variant)
    Dim t As Long, i As Long
    For t = 0 To NLV - 1
        For i = 2 To nRun
            If Not (td(t, i) And td(t, i - 1)) Then
                NaoAvaliavel "2_2s", "ACROSS_RUN_SAME_LEVEL"
            ElseIf (z(t, i) > 2 And z(t, i - 1) > 2) Or _
                   (z(t, i) < -2 And z(t, i - 1) < -2) Then
                r22(t, i) = 1
                Evid "2_2s", "ACROSS_RUN_SAME_LEVEL", "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     "N" & CStr(t + 1) & " " & Format$(z(t, i - 1), "0.00") & _
                     "->" & Format$(z(t, i), "0.00")
            End If
        Next i
    Next t
End Sub


' --- 2of3_2s WITHIN_RUN: dois dos tres niveis da corrida alem do MESMO 2s.
' Lados opostos NAO contam -- isso e R_4s, fenomeno diferente.
Private Sub Det_2of3_2s_WithinRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                  ByRef r22 As Variant)
    Dim i As Long, t As Long, acima As Long, abaixo As Long, comDado As Long
    For i = 1 To nRun
        acima = 0: abaixo = 0: comDado = 0
        For t = 0 To NLV - 1
            If td(t, i) Then
                comDado = comDado + 1
                If z(t, i) > 2 Then acima = acima + 1
                If z(t, i) < -2 Then abaixo = abaixo + 1
            End If
        Next t
        If comDado < 2 Then
            NaoAvaliavel "2of3_2s", "WITHIN_RUN"
        ElseIf acima >= 2 Or abaixo >= 2 Then
            Dim lado As Long, det As String
            If acima >= 2 Then lado = 1 Else lado = -1
            det = ""
            For t = 0 To NLV - 1
                If td(t, i) Then
                    If (lado = 1 And z(t, i) > 2) Or (lado = -1 And z(t, i) < -2) Then
                        r22(t, i) = 1
                        det = det & "N" & CStr(t + 1) & "=" & Format$(z(t, i), "0.00") & " "
                    End If
                End If
            Next t
            Evid "2of3_2s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, Trim$(det)
        End If
    Next i
End Sub


' --- 3_1s N3/R1: os TRES niveis da mesma corrida alem do mesmo 1s.
Private Sub Det_3_1s_N3R1(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                          ByRef r41 As Variant)
    Dim i As Long, t As Long, acima As Long, abaixo As Long
    For i = 1 To nRun
        If Not CorridaCompleta(td, i) Then
            NaoAvaliavel "3_1s", "N3_R1"
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For t = 0 To NLV - 1
            If z(t, i) > 1 Then acima = acima + 1
            If z(t, i) < -1 Then abaixo = abaixo + 1
        Next t
        If acima = NLV Or abaixo = NLV Then
            Dim det As String
            det = ""
            For t = 0 To NLV - 1
                r41(t, i) = 1
                det = det & "N" & CStr(t + 1) & "=" & Format$(z(t, i), "0.00") & " "
            Next t
            Evid "3_1s", "N3_R1", "WITHIN_RUN_ACROSS_MATERIALS", i, Trim$(det)
        End If
prox:
    Next i
End Sub


' --- 4_1s N2/R2: dois niveis x duas corridas = quatro observacoes.
Private Sub Det_4_1s_N2R2(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                          ByRef r41 As Variant)
    Dim i As Long, t As Long, k As Long, acima As Long, abaixo As Long
    For i = 2 To nRun
        If Not (CorridaCompleta(td, i) And CorridaCompleta(td, i - 1)) Then
            NaoAvaliavel "4_1s", "N2_R2"
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For k = i - 1 To i
            For t = 0 To NLV - 1
                If z(t, k) > 1 Then acima = acima + 1
                If z(t, k) < -1 Then abaixo = abaixo + 1
            Next t
        Next k
        If acima = NLV * 2 Or abaixo = NLV * 2 Then
            For t = 0 To NLV - 1
                r41(t, i) = 1
            Next t
            Evid "4_1s", "N2_R2", "ACROSS_RUN_ACROSS_MATERIALS", i, _
                 "runs " & CStr(i - 1) & ".." & CStr(i) & " " & CStr(NLV * 2) & " obs"
        End If
prox:
    Next i
End Sub


' --- 8x N2/R4: dois niveis x quatro corridas = oito observacoes do mesmo
' LADO DA MEDIA. E regra de deslocamento, nao de tendencia: nao exige que os
' valores crescam, so que fiquem do mesmo lado.
Private Sub Det_8x_N2R4(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                        ByRef r10 As Variant)
    Det_JanelaNR z, td, nRun, "8x", "N2_R4", 4, r10
End Sub


' --- 6x N3/R2: tres niveis x duas corridas = seis observacoes do mesmo lado.
Private Sub Det_6x_N3R2(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                        ByRef r10 As Variant)
    Det_JanelaNR z, td, nRun, "6x", "N3_R2", 2, r10
End Sub


' Janela de nRunsJanela corridas COMPLETAS: todas as NLV*nRunsJanela
' observacoes do mesmo lado da media.
Private Sub Det_JanelaNR(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                         ByVal regra As String, ByVal detector As String, _
                         ByVal nRunsJanela As Long, ByRef saida As Variant)
    Dim i As Long, k As Long, t As Long, acima As Long, abaixo As Long, total As Long
    total = NLV * nRunsJanela
    For i = nRunsJanela To nRun
        Dim completa As Boolean
        completa = True
        For k = i - nRunsJanela + 1 To i
            If Not CorridaCompleta(td, k) Then completa = False
        Next k
        If Not completa Then
            NaoAvaliavel regra, detector
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For k = i - nRunsJanela + 1 To i
            For t = 0 To NLV - 1
                If z(t, k) > 0 Then acima = acima + 1
                If z(t, k) < 0 Then abaixo = abaixo + 1
            Next t
        Next k
        If acima = total Or abaixo = total Then
            For t = 0 To NLV - 1
                saida(t, i) = 1
            Next t
            Dim sinal As String
            If acima = total Then sinal = String$(total, "+") Else sinal = String$(total, "-")
            Evid regra, detector, "ACROSS_RUN_ACROSS_MATERIALS", i, _
                 "runs " & CStr(i - nRunsJanela + 1) & ".." & CStr(i) & " " & sinal
        End If
prox:
    Next i
End Sub


' --- COMPLEMENTAR: n corridas seguidas do mesmo lado, no MESMO nivel.
' Registrado e nao consolidado, salvo se DetectorAtivo disser o contrario.
Private Sub Det_MesmoNivel_Lado(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                ByVal regra As String, ByVal detector As String, _
                                ByVal n As Long, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > 0 Then acima = acima + 1
                If z(t, k) < 0 Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub


' --- COMPLEMENTAR: n corridas seguidas alem do mesmo limite, no MESMO nivel.
Private Sub Det_MesmoNivel_Sequencia(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                     ByVal regra As String, ByVal detector As String, _
                                     ByVal n As Long, ByVal limite As Double, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > limite Then acima = acima + 1
                If z(t, k) < -limite Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub


' Os nomes da matriz vigente neste produto, na ordem das saidas de
' AvaliarWestgard. Existe para que Painel, Estatistica e a camada de BI leiam
' os rotulos do MESMO lugar que decide as regras -- uma lista escrita a mao
' noutro modulo divergiria no primeiro ajuste.
Public Function MatrizWestgard() As String
    If NLV >= 3 Then
        MatrizWestgard = "1_3s;2of3_2s;R_4s;3_1s;6x"
    Else
        MatrizWestgard = "1_3s;2_2s;R_4s;4_1s;8x"
    End If
End Function


' O nome da regra que ocupa uma posicao das saidas (1..5).
Public Function NomeRegraWestgard(ByVal pos As Long) As String
    Dim p As Variant
    p = Split(MatrizWestgard(), ";")
    If pos >= 1 And pos <= UBound(p) + 1 Then NomeRegraWestgard = p(pos - 1)
End Function


' A matriz ativa pertence mesmo a este produto? (ADR-041, secao 8)
'
' Devolve "" quando esta coerente, ou a descricao do problema. Existe porque a
' contaminacao entre modulos e silenciosa: uma matriz de dois niveis rodando
' num setor de tres continua produzindo numeros plausiveis -- so que com a
' sensibilidade e a taxa de falsa rejeicao do desenho errado.
'
' Verifica os DOIS sentidos. Faltar regra da propria familia e tao defeito
' quanto sobrar regra da outra.
Public Function ValidarMatrizWestgard() As String
    Dim m As String, faltando As String, intrusas As String
    Dim proprias As Variant, alheias As Variant, x As Variant

    m = ";" & MatrizWestgard() & ";"

    If NLV >= 3 Then
        proprias = Array("1_3s", "2of3_2s", "R_4s", "3_1s", "6x")
        alheias = Array("2_2s", "4_1s", "8x", "10x")
    Else
        proprias = Array("1_3s", "2_2s", "R_4s", "4_1s", "8x")
        alheias = Array("2of3_2s", "3_1s", "6x", "10x")
    End If

    For Each x In proprias
        If InStr(1, m, ";" & x & ";", vbTextCompare) = 0 Then _
            faltando = faltando & x & " "
    Next x
    For Each x In alheias
        If InStr(1, m, ";" & x & ";", vbTextCompare) > 0 Then _
            intrusas = intrusas & x & " "
    Next x

    If Len(faltando) > 0 Then
        ValidarMatrizWestgard = "ERRO DE CONFIGURACAO: faltam na matriz de " & _
            CStr(NLV) & " niveis: " & Trim$(faltando)
    End If
    If Len(intrusas) > 0 Then
        If Len(ValidarMatrizWestgard) > 0 Then ValidarMatrizWestgard = _
            ValidarMatrizWestgard & " | "
        ValidarMatrizWestgard = ValidarMatrizWestgard & _
            "ERRO DE CONFIGURACAO: regras de outro modulo na matriz de " & _
            CStr(NLV) & " niveis: " & Trim$(intrusas)
    End If
End Function


' O limiar que o motor REALMENTE usa para cada regra sequencial, publicado
' para a cobertura poder conferir semantica e nao so nome. Sem isto,
' "plano diz 8x, motor conta 10" volta a passar como TOTAL.
Public Function LimiarSequencialWestgard() As Long
    If NLV >= 3 Then LimiarSequencialWestgard = 6 Else LimiarSequencialWestgard = 8
End Function


' Quantos resultados consecutivos a regra de 1s exige neste produto
' (4_1s com dois niveis, 3_1s com tres).
Public Function LimiarUmSigmaWestgard() As Long
    If NLV >= 3 Then LimiarUmSigmaWestgard = 3 Else LimiarUmSigmaWestgard = 4
End Function


' ============================ MOTOR: MONTAR Calc ============================
' Produz TODA a matriz do Calc (valores, z, limites, regras, marcadores) para o
' analito selecionado. Uma unica escrita por bloco no final.
Public Sub AtualizarCalc()
    Dim ws As Worksheet, analito As String, lote As String
    Dim i As Long, t As Long, nRun As Long, r As Long
    Dim runs() As Long, dts() As Double, valor() As Double, temDado() As Boolean, z() As Double
    Dim r13 As Variant, r22 As Variant, rR4 As Variant, r41 As Variant, r10 As Variant, a12 As Variant
    Dim alvoM() As Double, alvoS() As Double, etp As Double
    Dim seen As Object, ordem As Object

    Set ws = ThisWorkbook.Sheets("Calc")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    lote = LoteAtivoCore()
    GarantirDB

    ' limpa a area de saida
    ws.Range(ws.Cells(KC0, 1), ws.Cells(KC0 + NK - 1, CF0 + NLV * NFD - 1)).ClearContents
    For i = 1 To NK
        ws.Cells(KC0 + i - 1, 1).Value = i
    Next i
    If analito = "" Or IsEmpty(mDB) Then Exit Sub

    ' ---- descobrir corridas (RUN) elegiveis do analito no lote ----
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim runs(1 To NK): ReDim dts(1 To NK)
    For i = 1 To UBound(mDB, 1)
        If StrComp(Trim$(CStr(mDB(i, COL_ANALITO))), analito, 1) = 0 Then
            If Mid$(CStr(mDB(i, COL_LOTE)), 4, 6) = lote Then
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

    ' ordenar por RUN (insercao — nRun e pequeno)
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
            If Mid$(CStr(mDB(i, COL_LOTE)), 4, 6) = lote Then
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

    ' ---- montar a matriz de saida ----
    Dim outBase() As Variant, outLvl() As Variant, rej As Boolean, filtro As Long
    ReDim outBase(1 To nRun, 1 To 5)
    For i = 1 To nRun
        outBase(i, 1) = i
        outBase(i, 2) = runs(i)
        If dts(i) > 0 Then outBase(i, 3) = CDate(dts(i))
        filtro = IIf(PassaFiltro(dts(i)), 1, 0)
        outBase(i, 4) = filtro
        If filtro = 1 Then outBase(i, 5) = runs(i) Else outBase(i, 5) = CVErr(xlErrNA)
    Next i
    ws.Range(ws.Cells(KC0, 1), ws.Cells(KC0 + nRun - 1, 5)).Value = outBase

    For t = 0 To NLV - 1
        ReDim outLvl(1 To nRun, 1 To NFD)
        For i = 1 To nRun
            filtro = IIf(PassaFiltro(dts(i)), 1, 0)
            If temDado(t, i) Then
                outLvl(i, 1) = valor(t, i)
                outLvl(i, 2) = z(t, i)
                rej = (r13(t, i) = 1 Or r22(t, i) = 1 Or rR4(t, i) = 1 Or r41(t, i) = 1 Or r10(t, i) = 1)
                outLvl(i, 6) = IIf(r13(t, i) = 1, 1, 0)
                outLvl(i, 7) = IIf(r22(t, i) = 1, 1, 0)
                outLvl(i, 8) = IIf(rR4(t, i) = 1, 1, 0)
                outLvl(i, 9) = IIf(r41(t, i) = 1, 1, 0)
                outLvl(i, 10) = IIf(r10(t, i) = 1, 1, 0)
                If rej Then
                    outLvl(i, 11) = "REJEITADO"
                ElseIf a12(t, i) = 1 Then
                    outLvl(i, 11) = "ALERTA"
                Else
                    outLvl(i, 11) = "OK"
                End If
                If filtro = 1 Then
                    outLvl(i, 3) = valor(t, i)
                    If rej Then
                        outLvl(i, 5) = valor(t, i): outLvl(i, 4) = CVErr(xlErrNA)
                    Else
                        outLvl(i, 4) = valor(t, i): outLvl(i, 5) = CVErr(xlErrNA)
                    End If
                Else
                    outLvl(i, 3) = CVErr(xlErrNA)
                    outLvl(i, 4) = CVErr(xlErrNA)
                    outLvl(i, 5) = CVErr(xlErrNA)
                End If
            Else
                outLvl(i, 3) = CVErr(xlErrNA)
                outLvl(i, 4) = CVErr(xlErrNA)
                outLvl(i, 5) = CVErr(xlErrNA)
                outLvl(i, 6) = 0: outLvl(i, 7) = 0: outLvl(i, 8) = 0
                outLvl(i, 9) = 0: outLvl(i, 10) = 0
            End If
            ' limites (12..18 = m3,m2,m1,med,p1,p2,p3)
            If filtro = 1 And alvoS(t) > 0 Then
                outLvl(i, 12) = alvoM(t) - 3 * alvoS(t)
                outLvl(i, 13) = alvoM(t) - 2 * alvoS(t)
                outLvl(i, 14) = alvoM(t) - 1 * alvoS(t)
                outLvl(i, 15) = alvoM(t)
                outLvl(i, 16) = alvoM(t) + 1 * alvoS(t)
                outLvl(i, 17) = alvoM(t) + 2 * alvoS(t)
                outLvl(i, 18) = alvoM(t) + 3 * alvoS(t)
            Else
                Dim q As Long
                For q = 12 To 18
                    outLvl(i, q) = CVErr(xlErrNA)
                Next q
            End If
            ' repeticoes / calibracao (aba Registros)
            Dim rp As Variant
            rp = MarcadoresRegistro(analito, t + 1, dts(i))
            outLvl(i, 19) = rp(0): outLvl(i, 20) = rp(1): outLvl(i, 21) = rp(2)
            If filtro = 1 And rp(3) = True And alvoM(t) <> 0 Then
                outLvl(i, 22) = alvoM(t)
            Else
                outLvl(i, 22) = CVErr(xlErrNA)
            End If
        Next i
        ws.Range(ws.Cells(KC0, CF0 + t * NFD), ws.Cells(KC0 + nRun - 1, CF0 + t * NFD + NFD - 1)).Value = outLvl
    Next t

    ' ---- parametros (media/DP alvo, limites de eixo, faixa do eixo X) ----
    Dim P1 As Long, P2 As Long, P3 As Long, lo As Double, hi As Double
    P1 = CF0 + NLV * NFD
    P2 = P1 + NLV * 2
    P3 = P2 + NLV * 2
    For t = 0 To NLV - 1
        ws.Cells(1, P1 + t * 2).Value = IIf(alvoM(t) = 0 And alvoS(t) = 0, "", alvoM(t))
        ws.Cells(1, P1 + t * 2 + 1).Value = IIf(alvoS(t) = 0, "", alvoS(t))
        If alvoS(t) > 0 Then
            lo = alvoM(t) - 3.3 * alvoS(t): hi = alvoM(t) + 3.3 * alvoS(t)
            For i = 1 To nRun
                If temDado(t, i) And PassaFiltro(dts(i)) Then
                    If valor(t, i) < lo Then lo = valor(t, i)
                    If valor(t, i) > hi Then hi = valor(t, i)
                End If
            Next i
            ws.Cells(1, P2 + t * 2).Value = lo
            ws.Cells(1, P2 + t * 2 + 1).Value = hi
        Else
            ws.Cells(1, P2 + t * 2).Value = ""
            ws.Cells(1, P2 + t * 2 + 1).Value = ""
        End If
    Next t
    ' eixo X em numero de RUN — o rotulo cai exatamente sob o ponto
    Dim rmin As Long, rmax As Long
    rmin = 0: rmax = 0
    For i = 1 To nRun
        If PassaFiltro(dts(i)) Then
            If rmin = 0 Or runs(i) < rmin Then rmin = runs(i)
            If runs(i) > rmax Then rmax = runs(i)
        End If
    Next i
    If rmin = 0 Then
        ws.Cells(1, P3).Value = "": ws.Cells(1, P3 + 1).Value = ""
    Else
        ws.Cells(1, P3).Value = rmin - 0.5
        ws.Cells(1, P3 + 1).Value = rmax + 0.5
    End If
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
                If loteCore = "" Or Mid$(CStr(mDB(i, COL_LOTE)), 4, 6) = loteCore Then
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
Public Sub AtualizarPainelEng()
    Dim ws As Worksheet, ca As Worksheet, analito As String, lote As String
    Dim t As Long, i As Long, n As Long, v() As Double, media As Double, dp As Double
    Dim cv As Double, bias As Double, et As Double, sg As Double, etp As Double
    Dim aM As Double, aS As Double, rr As Long, rejTot As Long
    Dim cnt(1 To 5) As Long, k2 As Long, bloco As Variant, blocoF As Variant
    Set ws = ThisWorkbook.Sheets("Painel")
    Set ca = ThisWorkbook.Sheets("Calc")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    lote = LoteAtivoCore()
    For t = 0 To NLV - 1
        rr = 7 + t
        n = 0: ReDim v(1 To NK)
        For i = 1 To 5
            cnt(i) = 0
        Next i
        rejTot = 0
        If IsEmpty(blocoF) Then blocoF = ca.Range(ca.Cells(KC0, 4), ca.Cells(KC0 + NK - 1, 4)).Value
        bloco = ca.Range(ca.Cells(KC0, CF0 + t * NFD), ca.Cells(KC0 + NK - 1, CF0 + t * NFD + 10)).Value
        For i = 1 To NK
            If Val(blocoF(i, 1)) = 1 Then
                If IsNumeric(bloco(i, 1)) And Trim$(CStr(bloco(i, 1))) <> "" Then
                    n = n + 1: v(n) = CDbl(bloco(i, 1))
                End If
                For k2 = 1 To 5
                    If Val(bloco(i, 5 + k2)) = 1 Then cnt(k2) = cnt(k2) + 1
                Next k2
                If CStr(bloco(i, 11)) = "REJEITADO" Then rejTot = rejTot + 1
            End If
        Next i
        media = CalcularMedia(v, n)
        dp = CalcularDP(v, n, media)
        cv = CalcularCV(dp, media)
        AlvoAnalito analito, t + 1, aM, aS, etp
        bias = CalcularBias(media, aM)
        et = CalcularErroTotal(cv, bias)
        sg = CalcularSigma(etp, bias, cv)

        ws.Cells(rr, 2).Value = n
        ws.Cells(rr, 3).Value = IIf(n = 0, "", media)
        ws.Cells(rr, 4).Value = IIf(n < 2, "", dp)
        ws.Cells(rr, 5).Value = IIf(n < 2 Or media = 0, "", cv)
        ws.Cells(rr, 6).Value = IIf(etp = 0, "", etp)
        ws.Cells(rr, 7).Value = IIf(n = 0 Or aM = 0, "", bias)
        ws.Cells(rr, 8).Value = IIf(n < 2 Or media = 0, "", et)
        ws.Cells(rr, 9).Value = IIf(n < 2 Or cv = 0 Or etp = 0, "", sg)
        ws.Cells(rr, 10).Value = IIf(n = 0, "", IIf(rejTot > 0, "REJEITADO", "OK"))
        For i = 1 To 5
            ws.Cells(rr, 12 + i).Value = cnt(i)
        Next i
        ws.Cells(rr, 18).Value = cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5)
        ' resumo — o historico completo fica na aba Eventos_Westgard
        Dim aa As Variant
        aa = AgregadoWestgard(analito, t + 1)
        If CLng(aa(0)) = 0 Then
            ws.Cells(rr, 19).Value = "—"
            ws.Cells(rr, 20).Value = ""
            ws.Cells(rr, 21).Value = ""
        Else
            ws.Cells(rr, 19).Value = aa(3) & " · RUN " & aa(4)
            ws.Cells(rr, 20).Value = RegraClassificacao(Split(CStr(aa(3)), "+")(0))
            ws.Cells(rr, 21).Value = aa(0) & "x  |  maior Z " & Format(aa(5), "+0.00;-0.00")
        End If
    Next t
    ws.Cells(6, 19).Value = "Últ. violação": ws.Cells(6, 19).Font.Bold = True
    ws.Cells(6, 20).Value = "Classificação": ws.Cells(6, 20).Font.Bold = True
    ws.Cells(6, 21).Value = "Histórico": ws.Cells(6, 21).Font.Bold = True
    ws.Columns(19).ColumnWidth = 18: ws.Columns(20).ColumnWidth = 18: ws.Columns(21).ColumnWidth = 22
End Sub

' ============================ ABA ESTATISTICA ============================
Public Sub AtualizarEstatisticaAba()
    Dim ws As Worksheet, wa As Worksheet, i As Long, t As Long, er As Long
    Dim analito As String, anoDe As Long, anoAte As Long, loteF As String
    Dim st As Variant, n As Long, media As Double, dp As Double, cv As Double
    Dim aM As Double, aS As Double, etp As Double, bias As Double, et As Double, sg As Double
    Dim outp() As Variant, linhas As Long
    Set ws = ThisWorkbook.Sheets("Estatística")
    Set wa = ThisWorkbook.Sheets("Analitos")
    anoDe = CLng(Val(ws.Range("B3").Value))
    anoAte = CLng(Val(ws.Range("D3").Value))
    If anoAte < anoDe Then anoAte = anoDe
    loteF = Trim$(CStr(ws.Range("B4").Value))
    If loteF = "" Then loteF = ""
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
                AlvoAnalito analito, t, aM, aS, etp
                bias = CalcularBias(media, aM)
                et = CalcularErroTotal(cv, bias)
                sg = CalcularSigma(etp, bias, cv)
                outp(er, 1) = n
                outp(er, 2) = IIf(n = 0, "", media)
                outp(er, 3) = IIf(n < 2, "", dp)
                outp(er, 4) = IIf(n < 2 Or media = 0, "", cv)
                outp(er, 5) = IIf(etp = 0, "", etp)
                outp(er, 6) = IIf(IsNumeric(wa.Cells(i, 19).Value), wa.Cells(i, 19).Value, "")
                outp(er, 7) = IIf(IsNumeric(wa.Cells(i, 20).Value), wa.Cells(i, 20).Value, "")
                outp(er, 8) = IIf(n = 0 Or aM = 0, "", bias)
                outp(er, 9) = IIf(n < 2 Or media = 0, "", et)
                outp(er, 10) = IIf(n < 2 Or cv = 0 Or etp = 0, "", sg)
                outp(er, 11) = IIf(n < 2 Or cv = 0 Or etp = 0, "", mQualidade.ClassificarSigma(sg))
            End If
        Next t
    Next i
    ws.Range(ws.Cells(E0, 3), ws.Cells(E0 + linhas - 1, 13)).Value = outp
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
            If Mid$(CStr(mDB(i, COL_LOTE)), 4, 6) = lote Then
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

                Dim aM As Double, aS As Double, etp As Double
                AlvoAnalito analitoN, nivelN, aM, aS, etp

                Dim zz() As Double, td() As Boolean
                Dim q13 As Variant, q22 As Variant, qR4 As Variant, q41 As Variant, q10 As Variant, q12 As Variant
                ReDim zz(0 To 0, 1 To nS): ReDim td(0 To 0, 1 To nS)
                For j = 1 To nS
                    If aS > 0 Then zz(0, j) = CalcularZ(ys(j), aM, aS)
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


