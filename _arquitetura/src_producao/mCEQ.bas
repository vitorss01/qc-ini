Attribute VB_Name = "mCEQ"
Option Explicit
' ===== CONTROLE EXTERNO / ENSAIO DE PROFICIENCIA (ADR-030 / ADR-032) =====
'
' O DEFEITO QUE ESTE MODULO CORRIGIU (ADR-030)
'
' O bias que alimentava Erro Total e Sigma vinha de
' mEstatistica.CalcularBias(mediaObs, alvo), onde:
'
'     mediaObs = media do CONTROLE INTERNO no periodo
'     alvo     = media atribuida AO LOTE do controle interno
'
' Isso mede a deriva do CQI contra o alvo do proprio lote -- util, mas NAO e
' erro sistematico. Erro sistematico se mede contra um valor externo e
' independente: o consenso do grupo no ensaio de proficiencia. Um laboratorio
' pode estar centrado no alvo do fabricante e, ainda assim, 8% acima do grupo.
'
' A formula sempre esteve certa dos dois lados: (X - ref)/ref * 100. O que
' estava errado era o "ref".
'
' O QUE O ADR-032 ACRESCENTA
'
' O laboratorio participa de MAIS DE UM programa -- Controllab com 4 rodadas
' anuais, CAP com 3 -- e a escolha de qual usar e decisao tecnica dele, nao do
' sistema. Todas as funcoes aqui aceitam agora provedor e rodada:
'
'     provedor  ""/"TODOS"  -> qualquer programa
'               "Controllab" | "CAP"
'     rodada    ""/"TODAS"   -> consolida as rodadas do ano
'               "A" | "B" | "C" | "D"
'
' A FONTE
'
'     EQC_Dados
'     A Analito   B Ano   C Rodada (A..D)   D Data   E Provedor   F Amostra
'     G Resultado lab (X_lab)   H Media grupo (X_ref)   I SD grupo
'     J SDI = (G-H)/I           K Lim.Inf   L Lim.Sup   M Status limites
'     N Bias % = (G-H)/H*100    O |Bias| %  P Status SDI
'
' G, H, I, K e L sao DIGITADOS pelo usuario. O resto e calculado na propria
' celula, onde fica visivel e auditavel -- este modulo CONSOME essas colunas,
' nao as recalcula. Reimplementar aqui criaria a segunda versao do indicador.
'
' CONSOLIDACAO DE MULTIPLAS RODADAS
'
' O valor que alimenta ET e Sigma e a MEDIA DAS MAGNITUDES:
'
'     |Bias|consolidado = soma(|Bias_i|) / n
'
' e nunca a media dos assinados. Rodadas de +5% e -5% descrevem um metodo que
' oscila 5% em torno do grupo; a media assinada daria 0% e afirmaria exatidao
' perfeita. Cancelamento de sinal apaga o erro que se quer medir.
'
' O bias ASSINADO continua disponivel em modo "SIGNED", para ler a direcao do
' desvio. Ele informa; nao entra nas metricas de magnitude.
'
' AUSENCIA DE DADO NAO E ZERO, E NAO PODE SER Empty
'
' Sem rodada utilizavel devolve-se o TEXTO "SEM EP". Devolver Empty parecia
' natural e estava errado: o Excel renderiza o Empty de uma UDF como ZERO na
' celula. Na primeira versao deste modulo as 80 linhas da Estatistica exibiram
' bias 0,00 -- inclusive analitos sem nenhuma rodada -- e esse zero entrou em
' ET e Sigma produzindo numeros de aparencia perfeita.
'
' VIGENCIA: A RODADA MAIS RECENTE QUE NAO ULTRAPASSA O ANO DE REFERENCIA
'
' O EP e anual e sai depois; o CQI e do mes corrente. Exigir coincidencia exata
' faria o bias sumir sempre que o CQI passasse na frente do ultimo ciclo
' publicado. Mesma regra de vigencia do ADR-022 para especificacao.

' ADR-034: A FONTE PASSOU A SER A EQA_Base
'
' Ate aqui este modulo lia a EQC_Dados, aba unica onde CAP e Controllab
' dividiam as mesmas colunas. Agora cada provedor tem a sua aba de digitacao,
' com a terminologia dele, e a EQA_Base normaliza as duas.
'
' Este e o UNICO lugar da pasta que le a aba de EP. As 403 celulas da
' Estatistica e do Painel, e a coluna de bias do BI, chamam as funcoes daqui --
' nenhuma delas aponta para a planilha. Por isso trocar a fonte foi trocar
' estas constantes, e nao 403 formulas.
'
' O analito casado e o CANONICO (coluna E), nao o nome do provedor (coluna D):
' o CAP reporta "Urea Nitrogen" e a Analitos chama "Ureia". A coluna D continua
' na base para rastrear ate o PDF.

Private Const EQ_ABA As String = "EQA_Base"
Private Const EQ_R0 As Long = 2
Private Const EQ_RN As Long = 5001
Private Const EQ_C_PROVEDOR As Long = 1
Private Const EQ_C_ANO As Long = 2
Private Const EQ_C_RODADA As Long = 3
Private Const EQ_C_ANALITO As Long = 5
Private Const EQ_C_XLAB As Long = 7
Private Const EQ_C_XREF As Long = 8
Private Const EQ_C_SDGRUPO As Long = 9
Private Const EQ_C_SDI As Long = 10
Private Const EQ_C_LIMINF As Long = 11
Private Const EQ_C_LIMSUP As Long = 12
Private Const EQ_C_BIAS As Long = 16
Private Const EQ_C_BIASABS As Long = 17
Private Const EQ_C_USO As Long = 20
Private Const EQ_NCOL As Long = 21

Public Const SEM_EP As String = "SEM EP"
Public Const LIM_SDI As Double = 2#


Private Function Igual(ByVal a As Variant, ByVal b As String) As Boolean
    Igual = (UCase$(Trim$(CStr(a))) = UCase$(Trim$(b)))
End Function

' Um filtro vazio, "TODOS" ou "TODAS" nao restringe nada.
Private Function Livre(ByVal f As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(f)))
    Livre = (s = "" Or s = "TODOS" Or s = "TODAS")
End Function

Private Function LerBanco() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(EQ_ABA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ' Da coluna 1 ate a ultima: assim d(i, EQ_C_XLAB) e literalmente a coluna
    ' EQ_C_XLAB. Ler a partir do analito (coluna 5) deslocaria todo indice em 4.
    LerBanco = ws.Range(ws.Cells(EQ_R0, 1), _
                        ws.Cells(EQ_RN, EQ_NCOL)).Value
End Function

' A linha pertence ao analito, ao provedor e a rodada pedidos?
Private Function Casa(ByRef d As Variant, ByVal i As Long, ByVal analito As String, _
                      ByVal provedor As Variant, ByVal rodada As Variant) As Boolean
    If Not Igual(d(i, EQ_C_ANALITO), analito) Then Exit Function
    ' Uso_Analitico = NAO marca dado preservado por historico que nao pode
    ' entrar em bias, Sigma nem ET -- hoje, os 90 registros de simulacao que
    ' vinham da EQC_Dados. Ver o cabecalho do mEQA.
    If Igual(d(i, EQ_C_USO), "NAO") Then Exit Function
    If Not Livre(provedor) Then
        If Not Igual(d(i, EQ_C_PROVEDOR), CStr(provedor)) Then Exit Function
    End If
    If Not Livre(rodada) Then
        If Not Igual(d(i, EQ_C_RODADA), CStr(rodada)) Then Exit Function
    End If
    Casa = True
End Function

' Maior ano que nao ultrapassa anoRef, ja respeitando provedor e rodada.
Private Function AnoVigente(ByRef d As Variant, ByVal analito As String, _
                            ByVal anoRef As Variant, ByVal provedor As Variant, _
                            ByVal rodada As Variant, ByVal colExigida As Long) As Long
    Dim i As Long, aRef As Long, ano As Long
    AnoVigente = -32768
    aRef = 32767
    If IsNumeric(anoRef) Then
        If Len(Trim$(CStr(anoRef))) > 0 Then aRef = CLng(Val(CStr(anoRef)))
    End If
    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada) Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_ANO)) Then GoTo prox
        If Not IsNumeric(d(i, colExigida)) Then GoTo prox
        ano = CLng(Val(CStr(d(i, EQ_C_ANO))))
        If ano <= aRef And ano > AnoVigente Then AnoVigente = ano
prox:
    Next i
End Function

' ---------------------------------------------------------------------------
' Bias do ensaio de proficiencia, consolidado.
'
'   modo  "ABS"    media das magnitudes -> alimenta ET e Sigma
'         "SIGNED" media dos assinados  -> leitura da direcao
'         "N"      quantas rodadas entraram
'         "ANO"    qual ano acabou vigente
'
' Devolve "SEM EP" quando nao ha rodada utilizavel.
Public Function BiasEQ(ByVal analito As String, ByVal anoRef As Variant, _
                       ByVal modo As String, _
                       Optional ByVal provedor As Variant = "", _
                       Optional ByVal rodada As Variant = "") As Variant
    ' -----------------------------------------------------------------
    ' CONSOLIDACAO EM DUAS ETAPAS (ADR-035)
    '
    '   etapa 1: para cada RODADA, a media dos |bias| das amostras dela
    '   etapa 2: a media dessas medias de rodada
    '
    ' A media simples de todas as amostras juntas dava peso maior a rodada
    ' que por acaso teve mais amostras. Com C-A e C-C de 5 amostras e uma
    ' rodada nova de 12, a rodada nova passaria a mandar no numero sem que
    ' ninguem tivesse decidido isso.
    '
    ' E as duas etapas trabalham sobre |bias|, nunca sobre o bias com sinal:
    ' +8% e -8% na mesma rodada descrevem um metodo que ninguem aprovaria, e
    ' a media assinada devolveria zero.
    '
    '   modo  "ABS"       magnitude -> alimenta ET e Sigma
    '         "SIGNED"    direcao do desvio, mesma consolidacao
    '         "N"         quantas AMOSTRAS entraram
    '         "NRODADAS"  quantas RODADAS entraram
    '         "ANO"       qual ano acabou vigente
    '         "DETALHE"   memoria de calculo, rodada a rodada
    '
    ' Devolve "SEM EP" quando nao ha rodada utilizavel -- nunca 0, que a
    ' celula exibiria como exatidao perfeita.
    ' -----------------------------------------------------------------
    Const MAX_ROD As Long = 64

    Dim d As Variant, i As Long, j As Long, k As Long
    Dim anoVig As Long, col As Long, md As String, r As String
    Dim rot(1 To MAX_ROD) As String
    Dim somaAbs(1 To MAX_ROD) As Double
    Dim somaSig(1 To MAX_ROD) As Double
    Dim cnt(1 To MAX_ROD) As Long
    Dim nRod As Long, nAmostras As Long
    Dim acc As Double, mr As Double, det As String

    BiasEQ = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco()
    If IsEmpty(d) Then Exit Function

    md = UCase$(Trim$(modo))
    Select Case md
        Case "ABS":                          col = EQ_C_BIASABS
        Case "SIGNED":                       col = EQ_C_BIAS
        Case "N", "NRODADAS", "ANO", "DETALHE": col = EQ_C_BIASABS
        Case Else:                           Exit Function
    End Select

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_BIASABS)
    If anoVig = -32768 Then Exit Function
    If md = "ANO" Then BiasEQ = anoVig: Exit Function

    ' ---- etapa 1: acumula por rodada ---------------------------------
    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada) Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_ANO)) Then GoTo prox
        If CLng(Val(CStr(d(i, EQ_C_ANO)))) <> anoVig Then GoTo prox
        If Not IsNumeric(d(i, col)) Then GoTo prox

        r = UCase$(Trim$(CStr(d(i, EQ_C_RODADA))))
        k = 0
        For j = 1 To nRod
            If rot(j) = r Then k = j: Exit For
        Next j
        If k = 0 Then
            If nRod >= MAX_ROD Then GoTo prox
            nRod = nRod + 1
            k = nRod
            rot(k) = r
        End If
        If IsNumeric(d(i, EQ_C_BIASABS)) Then _
            somaAbs(k) = somaAbs(k) + CDbl(d(i, EQ_C_BIASABS))
        If IsNumeric(d(i, EQ_C_BIAS)) Then _
            somaSig(k) = somaSig(k) + CDbl(d(i, EQ_C_BIAS))
        cnt(k) = cnt(k) + 1
        nAmostras = nAmostras + 1
prox:
    Next i

    If nRod = 0 Then Exit Function
    If md = "N" Then BiasEQ = nAmostras: Exit Function
    If md = "NRODADAS" Then BiasEQ = nRod: Exit Function

    ' ---- etapa 2: media das medias de rodada -------------------------
    acc = 0
    det = ""
    For k = 1 To nRod
        If cnt(k) > 0 Then
            If md = "SIGNED" Then
                mr = somaSig(k) / cnt(k)
            Else
                mr = somaAbs(k) / cnt(k)
            End If
            acc = acc + mr
            det = det & rot(k) & " = " & Format$(mr, "0.0000") & _
                  " (n=" & cnt(k) & ")"
            If k < nRod Then det = det & "  |  "
        End If
    Next k

    If md = "DETALHE" Then
        BiasEQ = det & "   ==>   media das " & nRod & " rodada(s) = " & _
                 Format$(acc / nRod, "0.000000")
        Exit Function
    End If

    BiasEQ = acc / nRod
End Function

' ---------------------------------------------------------------------------
' SDI consolidado.
'
'   modo  "MEDIA"  media dos SDI assinados
'         "MAX"    maior |SDI| -- e este que decide o status
'         "N"      quantas amostras entraram
'
' O SDI mede o desvio em unidades de DP DO GRUPO: (X_lab - media grupo)/SD grupo.
' Ele nao substitui o bias: bias e magnitude relativa, SDI e posicao dentro da
' dispersao do grupo. Os dois respondem perguntas diferentes.
Public Function SDIeq(ByVal analito As String, ByVal anoRef As Variant, _
                      ByVal modo As String, _
                      Optional ByVal provedor As Variant = "", _
                      Optional ByVal rodada As Variant = "") As Variant
    Dim d As Variant, i As Long, anoVig As Long
    Dim soma As Double, n As Long, maxAbs As Double, v As Variant

    SDIeq = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco()
    If IsEmpty(d) Then Exit Function

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_SDI)
    If anoVig = -32768 Then Exit Function

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada) Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_ANO)) Then GoTo prox
        If CLng(Val(CStr(d(i, EQ_C_ANO)))) <> anoVig Then GoTo prox
        v = d(i, EQ_C_SDI)
        If Not IsNumeric(v) Then GoTo prox
        soma = soma + CDbl(v)
        If Abs(CDbl(v)) > maxAbs Then maxAbs = Abs(CDbl(v))
        n = n + 1
prox:
    Next i

    If n = 0 Then Exit Function
    Select Case UCase$(Trim$(modo))
        Case "MEDIA": SDIeq = soma / n
        Case "MAX":   SDIeq = maxAbs
        Case "N":     SDIeq = n
        Case Else:    SDIeq = SEM_EP
    End Select
End Function

' Status do SDI: nenhuma amostra pode passar de |2|.
'
' O criterio e por PIOR AMOSTRA, nao pela media. Uma rodada com SDI +3 e outra
' com -3 dao media zero e descrevem um desempenho que ninguem aprovaria; e a
' pior que reprova o conjunto.
Public Function StatusSDIeq(ByVal analito As String, ByVal anoRef As Variant, _
                            Optional ByVal provedor As Variant = "", _
                            Optional ByVal rodada As Variant = "") As Variant
    Dim m As Variant, n As Variant
    m = SDIeq(analito, anoRef, "MAX", provedor, rodada)
    If Not IsNumeric(m) Then StatusSDIeq = SEM_EP: Exit Function
    n = SDIeq(analito, anoRef, "N", provedor, rodada)
    If CDbl(m) <= LIM_SDI Then
        StatusSDIeq = "OK (|SDI| max " & Format$(m, "0.00") & " em " & n & " amostra(s))"
    Else
        StatusSDIeq = "FORA (|SDI| max " & Format$(m, "0.00") & " > " & _
                      Format$(LIM_SDI, "0") & ")"
    End If
End Function

' Status dos limites do grupo: o resultado do laboratorio caiu dentro da faixa
' informada pelo provedor em TODAS as amostras?
'
' Limite ausente nao e aprovacao: a amostra entra como NAO AVALIADA e aparece na
' contagem, para nao passar por conforme quem ninguem conferiu.
Public Function StatusLimitesEQ(ByVal analito As String, ByVal anoRef As Variant, _
                                Optional ByVal provedor As Variant = "", _
                                Optional ByVal rodada As Variant = "") As Variant
    Dim d As Variant, i As Long, anoVig As Long
    Dim dentro As Long, fora As Long, semLim As Long
    Dim x As Variant, li As Variant, ls As Variant

    StatusLimitesEQ = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco()
    If IsEmpty(d) Then Exit Function

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_XLAB)
    If anoVig = -32768 Then Exit Function

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada) Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_ANO)) Then GoTo prox
        If CLng(Val(CStr(d(i, EQ_C_ANO)))) <> anoVig Then GoTo prox
        x = d(i, EQ_C_XLAB)
        If Not IsNumeric(x) Then GoTo prox
        li = d(i, EQ_C_LIMINF)
        ls = d(i, EQ_C_LIMSUP)
        If Not IsNumeric(li) Or Not IsNumeric(ls) Then
            semLim = semLim + 1
        ElseIf CDbl(x) < CDbl(li) Or CDbl(x) > CDbl(ls) Then
            fora = fora + 1
        Else
            dentro = dentro + 1
        End If
prox:
    Next i

    If dentro + fora + semLim = 0 Then Exit Function
    If fora > 0 Then
        StatusLimitesEQ = "NAO OK (" & fora & " fora dos limites)"
    ElseIf semLim > 0 Then
        StatusLimitesEQ = "OK (" & dentro & " dentro; " & semLim & " sem limite)"
    Else
        StatusLimitesEQ = "OK (" & dentro & " dentro dos limites)"
    End If
End Function

' ---------------------------------------------------------------------------
' Rastreabilidade: memoria de calculo, para auditoria.
Public Function BiasEQMemoria(ByVal analito As String, ByVal anoRef As Variant, _
                              Optional ByVal provedor As Variant = "", _
                              Optional ByVal rodada As Variant = "") As String
    Dim d As Variant, i As Long, anoVig As Long, s As String, n As Long
    Dim somaAbs As Double, somaSig As Double

    d = LerBanco()
    If IsEmpty(d) Then BiasEQMemoria = "EQA_Base ausente": Exit Function
    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_BIASABS)
    If anoVig = -32768 Then BiasEQMemoria = "sem rodada utilizavel": Exit Function

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada) Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_ANO)) Then GoTo prox
        If CLng(Val(CStr(d(i, EQ_C_ANO)))) <> anoVig Then GoTo prox
        If Not IsNumeric(d(i, EQ_C_BIAS)) Then GoTo prox
        n = n + 1
        somaSig = somaSig + CDbl(d(i, EQ_C_BIAS))
        somaAbs = somaAbs + CDbl(d(i, EQ_C_BIASABS))
        s = s & CStr(d(i, EQ_C_PROVEDOR)) & "/" & CStr(d(i, EQ_C_RODADA)) & _
                "|ano=" & CStr(anoVig) & _
                "|Xlab=" & Format$(d(i, EQ_C_XLAB), "0.####") & _
                "|Xref=" & Format$(d(i, EQ_C_XREF), "0.####") & _
                "|SDI=" & Format$(d(i, EQ_C_SDI), "0.##") & _
                "|bias=" & Format$(d(i, EQ_C_BIAS), "0.####") & ";"
prox:
    Next i
    If n = 0 Then BiasEQMemoria = "sem rodada utilizavel": Exit Function
    BiasEQMemoria = s & " CONSOLIDADO n=" & n & _
                    " media|bias|=" & Format$(somaAbs / n, "0.######") & _
                    " mediaAssinada=" & Format$(somaSig / n, "0.######")
End Function

' Resumo do filtro em uso, para a propria aba dizer o que esta usando.
Public Function ResumoFiltroEQ(ByVal provedor As Variant, ByVal ano As Variant, _
                               ByVal rodada As Variant) As String
    Dim p As String, r As String, a As String
    p = IIf(Livre(provedor), "todos os provedores", Trim$(CStr(provedor)))
    r = IIf(Livre(rodada), "media de todas as rodadas", "rodada " & Trim$(CStr(rodada)))
    a = IIf(Len(Trim$(CStr(ano))) = 0, "ano vigente do periodo", "ano " & Trim$(CStr(ano)))
    ResumoFiltroEQ = "Bias do EP: " & p & " | " & a & " | " & r
End Function
