Attribute VB_Name = "mCEQ"
Option Explicit
' ===== BIAS DO ENSAIO DE PROFICIENCIA (ADR-030) =====
'
' O DEFEITO QUE ESTE MODULO CORRIGE
'
' O bias que alimentava Erro Total e Sigma vinha de
' mEstatistica.CalcularBias(mediaObs, alvo), onde:
'
'     mediaObs = media do CONTROLE INTERNO no periodo
'     alvo     = media atribuida AO LOTE do controle interno
'
' Isso mede a deriva do CQI em relacao ao proprio alvo do lote -- e uma medida
' util, mas NAO e o erro sistematico do metodo. Erro sistematico se mede contra
' um valor externo e independente: o consenso do grupo no ensaio de
' proficiencia. Um laboratorio pode estar perfeitamente centrado no alvo do
' fabricante e, ainda assim, 8% acima do grupo -- e era esse 8% que sumia.
'
' A formula em si sempre esteve certa nos dois lados: (X - ref)/ref * 100.
' O que estava errado era o "ref".
'
' A FONTE, E POR QUE ELA JA EXISTIA
'
' A aba EQC_Dados ja guarda o ensaio de proficiencia linha a linha:
'
'     A Analito   B Ano   C Rodada   D Data   E Provedor   F Amostra
'     G Resultado lab (X_lab)        H Media grupo (X_ref)  I SD grupo
'     N Bias %  = (G-H)/H*100        O |Bias| % = ABS(N)
'
' As colunas N e O ja calculavam o bias corretamente, com o sinal preservado em
' N e a magnitude em O. Ninguem as consumia para ET e Sigma.
'
' Este modulo NAO recalcula o bias de cada linha: ele CONSOME N e O. A conta
' acontece uma vez so, na celula, onde e visivel e auditavel -- que e o que a
' regra de fonte unica pede. Reimplementar aqui criaria a segunda versao do
' mesmo indicador.
'
' CONSOLIDACAO DE MULTIPLAS RODADAS
'
' Com varias rodadas no periodo, o valor que alimenta ET e Sigma e a
'
'     MEDIA DAS MAGNITUDES:  soma(|Bias_i|) / n
'
' e nunca a media dos valores assinados. Rodadas de +5% e -5% descrevem um
' metodo que oscila 5% em torno do grupo; a media assinada daria 0% e afirmaria
' exatidao perfeita. Cancelamento de sinal aqui apaga exatamente o erro que se
' quer medir.
'
' O bias ASSINADO continua disponivel em modo "SIGNED", para leitura e
' interpretacao da direcao do desvio. Ele informa; nao entra nas metricas de
' magnitude.
'
' AUSENCIA DE DADO NAO E ZERO, E NAO PODE SER Empty
'
' Sem rodada utilizavel a funcao devolve o TEXTO "SEM EP".
'
' Devolver Empty parecia natural e estava errado: o Excel renderiza o Empty de
' uma UDF como ZERO na celula. Na primeira versao deste modulo as 80 linhas da
' Estatistica exibiram bias 0,00 -- inclusive analitos sem nenhuma rodada de EP
' -- e esse zero entrou em ET e em Sigma produzindo numeros de aparencia
' perfeita. E o mesmo defeito que AlvoDoLote tem em mEstatPeriodo.
'
' Texto nao e numero: ISNUMBER devolve falso, e ET e Sigma ficam vazios em vez
' de inventar exatidao.
'
' VIGENCIA: A RODADA MAIS RECENTE QUE NAO ULTRAPASSA O ANO EM ANALISE
'
' O EP e anual e sai depois; o CQI e do mes corrente. Exigir que os dois anos
' coincidam faria o bias sumir sempre que o ano do CQI passasse na frente do
' ultimo ciclo publicado -- foi o que aconteceu aqui: EP de 2025, analise de
' 2026, zero rodadas elegiveis.
'
' A regra adotada e a MESMA que o ADR-022 ja usa para a especificacao vigente:
' vale o maior ano que nao ultrapassa o ano de referencia. Uma unica semantica
' de vigencia no sistema inteiro.
'
' X_ref = 0 tambem nao e calculavel: a divisao nao existe. A propria EQC_Dados!N
' ja devolve vazio nesse caso, e a linha simplesmente nao entra na media.

Private Const EQ_ABA As String = "EQC_Dados"
Private Const EQ_R0 As Long = 4
Private Const EQ_RN As Long = 1003
Private Const EQ_C_ANALITO As Long = 1
Private Const EQ_C_ANO As Long = 2
Private Const EQ_C_RODADA As Long = 3
Private Const EQ_C_XLAB As Long = 7
Private Const EQ_C_XREF As Long = 8
Private Const EQ_C_BIAS As Long = 14      ' N = bias assinado
Private Const EQ_C_BIASABS As Long = 15   ' O = |bias|

' Bias do ensaio de proficiencia, consolidado no periodo.
'
'   analito  nome exatamente como na aba Analitos
'   anoRef   ano em analise; vale o maior ano de EP que nao o ultrapasse
'   modo     "ABS"    media das magnitudes -> alimenta ET e Sigma
'            "SIGNED" media dos assinados  -> leitura da direcao
'            "N"      quantas rodadas entraram na consolidacao
'            "ANO"    qual ano de EP acabou vigente
'
' Devolve o TEXTO "SEM EP" quando nao ha rodada utilizavel -- nunca Empty, que
' o Excel renderiza como zero.
Public Function BiasEQ(ByVal analito As String, ByVal anoRef As Variant, _
                       ByVal modo As String) As Variant
    Dim ws As Worksheet, dados As Variant, i As Long
    Dim aRef As Long, ano As Long, anoVig As Long
    Dim soma As Double, n As Long, col As Long, v As Variant

    BiasEQ = "SEM EP"
    If Len(Trim$(analito)) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(EQ_ABA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    aRef = 32767
    If IsNumeric(anoRef) Then
        If Len(Trim$(CStr(anoRef))) > 0 Then aRef = CLng(Val(CStr(anoRef)))
    End If

    Select Case UCase$(Trim$(modo))
        Case "ABS":    col = EQ_C_BIASABS
        Case "SIGNED": col = EQ_C_BIAS
        Case "N", "ANO": col = EQ_C_BIASABS
        Case Else:     Exit Function
    End Select

    dados = ws.Range(ws.Cells(EQ_R0, EQ_C_ANALITO), ws.Cells(EQ_RN, EQ_C_BIASABS)).Value

    ' 1a passada: qual e o ano vigente para este analito
    anoVig = -32768
    For i = 1 To UBound(dados, 1)
        If UCase$(Trim$(CStr(dados(i, EQ_C_ANALITO)))) <> UCase$(Trim$(analito)) Then GoTo p1
        If Not IsNumeric(dados(i, EQ_C_ANO)) Then GoTo p1
        If Not IsNumeric(dados(i, EQ_C_BIASABS)) Then GoTo p1
        ano = CLng(Val(CStr(dados(i, EQ_C_ANO))))
        If ano <= aRef And ano > anoVig Then anoVig = ano
p1:
    Next i
    If anoVig = -32768 Then Exit Function

    If UCase$(Trim$(modo)) = "ANO" Then BiasEQ = anoVig: Exit Function

    ' 2a passada: consolida as rodadas DAQUELE ano
    For i = 1 To UBound(dados, 1)
        If UCase$(Trim$(CStr(dados(i, EQ_C_ANALITO)))) <> UCase$(Trim$(analito)) Then GoTo p2
        If Not IsNumeric(dados(i, EQ_C_ANO)) Then GoTo p2
        If CLng(Val(CStr(dados(i, EQ_C_ANO)))) <> anoVig Then GoTo p2
        v = dados(i, col)
        If Not IsNumeric(v) Then GoTo p2
        If Len(Trim$(CStr(v))) = 0 Then GoTo p2
        soma = soma + CDbl(v)
        n = n + 1
p2:
    Next i

    If n = 0 Then Exit Function
    If UCase$(Trim$(modo)) = "N" Then BiasEQ = n Else BiasEQ = soma / n
End Function

' Rastreabilidade: devolve a memoria de calculo do bias de um analito.
'
' Existe para o item de auditoria -- permite demonstrar, sem abrir a planilha,
' de quais X_lab e X_ref cada rodada saiu e como elas foram consolidadas.
' Formato: "rodada=..|ano=..|Xlab=..|Xref=..|bias=..;" por linha, e ao fim o
' consolidado.
Public Function BiasEQMemoria(ByVal analito As String, ByVal anoIni As Variant, _
                              ByVal anoFim As Variant) As String
    Dim ws As Worksheet, dados As Variant, i As Long
    Dim a1 As Long, a2 As Long, ano As Long, s As String, n As Long
    Dim somaAbs As Double, somaSig As Double

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(EQ_ABA)
    On Error GoTo 0
    If ws Is Nothing Then BiasEQMemoria = "EQC_Dados ausente": Exit Function

    a1 = -32768: a2 = 32767
    If IsNumeric(anoIni) Then If Len(Trim$(CStr(anoIni))) > 0 Then a1 = CLng(Val(CStr(anoIni)))
    If IsNumeric(anoFim) Then If Len(Trim$(CStr(anoFim))) > 0 Then a2 = CLng(Val(CStr(anoFim)))

    dados = ws.Range(ws.Cells(EQ_R0, EQ_C_ANALITO), ws.Cells(EQ_RN, EQ_C_BIASABS)).Value
    For i = 1 To UBound(dados, 1)
        If UCase$(Trim$(CStr(dados(i, EQ_C_ANALITO)))) <> UCase$(Trim$(analito)) Then GoTo prox
        If Not IsNumeric(dados(i, EQ_C_ANO)) Then GoTo prox
        ano = CLng(Val(CStr(dados(i, EQ_C_ANO))))
        If ano < a1 Or ano > a2 Then GoTo prox
        If Not IsNumeric(dados(i, EQ_C_BIAS)) Then GoTo prox
        n = n + 1
        somaSig = somaSig + CDbl(dados(i, EQ_C_BIAS))
        somaAbs = somaAbs + CDbl(dados(i, EQ_C_BIASABS))
        s = s & "rodada=" & CStr(dados(i, EQ_C_RODADA)) & _
                "|ano=" & CStr(ano) & _
                "|Xlab=" & Format$(dados(i, EQ_C_XLAB), "0.####") & _
                "|Xref=" & Format$(dados(i, EQ_C_XREF), "0.####") & _
                "|bias=" & Format$(dados(i, EQ_C_BIAS), "0.####") & ";"
prox:
    Next i
    If n = 0 Then BiasEQMemoria = "sem rodada utilizavel": Exit Function
    BiasEQMemoria = s & " CONSOLIDADO n=" & n & _
                    " media|bias|=" & Format$(somaAbs / n, "0.######") & _
                    " mediaAssinada=" & Format$(somaSig / n, "0.######")
End Function
