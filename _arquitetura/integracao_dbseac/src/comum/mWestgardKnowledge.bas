Attribute VB_Name = "mWestgardKnowledge"
Option Explicit
' Identificadores fortes — evitam comparacao por string espalhada pelo codigo.
Public Enum eWestgard
    wg12s = 1
    wg13s = 2
    wg22s = 3
    wgR4s = 4
    wg41s = 5
    wg10x = 6
End Enum

' Estrutura unica que descreve uma regra.
Public Type WestgardRuleInfo
    codigo As String
    categoria As String
    Criterio As String
    Interpretacao As String
    Causas As String
    Sugestoes As String
End Type

Public Function RegraCodigo(ByVal r As eWestgard) As String
    Select Case r
        Case wg12s: RegraCodigo = "12s"
        Case wg13s: RegraCodigo = "13s"
        Case wg22s: RegraCodigo = "22s"
        Case wgR4s: RegraCodigo = "R4s"
        Case wg41s: RegraCodigo = "41s"
        Case wg10x: RegraCodigo = "10x"
    End Select
End Function

Public Function RegraDe(ByVal codigo As String) As eWestgard
    Select Case UCase$(Trim$(codigo))
        Case "12S": RegraDe = wg12s
        Case "13S": RegraDe = wg13s
        Case "22S": RegraDe = wg22s
        Case "R4S": RegraDe = wgR4s
        Case "41S": RegraDe = wg41s
        Case "10X": RegraDe = wg10x
    End Select
End Function

' Ponto unico de consulta: devolve TUDO sobre a regra numa estrutura so.
Public Function RegraInfo(ByVal r As eWestgard) As WestgardRuleInfo
    Dim info As WestgardRuleInfo, c As String
    c = RegraCodigo(r)
    info.codigo = c
    info.categoria = RegraClassificacao(c)
    info.Criterio = RegraCriterio(c)
    info.Interpretacao = RegraInterpretacao(c)
    info.Causas = RegraCausas(c)
    info.Sugestoes = RegraSugestoes(c)
    RegraInfo = info
End Function
' ============================================================================
'  mWestgardKnowledge — conhecimento tecnico das regras de Westgard
'  Fonte unica de verdade: classificacao, interpretacao, causas e sugestoes.
'  NENHUM outro modulo deve duplicar estes textos.
'  Referencia: Westgard JO et al. Clin Chem 1981;27:493-501; CLSI C24.
' ============================================================================

' O motor (mEstatistica) emite os codigos COM separador: 1_3s, 2_2s, 2of3_2s, R_4s, 4_1s,
' 3_1s, 3x/6x/8x/10x. Esta base so conhecia 12S/13S/22S/R4S/41S/10X, entao NENHUMA regra
' real casava: o Painel (Classificacao) e a aba Eventos_Westgard mostravam "Nao classificada"
' para tudo (achado de 03/10/2026). Familia junta cada codigo a regra-mae da base.
Private Function Familia(ByVal regra As String) As String
    Dim t As String
    t = UCase$(Replace$(Replace$(Trim$(regra), "_", ""), " ", ""))
    Select Case t
        Case "12S", "13S", "22S", "R4S", "41S", "10X": Familia = t
        Case "2OF32S": Familia = "22S"
        Case "31S": Familia = "41S"
        Case Else
            Familia = t
            If Len(t) > 1 Then
                If Right$(t, 1) = "X" And IsNumeric(Left$(t, Len(t) - 1)) Then Familia = "10X"
            End If
    End Select
End Function

' ALERTA -> 12s | ERRO ALEATORIO -> 13s, R4s | ERRO SISTEMATICO -> 22s (2of3_2s), 41s (3_1s), Nx
Public Function RegraClassificacao(ByVal regra As String) As String
    Select Case Familia(regra)
        Case "12S": RegraClassificacao = "Alerta"
        Case "13S", "R4S": RegraClassificacao = "Erro Aleatório"
        Case "22S", "41S", "10X": RegraClassificacao = "Erro Sistemático"
        Case Else: RegraClassificacao = "Não classificada"
    End Select
End Function

Public Function RegraCriterio(ByVal regra As String) As String
    Dim t As String
    t = UCase$(Replace$(Replace$(Trim$(regra), "_", ""), " ", ""))
    If t = "31S" Then RegraCriterio = "3 controles consecutivos excedem 1DP do mesmo lado": Exit Function
    If t = "2OF32S" Then RegraCriterio = "2 de 3 controles (3 níveis da mesma RUN ou 3 RUNs do mesmo nível) excedem 2DP do mesmo lado": Exit Function
    If Familia(regra) = "10X" And t <> "10X" Then
        RegraCriterio = Left$(t, Len(t) - 1) & " controles consecutivos do mesmo lado da média": Exit Function
    End If
    Select Case Familia(regra)
        Case "12S": RegraCriterio = "1 controle excede média ± 2DP"
        Case "13S": RegraCriterio = "1 controle excede média ± 3DP"
        Case "22S": RegraCriterio = "2 controles consecutivos (ou 2 níveis da mesma RUN) excedem 2DP do mesmo lado"
        Case "R4S": RegraCriterio = "Amplitude entre 2 controles da mesma RUN excede 4DP"
        Case "41S": RegraCriterio = "4 controles consecutivos excedem 1DP do mesmo lado"
        Case "10X": RegraCriterio = "10 controles consecutivos do mesmo lado da média"
        Case Else: RegraCriterio = ""
    End Select
End Function

Public Function RegraInterpretacao(ByVal regra As String) As String
    Select Case Familia(regra)
        Case "12S"
            RegraInterpretacao = "Um controle ultrapassou 2 DP. É apenas um ALERTA — isoladamente não " & _
                "rejeita a corrida, pois cerca de 5% dos resultados excedem 2DP por variação aleatória esperada."
        Case "13S"
            RegraInterpretacao = "Um controle ultrapassou 3 DP. Indica erro aleatório de magnitude " & _
                "incompatível com a variação esperada do processo. A corrida deve ser rejeitada."
        Case "22S"
            RegraInterpretacao = "Dois controles consecutivos, ou dois níveis da mesma corrida, " & _
                "ultrapassaram 2 DP para o MESMO lado. Padrão típico de erro sistemático (deslocamento)."
        Case "R4S"
            RegraInterpretacao = "Dois controles da mesma corrida apresentaram diferença superior a 4 DP " & _
                "(um acima e outro abaixo). Indica aumento da imprecisão — erro aleatório."
        Case "41S"
            RegraInterpretacao = "Controles consecutivos (3 ou 4, conforme a regra) ultrapassaram 1 DP para o mesmo lado. " & _
                "Erro sistemático de menor magnitude, geralmente progressivo."
        Case "10X"
            RegraInterpretacao = "Controles consecutivos (3, 6, 8 ou 10, conforme a regra) ficaram do mesmo lado da média. " & _
                "Indica deslocamento de patamar (shift), mesmo sem violar limites de DP."
        Case Else: RegraInterpretacao = ""
    End Select
End Function

Public Function RegraCausas(ByVal regra As String) As String
    Select Case Familia(regra)
        Case "12S"
            RegraCausas = "• Variação aleatória esperada do processo" & vbCrLf & _
                          "• Início de instabilidade (observar as próximas corridas)"
        Case "13S"
            RegraCausas = "• Bolha de ar na aspiração" & vbCrLf & "• Erro de pipetagem" & vbCrLf & _
                          "• Material de controle mal homogeneizado" & vbCrLf & "• Instabilidade momentânea do equipamento"
        Case "22S"
            RegraCausas = "• Calibração deslocada" & vbCrLf & "• Troca de lote de reagente" & vbCrLf & _
                          "• Deterioração do material de controle" & vbCrLf & "• Variação de temperatura"
        Case "R4S"
            RegraCausas = "• Erro de pipetagem" & vbCrLf & "• Formação de bolhas" & vbCrLf & _
                          "• Aspiração irregular" & vbCrLf & "• Oscilação hidráulica do sistema"
        Case "41S"
            RegraCausas = "• Desgaste progressivo de reagente ou lâmpada" & vbCrLf & _
                          "• Calibração derivando" & vbCrLf & "• Envelhecimento do material de controle"
        Case "10X"
            RegraCausas = "• Recalibração recente" & vbCrLf & "• Mudança de lote de reagente ou controle" & vbCrLf & _
                          "• Média/DP do lote mal estabelecidos" & vbCrLf & "• Deterioração progressiva"
        Case Else: RegraCausas = ""
    End Select
End Function

Public Function RegraSugestoes(ByVal regra As String) As String
    Select Case Familia(regra)
        Case "12S"
            RegraSugestoes = "• Não repetir automaticamente" & vbCrLf & "• Registrar e observar a próxima corrida"
        Case "13S"
            RegraSugestoes = "• Repetir o controle" & vbCrLf & "• Homogeneizar o material" & vbCrLf & _
                             "• Conferir pipetas e sistema de aspiração"
        Case "22S"
            RegraSugestoes = "• Verificar calibração" & vbCrLf & "• Conferir lote e validade do reagente" & vbCrLf & _
                             "• Avaliar temperatura do equipamento" & vbCrLf & "• Não liberar resultados até resolver"
        Case "R4S"
            RegraSugestoes = "• Repetir o controle" & vbCrLf & "• Conferir pipetas" & vbCrLf & _
                             "• Verificar sistema de aspiração" & vbCrLf & "• Avaliar estabilidade do equipamento"
        Case "41S"
            RegraSugestoes = "• Verificar necessidade de recalibração" & vbCrLf & _
                             "• Conferir validade do reagente" & vbCrLf & "• Avaliar manutenção preventiva"
        Case "10X"
            RegraSugestoes = "• Reavaliar média e DP estabelecidos para o lote" & vbCrLf & _
                             "• Verificar recalibração recente" & vbCrLf & "• Confirmar integridade do material de controle"
        Case Else: RegraSugestoes = ""
    End Select
End Function

