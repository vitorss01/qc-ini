Attribute VB_Name = "mQualidade"
Option Explicit

' mQualidade - ADR-033: as escadas de classificacao, em um lugar so
'
' POR QUE UM MODULO SO PARA ISTO
'
' A classificacao de Sigma aparece em tres consumidores: a coluna M da
' Estatistica, a coluna O do Painel e a tabela fato do Power BI. Escrita tres
' vezes, ela divergiria na primeira vez que alguem mexesse em uma faixa -- e a
' divergencia so apareceria quando o gestor comparasse a planilha com o
' relatorio do BI, tarde demais. Aqui ela existe uma vez; os tres a chamam.
'
' AS FAIXAS
'
' Westgard S, Bayat H, Westgard JO. Analytical Sigma metrics: A review of Six
' Sigma implementation tools for medical laboratories. Biochem Med (Zagreb)
' 2018;28(2):020502, pagina 6, descrevendo as zonas do Method Decision Chart:
' "world class quality (Six Sigma), the zone closest to the graph's origin,
' followed by a Five Sigma zone (Excellent), Four Sigma zone (Good), Three
' Sigma zone (Marginal), Two Sigma zone (Poor), and the rest of the graph, for
' Sigma metrics performance below Two Sigma, is labelled unacceptable."
'
' O artigo distingue SEIS zonas; o produto usa CINCO, unindo "Poor" (2 a 3) e
' "unacceptable" (abaixo de 2) sob "Inadequado". A uniao e decisao de produto,
' nao do artigo: abaixo de 3 Sigma a conduta operacional e a mesma, e o gestor
' pediu cinco faixas. Separar as duas e trocar uma linha em ClassificarSigma.
'
' SIGMA BAIXO NAO REPROVA CORRIDA
'
' Mesmo artigo, paginas 8 e 9: metodo de Sigma baixo exige mais regras, limites
' mais estreitos, mais controles e CQ mais frequente -- "For Three Sigma
' methods and lower, however, QC frequency must be greatly increased, to
' something closer to one control per 100 or per 50 patient samples". A
' classificacao qualifica o METODO. Quem reprova CORRIDA e Westgard.
'
' NUNCA DEVOLVER Empty
'
' IsNumeric(Empty) devolve True em VBA, e CDbl(Empty) devolve 0. Sem guarda de
' IsEmpty/IsNull, uma celula vazia seria classificada como "Inadequado" e uma
' margem ausente como "ETp excedido" -- reprovacao inventada em cima de nada.
'
' Uma UDF que devolve Empty e renderizada como 0 pela celula. O projeto ja
' pagou esse pedagio duas vezes (AlvoDoLote e BiasEQ). Estas funcoes devolvem
' String; quando nao ha o que classificar, devolvem vazio ("").

' ADR-103 (26/11/2026) -- REGRA INSTITUCIONAL E IMUTAVEL DA METRICA SIGMA (decisao do usuario).
' Faixas, operadores, classificacoes e status reproduzidos LITERALMENTE; nao alterar em revisao alguma:
'   sigma >= 6      Classificacao "Excelente — Classe mundial"   Status "Excelente"
'   5 < sigma < 6   Classificacao "Excelente"                   Status "Alto desempenho"
'   4 < sigma <= 5  Classificacao "Bom"                         Status "Desempenho adequado"
'   3 <= sigma <= 4 Classificacao "Marginal"                    Status "Atenção"
'   sigma < 3       Classificacao "Insatisfatório"              Status "método mais instável ou tendencioso"
' sigma = 4 e MARGINAL (expressao matematica; decisao do usuario em 26/11/2026). sigma = 5 e BOM.
' A classificacao NAO le mais a tblPlanoQC_Sigma: a tabela tem faixas fechadas embaixo (Sigma_Min <= s < Sigma_Max) e
' continua sendo SO a escolha das regras de Westgard/N/run size (ADR-035/101), que nao muda. Quem conflitar com esta
' regra se adapta a ela. Acentos e travessao por ChrW (a fonte e UTF-8; o VBE e ANSI).

Public Const MRG_EXCEDIDO As String = "ETp excedido"
Public Const MRG_CRITICA As String = "Margem critica"
Public Const MRG_DENTRO As String = "Dentro do orcamento"

' Abaixo deste limite a margem de ETp e considerada critica (em % do ETp).
Public Const MARGEM_CRITICA_PCT As Double = 10#


' Faixa institucional do Sigma: 1 (>= 6) a 5 (< 3); 0 = sem Sigma (vazio, texto, Null). Unica escada do produto.
Public Function FaixaSigma(ByVal sigma As Variant) As Long
    FaixaSigma = 0
    If IsEmpty(sigma) Then Exit Function
    If IsNull(sigma) Then Exit Function
    If IsError(sigma) Then Exit Function
    If Not IsNumeric(sigma) Then Exit Function
    If VarType(sigma) = vbString Then
        If Trim$(CStr(sigma)) = "" Then Exit Function
    End If
    Dim s As Double
    s = CDbl(sigma)
    If s >= 6# Then
        FaixaSigma = 1                          ' sigma >= 6
    ElseIf s > 5# And s < 6# Then
        FaixaSigma = 2                          ' 5 < sigma < 6
    ElseIf s > 4# And s <= 5# Then
        FaixaSigma = 3                          ' 4 < sigma <= 5
    ElseIf s >= 3# And s <= 4# Then
        FaixaSigma = 4                          ' 3 <= sigma <= 4
    ElseIf s < 3# Then
        FaixaSigma = 5                          ' sigma < 3
    End If
End Function

' Classificacao institucional (texto literal). "" quando nao ha Sigma.
Public Function ClassificarSigma(ByVal sigma As Variant) As String
    Select Case FaixaSigma(sigma)
        Case 1: ClassificarSigma = "Excelente " & ChrW(8212) & " Classe mundial"
        Case 2: ClassificarSigma = "Excelente"
        Case 3: ClassificarSigma = "Bom"
        Case 4: ClassificarSigma = "Marginal"
        Case 5: ClassificarSigma = "Insatisfat" & ChrW(243) & "rio"
        Case Else: ClassificarSigma = ""
    End Select
End Function

' Status institucional (texto literal). "" quando nao ha Sigma.
Public Function StatusSigma(ByVal sigma As Variant) As String
    Select Case FaixaSigma(sigma)
        Case 1: StatusSigma = "Excelente"
        Case 2: StatusSigma = "Alto desempenho"
        Case 3: StatusSigma = "Desempenho adequado"
        Case 4: StatusSigma = "Aten" & ChrW(231) & ChrW(227) & "o"
        Case 5: StatusSigma = "m" & ChrW(233) & "todo mais inst" & ChrW(225) & "vel ou tendencioso"
        Case Else: StatusSigma = ""
    End Select
End Function


' Situacao do orcamento de erro: quanto do ETp ainda sobra depois do ET
' observado. Recebe a margem JA em porcentagem do ETp.
Public Function ClassificarMargem(ByVal margemPct As Variant) As String
    ClassificarMargem = ""
    If IsEmpty(margemPct) Then Exit Function
    If IsNull(margemPct) Then Exit Function
    If Not IsNumeric(margemPct) Then Exit Function
    If VarType(margemPct) = vbString Then
        If Trim$(CStr(margemPct)) = "" Then Exit Function
    End If

    Dim m As Double
    m = CDbl(margemPct)

    If m < 0# Then
        ' O erro total observado ja passou do permitido: nao ha orcamento.
        ClassificarMargem = MRG_EXCEDIDO
    ElseIf m <= MARGEM_CRITICA_PCT Then
        ' Margem zero cai aqui, e nao em "excedido": ainda nao estourou.
        ClassificarMargem = MRG_CRITICA
    Else
        ClassificarMargem = MRG_DENTRO
    End If
End Function


' Margem de ETp em pontos percentuais. Devolve Null (nao Empty, nao 0) quando
' falta insumo -- a celula mostra vazio, e o BI grava campo nulo.
Public Function MargemETp(ByVal etp As Variant, ByVal et As Variant) As Variant
    MargemETp = Null
    If IsEmpty(etp) Or IsEmpty(et) Or IsNull(etp) Or IsNull(et) Then Exit Function
    If Not IsNumeric(etp) Or Not IsNumeric(et) Then Exit Function
    MargemETp = CDbl(etp) - CDbl(et)
End Function


' Margem de ETp como porcentagem do proprio ETp.
Public Function MargemETpPct(ByVal etp As Variant, ByVal et As Variant) As Variant
    MargemETpPct = Null
    If IsEmpty(etp) Or IsEmpty(et) Or IsNull(etp) Or IsNull(et) Then Exit Function
    If Not IsNumeric(etp) Or Not IsNumeric(et) Then Exit Function
    If CDbl(etp) = 0# Then Exit Function
    MargemETpPct = (CDbl(etp) - CDbl(et)) / CDbl(etp) * 100#
End Function
