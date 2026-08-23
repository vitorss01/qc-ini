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

Public Const CLS_MUNDIAL As String = "Classe mundial"
Public Const CLS_EXCELENTE As String = "Excelente"
Public Const CLS_BOM As String = "Bom"
Public Const CLS_MARGINAL As String = "Marginal"
Public Const CLS_INADEQUADO As String = "Desempenho inadequado"

Public Const MRG_EXCEDIDO As String = "ETp excedido"
Public Const MRG_CRITICA As String = "Margem critica"
Public Const MRG_DENTRO As String = "Dentro do orcamento"

' Abaixo deste limite a margem de ETp e considerada critica (em % do ETp).
Public Const MARGEM_CRITICA_PCT As Double = 10#


' Classifica o desempenho analitico de um metodo pela metrica Sigma.
' Devolve "" quando sigma nao e numero (sem EP, sem CV, sem ETp).
Public Function ClassificarSigma(ByVal sigma As Variant) As String
    ClassificarSigma = ""
    If IsEmpty(sigma) Then Exit Function
    If IsNull(sigma) Then Exit Function
    If Not IsNumeric(sigma) Then Exit Function
    If VarType(sigma) = vbString Then
        If Trim$(CStr(sigma)) = "" Then Exit Function
    End If

    ' A TABELA MANDA (ADR-035).
    '
    ' tblPlanoQC_Sigma ja carrega a classificacao de cada faixa, e e ela que
    ' define regras, N e run size. Manter aqui uma segunda escada de IFs faria
    ' as duas divergirem no dia em que alguem mexesse numa faixa -- e a
    ' divergencia so apareceria quando o gestor comparasse a coluna M com o
    ' plano de CQ ao lado. Foi assim que "Inadequado" e "Desempenho
    ' inadequado" conviveram no mesmo arquivo.
    '
    ' A escada abaixo continua como RESERVA: se a Cfg_PlanoQC nao existir na
    ' pasta, a classificacao ainda sai, em vez de a coluna inteira esvaziar.
    Dim daTabela As Variant
    daTabela = mPlanoQC.PlanoQC(sigma, "CLASSE")
    If VarType(daTabela) = vbString Then
        If Trim$(CStr(daTabela)) <> "" Then
            ClassificarSigma = CStr(daTabela)
            Exit Function
        End If
    End If

    Dim s As Double
    s = CDbl(sigma)

    If s >= 6# Then
        ClassificarSigma = CLS_MUNDIAL
    ElseIf s >= 5# Then
        ClassificarSigma = CLS_EXCELENTE
    ElseIf s >= 4# Then
        ClassificarSigma = CLS_BOM
    ElseIf s >= 3# Then
        ClassificarSigma = CLS_MARGINAL
    Else
        ' Aqui ficaria a separacao entre "Poor" (2 a 3) e "Inaceitavel" (< 2)
        ' do artigo, se o produto passar a usar as seis zonas.
        ClassificarSigma = CLS_INADEQUADO
    End If
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
