Attribute VB_Name = "mPeriodo"
Option Explicit
' ============================================================================
'  JANELA DE PERIODO -- a conta de "T1 de 2026" virar duas datas (ADR-054)
'
'  Modulo PURO: nao le celula, nao escreve celula, nao conhece produto. Existe
'  porque a mesma pergunta e feita em dois lugares --
'
'    mEstatPeriodo  (aba Estatistica, so na Bioquimica)
'    mApp           (Ano/Periodo do Painel, nos DOIS produtos)
'
'  -- e porque a Hematologia nao tem mEstatPeriodo. Enquanto mApp chamava
'  mEstatPeriodo diretamente, o projeto VBA da Hematologia nao compilava:
'  "variavel nao definida" no qualificador de modulo derruba a compilacao
'  INTEIRA, e ai nenhuma macro do arquivo responde.
'
'  Reimplementar a conta no mApp resolveria o compilador e criaria o problema
'  de sempre: duas contas de trimestre que um dia divergem. Uma conta, um
'  lugar, os dois produtos.
' ============================================================================

' parte = mes (1..12) quando MENSAL, trimestre (1..4), semestre (1..2);
' ignorado quando ANUAL (o default seguro para qualquer rotulo desconhecido).
Public Function PeriodoInicio(ByVal visao As String, ByVal ano As Variant, _
                              ByVal parte As Variant) As Variant
    PeriodoInicio = Faixa(visao, ano, parte, True)
End Function

Public Function PeriodoFim(ByVal visao As String, ByVal ano As Variant, _
                           ByVal parte As Variant) As Variant
    PeriodoFim = Faixa(visao, ano, parte, False)
End Function

Private Function Faixa(ByVal visao As String, ByVal ano As Variant, _
                       ByVal parte As Variant, ByVal querInicio As Boolean) As Variant
    Dim v As String, a As Long, p As Long, m1 As Long, m2 As Long
    v = UCase$(Trim$(visao))
    If Not IsNumeric(ano) Then Faixa = "": Exit Function
    a = CLng(Val(CStr(ano)))
    If a < 1900 Then Faixa = "": Exit Function
    p = CLng(Val(CStr(parte)))

    Select Case v
        Case "MENSAL"
            If p < 1 Or p > 12 Then p = 1
            m1 = p: m2 = p
        Case "TRIMESTRAL"
            If p < 1 Or p > 4 Then p = 1
            m1 = (p - 1) * 3 + 1: m2 = m1 + 2
        Case "SEMESTRAL"
            If p < 1 Or p > 2 Then p = 1
            m1 = (p - 1) * 6 + 1: m2 = m1 + 5
        Case Else                       ' ANUAL e o default seguro
            m1 = 1: m2 = 12
    End Select

    If querInicio Then
        Faixa = DateSerial(a, m1, 1)
    Else
        Faixa = DateSerial(a, m2 + 1, 0)      ' dia 0 do mes seguinte = ultimo dia
    End If
End Function

' Traduz o rotulo do Painel ("T1", "S2", "Mar", "(todos)") para visao+parte.
' Um so lugar decide o que cada rotulo significa.
Public Sub VisaoDoRotulo(ByVal s As String, ByRef visao As String, ByRef parte As Long)
    Dim u As String, m As Long
    u = UCase$(Trim$(s))
    visao = "ANUAL": parte = 0
    If Left$(u, 4) = "TUDO" Then visao = "TUDO": Exit Sub
    If Len(u) = 2 And (Left$(u, 1) = "T" Or Left$(u, 1) = "S") And IsNumeric(Mid$(u, 2, 1)) Then
        If Left$(u, 1) = "T" Then
            visao = "TRIMESTRAL": parte = CLng(Mid$(u, 2, 1))
        Else
            visao = "SEMESTRAL": parte = CLng(Mid$(u, 2, 1))
        End If
        Exit Sub
    End If
    m = MesDoRotulo(s)
    If m > 0 Then
        visao = "MENSAL": parte = m
    End If
End Sub

Private Function MesDoRotulo(ByVal s As String) As Long
    Dim v As Variant, i As Long
    v = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", _
              "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
    For i = 0 To 11
        If UCase$(Left$(s, 3)) = v(i) Then MesDoRotulo = i + 1: Exit Function
    Next i
End Function
