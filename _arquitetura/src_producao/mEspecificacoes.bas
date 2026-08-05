Attribute VB_Name = "mEspecificacoes"
Option Explicit
' ===== ESPECIFICACOES DE QUALIDADE (ADR-022) =====
'
' A meta analitica define CONTRA O QUE o laboratorio e julgado. Este modulo e a
' UNICA camada que a resolve. Nenhuma aba de interface volta a calcular meta --
' vale o ADR-019 sem excecao.
'
'   DB_Especificacoes  ->  mEspecificacoes  ->  Painel / Estatistica / Calc
'
' FONTE E DADO; MODELO E UM CONJUNTO FECHADO
'
' Cada fonte informa grandezas diferentes, e e isso que decide a conta:
'
'   ETP_DIRETO      a fonte da o ETp        -> CVtp = ETp/3
'   VB              a fonte da CVi/CVg      -> CVtp = CVi*fi
'                                              BIAStp = RAIZ(CVi^2+CVg^2)*fb
'                                              ETp = BIAStp + 1,65*CVtp
'   CV_BIAS_DIRETO  a fonte da CVtp/BIAStp  -> ETp = BIAStp + 1,65*CVtp
'
' CLIA usa ETP_DIRETO, Variacao Biologica usa VB, Fabricante usa
' CV_BIAS_DIRETO. Ricos, EFLM, CAP, RCPA e Rilibak entram como LINHA no banco,
' escolhendo um dos tres modelos -- sem tocar em codigo. Uma fonte com
' matematica realmente nova exige um modelo novo, e dizer isso e mais honesto
' do que prometer extensibilidade infinita.
'
' A META E RESOLVIDA PELO ANO DO RESULTADO, NUNCA PELO ANO CORRENTE
'
' Regra de VIGENCIA: vale a especificacao de maior Ano que seja <= ano do
' resultado. Um resultado de 2025 continua julgado pela meta de 2025 em 2035, e
' cadastrar a meta de 2027 nao reescreve o passado. Casar por ano exato seria
' pior: abriria buraco em todo ano sem cadastro.

Public Const ESPEC As String = "DB_Especificacoes"
Public Const ESPEC_R0 As Long = 4

' colunas de DB_Especificacoes
Public Const ES_ID As Long = 1
Public Const ES_ANO As Long = 2
Public Const ES_FONTE As Long = 3
Public Const ES_MODELO As Long = 4
Public Const ES_ANALITO As Long = 5
Public Const ES_ETP As Long = 6        ' entrada do modelo ETP_DIRETO
Public Const ES_CVI As Long = 7        ' entradas do modelo VB
Public Const ES_CVG As Long = 8
Public Const ES_RIGOR As Long = 9
Public Const ES_CVTP As Long = 10      ' entradas do modelo CV_BIAS_DIRETO
Public Const ES_BIASTP As Long = 11
Public Const ES_ATIVO As Long = 12
Public Const ES_USUARIO As Long = 13
Public Const ES_DATACAD As Long = 14
Public Const ES_NCOL As Long = 14

Public Const MOD_ETP As String = "ETP_DIRETO"
Public Const MOD_VB As String = "VB"
Public Const MOD_CVBIAS As String = "CV_BIAS_DIRETO"

' Cache da resolucao. Sem ele, desenhar o Painel refaz a mesma varredura do
' banco para cada analito e cada nivel.
Private mCache As Object
Private mCacheCarimbo As String

' ---------------------------------------------------------------------------
' Fatores de rigor. Sao os mesmos ja usados na aba Analitos (U4:W6) e os mesmos
' que o gestor especificou: nao ha duas tabelas de fator no sistema.
Public Function FatorCV(ByVal rigor As String) As Double
    Select Case NormRigor(rigor)
        Case "OTI": FatorCV = 0.25
        Case "DES": FatorCV = 0.5
        Case Else:  FatorCV = 0.75      ' MIN, e tambem o default seguro
    End Select
End Function

Public Function FatorBias(ByVal rigor As String) As Double
    Select Case NormRigor(rigor)
        Case "OTI": FatorBias = 0.125
        Case "DES": FatorBias = 0.25
        Case Else:  FatorBias = 0.375
    End Select
End Function

' "OTI" e "OTI" com acento sao o mesmo rigor. O usuario digita como quiser; o
' banco guarda ASCII.
Public Function NormRigor(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, ChrW(211), "O")   ' O agudo
    t = Replace(t, ChrW(243), "O")
    If Left$(t, 3) = "OTI" Then NormRigor = "OTI": Exit Function
    If Left$(t, 3) = "DES" Then NormRigor = "DES": Exit Function
    NormRigor = "MIN"
End Function

' ---------------------------------------------------------------------------
' Calcula as tres metas de UMA linha do banco, conforme o modelo dela.
' Devolve "CVtp|BIAStp|ETp", com "" onde a grandeza nao e definida pelo modelo.
Public Function MetasDaLinha(ByVal modelo As String, ByVal etp As Variant, _
                             ByVal cvi As Variant, ByVal cvg As Variant, _
                             ByVal rigor As String, ByVal cvtp As Variant, _
                             ByVal biastp As Variant) As String
    Dim vCV As Double, vBias As Double, vETP As Double
    Dim temCV As Boolean, temBias As Boolean, temETP As Boolean

    Select Case UCase$(Trim$(modelo))
        Case MOD_ETP
            If IsNumeric(etp) Then
                If Len(Trim$(CStr(etp))) > 0 Then
                    vETP = CDbl(etp): temETP = True
                    vCV = vETP / 3: temCV = True
                End If
            End If

        Case MOD_VB
            If IsNumeric(cvi) And IsNumeric(cvg) Then
                If Len(Trim$(CStr(cvi))) > 0 And Len(Trim$(CStr(cvg))) > 0 Then
                    vCV = CDbl(cvi) * FatorCV(rigor): temCV = True
                    vBias = Sqr(CDbl(cvi) ^ 2 + CDbl(cvg) ^ 2) * FatorBias(rigor): temBias = True
                    vETP = vBias + 1.65 * vCV: temETP = True
                End If
            End If

        Case MOD_CVBIAS
            If IsNumeric(cvtp) And IsNumeric(biastp) Then
                If Len(Trim$(CStr(cvtp))) > 0 And Len(Trim$(CStr(biastp))) > 0 Then
                    vCV = CDbl(cvtp): temCV = True
                    vBias = CDbl(biastp): temBias = True
                    vETP = vBias + 1.65 * vCV: temETP = True
                End If
            End If
    End Select

    ' Str(), nao CStr().
    '
    ' CStr usa o separador decimal DA MAQUINA: numa maquina pt-BR devolve
    ' "3,33333" e o "," vira parte do protocolo. Protocolo legivel por maquina
    ' nao pode depender de localidade -- e a mesma familia do defeito que
    ' gravava "92,0028" e o Excel relia como 920028 (item 2.2 do gate).
    ' Str/Val sao o par invariante do VBA: sempre ponto, sempre reversivel.
    MetasDaLinha = IIf(temCV, Trim$(Str$(vCV)), "") & "|" & _
                   IIf(temBias, Trim$(Str$(vBias)), "") & "|" & _
                   IIf(temETP, Trim$(Str$(vETP)), "")
End Function

' ---------------------------------------------------------------------------
' Resolve a especificacao vigente.
'
' Devolve "OK|CVtp|BIAStp|ETp|Fonte|AnoVigente|Rigor|ID"  ou  "NAO_CADASTRADA".
'
' fonte = "" usa a fonte padrao de Cfg_Especificacoes; qualquer outra usa a
' fonte pedida. O ano e o DO RESULTADO, e a vigencia e "maior Ano <= ano".
Public Function ResolverEspec(ByVal analito As String, ByVal ano As Long, _
                              Optional ByVal fonte As String = "") As String
    Dim chave As String
    If Trim$(fonte) = "" Then fonte = FontePadrao()
    chave = UCase$(Trim$(analito)) & "|" & ano & "|" & UCase$(Trim$(fonte))

    If mCache Is Nothing Or mCacheCarimbo <> CarimboBanco() Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCacheCarimbo = CarimboBanco()
    End If
    If mCache.Exists(chave) Then ResolverEspec = mCache(chave): Exit Function

    Dim ws As Worksheet, dados As Variant, i As Long
    Dim melhorAno As Long, melhorLin As Long
    Set ws = ThisWorkbook.Sheets(ESPEC)
    dados = CarregarEspec()
    melhorAno = -1: melhorLin = 0

    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Trim$(CStr(dados(i, ES_ANALITO))) <> "" Then
                If UCase$(Trim$(CStr(dados(i, ES_ANALITO)))) = UCase$(Trim$(analito)) Then
                    If UCase$(Trim$(CStr(dados(i, ES_FONTE)))) = UCase$(Trim$(fonte)) Then
                        If UCase$(Trim$(CStr(dados(i, ES_ATIVO)))) <> "NAO" Then
                            If IsNumeric(dados(i, ES_ANO)) Then
                                Dim a As Long: a = CLng(dados(i, ES_ANO))
                                ' VIGENCIA: maior ano que nao ultrapassa o ano
                                ' do resultado. Empate fica com a linha mais
                                ' recente, que e a de baixo.
                                If a <= ano And a >= melhorAno Then
                                    melhorAno = a: melhorLin = i
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If melhorLin = 0 Then
        mCache.Add chave, "NAO_CADASTRADA"
        ResolverEspec = "NAO_CADASTRADA"
        Exit Function
    End If

    Dim m As String, p() As String
    m = MetasDaLinha(CStr(dados(melhorLin, ES_MODELO)), dados(melhorLin, ES_ETP), _
                     dados(melhorLin, ES_CVI), dados(melhorLin, ES_CVG), _
                     CStr(dados(melhorLin, ES_RIGOR)), dados(melhorLin, ES_CVTP), _
                     dados(melhorLin, ES_BIASTP))
    p = Split(m, "|")

    Dim r As String
    r = "OK|" & p(0) & "|" & p(1) & "|" & p(2) & "|" & _
        CStr(dados(melhorLin, ES_FONTE)) & "|" & melhorAno & "|" & _
        NormRigor(CStr(dados(melhorLin, ES_RIGOR))) & "|" & _
        CStr(dados(melhorLin, ES_ID))
    mCache.Add chave, r
    ResolverEspec = r
End Function

' Atalhos de leitura. Devolvem Empty quando nao ha especificacao -- e nao zero:
' zero seria interpretado como "meta zero", que reprova tudo.
Public Function EspecCVtp(ByVal analito As String, ByVal ano As Long, Optional ByVal fonte As String = "") As Variant
    EspecCVtp = CampoEspec(analito, ano, fonte, 1)
End Function
Public Function EspecBIAStp(ByVal analito As String, ByVal ano As Long, Optional ByVal fonte As String = "") As Variant
    EspecBIAStp = CampoEspec(analito, ano, fonte, 2)
End Function
Public Function EspecETp(ByVal analito As String, ByVal ano As Long, Optional ByVal fonte As String = "") As Variant
    EspecETp = CampoEspec(analito, ano, fonte, 3)
End Function

Private Function CampoEspec(ByVal analito As String, ByVal ano As Long, _
                            ByVal fonte As String, ByVal pos As Long) As Variant
    Dim r As String, p() As String
    r = ResolverEspec(analito, ano, fonte)
    If Left$(r, 2) <> "OK" Then CampoEspec = Empty: Exit Function
    p = Split(r, "|")
    If Len(p(pos)) = 0 Then CampoEspec = Empty Else CampoEspec = Val(p(pos))
End Function

' ---------------------------------------------------------------------------
' Conformidade: o desempenho medido esta dentro da meta?
'
' Devolve "CONFORME" | "NAO CONFORME" | "SEM META".
' SEM META nao e aprovacao: e ausencia de criterio, e precisa aparecer como
' tal. Tratar falta de especificacao como conformidade e o jeito mais silencioso
' de um laboratorio se achar em dia.
Public Function AvaliarConformidade(ByVal analito As String, ByVal ano As Long, _
                                    ByVal cvReal As Variant, ByVal biasReal As Variant, _
                                    Optional ByVal fonte As String = "") As String
    Dim r As String, p() As String
    r = ResolverEspec(analito, ano, fonte)
    If Left$(r, 2) <> "OK" Then AvaliarConformidade = "SEM META": Exit Function
    p = Split(r, "|")

    Dim etReal As Variant, ruim As Boolean, avaliou As Boolean

    If Len(p(1)) > 0 And IsNumeric(cvReal) Then
        avaliou = True
        If CDbl(cvReal) > Val(p(1)) Then ruim = True
    End If
    If Len(p(2)) > 0 And IsNumeric(biasReal) Then
        avaliou = True
        If Abs(CDbl(biasReal)) > Val(p(2)) Then ruim = True
    End If
    If Len(p(3)) > 0 And IsNumeric(cvReal) And IsNumeric(biasReal) Then
        avaliou = True
        etReal = Abs(CDbl(biasReal)) + 1.65 * CDbl(cvReal)
        If CDbl(etReal) > Val(p(3)) Then ruim = True
    End If

    If Not avaliou Then AvaliarConformidade = "SEM META": Exit Function
    AvaliarConformidade = IIf(ruim, "NAO CONFORME", "CONFORME")
End Function

' ---------------------------------------------------------------------------
Public Function UltimaLinhaEspec() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ESPEC)
    UltimaLinhaEspec = ws.Cells(ws.Rows.Count, ES_ID).End(xlUp).Row
    If UltimaLinhaEspec < ESPEC_R0 Then UltimaLinhaEspec = ESPEC_R0 - 1
End Function

Public Function CarregarEspec() As Variant
    Dim ws As Worksheet, ult As Long
    Set ws = ThisWorkbook.Sheets(ESPEC)
    ult = UltimaLinhaEspec()
    If ult < ESPEC_R0 Then
        CarregarEspec = Empty
    Else
        CarregarEspec = ws.Range(ws.Cells(ESPEC_R0, ES_ID), ws.Cells(ult, ES_NCOL)).Value
    End If
End Function

' Carimbo barato para invalidar o cache quando o banco muda.
Private Function CarimboBanco() As String
    CarimboBanco = CStr(UltimaLinhaEspec()) & "|" & CStr(FontePadrao())
End Function

Public Sub InvalidarCacheEspec()
    Set mCache = Nothing
    mCacheCarimbo = ""
End Sub

Public Function FontePadrao() As String
    Dim ws As Worksheet, v As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    On Error GoTo 0
    If ws Is Nothing Then FontePadrao = "CLIA": Exit Function
    v = Trim$(CStr(ws.Range("B2").Value))
    If v = "" Then v = "CLIA"
    FontePadrao = v
End Function

Public Function RigorPadrao() As String
    Dim ws As Worksheet, v As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    On Error GoTo 0
    If ws Is Nothing Then RigorPadrao = "MIN": Exit Function
    v = Trim$(CStr(ws.Range("B3").Value))
    RigorPadrao = NormRigor(v)
End Function

' Fontes cadastradas, na ordem em que aparecem em Cfg_Especificacoes.
Public Function ListaFontes() As Collection
    Dim ws As Worksheet, c As Collection, i As Long, v As String
    Set c = New Collection
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    On Error GoTo 0
    If ws Is Nothing Then Set ListaFontes = c: Exit Function
    For i = 7 To 40
        v = Trim$(CStr(ws.Cells(i, 1).Value))
        If v <> "" Then c.Add v
    Next i
    Set ListaFontes = c
End Function

' Modelo declarado para uma fonte no cadastro de fontes.
Public Function ModeloDaFonte(ByVal fonte As String) As String
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    On Error GoTo 0
    If ws Is Nothing Then ModeloDaFonte = MOD_ETP: Exit Function
    For i = 7 To 40
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(fonte)) Then
            ModeloDaFonte = Trim$(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    ModeloDaFonte = MOD_ETP
End Function

' ---------------------------------------------------------------------------
' Grava uma especificacao. Chave logica: Ano + Fonte + Analito -- reenviar a
' mesma chave ATUALIZA, nao duplica, pelo mesmo principio do UpsertResultados.
Public Function GravarEspec(ByVal ano As Long, ByVal fonte As String, ByVal analito As String, _
                            ByVal etp As Variant, ByVal cvi As Variant, ByVal cvg As Variant, _
                            ByVal rigor As String, ByVal cvtp As Variant, ByVal biastp As Variant) As String
    Dim ws As Worksheet, dados As Variant, i As Long, lin As Long, modelo As String
    Set ws = ThisWorkbook.Sheets(ESPEC)
    modelo = ModeloDaFonte(fonte)

    dados = CarregarEspec()
    lin = 0
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If CStr(dados(i, ES_ANO)) = CStr(ano) _
               And UCase$(Trim$(CStr(dados(i, ES_FONTE)))) = UCase$(Trim$(fonte)) _
               And UCase$(Trim$(CStr(dados(i, ES_ANALITO)))) = UCase$(Trim$(analito)) Then
                lin = ESPEC_R0 + i - 1
                Exit For
            End If
        Next i
    End If

    Dim novoID As String
    If lin = 0 Then
        lin = UltimaLinhaEspec() + 1
        If lin < ESPEC_R0 Then lin = ESPEC_R0
        novoID = "ESP-" & Format$(lin - ESPEC_R0 + 1, "000000")
        ws.Cells(lin, ES_ID).Value = novoID
    Else
        novoID = CStr(ws.Cells(lin, ES_ID).Value)
    End If

    ws.Cells(lin, ES_ANO).Value = ano
    ws.Cells(lin, ES_FONTE).Value = fonte
    ws.Cells(lin, ES_MODELO).Value = modelo
    ws.Cells(lin, ES_ANALITO).Value = analito
    ws.Cells(lin, ES_ETP).Value = etp
    ws.Cells(lin, ES_CVI).Value = cvi
    ws.Cells(lin, ES_CVG).Value = cvg
    ws.Cells(lin, ES_RIGOR).Value = IIf(Trim$(rigor) = "", "", NormRigor(rigor))
    ws.Cells(lin, ES_CVTP).Value = cvtp
    ws.Cells(lin, ES_BIASTP).Value = biastp
    ws.Cells(lin, ES_ATIVO).Value = "Sim"
    ws.Cells(lin, ES_USUARIO).Value = Environ$("USERNAME")
    ws.Cells(lin, ES_DATACAD).Value = Now

    InvalidarCacheEspec
    RegistrarLog "ESPEC_GRAVADA", novoID & " " & fonte & " " & ano & " " & analito
    GravarEspec = novoID
End Function
