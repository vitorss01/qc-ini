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

' Camada de saida. Const de MODULO tem de ficar ANTES da primeira rotina: em
' VBA, declarada no meio do arquivo ela simplesmente nao existe para o
' compilador, e o erro sai como "variavel nao definida" na rotina que a usa --
' longe da causa. Mesmo lint que ja existe em gerar_mDados_audit.ps1.
Public Const ENGE As String = "Eng_Especificacoes"
Public Const ENGE_R0 As Long = 4
Private Const SENHA_ESPEC As String = "qcini2025"

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
        Case "MIN": FatorCV = 0.75
    End Select
End Function

Public Function FatorBias(ByVal rigor As String) As Double
    Select Case NormRigor(rigor)
        Case "OTI": FatorBias = 0.125
        Case "DES": FatorBias = 0.25
        Case "MIN": FatorBias = 0.375
    End Select
End Function

' "OTI" e "OTI" com acento sao o mesmo rigor. O usuario digita como quiser; o
' banco guarda ASCII.
Public Function NormRigor(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, ChrW(211), "O")
    t = Replace(t, ChrW(243), "O")
    If Left$(t, 3) = "OTI" Then NormRigor = "OTI": Exit Function
    If Left$(t, 3) = "DES" Then NormRigor = "DES": Exit Function
    If Left$(t, 3) = "MIN" Then NormRigor = "MIN"
End Function

' ---------------------------------------------------------------------------
' A linha do banco ainda esta em vigencia?
'
' O ACENTO IMPORTA, E ESTA E A RAZAO DESTA FUNCAO EXISTIR.
'
' Quem desativa uma especificacao e o gestor, digitando na propria aba. E "nao"
' se escreve "Nao" com til -- que e inclusive como o proprio sistema grava a
' coluna Elegivel da Cfg_Status. A comparacao anterior era contra "NAO" cru:
' UCase$("Nao" com til) devolve "NAO" com til, que NAO e igual a "NAO", entao a
' especificacao recem-desativada continuava valendo. Sem erro, sem aviso, e
' mudando veredito de conformidade -- o gestor via "Nao" na tela e o motor
' seguia aplicando a meta.
'
' POLARIDADE, e por que ela e o INVERSO da de EhElegivel (mEstatistica).
'
' La, estado desconhecido NUNCA entra no calculo: elegibilidade decide se um
' DADO entra na conta, e no escuro o certo e excluir o dado. Aqui, valor
' desconhecido CONTINUA valendo: vigencia decide se existe CRITERIO, e no
' escuro o certo e manter o criterio. Perder a meta em silencio jogaria o
' analito em "SEM ESPECIFICACAO" e o tiraria da avaliacao -- exatamente o que o
' ADR-023 existe para impedir. Em branco tambem vale, porque todo escritor do
' pipeline grava "Sim" e uma linha acrescentada a mao nao pode sumir calada.
Public Function EspecAtiva(ByVal v As Variant) As Boolean
    Dim t As String
    t = UCase$(Trim$(CStr(v)))
    t = Replace(t, ChrW(195), "A")      ' A com til (maiusculo, apos UCase)
    t = Replace(t, ChrW(227), "A")      ' a com til (minusculo, por seguranca)
    Select Case t
        Case "NAO", "N", "INATIVO", "INATIVA", "FALSO", "FALSE", "0"
            EspecAtiva = False
        Case Else
            EspecAtiva = True
    End Select
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
            If IsNumeric(etp) And Len(Trim$(CStr(etp))) > 0 Then
                vETP = CDbl(etp): temETP = True
                vCV = vETP / 3: temCV = True
            End If
        Case MOD_VB
            If NormRigor(rigor) <> "" Then
                If IsNumeric(cvi) And IsNumeric(cvg) Then
                    If Len(Trim$(CStr(cvi))) > 0 And Len(Trim$(CStr(cvg))) > 0 Then
                        vCV = CDbl(cvi) * FatorCV(rigor): temCV = True
                        vBias = Sqr(CDbl(cvi) ^ 2 + CDbl(cvg) ^ 2) * FatorBias(rigor): temBias = True
                        vETP = vBias + 1.65 * vCV: temETP = True
                    End If
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
                        If EspecAtiva(dados(i, ES_ATIVO)) Then
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

    ' "OK" SO QUANDO HA META UTILIZAVEL.
    '
    ' Antes, linha encontrada bastava para o prefixo ser OK -- mesmo com as
    ' tres grandezas vazias, por ETp textual ou modelo inexistente. Nenhum
    ' consumidor atual se enganava (todos conferem se o valor existe antes de
    ' usar), mas um consumidor futuro que testasse so Left(r,2)="OK" concluiria
    ' que ha meta. Contrato que so funciona porque quem chama e cuidadoso e
    ' contrato quebrado esperando a hora.
    '
    ' ESPEC_INVALIDA e distinto de NAO_CADASTRADA de proposito: alguem TENTOU
    ' cadastrar e o dado esta ruim. Tratar os dois como "nao existe" esconderia
    ' um problema de qualidade de dado atras de uma lacuna de cadastro.
    Dim r As String
    If Len(p(0)) = 0 And Len(p(1)) = 0 And Len(p(2)) = 0 Then
        r = "ESPEC_INVALIDA|" & CStr(dados(melhorLin, ES_ID)) & "|" & _
            MotivoInvalidez(CStr(dados(melhorLin, ES_MODELO)))
        mCache.Add chave, r
        ResolverEspec = r
        Exit Function
    End If
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

' Por que a linha cadastrada nao produziu meta.
Private Function MotivoInvalidez(ByVal modelo As String) As String
    Select Case UCase$(Trim$(modelo))
        Case MOD_ETP:    MotivoInvalidez = "ETp ausente ou nao numerico"
        Case MOD_VB:     MotivoInvalidez = "CVi/CVg ausentes ou nao numericos"
        Case MOD_CVBIAS: MotivoInvalidez = "CVtp/BIAStp ausentes ou nao numericos"
        Case "":         MotivoInvalidez = "modelo nao definido"
        Case Else:       MotivoInvalidez = "modelo desconhecido: " & modelo
    End Select
End Function

' ---------------------------------------------------------------------------
' Conformidade: QUATRO estados.
'
'   CONFORME           mediu e esta dentro        ninguem age
'   NAO CONFORME       mediu e esta fora          analista investiga a corrida
'   SEM ESPECIFICACAO  nao ha meta cadastrada     gestor da qualidade cadastra
'   SEM DADOS          ha meta, faltam medicoes   bancada roda controles
'
' Os dois ultimos estavam fundidos em "SEM META", e sao CAUSAS OPOSTAS COM
' DONOS OPOSTOS. Numa auditoria ISO 15189 "nao temos criterio" e "nao temos
' dado" sao achados diferentes: o primeiro e falha de gestao da qualidade, o
' segundo e lacuna operacional. Fundir os dois esconde qual e.
'
' Nenhum deles e aprovacao. Tratar ausencia -- de meta ou de dado -- como
' conformidade e o jeito mais silencioso de um laboratorio se achar em dia.
'
' Especificacao cadastrada mas INVALIDA cai em SEM ESPECIFICACAO: para a
' avaliacao nao ha criterio utilizavel. O motivo nao se perde -- fica visivel
' em ResolverEspec (ESPEC_INVALIDA) e em SituacaoEspec, onde e acionavel.
Public Function AvaliarConformidade(ByVal analito As String, ByVal ano As Long, _
                                    ByVal cvReal As Variant, ByVal biasReal As Variant, _
                                    Optional ByVal fonte As String = "") As String
    Dim r As String, p() As String
    r = ResolverEspec(analito, ano, fonte)
    If Left$(r, 2) <> "OK" Then AvaliarConformidade = "SEM ESPECIFICACAO": Exit Function
    p = Split(r, "|")

    ' Sem medicao nao ha o que comparar -- e isso NAO e falta de especificacao.
    Dim temCV As Boolean, temBias As Boolean
    temCV = Num(cvReal)
    temBias = Num(biasReal)
    If Not temCV And Not temBias Then AvaliarConformidade = "SEM DADOS": Exit Function

    Dim ruim As Boolean, comparou As Boolean, etReal As Double

    If Len(p(1)) > 0 And temCV Then
        comparou = True
        If CDbl(cvReal) > Val(p(1)) Then ruim = True
    End If
    If Len(p(2)) > 0 And temBias Then
        comparou = True
        If Abs(CDbl(biasReal)) > Val(p(2)) Then ruim = True
    End If
    If Len(p(3)) > 0 And temCV And temBias Then
        comparou = True
        etReal = Abs(CDbl(biasReal)) + 1.65 * CDbl(cvReal)
        If etReal > Val(p(3)) Then ruim = True
    End If

    ' Ha meta e ha medida, mas sem grandeza em comum entre as duas -- a meta so
    ' define BIAStp e so temos CV, por exemplo. Nao da para afirmar nada.
    If Not comparou Then AvaliarConformidade = "SEM DADOS": Exit Function
    AvaliarConformidade = IIf(ruim, "NAO CONFORME", "CONFORME")
End Function

' Numero utilizavel? String vazia e Empty passam por IsNumeric em alguns
' caminhos; aqui a pergunta e "da para comparar com isto".
Private Function Num(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    Num = IsNumeric(v)
End Function

' Situacao do CADASTRO, para exibicao e auditoria. Responde uma pergunta
' diferente da conformidade: nao "o laboratorio esta dentro da meta", e sim
' "existe meta utilizavel, e se nao, por que".
' "CADASTRADA" | "NAO CADASTRADA" | "INVALIDA: <motivo>"
Public Function SituacaoEspec(ByVal analito As String, ByVal ano As Long, _
                              Optional ByVal fonte As String = "") As String
    Dim r As String, p() As String
    r = ResolverEspec(analito, ano, fonte)
    If Left$(r, 2) = "OK" Then SituacaoEspec = "CADASTRADA": Exit Function
    If r = "NAO_CADASTRADA" Then SituacaoEspec = "NAO CADASTRADA": Exit Function
    p = Split(r, "|")
    SituacaoEspec = "INVALIDA: " & p(2)
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

Private Function ModeloSuportado(ByVal modelo As String) As Boolean
    Select Case UCase$(Trim$(modelo))
        Case MOD_ETP, MOD_VB, MOD_CVBIAS
            ModeloSuportado = True
    End Select
End Function

Public Function FontePadrao() As String
    Dim ws As Worksheet, v As String, i As Long
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    v = Trim$(CStr(ws.Range("B2").Value))
    If v = "" Then Exit Function

    For i = 7 To 40
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(v) Then
            If ModeloSuportado(CStr(ws.Cells(i, 2).Value)) Then FontePadrao = v
            Exit Function
        End If
    Next i
fim:
End Function

Public Function RigorPadrao() As String
    Dim ws As Worksheet, v As String
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    v = UCase$(Trim$(CStr(ws.Range("B3").Value)))
    If Left$(v, 3) = "OTI" Then RigorPadrao = "OTI": Exit Function
    If Left$(v, 3) = "DES" Then RigorPadrao = "DES": Exit Function
    If Left$(v, 3) = "MIN" Then RigorPadrao = "MIN"
fim:
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
    Dim ws As Worksheet, i As Long, modelo As String
    If Trim$(fonte) = "" Then Exit Function
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    For i = 7 To 40
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(fonte)) Then
            modelo = UCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If ModeloSuportado(modelo) Then ModeloDaFonte = modelo
            Exit Function
        End If
    Next i
fim:
End Function

' ---------------------------------------------------------------------------
' Grava uma especificacao. Chave logica: Ano + Fonte + Analito -- reenviar a
' mesma chave ATUALIZA, nao duplica, pelo mesmo principio do UpsertResultados.
Public Function GravarEspec(ByVal ano As Long, ByVal fonte As String, ByVal analito As String, _
                            ByVal etp As Variant, ByVal cvi As Variant, ByVal cvg As Variant, _
                            ByVal rigor As String, ByVal cvtp As Variant, ByVal biastp As Variant) As String
    Dim ws As Worksheet, dados As Variant, i As Long, lin As Long, modelo As String
    Dim metas As String, partes() As String

    fonte = Trim$(fonte)
    analito = Trim$(analito)
    If ano < 1900 Or ano > 2200 Then Err.Raise vbObjectError + 620, "mEspecificacoes.GravarEspec", "Ano invalido."
    If fonte = "" Then Err.Raise vbObjectError + 621, "mEspecificacoes.GravarEspec", "Cadastre e selecione uma fonte."
    If analito = "" Then Err.Raise vbObjectError + 622, "mEspecificacoes.GravarEspec", "Analito obrigatorio."

    modelo = ModeloDaFonte(fonte)
    If modelo = "" Then Err.Raise vbObjectError + 623, "mEspecificacoes.GravarEspec", "Fonte sem modelo suportado."
    If modelo = MOD_VB And NormRigor(rigor) = "" Then
        Err.Raise vbObjectError + 624, "mEspecificacoes.GravarEspec", "Selecione um rigor valido para a fonte VB."
    End If

    metas = MetasDaLinha(modelo, etp, cvi, cvg, rigor, cvtp, biastp)
    partes = Split(metas, "|")
    If Len(partes(0)) = 0 And Len(partes(1)) = 0 And Len(partes(2)) = 0 Then
        Err.Raise vbObjectError + 625, "mEspecificacoes.GravarEspec", "A especificacao nao contem entradas validas para o modelo."
    End If

    Set ws = ThisWorkbook.Sheets(ESPEC)
    dados = CarregarEspec()
    lin = 0
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If CStr(dados(i, ES_ANO)) = CStr(ano) _
               And UCase$(Trim$(CStr(dados(i, ES_FONTE)))) = UCase$(fonte) _
               And UCase$(Trim$(CStr(dados(i, ES_ANALITO)))) = UCase$(analito) Then
                lin = ESPEC_R0 + i - 1
                Exit For
            End If
        Next i
    End If

    Dim novoID As String, prot As Boolean
    If lin = 0 Then
        lin = UltimaLinhaEspec() + 1
        If lin < ESPEC_R0 Then lin = ESPEC_R0
        novoID = "ESP-" & Format$(lin - ESPEC_R0 + 1, "000000")
    Else
        novoID = CStr(ws.Cells(lin, ES_ID).Value)
    End If

    prot = ws.ProtectContents
    On Error GoTo falha
    If prot Then ws.Unprotect Password:=SENHA_ESPEC
    ws.Cells(lin, ES_ID).Value = novoID
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
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True

    InvalidarCacheEspec
    RegistrarLog "ESPEC_GRAVADA", novoID & " " & fonte & " " & ano & " " & analito
    GravarEspec = novoID
    Exit Function
falha:
    Dim numErro As Long, descErro As String
    numErro = Err.Number: descErro = Err.Description
    On Error Resume Next
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise numErro, "mEspecificacoes.GravarEspec", descErro
End Function

' ===========================================================================
' CAMADA DE SAIDA: Eng_Especificacoes
'
' Mesmo padrao do Eng_Saida. O motor ESCREVE aqui; Painel e Estatistica apenas
' REFERENCIAM. Nenhuma aba de interface chama o motor por UDF nem repete a
' formula -- a interface seleciona e formata, nunca calcula (ADR-019).
'
' O ANO DE CONTEXTO e o do resultado mais recente do lote em uso, e fica
' GRAVADO em B1 para o usuario ver contra qual ano a meta esta sendo aplicada.
' Meta aplicada em silencio, sem dizer de que ano e, e meta que ninguem audita.

Private Function EspecLegada(ByVal analito As String, ByRef fonte As String, ByRef etp As Variant) As Boolean
    Dim ws As Worksheet, i As Long, v As Variant
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(analito)) Then
            v = ws.Cells(i, 18).Value
            If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
                etp = CDbl(v)
                fonte = Trim$(CStr(ws.Cells(i, 17).Value))
                If fonte = "" Or fonte = "-" Then fonte = "LEGADO_ANALITOS"
                EspecLegada = True
            End If
            Exit Function
        End If
    Next i
fim:
End Function

Public Sub AtualizarEngEspec()
    Dim ws As Worksheet, c As Collection, i As Long
    Dim ano As Long, fonteFormal As String, fonteLegada As String
    Dim r As String, p() As String, etpLegado As Variant
    Dim buf() As Variant, n As Long, prot As Boolean

    Set ws = ThisWorkbook.Sheets(ENGE)
    fonteFormal = FontePadrao()
    ano = AnoDeContexto()
    Set c = ListaAnalitos()
    n = c.Count

    prot = ws.ProtectContents
    On Error GoTo falha
    If prot Then ws.Unprotect Password:=SENHA_ESPEC
    ws.Range(ws.Cells(ENGE_R0, 1), ws.Cells(ENGE_R0 + 39, 9)).ClearContents
    ws.Range("B1").Value = ano
    ws.Range("D1").Value = fonteFormal
    If n = 0 Then GoTo concluir

    ReDim buf(1 To n, 1 To 9)
    For i = 1 To n
        buf(i, 1) = c(i)
        r = ResolverEspec(CStr(c(i)), ano, fonteFormal)
        If fonteFormal <> "" And Left$(r, 2) = "OK" Then
            p = Split(r, "|")
            buf(i, 2) = p(4)
            buf(i, 3) = Val(p(5))
            buf(i, 4) = IIf(Len(p(1)) > 0, Val(p(1)), "")
            buf(i, 5) = IIf(Len(p(2)) > 0, Val(p(2)), "")
            buf(i, 6) = IIf(Len(p(3)) > 0, Val(p(3)), "")
            buf(i, 7) = p(6)
            buf(i, 8) = p(7)
            buf(i, 9) = "CADASTRADA"
        Else
            fonteLegada = "": etpLegado = Empty
            If EspecLegada(CStr(c(i)), fonteLegada, etpLegado) Then
                buf(i, 2) = fonteLegada
                buf(i, 6) = etpLegado
                buf(i, 9) = "FALLBACK LEGADO"
            ElseIf fonteFormal = "" Then
                buf(i, 9) = "FONTE NAO CONFIGURADA"
            ElseIf r = "NAO_CADASTRADA" Then
                buf(i, 2) = fonteFormal
                buf(i, 9) = "NAO CADASTRADA"
            Else
                p = Split(r, "|")
                buf(i, 2) = fonteFormal
                If UBound(p) >= 1 Then buf(i, 8) = p(1)
                buf(i, 9) = "FORMAL INVALIDA; SEM FALLBACK"
            End If
        End If
    Next i
    ws.Range(ws.Cells(ENGE_R0, 1), ws.Cells(ENGE_R0 + n - 1, 9)).Value = buf

concluir:
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    Exit Sub
falha:
    Dim numErro As Long, descErro As String
    numErro = Err.Number: descErro = Err.Description
    On Error Resume Next
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise numErro, "mEspecificacoes.AtualizarEngEspec", descErro
End Sub

' Ano contra o qual a meta e aplicada no Painel e na Estatistica.
'
' E o ano do resultado MAIS RECENTE do lote em uso -- nao Year(Date). O sistema
' precisa poder reabrir um lote de 2025 daqui a cinco anos e reproduzir a meta
' vigente naquela epoca; usar o ano corrente destruiria isso em silencio.
'
' Cfg_Especificacoes!B4 > 0 forca um ano, para conferencia e reprocessamento.
Public Function AnoDeContexto() As Long
    Dim ws As Worksheet, forcado As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    If Not ws Is Nothing Then forcado = CLng(Val(CStr(ws.Range("B4").Value)))
    On Error GoTo 0
    If forcado > 0 Then AnoDeContexto = forcado: Exit Function

    Dim dados As Variant, i As Long, mx As Date, lote As String
    lote = LoteAtivoCore()
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If IsDate(dados(i, COL_DATA)) Then
                If lote = "" Or NucleoLote(CStr(dados(i, COL_LOTE))) = lote Then
                    If CDate(dados(i, COL_DATA)) > mx Then mx = CDate(dados(i, COL_DATA))
                End If
            End If
        Next i
    End If
    If mx = 0 Then AnoDeContexto = Year(Date) Else AnoDeContexto = Year(mx)
End Function

' Ponto de entrada do botao da aba Analitos.
'
' JA SUMIU UMA VEZ. Reinstalar este modulo por script apaga o que foi
' anexado ao fim dele -- e o botao fica apontando para um nome que nao
' existe mais, com cara de codigo orfao. Ele mora AQUI, na fonte
' versionada, e nao so no arquivo: foi a verificacao 1.1 (modulo
' identico a fonte) que acusou a divergencia quando eu restaurei so no
' artefato.
Public Sub AbrirFormEspecificacoes()
    frmEspecificacoes.Show
End Sub
