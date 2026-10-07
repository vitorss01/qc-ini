Option Explicit
' ===== INCERTEZA DE MEDICAO (ADR-064) =====
'
' Modelo: Nordtest TR 537, top-down por DESEMPENHO (ADR-067, 05/10/2026, decisao do
' laboratorio: so CIQ + CEQ, sem certificado de calibrador; Magnusson 2012, Cui 2017,
' Hann 2017, Martinello 2020, Michaelis 2026).
'     uc = raiz( u(Rw)^2 + u(bias)^2 )       U = 2 x uc  (~95%, bilateral)
' u(Rw): imprecisao intermediaria de LONGO PRAZO do CIQ (janela de 12 meses), por
' analito x nivel x equipamento, com os lotes de controle SEPARADOS e agrupados
' pelos graus de liberdade:
'     u(Rw)% = raiz( soma((n_i - 1) x CV_i^2) / soma(n_i - 1) ),  so lotes com n_i >= 20
' Um DP unico sobre lotes diferentes inflaria u(Rw) com a diferenca entre os alvos
' dos materiais; a media simples dos CVs so vale com n iguais (Coskun 2022).
' u(bias): do CEQ (mCEQ.ViesEQ "UBIAS") = raiz(RMS dos bias^2 + u(Cref)^2), com
' u(Cref) = DP do grupo / raiz(n laboratorios). Exige >= 6 amostras de >= 2 rodadas
' (preferencial 10) e u(Cref) informado. Sem u(bias) a incerteza NAO e estimada:
' u(Rw) sozinha nao e a incerteza (Panteghini 2022) -- u(Rw) aparece, U nao.
' A triagem do vies (mCEQ "TRIAGEM") continua: vies relevante se investiga.
'
' Validade (regras INTERNAS documentadas no ADR-064; a norma nao fixa n):
'   dias cobertos pelos lotes elegiveis >= 180 e gl >= 100  -> valida
'   90 a 179 dias ou gl 30 a 99                           -> PROVISORIA
'   abaixo disso                                          -> nao exibida
' Alertas que nao bloqueiam: lotes curtos (mediana < 90 dias -- deslocamentos na
' troca de lote nao entram em u(Rw)), CVmax/CVmin > 2, excluidos > 5%.
'
' Mesma populacao do motor: tblCQ_Final (mDados.CarregarDB), equipamento do
' Painel, janela de datas e exclusoes da aba Estatistica. So resultado ATIVO
' (participa da estatistica) entra em u(Rw); o inativado so e CONTADO, para o
' alerta de excluidos (exclusao sem erro documentado reduz u(Rw) artificialmente).

' ADR-071: o MESMO estimador (CV pooled dos lotes n >= 20 pelos graus de liberdade) e o denominador
' do Sigma (Estatistica!L) numa janela de N meses (Sigma_Ini..MU_Fim, padrao 6), ao lado da janela de
' 12 meses da incerteza. Cache por JANELA (carimbo -> agregado): na mesma linha da tabela ~11 chamadas
' usam 12 m e 1 usa N m; com um slot so, a alternancia refazia a varredura da tblCQ_Final inteira a
' cada chamada (I09 < 10 s e qa_desempenho estourariam).

Private mSnapMU As Variant
Private mAggMU As Object          ' "ANALITO|NIVEL" -> Dictionary(lote -> Array(n, soma, somaQ, dMin, dMax))
Private mContMU As Object         ' "ANALITO|NIVEL" -> Array(total na janela, inativos)
Private mCarimboMU As String      ' carimbo da janela em mAggMU/mContMU
Private mCacheMU As Object        ' carimbo -> Array(agregado, contagem)  (ADR-071)
Private Const MU_CACHE_MAX As Long = 6

Public Const MU_N_LOTE_MIN As Long = 20       ' lote entra no agrupamento com n >= 20
Public Const MU_DIAS_VALIDA As Long = 180
Public Const MU_DIAS_MIN As Long = 90
Public Const MU_GL_VALIDA As Long = 100
Public Const MU_GL_MIN As Long = 30
Public Const MU_K As Double = 2               ' fator de abrangencia (~95%)
Public Const MU_LOTE_CURTO As Long = 90       ' duracao mediana do lote, dias
Public Const MU_RAZAO_CV As Double = 2        ' CVmax / CVmin entre lotes
Public Const MU_PEXCL As Double = 5           ' % de resultados inativados na janela
' Meta pela variacao biologica (EFLM, modelo 2 de Milao): u (k = 1) <= f x CVI
Public Const MU_F_OTIMO As Double = 0.25
Public Const MU_F_DESEJAVEL As Double = 0.5
Public Const MU_F_MINIMO As Double = 0.75
Public Const MU_CEQ_PREF As Long = 10        ' amostras de CEQ preferenciais (Nordtest TR 537)


Public Sub InvalidarIncerteza()
    mSnapMU = Empty
    Set mAggMU = Nothing
    Set mContMU = Nothing
    Set mCacheMU = Nothing
    mCarimboMU = ""
End Sub

' Depois de um refresh (ATUALIZAR DADOS) ou da troca de equipamento nenhum ARGUMENTO das
' formulas de incerteza muda -- o Excel nao as recalcularia. Marca a faixa (nome MU_Faixa,
' criado pelo instalador do ADR-064) como suja e recalcula. ADR-071: tambem o Sigma
' (Sigma_Faixa = Estatistica!L, CV pooled de N meses) -- senao L ficaria com o valor velho
' (o Painel I7:I9 le L e acompanha).
Public Sub RecalcularIncerteza()
    On Error Resume Next
    InvalidarIncerteza
    mEQA.InvalidarNLabs                  ' n de laboratorios digitado depois da consolidacao
    ThisWorkbook.Names("MU_Faixa").RefersToRange.Dirty
    ThisWorkbook.Names("Sigma_Faixa").RefersToRange.Dirty
    Application.Calculate
End Sub

Private Function ComoDataMU(ByVal v As Variant) As Double
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If IsDate(v) Then
        ComoDataMU = Int(CDbl(CDate(v)))
    ElseIf IsNumeric(v) Then
        ComoDataMU = Int(CDbl(v))
    End If
End Function

' Exclusoes: intervalo de 2 colunas (inicio, fim) por linha -- o formato da aba
' Estatistica (Estat_Exclusoes). Pares incompletos ou invertidos nao valem.
Private Function LerExclusoesMU(ByVal exclusoes As Variant, ByRef ex() As Double) As Long
    Dim v As Variant, i As Long, n As Long, a As Double, b As Double
    ReDim ex(1 To 1, 1 To 2)
    On Error GoTo fim
    If IsObject(exclusoes) Then v = exclusoes.Value Else v = exclusoes
    If Not IsArray(v) Then GoTo fim
    If UBound(v, 2) - LBound(v, 2) + 1 < 2 Then GoTo fim
    ReDim ex(1 To UBound(v, 1) - LBound(v, 1) + 1, 1 To 2)
    For i = LBound(v, 1) To UBound(v, 1)
        a = ComoDataMU(v(i, LBound(v, 2)))
        b = ComoDataMU(v(i, LBound(v, 2) + 1))
        If a > 0 And b > 0 And b >= a Then
            n = n + 1
            ex(n, 1) = a
            ex(n, 2) = b
        End If
    Next i
fim:
    LerExclusoesMU = n
End Function

Private Sub AgregarMU(ByVal dtIni As Double, ByVal dtFim As Double, ByRef ex() As Double, ByVal nEx As Long)
    Dim dados As Variant, i As Long, j As Long, d As Double, v As Double
    Dim eq As String, c As String, k As String, lote As String, reg As Variant, porLote As Object, ct As Variant
    Dim sa As String, par As Variant

    eq = mEstatistica.EquipFiltro()
    c = CStr(dtIni) & "|" & CStr(dtFim) & "|" & eq & "|" & nEx
    For j = 1 To nEx
        c = c & "|" & CStr(ex(j, 1)) & "-" & CStr(ex(j, 2))
    Next j
    If Not mAggMU Is Nothing Then
        If c = mCarimboMU Then Exit Sub
    End If
    ' ADR-071: janela ja agregada nesta sessao (12 m da incerteza, N m do Sigma)?
    If mCacheMU Is Nothing Then Set mCacheMU = CreateObject("Scripting.Dictionary")
    If mCacheMU.Exists(c) Then
        par = mCacheMU.Item(c)
        Set mAggMU = par(0)
        Set mContMU = par(1)
        mCarimboMU = c
        Exit Sub
    End If
    If mCacheMU.Count >= MU_CACHE_MAX Then mCacheMU.RemoveAll
    Set mAggMU = CreateObject("Scripting.Dictionary")
    mAggMU.CompareMode = 1
    Set mContMU = CreateObject("Scripting.Dictionary")
    mContMU.CompareMode = 1
    mCarimboMU = c
    mCacheMU.Add c, Array(mAggMU, mContMU)        ' os objetos sao preenchidos abaixo (referencia)
    If IsEmpty(mSnapMU) Then mSnapMU = mDados.CarregarDB()
    If IsEmpty(mSnapMU) Then Exit Sub
    dados = mSnapMU

    For i = 1 To UBound(dados, 1)
        If Trim$(CStr(dados(i, COL_ANALITO))) = "" Then GoTo proxima
        If Len(eq) > 0 Then
            If UCase$(Trim$(CStr(dados(i, COL_EQUIP)))) <> eq Then GoTo proxima
        End If
        If Not IsDate(dados(i, COL_DATA)) Then GoTo proxima
        d = Int(CDbl(CDate(dados(i, COL_DATA))))
        If dtIni > 0 Then
            If d < dtIni Then GoTo proxima
        End If
        If dtFim > 0 Then
            If d > dtFim Then GoTo proxima
        End If
        For j = 1 To nEx
            If d >= ex(j, 1) And d <= ex(j, 2) Then GoTo proxima
        Next j
        k = UCase$(Trim$(CStr(dados(i, COL_ANALITO)))) & "|" & Trim$(CStr(dados(i, COL_NIVEL)))
        ' alerta de excluidos: so a INATIVACAO do analista (STATUS_ANALITICO = INATIVADO) sobre os
        ' resultados reais (ATIVO + INATIVADO). Sem valor, retransmissao, conflito e manual incompleto
        ' sao exclusoes tecnicas do Power Query, nao escolha de quem analisa (auditoria 04/10/2026)
        sa = UCase$(Trim$(CStr(dados(i, COL_STATUS_AN))))
        If sa = "ATIVO" Or sa = "INATIVADO" Then
            If Not mContMU.Exists(k) Then mContMU.Add k, Array(0#, 0#)
            ct = mContMU(k)
            ct(0) = ct(0) + 1
            If sa = "INATIVADO" Then ct(1) = ct(1) + 1
            mContMU(k) = ct
        End If
        If CStr(dados(i, COL_STATUS)) <> ST_ATIVO Then GoTo proxima
        If Not IsNumeric(dados(i, COL_RESULT)) Then GoTo proxima
        lote = Trim$(CStr(dados(i, COL_LOTE)))
        If Len(lote) = 0 Then GoTo proxima

        v = CDbl(dados(i, COL_RESULT))
        If Not mAggMU.Exists(k) Then
            Set porLote = CreateObject("Scripting.Dictionary")
            porLote.CompareMode = 1
            mAggMU.Add k, porLote
        End If
        Set porLote = mAggMU(k)
        If porLote.Exists(lote) Then
            reg = porLote(lote)
            reg(0) = reg(0) + 1
            reg(1) = reg(1) + v
            reg(2) = reg(2) + v * v
            If d < reg(3) Then reg(3) = d
            If d > reg(4) Then reg(4) = d
            porLote(lote) = reg
        Else
            porLote.Add lote, Array(1#, v, v * v, d, d)
        End If
proxima:
    Next i
End Sub

' u(Rw) pelo CIQ, lotes separados e agrupados (so lotes com n >= MU_N_LOTE_MIN).
'   "CV"        u(Rw) % agrupada                  "GL"       soma(n_i - 1)
'   "N"         resultados nos lotes elegiveis    "LOTES"    lotes elegiveis
'   "FORA"      lotes com n < 20 (fora)          "DIAS"     ultimo - primeiro resultado dos lotes elegiveis
'   "MEDIA"     media ponderada dos lotes (x_L)   "DPABS"    u(Rw) absoluta (DP agrupado)
'   "RAZAO"     CVmax / CVmin entre os lotes      "DURMED"   duracao mediana dos lotes, dias
'   "PEXCL"     % de inativados na janela         "VALIDADE" VALIDA | PROVISORIA | INSUFICIENTE
Public Function IncertezaCIQ(ByVal analito As String, ByVal nivel As Variant, ByVal metrica As String, _
                             ByVal dtIni As Variant, ByVal dtFim As Variant, _
                             Optional ByVal exclusoes As Variant) As Variant
    Dim ex() As Double, nEx As Long, k As String, porLote As Object, lt As Variant, reg As Variant
    Dim n As Double, media As Double, num As Double, s As Double, cvj As Double, md As String
    Dim somaGL As Double, somaQcv As Double, somaQs As Double, nTot As Double, somaX As Double
    Dim nLotes As Long, nFora As Long, dMin As Double, dMax As Double
    Dim cvMin As Double, cvMax As Double, dur() As Double, nDur As Long, ct As Variant

    IncertezaCIQ = ""
    If Len(Trim$(analito)) = 0 Then Exit Function
    On Error GoTo falhou
    md = UCase$(Trim$(metrica))
    If IsMissing(exclusoes) Then
        nEx = 0
        ReDim ex(1 To 1, 1 To 2)
    Else
        nEx = LerExclusoesMU(exclusoes, ex)
    End If
    AgregarMU ComoDataMU(dtIni), ComoDataMU(dtFim), ex, nEx
    k = UCase$(Trim$(analito)) & "|" & Trim$(CStr(nivel))

    If md = "PEXCL" Then
        If mContMU.Exists(k) Then
            ct = mContMU(k)
            If ct(0) > 0 Then IncertezaCIQ = ct(1) / ct(0) * 100#
        End If
        Exit Function
    End If
    If Not mAggMU.Exists(k) Then
        Select Case md
            Case "N", "LOTES", "GL", "FORA": IncertezaCIQ = 0
            Case "VALIDADE": IncertezaCIQ = "INSUFICIENTE"
        End Select
        Exit Function
    End If

    Set porLote = mAggMU(k)
    ReDim dur(1 To porLote.Count)
    cvMin = 1E+300
    For Each lt In porLote.Keys
        reg = porLote(lt)
        n = reg(0)
        If n >= MU_N_LOTE_MIN Then media = reg(1) / n Else media = 0
        If n < MU_N_LOTE_MIN Or media = 0 Then
            ' n < 20 ou media zero (nivel todo em 0,00: CV indefinido) -- fora do agrupamento.
            ' Somar os gl desse lote sem somar CV encolheria u(Rw) (auditoria 04/10/2026)
            nFora = nFora + 1
        Else
            num = reg(2) - n * media * media
            If num < 0 Then num = 0
            s = Sqr(num / (n - 1))
            cvj = s / Abs(media) * 100#
            somaQcv = somaQcv + (n - 1) * cvj * cvj
            If cvj < cvMin Then cvMin = cvj
            If cvj > cvMax Then cvMax = cvj
            somaQs = somaQs + (n - 1) * s * s
            somaGL = somaGL + (n - 1)
            nTot = nTot + n
            somaX = somaX + reg(1)
            nLotes = nLotes + 1
            If nLotes = 1 Then
                dMin = reg(3): dMax = reg(4)
            Else
                If reg(3) < dMin Then dMin = reg(3)
                If reg(4) > dMax Then dMax = reg(4)
            End If
            nDur = nDur + 1
            dur(nDur) = reg(4) - reg(3) + 1
        End If
    Next lt

    Select Case md
        Case "CV":     If somaGL > 0 Then IncertezaCIQ = Sqr(somaQcv / somaGL)
        Case "DPABS":  If somaGL > 0 Then IncertezaCIQ = Sqr(somaQs / somaGL)
        Case "MEDIA":  If nTot > 0 Then IncertezaCIQ = somaX / nTot
        Case "N":      IncertezaCIQ = nTot
        Case "GL":     IncertezaCIQ = somaGL
        Case "LOTES":  IncertezaCIQ = nLotes
        Case "FORA":   IncertezaCIQ = nFora
        Case "DIAS":   If nLotes > 0 Then IncertezaCIQ = dMax - dMin
        Case "RAZAO":  If nLotes >= 2 And cvMin > 0 Then IncertezaCIQ = cvMax / cvMin
        Case "DURMED": If nDur > 0 Then IncertezaCIQ = Mediana(dur, nDur)
        Case "VALIDADE"
            If nLotes = 0 Then
                IncertezaCIQ = "INSUFICIENTE"
            ElseIf (dMax - dMin) < MU_DIAS_MIN Or somaGL < MU_GL_MIN Then
                IncertezaCIQ = "INSUFICIENTE"
            ElseIf (dMax - dMin) < MU_DIAS_VALIDA Or somaGL < MU_GL_VALIDA Then
                IncertezaCIQ = "PROVISORIA"
            Else
                IncertezaCIQ = "VALIDA"
            End If
        Case Else:     IncertezaCIQ = CVErr(xlErrValue)
    End Select
    Exit Function
falhou:
    IncertezaCIQ = CVErr(xlErrValue)
End Function

Private Function Mediana(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, j As Long, t As Double, a() As Double
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = v(i)
    Next i
    For i = 2 To n                       ' insercao: poucos lotes
        t = a(i)
        j = i - 1
        Do While j >= 1
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
    If n Mod 2 = 1 Then Mediana = a((n + 1) \ 2) Else Mediana = (a(n \ 2) + a(n \ 2 + 1)) / 2
End Function

' CVI para a meta: SO o cadastrado (EFLM). Sem CVI, sem meta -- e NUNCA do TEa/ETp CLIA.
' A aproximacao "CVI = 2 x CVTp" saiu (auditoria 04/10/2026): o CVTp de VB da pasta usa o nivel
' do analito (otimo 0,25, desejavel 0,50, minimo 0,75 x CVI), e ele mesmo vem do CVI -- sem CVI
' nao ha CVTp de VB. fonte e cvtp ficam na assinatura so por compatibilidade das formulas.
Public Function CVIMeta(ByVal cvi As Variant, ByVal fonte As Variant, ByVal cvtp As Variant) As Variant
    CVIMeta = ""
    If EhNumero(cvi) Then
        If CDbl(cvi) > 0 Then CVIMeta = CDbl(cvi)
    End If
End Function

' Numero de verdade (nao texto com cara de numero): IsNumeric("1,2") e True no VBA em portugues,
' e as formulas tratam texto como zero (N()) -- os dois lados tem de concordar (auditoria 04/10/2026).
Private Function EhNumero(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal: EhNumero = True
    End Select
End Function

' Classe da incerteza (para cor e leitura rapida): compara uc (k = 1) com a MAU.
Public Function ClasseIncerteza(ByVal uc As Variant, ByVal cviMeta As Variant, ByVal validade As Variant) As String
    Dim r As Double
    If UCase$(CStr(validade)) = "INSUFICIENTE" Or Not IsNumeric(uc) Or Len(Trim$(CStr(uc))) = 0 Then
        ClasseIncerteza = "Insuficiente"
        Exit Function
    End If
    If Not IsNumeric(cviMeta) Or Len(Trim$(CStr(cviMeta))) = 0 Then
        ClasseIncerteza = "Sem meta"
        Exit Function
    End If
    If CDbl(cviMeta) <= 0 Then
        ClasseIncerteza = "Sem meta"
        Exit Function
    End If
    r = CDbl(uc) / CDbl(cviMeta)
    If r <= MU_F_OTIMO Then
        ClasseIncerteza = "Ótimo"
    ElseIf r <= MU_F_DESEJAVEL Then
        ClasseIncerteza = "Desejável"
    ElseIf r <= MU_F_MINIMO Then
        ClasseIncerteza = "Mínimo"
    Else
        ClasseIncerteza = "Não atende"
    End If
End Function

' Situacao da estimativa. validade = do CIQ; ubias/sitCEQ/nCEQ = do CEQ (mCEQ.ViesEQ
' "UBIAS", "SITUACAO", "N"). Sem u(bias) o U nao e calculado e o motivo aparece aqui.
Public Function SituacaoIncerteza(ByVal validade As Variant, ByVal ubias As Variant, _
                                  Optional ByVal sitCEQ As Variant = "", Optional ByVal nCEQ As Variant = "") As String
    Dim s As String
    Select Case UCase$(CStr(validade))
        Case "VALIDA":     s = "Válida"
        Case "PROVISORIA": s = "PROVISÓRIA (CIQ < 180 dias ou gl < 100)"
        Case Else:         SituacaoIncerteza = "Não estimada: CIQ < 90 dias ou gl < 30": Exit Function
    End Select
    If Not EhNumero(ubias) Then
        If IsError(sitCEQ) Then sitCEQ = ""
        If Len(Trim$(CStr(sitCEQ))) = 0 Then sitCEQ = "sem u(bias) do CEQ"
        SituacaoIncerteza = "U não estimada: " & CStr(sitCEQ)
        Exit Function
    End If
    If EhNumero(nCEQ) Then
        If CDbl(nCEQ) < MU_CEQ_PREF Then s = s & " · CEQ com " & CStr(nCEQ) & " amostras (preferível >= " & MU_CEQ_PREF & ")"
    End If
    SituacaoIncerteza = s
End Function

' Alertas que nao bloqueiam o valor.
Public Function AlertasIncerteza(ByVal analito As String, ByVal nivel As Variant, _
                                 ByVal dtIni As Variant, ByVal dtFim As Variant, _
                                 ByVal urw As Variant, ByVal cviMeta As Variant, ByVal ubias As Variant, _
                                 Optional ByVal exclusoes As Variant) As String
    Dim s As String, v As Variant
    If Len(Trim$(analito)) = 0 Then Exit Function
    If IsMissing(exclusoes) Then
        v = IncertezaCIQ(analito, nivel, "DURMED", dtIni, dtFim)
    Else
        v = IncertezaCIQ(analito, nivel, "DURMED", dtIni, dtFim, exclusoes)
    End If
    If IsNumeric(v) And Len(CStr(v)) > 0 Then
        If CDbl(v) < MU_LOTE_CURTO Then s = s & "lotes curtos (mediana " & Format$(v, "0") & " d); "
    End If
    If IsMissing(exclusoes) Then
        v = IncertezaCIQ(analito, nivel, "RAZAO", dtIni, dtFim)
    Else
        v = IncertezaCIQ(analito, nivel, "RAZAO", dtIni, dtFim, exclusoes)
    End If
    If IsNumeric(v) And Len(CStr(v)) > 0 Then
        If CDbl(v) > MU_RAZAO_CV Then s = s & "CV heterogêneo entre lotes (" & Format$(v, "0.0") & "x); "
    End If
    If IsMissing(exclusoes) Then
        v = IncertezaCIQ(analito, nivel, "PEXCL", dtIni, dtFim)
    Else
        v = IncertezaCIQ(analito, nivel, "PEXCL", dtIni, dtFim, exclusoes)
    End If
    If IsNumeric(v) And Len(CStr(v)) > 0 Then
        If CDbl(v) > MU_PEXCL Then s = s & Format$(v, "0.0") & "% inativados; "
    End If
    ' componente dominante: orienta a acao (precisao -> CIQ/manutencao; vies -> calibracao/rastreabilidade)
    If EhNumero(urw) And EhNumero(ubias) Then
        If CDbl(ubias) > 0 And CDbl(urw) > 0 Then
            If CDbl(ubias) > 2# * CDbl(urw) Then
                s = s & "u(bias) domina (" & Format$(CDbl(ubias) / CDbl(urw), "0.0") & "× u(Rw)): investigar vies/calibração; "
            ElseIf CDbl(urw) > 2# * CDbl(ubias) Then
                s = s & "u(Rw) domina: precisão é o componente principal; "
            End If
        End If
    End If
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    AlertasIncerteza = s
End Function

' Nota do bloco de incerteza no Painel. As datas saem do VBA (Format com codigo fixo):
' TEXTO() da planilha depende do idioma da instalacao ("aaaa" x "yyyy").
Public Function NotaIncerteza(ByVal ini As Variant, ByVal fim As Variant) As String
    Dim j As String, a As Double, b As Double
    a = ComoDataMU(ini)
    b = ComoDataMU(fim)
    If a > 0 And b > 0 Then
        j = "Janela " & Format$(CDate(a), "dd/mm/yyyy") & " a " & Format$(CDate(b), "dd/mm/yyyy") & " (12 meses)"
    Else
        j = "Sem dados de CIQ na janela"
    End If
    NotaIncerteza = j & "  ·  Nordtest: uc = raiz(u(Rw)² + u(bias)²)  ·  u(Rw) = CIQ de longo prazo, lotes (n >= 20) agrupados  ·  " & _
                    "u(bias) = CEQ: raiz(RMS viés² + u(Cref)²), >= 6 amostras de >= 2 rodadas  ·  U = 2 × uc (~95%)  ·  meta: u <= 0,50 × CVI (EFLM)"
End Function
