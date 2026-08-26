Attribute VB_Name = "mEstatPeriodo"
Option Explicit
' ===== MOTOR DE ESTATISTICA POR PERIODO =====
'
' Responde n, Media, DP, CV%, Bias% e ET% de um analito+nivel dentro de uma
' JANELA DE VISAO, descontando uma JANELA DE EXCLUSAO.
'
' POR QUE AS DATAS SAO ARGUMENTO, E NAO LEITURA DE CELULA NOMEADA
'
' Seria mais curto escrever =EstatPeriodo("Glicose";1;"CV") e a funcao ler as
' celulas de filtro por dentro. Seria tambem ERRADO: o Excel so recalcula uma
' UDF quando muda alguma celula que ela recebeu COMO ARGUMENTO. Lendo por
' dentro, trocar o periodo nao recalcularia nada -- a tela mostraria o numero do
' periodo anterior, com cara de certo. Application.Volatile "resolveria"
' recalculando a pasta inteira a cada tecla, o que num banco de 15.000 linhas
' trava a planilha.
'
' Entao as janelas entram por argumento. A formula fica longa; o numero fica
' certo.
'
' CACHE
'
' Sao 62 linhas x 6 metricas = 372 chamadas por recalculo. Sem cache, cada uma
' varreria o banco inteiro. O cache guarda o AGREGADO por analito+nivel e e
' invalidado quando muda a janela, o lote ou o tamanho do banco.
'
' N<2 DEVOLVE VAZIO, NUNCA ZERO
'
' DP zero seria lido como "dispersao perfeita" e CV zero como "imprecisao zero".
' Uma corrida unica nao tem dispersao estimavel -- e dizer isso e diferente de
' dizer que ela e zero.
'
' BIAS: DE ONDE VEM O ALVO
'
' O bias correto para avaliacao de METODO vem de ensaio de proficiencia (EQC),
' e a coluna antiga da aba lia EQC_Dados. Nao ha nenhum registro de EQC hoje,
' entao aquela coluna vinha vazia e levava ET% e Sigma vazios junto.
' Este motor usa o ALVO DO LOTE (LotesStore), que e o unico alvo com dado. Isso
' e bias contra o valor designado do controle -- util para acompanhar deriva,
' e NAO substitui o EQC para avaliacao de exatidao do metodo. A aba rotula a
' coluna com a origem para nao haver confusao.

Public Const EP_BANCO As String = "DB_Resultados"
Public Const EP_R0 As Long = 4

' colunas do banco
Private Const EP_C_RUN As Long = 1
Private Const EP_C_DATA As Long = 2
Private Const EP_C_NIVEL As Long = 3
Private Const EP_C_LOTE As Long = 4
Private Const EP_C_ANALITO As Long = 5
Private Const EP_C_VALOR As Long = 6
Private Const EP_C_STATUS As Long = 7

Private mAgg As Object          ' chave "ANALITO|NIVEL" -> Array(n, soma, somaQ)
Private mCarimbo As String

' ---------------------------------------------------------------------------
' Serializa a janela de forma INVARIANTE DE LOCALIDADE.
' CStr de data numa maquina pt-BR devolve "05/01/2026"; noutra devolve outra
' coisa. O carimbo do cache precisa ser estavel, entao usa o serial numerico.
Private Function Carimbo(ByVal dtIni As Double, ByVal dtFim As Double, _
                         ByRef ex() As Double, ByVal nEx As Long, _
                         ByVal lote As String, ByVal ultima As Long) As String
    Dim s As String, i As Long
    s = Trim$(Str$(dtIni)) & "|" & Trim$(Str$(dtFim))
    For i = 1 To nEx
        s = s & "|" & Trim$(Str$(ex(i, 1))) & "-" & Trim$(Str$(ex(i, 2)))
    Next i
    Carimbo = s & "|" & UCase$(Trim$(lote)) & "|" & Trim$(Str$(ultima))
End Function

' ---------------------------------------------------------------------------
' Le o bloco de exclusoes (intervalo Nx2: inicio na coluna 1, fim na coluna 2).
'
' UM ARGUMENTO, NAO DEZ.
'
' Acrescentar exIni2..exFim5 a assinatura resolveria hoje e cobraria caro
' depois: dez parametros para ler, uma formula ilegivel em 372 celulas, e uma
' sexta exclusao exigindo mexer no codigo e reescrever a aba inteira. Um
' intervalo e um argumento so -- e o Excel continua enxergando a dependencia,
' que e o que faz a aba recalcular quando o usuario digita uma data.
'
' Par incompleto (uma ponta so) e IGNORADO, nao adivinhado: exclusao pela
' metade descartaria historico por um preenchimento inacabado.
' Par invertido (fim antes do inicio) tambem e ignorado.
Private Function LerExclusoes(ByVal exclusoes As Variant, ByRef ex() As Double) As Long
    Dim v As Variant, i As Long, n As Long, a As Double, b As Double
    ReDim ex(1 To 50, 1 To 2)
    n = 0
    If IsMissing(exclusoes) Then LerExclusoes = 0: Exit Function
    If IsEmpty(exclusoes) Then LerExclusoes = 0: Exit Function

    On Error GoTo fim
    v = exclusoes
    If Not IsArray(v) Then
        ' celula unica nao forma par: nao ha o que excluir
        LerExclusoes = 0
        Exit Function
    End If

    For i = LBound(v, 1) To UBound(v, 1)
        If n >= 50 Then Exit For
        a = ComoData(v(i, LBound(v, 2)))
        b = ComoData(v(i, LBound(v, 2) + 1))
        If a > 0 And b > 0 Then
            If b >= a Then
                n = n + 1
                ex(n, 1) = a
                ex(n, 2) = b
            End If
        End If
    Next i
fim:
    LerExclusoes = n
End Function

Private Function UltimaLinhaEP() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(EP_BANCO)
    UltimaLinhaEP = ws.Cells(ws.Rows.Count, EP_C_RUN).End(xlUp).Row
End Function

' Converte o que vier da celula em serial de data. Celula vazia vira 0, e 0
' significa "sem limite deste lado" -- nao "01/01/1900".
'
' TRUNCA PARA O DIA, e isso NAO e detalhe.
'
' Um filtro de periodo de CQ e por DIA, nunca por instante. Se a data levar
' qualquer fracao de tempo -- e leva com facilidade: um Now() gravado, um
' Get-Date de script que nao zerou os milissegundos, uma importacao com hora --
' a comparacao vira instante contra instante e o resultado erra por UM registro,
' calado.
'
' Foi exatamente o que aconteceu no teste desta aba: a janela 05/01 a 09/01
' devolveu n=4 onde havia 5 registros, porque o limite era 05/01 00:00:00,050 e
' o registro das 00:00:00,000 ficou "antes do inicio". Pelo mesmo motivo o aviso
' de "exclusao cobre todo o periodo" nao disparava. Um erro de 50 milissegundos
' descartando uma corrida inteira de um calculo de imprecisao.
'
' Int() sobre o serial descarta a fracao e deixa a comparacao no dia.
Private Function ComoData(ByVal v As Variant) As Double
    Dim d As Double
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If IsDate(v) Then
        d = CDbl(CDate(v))
    ElseIf IsNumeric(v) Then
        d = CDbl(v)
    Else
        Exit Function
    End If
    ComoData = Int(d)
End Function

' ---------------------------------------------------------------------------
' Varre o banco UMA vez e agrega por analito+nivel.
Private Sub Agregar(ByVal dtIni As Double, ByVal dtFim As Double, _
                    ByRef ex() As Double, ByVal nEx As Long, _
                    ByVal lote As String)
    Dim ws As Worksheet, ult As Long, dados As Variant, i As Long, j As Long
    Dim c As String, k As String, d As Double, v As Double
    Dim reg As Variant

    ult = UltimaLinhaEP()
    c = Carimbo(dtIni, dtFim, ex, nEx, lote, ult)
    If Not mAgg Is Nothing Then
        If mCarimbo = c Then Exit Sub
    End If

    Set mAgg = CreateObject("Scripting.Dictionary")
    mAgg.CompareMode = 1
    mCarimbo = c
    If ult < EP_R0 Then Exit Sub

    Set ws = ThisWorkbook.Sheets(EP_BANCO)
    dados = ws.Range(ws.Cells(EP_R0, EP_C_RUN), ws.Cells(ult, EP_C_STATUS)).Value

    For i = 1 To UBound(dados, 1)
        If Trim$(CStr(dados(i, EP_C_ANALITO))) = "" Then GoTo proxima
        If Trim$(CStr(dados(i, EP_C_STATUS))) <> "Ativo" Then GoTo proxima
        If Not IsDate(dados(i, EP_C_DATA)) Then GoTo proxima
        If Not IsNumeric(dados(i, EP_C_VALOR)) Then GoTo proxima

        ' o registro tambem entra pelo DIA: a hora gravada no banco nao pode
        ' decidir se a corrida pertence ao periodo
        d = Int(CDbl(CDate(dados(i, EP_C_DATA))))

        ' janela de VISAO (0 = sem limite deste lado)
        If dtIni > 0 Then If d < dtIni Then GoTo proxima
        If dtFim > 0 Then If d > dtFim Then GoTo proxima

        ' janelas de EXCLUSAO: cair em QUALQUER uma tira o registro do calculo.
        ' Pares incompletos ou invertidos ja foram descartados em LerExclusoes.
        For j = 1 To nEx
            If d >= ex(j, 1) And d <= ex(j, 2) Then GoTo proxima
        Next j

        ' lote pelo NUCLEO (NucleoLote vive em mEntrada: uma conta, um lugar)
        If Len(Trim$(lote)) > 0 Then
            If NucleoLote(CStr(dados(i, EP_C_LOTE))) <> Trim$(lote) Then GoTo proxima
        End If

        v = CDbl(dados(i, EP_C_VALOR))
        k = UCase$(Trim$(CStr(dados(i, EP_C_ANALITO)))) & "|" & Trim$(CStr(dados(i, EP_C_NIVEL)))
        If mAgg.Exists(k) Then
            reg = mAgg(k)
            reg(0) = reg(0) + 1
            reg(1) = reg(1) + v
            reg(2) = reg(2) + v * v
            mAgg(k) = reg
        Else
            mAgg.Add k, Array(1#, v, v * v)
        End If
proxima:
    Next i
End Sub

' ---------------------------------------------------------------------------
' UDF principal.
'
' metrica: "N" | "MEDIA" | "DP" | "CV" | "BIAS" | "ET"
'
' Uso na planilha (as janelas ENTRAM como argumento para o Excel saber
' recalcular):
'   =EstatPeriodo($A7;$B7;"CV";Data_Inicio_Efetiva;Data_Fim_Efetiva;
'                 Data_Inicio_Exclusao;Data_Fim_Exclusao;$B$4)
Public Function EstatPeriodo(ByVal analito As String, ByVal nivel As Variant, _
                             ByVal metrica As String, _
                             ByVal dtIni As Variant, ByVal dtFim As Variant, _
                             ByVal exclusoes As Variant, _
                             Optional ByVal lote As Variant = "") As Variant
    Dim k As String, reg As Variant, n As Double, soma As Double, somaQ As Double
    Dim media As Double, dp As Double, cv As Double, bias As Variant

    If Trim$(analito) = "" Then EstatPeriodo = "": Exit Function

    On Error GoTo falhou
    Dim ex() As Double, nEx As Long
    nEx = LerExclusoes(exclusoes, ex)
    Agregar ComoData(dtIni), ComoData(dtFim), ex, nEx, Trim$(CStr(lote))

    k = UCase$(Trim$(analito)) & "|" & Trim$(CStr(nivel))
    If Not mAgg.Exists(k) Then
        If UCase$(metrica) = "N" Then EstatPeriodo = 0 Else EstatPeriodo = ""
        Exit Function
    End If

    reg = mAgg(k)
    n = reg(0): soma = reg(1): somaQ = reg(2)

    Select Case UCase$(Trim$(metrica))
        Case "N"
            EstatPeriodo = n
            Exit Function

        Case "MEDIA"
            If n < 1 Then EstatPeriodo = "" Else EstatPeriodo = soma / n
            Exit Function
    End Select

    If n < 1 Then EstatPeriodo = "": Exit Function
    media = soma / n

    ' DP amostral por soma de quadrados. O termo pode dar levemente negativo por
    ' erro de arredondamento quando todos os valores sao iguais; zerar antes da
    ' raiz evita erro de dominio.
    Dim num As Double
    If n >= 2 Then
        num = somaQ - n * media * media
        If num < 0 Then num = 0
        dp = Sqr(num / (n - 1))
    End If

    Select Case UCase$(Trim$(metrica))
        Case "DP"
            If n < 2 Then EstatPeriodo = "" Else EstatPeriodo = dp
            Exit Function

        Case "CV"
            If n < 2 Or media = 0 Then EstatPeriodo = "" Else EstatPeriodo = dp / media * 100#
            Exit Function

        Case "BIAS"
            EstatPeriodo = BiasContraAlvo(analito, nivel, media)
            Exit Function

        Case "ET"
            ' ET% = |Bias%| + 1,65 * CV%
            If n < 2 Or media = 0 Then EstatPeriodo = "": Exit Function
            bias = BiasContraAlvo(analito, nivel, media)
            If Not IsNumeric(bias) Then EstatPeriodo = "": Exit Function
            cv = dp / media * 100#
            EstatPeriodo = Abs(CDbl(bias)) + 1.65 * cv
            Exit Function
    End Select

    EstatPeriodo = CVErr(xlErrValue)
    Exit Function
falhou:
    EstatPeriodo = CVErr(xlErrValue)
End Function

' ---------------------------------------------------------------------------
' Bias % contra o alvo do lote, lido de LotesStore.
'
' LotesStore: A=Lote  B=idx do analito  C=MedN1  D=DPN1  E=MedN2  F=DPN2 ...
' O idx e a POSICAO do analito na aba Analitos.
'
' Devolve Empty quando nao ha alvo -- e nao zero. Zero seria lido como
' "exatidao perfeita" num analito que nunca teve alvo cadastrado.
Public Function BiasContraAlvo(ByVal analito As String, ByVal nivel As Variant, _
                               ByVal mediaObs As Double) As Variant
    Dim alvo As Variant
    alvo = AlvoDoLote(analito, nivel)
    If Not IsNumeric(alvo) Then Exit Function
    If CDbl(alvo) = 0 Then Exit Function
    BiasContraAlvo = (mediaObs - CDbl(alvo)) / CDbl(alvo) * 100#
End Function

Public Function AlvoDoLote(ByVal analito As String, ByVal nivel As Variant) As Variant
    Dim wsA As Worksheet, wsL As Worksheet, i As Long, idx As Long, col As Long
    Dim lote As String

    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Analitos")
    Set wsL = ThisWorkbook.Sheets("LotesStore")
    On Error GoTo 0
    If wsA Is Nothing Or wsL Is Nothing Then Exit Function

    idx = 0
    For i = 4 To 43
        If Trim$(CStr(wsA.Cells(i, 1).Value)) <> "" Then
            idx = idx + 1
            If UCase$(Trim$(CStr(wsA.Cells(i, 1).Value))) = UCase$(Trim$(analito)) Then Exit For
        End If
    Next i
    If i > 43 Then Exit Function

    ' C=MedN1, E=MedN2, G=MedN3 -> 1 + 2*nivel
    col = 1 + 2 * CLng(Val(CStr(nivel)))
    lote = LoteAtivoCore()

    ' TETO VINDO DO DADO, NAO DE UM NUMERO ESCRITO A MAO.
    '
    ' O limite era 200. O LotesStore guarda 40 linhas por lote (mBI.LS_CAP),
    ' entao 200 cobre quatro blocos e meio: a partir do 5o lote o alvo ficava
    ' fora da varredura e esta funcao devolvia vazio -- sem Z, sem bias, e sem
    ' erro. Num sistema dimensionado para 60 meses, trocar de lote a cada
    ' semestre chega la dentro da vida util do arquivo.
    '
    ' Nota de divida tecnica: mBI.AlvoDoLote le ESTE MESMO alvo por aritmetica
    ' de bloco, e nao por varredura. Sao duas implementacoes do mesmo contrato
    ' de layout; concordam hoje por coincidencia de duas leituras corretas, nao
    ' por terem uma fonte so. Unificar exige mexer na Estatistica por periodo,
    ' que nao tem prova automatizada -- fica registrado, nao remendado aqui.
    Dim ultLote As Long
    ultLote = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row
    If ultLote < 2 Then Exit Function
    For i = 2 To ultLote
        If Trim$(CStr(wsL.Cells(i, 1).Value)) = "" Then Exit For
        If Trim$(CStr(wsL.Cells(i, 1).Value)) = lote Then
            If CLng(Val(CStr(wsL.Cells(i, 2).Value))) = idx Then
                If IsNumeric(wsL.Cells(i, col).Value) Then
                    If Len(Trim$(CStr(wsL.Cells(i, col).Value))) > 0 Then
                        AlvoDoLote = CDbl(wsL.Cells(i, col).Value)
                    End If
                End If
                Exit Function
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------------------
' LIMITE VINDO DO MOTOR DE ESPECIFICACOES, EM FORMA SEGURA PARA A PLANILHA.
'
' POR QUE ESTE INVOLUCRO EXISTE -- e o defeito que ele impede.
'
' EspecCVtp/EspecBIAStp/EspecETp devolvem Empty quando NAO HA especificacao
' cadastrada. Isso e correto em VBA. Numa CELULA, porem, Empty e renderizado
' como ZERO -- e dai em diante o estrago e silencioso e total:
'
'   Lim CV VB = 0  ->  StatusCV ve um limite valido de 0,00%
'                  ->  qualquer CV real > 0 "excede" esse limite
'                  ->  a coluna acusa "FORA VB" em TODOS os analitos
'
' Ou seja: analito que nunca teve especificacao de Variacao Biologica apareceria
' reprovado nela. Hoje isso valeria para os 31 analitos, porque nao ha nenhuma
' linha de VB nem de Fabricante no banco.
'
' Devolvendo "" em vez de Empty, a celula fica visualmente vazia e o NumOK do
' StatusCV a reconhece como "limite ausente" -- que cai em SEM LIMITE, nunca em
' aprovacao nem em reprovacao.
'
' tipo: "CV" | "BIAS" | "ETP"
Public Function LimEspec(ByVal analito As String, ByVal ano As Variant, _
                         ByVal fonte As String, ByVal tipo As String) As Variant
    Dim v As Variant, a As Long
    If Trim$(analito) = "" Then LimEspec = "": Exit Function
    If Not IsNumeric(ano) Then LimEspec = "": Exit Function
    a = CLng(Val(CStr(ano)))
    If a < 1900 Then LimEspec = "": Exit Function

    ' VINCULO TARDIO, PELO MESMO MOTIVO DO ADR-041.
    '
    ' EspecCVtp/EspecBIAStp/EspecETp vivem no mEspecificacoes, que HOJE nao
    ' existe em nenhum dos dois produtos -- so a Hematologia tem etapa de build
    ' que o gera, e nem o artefato dela o carrega ainda. Chamada direta cria
    ' dependencia de COMPILACAO: numa pasta sem o modulo, o projeto VBA INTEIRO
    ' deixa de compilar, e On Error nao captura erro de compilacao. O sintoma e
    ' um dialogo modal invisivel que trava qualquer automacao -- ja aconteceu
    ' tres vezes neste projeto.
    '
    ' Com Application.Run a ausencia vira erro de EXECUCAO (1004), que o
    ' On Error abaixo captura e converte em "sem limite". E a degradacao certa:
    ' StatusCV/StatusETP ja tratam "SEM LIMITE" como NAO-aprovacao (ADR-023),
    ' entao ninguem se ve aprovado por um criterio que nao foi avaliado.
    On Error GoTo semLimite
    Select Case UCase$(Trim$(tipo))
        Case "CV":   v = Application.Run("EspecCVtp", analito, a, fonte)
        Case "BIAS": v = Application.Run("EspecBIAStp", analito, a, fonte)
        Case "ETP":  v = Application.Run("EspecETp", analito, a, fonte)
        Case Else:   LimEspec = "": Exit Function
    End Select

    If IsEmpty(v) Or IsNull(v) Then LimEspec = "": Exit Function
    If Not IsNumeric(v) Then LimEspec = "": Exit Function
    LimEspec = CDbl(v)
    Exit Function
semLimite:
    LimEspec = ""
End Function

' ---------------------------------------------------------------------------
' STATUS DA IMPRECISAO (CV)
'
' A REGRA QUE GOVERNA ESTA FUNCAO: LIMITE AUSENTE NAO E APROVACAO.
'
' Hoje o banco de especificacoes so tem CLIA -- nao ha nenhuma linha de Variacao
' Biologica nem de Fabricante. Se a funcao devolvesse "OK" por passar no unico
' limite que existe, o laboratorio se veria aprovado em tres criterios tendo
' sido avaliado por um. E o mesmo defeito que o ADR-023 trata na conformidade:
' ausencia tratada como aprovacao e o jeito mais silencioso de um laboratorio se
' achar em dia.
'
' Por isso o retorno diz CONTRA QUANTOS criterios a aprovacao vale:
'   ""                    sem medida (n<2)
'   "SEM LIMITE"          nenhum dos tres cadastrado: nada foi avaliado
'   "OK (CLIA)"           passou, e so o CLIA existia
'   "OK (CLIA/VB)"        passou nos dois que existiam
'   "FORA CLIA"           excedeu o CLIA
'   "FORA CLIA/VB"        excedeu os dois
Public Function StatusCV(ByVal cvReal As Variant, ByVal limCLIA As Variant, _
                         ByVal limVB As Variant, ByVal limFAB As Variant) As String
    Dim temMedida As Boolean
    temMedida = NumOK(cvReal)
    If Not temMedida Then StatusCV = "": Exit Function

    Dim existentes As String, fora As String
    existentes = ""
    fora = ""

    If NumOK(limCLIA) Then
        existentes = existentes & "/CLIA"
        If CDbl(cvReal) > CDbl(limCLIA) Then fora = fora & "/CLIA"
    End If
    If NumOK(limVB) Then
        existentes = existentes & "/VB"
        If CDbl(cvReal) > CDbl(limVB) Then fora = fora & "/VB"
    End If
    If NumOK(limFAB) Then
        existentes = existentes & "/FAB"
        If CDbl(cvReal) > CDbl(limFAB) Then fora = fora & "/FAB"
    End If

    If Len(existentes) = 0 Then StatusCV = "SEM LIMITE": Exit Function
    If Len(fora) > 0 Then StatusCV = "FORA " & Mid$(fora, 2): Exit Function
    StatusCV = "OK (" & Mid$(existentes, 2) & ")"
End Function

' ---------------------------------------------------------------------------
' STATUS DO ERRO TOTAL
'
' Mesma regra: sem limite nao ha aprovacao. "APROVADO" so aparece quando havia
' criterio para aprovar, e diz qual.
'   ""                        sem ET calculado
'   "SEM LIMITE"              nem CLIA nem VB cadastrados
'   "APROVADO (CLIA)"         passou, e so o CLIA existia
'   "APROVADO (CLIA/VB)"      passou nos dois
'   "FORA CLIA" / "FORA VB"
'   "CRITICO: FORA AMBOS"     excedeu os dois
Public Function StatusETP(ByVal etReal As Variant, ByVal etpCLIA As Variant, _
                          ByVal etpVB As Variant) As String
    If Not NumOK(etReal) Then StatusETP = "": Exit Function

    Dim temC As Boolean, temV As Boolean, foraC As Boolean, foraV As Boolean
    temC = NumOK(etpCLIA)
    temV = NumOK(etpVB)
    If Not temC And Not temV Then StatusETP = "SEM LIMITE": Exit Function

    If temC Then foraC = (CDbl(etReal) > CDbl(etpCLIA))
    If temV Then foraV = (CDbl(etReal) > CDbl(etpVB))

    If foraC And foraV Then StatusETP = "CRITICO: FORA AMBOS": Exit Function
    If foraC Then StatusETP = "FORA CLIA": Exit Function
    If foraV Then StatusETP = "FORA VB": Exit Function

    Dim quais As String
    quais = ""
    If temC Then quais = quais & "/CLIA"
    If temV Then quais = quais & "/VB"
    StatusETP = "APROVADO (" & Mid$(quais, 2) & ")"
End Function

' Numero utilizavel para comparacao? Vazio, texto e erro nao sao.
Private Function NumOK(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    NumOK = IsNumeric(v)
End Function

' ---------------------------------------------------------------------------
' JANELA EFETIVA
'
' UMA pergunta, UMA resposta. O seletor de visao decide; as celulas de periodo
' customizado so valem quando a visao e "PERSONALIZADO". Aceitar as duas fontes
' ao mesmo tempo produziria numero certo para uma janela que o usuario nao
' escolheu -- e ele nao teria como saber qual venceu.
'
' parte = mes (1..12) quando MENSAL, trimestre (1..4), semestre (1..2); ignorado
' quando ANUAL.
Public Function JanelaInicio(ByVal visao As String, ByVal ano As Variant, _
                             ByVal parte As Variant, ByVal custIni As Variant, _
                             ByVal custFim As Variant) As Variant
    JanelaInicio = Janela(visao, ano, parte, custIni, custFim, True)
End Function

Public Function JanelaFim(ByVal visao As String, ByVal ano As Variant, _
                          ByVal parte As Variant, ByVal custIni As Variant, _
                          ByVal custFim As Variant) As Variant
    JanelaFim = Janela(visao, ano, parte, custIni, custFim, False)
End Function

Private Function Janela(ByVal visao As String, ByVal ano As Variant, _
                        ByVal parte As Variant, ByVal custIni As Variant, _
                        ByVal custFim As Variant, ByVal querInicio As Boolean) As Variant
    Dim v As String, a As Long, p As Long, m1 As Long, m2 As Long
    v = UCase$(Trim$(visao))

    If v = "PERSONALIZADO" Then
        If querInicio Then Janela = custIni Else Janela = custFim
        Exit Function
    End If

    If Not IsNumeric(ano) Then Janela = "": Exit Function
    a = CLng(Val(CStr(ano)))
    If a < 1900 Then Janela = "": Exit Function
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
        Janela = DateSerial(a, m1, 1)
    Else
        Janela = DateSerial(a, m2 + 1, 0)      ' dia 0 do mes seguinte = ultimo dia
    End If
End Function

' Texto do periodo realmente aplicado, para a aba MOSTRAR ao usuario qual janela
' venceu -- e nao deixa-lo deduzir.
' NAO USAR IsDate AQUI.
'
' JanelaInicio devolve um Date, mas a CELULA o armazena como serial numerico --
' e IsDate(45658) e False em VBA (IsDate so aceita o tipo Date ou texto com cara
' de data). Testando com IsDate, esta funcao dizia "periodo invalido" para uma
' janela perfeitamente valida, enquanto o motor calculava certo porque ComoData
' aceita numero. Duas leituras diferentes da mesma celula e como se produz uma
' tela que se contradiz.
Public Function PeriodoEfetivo(ByVal dtIni As Variant, ByVal dtFim As Variant, _
                               ByVal exclusoes As Variant) As String
    Dim s As String, i As Double, f As Double, j As Long
    Dim ex() As Double, nEx As Long, cobre As Boolean

    i = ComoData(dtIni): f = ComoData(dtFim)
    If i = 0 Or f = 0 Then PeriodoEfetivo = "periodo nao definido": Exit Function
    If f < i Then PeriodoEfetivo = "FIM ANTES DO INICIO - corrija": Exit Function

    s = Format$(CDate(i), "dd/mm/yyyy") & " a " & Format$(CDate(f), "dd/mm/yyyy")

    nEx = LerExclusoes(exclusoes, ex)
    If nEx = 0 Then
        PeriodoEfetivo = s & "  |  sem exclusoes"
        Exit Function
    End If

    s = s & "  |  EXCLUINDO " & nEx & " periodo(s): "
    For j = 1 To nEx
        If j > 1 Then s = s & " ; "
        s = s & Format$(CDate(ex(j, 1)), "dd/mm") & "-" & Format$(CDate(ex(j, 2)), "dd/mm/yy")
        If ex(j, 1) <= i And ex(j, 2) >= f Then cobre = True
    Next j
    If cobre Then s = s & "   >>> UMA EXCLUSAO COBRE TODO O PERIODO: nao sobra dado"

    PeriodoEfetivo = s
End Function

' ---------------------------------------------------------------------------
' Abre a tela de configuracao da analise.
'
' Fica AQUI, e nao dentro do formulario, por dois motivos: o botao da aba
' precisa de um destino estavel mesmo que o form seja reconstruido, e a fonte
' versionada deste modulo tem de conter tudo que roda dentro do arquivo -- a
' prova 1.1 compara os dois e acusa qualquer diferenca. Acrescentar esta rotina
' no momento da instalacao, direto no .xlsm, e o que fez a 1.1 reprovar.
Public Sub AbrirConfigEstatistica()
    frmConfigEstatistica.Show
End Sub

' Invalida o cache. Chamar apos importacao ou exclusao logica.
Public Sub InvalidarCacheEstat()
    Set mAgg = Nothing
    mCarimbo = ""
End Sub
