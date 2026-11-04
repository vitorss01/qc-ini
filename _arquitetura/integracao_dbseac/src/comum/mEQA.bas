Option Explicit

' mEQA - ADR-034: interface separada, motor consolidado
'
' O DESENHO
'
'   EQA.CAP_Dados ---------+
'                          +--> EQA_Base --> mCEQ --> Estatistica / Painel / BI
'   EQA.Controllab_Dados --+
'
' Cada provedor tem a SUA aba de digitacao, com a terminologia dele: o CAP fala
' "Survey C-A 2025", "Acceptable", "CHM-01"; o Controllab fala outra coisa.
' Forcar os dois na mesma planilha faria um dos dois mentir.
'
' Mas o MOTOR e um so. EQA_Base normaliza os dois num vocabulario unico, e e
' dela -- nao das abas de digitacao -- que saem bias, Sigma, ET e Power BI. Duas
' implementacoes das mesmas contas divergiriam no dia em que alguem corrigisse
' uma e esquecesse a outra.
'
' ANALITO ORIGINAL E ANALITO CANONICO SAO COISAS DIFERENTES
'
' O CAP reporta "Urea Nitrogen"; a pasta chama "Ureia". "Glucose, serum" e
' "Glicose". Guardar so um dos dois nomes destroi alguma coisa: guardar so o do
' provedor impede o cruzamento com a Analitos; guardar so o canonico apaga a
' rastreabilidade ate o PDF. EQA_Base guarda OS DOIS -- coluna D o nome do
' provedor, coluna E o nome canonico -- e o mapa entre eles vive em W:Y desta
' mesma aba, editavel.
'
' Analito sem correspondente canonico (Ferritin e TSH nao existem na
' Bioquimica) entra na base com E vazio. Ele NAO alimenta Estatistica, e isso e
' correto: nao ha especificacao de qualidade para cruzar. Mas continua visivel,
' contado e rastreavel. Inventar um correspondente seria pior do que a lacuna.
'
' USO_ANALITICO: PRESERVAR NAO E MISTURAR
'
' A EQC_Dados trazia 90 registros que nao sao resultado real de EP -- glicose
' entre 250 e 262 mg/dL nas quinze amostras, com quatro casas decimais, e
' nenhum valor em comum com o relatorio do CAP. Sao dado de simulacao.
'
' Apagar violaria "nao apagar historico". Misturar com o dado real contaminaria
' todo Sigma e todo ET da pasta. Entao a base tem a coluna Uso_Analitico: a
' linha continua ali, visivel e contada, mas com "NAO" ela nao entra em conta
' nenhuma. Quem discordar troca a celula para "SIM".

Public Const EQA_BASE As String = "EQA_Base"
Public Const ABA_CAP As String = "EQA.CAP_Dados"
Public Const ABA_CTL As String = "EQA.Controllab_Dados"

Public Const BASE_R0 As Long = 2          ' linha 1 e cabecalho
Public Const BASE_RN As Long = 5001
Public Const BASE_NCOL As Long = 21

' colunas da EQA_Base
Public Const B_PROVEDOR As Long = 1
Public Const B_ANO As Long = 2
Public Const B_RODADA As Long = 3
Public Const B_ANALITO As Long = 4        ' como o provedor escreve
Public Const B_CANONICO As Long = 5       ' como a pasta chama
Public Const B_AMOSTRA As Long = 6
Public Const B_RESULTADO As Long = 7
Public Const B_ALVO As Long = 8
Public Const B_SD As Long = 9
Public Const B_SDI As Long = 10
Public Const B_LIMINF As Long = 11
Public Const B_LIMSUP As Long = 12
Public Const B_AVAL_ORIG As Long = 13
Public Const B_STATUS As Long = 14
Public Const B_UNIDADE As Long = 15
Public Const B_BIAS As Long = 16
Public Const B_BIASABS As Long = 17
Public Const B_PAGINA As Long = 18
Public Const B_ARQUIVO As Long = 19
Public Const B_USO As Long = 20
Public Const B_CHAVE As Long = 21

' colunas das abas de digitacao (identicas nas duas, por desenho)
Public Const D_R0 As Long = 2
Public Const D_PROVEDOR As Long = 1
Public Const D_RODADA As Long = 2
Public Const D_ANO As Long = 3
Public Const D_ANALITO As Long = 4
Public Const D_AMOSTRA As Long = 5
Public Const D_RESULTADO As Long = 6
Public Const D_ALVO As Long = 7
Public Const D_SD As Long = 8
Public Const D_SDI As Long = 9
Public Const D_LIMINF As Long = 10
Public Const D_LIMSUP As Long = 11
Public Const D_AVALIACAO As Long = 12
Public Const D_NLABS As Long = 13
Public Const D_UNIDADE As Long = 14
Public Const D_BIAS As Long = 15
Public Const D_BIASABS As Long = 16
Public Const D_PAGINA As Long = 17
Public Const D_ARQUIVO As Long = 18
Public Const D_NCOL As Long = 18

' mapa de analitos, na propria EQA_Base
Public Const MAPA_C0 As Long = 23         ' W = provedor
Public Const MAPA_R0 As Long = 2
Public Const MAPA_RN As Long = 201

' rotulos padronizados de status
Public Const ST_ACEITO As String = "ACEITO"
Public Const ST_NAO_ACEITO As String = "NAO ACEITO"
Public Const ST_NAO_AVALIADO As String = "NAO AVALIADO"

' marcador das linhas que nao alimentam analise
Public Const USO_SIM As String = "SIM"
Public Const USO_NAO As String = "NAO"

' onde AtualizarEQABase deixa o carimbo da ultima consolidacao
Public Const CARIMBO_LIN As Long = 1
Public Const CARIMBO_COL As Long = 26     ' Z1
' ADR-089 (OC-0028): instante NUMERICO da ultima consolidacao (EQA_Base!AA1, formato oculto) e a marca de edicao
' direta de cada aba de digitacao (Z4, formato oculto, gravada pelo Worksheet_Change da aba). Marca mais nova que a
' consolidacao = ha edicao na tabela que ainda nao chegou a EQA_Base (o Z3 da aba avisa; o ATUALIZAR consolida).
Public Const CARIMBO_NUM_COL As Long = 27 ' AA1
Public Const MARCA_EDICAO As String = "Z4"

' ADR-064: n de laboratorios do grupo por chave (NLabsPorChave)
Private mNLabs As Object
Private mNLabsSig As String


' NAO renomear para "Ws". VBA nao distingue maiusculas: uma variavel local
' chamada ws passaria a sombrear esta funcao, e "Set ws = Ws(nome)" viraria
' chamada ao membro padrao da propria variavel vazia -- erro 91, em tempo de
' execucao, dentro de uma instancia sem tela para mostrar a caixa.
Private Function AbaPorNome(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set AbaPorNome = ThisWorkbook.Sheets(nome)
    On Error GoTo 0
End Function


Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function


' Ultima linha realmente preenchida de uma aba de digitacao. Nao usa
' UsedRange: formatacao propagada por tabela estruturada infla o UsedRange e
' faria a consolidacao varrer milhares de linhas vazias.
Private Function UltimaLinha(ByVal ws As Worksheet) As Long
    UltimaLinha = ws.Cells(ws.Rows.Count, D_ANALITO).End(xlUp).Row
    If UltimaLinha < D_R0 Then UltimaLinha = D_R0 - 1
End Function


' ---------------------------------------------------------------------------
' Mapa provedor+nome do provedor -> nome canonico.
'
' Chave em UCase para nao depender de capitalizacao; o provedor entra na chave
' porque "Lactate" do CAP e "Lactato" do Controllab podem, em tese, apontar
' para canonicos diferentes se um dia medirem coisas diferentes.
' ---------------------------------------------------------------------------
Private Function CarregarMapa() As Object
    Dim ws As Worksheet, i As Long, k As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = AbaPorNome(EQA_BASE)
    If ws Is Nothing Then Set CarregarMapa = d: Exit Function

    Dim m As Variant
    m = ws.Range(ws.Cells(MAPA_R0, MAPA_C0), _
                 ws.Cells(MAPA_RN, MAPA_C0 + 2)).Value
    For i = 1 To UBound(m, 1)
        If Txt(m(i, 2)) <> "" Then
            k = UCase$(Txt(m(i, 1))) & "|" & UCase$(Txt(m(i, 2)))
            If Not d.Exists(k) Then d(k) = Txt(m(i, 3))
        End If
    Next i
    Set CarregarMapa = d
End Function


' Nome canonico do analito, ou "" quando o provedor reporta algo que a pasta
' nao mede. "" e resposta legitima -- ver o cabecalho do modulo.
Public Function MapearAnalito(ByVal provedor As String, ByVal analito As String) As String
    Dim d As Object
    Set d = CarregarMapa()
    MapearAnalito = ""
    Dim k As String
    k = UCase$(Trim$(provedor)) & "|" & UCase$(Trim$(analito))
    If d.Exists(k) Then MapearAnalito = d(k)
End Function


' Traduz a avaliacao do provedor para o vocabulario unico da base.
'
' NAO tenta adivinhar: so classifica o que reconhece. Termo desconhecido volta
' como NAO AVALIADO, e nao como ACEITO -- aprovar por omissao seria o pior
' erro possivel aqui.
Public Function PadronizarStatus(ByVal avaliacao As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(avaliacao)))
    PadronizarStatus = ST_NAO_AVALIADO
    If s = "" Then Exit Function

    Select Case s
        Case "ACCEPTABLE", "ACEITAVEL", "ACEITÁVEL", "ACEITO", "SATISFATORIO", _
             "SATISFATÓRIO", "CONFORME", "OK", "APROVADO"
            PadronizarStatus = ST_ACEITO
        Case "UNACCEPTABLE", "NAO ACEITAVEL", "NÃO ACEITÁVEL", "NAO ACEITO", _
             "NÃO ACEITO", "INSATISFATORIO", "INSATISFATÓRIO", "NAO CONFORME", _
             "NÃO CONFORME", "DESVIO", "REPROVADO"
            PadronizarStatus = ST_NAO_ACEITO
    End Select
End Function


' ---------------------------------------------------------------------------
' Consolida as duas abas de digitacao na EQA_Base.
'
' Escreve VALORES, nao formulas. A base tem ate 5000 linhas e e lida celula a
' celula pelo mCEQ a cada recalculo; 5000 x 21 formulas vivas custariam caro
' por um dado que so muda quando alguem digita uma rodada nova.
' ---------------------------------------------------------------------------
Public Sub AtualizarEQABase()
    ConsolidarBase True, "botão Consolidar / chamada direta"
End Sub

' ADR-089: a janela do CEQ consolida dizendo de onde veio (trilha EQA_BASE_CONSOLIDADA).
Public Sub ConsolidarEQABaseDe(ByVal origem As String)
    ConsolidarBase True, origem
End Sub

' ADR-089 (OC-0028): o ATUALIZAR DADOS consolida so quando ha pendencia (edicao direta na tabela depois da ultima
' consolidacao, ou quantidade de linhas diferente). Sem as listas e a incerteza: o ATUALIZAR as refaz logo depois.
Public Function ConsolidarSePendente() As Boolean
    If Not ConsolidacaoPendente() Then Exit Function
    ConsolidarBase False, "ATUALIZAR DADOS (edição pendente)"
    ConsolidarSePendente = True
End Function

' Instante com fracao de segundo (Now so tem segundos: editar e consolidar no mesmo segundo empatariam).
Public Function InstanteCEQ() As Double
    InstanteCEQ = CDbl(Date) + Timer / 86400#
End Function

Private Function NumCel(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NumCel = CDbl(v)
End Function

' Ha edicao das abas de digitacao que a EQA_Base ainda nao recebeu? (a mesma pergunta do Z3 das abas)
Public Function ConsolidacaoPendente() As Boolean
    Dim b As Worksheet, ws As Worksheet, t As Double, p As Variant, prov As String
    On Error GoTo sai
    Set b = AbaPorNome(EQA_BASE)
    If b Is Nothing Then Exit Function
    t = NumCel(b.Cells(CARIMBO_LIN, CARIMBO_NUM_COL).Value)
    For Each p In Array(ABA_CAP, ABA_CTL)
        Set ws = AbaPorNome(CStr(p))
        If Not ws Is Nothing Then
            If CStr(p) = ABA_CAP Then prov = "CAP" Else prov = "Controllab"
            If NumCel(ws.Range(MARCA_EDICAO).Value) > t Then ConsolidacaoPendente = True: Exit Function
            If Application.WorksheetFunction.CountA(ws.Range("D2:D5000")) <> _
               Application.WorksheetFunction.CountIf(b.Range("A2:A5001"), prov) Then ConsolidacaoPendente = True: Exit Function
        End If
    Next p
sai:
End Function

' Marca de edicao direta numa aba de digitacao (Worksheet_Change da aba): o Z3 passa a avisar.
Public Sub MarcarEdicaoCEQ(ByVal ws As Worksheet)
    Dim prot As Boolean, ev As Boolean
    On Error Resume Next
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(ws)
    ws.Range(MARCA_EDICAO).Value = InstanteCEQ()
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub

' Retrato das linhas da base: chave -> conteudo das 21 colunas (chave repetida ganha o numero da linha).
Private Function RetratoBase(ByVal v As Variant, ByVal n As Long) As Object
    Dim d As Object, i As Long, c As Long, s As String, k As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To n
        k = Txt(v(i, B_CHAVE))
        If k <> "" Then
            s = ""
            For c = 1 To BASE_NCOL
                s = s & Chr$(31) & Txt(v(i, c))
            Next c
            If d.Exists(k) Then k = k & "#" & i
            d(k) = s
        End If
    Next i
    Set RetratoBase = d
End Function

Private Sub ConsolidarBase(ByVal recalcular As Boolean, ByVal origem As String)
    Dim wb As Worksheet, protB As Boolean, carAntes As String
    Set wb = AbaPorNome(EQA_BASE)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 3401, "mEQA", "Aba " & EQA_BASE & " nao existe."
    End If
    carAntes = Txt(wb.Cells(CARIMBO_LIN, CARIMBO_COL).Value)

    Dim mapa As Object
    Set mapa = CarregarMapa()

    Dim saida() As Variant
    ReDim saida(1 To BASE_RN - BASE_R0 + 1, 1 To BASE_NCOL)

    Dim k As Long, nCAP As Long, nCTL As Long, semCanon As Long, dup As Long
    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")

    ' Decisao MANUAL de Uso_Analitico (o cabecalho manda trocar a celula para SIM/NAO) sobrevive a
    ' consolidacao: antes de limpar, guarda por chave o Uso que DIFERE da regra automatica e o reaplica
    ' (auditoria 04/10/2026 -- a consolidacao regravava a regra e apagava a decisao sem aviso).
    Dim manual As Object, antigo As Variant, ia As Long, nMan As Long
    Set manual = CreateObject("Scripting.Dictionary")
    manual.CompareMode = 1
    antigo = wb.Range(wb.Cells(BASE_R0, 1), wb.Cells(BASE_RN, BASE_NCOL)).Value
    For ia = 1 To UBound(antigo, 1)
        If Txt(antigo(ia, B_CHAVE)) <> "" Then
            If UCase$(Txt(antigo(ia, B_USO))) = USO_SIM Or UCase$(Txt(antigo(ia, B_USO))) = USO_NAO Then
                If UCase$(Txt(antigo(ia, B_USO))) <> UsoAnalitico(antigo(ia, B_ARQUIVO)) Then
                    manual(UCase$(Txt(antigo(ia, B_CHAVE)))) = UCase$(Txt(antigo(ia, B_USO)))
                End If
            End If
        End If
    Next ia

    k = 0
    nCAP = Consolidar(AbaPorNome(ABA_CAP), "CAP", mapa, saida, k, vistos, semCanon, dup)
    nCTL = Consolidar(AbaPorNome(ABA_CTL), "Controllab", mapa, saida, k, vistos, semCanon, dup)
    For ia = 1 To k
        If manual.Exists(UCase$(Txt(saida(ia, B_CHAVE)))) Then
            saida(ia, B_USO) = manual(UCase$(Txt(saida(ia, B_CHAVE))))
            nMan = nMan + 1
        End If
    Next ia

    ' limpa o corpo inteiro antes de escrever: sobra de consolidacao anterior
    ' seria linha fantasma, e o mCEQ nao tem como saber que ela e resto.
    ' ADR-089: a EQA_Base e protegida; sem o UserInterfaceOnly da sessao (automacao, arquivo recem-aberto) a escrita
    ' dava erro 1004 (ADR-046: cada rotina que escreve trata a propria janela)
    protB = mSeguranca.LiberarEscritaRapida(wb)
    On Error GoTo devolve
    wb.Range(wb.Cells(BASE_R0, 1), wb.Cells(BASE_RN, BASE_NCOL)).ClearContents

    If k > 0 Then
        Dim bloco() As Variant, r As Long, c As Long
        ReDim bloco(1 To k, 1 To BASE_NCOL)
        For r = 1 To k
            For c = 1 To BASE_NCOL
                bloco(r, c) = saida(r, c)
            Next c
        Next r
        wb.Range(wb.Cells(BASE_R0, 1), _
                 wb.Cells(BASE_R0 + k - 1, BASE_NCOL)).Value = bloco
    End If

    wb.Cells(CARIMBO_LIN, CARIMBO_COL).Value = _
        "consolidado: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | CAP " & nCAP & " | Controllab " & nCTL & _
        " | total " & k & " | sem analito canonico " & semCanon & _
        " | chaves repetidas " & dup & " | uso manual mantido " & nMan
    wb.Cells(CARIMBO_LIN, CARIMBO_NUM_COL).Value = InstanteCEQ()      ' ADR-089: o Z3 das abas compara com este
    mSeguranca.RestaurarProtecao wb, protB
    protB = False
    On Error GoTo 0

    ' ADR-089 (OC-0032): trilha da consolidacao, com o antes e o depois
    Dim rA As Object, rD As Object, kk As Variant, nInc As Long, nAlt As Long, nRem As Long
    On Error Resume Next
    Set rA = RetratoBase(antigo, UBound(antigo, 1))
    Set rD = RetratoBase(saida, k)
    For Each kk In rD.Keys
        If Not rA.Exists(kk) Then
            nInc = nInc + 1
        ElseIf rA(kk) <> rD(kk) Then
            nAlt = nAlt + 1
        End If
    Next kk
    For Each kk In rA.Keys
        If Not rD.Exists(kk) Then nRem = nRem + 1
    Next kk
    mAuditoria.RegistrarLog "EQA_BASE_CONSOLIDADA", "origem: " & origem & " | linhas " & k & " (CAP " & nCAP & _
        ", Controllab " & nCTL & ", sem analito canônico " & semCanon & ") | incluídas " & nInc & " · alteradas " & _
        nAlt & " · removidas " & nRem & " | antes: " & IIf(carAntes = "", "(sem consolidação)", carAntes) & _
        " | depois: " & Txt(wb.Cells(CARIMBO_LIN, CARIMBO_COL).Value)
    Err.Clear
    On Error GoTo 0

    ' As formulas do vies/incerteza (Estatistica AR:AU) leem a EQA_Base pelo VBA: o Excel nao sabe
    ' que dependem dela e mostrariam o CEQ antigo ate o proximo ATUALIZAR DADOS (auditoria 04/10/2026).
    On Error Resume Next
    mCEQ.InvalidarEQ
    ' ADR-071: rodada/ano novos entram nas listas de selecao (Estatistica!N4/P4) sem reabrir o arquivo
    ' (ADR-089: no ATUALIZAR, recalcular = False -- ele mesmo refaz as listas e a incerteza logo depois)
    If recalcular Then mDados.AtualizarListasAno
    ' ADR-071 (revisao): G/R/S/T/AC/AD e K5 leem a EQA_Base pelo VBA e nao dependem de celula que mudou:
    ' marca-os sujos para o Calculate do RecalcularIncerteza (senao ficavam com o conjunto antigo)
    mCEQ.MarcarCEQSujo
    If recalcular Then mIncerteza.RecalcularIncerteza
    Exit Sub
devolve:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If protB Then mSeguranca.RestaurarProtecao wb, protB
    On Error GoTo 0
    Err.Raise nE, "mEQA.AtualizarEQABase", sE
End Sub

' O n de laboratorios e digitado depois da consolidacao (nao vai para a EQA_Base): o cache e
' renovado a cada recalculo da incerteza, nao so quando o carimbo da consolidacao muda.
Public Sub InvalidarNLabs()
    Set mNLabs = Nothing
    mNLabsSig = ""
End Sub


' Le uma aba de digitacao e empilha em saida(). Devolve quantas linhas leu.
'
' Chave repetida NAO e descartada: entra na base e e CONTADA. Sumir com
' duplicata em silencio esconderia digitacao dobrada -- exatamente o erro que
' a chave existe para revelar.
Private Function Consolidar(ByVal ws As Worksheet, ByVal provedor As String, _
                            ByVal mapa As Object, ByRef saida() As Variant, _
                            ByRef k As Long, ByVal vistos As Object, _
                            ByRef semCanon As Long, ByRef dup As Long) As Long
    Consolidar = 0
    If ws Is Nothing Then Exit Function

    Dim ult As Long
    ult = UltimaLinha(ws)
    If ult < D_R0 Then Exit Function

    Dim d As Variant
    d = ws.Range(ws.Cells(D_R0, 1), ws.Cells(ult, D_NCOL)).Value

    Dim i As Long, n As Long, canon As String, ch As String, mk As String
    For i = 1 To UBound(d, 1)
        If Txt(d(i, D_ANALITO)) = "" Then GoTo prox
        If k >= BASE_RN - BASE_R0 + 1 Then Exit For

        mk = UCase$(provedor) & "|" & UCase$(Txt(d(i, D_ANALITO)))
        canon = ""
        If mapa.Exists(mk) Then canon = mapa(mk)
        If canon = "" Then semCanon = semCanon + 1

        ch = provedor & "|" & Txt(d(i, D_ANO)) & "|" & Txt(d(i, D_RODADA)) & _
             "|" & Txt(d(i, D_ANALITO)) & "|" & Txt(d(i, D_AMOSTRA))
        If vistos.Exists(UCase$(ch)) Then
            dup = dup + 1
        Else
            vistos(UCase$(ch)) = 1
        End If

        k = k + 1
        n = n + 1
        saida(k, B_PROVEDOR) = provedor
        saida(k, B_ANO) = d(i, D_ANO)
        saida(k, B_RODADA) = d(i, D_RODADA)
        saida(k, B_ANALITO) = d(i, D_ANALITO)
        saida(k, B_CANONICO) = canon
        saida(k, B_AMOSTRA) = d(i, D_AMOSTRA)
        saida(k, B_RESULTADO) = d(i, D_RESULTADO)
        saida(k, B_ALVO) = d(i, D_ALVO)
        saida(k, B_SD) = d(i, D_SD)
        saida(k, B_SDI) = d(i, D_SDI)
        saida(k, B_LIMINF) = d(i, D_LIMINF)
        saida(k, B_LIMSUP) = d(i, D_LIMSUP)
        saida(k, B_AVAL_ORIG) = d(i, D_AVALIACAO)
        saida(k, B_STATUS) = PadronizarStatus(d(i, D_AVALIACAO))
        saida(k, B_UNIDADE) = d(i, D_UNIDADE)
        saida(k, B_BIAS) = d(i, D_BIAS)
        saida(k, B_BIASABS) = d(i, D_BIASABS)
        saida(k, B_PAGINA) = d(i, D_PAGINA)
        saida(k, B_ARQUIVO) = d(i, D_ARQUIVO)
        saida(k, B_USO) = UsoAnalitico(d(i, D_ARQUIVO))
        saida(k, B_CHAVE) = ch
prox:
    Next i
    Consolidar = n
End Function


' Uma linha entra em bias, Sigma e ET? Sai "NAO" so para o que esta marcado na
' origem como simulacao. O criterio e explicito de proposito: qualquer outro
' dado entra, e a excecao fica visivel na coluna.
Public Function UsoAnalitico(ByVal arquivoFonte As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(arquivoFonte)))
    If InStr(s, "SIMULA") > 0 Then
        UsoAnalitico = USO_NAO
    Else
        UsoAnalitico = USO_SIM
    End If
End Function


' Quantas linhas a base tem agora, para as abas de digitacao avisarem quando
' estiverem na frente dela.
Public Function LinhasEQABase() As Long
    Dim ws As Worksheet
    Set ws = AbaPorNome(EQA_BASE)
    If ws Is Nothing Then Exit Function
    LinhasEQABase = ws.Cells(ws.Rows.Count, B_PROVEDOR).End(xlUp).Row - BASE_R0 + 1
    If LinhasEQABase < 0 Then LinhasEQABase = 0
End Function


' Texto do carimbo da ultima consolidacao, para exibir nas abas de digitacao.
' ADR-089 (OC-0029/0030): o mapa de UM provedor, para a janela do CEQ mostrar o canonico e os pares %/#.
' Matriz (n x 2) nome do provedor | canonico; Empty sem nenhuma entrada.
Public Function MapaDoProvedor(ByVal provedor As String) As Variant
    Dim ws As Worksheet, m As Variant, i As Long, n As Long, out() As Variant, idx() As Long
    Set ws = AbaPorNome(EQA_BASE)
    If ws Is Nothing Then Exit Function
    m = ws.Range(ws.Cells(MAPA_R0, MAPA_C0), ws.Cells(MAPA_RN, MAPA_C0 + 2)).Value
    ReDim idx(1 To UBound(m, 1))
    For i = 1 To UBound(m, 1)
        If Txt(m(i, 2)) <> "" And StrComp(Txt(m(i, 1)), Trim$(provedor), vbTextCompare) = 0 Then n = n + 1: idx(n) = i
    Next i
    If n = 0 Then Exit Function
    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = Txt(m(idx(i), 2))
        out(i, 2) = Txt(m(idx(i), 3))
    Next i
    MapaDoProvedor = out
End Function

Public Function CarimboEQA() As String
    Dim ws As Worksheet
    Set ws = AbaPorNome(EQA_BASE)
    If ws Is Nothing Then CarimboEQA = "EQA_Base ausente": Exit Function
    CarimboEQA = Txt(ws.Cells(CARIMBO_LIN, CARIMBO_COL).Value)
    If CarimboEQA = "" Then CarimboEQA = "ainda nao consolidada"
End Function

' ---------------------------------------------------------------------------
' ADR-064: numero de laboratorios do grupo (participantes), por CHAVE da EQA_Base.
'
' Fica nas abas de digitacao (coluna "Nº Laboratórios") e NAO e copiado para a
' EQA_Base: a base tem o mapa de analitos colado a direita (W:Y) e crescer a
' tabela para la arriscaria engolir o mapa. A chave e montada aqui pela MESMA
' regra do Consolidar -- um dono so para a chave. Serve a u(Cref) do Nordtest
' (DP do grupo / raiz(n laboratorios)), em mCEQ.
' Cache valido enquanto o carimbo da ultima consolidacao nao mudar.
Public Function NLabsPorChave(ByVal chave As String) As Variant
    Dim sig As String
    NLabsPorChave = Empty
    sig = CarimboEQA()
    If mNLabs Is Nothing Or sig <> mNLabsSig Then
        Set mNLabs = CreateObject("Scripting.Dictionary")
        mNLabs.CompareMode = 1
        mNLabsSig = sig
        CarregarNLabs AbaPorNome(ABA_CAP), "CAP"
        CarregarNLabs AbaPorNome(ABA_CTL), "Controllab"
    End If
    If mNLabs.Exists(UCase$(Trim$(chave))) Then NLabsPorChave = mNLabs(UCase$(Trim$(chave)))
End Function

Private Sub CarregarNLabs(ByVal ws As Worksheet, ByVal provedor As String)
    Dim ult As Long, d As Variant, i As Long, ch As String
    If ws Is Nothing Then Exit Sub
    ult = UltimaLinha(ws)
    If ult < D_R0 Then Exit Sub
    d = ws.Range(ws.Cells(D_R0, 1), ws.Cells(ult, D_NCOL)).Value
    For i = 1 To UBound(d, 1)
        If Txt(d(i, D_ANALITO)) <> "" Then
            ch = provedor & "|" & Txt(d(i, D_ANO)) & "|" & Txt(d(i, D_RODADA)) & _
                 "|" & Txt(d(i, D_ANALITO)) & "|" & Txt(d(i, D_AMOSTRA))
            If IsNumeric(d(i, D_NLABS)) And Len(Txt(d(i, D_NLABS))) > 0 Then
                If CDbl(d(i, D_NLABS)) > 0 Then mNLabs(UCase$(ch)) = CDbl(d(i, D_NLABS))
            End If
        End If
    Next i
End Sub

