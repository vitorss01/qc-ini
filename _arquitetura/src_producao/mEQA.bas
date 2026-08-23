Attribute VB_Name = "mEQA"
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
    Dim wb As Worksheet
    Set wb = AbaPorNome(EQA_BASE)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 3401, "mEQA", "Aba " & EQA_BASE & " nao existe."
    End If

    Dim mapa As Object
    Set mapa = CarregarMapa()

    Dim saida() As Variant
    ReDim saida(1 To BASE_RN - BASE_R0 + 1, 1 To BASE_NCOL)

    Dim k As Long, nCAP As Long, nCTL As Long, semCanon As Long, dup As Long
    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")

    k = 0
    nCAP = Consolidar(AbaPorNome(ABA_CAP), "CAP", mapa, saida, k, vistos, semCanon, dup)
    nCTL = Consolidar(AbaPorNome(ABA_CTL), "Controllab", mapa, saida, k, vistos, semCanon, dup)

    ' limpa o corpo inteiro antes de escrever: sobra de consolidacao anterior
    ' seria linha fantasma, e o mCEQ nao tem como saber que ela e resto.
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
        " | chaves repetidas " & dup
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
Public Function CarimboEQA() As String
    Dim ws As Worksheet
    Set ws = AbaPorNome(EQA_BASE)
    If ws Is Nothing Then CarimboEQA = "EQA_Base ausente": Exit Function
    CarimboEQA = Txt(ws.Cells(CARIMBO_LIN, CARIMBO_COL).Value)
    If CarimboEQA = "" Then CarimboEQA = "ainda nao consolidada"
End Function
