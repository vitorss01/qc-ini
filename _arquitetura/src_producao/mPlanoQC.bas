Attribute VB_Name = "mPlanoQC"
Option Explicit

' mPlanoQC - ADR-035: do Sigma ate a decisao operacional
'
' A CADEIA
'
'   Sigma -> DPM teorico -> Rendimento teorico -> Plano de CQ
'                                                 (regras, N, run size)
'
' Um numero Sigma isolado nao decide nada. O que decide e a consequencia:
' quantas regras rodar, quantos controles medir, quantos pacientes podem passar
' entre um evento de CQ e o proximo. Este modulo faz essa traducao, e faz UMA
' vez -- Estatistica, Painel e Power BI chamam as mesmas funcoes.
'
' DPM E BENCHMARK TEORICO, NAO CONTAGEM DE ERRO
'
' A conversao Sigma -> DPM usa a convencao de SHORT-TERM SIGMA, com o
' deslocamento convencional de 1,5 SD:
'
'     DPM = [1 - Phi(Sigma - 1,5)] x 1.000.000
'
' Westgard S, Bayat H, Westgard JO. Analytical Sigma metrics. Biochem Med
' (Zagreb) 2018;28(2):020502, Table 1 -- "Short-term Sigma is the most
' commonly-used and cited metric, and it assumes there is a 1.5 SD shift that
' occurs as the "natural" variation in a process over the long-term operation
' of a process."
'
' Isso NAO e "o laboratorio produziu N resultados errados". E a expectativa
' teorica de defeitos associada aquele Sigma. A distincao e obrigatoria em
' qualquer tela que exiba o numero.
'
' CALCULO CONTINUO, NAO TABELA
'
' Sigma real e 4,27 ou 5,43, nao 4 ou 5. Arredondar antes de converter jogaria
' fora justamente a resolucao que o indicador tem. A tabela publicada serve de
' conferencia -- e as provas conferem contra ela nos nove pontos.
'
' O PLANO VEM DE UMA TABELA, NAO DE IFs ESPALHADOS
'
' As faixas vivem em tblPlanoQC_Sigma (aba Cfg_PlanoQC). Mudar uma faixa e
' editar uma linha da tabela; nenhuma formula precisa ser reescrita. IFs
' encadeados em tres abas divergiriam no primeiro esquecimento.
'
' N NAO E NIVEL DE CONTROLE
'
' N e o numero TOTAL de medicoes de controle no evento. N4 pode ser dois niveis
' em duplicata, ou quatro niveis, ou outra combinacao -- depende da configuracao
' real do laboratorio, e o sistema nao presume qual.
'
' RUN SIZE NAO E R_4s
'
' R_4s e regra de Westgard. Run size e quantos pacientes podem passar entre
' eventos de CQ. Sao coisas diferentes e a interface nunca as abrevia junto.

Public Const CFG_ABA As String = "Cfg_PlanoQC"
Public Const CFG_R0 As Long = 3           ' primeira linha de faixa
Public Const CFG_RN As Long = 20

' colunas de tblPlanoQC_Sigma
Public Const P_SMIN As Long = 1
Public Const P_SMAX As Long = 2
Public Const P_CLASSE As Long = 3
Public Const P_REGRAS As Long = 4
Public Const P_N As Long = 5
Public Const P_RUNSIZE As Long = 6
Public Const P_FREQ As Long = 7
Public Const P_REFER As Long = 8
Public Const P_NCOL As Long = 8

' O deslocamento convencional do modelo de curto prazo. Nao e ajuste livre:
' mudar isto muda o significado de todo DPM publicado pela pasta.
Public Const DESLOCAMENTO_SHORT_TERM As Double = 1.5

Public Const NOTA_DPM As String = _
    "DPM teórico estimado pelo Sigma: usa a convenção de short-term Sigma " & _
    "com deslocamento de 1,5 SD. É um benchmark teórico de desempenho, e " & _
    "não uma contagem observada de erros em resultados de pacientes."

Public Const NOTA_RUNSIZE As String = _
    "O run size é recomendação de planejamento de SQC baseada em desempenho " & _
    "Sigma e risco. Não substitui requisitos regulatórios, de acreditação, " & _
    "instruções do fabricante ou procedimentos internos mais restritivos."

' Regras que o motor desta pasta realmente avalia no Calc. Conferido coluna a
' coluna: K=1_3s, L=2_2s, M=R_4s, N=4_1s, O=8x.
'
' AS REGRAS QUE O MOTOR AVALIA VEM DO PROPRIO MOTOR (ADR-041)
'
' Era uma constante escrita aqui a mao: "1_3s;2_2s;R_4s;4_1s;8x". Com dois
' produtos isso vira defeito imediato -- a Hematologia avalia 2of3_2s, 3_1s e
' 6x, e a lista fixa afirmaria o contrario, fazendo CoberturaWestgard responder
' TOTAL sobre regras que aquele motor nao executa.
'
' Agora a lista vem de mEstatistica.MatrizWestgard(), que e o mesmo lugar que
' DECIDE as regras a partir de NLV. Uma fonte so: se a matriz mudar, a
' cobertura acompanha sem ninguem lembrar de editar dois arquivos.
' O CONTRATO ESPERADO DE CADA REGRA (ADR-044)
'
' Isto e a ESPECIFICACAO; DETECTORES, no mEstatistica, e a DECLARACAO do que
' o motor implementa. Comparar as duas nao e ter duas fontes da mesma coisa:
' e o mesmo papel de um teste que afirma o valor esperado. Se fossem o mesmo
' texto lido do mesmo lugar, a conferencia nao provaria nada.
'
'   Area|Regra|Detector|N|R|Escopo
Private Const CONTRATO As String = _
    "BIOQUIMICA|1_3s|INDIVIDUAL|1|1|WITHIN_RUN;" & _
    "BIOQUIMICA|2_2s|WITHIN_RUN|2|1|WITHIN_RUN_ACROSS_MATERIALS;" & _
    "BIOQUIMICA|R_4s|WITHIN_RUN|2|1|WITHIN_RUN_ACROSS_MATERIALS;" & _
    "BIOQUIMICA|4_1s|N2_R2|2|2|ACROSS_RUN_ACROSS_MATERIALS;" & _
    "BIOQUIMICA|8x|N2_R4|2|4|ACROSS_RUN_ACROSS_MATERIALS;" & _
    "HEMATOLOGIA|1_3s|INDIVIDUAL|1|1|WITHIN_RUN;" & _
    "HEMATOLOGIA|2of3_2s|WITHIN_RUN|3|1|WITHIN_RUN_ACROSS_MATERIALS;" & _
    "HEMATOLOGIA|R_4s|WITHIN_RUN|3|1|WITHIN_RUN_ACROSS_MATERIALS;" & _
    "HEMATOLOGIA|3_1s|N3_R1|3|1|WITHIN_RUN_ACROSS_MATERIALS;" & _
    "HEMATOLOGIA|6x|N3_R2|3|2|ACROSS_RUN_ACROSS_MATERIALS"

Private Function RegrasDoMotor() As String
    ' Application.Run e nao mEstatistica.MatrizWestgard(): a chamada direta
    ' cria dependencia de COMPILACAO. Numa pasta cujo mEstatistica ainda nao
    ' tenha a funcao, o projeto VBA inteiro deixa de compilar -- e erro de
    ' compilacao NAO e capturado por On Error, entao a pasta quebra em vez de
    ' degradar. Com vinculo tardio a ausencia vira erro de execucao, que o
    ' tratamento abaixo absorve.
    On Error GoTo semMotor
    RegrasDoMotor = CStr(Application.Run("MatrizWestgard"))
    If Len(Trim$(RegrasDoMotor)) > 0 Then Exit Function
semMotor:
    ' Sem o motor carregado nao ha o que declarar. Devolver a matriz de dois
    ' niveis "por seguranca" seria afirmar cobertura que ninguem verificou.
    RegrasDoMotor = ""
End Function


Private Function AbaCfg() As Worksheet
    On Error Resume Next
    Set AbaCfg = ThisWorkbook.Sheets(CFG_ABA)
    On Error GoTo 0
End Function


Private Function EhNumero(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(CStr(v)) = "" Then Exit Function
    End If
    EhNumero = IsNumeric(v)
End Function


' Phi(z) -- funcao de distribuicao acumulada da normal padrao.
'
' Usa a funcao da planilha, que e a mesma que o Excel expoe ao usuario: assim o
' numero da celula e o numero do VBA nao podem divergir. Norm_S_Dist existe do
' Excel 2010 em diante; NormSDist e o nome antigo, mantido como reserva.
Private Function Phi(ByVal z As Double) As Double
    On Error GoTo antigo
    Phi = Application.WorksheetFunction.Norm_S_Dist(z, True)
    Exit Function
antigo:
    On Error GoTo falhou
    Phi = Application.WorksheetFunction.NormSDist(z)
    Exit Function
falhou:
    Phi = -1
End Function


' DPM teorico estimado pelo Sigma. Devolve "" quando nao ha Sigma -- nunca 0,
' que seria ler "processo perfeito" onde nao ha medida nenhuma.
Public Function DPMdoSigma(ByVal sigma As Variant) As Variant
    DPMdoSigma = ""
    If Not EhNumero(sigma) Then Exit Function

    Dim z As Double, p As Double
    z = CDbl(sigma) - DESLOCAMENTO_SHORT_TERM
    p = Phi(z)
    If p < 0 Then Exit Function          ' nem Norm_S_Dist nem NormSDist responderam
    DPMdoSigma = (1# - p) * 1000000#
End Function


' Rendimento teorico = 1 - DPM/1.000.000, em porcentagem.
Public Function RendimentoDoSigma(ByVal sigma As Variant) As Variant
    RendimentoDoSigma = ""
    Dim d As Variant
    d = DPMdoSigma(sigma)
    If Not EhNumero(d) Then Exit Function
    RendimentoDoSigma = (1# - CDbl(d) / 1000000#) * 100#
End Function


' Qual linha de tblPlanoQC_Sigma vale para este Sigma.
' Faixa fechada embaixo e aberta em cima: Sigma_Min <= s < Sigma_Max.
Private Function LinhaDoPlano(ByVal sigma As Variant) As Long
    LinhaDoPlano = 0
    If Not EhNumero(sigma) Then Exit Function
    Dim ws As Worksheet
    Set ws = AbaCfg()
    If ws Is Nothing Then Exit Function

    Dim d As Variant, i As Long, s As Double
    s = CDbl(sigma)
    d = ws.Range(ws.Cells(CFG_R0, 1), ws.Cells(CFG_RN, P_NCOL)).Value
    For i = 1 To UBound(d, 1)
        If EhNumero(d(i, P_SMIN)) And EhNumero(d(i, P_SMAX)) Then
            If s >= CDbl(d(i, P_SMIN)) And s < CDbl(d(i, P_SMAX)) Then
                LinhaDoPlano = CFG_R0 + i - 1
                Exit Function
            End If
        End If
    Next i
End Function


' Um campo do plano de CQ recomendado para este Sigma.
'
'   campo  "CLASSE" | "REGRAS" | "N" | "RUNSIZE" | "FREQUENCIA" | "REFERENCIA"
'
' Devolve "" quando nao ha Sigma. Abaixo de 3 Sigma a tabela deixa N e run size
' VAZIOS de proposito: atribuir numero ali sugeriria que existe plano de CQ
' estatistico capaz de sustentar o metodo, e nao existe -- o caminho e
' investigar e melhorar o metodo.
Public Function PlanoQC(ByVal sigma As Variant, ByVal campo As String) As Variant
    PlanoQC = ""
    Dim lin As Long
    lin = LinhaDoPlano(sigma)
    If lin = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = AbaCfg()
    If ws Is Nothing Then Exit Function

    Dim c As Long
    Select Case UCase$(Trim$(campo))
        Case "CLASSE", "CLASSIFICACAO": c = P_CLASSE
        Case "REGRAS": c = P_REGRAS
        Case "N", "N_CONTROLE": c = P_N
        Case "RUNSIZE", "RUN_SIZE": c = P_RUNSIZE
        Case "FREQUENCIA", "FREQ": c = P_FREQ
        Case "REFERENCIA", "REF": c = P_REFER
        Case Else: Exit Function
    End Select

    ' A celula das faixas abaixo de 3 Sigma esta VAZIA de proposito, e uma UDF
    ' que devolve Empty e renderizada pela celula como ZERO. Sem esta conversao
    ' a tela mostrava "N = 0" e "Run Size = 0" -- que nao e "nao ha plano
    ' automatico", e sim "rode zero controles". Quarta vez que este projeto
    ' tropeca no mesmo Empty; por isso a conversao mora aqui, na saida.
    Dim v As Variant
    v = ws.Cells(lin, c).Value
    If IsEmpty(v) Or IsNull(v) Then
        PlanoQC = ""
    Else
        PlanoQC = v
    End If
End Function


' As regras recomendadas para este Sigma estao todas implementadas no motor?
'
' Devolve "TOTAL", ou "PARCIAL - falta ..." nomeando o que falta. NAO ajusta a
' recomendacao para caber no que o codigo sabe fazer: se um dia a tabela pedir
' uma regra que o motor nao avalia, quem muda e o rotulo de cobertura, nao a
' recomendacao. Fingir suporte seria a pior saida. Hoje a lista recomendada e
' a executada coincidem, e a funcao existe para o dia em que deixarem de
' coincidir.
Public Function CoberturaWestgard(ByVal sigma As Variant) As String
    CoberturaWestgard = ""
    ' Sem faixa nao ha o que cobrir: Sigma ausente devolve "", e nao um
    ' veredito. Com faixa mas sem regras -- a faixa abaixo de 3 Sigma --, a
    ' resposta e "nao aplicavel", que e diferente de "nao sei".
    If LinhaDoPlano(sigma) = 0 Then Exit Function

    Dim regras As Variant
    regras = PlanoQC(sigma, "REGRAS")
    ' Celula vazia volta do Excel como Empty, nao como String vazia: testar
    ' VarType = vbString descartaria justamente a faixa que precisa responder.
    If IsEmpty(regras) Or IsNull(regras) Then regras = ""
    If Trim$(CStr(regras)) = "" Then
        CoberturaWestgard = "não aplicável"
        Exit Function
    End If

    Dim tem As Object
    Set tem = CreateObject("Scripting.Dictionary")
    Dim p As Variant, x As Variant
    For Each x In Split(RegrasDoMotor(), ";")
        tem(UCase$(Trim$(CStr(x)))) = 1
    Next x

    Dim falta As String
    ' as regras vem separadas por "/" no formato de leitura do gestor
    For Each p In Split(Replace(CStr(regras), " ", ""), "/")
        If Trim$(CStr(p)) <> "" Then
            If Not tem.Exists(UCase$(Trim$(CStr(p)))) Then
                If falta <> "" Then falta = falta & ", "
                falta = falta & Trim$(CStr(p))
            End If
        End If
    Next p

    If falta <> "" Then
        CoberturaWestgard = "PARCIAL - falta " & falta
        Exit Function
    End If

    ' NOME BATENDO NAO E COBERTURA (ADR-044)
    '
    ' Foi assim que "plano diz 8x, motor conta 10" atravessou o projeto
    ' respondendo TOTAL: a conferencia comparava apenas os rotulos. Agora o
    ' contrato inteiro e conferido -- detector oficial, N, R e escopo.
    Dim erroCfg As String
    erroCfg = ValidacaoDaMatriz()
    If Len(erroCfg) > 0 Then
        CoberturaWestgard = erroCfg
        Exit Function
    End If

    Dim contrato As String
    contrato = ConferirContrato(CStr(regras), TabelaDoMotor())
    If Len(contrato) > 0 Then
        CoberturaWestgard = "ERRO DE COBERTURA - " & contrato
        Exit Function
    End If

    CoberturaWestgard = "TOTAL"
End Function




Public Function ContratoWestgard() As String
    ContratoWestgard = CONTRATO
End Function


' A tabela que o motor declara. Vinculo tardio: pasta sem o motor carregado
' devolve vazio e a cobertura diz que nao pode conferir, em vez de a pasta
' inteira deixar de compilar.
Private Function TabelaDoMotor() As String
    On Error GoTo semMotor
    TabelaDoMotor = CStr(Application.Run("DetectoresWestgard"))
    Exit Function
semMotor:
    TabelaDoMotor = ""
End Function


' Confere TODAS as regras que o plano pede contra a tabela informada.
' Publica para o QA poder passar uma tabela MUTADA e provar que reprova.
Public Function ConferirContrato(ByVal regras As String, ByVal tabela As String) As String
    If Len(Trim$(tabela)) = 0 Then
        ConferirContrato = "o motor nao publica a tabela de detectores"
        Exit Function
    End If

    Dim area As String
    On Error Resume Next
    area = CStr(Application.Run("AreaDoProduto"))
    On Error GoTo 0
    If Len(area) = 0 Then
        ConferirContrato = "o motor nao publica a area"
        Exit Function
    End If

    Dim p As Variant, x As Variant, c As Variant, erro As String
    For Each p In Split(Replace(CStr(regras), " ", ""), "/")
        If Len(Trim$(CStr(p))) > 0 Then
            Dim esperado As String
            esperado = ""
            For Each x In Split(CONTRATO, ";")
                c = Split(CStr(x), "|")
                If UCase$(CStr(c(0))) = UCase$(area) And _
                   UCase$(CStr(c(1))) = UCase$(Trim$(CStr(p))) Then
                    esperado = CStr(x)
                    Exit For
                End If
            Next x
            If Len(esperado) = 0 Then
                ConferirContrato = Trim$(CStr(p)) & ": sem contrato definido para " & area
                Exit Function
            End If
            c = Split(esperado, "|")
            erro = CStr(Application.Run("ConferirRegra", tabela, area, _
                        CStr(c(1)), CStr(c(2)), CLng(Val(CStr(c(3)))), _
                        CLng(Val(CStr(c(4)))), CStr(c(5))))
            If Len(erro) > 0 Then
                ConferirContrato = erro
                Exit Function
            End If
        End If
    Next p
End Function


' O nome da regra sequencial recomendada casa com o limiar que o motor usa?
Private Function DivergenciaDeLimiar(ByVal regras As String) As String
    Dim seq As Long, um As Long, r As String
    On Error GoTo semMotor
    seq = CLng(Application.Run("LimiarSequencialWestgard"))
    um = CLng(Application.Run("LimiarUmSigmaWestgard"))
    On Error GoTo 0

    r = Replace(LCase$(regras), " ", "")

    ' tendencia: o nome traz o proprio numero ("8x", "6x", "10x")
    If InStr(r, "/10x") > 0 Or Left$(r, 4) = "10x/" Or r = "10x" Then
        If seq <> 10 Then DivergenciaDeLimiar = "plano pede 10x e o motor conta " & seq
    ElseIf InStr(r, "8x") > 0 Then
        If seq <> 8 Then DivergenciaDeLimiar = "plano pede 8x e o motor conta " & seq
    ElseIf InStr(r, "6x") > 0 Then
        If seq <> 6 Then DivergenciaDeLimiar = "plano pede 6x e o motor conta " & seq
    End If
    If Len(DivergenciaDeLimiar) > 0 Then Exit Function

    ' regra de 1s: 4_1s exige quatro consecutivos, 3_1s exige tres
    If InStr(r, "4_1s") > 0 Then
        If um <> 4 Then DivergenciaDeLimiar = "plano pede 4_1s e o motor conta " & um
    ElseIf InStr(r, "3_1s") > 0 Then
        If um <> 3 Then DivergenciaDeLimiar = "plano pede 3_1s e o motor conta " & um
    End If
    Exit Function
semMotor:
    DivergenciaDeLimiar = "motor nao publica os limiares"
End Function


' A matriz ativa e coerente com o produto? Vinculo tardio pelo mesmo motivo
' de RegrasDoMotor: ausencia da funcao nao pode impedir a pasta de compilar.
Private Function ValidacaoDaMatriz() As String
    On Error GoTo semMotor
    ValidacaoDaMatriz = CStr(Application.Run("ValidarMatrizWestgard"))
    Exit Function
semMotor:
    ValidacaoDaMatriz = ""
End Function


' Quais regras o motor desta pasta avalia. Publicado para o BI e para a tela,
' para ninguem precisar deduzir do codigo.
Public Function RegrasImplementadas() As String
    RegrasImplementadas = Replace(RegrasDoMotor(), ";", " / ")
End Function


' Esta regra faz parte do plano recomendado para este Sigma?
'
' Existe para o BI publicar CINCO COLUNAS BOOLEANAS em vez de uma cadeia de
' texto. Uma medida DAX que procurasse "4_1s" dentro de
' "1_3s / 2_2s / R_4s / 4_1s / 8x" funcionaria hoje e quebraria no dia em que
' alguem escrevesse "4-1s" ou trocasse o separador -- e quebraria em silencio,
' devolvendo FALSE e apagando a regra da tela. O reconhecimento fica aqui, no
' mesmo modulo que produz a cadeia.
'
' Compara sobre a cadeia normalizada: separadores viram espaco e o traco vira
' sublinhado, entao "4-1s", "4_1s" e "4 1s" sao a mesma regra. A comparacao e
' por TOKEN INTEIRO -- sem isso "1_3s" casaria dentro de "11_3s".
Public Function RegraNoPlano(ByVal sigma As Variant, ByVal regra As String) As Boolean
    RegraNoPlano = False
    Dim s As String
    s = CStr(PlanoQC(sigma, "REGRAS"))
    If Len(Trim$(s)) = 0 Then Exit Function

    Dim alvo As String
    alvo = NormalizarRegra(regra)
    If Len(alvo) = 0 Then Exit Function

    Dim partes As Variant, i As Long
    partes = Split(NormalizarRegra(s), " ")
    For i = LBound(partes) To UBound(partes)
        If partes(i) = alvo Then
            RegraNoPlano = True
            Exit Function
        End If
    Next i
End Function


Private Function NormalizarRegra(ByVal t As String) As String
    Dim s As String
    s = LCase$(Trim$(t))
    s = Replace(s, "/", " ")
    s = Replace(s, ";", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, "-", "_")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizarRegra = Trim$(s)
End Function
