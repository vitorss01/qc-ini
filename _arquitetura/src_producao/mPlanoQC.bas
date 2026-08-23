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
' A REGRA DE SEQUENCIA DO PRODUTO E 8x.
'
' A familia 6x / 8x / 10x responde a mesma pergunta: quantos resultados
' consecutivos do mesmo lado da media denunciam desvio sistematico. Tabelas
' publicadas de Sigma rules trazem ora 6x, ora 8x, ora 10x -- escolher UMA e
' decisao operacional do laboratorio, e o QC_INI escolheu 8x. tblPlanoQC_Sigma
' recomenda 8x, o motor avalia 8x, e por isso a cobertura fecha TOTAL: os dois
' falam da mesma regra.
Private Const REGRAS_IMPLEMENTADAS As String = "1_3s;2_2s;R_4s;4_1s;8x"


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
' recomendacao. Fingir suporte seria a pior saida.
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
    For Each x In Split(REGRAS_IMPLEMENTADAS, ";")
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

    If falta = "" Then
        CoberturaWestgard = "TOTAL"
    Else
        CoberturaWestgard = "PARCIAL - falta " & falta
    End If
End Function


' Quais regras o motor desta pasta avalia. Publicado para o BI e para a tela,
' para ninguem precisar deduzir do codigo.
Public Function RegrasImplementadas() As String
    RegrasImplementadas = Replace(REGRAS_IMPLEMENTADAS, ";", " / ")
End Function
