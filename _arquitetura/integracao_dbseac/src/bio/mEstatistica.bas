Attribute VB_Name = "mEstatistica"
Option Explicit
' ============================================================================
'  MOTOR ESTATISTICO (Fase 3)
'  Toda a estatistica do sistema vive aqui. Nenhum UserForm e nenhuma planilha
'  calcula nada — o Excel guarda apenas o RESULTADO produzido por este modulo.
'
'  Elegibilidade (CLSI EP05 / C24): apenas resultados ELEGIVEIS compoem
'  media, DP, CV, Bias, Sigma e as regras de Westgard. Quem decide e o Power
'  Query (tblCQ_Final.PARTICIPA_ESTATISTICA, ADR-057); o motor so le. O
'  resultado inativado com REGISTRAR - LJ marcado vira X vermelho no grafico e
'  nunca entra em calculo.
' ============================================================================
Public Const NLV As Long = 2          ' niveis deste setor

Public Const EV_NCOL As Long = 14     ' colunas de Eventos_Westgard (ADR-045)
Private Const KC0 As Long = 3          ' Calc: 1a linha
Private Const NK  As Long = 180        ' Calc: nº de corridas exibiveis
Private Const CF0 As Long = 6          ' Calc: 1a coluna de nivel (F)
Private Const NFD As Long = 22         ' Calc: campos por nivel
Private Const EF0 As Long = 3          ' Eng_Saida: 1a coluna de bloco de nivel
Private Const NEF As Long = 7          ' Eng_Saida: campos por nivel
Private Const COL_FILTRO As Long = 24  ' Eng_Saida: filtro de data por corrida
Private Const COL_VALOR0 As Long = 25  ' Eng_Saida: 1a coluna de valor por nivel
Private Const COL_CHAVE As Long = 28   ' Eng_Saida: chave logica ANALITO|RUN
Private Const LINHA_STAT As Long = 185 ' Eng_Saida: 1a linha do bloco de estatistica
Private Const LINHA_EST As Long = 190  ' Eng_Saida: 1a linha da tabela de parametros
Private Const COL_DATA_ENG As Long = 29 ' Eng_Saida: data da corrida (AC) -- o Calc le daqui (ADR-050)
Private Const COL_X0 As Long = 30      ' Eng_Saida: 1o X vermelho (AD) -- ADR-057
Private Const NXN As Long = 3          ' X vermelhos por corrida e nivel
Private Const NX As Long = 9           ' colunas do bloco X (3 niveis x 3), iguais nos dois produtos
' ADR-070: de quem e cada ponto do LJ -- o ID_REGISTRO (o numero que o usuario digita na aba Inativar) e a
' DATA_HORA do resultado, por ponto. Mesmo arranjo das colunas de valor (um por nivel) e do bloco X (3 por
' nivel), com 3 niveis fixos nos dois produtos. Lidos por mUI.DicaPonto (dica do grafico).
Public Const ENG_COL_ID0 As Long = 39   ' Eng_Saida AM..AO: ID do ponto normal de N1..N3
Public Const ENG_COL_IDX0 As Long = 42  ' Eng_Saida AP..AX: ID de cada X vermelho (N1 X1..X3, N2 X1..X3, N3 X1..X3)
Public Const ENG_COL_DH0 As Long = 51   ' Eng_Saida AY..BA: DATA_HORA do ponto normal de N1..N3
Public Const ENG_COL_DHX0 As Long = 54  ' Eng_Saida BB..BJ: DATA_HORA de cada X vermelho
Public Const ENG_COL_ULT As Long = 62   ' Eng_Saida BJ: ultima coluna publicada por corrida
Public Const ENG_COL_VALOR0 As Long = 25 ' Eng_Saida Y: valor do N1 (mesmo que COL_VALOR0, para o mUI)
Public Const ENG_COL_X0 As Long = 30    ' Eng_Saida AD: 1o X (mesmo que COL_X0, para o mUI)
Private Const NIDV As Long = 3          ' colunas de ID/DATA_HORA do ponto normal (3 niveis fixos)
Private Const AR0 As Long = 4          ' Analitos: 1a linha
Private Const ARN As Long = 43
Private Const E0  As Long = 7          ' Estatistica: 1a linha

' ---- cache ----
Private mIdxX As Object                ' "ANALITO|LOTE" -> linhas X_VERMELHO de mDB (ADR-057)
Private mEquipIdx As String            ' equipamento para o qual mIdx/mIdxX foram montados ("" = todos)
Private mCache As Object               ' chave -> Array(n, media, dp)
Private mDB As Variant                 ' snapshot do banco
Private mDBok As Boolean
Private mAnal As Variant               ' snapshot de Analitos (alvos)
Private mIdxAnal As Object             ' analito -> linha em mAnal
Private mAgg As Object                 ' (analito|nivel) -> agregados de Westgard
Private mEvLote As String              ' lote cujos eventos estao em mAgg / Eventos_Westgard
Private mIdx As Object                 ' "ANALITO|LOTE" -> Collection de linhas ELEGIVEIS de mDB
' filtro do Painel, lido UMA vez por operacao (CarregarFiltro), nao uma vez por corrida
Private fDe As Double, fAte As Double

' ---- Westgard por detector (ADR-042) ----
' Declaracao de MODULO: em VBA tudo isto tem de vir antes da primeira
' procedure. Ficaram no meio do arquivo quando o motor foi substituido, e
' o resultado foi 'variavel nao definida' em gNaoAval.
Private Const DETECTORES As String = _
    "BIOQUIMICA|2|1_3s|INDIVIDUAL|1|1|1|WITHIN_RUN|3;" & _
    "BIOQUIMICA|2|2_2s|WITHIN_RUN|1|2|1|WITHIN_RUN_ACROSS_MATERIALS|2;" & _
    "BIOQUIMICA|2|2_2s|ACROSS_RUN_SAME_LEVEL|1|1|2|WITHIN_MATERIAL_ACROSS_RUN|2;" & _
    "BIOQUIMICA|2|R_4s|WITHIN_RUN|1|2|1|WITHIN_RUN_ACROSS_MATERIALS|4;" & _
    "BIOQUIMICA|2|4_1s|N2_R2|1|2|2|ACROSS_RUN_ACROSS_MATERIALS|1;" & _
    "BIOQUIMICA|2|4_1s|SAME_LEVEL_R4|0|1|4|WITHIN_MATERIAL_ACROSS_RUN|1;" & _
    "BIOQUIMICA|2|8x|N2_R4|1|2|4|ACROSS_RUN_ACROSS_MATERIALS|0;" & _
    "BIOQUIMICA|2|8x|SAME_LEVEL_R8|0|1|8|WITHIN_MATERIAL_ACROSS_RUN|0;" & _
    "HEMATOLOGIA|3|1_3s|INDIVIDUAL|1|1|1|WITHIN_RUN|3;" & _
    "HEMATOLOGIA|3|2of3_2s|WITHIN_RUN|1|3|1|WITHIN_RUN_ACROSS_MATERIALS|2;" & _
    "HEMATOLOGIA|3|2of3_2s|SAME_LEVEL_R3|0|1|3|WITHIN_MATERIAL_ACROSS_RUN|2;" & _
    "HEMATOLOGIA|3|R_4s|WITHIN_RUN|1|3|1|WITHIN_RUN_ACROSS_MATERIALS|4;" & _
    "HEMATOLOGIA|3|3_1s|N3_R1|1|3|1|WITHIN_RUN_ACROSS_MATERIALS|1;" & _
    "HEMATOLOGIA|3|3_1s|SAME_LEVEL_R3|0|1|3|WITHIN_MATERIAL_ACROSS_RUN|1;" & _
    "HEMATOLOGIA|3|6x|N3_R2|1|3|2|ACROSS_RUN_ACROSS_MATERIALS|0;" & _
    "HEMATOLOGIA|3|6x|SAME_LEVEL_R6|0|1|6|WITHIN_MATERIAL_ACROSS_RUN|0"

Private gTrace As Object            ' evidencias da ultima avaliacao
Private gDetMemo As Object          ' "ATIVO|regra|detector" / "ESC|regra|detector" -> resposta (ADR-050)
Private gNaoAval As Object          ' regra|detector -> janelas nao avaliaveis

' ============================ ELEGIBILIDADE (ADR-057) ============================
' Quem decide se um resultado participa da estatistica e o Power Query
' (tblCQ_Final.PARTICIPA_ESTATISTICA); mDados.CarregarDB traduz SIM para "Ativo".
' A aba Cfg_Status saiu: uma segunda regra de elegibilidade aqui seria uma
' populacao analitica paralela, exatamente o que a arquitetura proibe.

' Equipamento em analise. Bioquimica: o seletor do Painel (selEquipamento) --
' DIMENSION 1 e DIMENSION 2 sao series proprias de Levey-Jennings e Westgard.
' Hematologia: vazio = todos (um equipamento so).
Public Function EquipFiltro() As String
    On Error Resume Next
    EquipFiltro = UCase$(Trim$(CStr(ThisWorkbook.Names("selEquipamento").RefersToRange.Value)))
End Function

' ============================ CACHE ============================
Public Sub InvalidarCache()
    mLotes.InvalidarLotes
    mCEQ.InvalidarEQ
    mIncerteza.InvalidarIncerteza      ' ADR-064
    Set mCache = Nothing
    mDBok = False
    Set mIdxX = Nothing
    mEquipIdx = ""
    Set mIdxAnal = Nothing
    Set mAgg = Nothing
    mEvLote = ""
    Set mIdx = Nothing
    mAnal = Empty
End Sub

' Troca de lote em analise e edicao de media/DP: o BANCO nao mudou, entao o
' snapshot e o indice continuam valendo -- descarta-los custava ~1 s por troca
' com cinco anos de banco. Muda o que depende do parametro: os eventos de
' Westgard (mAgg) e o snapshot de Analitos.
Public Sub InvalidarParametros()
    mLotes.InvalidarLotes
    Set mAgg = Nothing
    mEvLote = ""
    Set mCache = Nothing
    If mDBok Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 1
    End If
End Sub

' ============================ INDICE (ADR-050) ============================
' Com cinco anos de banco (~100 mil linhas), varrer o banco inteiro a cada
' clique do spinner custava meio segundo so para achar as ~150 linhas do
' analito no lote. O indice e montado UMA vez por versao do banco (qualquer
' gravacao chama InvalidarCache) e cada consulta toca so as linhas do par.
' So entram linhas ELEGIVEIS: e o mesmo filtro que todas as consultas faziam.
Private Sub GarantirIndice()
    Dim i As Long, k As String, an As String, cod As String, eq As String
    GarantirDB
    eq = EquipFiltro()
    If Not mIdx Is Nothing Then
        If eq = mEquipIdx Then Exit Sub
        ' ADR-057: o equipamento do Painel mudou -- o indice, o cache de
        ' estatistica e os eventos de Westgard eram do outro equipamento
        Set mIdx = Nothing: Set mIdxX = Nothing: Set mAgg = Nothing: mEvLote = ""
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 1
    End If
    mEquipIdx = eq
    Set mIdx = CreateObject("Scripting.Dictionary")
    mIdx.CompareMode = 0                     ' lote exato; o analito vai em maiusculas
    Set mIdxX = CreateObject("Scripting.Dictionary")
    mIdxX.CompareMode = 0
    If IsEmpty(mDB) Then Exit Sub
    For i = 1 To UBound(mDB, 1)
        an = Trim$(CStr(mDB(i, COL_ANALITO)))
        If Len(an) > 0 Then
            If Len(eq) = 0 Or UCase$(Trim$(CStr(mDB(i, COL_EQUIP)))) = eq Then
                ' ADR-057: o lote ja vem como NUCLEO da tblCQ_Final
                cod = Trim$(CStr(mDB(i, COL_LOTE)))
                k = UCase$(an) & "|" & cod
                If mDB(i, COL_STATUS) = ST_ATIVO Then
                    If Not mIdx.Exists(k) Then mIdx.Add k, New Collection
                    mIdx(k).Add i
                ElseIf CStr(mDB(i, COL_PLOT)) = PLOT_X Then
                    ' inativado com REGISTRAR - LJ: so o X do grafico, nunca o calculo
                    If Not mIdxX.Exists(k) Then mIdxX.Add k, New Collection
                    mIdxX(k).Add i
                End If
            End If
        End If
    Next i
End Sub

' Linhas elegiveis de (analito, lote). Lote vazio = todos os lotes do analito.
Private Function LinhasDe(ByVal analito As String, ByVal lote As String) As Collection
    Dim k As Variant, pre As String, c As Collection, x As Variant
    GarantirIndice
    pre = UCase$(Trim$(analito)) & "|"
    If Len(lote) > 0 Then
        If mIdx.Exists(pre & lote) Then Set LinhasDe = mIdx(pre & lote) Else Set LinhasDe = New Collection
        Exit Function
    End If
    Set c = New Collection
    For Each k In mIdx.Keys
        If Left$(k, Len(pre)) = pre Then
            For Each x In mIdx(k)
                c.Add x
            Next x
        End If
    Next k
    Set LinhasDe = c
End Function

' Linhas X VERMELHO (inativadas com REGISTRAR - LJ) de (analito, lote) -- ADR-057.
' Nome com sufixo De, como LinhasDe: o VBA nao diferencia maiusculas, e uma
' variavel local "linhasX" esconderia uma funcao "LinhasX" (erro 450).
Private Function LinhasXDe(ByVal analito As String, ByVal lote As String) As Collection
    Dim k As String
    GarantirIndice
    k = UCase$(Trim$(analito)) & "|" & lote
    If mIdxX.Exists(k) Then Set LinhasXDe = mIdxX(k) Else Set LinhasXDe = New Collection
End Function

' Lote de uma chave do indice ("ANALITO|LOTE").
Private Function LoteDaChave(ByVal k As String) As String
    LoteDaChave = Mid$(k, InStrRev(k, "|") + 1)
End Function

Private Sub GarantirDB()
    Dim wsG As Worksheet, ig As Long, nmG As String
    If Not mDBok Then
        mDB = CarregarDB()
        ' snapshots em bloco — uma leitura por aba, nunca celula a celula
        Set wsG = ThisWorkbook.Sheets("Analitos")
        mAnal = wsG.Range(wsG.Cells(AR0, 1), wsG.Cells(ARN, 20)).Value  ' A..T
        Set mIdxAnal = CreateObject("Scripting.Dictionary")
        mIdxAnal.CompareMode = 1
        For ig = 1 To UBound(mAnal, 1)
            nmG = Trim$(CStr(mAnal(ig, 1)))
            If nmG <> "" Then
                If Not mIdxAnal.Exists(nmG) Then mIdxAnal.Add nmG, ig
            End If
        Next ig
        mDBok = True
    End If
    If mCache Is Nothing Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 1
    End If
End Sub

' ============================ PRIMITIVAS ============================
Public Function CalcularMedia(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double
    If n < 1 Then Exit Function
    For i = 1 To n
        s = s + v(i)
    Next i
    CalcularMedia = s / n
End Function

' DP amostral (n-1) — CLSI. Com n<2 nao ha dispersao estimavel: devolve 0.
Public Function CalcularDP(ByRef v() As Double, ByVal n As Long, ByVal media As Double) As Double
    Dim i As Long, s As Double
    If n < 2 Then Exit Function
    For i = 1 To n
        s = s + (v(i) - media) ^ 2
    Next i
    If s <= 0 Then Exit Function
    CalcularDP = Sqr(s / (n - 1))
End Function

Public Function CalcularCV(ByVal dp As Double, ByVal media As Double) As Double
    If media = 0 Then Exit Function
    CalcularCV = dp / media * 100#
End Function

Public Function CalcularBias(ByVal mediaObs As Double, ByVal alvo As Double) As Double
    If alvo = 0 Then Exit Function
    CalcularBias = (mediaObs - alvo) / alvo * 100#
End Function

Public Function CalcularErroTotal(ByVal cv As Double, ByVal bias As Double) As Double
    CalcularErroTotal = Abs(bias) + 1.65 * cv
End Function

Public Function CalcularSigma(ByVal etp As Double, ByVal bias As Double, ByVal cv As Double) As Double
    If cv = 0 Then Exit Function
    CalcularSigma = (etp - Abs(bias)) / cv
End Function

Public Function CalcularZ(ByVal valor As Double, ByVal media As Double, ByVal dp As Double) As Double
    If dp = 0 Then Exit Function
    CalcularZ = (valor - media) / dp
End Function

' A CLASSIFICACAO DE SIGMA MORA EM mQualidade (ADR-043)
'
' Existia aqui uma segunda escada, com TRES defeitos: >=6 devolvia
' "Excelente" em vez de "Classe mundial"; a faixa 5 a <6 nao existia, entao
' Sigma 5,5 caia em "Bom"; e abaixo de 3 dizia "Inaceitavel" enquanto o resto
' do projeto diz "Desempenho inadequado".
'
' Pior que os defeitos: ela vencia a canonica. AtualizarEstatisticaAba
' chamava ClassificarSigma SEM QUALIFICAR, e chamada nao qualificada resolve
' para a funcao do PROPRIO modulo -- entao a coluna de classificacao dos dois
' produtos vinha daqui, e nao de mQualidade, sem ninguem perceber.
'
' Removida. Quem precisa da classificacao chama mQualidade.ClassificarSigma,
' que le as faixas de Cfg_PlanoQC e cai numa escada de reserva identica.

' ============================ ALVOS (Analitos) ============================
Private Function LinhaAnalito(ByVal analito As String) As Long
    GarantirDB
    If mIdxAnal.Exists(Trim$(analito)) Then LinhaAnalito = mIdxAnal(Trim$(analito))
End Function

' Media/DP-alvo DO LOTE pedido, para o nivel (ADR-049).
'
' ANTES: lia Analitos!E:J -- a tela do lote carregado --, qualquer que fosse o
' lote dos resultados avaliados. Com o Painel num ano de outro lote, o z de um
' lote era calculado com a media/DP de outro. Numero plausivel, errado, calado.
' AGORA: a fonte e mLotes.ParametrosLote(lote, analito, nivel), pela chave.
' Devolve False quando o lote nao tem parametro valido: quem chama NAO calcula z.
'
' ETp: vem da coluna "ETp em uso final %", achada pelo CABECALHO. O indice fixo
' 18 apontava para "ETp VB" desde que as colunas de especificacao mudaram.
Public Function AlvoAnalito(ByVal analito As String, ByVal nivel As Long, _
                            ByRef alvoMedia As Double, ByRef alvoDP As Double, _
                            ByRef etp As Double, ByVal lote As String) As Boolean
    Dim r As Long, cE As Long
    alvoMedia = 0: alvoDP = 0: etp = 0
    r = LinhaAnalito(analito)
    If r > 0 Then
        cE = ColunaEtpEmUso()
        If cE > 0 And cE <= UBound(mAnal, 2) Then
            If IsNumeric(mAnal(r, cE)) And Not IsEmpty(mAnal(r, cE)) Then etp = CDbl(mAnal(r, cE))
        End If
    End If
    AlvoAnalito = mLotes.ParametrosLote(lote, analito, nivel, alvoMedia, alvoDP)
End Function

' Coluna de "ETp em uso final %" na aba Analitos, pelo rotulo da linha 3.
Private Function ColunaEtpEmUso() As Long
    Static c As Long
    Dim ws As Worksheet, j As Long, s As String
    If c > 0 Then ColunaEtpEmUso = c: Exit Function
    Set ws = ThisWorkbook.Sheets("Analitos")
    For j = 1 To 20
        s = UCase$(Replace(Trim$(CStr(ws.Cells(3, j).Value)), "  ", " "))
        If Left$(s, 10) = "ETP EM USO" Then c = j: Exit For
    Next j
    ' Hematologia: o mesmo papel se chama "ETp final %" (coluna R)
    If c = 0 Then
        For j = 1 To 20
            s = UCase$(Replace(Trim$(CStr(ws.Cells(3, j).Value)), "  ", " "))
            If Left$(s, 9) = "ETP FINAL" Then c = j: Exit For
        Next j
    End If
    If c = 0 Then c = IIf(NLV >= 3, 18, 20)  ' R (Hematologia) / T (Bioquimica), layout de 08/2026
    ColunaEtpEmUso = c
End Function

' Ordem CRONOLOGICA das corridas: data, e o RUN so desempata (ADR-049).
'
' O RUN e unico por (data, lote), mas NAO e cronologico: dado historico
' importado depois ganha RUN maior. No arquivo real, marco/2025 do lote 8974
' tem RUN 151..168 e junho/2026 tem RUN ~114 -- ordenar por RUN punha 2025
' DEPOIS de 2026, e as regras de sequencia (4_1s, 8x) liam uma serie fora do
' tempo. Um so comparador para o motor, os eventos e o BI.
Public Function CorridaAntes(ByVal d1 As Double, ByVal r1 As Long, _
                             ByVal d2 As Double, ByVal r2 As Long) As Boolean
    If d1 < d2 Then
        CorridaAntes = True
    ElseIf d1 = d2 Then
        CorridaAntes = (r1 < r2)
    End If
End Function

Public Sub OrdenarCorridas(ByRef runs() As Long, ByRef dts() As Double, ByVal n As Long)
    Dim a As Long, b As Long, tr As Long, td As Double
    For a = 2 To n
        tr = runs(a): td = dts(a): b = a - 1
        Do While b >= 1
            If Not CorridaAntes(td, tr, dts(b), runs(b)) Then Exit Do
            runs(b + 1) = runs(b): dts(b + 1) = dts(b): b = b - 1
        Loop
        runs(b + 1) = tr: dts(b + 1) = td
    Next a
End Sub

' ============================ ESTATISTICA BASICA ============================
' n/media/DP dos resultados ELEGIVEIS de (analito, nivel) no lote, com filtro de ano.
Public Function EstatBasica(ByVal analito As String, ByVal nivel As Long, ByVal loteCore As String, _
                            Optional ByVal anoDe As Long = 0, Optional ByVal anoAte As Long = 0) As Variant
    Dim k As String, i As Long, n As Long, v() As Double, media As Double, dp As Double
    GarantirDB
    k = UCase$(analito) & "|" & nivel & "|" & loteCore & "|" & anoDe & "|" & anoAte
    If mCache.Exists(k) Then EstatBasica = mCache(k): Exit Function
    If IsEmpty(mDB) Then
        EstatBasica = Array(0, 0, 0): Exit Function
    End If
    ' ADR-050: so as linhas elegiveis do par (analito, lote), pelo indice
    Dim linhas As Collection, x As Variant
    Set linhas = LinhasDe(analito, loteCore)
    ReDim v(1 To linhas.Count + 1)
    For Each x In linhas
        i = x
        If CLng(Val(mDB(i, COL_NIVEL))) = nivel Then
            If IsNumeric(mDB(i, COL_RESULT)) Then
                If anoDe = 0 Or (IsDate(mDB(i, COL_DATA)) And _
                   Year(mDB(i, COL_DATA)) >= anoDe And Year(mDB(i, COL_DATA)) <= anoAte) Then
                    n = n + 1: v(n) = CDbl(mDB(i, COL_RESULT))
                End If
            End If
        End If
    Next x
    media = CalcularMedia(v, n)
    dp = CalcularDP(v, n, media)
    EstatBasica = Array(n, media, dp)
    mCache(k) = EstatBasica
End Function

' ============================ WESTGARD ============================
' z(t, i) = z-score do nivel t na i-esima corrida (ordenada). ok(t,i) = tem dado.
' Regras padrao: 12s (alerta) e 13s/22s/R4s/41s/10x (rejeicao).
' ===========================================================================
' MOTOR WESTGARD POR DETECTOR E ESCOPO (ADR-042)
' ===========================================================================
'
' O QUE MUDOU, E POR QUE
'
' A versao anterior avaliava "a regra 8x" como oito CORRIDAS consecutivas do
' mesmo lado da media, num unico nivel. Isso e o detector LONGITUDINAL. A
' matriz Sigma da Bioquimica pede outra coisa: N=2, R=4 -- dois niveis por
' corrida, quatro corridas, oito OBSERVACOES. Sao janelas diferentes, com
' sensibilidade diferente, e a versao anterior chamava uma pelo nome da outra.
' O mesmo valia para 6x: a matriz de tres niveis pede N=3, R=2.
'
' ESCOPOS
'
'   WITHIN_RUN ......... so a corrida atual
'   ACROSS_RUN ......... usa corridas diferentes
'   WITHIN_MATERIAL .... o mesmo nivel ao longo do tempo
'   ACROSS_MATERIALS ... combina niveis diferentes
'
' R_4s e WITHIN_RUN e ponto. Amplitude de 4 DP entre a corrida de ontem e a
' de hoje NAO e R_4s -- e deriva, que outras regras pegam. Aplicar R_4s
' atraves de corridas inventa rejeicao que o metodo nao cometeu.
'
' OFICIAL x COMPLEMENTAR
'
' Detectores adicionais NAO entram automaticamente na decisao. Somar
' longitudinal com N/R multiplica as oportunidades de rejeicao e sobe a
' probabilidade de falsa rejeicao sem ninguem decidir isso. DetectorAtivo()
' e a unica porta: o que nao esta la e calculado, registrado e NAO consolidado.
'
' NAO_AVALIAVEL x FALSE
'
' FALSE quer dizer "avaliei e nao violou". Janela sem os dados que a regra
' exige nao e FALSE -- e ausencia de avaliacao, e some do QA se for tratada
' como aprovacao. 6x N3/R2 exige tres niveis validos nas DUAS corridas; com
' N3 faltando numa delas, a regra nao foi avaliada.

' A tabela de detectores. UMA fonte: o motor consulta, e Cfg_Westgard_Escopo
' e materializada a partir daqui -- planilha e codigo nao podem divergir.
'   Area|Niveis|Regra|Detector|Ativo|N|R|Escopo|Limiar
'
' Limiar 0 nas regras de deslocamento (8x, 6x): o corte e a MEDIA,
' nao um multiplo de DP. Deixar em branco confundiria com ausencia.



Public Function AreaDoProduto() As String
    If NLV >= 3 Then AreaDoProduto = "HEMATOLOGIA" Else AreaDoProduto = "BIOQUIMICA"
End Function


' A tabela inteira, para a planilha de configuracao e para o QA.
Public Function DetectoresWestgard() As String
    DetectoresWestgard = DETECTORES
End Function


' O contrato do motor para UMA regra, comparado contra a tabela informada.
' Devolve "" quando confere, ou a divergencia.
'
' Confere o que da para conferir DECLARATIVAMENTE: existencia do detector
' oficial, se esta ativo, N, R, escopo e limiar. NAO prova comportamento --
' isso e papel da suite de testes, e misturar as duas garantias faria a
' cobertura parecer mais forte do que e.
Public Function ConferirRegra(ByVal tabela As String, ByVal area As String, _
                              ByVal regra As String, ByVal detectorEsperado As String, _
                              ByVal nEsperado As Long, ByVal rEsperado As Long, _
                              ByVal escopoEsperado As String) As String
    Dim linha As String, c As Variant
    linha = DetectorOficialDe(tabela, area, regra)
    If Len(linha) = 0 Then
        ConferirRegra = regra & ": sem detector oficial ativo em " & area
        Exit Function
    End If
    c = Split(linha, "|")
    If UCase$(CStr(c(3))) <> UCase$(detectorEsperado) Then
        ConferirRegra = regra & ": detector oficial e " & CStr(c(3)) & _
                        ", esperado " & detectorEsperado
        Exit Function
    End If
    If CLng(Val(CStr(c(5)))) <> nEsperado Then
        ConferirRegra = regra & ": N=" & CStr(c(5)) & ", esperado " & CStr(nEsperado)
        Exit Function
    End If
    If CLng(Val(CStr(c(6)))) <> rEsperado Then
        ConferirRegra = regra & ": R=" & CStr(c(6)) & ", esperado " & CStr(rEsperado)
        Exit Function
    End If
    If Len(escopoEsperado) > 0 Then
        If UCase$(CStr(c(7))) <> UCase$(escopoEsperado) Then
            ConferirRegra = regra & ": escopo " & CStr(c(7)) & _
                            ", esperado " & escopoEsperado
        End If
    End If
End Function


' Este detector participa da DECISAO oficial deste produto?
Public Function DetectorAtivo(ByVal regra As String, ByVal detector As String) As Boolean
    Dim p As Variant, c As Variant, x As Variant, k As String
    ' ADR-050: a tabela e constante; a resposta tambem. Um lote de cinco meses
    ' gera ~8 mil evidencias, e cada uma redesmontava a tabela duas vezes.
    If gDetMemo Is Nothing Then Set gDetMemo = CreateObject("Scripting.Dictionary")
    k = "ATIVO|" & UCase$(Trim$(regra)) & "|" & UCase$(Trim$(detector))
    If gDetMemo.Exists(k) Then DetectorAtivo = gDetMemo(k): Exit Function
    DetectorAtivo = DetectorAtivoTabela(regra, detector)
    gDetMemo(k) = DetectorAtivo
End Function

Private Function DetectorAtivoTabela(ByVal regra As String, ByVal detector As String) As Boolean
    Dim p As Variant, c As Variant, x As Variant
    p = Split(DETECTORES, ";")
    For Each x In p
        c = Split(CStr(x), "|")
        If UCase$(CStr(c(0))) = AreaDoProduto() Then
            If UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And _
               UCase$(CStr(c(3))) = UCase$(Trim$(detector)) Then
                DetectorAtivoTabela = (CStr(c(4)) = "1")
                Exit Function
            End If
        End If
    Next x
End Function


' Metadata completa de um detector, na tabela informada.
'
' Recebe a TABELA como parametro para que o QA possa passar uma versao
' mutada e provar que a cobertura reprova um motor com detector errado. Sem
' isso os testes negativos seriam encenacao: nao existe como alterar uma
' Const em tempo de execucao.
'
' Devolve "" quando nao ha o detector; senao
'   Area|Niveis|Regra|Detector|Ativo|N|R|Escopo|Limiar
Public Function MetadataDetector(ByVal tabela As String, ByVal area As String, _
                                 ByVal regra As String, ByVal detector As String) As String
    Dim c As Variant, x As Variant
    For Each x In Split(tabela, ";")
        c = Split(CStr(x), "|")
        If UBound(c) >= 8 Then
            If UCase$(CStr(c(0))) = UCase$(Trim$(area)) And _
               UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And _
               UCase$(CStr(c(3))) = UCase$(Trim$(detector)) Then
                MetadataDetector = CStr(x)
                Exit Function
            End If
        End If
    Next x
End Function


' O detector OFICIAL de uma regra: o unico com Ativo=1 naquela area.
' Devolve "" se nao houver -- e isso e ERRO DE COBERTURA, nao ausencia banal.
Public Function DetectorOficialDe(ByVal tabela As String, ByVal area As String, _
                                  ByVal regra As String) As String
    Dim c As Variant, x As Variant
    For Each x In Split(tabela, ";")
        c = Split(CStr(x), "|")
        If UBound(c) >= 8 Then
            If UCase$(CStr(c(0))) = UCase$(Trim$(area)) And _
               UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And CStr(c(4)) = "1" Then
                DetectorOficialDe = CStr(x)
                Exit Function
            End If
        End If
    Next x
End Function


' N e R declarados para um detector, no formato "N=2 R=4".
Public Function EscalaDoDetector(ByVal regra As String, ByVal detector As String) As String
    Dim c As Variant, x As Variant, k As String
    If gDetMemo Is Nothing Then Set gDetMemo = CreateObject("Scripting.Dictionary")
    k = "ESC|" & UCase$(Trim$(regra)) & "|" & UCase$(Trim$(detector))
    If gDetMemo.Exists(k) Then EscalaDoDetector = gDetMemo(k): Exit Function
    For Each x In Split(DETECTORES, ";")
        c = Split(CStr(x), "|")
        If UCase$(CStr(c(0))) = AreaDoProduto() Then
            If UCase$(CStr(c(2))) = UCase$(Trim$(regra)) And _
               UCase$(CStr(c(3))) = UCase$(Trim$(detector)) Then
                EscalaDoDetector = "N=" & CStr(c(5)) & " R=" & CStr(c(6))
                gDetMemo(k) = EscalaDoDetector
                Exit Function
            End If
        End If
    Next x
End Function


' Uma evidencia = UM evento de Westgard. O par (niveis, runIni) nao e enfeite:
' e o que torna o trace utilizavel como FATO. Sem eles, quem consome so teria o
' texto livre de "detalhe" e precisaria adivinhar, por arqueologia de string,
' quais resultados o evento marcou -- exatamente o tipo de segunda leitura da
' regra que o ADR-042 existe para eliminar.
'
'   niveis  1-based, separados por virgula ("1" / "1,3" / "1,2,3")
'   runIni  primeira corrida da janela; = run nos detectores intra-corrida
Private Sub Evid(ByVal regra As String, ByVal detector As String, ByVal escopo As String, _
                 ByVal run As Long, ByVal niveis As String, ByVal runIni As Long, _
                 ByVal detalhe As String)
    If gTrace Is Nothing Then Set gTrace = New Collection
    Dim oficial As String
    If DetectorAtivo(regra, detector) Then oficial = "OFICIAL" Else oficial = "COMPLEMENTAR"
    gTrace.Add regra & "|" & detector & "|" & escopo & "|" & oficial & _
               "|run=" & CStr(run) & "|runIni=" & CStr(runIni) & _
               "|niveis=" & niveis & "|" & EscalaDoDetector(regra, detector) & _
               "|" & detalhe
End Sub


' Lista "1,2,3" com todos os niveis do produto.
Private Function TodosOsNiveis() As String
    Dim t As Long, s As String
    For t = 1 To NLV
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(t)
    Next t
    TodosOsNiveis = s
End Function


Private Sub NaoAvaliavel(ByVal regra As String, ByVal detector As String)
    If gNaoAval Is Nothing Then
        Set gNaoAval = CreateObject("Scripting.Dictionary")
        gNaoAval.CompareMode = 1
    End If
    Dim k As String
    k = regra & "|" & detector
    If gNaoAval.Exists(k) Then
        gNaoAval(k) = CLng(gNaoAval(k)) + 1
    Else
        gNaoAval.Add k, 1
    End If
End Sub


' As evidencias da ultima avaliacao, uma por linha.
Public Function TraceWestgard() As String
    ' ADR-050: Join, e nao s = s & x -- a concatenacao era quadratica no numero
    ' de evidencias (centenas por analito num lote longo). Mesmo texto.
    Dim a() As String, x As Variant, i As Long
    If gTrace Is Nothing Then Exit Function
    If gTrace.Count = 0 Then Exit Function
    ReDim a(0 To gTrace.Count - 1)
    For Each x In gTrace
        a(i) = CStr(x): i = i + 1
    Next x
    TraceWestgard = Join(a, vbLf)
End Function


' Quantas janelas ficaram sem avaliacao por falta de dado, por detector.
Public Function NaoAvaliaveisWestgard() As String
    Dim s As String, k As Variant
    If gNaoAval Is Nothing Then Exit Function
    For Each k In gNaoAval.Keys
        If Len(s) > 0 Then s = s & vbLf
        s = s & CStr(k) & "=" & CStr(gNaoAval(k))
    Next k
    NaoAvaliaveisWestgard = s
End Function


' Todos os niveis da corrida tem dado?
Private Function CorridaCompleta(ByRef td() As Boolean, ByVal i As Long) As Boolean
    Dim t As Long
    For t = 0 To NLV - 1
        If Not td(t, i) Then Exit Function
    Next t
    CorridaCompleta = True
End Function


Public Sub AvaliarWestgard(ByRef z() As Double, ByRef temDado() As Boolean, ByVal nRun As Long, _
                           ByRef r13 As Variant, ByRef r2sMulti As Variant, ByRef rR4 As Variant, _
                           ByRef r1sMulti As Variant, ByRef rSeq As Variant, ByRef a12 As Variant)
    Set gTrace = New Collection
    Set gNaoAval = CreateObject("Scripting.Dictionary")
    gNaoAval.CompareMode = 1

    Det_1_3s z, temDado, nRun, r13, a12
    Det_R_4s z, temDado, nRun, rR4

    If NLV >= 3 Then
        Det_2of3_2s_WithinRun z, temDado, nRun, r2sMulti
        Det_3_1s_N3R1 z, temDado, nRun, r1sMulti
        Det_6x_N3R2 z, temDado, nRun, rSeq
        ' Complementares: calculados e registrados, NAO consolidados.
        Det_MesmoNivel_Sequencia z, temDado, nRun, "2of3_2s", "SAME_LEVEL_R3", 3, 2#, r2sMulti
        Det_MesmoNivel_Sequencia z, temDado, nRun, "3_1s", "SAME_LEVEL_R3", 3, 1#, r1sMulti
        Det_MesmoNivel_Lado z, temDado, nRun, "6x", "SAME_LEVEL_R6", 6, rSeq
    Else
        Det_2_2s_WithinRun z, temDado, nRun, r2sMulti
        Det_2_2s_AcrossRun z, temDado, nRun, r2sMulti
        Det_4_1s_N2R2 z, temDado, nRun, r1sMulti
        Det_8x_N2R4 z, temDado, nRun, rSeq
        Det_MesmoNivel_Sequencia z, temDado, nRun, "4_1s", "SAME_LEVEL_R4", 4, 1#, r1sMulti
        Det_MesmoNivel_Lado z, temDado, nRun, "8x", "SAME_LEVEL_R8", 8, rSeq
    End If
End Sub


' --- 1_3s: INDIVIDUAL. Nao depende de outro nivel nem da corrida anterior.
Private Sub Det_1_3s(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                     ByRef r13 As Variant, ByRef a12 As Variant)
    Dim t As Long, i As Long
    For t = 0 To NLV - 1
        For i = 1 To nRun
            If td(t, i) Then
                If Abs(z(t, i)) > 2 Then a12(t, i) = 1
                If Abs(z(t, i)) > 3 Then
                    r13(t, i) = 1
                    Evid "1_3s", "INDIVIDUAL", "WITHIN_RUN", i, CStr(t + 1), i, _
                         "N" & CStr(t + 1) & " z=" & Format$(z(t, i), "0.00")
                End If
            End If
        Next i
    Next t
End Sub


' --- R_4s: WITHIN_RUN ONLY, todos os pares de niveis da MESMA corrida.
' Com tres niveis sao N1xN2, N1xN3 e N2xN3. O par fica registrado com o menor
' indice primeiro, para N1xN3 e N3xN1 nao virarem duas violacoes.
Private Sub Det_R_4s(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                     ByRef rR4 As Variant)
    Dim i As Long, t As Long, j As Long, comDado As Long
    For i = 1 To nRun
        comDado = 0
        For t = 0 To NLV - 1
            If td(t, i) Then comDado = comDado + 1
        Next t
        If comDado < 2 Then
            NaoAvaliavel "R_4s", "WITHIN_RUN"
            GoTo proximaRun
        End If
        For t = 0 To NLV - 2
            For j = t + 1 To NLV - 1
                If td(t, i) And td(j, i) Then
                    If Abs(z(t, i) - z(j, i)) > 4 Then
                        rR4(t, i) = 1
                        rR4(j, i) = 1
                        Evid "R_4s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, _
                             CStr(t + 1) & "," & CStr(j + 1), i, _
                             "N" & CStr(t + 1) & "xN" & CStr(j + 1) & _
                             " z=" & Format$(z(t, i), "0.00") & _
                             "/" & Format$(z(j, i), "0.00")
                    End If
                End If
            Next j
        Next t
proximaRun:
    Next i
End Sub


' --- 2_2s WITHIN_RUN (N=2): os dois niveis da corrida, mesmo lado, alem de 2s.
Private Sub Det_2_2s_WithinRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                               ByRef r2sMulti As Variant)
    Dim i As Long
    For i = 1 To nRun
        If Not CorridaCompleta(td, i) Then
            NaoAvaliavel "2_2s", "WITHIN_RUN"
        ElseIf (z(0, i) > 2 And z(1, i) > 2) Or (z(0, i) < -2 And z(1, i) < -2) Then
            r2sMulti(0, i) = 1: r2sMulti(1, i) = 1
            Evid "2_2s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, "1,2", i, _
                 "N1=" & Format$(z(0, i), "0.00") & " N2=" & Format$(z(1, i), "0.00")
        End If
    Next i
End Sub


' --- 2_2s ACROSS_RUN, MESMO nivel: duas corridas seguidas alem do mesmo 2s.
Private Sub Det_2_2s_AcrossRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                               ByRef r2sMulti As Variant)
    Dim t As Long, i As Long
    For t = 0 To NLV - 1
        For i = 2 To nRun
            If Not (td(t, i) And td(t, i - 1)) Then
                NaoAvaliavel "2_2s", "ACROSS_RUN_SAME_LEVEL"
            ElseIf (z(t, i) > 2 And z(t, i - 1) > 2) Or _
                   (z(t, i) < -2 And z(t, i - 1) < -2) Then
                r2sMulti(t, i) = 1
                Evid "2_2s", "ACROSS_RUN_SAME_LEVEL", "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - 1, _
                     "N" & CStr(t + 1) & " " & Format$(z(t, i - 1), "0.00") & _
                     "->" & Format$(z(t, i), "0.00")
            End If
        Next i
    Next t
End Sub


' --- 2of3_2s WITHIN_RUN: dois dos tres niveis da corrida alem do MESMO 2s.
' Lados opostos NAO contam -- isso e R_4s, fenomeno diferente.
Private Sub Det_2of3_2s_WithinRun(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                  ByRef r2sMulti As Variant)
    Dim i As Long, t As Long, acima As Long, abaixo As Long, comDado As Long
    For i = 1 To nRun
        acima = 0: abaixo = 0: comDado = 0
        For t = 0 To NLV - 1
            If td(t, i) Then
                comDado = comDado + 1
                If z(t, i) > 2 Then acima = acima + 1
                If z(t, i) < -2 Then abaixo = abaixo + 1
            End If
        Next t
        If comDado < 2 Then
            NaoAvaliavel "2of3_2s", "WITHIN_RUN"
        ElseIf acima >= 2 Or abaixo >= 2 Then
            Dim lado As Long, det As String, nivs As String
            If acima >= 2 Then lado = 1 Else lado = -1
            det = "": nivs = ""
            For t = 0 To NLV - 1
                If td(t, i) Then
                    If (lado = 1 And z(t, i) > 2) Or (lado = -1 And z(t, i) < -2) Then
                        r2sMulti(t, i) = 1
                        If Len(nivs) > 0 Then nivs = nivs & ","
                        nivs = nivs & CStr(t + 1)
                        det = det & "N" & CStr(t + 1) & "=" & Format$(z(t, i), "0.00") & " "
                    End If
                End If
            Next t
            Evid "2of3_2s", "WITHIN_RUN", "WITHIN_RUN_ACROSS_MATERIALS", i, nivs, i, Trim$(det)
        End If
    Next i
End Sub


' --- 3_1s N3/R1: os TRES niveis da mesma corrida alem do mesmo 1s.
Private Sub Det_3_1s_N3R1(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                          ByRef r1sMulti As Variant)
    Dim i As Long, t As Long, acima As Long, abaixo As Long
    For i = 1 To nRun
        If Not CorridaCompleta(td, i) Then
            NaoAvaliavel "3_1s", "N3_R1"
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For t = 0 To NLV - 1
            If z(t, i) > 1 Then acima = acima + 1
            If z(t, i) < -1 Then abaixo = abaixo + 1
        Next t
        If acima = NLV Or abaixo = NLV Then
            Dim det As String
            det = ""
            For t = 0 To NLV - 1
                r1sMulti(t, i) = 1
                det = det & "N" & CStr(t + 1) & "=" & Format$(z(t, i), "0.00") & " "
            Next t
            Evid "3_1s", "N3_R1", "WITHIN_RUN_ACROSS_MATERIALS", i, TodosOsNiveis(), i, Trim$(det)
        End If
prox:
    Next i
End Sub


' --- 4_1s N2/R2: dois niveis x duas corridas = quatro observacoes.
Private Sub Det_4_1s_N2R2(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                          ByRef r1sMulti As Variant)
    Dim i As Long, t As Long, k As Long, acima As Long, abaixo As Long
    For i = 2 To nRun
        If Not (CorridaCompleta(td, i) And CorridaCompleta(td, i - 1)) Then
            NaoAvaliavel "4_1s", "N2_R2"
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For k = i - 1 To i
            For t = 0 To NLV - 1
                If z(t, k) > 1 Then acima = acima + 1
                If z(t, k) < -1 Then abaixo = abaixo + 1
            Next t
        Next k
        If acima = NLV * 2 Or abaixo = NLV * 2 Then
            For t = 0 To NLV - 1
                r1sMulti(t, i) = 1
            Next t
            Evid "4_1s", "N2_R2", "ACROSS_RUN_ACROSS_MATERIALS", i, TodosOsNiveis(), i - 1, _
                 "runs " & CStr(i - 1) & ".." & CStr(i) & " " & CStr(NLV * 2) & " obs"
        End If
prox:
    Next i
End Sub


' --- 8x N2/R4: dois niveis x quatro corridas = oito observacoes do mesmo
' LADO DA MEDIA. E regra de deslocamento, nao de tendencia: nao exige que os
' valores crescam, so que fiquem do mesmo lado.
Private Sub Det_8x_N2R4(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                        ByRef rSeq As Variant)
    Det_JanelaNR z, td, nRun, "8x", "N2_R4", 4, rSeq
End Sub


' --- 6x N3/R2: tres niveis x duas corridas = seis observacoes do mesmo lado.
Private Sub Det_6x_N3R2(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                        ByRef rSeq As Variant)
    Det_JanelaNR z, td, nRun, "6x", "N3_R2", 2, rSeq
End Sub


' Janela de nRunsJanela corridas COMPLETAS: todas as NLV*nRunsJanela
' observacoes do mesmo lado da media.
Private Sub Det_JanelaNR(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                         ByVal regra As String, ByVal detector As String, _
                         ByVal nRunsJanela As Long, ByRef saida As Variant)
    Dim i As Long, k As Long, t As Long, acima As Long, abaixo As Long, total As Long
    total = NLV * nRunsJanela
    For i = nRunsJanela To nRun
        Dim completa As Boolean
        completa = True
        For k = i - nRunsJanela + 1 To i
            If Not CorridaCompleta(td, k) Then completa = False
        Next k
        If Not completa Then
            NaoAvaliavel regra, detector
            GoTo prox
        End If
        acima = 0: abaixo = 0
        For k = i - nRunsJanela + 1 To i
            For t = 0 To NLV - 1
                If z(t, k) > 0 Then acima = acima + 1
                If z(t, k) < 0 Then abaixo = abaixo + 1
            Next t
        Next k
        If acima = total Or abaixo = total Then
            For t = 0 To NLV - 1
                saida(t, i) = 1
            Next t
            Dim sinal As String
            If acima = total Then sinal = String$(total, "+") Else sinal = String$(total, "-")
            Evid regra, detector, "ACROSS_RUN_ACROSS_MATERIALS", i, _
                 TodosOsNiveis(), i - nRunsJanela + 1, _
                 "runs " & CStr(i - nRunsJanela + 1) & ".." & CStr(i) & " " & sinal
        End If
prox:
    Next i
End Sub


' --- COMPLEMENTAR: n corridas seguidas do mesmo lado, no MESMO nivel.
' Registrado e nao consolidado, salvo se DetectorAtivo disser o contrario.
Private Sub Det_MesmoNivel_Lado(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                ByVal regra As String, ByVal detector As String, _
                                ByVal n As Long, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > 0 Then acima = acima + 1
                If z(t, k) < 0 Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - n + 1, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub


' --- COMPLEMENTAR: n corridas seguidas alem do mesmo limite, no MESMO nivel.
Private Sub Det_MesmoNivel_Sequencia(ByRef z() As Double, ByRef td() As Boolean, ByVal nRun As Long, _
                                     ByVal regra As String, ByVal detector As String, _
                                     ByVal n As Long, ByVal limite As Double, ByRef saida As Variant)
    Dim t As Long, i As Long, k As Long, acima As Long, abaixo As Long
    Dim ativo As Boolean
    ativo = DetectorAtivo(regra, detector)
    For t = 0 To NLV - 1
        For i = n To nRun
            Dim ok As Boolean
            ok = True
            For k = i - n + 1 To i
                If Not td(t, k) Then ok = False
            Next k
            If Not ok Then
                NaoAvaliavel regra, detector
                GoTo prox
            End If
            acima = 0: abaixo = 0
            For k = i - n + 1 To i
                If z(t, k) > limite Then acima = acima + 1
                If z(t, k) < -limite Then abaixo = abaixo + 1
            Next k
            If acima = n Or abaixo = n Then
                If ativo Then saida(t, i) = 1
                Evid regra, detector, "WITHIN_MATERIAL_ACROSS_RUN", i, _
                     CStr(t + 1), i - n + 1, _
                     "N" & CStr(t + 1) & " runs " & CStr(i - n + 1) & ".." & CStr(i)
            End If
prox:
        Next i
    Next t
End Sub


' Os nomes da matriz vigente neste produto, na ordem das saidas de
' AvaliarWestgard. Existe para que Painel, Estatistica e a camada de BI leiam
' os rotulos do MESMO lugar que decide as regras -- uma lista escrita a mao
' noutro modulo divergiria no primeiro ajuste.
Public Function MatrizWestgard() As String
    If NLV >= 3 Then
        MatrizWestgard = "1_3s;2of3_2s;R_4s;3_1s;6x"
    Else
        MatrizWestgard = "1_3s;2_2s;R_4s;4_1s;8x"
    End If
End Function


' O nome da regra que ocupa uma posicao das saidas (1..5).
Public Function NomeRegraWestgard(ByVal pos As Long) As String
    Dim p As Variant
    p = Split(MatrizWestgard(), ";")
    If pos >= 1 And pos <= UBound(p) + 1 Then NomeRegraWestgard = p(pos - 1)
End Function


' A matriz ativa pertence mesmo a este produto? (ADR-041, secao 8)
'
' Devolve "" quando esta coerente, ou a descricao do problema. Existe porque a
' contaminacao entre modulos e silenciosa: uma matriz de dois niveis rodando
' num setor de tres continua produzindo numeros plausiveis -- so que com a
' sensibilidade e a taxa de falsa rejeicao do desenho errado.
'
' Verifica os DOIS sentidos. Faltar regra da propria familia e tao defeito
' quanto sobrar regra da outra.
Public Function ValidarMatrizWestgard() As String
    Dim m As String, faltando As String, intrusas As String
    Dim proprias As Variant, alheias As Variant, x As Variant

    m = ";" & MatrizWestgard() & ";"

    If NLV >= 3 Then
        proprias = Array("1_3s", "2of3_2s", "R_4s", "3_1s", "6x")
        alheias = Array("2_2s", "4_1s", "8x", "10x")
    Else
        proprias = Array("1_3s", "2_2s", "R_4s", "4_1s", "8x")
        alheias = Array("2of3_2s", "3_1s", "6x", "10x")
    End If

    For Each x In proprias
        If InStr(1, m, ";" & x & ";", vbTextCompare) = 0 Then _
            faltando = faltando & x & " "
    Next x
    For Each x In alheias
        If InStr(1, m, ";" & x & ";", vbTextCompare) > 0 Then _
            intrusas = intrusas & x & " "
    Next x

    If Len(faltando) > 0 Then
        ValidarMatrizWestgard = "ERRO DE CONFIGURACAO: faltam na matriz de " & _
            CStr(NLV) & " niveis: " & Trim$(faltando)
    End If
    If Len(intrusas) > 0 Then
        If Len(ValidarMatrizWestgard) > 0 Then ValidarMatrizWestgard = _
            ValidarMatrizWestgard & " | "
        ValidarMatrizWestgard = ValidarMatrizWestgard & _
            "ERRO DE CONFIGURACAO: regras de outro modulo na matriz de " & _
            CStr(NLV) & " niveis: " & Trim$(intrusas)
    End If
End Function


' O limiar que o motor REALMENTE usa para cada regra sequencial, publicado
' para a cobertura poder conferir semantica e nao so nome. Sem isto,
' "plano diz 8x, motor conta 10" volta a passar como TOTAL.
Public Function LimiarSequencialWestgard() As Long
    If NLV >= 3 Then LimiarSequencialWestgard = 6 Else LimiarSequencialWestgard = 8
End Function


' Quantos resultados consecutivos a regra de 1s exige neste produto
' (4_1s com dois niveis, 3_1s com tres).
Public Function LimiarUmSigmaWestgard() As Long
    If NLV >= 3 Then LimiarUmSigmaWestgard = 3 Else LimiarUmSigmaWestgard = 4
End Function


' O par LiberarEscrita/RestaurarProtecao mudou para mSeguranca (ADR-046).
'
' Ele nasceu aqui, mas a auditoria estatica achou oito pontos com o mesmo
' defeito em mAuditoria, mConfig, mLogDB, mRegistros e mImportar -- nenhum
' deles com relacao alguma com estatistica. Pior: a Imunologia ainda nao tem
' motor, entao um primitivo compartilhado morando aqui a deixaria sem ele.
'
' mSeguranca existe em todos os produtos e e a dona da protecao. Instalado no
' artefato por instalar_guarda_protecao.ps1 (mSeguranca e fonte de producao,
' ADR-021).

' ============================ MOTOR: ESCREVER Eng_Saida ============================
' Marco 2 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: esta rotina dava ClearContents na area inteira do Calc e reescrevia a
' matriz toda com valores, destruindo as 12.614 formulas da aba.
'
' AGORA: nao toca no Calc. Calcula z e Westgard e publica SO o veredicto em
' Eng_Saida. O Calc continua fazendo por formula o que a planilha faz bem --
' buscar, filtrar, ordenar e plotar -- e le as regras daqui.
'
' Motivo de existir: as formulas do Calc implementavam Westgard com definicoes
' diferentes das do motor (41s com 3 pontos em vez de 4, 10x com 6 em vez de 10,
' R4s como max>2 E min<-2 em vez de amplitude>4, e 22s so entre niveis). Duas
' camadas emitiam veredictos diferentes sobre o mesmo dado. Agora ha uma so.
'
' O nome da rotina foi mantido para nao quebrar as chamadas existentes em
' AtualizarEstatistica, RecalcularAnalitoAtual e mOperacao.
Public Sub AtualizarCalc()
    Dim ws As Worksheet, analito As String, lote As String, equip As String
    Dim i As Long, t As Long, nRun As Long, r As Long, e As Long
    Dim runs() As Long, dts() As Double, valor() As Double, temDado() As Boolean
    Dim r13 As Variant, r2sMulti As Variant, rR4 As Variant, r1sMulti As Variant, rSeq As Variant, a12 As Variant
    Dim alvoM() As Double, alvoS() As Double, etp As Double
    Dim seen As Object, ordem As Object
    Dim protEstava As Boolean

    Set ws = ThisWorkbook.Sheets("Eng_Saida")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))
    GarantirIndice
    CarregarFiltro
    ' ADR-049: o lote e o do PAINEL (loteSel), escolhido pelo usuario.
    lote = mLotes.LotePainel()
    ' ADR-057: na Bioquimica, o equipamento tambem (cada um e uma serie propria)
    equip = mEquipIdx

    ' Eng_Saida e aba tecnica e protegida. Ver o cabecalho de LiberarEscrita.
    On Error GoTo restaura
    protEstava = LiberarEscritaRapida(ws)   ' ADR-062

    ' limpa a area de saida (colunas B em diante; a coluna A guarda os slots
    ' fixos). Vai ate o bloco do X vermelho (ADR-057): sem isso, o X do analito
    ' anterior ficaria na tela.
    ' ADR-070: e as colunas de ID/DATA_HORA por ponto (ate BJ)
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + NK - 1, ENG_COL_ULT)).ClearContents
    ws.Range("C1").Value = analito
    ws.Range("E1").NumberFormat = "@"       ' lote "010" nao pode virar 10
    ws.Range("E1").Value = lote
    ws.Range("G1").Value = Now
    ws.Range("I1").Value = 0
    ' ADR-070 (revisao): nova publicacao -- a dica do LJ na tela pode ser de outro ponto agora
    mUI.LimparDicas
    ws.Range("K1").Value = 0
    ws.Range("M1").Value = 0
    ws.Range("O1").Value = ""
    ws.Range("Q1").Value = ""
    ws.Range("R1").Value = "equip:"
    ws.Range("S1").Value = equip
    ws.Range("T1").Value = "X além de " & NXN & ":"
    ws.Range("U1").Value = 0
    If analito = "" Or IsEmpty(mDB) Then GoTo restaura
    ws.Range("O1").Value = LotesNoPeriodo(analito)

    ' ---- descobrir as corridas do analito no lote ----
    ' DOIS conjuntos de linhas (ADR-057):
    '   linhas  = resultados ELEGIVEIS (PARTICIPA_ESTATISTICA = SIM): ponto,
    '             z, Westgard, media/DP/CV;
    '   linhasX = resultados INATIVADOS com REGISTRAR - LJ marcado
    '             (TIPO_PLOTAGEM_LJ = X_VERMELHO): so o X no grafico.
    ' O eixo de PLOTAGEM e a uniao das corridas dos dois (uma corrida em que o
    ' unico resultado do analito foi inativado ainda ganha posicao, com o X).
    ' O eixo do WESTGARD e so o das corridas com resultado elegivel -- o mesmo
    ' de antes da inativacao existir: o X nunca entra em regra, media ou DP.
    ' ADR-049: sem teto na descoberta. ADR-050: pelo indice.
    Dim linhas As Collection, linhasX As Collection, x As Variant
    Set linhas = LinhasDe(analito, lote)
    Set linhasX = LinhasXDe(analito, lote)    ' LinhasXDe, nao LinhasX: VBA nao diferencia maiusculas e a variavel esconderia a funcao
    Set seen = CreateObject("Scripting.Dictionary")
    Dim elegivel As Object: Set elegivel = CreateObject("Scripting.Dictionary")
    Dim nTodas As Long, runsT() As Long, dtsT() As Double, dd As Double, mxDe As Object, passada As Long
    ReDim runsT(1 To linhas.Count + linhasX.Count + 1): ReDim dtsT(1 To linhas.Count + linhasX.Count + 1)
    Set mxDe = CreateObject("Scripting.Dictionary")     ' RUN -> maior data/hora
    Dim fonte As Collection
    For passada = 1 To 2
        If passada = 1 Then Set fonte = linhas Else Set fonte = linhasX
        For Each x In fonte
            i = x
            dd = 0
            If IsDate(mDB(i, COL_DATA)) Then dd = CDbl(CDate(mDB(i, COL_DATA)))
            If passada = 1 Then elegivel(CStr(mDB(i, COL_RUN))) = True
            If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                nTodas = nTodas + 1
                runsT(nTodas) = CLng(Val(mDB(i, COL_RUN)))
                dtsT(nTodas) = dd
                seen.Add CStr(mDB(i, COL_RUN)), nTodas
                mxDe(CStr(mDB(i, COL_RUN))) = dd
            ElseIf dd > mxDe(CStr(mDB(i, COL_RUN))) Then
                ' data EXIBIDA = a maior data/hora da corrida no analito; a ORDEM
                ' continua pela data da 1a linha (ADR-049)
                mxDe(CStr(mDB(i, COL_RUN))) = dd
            End If
        Next x
    Next passada
    ws.Range("K1").Value = nTodas
    If nTodas = 0 Then GoTo restaura
    OrdenarCorridas runsT, dtsT, nTodas

    ' ---- janela: as NK corridas mais recentes ate o fim do periodo ----
    Dim fimJan As Long, iniJan As Long, nCort As Long, vAte As Variant
    fimJan = nTodas
    On Error Resume Next
    vAte = ThisWorkbook.Names("filtroAte").RefersToRange.Value
    On Error GoTo restaura
    If IsDate(vAte) Then
        Do While fimJan >= 1
            If Int(dtsT(fimJan)) <= Int(CDbl(CDate(vAte))) Then Exit Do
            fimJan = fimJan - 1
        Loop
    End If
    If fimJan < 1 Then GoTo restaura
    ' D10 (QA-ETL-001): a janela e das NK corridas ELEGIVEIS mais recentes. Corrida que so tem X nao consome
    ' posicao -- antes, com a janela cheia, inativar com o X empurrava a corrida elegivel mais antiga para fora
    ' e n/media/DP/CV/Sigma do Painel mudavam por causa do X. Corridas so-X dentro da janela entram enquanto
    ' houver posicao; faltando, saem as mais antigas (o X nunca tira lugar de resultado elegivel).
    Dim nEleg As Long, nSoX As Long, pularX As Long, j As Long
    iniJan = 1
    For i = fimJan To 1 Step -1
        If elegivel.Exists(CStr(runsT(i))) Then nEleg = nEleg + 1
        If nEleg >= NK Then iniJan = i: Exit For
    Next i
    If nEleg > NK Then nEleg = NK
    For i = iniJan To fimJan
        If Not elegivel.Exists(CStr(runsT(i))) Then nSoX = nSoX + 1
    Next i
    pularX = nSoX - (NK - nEleg)
    If pularX < 0 Then pularX = 0
    For i = 1 To iniJan - 1
        If PassaFiltro(dtsT(i)) Then nCort = nCort + 1
    Next i
    ws.Range("M1").Value = nCort + pularX
    ReDim runs(1 To NK): ReDim dts(1 To NK)
    nRun = 0
    j = 0
    For i = iniJan To fimJan
        If elegivel.Exists(CStr(runsT(i))) Then
            nRun = nRun + 1
            runs(nRun) = runsT(i): dts(nRun) = dtsT(i)
        ElseIf j < pularX Then
            j = j + 1                                  ' so-X mais antiga sem posicao
        Else
            nRun = nRun + 1
            runs(nRun) = runsT(i): dts(nRun) = dtsT(i)
        End If
    Next i
    If nRun < 1 Then GoTo restaura

    Set ordem = CreateObject("Scripting.Dictionary")
    For i = 1 To nRun
        ordem(CStr(runs(i))) = i
    Next i

    ' ---- coletar valores ELEGIVEIS por nivel (eixo de plotagem) ----
    ReDim valor(0 To NLV - 1, 1 To nRun)
    ReDim temDado(0 To NLV - 1, 1 To nRun)
    Dim outIdV() As Variant, outDhV() As Variant          ' ADR-070: de quem e cada ponto
    ReDim outIdV(1 To nRun, 1 To NIDV): ReDim outDhV(1 To nRun, 1 To NIDV)
    For Each x In linhas
        i = x
        If IsNumeric(mDB(i, COL_RESULT)) Then
            If ordem.Exists(CStr(mDB(i, COL_RUN))) Then
                t = CLng(Val(mDB(i, COL_NIVEL))) - 1
                If t >= 0 And t <= NLV - 1 Then
                    r = ordem(CStr(mDB(i, COL_RUN)))
                    valor(t, r) = CDbl(mDB(i, COL_RESULT))
                    temDado(t, r) = True
                    outIdV(r, t + 1) = CStr(mDB(i, COL_ID))
                    outDhV(r, t + 1) = mDB(i, COL_DATA)
                End If
            End If
        End If
    Next x

    ' ---- coletar os X vermelhos (ate NXN por corrida e nivel) ----
    Dim outX() As Variant, s As Long, nExc As Long
    Dim outIdX() As Variant, outDhX() As Variant           ' ADR-070: de quem e cada X
    ReDim outX(1 To nRun, 1 To NX)
    ReDim outIdX(1 To nRun, 1 To NX): ReDim outDhX(1 To nRun, 1 To NX)
    For Each x In linhasX
        i = x
        If IsNumeric(mDB(i, COL_RESULT)) Then
            If ordem.Exists(CStr(mDB(i, COL_RUN))) Then
                t = CLng(Val(mDB(i, COL_NIVEL))) - 1
                If t >= 0 And t <= NLV - 1 Then
                    r = ordem(CStr(mDB(i, COL_RUN)))
                    For s = 1 To NXN
                        If IsEmpty(outX(r, t * NXN + s)) Then
                            outX(r, t * NXN + s) = CDbl(mDB(i, COL_RESULT))
                            outIdX(r, t * NXN + s) = CStr(mDB(i, COL_ID))
                            outDhX(r, t * NXN + s) = mDB(i, COL_DATA)
                            Exit For
                        End If
                    Next s
                    If s > NXN Then nExc = nExc + 1
                End If
            End If
        End If
    Next x
    ws.Range("U1").Value = nExc

    ' ---- eixo do Westgard: so as corridas com resultado elegivel ----
    ' A sequencia avaliada e identica a de antes da inativacao existir: as
    ' corridas so-X ficam de fora, e o resultado inativado nao entra em z.
    Dim pDeE() As Long, nE As Long
    ReDim pDeE(1 To nRun)
    For i = 1 To nRun
        If elegivel.Exists(CStr(runs(i))) Then nE = nE + 1: pDeE(nE) = i
    Next i

    ' ---- alvos e z-scores (no eixo compactado) ----
    ' Nivel sem media/DP do lote NAO entra no Westgard: z = 0 seria lido como
    ' "no alvo" e o grafico sairia verde sobre um lote nunca configurado.
    Dim parOK() As Boolean, temAval() As Boolean, stPar As String
    Dim zc() As Double, tac() As Boolean
    ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1): ReDim parOK(0 To NLV - 1)
    ReDim temAval(0 To NLV - 1, 1 To nRun)
    ReDim r13(0 To NLV - 1, 1 To nRun): ReDim r2sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rR4(0 To NLV - 1, 1 To nRun): ReDim r1sMulti(0 To NLV - 1, 1 To nRun)
    ReDim rSeq(0 To NLV - 1, 1 To nRun): ReDim a12(0 To NLV - 1, 1 To nRun)
    For t = 0 To NLV - 1
        parOK(t) = AlvoAnalito(analito, t + 1, alvoM(t), alvoS(t), etp, lote)
        stPar = stPar & IIf(t > 0, ";", "") & "N" & (t + 1) & "=" & IIf(parOK(t), "OK", "SEM")
    Next t
    ws.Range("Q1").Value = stPar
    If nE > 0 Then
        ReDim zc(0 To NLV - 1, 1 To nE): ReDim tac(0 To NLV - 1, 1 To nE)
        For t = 0 To NLV - 1
            For e = 1 To nE
                i = pDeE(e)
                If temDado(t, i) And parOK(t) Then
                    zc(t, e) = CalcularZ(valor(t, i), alvoM(t), alvoS(t))
                    tac(t, e) = True
                    temAval(t, i) = True
                End If
            Next e
        Next t

        ' ---- Westgard ----
        Dim c13 As Variant, c2s As Variant, cR4 As Variant, c1s As Variant, cSeq As Variant, c12 As Variant
        ReDim c13(0 To NLV - 1, 1 To nE): ReDim c2s(0 To NLV - 1, 1 To nE)
        ReDim cR4(0 To NLV - 1, 1 To nE): ReDim c1s(0 To NLV - 1, 1 To nE)
        ReDim cSeq(0 To NLV - 1, 1 To nE): ReDim c12(0 To NLV - 1, 1 To nE)
        AvaliarWestgard zc, tac, nE, c13, c2s, cR4, c1s, cSeq, c12
        ' devolve as marcas ao eixo de plotagem
        For t = 0 To NLV - 1
            For e = 1 To nE
                i = pDeE(e)
                r13(t, i) = c13(t, e): r2sMulti(t, i) = c2s(t, e): rR4(t, i) = cR4(t, e)
                r1sMulti(t, i) = c1s(t, e): rSeq(t, i) = cSeq(t, e): a12(t, i) = c12(t, e)
            Next e
        Next t
    End If

    ' ---- publicar em Eng_Saida ----
    ' Sem filtro de data de proposito: as regras dependem da sequencia completa
    ' de corridas. Quem recorta o periodo exibido e o Calc, na hora de plotar.
    Dim outRun() As Variant, outLvl() As Variant, rej As Boolean
    ReDim outRun(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outRun(i, 1) = runs(i)
    Next i
    ws.Range(ws.Cells(KC0, 2), ws.Cells(KC0 + nRun - 1, 2)).Value = outRun

    Dim outFil() As Variant, outVal() As Variant, outChv() As Variant, outDt() As Variant
    ReDim outFil(1 To nRun, 1 To 1)
    ReDim outVal(1 To nRun, 1 To NLV)
    ReDim outChv(1 To nRun, 1 To 1)
    ReDim outDt(1 To nRun, 1 To 1)
    For i = 1 To nRun
        outFil(i, 1) = IIf(PassaFiltro(dts(i)), 1, 0)
        Dim dMx As Double
        dMx = mxDe(CStr(runs(i)))
        If dMx > 0 Then outDt(i, 1) = CDate(dMx) Else outDt(i, 1) = ""
        ' Chave logica da linha: ANALITO|RUN (o RUN e compartilhado entre analitos)
        outChv(i, 1) = analito & "|" & runs(i)
        For t = 0 To NLV - 1
            If temDado(t, i) Then outVal(i, t + 1) = valor(t, i) Else outVal(i, t + 1) = ""
        Next t
    Next i
    ws.Range(ws.Cells(KC0, COL_FILTRO), ws.Cells(KC0 + nRun - 1, COL_FILTRO)).Value = outFil
    ws.Range(ws.Cells(KC0, COL_VALOR0), ws.Cells(KC0 + nRun - 1, COL_VALOR0 + NLV - 1)).Value = outVal
    ws.Range(ws.Cells(KC0, COL_CHAVE), ws.Cells(KC0 + nRun - 1, COL_CHAVE)).Value = outChv
    ws.Range(ws.Cells(KC0, COL_DATA_ENG), ws.Cells(KC0 + nRun - 1, COL_DATA_ENG)).Value = outDt
    ' X vermelho: bloco proprio, FORA das colunas de valor (que o Painel soma em n/media/DP)
    ws.Range(ws.Cells(KC0, COL_X0), ws.Cells(KC0 + nRun - 1, COL_X0 + NX - 1)).Value = outX
    ' ADR-070: ID_REGISTRO e DATA_HORA de cada ponto (dica do grafico: mUI.DicaPonto)
    ws.Range(ws.Cells(KC0, ENG_COL_ID0), ws.Cells(KC0 + nRun - 1, ENG_COL_ID0 + NIDV - 1)).Value = outIdV
    ws.Range(ws.Cells(KC0, ENG_COL_IDX0), ws.Cells(KC0 + nRun - 1, ENG_COL_IDX0 + NX - 1)).Value = outIdX
    ws.Range(ws.Cells(KC0, ENG_COL_DH0), ws.Cells(KC0 + nRun - 1, ENG_COL_DH0 + NIDV - 1)).Value = outDhV
    ws.Range(ws.Cells(KC0, ENG_COL_DHX0), ws.Cells(KC0 + nRun - 1, ENG_COL_DHX0 + NX - 1)).Value = outDhX

    For t = 0 To NLV - 1
        ReDim outLvl(1 To nRun, 1 To NEF)
        For i = 1 To nRun
            If temDado(t, i) And Not temAval(t, i) Then
                ' resultado existe, parametro do lote nao: sem veredicto
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = "SEM PARAMETROS"
            ElseIf temDado(t, i) Then
                outLvl(i, 1) = IIf(r13(t, i) = 1, 1, 0)
                outLvl(i, 2) = IIf(r2sMulti(t, i) = 1, 1, 0)
                outLvl(i, 3) = IIf(rR4(t, i) = 1, 1, 0)
                outLvl(i, 4) = IIf(r1sMulti(t, i) = 1, 1, 0)
                outLvl(i, 5) = IIf(rSeq(t, i) = 1, 1, 0)
                outLvl(i, 6) = IIf(a12(t, i) = 1, 1, 0)
                rej = (r13(t, i) = 1 Or r2sMulti(t, i) = 1 Or rR4(t, i) = 1 Or r1sMulti(t, i) = 1 Or rSeq(t, i) = 1)
                ' So REJEITADO/OK. O alerta 12s vai na coluna propria.
                outLvl(i, 7) = IIf(rej, "REJEITADO", "OK")
            Else
                ' corrida so com X (ou nivel sem resultado): sem ponto, sem veredicto
                outLvl(i, 1) = 0: outLvl(i, 2) = 0: outLvl(i, 3) = 0
                outLvl(i, 4) = 0: outLvl(i, 5) = 0: outLvl(i, 6) = 0
                outLvl(i, 7) = ""
            End If
        Next i
        ws.Range(ws.Cells(KC0, EF0 + t * NEF), ws.Cells(KC0 + nRun - 1, EF0 + t * NEF + NEF - 1)).Value = outLvl
    Next t

    ws.Range("I1").Value = nRun
restaura:
    ' Restauro garantido: a saida antecipada e o erro passam por aqui.
    Dim nErrP As Long, sErrP As String
    nErrP = Err.Number: sErrP = Err.Description
    RestaurarProtecao ws, protEstava
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mEstatistica.AtualizarCalc", sErrP
End Sub

' Filtro de data/trimestre do Painel.
Private Function PassaFiltro(ByVal dserial As Double) As Boolean
    PassaFiltro = True
    If dserial <= 0 Then Exit Function
    ' ADR-049: por DIA. Resultado gravado com hora (a importacao de 2025 veio
    ' com 03:00) ficava fora do ultimo dia do periodo: 29/05 03:00 > 29/05 00:00.
    ' mEstatPeriodo ja truncava; o Painel nao.
    If fDe > 0 Then If Int(dserial) < fDe Then PassaFiltro = False: Exit Function
    If fAte > 0 Then If Int(dserial) > fAte Then PassaFiltro = False: Exit Function
End Function

' Le o filtro do Painel (De/Ate) para as variaveis do modulo.
' ADR-050: antes cada chamada de PassaFiltro relia seis nomes da pasta --
' milhares de leituras por clique num lote longo.
' ADR-054: as quatro caixas de trimestre (qsel1..4) sairam. Eram um SEGUNDO
' caminho de filtragem, paralelo a De/Ate: com "2o Tri" marcado e T1 escolhido
' na lista, o Painel ficava vazio e nao dizia por que. O periodo agora tem um
' dono so -- Ano + Periodo escrevem De/Ate, e o motor le De/Ate.
Private Sub CarregarFiltro()
    Dim de As Variant, ate As Variant
    fDe = 0: fAte = 0
    On Error Resume Next
    de = ThisWorkbook.Names("filtroDe").RefersToRange.Value
    ate = ThisWorkbook.Names("filtroAte").RefersToRange.Value
    On Error GoTo 0
    If IsDate(de) Then fDe = Int(CDbl(CDate(de)))
    If IsDate(ate) Then fAte = Int(CDbl(CDate(ate)))
End Sub

' Repeticoes/calibracao vindas da aba Registros (marcadores do grafico).


' ============================================================================
'  VIOLACOES DE WESTGARD — estrutura completa e explicavel
'  Escopo revisado: sem Trend/Slope/Deriva/Shift (pertencem a um modulo futuro
'  de Analytics). O comportamento sistematico e capturado por 22s, 41s e 10x.
'
'  Devolve, para (analito, nivel): a ULTIMA violacao da serie elegivel com
'  Regra | Classificacao | RUN | Analito | Nivel | Resultado | Media | DP | Z | Data
'  Referencia: Westgard JO, Barry PL, Hunt MR, Groth T. Clin Chem 1981;27:493-501.
'  Elegibilidade conforme CLSI C24 e EP05 (ver mEstatistica.EhElegivel).
' ============================================================================
Public Function AnalisarViolacoes(ByVal analito As String, ByVal nivel As Long, ByVal loteCore As String) As Variant
    Dim i As Long, j As Long, nS As Long
    Dim ys() As Double, rs() As Long, ds() As Double
    Dim media As Double, dp As Double, alvoM As Double, alvoS As Double, etp As Double
    Dim seen As Object, regra As String, runV As Long, valV As Double, zV As Double, dtv As Double

    GarantirDB
    ' n, media, dp, regra, RUN, valor, z, dataDeteccao, totalViolacoes
    AnalisarViolacoes = Array(0, 0, 0, "", 0, 0, 0, 0, 0)
    If IsEmpty(mDB) Then Exit Function

    Dim linhas As Collection, x As Variant
    Set linhas = LinhasDe(analito, loteCore)       ' ADR-050: indice
    ReDim ys(1 To linhas.Count + 1): ReDim rs(1 To linhas.Count + 1): ReDim ds(1 To linhas.Count + 1)
    Set seen = CreateObject("Scripting.Dictionary")
    For Each x In linhas
        i = x
        If CLng(Val(mDB(i, COL_NIVEL))) = nivel Then
            If IsNumeric(mDB(i, COL_RESULT)) Then
                If Not seen.Exists(CStr(mDB(i, COL_RUN))) Then
                    seen.Add CStr(mDB(i, COL_RUN)), 1
                    nS = nS + 1
                    rs(nS) = CLng(Val(mDB(i, COL_RUN)))
                    ys(nS) = CDbl(mDB(i, COL_RESULT))
                    If IsDate(mDB(i, COL_DATA)) Then ds(nS) = CDbl(CDate(mDB(i, COL_DATA)))
                End If
            End If
        End If
    Next x
    If nS = 0 Then Exit Function

    Dim tr As Long, ty As Double, td2 As Double
    For i = 2 To nS
        tr = rs(i): ty = ys(i): td2 = ds(i): j = i - 1
        Do While j >= 1
            If Not CorridaAntes(td2, tr, ds(j), rs(j)) Then Exit Do
            rs(j + 1) = rs(j): ys(j + 1) = ys(j): ds(j + 1) = ds(j): j = j - 1
        Loop
        rs(j + 1) = tr: ys(j + 1) = ty: ds(j + 1) = td2
    Next i

    Dim vv() As Double
    ReDim vv(1 To nS)
    For i = 1 To nS
        vv(i) = ys(i)
    Next i
    media = CalcularMedia(vv, nS)
    dp = CalcularDP(vv, nS, media)
    If Not AlvoAnalito(analito, nivel, alvoM, alvoS, etp, loteCore) Then Exit Function

    ' --- avaliar Westgard sobre a serie ---
    Dim zz() As Double, tdd() As Boolean, tot As Long
    Dim q13 As Variant, q2sMulti As Variant, qR4 As Variant, q1sMulti As Variant, qSeq As Variant, q12 As Variant
    ReDim zz(0 To NLV - 1, 1 To nS): ReDim tdd(0 To NLV - 1, 1 To nS)
    For i = 1 To nS
        If alvoS > 0 Then zz(nivel - 1, i) = CalcularZ(ys(i), alvoM, alvoS)
        tdd(nivel - 1, i) = True
    Next i
    ReDim q13(0 To NLV - 1, 1 To nS): ReDim q2sMulti(0 To NLV - 1, 1 To nS)
    ReDim qR4(0 To NLV - 1, 1 To nS): ReDim q1sMulti(0 To NLV - 1, 1 To nS)
    ReDim qSeq(0 To NLV - 1, 1 To nS): ReDim q12(0 To NLV - 1, 1 To nS)
    AvaliarWestgard zz, tdd, nS, q13, q2sMulti, qR4, q1sMulti, qSeq, q12

    For i = 1 To nS
        Dim rg As String
        rg = ""
        If Val(q13(nivel - 1, i)) = 1 Then rg = "13s"
        If Val(q2sMulti(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "22s", rg & "+22s")
        If Val(q1sMulti(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "41s", rg & "+41s")
        If Val(qSeq(nivel - 1, i)) = 1 Then rg = IIf(rg = "", "10x", rg & "+10x")
        If rg <> "" Then
            tot = tot + 1
            regra = rg: runV = rs(i): valV = ys(i): zV = zz(nivel - 1, i): dtv = ds(i)
        End If
    Next i

    AnalisarViolacoes = Array(nS, media, dp, regra, runV, valV, zV, dtv, tot)
End Function

' Bloco descritivo completo da ultima violacao — consulta o modulo de conhecimento.
Public Function DetalheViolacao(ByVal analito As String, ByVal nivel As Long) As String
    Dim a As Variant, r As String, primeira As String
    a = AnalisarViolacoes(analito, nivel, mLotes.LotePainel())
    If CStr(a(3)) = "" Then
        DetalheViolacao = "Sem violação de Westgard na série elegível."
        Exit Function
    End If
    primeira = Split(CStr(a(3)), "+")(0)
    r = "? Regra: " & a(3) & vbCrLf
    r = r & "Classificação: " & RegraClassificacao(primeira) & vbCrLf
    r = r & "RUN: " & a(4) & vbCrLf
    r = r & "Analito: " & analito & vbCrLf
    r = r & "Nível: " & nivel & vbCrLf
    r = r & "Resultado: " & Format(a(5), "0.0000") & vbCrLf
    r = r & "Média: " & Format(a(1), "0.0000") & vbCrLf
    r = r & "DP: " & Format(a(2), "0.0000") & vbCrLf
    r = r & "Z-Score: " & Format(a(6), "+0.00;-0.00") & vbCrLf
    If a(7) > 0 Then r = r & "Data/Hora da detecção: " & Format(CDate(a(7)), "dd/mm/yyyy") & vbCrLf
    r = r & vbCrLf & "Interpretação:" & vbCrLf & RegraInterpretacao(primeira) & vbCrLf
    r = r & vbCrLf & "Prováveis causas:" & vbCrLf & RegraCausas(primeira) & vbCrLf
    r = r & vbCrLf & "Sugestões:" & vbCrLf & RegraSugestoes(primeira)
    DetalheViolacao = r
End Function

' ============================ PAINEL ============================
' ======================= MOTOR: ESTATISTICA POR NIVEL =======================
' Marco 3 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: lia o Calc, calculava n/media/DP/CV/bias/ET/Sigma e gravava direto nas
' celulas do Painel, destruindo 45 das 48 formulas da aba.
'
' AGORA: le Eng_Saida (nunca o Calc) e publica em Eng_Saida. O Painel virou
' camada de apresentacao pura: le por formula e nao calcula mais nada.
'
' Por que parou de ler o Calc: depois do Marco 2 as colunas de regra do Calc sao
' formulas que apontam para Eng_Saida. Se o motor as lesse, passaria a depender
' de o Excel ter recalculado a planilha antes -- com Calculation manual, ou com
' a rotina chamada em sequencia dentro de outra, leria valor velho sem avisar.
' Lendo Eng_Saida, le o que o proprio motor acabou de escrever.
'
' Correcao de conta que vem junto (as formulas do Painel estavam erradas):
'   Erro Total  antes 1,65*CV + bias        agora Abs(bias) + 1,65*CV
'   Sigma       antes (ETp - bias)/CV       agora (ETp - Abs(bias))/CV
' Sem o Abs, bias negativo -- media abaixo do alvo, situacao corriqueira --
' subestimava o Erro Total e superestimava o Sigma. CLSI/Westgard: TE = |bias| +
' 1,65*CV. O motor ja estava certo; a interface e que discordava dele.
Public Sub AtualizarPainelEng()
    Dim eng As Worksheet, analito As String
    Dim t As Long, i As Long, n As Long, v() As Double, media As Double, dp As Double
    Dim cv As Double, bias As Double, et As Double, sg As Double, etp As Double
    ' alvoM/alvoS, nunca alvoM/alvoS: VBA e insensivel a maiusculas, entao "alvoS" e o
    ' mesmo token que a palavra reservada "As" e o modulo nao compila.
    Dim alvoM As Double, alvoS As Double, rejTot As Long
    Dim cnt(1 To 5) As Long, k2 As Long
    Dim flags As Variant, vals As Variant, filtros As Variant
    Dim outStat() As Variant

    ' O motor nao referencia mais a aba Painel, nem para leitura.
    Set eng = ThisWorkbook.Sheets("Eng_Saida")
    analito = Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value))

    ' uma leitura em bloco de tudo que veio do motor
    filtros = eng.Range(eng.Cells(KC0, COL_FILTRO), eng.Cells(KC0 + NK - 1, COL_FILTRO)).Value
    vals = eng.Range(eng.Cells(KC0, COL_VALOR0), eng.Cells(KC0 + NK - 1, COL_VALOR0 + NLV - 1)).Value
    flags = eng.Range(eng.Cells(KC0, EF0), eng.Cells(KC0 + NK - 1, EF0 + NLV * NEF - 1)).Value

    ReDim outStat(1 To NLV, 1 To 21)

    For t = 0 To NLV - 1
        n = 0: ReDim v(1 To NK)
        For i = 1 To 5
            cnt(i) = 0
        Next i
        rejTot = 0

        For i = 1 To NK
            If Val(filtros(i, 1)) = 1 Then
                If IsNumeric(vals(i, t + 1)) And Trim$(CStr(vals(i, t + 1))) <> "" Then
                    n = n + 1: v(n) = CDbl(vals(i, t + 1))
                End If
                ' campos 1..5 do bloco do nivel = R13s, R22s, RR4s, R41s, R10x
                For k2 = 1 To 5
                    If Val(flags(i, t * NEF + k2)) = 1 Then cnt(k2) = cnt(k2) + 1
                Next k2
                ' campo 7 = Veredicto
                If CStr(flags(i, t * NEF + 7)) = "REJEITADO" Then rejTot = rejTot + 1
            End If
        Next i

        media = CalcularMedia(v, n)
        dp = CalcularDP(v, n, media)
        cv = CalcularCV(dp, media)
        AlvoAnalito analito, t + 1, alvoM, alvoS, etp, mLotes.LotePainel()
        bias = CalcularBias(media, alvoM)
        et = CalcularErroTotal(cv, bias)
        sg = CalcularSigma(etp, bias, cv)

        outStat(t + 1, 2) = n
        outStat(t + 1, 3) = IIf(n = 0, "", media)
        outStat(t + 1, 4) = IIf(n < 2, "", dp)
        outStat(t + 1, 5) = IIf(n < 2 Or media = 0, "", cv)
        ' ADR-068: sem ETp cadastrado (vazio ou <= 0) ETp e Sigma saem "-" -- nunca 0 nem Sigma negativo
        outStat(t + 1, 6) = IIf(etp <= 0, "-", etp)
        outStat(t + 1, 7) = IIf(n = 0 Or alvoM = 0, "", bias)
        outStat(t + 1, 8) = IIf(n < 2 Or media = 0, "", et)
        outStat(t + 1, 9) = IIf(etp <= 0, "-", IIf(n < 2 Or cv = 0, "", sg))
        outStat(t + 1, 10) = IIf(n = 0, "", IIf(rejTot > 0, "REJEITADO", "OK"))
        outStat(t + 1, 11) = ""
        outStat(t + 1, 12) = ""
        For i = 1 To 5
            outStat(t + 1, 12 + i) = cnt(i)
        Next i
        outStat(t + 1, 18) = cnt(1) + cnt(2) + cnt(3) + cnt(4) + cnt(5)

        ' resumo: o historico completo fica na aba Eventos_Westgard
        Dim aa As Variant
        aa = AgregadoWestgard(analito, t + 1)
        If CLng(aa(0)) = 0 Then
            outStat(t + 1, 19) = Chr$(151)
            outStat(t + 1, 20) = ""
            outStat(t + 1, 21) = ""
        Else
            outStat(t + 1, 19) = aa(3) & " " & Chr$(183) & " RUN " & aa(4)
            outStat(t + 1, 20) = RegraClassificacao(Split(CStr(aa(3)), "+")(0))
            ' ADR-045: "3x" era ambiguo -- 3 eventos ou 3 resultados fora?
            ' Um 6x N3/R2 e 1 evento e 6 resultados marcados. Mostrar os dois.
            outStat(t + 1, 21) = aa(0) & " ev / " & aa(6) & " res  |  maior Z " & _
                                 Format(aa(5), "+0.00;-0.00")
        End If
    Next t

    ' escrita unica do bloco A..U; a coluna A repoe o proprio numero do nivel
    For t = 1 To NLV
        outStat(t, 1) = t
    Next t
    Dim protEstava As Boolean
    On Error GoTo restaura
    protEstava = LiberarEscritaRapida(eng)   ' ADR-062
    eng.Range(eng.Cells(LINHA_STAT, 1), eng.Cells(LINHA_STAT + NLV - 1, 21)).Value = outStat

restaura:
    Dim nErrP As Long, sErrP As String
    nErrP = Err.Number: sErrP = Err.Description
    RestaurarProtecao eng, protEstava
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mEstatistica.AtualizarPainelEng", sErrP
End Sub

' ============================ ABA ESTATISTICA ============================
' ========================= MOTOR: TABELA DE PARAMETROS =========================
' Marco 4 do Sprint HARDENING 1 (ADR-019).
'
' ANTES: gravava direto em Estatistica!C7:M126, destruindo as 1.320 formulas
' daquele bloco. Nao rodava na producao porque o procedimento nao compilava
' (identificador alvoS colidindo com a palavra reservada As) -- e era exatamente
' isso que mantinha as formulas vivas la. Com o alvoS corrigido, passaria a rodar
' e a destruir a aba; por isso o desvio para Eng_Saida vem junto da correcao.
'
' AGORA: publica em Eng_Saida (linhas 190..309, colunas C..M) e a aba
' Estatistica le por INDEX(engEstat, linha, coluna).
'
' Continua lendo B3/D3/B4 da propria aba Estatistica: sao filtros que o usuario
' digita (ano inicial, ano final, lote), entrada de interface e nao calculo.
Public Sub AtualizarEstatisticaAba()
    Dim ws As Worksheet, wa As Worksheet, eng As Worksheet, i As Long, t As Long, er As Long
    Dim analito As String, anoDe As Long, anoAte As Long, loteF As String
    Dim st As Variant, n As Long, media As Double, dp As Double, cv As Double
    ' alvoM/alvoS, nunca alvoM/alvoS: VBA e insensivel a maiusculas, entao "alvoS" e o
    ' mesmo token que a palavra reservada "As" e o modulo nao compila.
    Dim alvoM As Double, alvoS As Double, etp As Double, bias As Double, et As Double, sg As Double
    Dim outp() As Variant, linhas As Long
    Dim protEstava As Boolean

    Set ws = ThisWorkbook.Sheets("Estatística")
    Set wa = ThisWorkbook.Sheets("Analitos")
    Set eng = ThisWorkbook.Sheets("Eng_Saida")

    Dim visao As String
    visao = UCase$(Trim$(CStr(ws.Range("B4").Value)))
    loteF = Trim$(CStr(ws.Range("H4").Value))
    If visao = "PERSONALIZADO" Then
        Dim dIniPer As Variant, dFimPer As Variant
        dIniPer = ws.Range("B5").Value
        dFimPer = ws.Range("D5").Value
        If IsDate(dIniPer) Then anoDe = Year(CDate(dIniPer)) Else anoDe = 0
        If IsDate(dFimPer) Then anoAte = Year(CDate(dFimPer)) Else anoAte = anoDe
    Else
        anoDe = CLng(Val(ws.Range("D4").Value))
        anoAte = anoDe
    End If
    If anoAte < anoDe Then anoAte = anoDe

    linhas = 40 * NLV
    ReDim outp(1 To linhas, 1 To 11)     ' C..M
    er = 0
    For i = AR0 To ARN
        analito = Trim$(CStr(wa.Cells(i, 1).Value))
        For t = 1 To NLV
            er = er + 1
            If analito = "" Then
                Dim c As Long
                For c = 1 To 11
                    outp(er, c) = ""
                Next c
            Else
                st = EstatBasica(analito, t, loteF, anoDe, anoAte)
                n = st(0): media = st(1): dp = st(2)
                cv = CalcularCV(dp, media)
                ' lote vazio = todos os lotes: nao ha UM alvo, entao nao ha bias
                AlvoAnalito analito, t, alvoM, alvoS, etp, loteF
                bias = CalcularBias(media, alvoM)
                et = CalcularErroTotal(cv, bias)
                sg = CalcularSigma(etp, bias, cv)
                outp(er, 1) = n
                outp(er, 2) = IIf(n = 0, "", media)
                outp(er, 3) = IIf(n < 2, "", dp)
                outp(er, 4) = IIf(n < 2 Or media = 0, "", cv)
                outp(er, 5) = IIf(etp <= 0, "-", etp)            ' ADR-068
                outp(er, 6) = IIf(IsNumeric(wa.Cells(i, 19).Value), wa.Cells(i, 19).Value, "")
                outp(er, 7) = IIf(IsNumeric(wa.Cells(i, 20).Value), wa.Cells(i, 20).Value, "")
                outp(er, 8) = IIf(n = 0 Or alvoM = 0, "", bias)
                outp(er, 9) = IIf(n < 2 Or media = 0, "", et)
                outp(er, 10) = IIf(etp <= 0, "-", IIf(n < 2 Or cv = 0, "", sg))
                outp(er, 11) = IIf(etp <= 0, "-", IIf(n < 2 Or cv = 0, "", mQualidade.ClassificarSigma(sg)))
            End If
        Next t
    Next i

    ' publica em Eng_Saida; a aba Estatistica le por formula
    On Error GoTo restaura
    protEstava = LiberarEscrita(eng)
    eng.Range(eng.Cells(LINHA_EST, 3), eng.Cells(LINHA_EST + linhas - 1, 13)).Value = outp
    RestaurarProtecao eng, protEstava
    On Error GoTo 0

    ' Fora do bloco protegido de proposito: RegistrarEventosWestgard cuida da
    ' PROPRIA aba (Eventos_Westgard). Chama-la com Eng_Saida ainda destrancada
    ' ampliaria a janela de exposicao sem necessidade nenhuma.
    RegistrarEventosWestgard
    Exit Sub

restaura:
    Dim nErrP As Long, sErrP As String
    nErrP = Err.Number: sErrP = Err.Description
    RestaurarProtecao eng, protEstava
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mEstatistica.AtualizarEstatisticaAba", sErrP
End Sub



' ============================================================================
'  EVENTOS_WESTGARD -- historico auditavel, granularidade de EVENTO

'

'  Uma linha = uma EVIDENCIA do motor: uma regra disparou, num detector, num

'  escopo, fechando numa corrida. Nao e uma linha por resultado marcado --

'  um 6x N3/R2 e UM evento e marca SEIS resultados, e confundir as duas

'  coisas era o defeito que o ADR-045 encerra.

'

'  A serie vem do trace de AvaliarWestgard. Nao existe aqui uma segunda

'  implementacao das regras: a anterior (AvaliarWestgard1N) era cega para

'  R_4s, 2of3_2s, 3_1s e 6x, e ainda materializava o 10x aposentado.

'

'  Colunas A..N: Data | RUN | Analito | Niveis | Regra | Detector | Escopo |

'                Classe | N | R | RUN_Inicial | Evidencia | Classificacao |

'                Z_Max

' ============================================================================

Public Sub RegistrarEventosWestgard()
    Dim ws As Worksheet, i As Long, t As Long, j As Long, k As Long
    Dim lote As String, chave As String, nEv As Long, nDescartados As Long
    Dim porAnalito As Object, nomeDe As Object
    Dim ev() As Variant, agg As Object
    Dim prot As Boolean

    GarantirIndice
    Set ws = ThisWorkbook.Sheets("Eventos_Westgard")

    ' ADR-050: os eventos sao do LOTE (todos os analitos dele) e so mudam quando
    ' muda o banco, o parametro ou o lote -- e os tres passam por InvalidarCache.
    ' Trocar de ANALITO nao muda nenhum: refazer a cada clique do spinner era
    ' 0,8 s jogados fora com cinco anos de banco.
    If Not mAgg Is Nothing Then
        If Len(mEvLote) > 0 And mEvLote = mLotes.LotePainel() & "|" & mEquipIdx Then Exit Sub
    End If

    ' Protecao tratada pelo par LiberarEscrita/RestaurarProtecao (ADR-046).
    On Error GoTo restaura
    prot = LiberarEscritaRapida(ws)   ' ADR-062

    ' Limpa ate a ultima linha REALMENTE usada, com piso na area antiga.
    ' Fixar num numero deixava linhas orfas embaixo quando o historico crescia;
    ' usar ws.Rows.Count limparia um milhao de linhas em toda operacao do dia.
    ' Linhas 1 a 3 (titulo e cabecalho) nunca sao tocadas.
    Dim ultimaLinhaEv As Long
    ultimaLinhaEv = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If ultimaLinhaEv < 5003 Then ultimaLinhaEv = 5003
    ws.Range(ws.Cells(4, 1), ws.Cells(ultimaLinhaEv, EV_NCOL)).ClearContents
    ' recomputando: o agregado antigo (de outro lote) nao vale mais
    Set mAgg = CreateObject("Scripting.Dictionary")
    mEvLote = ""
    If IsEmpty(mDB) Then GoTo restaura
    lote = mLotes.LotePainel()
    ws.Range("K2").Value = "Lote:"
    ws.Range("L2").NumberFormat = "@"
    ws.Range("L2").Value = lote
    ws.Range("M2").Value = IIf(Len(mEquipIdx) > 0, "Equip.:", "")
    ws.Range("N2").Value = mEquipIdx

    ' ---- serie elegivel agrupada por ANALITO (todos os niveis juntos) -------
    Set porAnalito = CreateObject("Scripting.Dictionary")
    Set nomeDe = CreateObject("Scripting.Dictionary")
    ' ADR-050: pelo indice -- so as chaves "ANALITO|<este lote>". A ordem das
    ' linhas dentro de cada chave e a do banco, como na varredura anterior.
    Dim chIdx As Variant, xLin As Variant
    For Each chIdx In mIdx.Keys
        If LoteDaChave(CStr(chIdx)) = lote Then
            For Each xLin In mIdx(chIdx)
                i = xLin
                If IsNumeric(mDB(i, COL_RESULT)) Then
                    Dim nv As Long
                    nv = CLng(Val(mDB(i, COL_NIVEL)))
                    If nv >= 1 And nv <= NLV Then
                        chave = UCase$(Trim$(CStr(mDB(i, COL_ANALITO))))
                        If Not porAnalito.Exists(chave) Then
                            porAnalito.Add chave, New Collection
                            nomeDe.Add chave, Trim$(CStr(mDB(i, COL_ANALITO)))
                        End If
                        ' nivel | RUN | valor | data
                        porAnalito(chave).Add Array(nv, CLng(Val(mDB(i, COL_RUN))), _
                                                    CDbl(mDB(i, COL_RESULT)), _
                                                    IIf(IsDate(mDB(i, COL_DATA)), CDbl(CDate(mDB(i, COL_DATA))), 0))
                    End If
                End If
            Next xLin
        End If
    Next chIdx
    If porAnalito.Count = 0 Then Set mAgg = CreateObject("Scripting.Dictionary"): mEvLote = lote & "|" & mEquipIdx: GoTo restaura

    Set agg = CreateObject("Scripting.Dictionary")
    Set mAgg = agg          ' definido JA: AgregadoWestgard nunca re-dispara esta rotina
    ' BUFFER SEM TETO, e a orientacao (coluna, evento) e o motivo de ser
    ' possivel: VBA so aceita ReDim Preserve na ULTIMA dimensao. Com a matriz
    ' na forma natural (evento, coluna) ela nao poderia crescer, e voltariamos
    ' a um limite arbitrario -- que ja custou 9.317 eventos descartados no
    ' teste de estresse de 04/08/2026.
    '
    ' O limite antigo (UBound(mDB,1)) tambem nao serve mais: com granularidade
    ' de EVENTO, dez detectores podem emitir evidencia sobre a mesma corrida,
    ' entao "um evento por linha do banco" deixou de ser verdade.
    Dim capEv As Long
    capEv = 1024
    ReDim ev(1 To EV_NCOL, 1 To capEv)
    nEv = 0: nDescartados = 0

    Dim ch As Variant
    For Each ch In porAnalito.Keys
        Dim col As Collection
        Set col = porAnalito(ch)

        ' --- eixo de corridas COMUM aos niveis -----------------------------
        ' O eixo tem de ser um so. Se cada nivel tivesse o proprio indice, a
        ' corrida i do N1 poderia ser outra corrida no N3, e R_4s compararia
        ' materiais de dias diferentes -- violacao silenciosa do escopo
        ' WITHIN_RUN do ADR-042.
        Dim runs() As Long, nRun As Long, vistos As Object, item As Variant
        Dim dtRun() As Double
        Set vistos = CreateObject("Scripting.Dictionary")
        ReDim runs(1 To col.Count): ReDim dtRun(1 To col.Count)
        nRun = 0
        For j = 1 To col.Count
            item = col(j)
            If Not vistos.Exists(CStr(item(1))) Then
                vistos.Add CStr(item(1)), 1
                nRun = nRun + 1
                runs(nRun) = CLng(item(1))
                dtRun(nRun) = CDbl(item(3))
            End If
        Next j
        If nRun > 0 Then
            OrdenarCorridas runs, dtRun, nRun   ' cronologico (ADR-049)
            Dim posDe As Object
            Set posDe = CreateObject("Scripting.Dictionary")
            For j = 1 To nRun
                posDe(CStr(runs(j))) = j
            Next j

            ' --- matriz (nivel, corrida) ------------------------------------
            Dim zz() As Double, td() As Boolean, vl() As Double, dtv() As Double
            ReDim zz(0 To NLV - 1, 1 To nRun)
            ReDim td(0 To NLV - 1, 1 To nRun)
            ReDim vl(0 To NLV - 1, 1 To nRun)
            ReDim dtv(0 To NLV - 1, 1 To nRun)

            Dim analitoN As String
            analitoN = CStr(nomeDe(ch))

            Dim alvoM() As Double, alvoS() As Double, alvoOK() As Boolean
            ReDim alvoM(0 To NLV - 1): ReDim alvoS(0 To NLV - 1): ReDim alvoOK(0 To NLV - 1)
            Dim mm As Double, ss As Double, ee As Double
            For t = 0 To NLV - 1
                alvoOK(t) = AlvoAnalito(analitoN, t + 1, mm, ss, ee, lote)
                alvoM(t) = mm: alvoS(t) = ss
            Next t

            For j = 1 To col.Count
                item = col(j)
                t = CLng(item(0)) - 1
                i = CLng(posDe(CStr(item(1))))
                ' repetido no mesmo (nivel,RUN): fica o primeiro. Nivel sem
                ' parametro do lote nao entra: z = 0 nao e "no alvo" (ADR-049).
                If Not td(t, i) And alvoOK(t) Then
                    td(t, i) = True
                    vl(t, i) = CDbl(item(2))
                    dtv(t, i) = CDbl(item(3))
                    zz(t, i) = CalcularZ(vl(t, i), alvoM(t), alvoS(t))
                End If
            Next j

            ' --- O MOTOR. Nao ha segunda implementacao das regras aqui. -----
            Dim q13 As Variant, q2sMulti As Variant, qR4 As Variant
            Dim q1sMulti As Variant, qSeq As Variant, q12 As Variant
            ReDim q13(0 To NLV - 1, 1 To nRun): ReDim q2sMulti(0 To NLV - 1, 1 To nRun)
            ReDim qR4(0 To NLV - 1, 1 To nRun): ReDim q1sMulti(0 To NLV - 1, 1 To nRun)
            ReDim qSeq(0 To NLV - 1, 1 To nRun): ReDim q12(0 To NLV - 1, 1 To nRun)
            AvaliarWestgard zz, td, nRun, q13, q2sMulti, qR4, q1sMulti, qSeq, q12

            ' --- EVENTOS: uma linha por evidencia do trace -------------------
            Dim linhas As Variant, ln As String, f As Variant
            Dim nEvIni As Long
            nEvIni = nEv
            linhas = Split(TraceWestgard(), vbLf)
            For k = LBound(linhas) To UBound(linhas)
                ln = CStr(linhas(k))
                If Len(Trim$(ln)) > 0 Then
                    f = Split(ln, "|")
                    If UBound(f) >= 8 Then
                        Dim iRun As Long, iRunIni As Long, nivs As String
                        iRun = CLng(Val(Mid$(CStr(f(4)), 5)))
                        iRunIni = CLng(Val(Mid$(CStr(f(5)), 8)))
                        nivs = Mid$(CStr(f(6)), 8)
                        If nEv >= capEv Then
                            capEv = capEv * 2
                            ReDim Preserve ev(1 To EV_NCOL, 1 To capEv)
                        End If
                        If nEv >= capEv Then
                            nDescartados = nDescartados + 1
                        Else
                            nEv = nEv + 1
                            ev(1, nEv) = DataDoEvento(dtv, nivs, iRun)
                            ev(2, nEv) = runs(iRun)
                            ev(3, nEv) = analitoN
                            ev(4, nEv) = nivs
                            ev(5, nEv) = CStr(f(0))
                            ev(6, nEv) = CStr(f(1))
                            ev(7, nEv) = CStr(f(2))
                            ev(8, nEv) = CStr(f(3))
                            ev(9, nEv) = CampoEscala(CStr(f(7)), "N=")
                            ev(10, nEv) = CampoEscala(CStr(f(7)), "R=")
                            ev(11, nEv) = runs(iRunIni)
                            ev(12, nEv) = CStr(f(8))
                            ev(13, nEv) = RegraClassificacao(CStr(f(0)))
                            ev(14, nEv) = ZDoEvento(zz, td, nivs, iRun)
                        End If
                    End If
                End If
            Next k

            ' --- eventos ordenados por corrida dentro do analito -------------
            OrdenarEventosPorRun ev, nEvIni + 1, nEv

            ' --- agregados por nivel: EVENTO e MARCA sao coisas diferentes ---
            ' Um 6x N3/R2 e UM evento e marca SEIS resultados. Somar as duas
            ' coisas na mesma celula era o defeito que o ADR-045 encerra.
            For t = 0 To NLV - 1
                Dim nEvNivel As Long, nMarcados As Long
                Dim priR As String, priRun As Long, ultR As String, ultRun As Long
                Dim maxZ As Double, corridasEnv As Object
                nEvNivel = 0: nMarcados = 0
                priR = "": priRun = 0: ultR = "": ultRun = 0: maxZ = 0
                Set corridasEnv = CreateObject("Scripting.Dictionary")

                For k = nEvIni + 1 To nEv
                    If CStr(ev(8, k)) = "OFICIAL" Then
                        If NivelNaLista(CStr(ev(4, k)), t + 1) Then
                            nEvNivel = nEvNivel + 1
                            If priR = "" Then priR = CStr(ev(5, k)): priRun = CLng(ev(2, k))
                            ultR = CStr(ev(5, k)): ultRun = CLng(ev(2, k))
                            ' A corrida ENVOLVIDA nao e so a de fechamento: um 6x
                            ' N3/R2 fechado na corrida 9 olhou tambem a 8. Contar
                            ' apenas o fechamento subestimaria o que esta sob
                            ' suspeita, que e a pergunta que o analista faz.
                            Dim wIni As Long, wFim As Long, wq As Long
                            wIni = CLng(posDe(CStr(ev(11, k))))
                            wFim = CLng(posDe(CStr(ev(2, k))))
                            For wq = wIni To wFim
                                corridasEnv(CStr(runs(wq))) = 1
                            Next wq
                        End If
                    End If
                Next k

                For i = 1 To nRun
                    If td(t, i) Then
                        If Val(q13(t, i)) = 1 Or Val(q2sMulti(t, i)) = 1 Or _
                           Val(qR4(t, i)) = 1 Or Val(q1sMulti(t, i)) = 1 Or _
                           Val(qSeq(t, i)) = 1 Then
                            nMarcados = nMarcados + 1
                            If Abs(zz(t, i)) > Abs(maxZ) Then maxZ = zz(t, i)
                        End If
                    End If
                Next i

                agg(UCase$(analitoN) & "|" & CStr(t + 1)) = _
                    Array(nEvNivel, priR, priRun, ultR, ultRun, maxZ, _
                          nMarcados, corridasEnv.Count)
            Next t
        End If
    Next ch

    If nEv > 0 Then
        Dim outp() As Variant
        ReDim outp(1 To nEv, 1 To EV_NCOL)
        For i = 1 To nEv
            For j = 1 To EV_NCOL
                outp(i, j) = ev(j, i)
            Next j
        Next i
        ' A coluna Niveis e TEXTO, sempre. Sem isto o Excel converte "3" em
        ' numero e deixa "1,2" como texto -- coluna de tipo misto, que o Power
        ' Query resolve escolhendo um dos dois e descartando o resto em erro.
        ws.Range(ws.Cells(4, 4), ws.Cells(3 + nEv, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + nEv, EV_NCOL)).Value = outp
    End If
    ws.Range("J2").Value = nEv

    ' Teto de eventos: descartar em silencio esconderia violacao de Westgard do
    ' analista e do auditor. Eventos_Westgard e DERIVADA -- pode ser
    ' reconstruida --, entao interromper aqui e seguro e forca a decisao.
    If nDescartados > 0 Then
        Auditar CAT_SIS, "EVENTOS_WESTGARD_ESTOUROU", "mEstatistica", _
                0, "", "", "", 0, "", nEv, nDescartados, "", "", _
                "Buffer dinamico de " & capEv & " eventos cheio (INVARIANTE VIOLADA); " & nDescartados & _
                " evento(s) NAO registrado(s)", _
                "Arquivar Eventos_Westgard e reexecutar para restaurar o historico completo"
        Err.Raise vbObjectError + 513, "RegistrarEventosWestgard", _
                  "Historico de Westgard incompleto: " & nDescartados & _
                  " evento(s) descartado(s) por buffer cheio (" & capEv & _
                  "). O evento foi registrado no Audit_Log."
    End If

    Set mAgg = agg
    mEvLote = lote & "|" & mEquipIdx   ' ADR-050/057: vale para este lote e equipamento ate InvalidarCache

restaura:
    ' Reprotege SEMPRE -- inclusive nos caminhos de saida antecipada e no erro.
    ' Deixar a aba destrancada por causa de uma excecao seria trocar um defeito
    ' visivel por um buraco de seguranca silencioso.
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    RestaurarProtecao ws, prot
    On Error GoTo 0
    If nErr <> 0 Then Err.Raise nErr, "mEstatistica.RegistrarEventosWestgard", sErr
End Sub


' O nivel n aparece na lista "1,2,3" do evento?
Private Function NivelNaLista(ByVal lista As String, ByVal n As Long) As Boolean
    Dim x As Variant
    For Each x In Split(lista, ",")
        If CLng(Val(x)) = n Then NivelNaLista = True: Exit Function
    Next x
End Function


' Le "N=3 R=2" devolvendo o inteiro do campo pedido.
Private Function CampoEscala(ByVal escala As String, ByVal campo As String) As Variant
    Dim x As Variant
    For Each x In Split(escala, " ")
        If Left$(CStr(x), Len(campo)) = campo Then
            CampoEscala = CLng(Val(Mid$(CStr(x), Len(campo) + 1)))
            Exit Function
        End If
    Next x
    CampoEscala = ""
End Function


' Data da corrida de fechamento, pelo primeiro nivel envolvido que tenha data.
Private Function DataDoEvento(ByRef dtv() As Double, ByVal niveis As String, _
                              ByVal iRun As Long) As Variant
    Dim x As Variant, t As Long
    For Each x In Split(niveis, ",")
        t = CLng(Val(x)) - 1
        If t >= 0 And t <= NLV - 1 Then
            If dtv(t, iRun) > 0 Then DataDoEvento = CDate(dtv(t, iRun)): Exit Function
        End If
    Next x
    DataDoEvento = ""
End Function


' Maior |Z| entre os niveis envolvidos, na corrida de fechamento. Devolve o
' valor COM sinal: o lado importa para quem le a evidencia.
Private Function ZDoEvento(ByRef z() As Double, ByRef td() As Boolean, _
                           ByVal niveis As String, ByVal iRun As Long) As Variant
    Dim x As Variant, t As Long, melhor As Double, achou As Boolean
    For Each x In Split(niveis, ",")
        t = CLng(Val(x)) - 1
        If t >= 0 And t <= NLV - 1 Then
            If td(t, iRun) Then
                If Not achou Or Abs(z(t, iRun)) > Abs(melhor) Then
                    melhor = z(t, iRun): achou = True
                End If
            End If
        End If
    Next x
    If achou Then ZDoEvento = melhor Else ZDoEvento = ""
End Function


' Insercao sobre a faixa [de..ate] do buffer, chave = corrida de fechamento.
' O trace sai na ordem dos DETECTORES; o historico se le por data.
Private Sub OrdenarEventosPorRun(ByRef ev() As Variant, ByVal de As Long, ByVal ate As Long)
    ' ADR-050: ordena INDICES por chave Long (insercao estavel, mesma comparacao
    ' e portanto mesma ordem final da versao anterior) e permuta as colunas UMA
    ' vez. Antes, cada deslocamento movia as 14 colunas Variant do evento:
    ' quadratico num lote com centenas de evidencias por analito.
    Dim n As Long, i As Long, j As Long, c As Long
    Dim idx() As Long, chave() As Long, tk As Long, ti As Long
    Dim tmp() As Variant
    n = ate - de + 1
    If n < 2 Then Exit Sub
    ReDim idx(1 To n): ReDim chave(1 To n)
    For i = 1 To n
        idx(i) = de + i - 1
        chave(i) = CLng(Val(ev(2, de + i - 1)))
    Next i
    For i = 2 To n
        tk = chave(i): ti = idx(i): j = i - 1
        Do While j >= 1
            If chave(j) <= tk Then Exit Do
            chave(j + 1) = chave(j): idx(j + 1) = idx(j): j = j - 1
        Loop
        chave(j + 1) = tk: idx(j + 1) = ti
    Next i
    ReDim tmp(1 To EV_NCOL, 1 To n)
    For i = 1 To n
        For c = 1 To EV_NCOL
            tmp(c, i) = ev(c, idx(i))
        Next c
    Next i
    For i = 1 To n
        For c = 1 To EV_NCOL
            ev(c, de + i - 1) = tmp(c, i)
        Next c
    Next i
End Sub

' Agregados por analito/nivel a partir dos eventos ja calculados (sem re-varredura).
' Devolve Array(nViolacoes, primeiraRegra, primeiroRUN, ultimaRegra, ultimoRUN, maiorZ)
' As tres metricas do ADR-045, para quem so precisa dos numeros.
'
' Elas contam coisas DIFERENTES e nao devem ser somadas nem trocadas:
'
'   N_Eventos_Violacao ..... quantas vezes uma REGRA disparou (evidencias
'                            oficiais do motor que envolvem este nivel)
'   N_Resultados_Marcados .. quantos RESULTADOS deste nivel ficaram marcados
'                            por alguma regra -- um evento 6x N3/R2 marca seis
'   N_Corridas_Envolvidas .. quantas CORRIDAS deste nivel tem ao menos um
'                            resultado marcado
'
' Devolve "eventos|marcados|corridas".
Public Function MetricasWestgard(ByVal analito As String, ByVal nivel As Long) As String
    Dim a As Variant
    a = AgregadoWestgard(analito, nivel)
    MetricasWestgard = CStr(a(0)) & "|" & CStr(a(6)) & "|" & CStr(a(7))
End Function


Public Function AgregadoWestgard(ByVal analito As String, ByVal nivel As Long) As Variant
    Dim k As String
    If mAgg Is Nothing Then RegistrarEventosWestgard
    k = UCase$(Trim$(analito)) & "|" & nivel
    If mAgg.Exists(k) Then
        AgregadoWestgard = mAgg(k)
    Else
        AgregadoWestgard = Array(0, "", 0, "", 0, 0, 0, 0)
    End If
End Function

' Avaliacao de Westgard restrita a UM nivel (series independentes por analito/nivel).
' 22s intra-corrida entre niveis e R4s sao avaliados no motor do Calc, que tem a
' visao multi-nivel; aqui tratamos a serie temporal do proprio nivel.
' AvaliarWestgard1N foi REMOVIDA (ADR-045).
'
' Ela era uma segunda implementacao das regras, escrita para uma serie de UM
' nivel: estruturalmente incapaz de ver R_4s, 2of3_2s, 3_1s e 6x, que exigem
' visao multi-nivel. Pior, materializava o proibido "10 corridas seguidas do
' mesmo lado" -- o 10x que o ADR-041 aposentou --, de modo que a aba
' Eventos_Westgard contava um fenomeno que o motor ja nao reconhecia.
'
' O historico agora vem do trace de AvaliarWestgard. Fonte unica.


Public Sub AbrirEventosWestgard()
    On Error Resume Next
    ThisWorkbook.Sheets("Eventos_Westgard").Activate
End Sub

' ============================ ORQUESTRACAO ============================
Public Sub AtualizarEstatistica()
    Dim calcAntes As Long, nE As Long, sE As String
    calcAntes = Application.Calculation
    Application.ScreenUpdating = False
    On Error GoTo fim
    Application.Calculation = xlCalculationManual     ' ver RecalcularAnalitoAtual
    InvalidarCache
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEstatisticaAba
fim:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    ' ADR-050: recalcula SEMPRE antes dos eixos -- AtualizarEixos le os limites
    ' do Calc, e dentro de uma operacao maior (mApp) o calculo segue em manual.
    ' So o que ficou sujo e recalculado; nada mais varre o banco por formula.
    Application.Calculate
    Application.Calculation = calcAntes
    AtualizarEixos
    Application.ScreenUpdating = True
    On Error GoTo 0
    If nE <> 0 Then Err.Raise nE, "mEstatistica.AtualizarEstatistica", sE
End Sub

' Recalculo incremental: so o analito atualmente exibido (nao varre a planilha toda).
' ADR-049: o motor grava em blocos (Eng_Saida, Eventos_Westgard). Com o calculo
' em automatico, CADA bloco disparava o recalculo das formulas dependentes --
' medido em ~11 s por recalculo no arquivo real, pago uma vez por escrita.
' Trocar de lote ou de analito virava minutos. Agora: calculo manual durante o
' motor, UM recalculo no fim, e o modo anterior restaurado tambem no erro. O
' resultado e o mesmo: nenhuma rotina do motor le formula que ele mesmo altere.
Public Sub RecalcularAnalitoAtual()
    Dim calcAntes As Long, nE As Long, sE As String, telaAntes As Boolean
    calcAntes = Application.Calculation
    telaAntes = Application.ScreenUpdating      ' ADR-062: a tela e de quem chamou (ADR-050)
    Application.ScreenUpdating = False
    On Error GoTo fim
    Application.Calculation = xlCalculationManual
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
fim:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    ' ADR-050: recalcula SEMPRE antes dos eixos -- AtualizarEixos le os limites
    ' do Calc, e dentro de uma operacao maior (mApp) o calculo segue em manual.
    ' So o que ficou sujo e recalculado; nada mais varre o banco por formula.
    Application.Calculate
    Application.Calculation = calcAntes
    AtualizarEixos
    Application.ScreenUpdating = telaAntes   ' religar aqui redesenhava os graficos NO MEIO do clique
    On Error GoTo 0
    If nE <> 0 Then Err.Raise nE, "mEstatistica.RecalcularAnalitoAtual", sE
End Sub




' LoteDoAnoOuAtivo removida (ADR-049): o lote em analise e escolhido no
' Painel (loteSel), nao deduzido do ano.



' Grava o provedor de EQA (CAP/Controllab) escolhido no Painel para UM
' analito especifico, na aba Analitos (coluna AR). Fonte unica: a Estatistica
' LE essa coluna por formula (INDEX/MATCH), nunca guarda copia propria.
Public Sub GravarProvedorAnalito(ByVal nomeAnalito As String, ByVal provedor As String)
    Dim wa As Worksheet, i As Long
    If Trim$(nomeAnalito) = "" Then Exit Sub
    Set wa = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        If UCase$(Trim$(CStr(wa.Cells(i, 1).Value))) = UCase$(Trim$(nomeAnalito)) Then
            wa.Cells(i, 44).Value = provedor
            Exit For
        End If
    Next i
End Sub

' Lotes com resultado elegivel do analito dentro do filtro do Painel, com o
' numero de corridas: "8973 (39) · 8974 (18)". Vai para Eng_Saida!O1 e o
' Painel mostra quando o lote escolhido nao tem corrida no periodo -- em vez
' de trocar o lote sozinho, como o LoteDoAnoOuAtivo fazia.
Public Function LotesNoPeriodo(ByVal analito As String) As String
    Dim i As Long, nl As String, k As String, d As Object, runs As Object, s As String, x As Variant
    Dim ch As Variant, pre As String, y As Variant
    GarantirIndice
    If IsEmpty(mDB) Then Exit Function
    CarregarFiltro
    Set d = CreateObject("Scripting.Dictionary")
    Set runs = CreateObject("Scripting.Dictionary")
    ' ADR-050: pelo indice, chave a chave dos lotes deste analito
    pre = UCase$(Trim$(analito)) & "|"
    For Each ch In mIdx.Keys
        If Left$(ch, Len(pre)) = pre Then
            nl = LoteDaChave(CStr(ch))
            For Each y In mIdx(ch)
                i = y
                If IsDate(mDB(i, COL_DATA)) Then
                    If PassaFiltro(CDbl(CDate(mDB(i, COL_DATA)))) Then
                        k = nl & "|" & CStr(mDB(i, COL_RUN))
                        If Not runs.Exists(k) Then
                            runs.Add k, 1
                            d(nl) = d(nl) + 1
                        End If
                    End If
                End If
            Next y
        End If
    Next ch
    For Each x In d.Keys
        s = s & IIf(Len(s) > 0, " " & ChrW(183) & " ", "") & x & " (" & d(x) & ")"
    Next x
    LotesNoPeriodo = s
End Function

' Spinner do Painel e De/Ate: o motor tem de acompanhar o analito e o periodo
' em tela. Antes o spinner so reajustava o eixo e o Eng_Saida ficava no
' analito anterior -- o Calc agora se recusa a plotar um motor de outro
' analito/lote, entao quem muda a tela precisa refazer o motor.
Public Sub PainelMudou()
    ' ADR-050: uma operacao so (tela congelada, um recalculo). O motor pula os
    ' eventos de Westgard do lote se ja estao feitos, e le o banco pelo indice.
    ' ADR-053: a tela do usuario e guardada e devolvida, e a falha vira mensagem.
    Dim volta As Object
    Set volta = mApp.TelaAtual()
    mApp.InicioUsuario "Carregando " & Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value)) & "..."
    On Error GoTo falha
    ' ADR-062: RecalcularAnalitoAtual ja ajusta os eixos; aqui so quando o motor estava em dia
    If Not GarantirMotorDoPainel() Then AtualizarEixos
    mApp.fim
    mApp.VoltarPara volta
    mUI.GrafVigiaLigar          ' ADR-061: uma troca longa pode ter pulado um tique (idempotente)
    Exit Sub
falha:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    mApp.fim
    mApp.VoltarPara volta
    MsgBox "Nao foi possivel atualizar o analito." & vbCrLf & vbCrLf & _
           "Erro " & nE & ": " & sE, vbExclamation, "Painel"
End Sub

' Filtro do Painel (De/Ate, trimestres) mudou: o motor refaz a janela de
' corridas e as estatisticas do periodo -- uma operacao so (ADR-050).
' ADR-053. Duas correcoes vieram do relato "cliquei no trimestre, deu erro e fui
' parar na aba de eventos de Westgard":
'
'   1. O filtro NAO precisa dos eventos de Westgard. Eles sao do lote inteiro e
'      nao olham periodo nenhum -- refaze-los aqui era trabalho jogado fora e
'      mexia numa aba que a tela do filtro nao tem nada a ver.
'   2. A tela ativa e guardada antes e devolvida depois, e o erro vira mensagem.
'      Assim nenhuma falha no meio do caminho deixa o usuario noutra aba, ainda
'      mais agora que as guias ficam ocultas no visual de aplicativo.
Public Sub FiltroMudou()
    Dim volta As Object, nE As Long, sE As String
    Set volta = mApp.TelaAtual()
    mApp.InicioUsuario "Aplicando o filtro do periodo..."
    On Error GoTo falha
    AtualizarCalc                ' a janela de corridas e o filtro corrida a corrida
    AtualizarPainelEng           ' n, media, DP, CV e as contagens DO PERIODO
    Application.Calculate        ' os eixos leem o Calc (ver RecalcularAnalitoAtual)
    AtualizarEixos
    mApp.fim
    mApp.VoltarPara volta
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    mApp.fim
    mApp.VoltarPara volta
    MsgBox "Nao foi possivel aplicar o filtro do periodo." & vbCrLf & vbCrLf & _
           "Erro " & nE & ": " & sE & vbCrLf & vbCrLf & _
           "Desmarque os trimestres para ver o periodo inteiro.", vbExclamation, "Filtro"
End Sub

' Painel ativado: se o motor ficou para tras (outro analito ou outro lote),
' refaz antes de mostrar.
Public Function GarantirMotorDoPainel() As Boolean
    Dim eng As Worksheet
    On Error Resume Next
    Set eng = ThisWorkbook.Sheets("Eng_Saida")
    If StrComp(Trim$(CStr(eng.Range("C1").Value)), _
               Trim$(CStr(ThisWorkbook.Names("selAnalito").RefersToRange.Value)), 1) <> 0 _
       Or Trim$(CStr(eng.Range("E1").Value)) <> mLotes.LotePainel() _
       Or UCase$(Trim$(CStr(eng.Range("S1").Value))) <> EquipFiltro() Then
        RecalcularAnalitoAtual
        GarantirMotorDoPainel = True
    End If
End Function


' ADR-057: o equipamento do Painel mudou (Bioquimica). Cada equipamento e uma
' serie propria: o indice e refeito para ele e TODO o motor acompanha --
' grafico, Westgard, eventos e a aba Estatistica.
Public Sub EquipMudou()
    Dim volta As Object, nE As Long, sE As String
    Set volta = mApp.TelaAtual()
    mApp.InicioUsuario "Trocando o equipamento para " & EquipFiltro() & "..."
    On Error GoTo falha
    AtualizarEstatistica
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!RecalcularEstatPeriodo"
    mIncerteza.RecalcularIncerteza          ' ADR-064: o equipamento muda a populacao do CIQ
    On Error GoTo falha
    mApp.fim
    mApp.VoltarPara volta
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    mApp.fim
    mApp.VoltarPara volta
    MsgBox "Nao foi possivel trocar o equipamento." & vbCrLf & vbCrLf & _
           "Erro " & nE & ": " & sE, vbExclamation, "Painel"
End Sub

' ADR-062: o 1o clique depois de abrir pagava ler a tblCQ_Final inteira (32-44 mil linhas),
' montar o indice analito|lote e os eventos de Westgard do lote -- 2,7 a 3,1 s medidos.
' Aquecer no login (e ao entrar no Painel) tira esse custo do clique. Quente: sai na hora.
Public Sub AquecerMotor()
    On Error Resume Next
    GarantirIndice
    If mAgg Is Nothing Then RegistrarEventosWestgard
    mIncerteza.AquecerJanelaSigma         ' ADR-071: CVp do Painel quente na 1a troca de analito
End Sub
