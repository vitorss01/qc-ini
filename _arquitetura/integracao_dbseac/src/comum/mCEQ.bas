Option Explicit
' ===== CONTROLE EXTERNO / ENSAIO DE PROFICIENCIA (ADR-030 / ADR-032) =====
'
' O DEFEITO QUE ESTE MODULO CORRIGIU (ADR-030)
'
' O bias que alimentava Erro Total e Sigma vinha de
' mEstatistica.CalcularBias(mediaObs, alvo), onde:
'
'     mediaObs = media do CONTROLE INTERNO no periodo
'     alvo     = media atribuida AO LOTE do controle interno
'
' Isso mede a deriva do CQI contra o alvo do proprio lote -- util, mas NAO e
' erro sistematico. Erro sistematico se mede contra um valor externo e
' independente: o consenso do grupo no ensaio de proficiencia. Um laboratorio
' pode estar centrado no alvo do fabricante e, ainda assim, 8% acima do grupo.
'
' A formula sempre esteve certa dos dois lados: (X - ref)/ref * 100. O que
' estava errado era o "ref".
'
' O QUE O ADR-032 ACRESCENTA
'
' O laboratorio participa de MAIS DE UM programa -- Controllab com 4 rodadas
' anuais, CAP com 3 -- e a escolha de qual usar e decisao tecnica dele, nao do
' sistema. Todas as funcoes aqui aceitam agora provedor e rodada:
'
'     provedor  ""/"TODOS"  -> qualquer programa
'               "Controllab" | "CAP"
'     rodada    ""/"TODAS"   -> consolida as rodadas do ano
'               "A" | "B" | "C" | "D"
'
' A FONTE
'
'     EQC_Dados
'     A Analito   B Ano   C Rodada (A..D)   D Data   E Provedor   F Amostra
'     G Resultado lab (X_lab)   H Media grupo (X_ref)   I SD grupo
'     J SDI = (G-H)/I           K Lim.Inf   L Lim.Sup   M Status limites
'     N Bias % = (G-H)/H*100    O |Bias| %  P Status SDI
'
' G, H, I, K e L sao DIGITADOS pelo usuario. O resto e calculado na propria
' celula, onde fica visivel e auditavel -- este modulo CONSOME essas colunas,
' nao as recalcula. Reimplementar aqui criaria a segunda versao do indicador.
'
' CONSOLIDACAO DE MULTIPLAS RODADAS
'
' O valor que alimenta ET e Sigma e a MEDIA DAS MAGNITUDES:
'
'     |Bias|consolidado = soma(|Bias_i|) / n
'
' e nunca a media dos assinados. Rodadas de +5% e -5% descrevem um metodo que
' oscila 5% em torno do grupo; a media assinada daria 0% e afirmaria exatidao
' perfeita. Cancelamento de sinal apaga o erro que se quer medir.
'
' O bias ASSINADO continua disponivel em modo "SIGNED", para ler a direcao do
' desvio. Ele informa; nao entra nas metricas de magnitude.
'
' AUSENCIA DE DADO NAO E ZERO, E NAO PODE SER Empty
'
' Sem rodada utilizavel devolve-se o TEXTO "SEM EP". Devolver Empty parecia
' natural e estava errado: o Excel renderiza o Empty de uma UDF como ZERO na
' celula. Na primeira versao deste modulo as 80 linhas da Estatistica exibiram
' bias 0,00 -- inclusive analitos sem nenhuma rodada -- e esse zero entrou em
' ET e Sigma produzindo numeros de aparencia perfeita.
'
' VIGENCIA: A RODADA MAIS RECENTE QUE NAO ULTRAPASSA O ANO DE REFERENCIA
'
' O EP e anual e sai depois; o CQI e do mes corrente. Exigir coincidencia exata
' faria o bias sumir sempre que o CQI passasse na frente do ultimo ciclo
' publicado. Mesma regra de vigencia do ADR-022 para especificacao.

' ADR-034: A FONTE PASSOU A SER A EQA_Base
'
' Ate aqui este modulo lia a EQC_Dados, aba unica onde CAP e Controllab
' dividiam as mesmas colunas. Agora cada provedor tem a sua aba de digitacao,
' com a terminologia dele, e a EQA_Base normaliza as duas.
'
' Este e o UNICO lugar da pasta que le a aba de EP. As 403 celulas da
' Estatistica e do Painel, e a coluna de bias do BI, chamam as funcoes daqui --
' nenhuma delas aponta para a planilha. Por isso trocar a fonte foi trocar
' estas constantes, e nao 403 formulas.
'
' O analito casado e o CANONICO (coluna E), nao o nome do provedor (coluna D):
' o CAP reporta "Urea Nitrogen" e a Analitos chama "Ureia". A coluna D continua
' na base para rastrear ate o PDF.

Private Const EQ_ABA As String = "EQA_Base"
Private Const EQ_R0 As Long = 2
Private Const EQ_RN As Long = 5001
Private Const EQ_C_PROVEDOR As Long = 1
Private Const EQ_C_ANO As Long = 2
Private Const EQ_C_RODADA As Long = 3
Private Const EQ_C_ANALITO As Long = 5
Private Const EQ_C_XLAB As Long = 7
Private Const EQ_C_XREF As Long = 8
Private Const EQ_C_SDGRUPO As Long = 9
Private Const EQ_C_SDI As Long = 10
Private Const EQ_C_LIMINF As Long = 11
Private Const EQ_C_LIMSUP As Long = 12
Private Const EQ_C_BIAS As Long = 16
Private Const EQ_C_BIASABS As Long = 17
Private Const EQ_C_USO As Long = 20
Private Const EQ_C_CHAVE As Long = 21     ' ADR-064: chave provedor|ano|rodada|analito|amostra
Private Const EQ_C_STATUS As Long = 14    ' ADR-064: Status_Padronizado (ACEITO | NAO ACEITO | NAO AVALIADO)
Public Const VIES_M_MIN As Long = 6       ' ADR-064: triagem do vies -- amostras minimas
Public Const VIES_ROD_MIN As Long = 2     '          e rodadas minimas
Private Const EQ_NCOL As Long = 21

Public Const SEM_EP As String = "SEM EP"
Public Const LIM_SDI As Double = 2#

' cache da EQA_Base (ADR-050) -- ver GarantirEQ
Private mEQSig As String
Private mEQPorAn As Object          ' ANALITO -> matriz (linhas do analito, colunas 1..EQ_NCOL)
Private mEQExiste As Boolean

' ---------------------------------------------------------------------------
' ADR-071: SELECAO DE RODADAS (texto de eqRodada, Estatistica!P4)
'
' Nenhuma formula da pasta mudou: as ~1.200 (Bio) / ~1.800 (Hema) chamadas
' continuam recebendo eqRodada. Muda o TEXTO aceito e a leitura dele aqui:
'
'   ""/TODAS/TODOS          todas as rodadas do ano vigente (<= eqAnoEP)  -- como antes
'   "C-A 2025"              uma rodada, rotulo sem ano                    -- como antes
'   "2025|C-A 2025"         uma rodada com o ano explicito
'   "2025|C-A 2025; 2026|C-B 2026; ..."  conjunto livre (pode atravessar anos)
'   "ULTIMAS n"             as n rodadas (ano|rodada) mais recentes do provedor, ano <= eqAnoEP
'   "ACUMULADAS"            todas as rodadas do provedor com ano <= eqAnoEP
'
' Nos modos novos o PROVEDOR e obrigatorio (D1): vazio = SEM EP, CAP e
' Controllab nunca se misturam. Texto que nao se le = SEM EP (K5 diz por que),
' nunca #VALOR!. O estimador (BiasEQ "ABS", duas etapas) NAO muda: muda so o
' conjunto de linhas que entra. 3 rodadas sao opcao operacional, nao
' equivalencia estatistica a 6.
Private Const MR_INVALIDO As Long = -1
Private Const MR_LIVRE As Long = 0
Private Const MR_ROTULO As Long = 1
Private Const MR_CONJUNTO As Long = 2
Private Const MR_ACUM As Long = 3
Private Const MR_ULTIMAS As Long = 4
Private Const MR_MAX_ULT As Long = 64
Private Const SEP_ROD As String = "|"
Private Const SEP_CONJ As String = ";"

' o texto lido por ultimo (todas as chamadas de um recalculo usam o mesmo)
Private mRodOk As Boolean
Private mRodTxt As String
Private mRodModo As Long
Private mRodN As Long
Private mRodConj As Object           ' "ANO|ROTULO" -> 1 (modo conjunto, na ordem digitada)
Private mRodErro As String
' rodadas existentes por provedor (montado com o cache da EQA_Base)
Private mRodPorProv As Object        ' PROVEDOR -> Dictionary("ANO|ROTULO" -> Array(ano, ROTULO))
Private mRodPorProvAn As Object      ' "PROVEDOR|ANALITO" -> idem, so as rodadas em que O ANALITO tem dado valido
Private mUltCache As Object          ' "PROV|ANALITO|TETO|N" -> Dictionary("ANO|ROTULO" -> 1)
Private mResumoCache As Object       ' "prov|ano|rodada" -> texto do K5
Private mEQErro As String            ' a EQA_Base nao pode ser lida (K5 diz); "" = lida
Private mEQSigFalha As String        ' carimbo da leitura que falhou (nao tenta a cada celula)


' ADR-071 (revisao 07/10/2026): UMA regra de "valor numerico de verdade" para toda leitura da EQA_Base --
' IsNumeric E nao vazio (e nunca booleano nem valor de erro). IsNumeric(Empty) = True no VBA: o |bias|
' de uma rodada ainda sem alvo (a celula fica vazia na consolidacao) entrava como 0% no BiasEQ, no
' AnoVigente, no SDI, nos limites e no K5 -- inclusive no modo TODAS (defeito anterior ao ADR-071) --,
' enquanto as listas e o ULTIMAS n ja o excluiam. Valor de erro (#DIV/0! colado) = False, sem CStr
' (que daria o erro 13 e derrubava a montagem do cache inteiro).
Private Function NumOk(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Or IsArray(v) Or IsObject(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    NumOk = (Len(Trim$(CStr(v))) > 0)
End Function

' CStr que nunca falha: valor de erro, Null ou matriz = "" (antes: erro 13). Intervalo -> o valor dele.
Private Function TxtEQ(ByVal v As Variant) As String
    On Error GoTo fim
    If IsObject(v) Then v = v.Value
    If IsError(v) Or IsNull(v) Or IsArray(v) Then Exit Function
    TxtEQ = CStr(v)
fim:
End Function

Private Function Igual(ByVal a As Variant, ByVal b As String) As Boolean
    If IsError(a) Then Exit Function
    Igual = (UCase$(Trim$(CStr(a))) = UCase$(Trim$(b)))
End Function

' Um filtro vazio, "TODOS" ou "TODAS" nao restringe nada.
Private Function Livre(ByVal f As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(TxtEQ(f)))
    Livre = (s = "" Or s = "TODOS" Or s = "TODAS")
End Function

' ---------------------------------------------------------------------------
' CACHE DA EQA_Base (ADR-050)
'
' Cada celula da Estatistica que chama BiasEQ/SDIeq/StatusLimitesEQ relia a
' EQA_Base INTEIRA (5.000 x 21 celulas) e a varria duas vezes. Sao ~480
' chamadas: ~10 s toda vez que o ano de EQA mudava -- inclusive, sem ninguem
' perceber, a cada importacao de corrida (a lista de anos era regravada).
'
' Agora a base e lida UMA vez, agrupada por analito, e cada chamada recebe so
' as linhas do analito dela (mesma ordem). A validade do cache e o CARIMBO que
' mEQA.AtualizarEQABase grava a cada consolidacao -- a unica rotina que escreve
' na EQA_Base. InvalidarEQ (chamada por mEstatistica.InvalidarCache) descarta
' o cache em qualquer operacao de dados.

Public Sub InvalidarEQ()
    mEQSig = ""
    mEQSigFalha = ""
    Set mEQPorAn = Nothing
    Set mRodPorProv = Nothing
    Set mRodPorProvAn = Nothing
    Set mUltCache = Nothing
    Set mResumoCache = Nothing
End Sub

' ADR-071 (revisao): o cache so passa a valer DEPOIS de montado. Antes o carimbo era gravado antes do
' laco: um erro no meio (CStr de uma celula com #DIV/0! colado) deixava o dicionario vazio com o carimbo
' novo, e toda a Estatistica ficava em SEM EP, sem aviso, ate a proxima consolidacao. Agora monta em
' variaveis locais, so entao publica; se mesmo assim falhar, publica um cache VAZIO com o motivo
' (mEQErro, que o K5 mostra) e nao marca o carimbo como lido -- InvalidarEQ ou nova consolidacao tentam
' de novo. Nunca propaga erro para quem chama (AtualizarListasAno, Workbook_Open).
Private Sub GarantirEQ()
    Dim sig As String, ws As Worksheet, d As Variant, i As Long, j As Long, c As Long
    Dim an As String, grupos As Object, k As Variant, x As Variant, m() As Variant
    Dim pv As String, ch As String, it As Variant
    Dim porAn As Object, rodProv As Object, rodProvAn As Object
    sig = mEQA.CarimboEQA()
    If Not mEQPorAn Is Nothing Then
        If sig = mEQSig Then Exit Sub
        If Len(mEQSigFalha) > 0 And sig = mEQSigFalha Then Exit Sub
    End If
    On Error GoTo falhou
    Set porAn = CreateObject("Scripting.Dictionary")
    Set rodProv = CreateObject("Scripting.Dictionary")
    Set rodProvAn = CreateObject("Scripting.Dictionary")
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(EQ_ABA)
    On Error GoTo falhou
    If Not ws Is Nothing Then
        ' Da coluna 1 ate a ultima: assim d(i, EQ_C_XLAB) e literalmente a coluna
        ' EQ_C_XLAB. Ler a partir do analito (coluna 5) deslocaria todo indice em 4.
        d = ws.Range(ws.Cells(EQ_R0, 1), ws.Cells(EQ_RN, EQ_NCOL)).Value
        Set grupos = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(d, 1)
            an = UCase$(Trim$(TxtEQ(d(i, EQ_C_ANALITO))))
            If Len(an) > 0 Then
                If Not grupos.Exists(an) Then grupos.Add an, New Collection
                grupos(an).Add i
                ' ADR-071: a rodada EXISTE para o provedor se tem linha utilizavel (Uso <> NAO,
                ' analito canonico, |bias| numerico de verdade) -- a mesma regra das listas (mDados)
                ' e do calculo. Guardada tambem por provedor|analito: ULTIMAS n e por analito.
                If Not Igual(d(i, EQ_C_USO), "NAO") Then
                    If NumOk(d(i, EQ_C_BIASABS)) Then
                        ch = ChaveRodada(d, i)
                        If Len(ch) > 0 Then
                            pv = UCase$(Trim$(TxtEQ(d(i, EQ_C_PROVEDOR))))
                            it = Array(CLng(Val(TxtEQ(d(i, EQ_C_ANO)))), UCase$(Trim$(TxtEQ(d(i, EQ_C_RODADA)))))
                            GuardarRodada rodProv, pv, ch, it
                            GuardarRodada rodProvAn, pv & SEP_ROD & an, ch, it
                        End If
                    End If
                End If
            End If
        Next i
        For Each k In grupos.Keys
            ReDim m(1 To grupos(k).Count, 1 To EQ_NCOL)
            j = 0
            For Each x In grupos(k)
                j = j + 1
                For c = 1 To EQ_NCOL
                    m(j, c) = d(x, c)
                Next c
            Next x
            porAn.Add k, m
        Next k
    End If
    ' so agora o cache vale
    Set mEQPorAn = porAn
    Set mRodPorProv = rodProv
    Set mRodPorProvAn = rodProvAn
    Set mUltCache = CreateObject("Scripting.Dictionary")
    Set mResumoCache = CreateObject("Scripting.Dictionary")
    mEQExiste = Not (ws Is Nothing)
    mEQErro = ""
    mEQSigFalha = ""
    mEQSig = sig
    Exit Sub
falhou:
    mEQErro = "EQA_Base ilegivel (" & Err.Description & ")"
    On Error Resume Next
    Set mEQPorAn = CreateObject("Scripting.Dictionary")
    Set mRodPorProv = CreateObject("Scripting.Dictionary")
    Set mRodPorProvAn = CreateObject("Scripting.Dictionary")
    Set mUltCache = CreateObject("Scripting.Dictionary")
    Set mResumoCache = CreateObject("Scripting.Dictionary")
    mEQExiste = False
    mEQSig = ""
    mEQSigFalha = sig
End Sub

' ADR-071 (revisao): depois de consolidar o CEQ (mEQA.AtualizarEQABase), as celulas que chamam as funcoes
' deste modulo (Estatistica G, R, S, T, AC, AD, K5 e as do vies AJ..AU) NAO dependem de nenhuma celula que
' mudou -- leem a EQA_Base pelo VBA, e os argumentos (eqProvedor, eqAnoEP, eqRodada) sao os mesmos. O Excel
' nao as recalcularia: G e K5 ficavam com o conjunto de rodadas antigo ate alguem mexer em N4/P4/R4,
' enquanto L (Sigma, sujo pelo RecalcularIncerteza) era recalculado sobre o G velho. Marca como sujas
' so as celulas da Estatistica cuja formula chama o CEQ (o Painel le a Estatistica e acompanha);
' quem chama roda o Application.Calculate (mIncerteza.RecalcularIncerteza).
Public Sub MarcarCEQSujo()
    Dim ws As Worksheet, f As Variant, r As Long, c As Long, t As String, rg As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Estat" & ChrW$(237) & "stica")
    If ws Is Nothing Then Exit Sub
    Set rg = ws.UsedRange
    f = rg.Formula
    If Not IsArray(f) Then Exit Sub
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            t = CStr(f(r, c))
            If Left$(t, 1) = "=" Then
                t = UCase$(t)
                If InStr(t, "BIASEQ(") > 0 Or InStr(t, "SDIEQ(") > 0 Or InStr(t, "LIMITESEQ(") > 0 Or _
                   InStr(t, "VIESEQ(") > 0 Or InStr(t, "RESUMOFILTROEQ(") > 0 Then rg.Cells(r, c).Dirty
            End If
        Next c
    Next r
End Sub

Private Sub GuardarRodada(ByVal dic As Object, ByVal chave As String, ByVal rodada As String, ByVal it As Variant)
    If Not dic.Exists(chave) Then dic.Add chave, CreateObject("Scripting.Dictionary")
    If Not dic(chave).Exists(rodada) Then dic(chave).Add rodada, it
End Sub

' So as linhas do analito (Empty se nao houver nenhuma). Os lacos que usam o
' resultado continuam conferindo analito/provedor/rodada/uso por Casa().
Private Function LerBanco(ByVal analito As String) As Variant
    Dim k As String
    GarantirEQ
    k = UCase$(Trim$(analito))
    If mEQPorAn.Exists(k) Then LerBanco = mEQPorAn(k)
End Function

' ---------------------------------------------------------------------------
' ADR-071: leitura do texto de eqRodada
Private Function SoDigitos(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    SoDigitos = True
End Function

' Texto da celula (ou do valor) em maiusculas; ok = False para erro, matriz ou intervalo de varias celulas.
Private Function TextoRodada(ByVal rodada As Variant, ByRef ok As Boolean) As String
    Dim v As Variant
    ok = False
    On Error GoTo fim
    If IsObject(rodada) Then v = rodada.Value Else v = rodada
    If IsArray(v) Or IsError(v) Then Exit Function
    TextoRodada = UCase$(Trim$(CStr(v)))
    ok = True
fim:
End Function

Private Sub LerRodada(ByVal s As String)
    Dim partes() As String, i As Long, tok As String, p As Long, a As String, r As String
    mRodTxt = s
    mRodOk = True
    mRodN = 0
    mRodErro = ""
    Set mRodConj = Nothing
    If s = "" Or s = "TODAS" Or s = "TODOS" Then mRodModo = MR_LIVRE: Exit Sub
    If s = "ACUMULADAS" Then mRodModo = MR_ACUM: Exit Sub
    If Left$(s, 7) = "ULTIMAS" Or Left$(s, 7) = "ÚLTIMAS" Then
        a = Trim$(Mid$(s, 8))
        mRodModo = MR_INVALIDO
        mRodErro = "ULTIMAS n: n inteiro de 1 a " & MR_MAX_ULT
        If SoDigitos(a) And Len(a) <= 3 Then
            If CLng(a) >= 1 And CLng(a) <= MR_MAX_ULT Then
                mRodModo = MR_ULTIMAS
                mRodN = CLng(a)
                mRodErro = ""
            End If
        End If
        Exit Sub
    End If
    If InStr(s, SEP_ROD) = 0 And InStr(s, SEP_CONJ) = 0 Then mRodModo = MR_ROTULO: Exit Sub
    ' conjunto: itens ano|rodada separados por ";"
    mRodModo = MR_INVALIDO
    Set mRodConj = CreateObject("Scripting.Dictionary")
    partes = Split(s, SEP_CONJ)
    For i = LBound(partes) To UBound(partes)
        tok = Trim$(partes(i))
        If Len(tok) > 0 Then
            p = InStr(tok, SEP_ROD)
            If p = 0 Then
                mRodErro = "item sem ano (use ano|rodada): " & tok
                Set mRodConj = Nothing
                Exit Sub
            End If
            a = Trim$(Left$(tok, p - 1))
            r = Trim$(Mid$(tok, p + 1))
            If Not SoDigitos(a) Or Len(a) <> 4 Or Len(r) = 0 Or InStr(r, SEP_ROD) > 0 Then
                mRodErro = "item invalido: " & tok
                Set mRodConj = Nothing
                Exit Sub
            End If
            mRodConj(a & SEP_ROD & r) = 1
        End If
    Next i
    If mRodConj.Count = 0 Then
        mRodErro = "conjunto vazio"
        Set mRodConj = Nothing
        Exit Sub
    End If
    mRodModo = MR_CONJUNTO
End Sub

' Modo do texto de eqRodada (cache do ultimo texto lido).
Private Function ModoRodada(ByVal rodada As Variant) As Long
    Dim s As String, ok As Boolean
    s = TextoRodada(rodada, ok)
    If Not ok Then
        ModoRodada = MR_INVALIDO
        Exit Function
    End If
    If Not (mRodOk And s = mRodTxt) Then LerRodada s
    ModoRodada = mRodModo
End Function

' "ANO|ROTULO" da linha ("" se o ano nao e numero).
Private Function ChaveRodada(ByRef d As Variant, ByVal i As Long) As String
    If Not NumOk(d(i, EQ_C_ANO)) Then Exit Function
    ChaveRodada = CStr(CLng(Val(CStr(d(i, EQ_C_ANO))))) & SEP_ROD & UCase$(Trim$(TxtEQ(d(i, EQ_C_RODADA))))
End Function

Private Function TetoAno(ByVal anoRef As Variant) As Long
    TetoAno = 32767
    On Error Resume Next
    If IsObject(anoRef) Then anoRef = anoRef.Value
    If IsNumeric(anoRef) Then
        If Len(Trim$(CStr(anoRef))) > 0 Then TetoAno = CLng(Val(CStr(anoRef)))
    End If
End Function

' Ordem cronologica das rodadas: ano e, no mesmo ano, o rotulo (numerico quando os dois sao numeros).
' A EQA_Base nao tem data da rodada -- e a unica ordem possivel (ADR-071, limitacao).
Private Function RodadaAntes(ByVal a1 As Long, ByVal r1 As String, ByVal a2 As Long, ByVal r2 As String) As Boolean
    If a1 <> a2 Then RodadaAntes = (a1 < a2): Exit Function
    If SoDigitos(r1) And SoDigitos(r2) And Len(r1) <= 9 And Len(r2) <= 9 Then
        RodadaAntes = (CLng(r1) < CLng(r2))
    Else
        RodadaAntes = (StrComp(r1, r2, vbBinaryCompare) < 0)
    End If
End Function

' Rodadas do provedor com ano <= teto, da mais recente para a mais antiga (matriz 1..n; Empty se
' nenhuma). provedor "" junta todos os provedores (so para a lista da Bioquimica).
Private Function RodadasDoProvedor(ByVal provedor As Variant, ByVal aRef As Long) As Variant
    Dim pv As String, k As Variant, pk As Variant, todos As Object
    pv = UCase$(Trim$(TxtEQ(provedor)))
    If mRodPorProv Is Nothing Then Exit Function
    Set todos = CreateObject("Scripting.Dictionary")
    For Each pk In mRodPorProv.Keys
        If pv = "" Or pk = pv Then
            For Each k In mRodPorProv(pk).Keys
                If Not todos.Exists(k) Then todos.Add k, mRodPorProv(pk)(k)
            Next k
        End If
    Next pk
    RodadasDoProvedor = OrdenarRodadas(todos, aRef)
End Function

' Rodadas ("ANO|ROTULO" -> Array(ano, ROTULO)) com ano <= aRef, da mais recente para a mais antiga
' (matriz 1..n; Empty se nenhuma).
Private Function OrdenarRodadas(ByVal todos As Object, ByVal aRef As Long) As Variant
    Dim k As Variant, it As Variant, n As Long, i As Long, j As Long
    Dim anos() As Long, rots() As String, chs() As String, ta As Long, tr As String, tc As String
    Dim tot As Long
    If todos Is Nothing Then Exit Function
    tot = todos.Count
    If tot = 0 Then Exit Function
    ReDim anos(1 To tot)
    ReDim rots(1 To tot)
    ReDim chs(1 To tot)
    For Each k In todos.Keys
        it = todos(k)
        If it(0) <= aRef Then
            n = n + 1
            anos(n) = it(0): rots(n) = it(1): chs(n) = CStr(k)
        End If
    Next k
    If n = 0 Then Exit Function
    For i = 2 To n                                 ' insercao, decrescente
        ta = anos(i): tr = rots(i): tc = chs(i)
        j = i - 1
        Do While j >= 1
            If Not RodadaAntes(anos(j), rots(j), ta, tr) Then Exit Do
            anos(j + 1) = anos(j): rots(j + 1) = rots(j): chs(j + 1) = chs(j)
            j = j - 1
        Loop
        anos(j + 1) = ta: rots(j + 1) = tr: chs(j + 1) = tc
    Next i
    ReDim Preserve chs(1 To n)
    OrdenarRodadas = chs
End Function

' Conjunto das ULTIMAS n rodadas DO ANALITO no provedor (ano <= teto): as n mais recentes em que ele tem
' linha valida (Uso <> NAO, |bias| numerico) -- em cache por provedor|analito|teto|n.
' ADR-071 (revisao): antes era um conjunto unico por provedor, montado com as rodadas de TODOS os analitos:
' com duas familias de survey no mesmo ano (C-x e LN2-x 2025), "ULTIMAS 3" virava {LN2-C, LN2-B, LN2-A}
' (StrComp poe LN2 depois de C) e todos os analitos da familia C ficavam SEM EP; um analito nao avaliado
' nas ultimas rodadas do provedor entrava com menos de n (ou nenhuma).
Private Function ConjuntoUltimas(ByVal provedor As Variant, ByVal analito As String, ByVal anoRef As Variant) As Object
    Dim aRef As Long, key As String, pk As String, lst As Variant, i As Long, s As Object
    aRef = TetoAno(anoRef)
    pk = UCase$(Trim$(TxtEQ(provedor))) & SEP_ROD & UCase$(Trim$(analito))
    key = pk & SEP_ROD & aRef & SEP_ROD & mRodN
    If mUltCache Is Nothing Then Set mUltCache = CreateObject("Scripting.Dictionary")
    If mUltCache.Exists(key) Then Set ConjuntoUltimas = mUltCache(key): Exit Function
    Set s = CreateObject("Scripting.Dictionary")
    If Not mRodPorProvAn Is Nothing Then
        If mRodPorProvAn.Exists(pk) Then lst = OrdenarRodadas(mRodPorProvAn(pk), aRef)
    End If
    If IsArray(lst) Then
        For i = 1 To UBound(lst)
            If i > mRodN Then Exit For
            s(lst(i)) = 1
        Next i
    End If
    mUltCache.Add key, s
    Set ConjuntoUltimas = s
End Function

' ADR-071: listas das validacoes (mDados.AtualizarListasAno) saem DAQUI -- a mesma regra de
' "rodada existente" que o BiasEQ usa (Uso <> NAO, analito canonico, |bias| numerico): a lista
' nunca oferece rodada de simulacao nem rodada que o calculo nao enxerga.
' Rodadas "ANO|ROTULO" do provedor ("" = todos), da mais recente para a mais antiga; matriz 0..n-1.
Public Function RodadasExistentes(Optional ByVal provedor As String = "") As Variant
    Dim lst As Variant, i As Long, out() As String
    GarantirEQ
    lst = RodadasDoProvedor(provedor, 32767)
    If Not IsArray(lst) Then RodadasExistentes = Array(): Exit Function
    ReDim out(0 To UBound(lst) - 1)
    For i = 1 To UBound(lst)
        out(i - 1) = lst(i)
    Next i
    RodadasExistentes = out
End Function

' Anos com rodada existente do provedor ("" = todos), em ordem crescente; matriz 0..n-1.
Public Function AnosExistentes(Optional ByVal provedor As String = "") As Variant
    Dim lst As Variant, i As Long, j As Long, a As Long, anos As Object, out() As Long, k As Variant, t As Long
    GarantirEQ
    lst = RodadasDoProvedor(provedor, 32767)
    If Not IsArray(lst) Then AnosExistentes = Array(): Exit Function
    Set anos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(lst)
        a = CLng(Val(Split(lst(i), SEP_ROD)(0)))
        anos(a) = 1
    Next i
    ReDim out(0 To anos.Count - 1)
    i = 0
    For Each k In anos.Keys
        out(i) = CLng(k): i = i + 1
    Next k
    For i = 1 To UBound(out)                         ' insercao, crescente
        t = out(i): j = i - 1
        Do While j >= 0
            If out(j) <= t Then Exit Do
            out(j + 1) = out(j): j = j - 1
        Loop
        out(j + 1) = t
    Next i
    AnosExistentes = out
End Function

' A linha pertence ao analito, ao provedor e a rodada pedidos?
Private Function Casa(ByRef d As Variant, ByVal i As Long, ByVal analito As String, _
                      ByVal provedor As Variant, ByVal rodada As Variant, _
                      Optional ByVal anoRef As Variant = "") As Boolean
    Dim md As Long
    If Not Igual(d(i, EQ_C_ANALITO), analito) Then Exit Function
    ' Uso_Analitico = NAO marca dado preservado por historico que nao pode
    ' entrar em bias, Sigma nem ET -- hoje, os 90 registros de simulacao que
    ' vinham da EQC_Dados. Ver o cabecalho do mEQA.
    If Igual(d(i, EQ_C_USO), "NAO") Then Exit Function
    md = ModoRodada(rodada)
    If md = MR_INVALIDO Then Exit Function
    If md >= MR_CONJUNTO Then
        ' ADR-071 (D1): modos novos exigem o provedor -- CAP e Controllab nunca se misturam
        If Livre(provedor) Then Exit Function
        If Not Igual(d(i, EQ_C_PROVEDOR), TxtEQ(provedor)) Then Exit Function
        If md = MR_CONJUNTO Then
            If Not mRodConj.Exists(ChaveRodada(d, i)) Then Exit Function
        ElseIf md = MR_ULTIMAS Then
            If Not ConjuntoUltimas(provedor, analito, anoRef).Exists(ChaveRodada(d, i)) Then Exit Function
        End If
        Casa = True
        Exit Function
    End If
    ' modos de antes (TODAS / rotulo unico): regra do ADR-032, intacta
    If Not Livre(provedor) Then
        If Not Igual(d(i, EQ_C_PROVEDOR), TxtEQ(provedor)) Then Exit Function
    End If
    If Not Livre(rodada) Then
        If Not Igual(d(i, EQ_C_RODADA), TxtEQ(rodada)) Then Exit Function
    End If
    Casa = True
End Function

' Maior ano que nao ultrapassa anoRef, ja respeitando provedor e rodada. No conjunto explicito
' (ADR-071) o ano vem de cada item: sem teto. A linha so conta com valor numerico de verdade (NumOk)
' na coluna exigida -- |bias| vazio de rodada sem alvo nao torna o ano vigente.
Private Function AnoVigente(ByRef d As Variant, ByVal analito As String, _
                            ByVal anoRef As Variant, ByVal provedor As Variant, _
                            ByVal rodada As Variant, ByVal colExigida As Long) As Long
    Dim i As Long, aRef As Long, ano As Long
    AnoVigente = -32768
    aRef = TetoAno(anoRef)
    If ModoRodada(rodada) = MR_CONJUNTO Then aRef = 32767
    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
        If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
        If Not NumOk(d(i, colExigida)) Then GoTo prox
        ano = CLng(Val(CStr(d(i, EQ_C_ANO))))
        If ano <= aRef And ano > AnoVigente Then AnoVigente = ano
prox:
    Next i
End Function

' O ano da linha entra? Modos de antes: SO o ano vigente (ADR-030). Modos novos: todo ano ate o
' vigente (que ja respeita o teto em ACUMULADAS/ULTIMAS; no conjunto o filtro e o proprio item).
' novo = ModoRodada(rodada) >= MR_CONJUNTO (calculado uma vez antes do laco).
Private Function AnoEntra(ByVal ano As Long, ByVal anoVig As Long, ByVal novo As Boolean) As Boolean
    If novo Then
        AnoEntra = (ano <= anoVig)
    Else
        AnoEntra = (ano = anoVig)
    End If
End Function

Private Function ModoNovo(ByVal rodada As Variant) As Boolean
    ModoNovo = (ModoRodada(rodada) >= MR_CONJUNTO)
End Function

' ---------------------------------------------------------------------------
' Bias do ensaio de proficiencia, consolidado.
'
'   modo  "ABS"    media das magnitudes -> alimenta ET e Sigma
'         "SIGNED" media dos assinados  -> leitura da direcao
'         "N"      quantas rodadas entraram
'         "ANO"    qual ano acabou vigente
'
' Devolve "SEM EP" quando nao ha rodada utilizavel.
Public Function BiasEQ(ByVal analito As String, ByVal anoRef As Variant, _
                       ByVal modo As String, _
                       Optional ByVal provedor As Variant = "", _
                       Optional ByVal rodada As Variant = "") As Variant
    ' -----------------------------------------------------------------
    ' CONSOLIDACAO EM DUAS ETAPAS (ADR-035)
    '
    '   etapa 1: para cada RODADA, a media dos |bias| das amostras dela
    '   etapa 2: a media dessas medias de rodada
    '
    ' A media simples de todas as amostras juntas dava peso maior a rodada
    ' que por acaso teve mais amostras. Com C-A e C-C de 5 amostras e uma
    ' rodada nova de 12, a rodada nova passaria a mandar no numero sem que
    ' ninguem tivesse decidido isso.
    '
    ' E as duas etapas trabalham sobre |bias|, nunca sobre o bias com sinal:
    ' +8% e -8% na mesma rodada descrevem um metodo que ninguem aprovaria, e
    ' a media assinada devolveria zero.
    '
    '   modo  "ABS"       magnitude -> alimenta ET e Sigma
    '         "SIGNED"    direcao do desvio, mesma consolidacao
    '         "N"         quantas AMOSTRAS entraram
    '         "NRODADAS"  quantas RODADAS entraram
    '         "ANO"       qual ano acabou vigente
    '         "DETALHE"   memoria de calculo, rodada a rodada
    '
    ' Devolve "SEM EP" quando nao ha rodada utilizavel -- nunca 0, que a
    ' celula exibiria como exatidao perfeita.
    ' -----------------------------------------------------------------
    ' ADR-071 (revisao): sem teto de rodadas. Havia MAX_ROD = 64 em vetores fixos: com ACUMULADAS
    ' (atravessa anos) ou um conjunto longo, a 65a rodada em diante era descartada EM SILENCIO, pela
    ' ordem fisica da EQA_Base (nao pela antiguidade), e NRODADAS dizia 64. Agora a rodada e achada
    ' num dicionario e os vetores crescem conforme precisam.
    Dim d As Variant, i As Long, k As Long, cap As Long
    Dim anoVig As Long, col As Long, md As String, r As String
    Dim idx As Object
    Dim rot() As String
    Dim somaAbs() As Double
    Dim somaSig() As Double
    Dim cnt() As Long
    Dim nRod As Long, nAmostras As Long
    Dim acc As Double, mr As Double, det As String, novo As Boolean

    BiasEQ = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco(analito)
    If IsEmpty(d) Then Exit Function

    md = UCase$(Trim$(modo))
    Select Case md
        Case "ABS":                          col = EQ_C_BIASABS
        Case "SIGNED":                       col = EQ_C_BIAS
        Case "N", "NRODADAS", "ANO", "DETALHE": col = EQ_C_BIASABS
        Case Else:                           Exit Function
    End Select

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_BIASABS)
    If anoVig = -32768 Then Exit Function
    If md = "ANO" Then BiasEQ = anoVig: Exit Function
    novo = ModoNovo(rodada)
    Set idx = CreateObject("Scripting.Dictionary")
    cap = 8
    ReDim rot(1 To cap)
    ReDim somaAbs(1 To cap)
    ReDim somaSig(1 To cap)
    ReDim cnt(1 To cap)

    ' ---- etapa 1: acumula por rodada ---------------------------------
    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
        If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
        If Not AnoEntra(CLng(Val(CStr(d(i, EQ_C_ANO)))), anoVig, novo) Then GoTo prox
        If Not NumOk(d(i, col)) Then GoTo prox

        ' ADR-071: a rodada e ANO|ROTULO (como no ViesEQ) -- rotulos repetidos em anos diferentes
        ' ('A', '1') nao viram uma rodada so. No modo TODAS so ha um ano: nada muda.
        r = ChaveRodada(d, i)
        If idx.Exists(r) Then
            k = idx(r)
        Else
            nRod = nRod + 1
            If nRod > cap Then
                cap = cap * 2
                ReDim Preserve rot(1 To cap)
                ReDim Preserve somaAbs(1 To cap)
                ReDim Preserve somaSig(1 To cap)
                ReDim Preserve cnt(1 To cap)
            End If
            k = nRod
            rot(k) = r
            idx.Add r, k
        End If
        If NumOk(d(i, EQ_C_BIASABS)) Then _
            somaAbs(k) = somaAbs(k) + CDbl(d(i, EQ_C_BIASABS))
        If NumOk(d(i, EQ_C_BIAS)) Then _
            somaSig(k) = somaSig(k) + CDbl(d(i, EQ_C_BIAS))
        cnt(k) = cnt(k) + 1
        nAmostras = nAmostras + 1
prox:
    Next i

    If nRod = 0 Then Exit Function
    If md = "N" Then BiasEQ = nAmostras: Exit Function
    If md = "NRODADAS" Then BiasEQ = nRod: Exit Function

    ' ---- etapa 2: media das medias de rodada -------------------------
    acc = 0
    det = ""
    For k = 1 To nRod
        If cnt(k) > 0 Then
            If md = "SIGNED" Then
                mr = somaSig(k) / cnt(k)
            Else
                mr = somaAbs(k) / cnt(k)
            End If
            acc = acc + mr
            det = det & rot(k) & " = " & Format$(mr, "0.0000") & _
                  " (n=" & cnt(k) & ")"
            If k < nRod Then det = det & "  |  "
        End If
    Next k

    If md = "DETALHE" Then
        BiasEQ = det & "   ==>   media das " & nRod & " rodada(s) = " & _
                 Format$(acc / nRod, "0.000000")
        Exit Function
    End If

    BiasEQ = acc / nRod
End Function

' ---------------------------------------------------------------------------
' SDI consolidado.
'
'   modo  "MEDIA"  media dos SDI assinados
'         "MAX"    maior |SDI| -- e este que decide o status
'         "N"      quantas amostras entraram
'
' O SDI mede o desvio em unidades de DP DO GRUPO: (X_lab - media grupo)/SD grupo.
' Ele nao substitui o bias: bias e magnitude relativa, SDI e posicao dentro da
' dispersao do grupo. Os dois respondem perguntas diferentes.
Public Function SDIeq(ByVal analito As String, ByVal anoRef As Variant, _
                      ByVal modo As String, _
                      Optional ByVal provedor As Variant = "", _
                      Optional ByVal rodada As Variant = "") As Variant
    Dim d As Variant, i As Long, anoVig As Long
    Dim soma As Double, n As Long, maxAbs As Double, v As Variant, novo As Boolean

    SDIeq = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco(analito)
    If IsEmpty(d) Then Exit Function

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_SDI)
    If anoVig = -32768 Then Exit Function
    novo = ModoNovo(rodada)

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
        If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
        If Not AnoEntra(CLng(Val(CStr(d(i, EQ_C_ANO)))), anoVig, novo) Then GoTo prox
        v = d(i, EQ_C_SDI)
        If Not NumOk(v) Then GoTo prox
        soma = soma + CDbl(v)
        If Abs(CDbl(v)) > maxAbs Then maxAbs = Abs(CDbl(v))
        n = n + 1
prox:
    Next i

    If n = 0 Then Exit Function
    Select Case UCase$(Trim$(modo))
        Case "MEDIA": SDIeq = soma / n
        Case "MAX":   SDIeq = maxAbs
        Case "N":     SDIeq = n
        Case Else:    SDIeq = SEM_EP
    End Select
End Function

' Status do SDI: nenhuma amostra pode passar de |2|.
'
' O criterio e por PIOR AMOSTRA, nao pela media. Uma rodada com SDI +3 e outra
' com -3 dao media zero e descrevem um desempenho que ninguem aprovaria; e a
' pior que reprova o conjunto.
Public Function StatusSDIeq(ByVal analito As String, ByVal anoRef As Variant, _
                            Optional ByVal provedor As Variant = "", _
                            Optional ByVal rodada As Variant = "") As Variant
    Dim m As Variant, n As Variant
    m = SDIeq(analito, anoRef, "MAX", provedor, rodada)
    If Not IsNumeric(m) Then StatusSDIeq = SEM_EP: Exit Function
    n = SDIeq(analito, anoRef, "N", provedor, rodada)
    If CDbl(m) <= LIM_SDI Then
        StatusSDIeq = "OK (|SDI| max " & Format$(m, "0.00") & " em " & n & " amostra(s))"
    Else
        StatusSDIeq = "FORA (|SDI| max " & Format$(m, "0.00") & " > " & _
                      Format$(LIM_SDI, "0") & ")"
    End If
End Function

' Status dos limites do grupo: o resultado do laboratorio caiu dentro da faixa
' informada pelo provedor em TODAS as amostras?
'
' Limite ausente nao e aprovacao: a amostra entra como NAO AVALIADA e aparece na
' contagem, para nao passar por conforme quem ninguem conferiu.
Public Function StatusLimitesEQ(ByVal analito As String, ByVal anoRef As Variant, _
                                Optional ByVal provedor As Variant = "", _
                                Optional ByVal rodada As Variant = "") As Variant
    Dim d As Variant, i As Long, anoVig As Long
    Dim dentro As Long, fora As Long, semLim As Long
    Dim x As Variant, li As Variant, ls As Variant, novo As Boolean

    StatusLimitesEQ = SEM_EP
    If Len(Trim$(analito)) = 0 Then Exit Function
    d = LerBanco(analito)
    If IsEmpty(d) Then Exit Function

    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_XLAB)
    If anoVig = -32768 Then Exit Function
    novo = ModoNovo(rodada)

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
        If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
        If Not AnoEntra(CLng(Val(CStr(d(i, EQ_C_ANO)))), anoVig, novo) Then GoTo prox
        x = d(i, EQ_C_XLAB)
        If Not NumOk(x) Then GoTo prox
        li = d(i, EQ_C_LIMINF)
        ls = d(i, EQ_C_LIMSUP)
        ' limite VAZIO e limite ausente (NAO AVALIADA), nunca 0 -- antes IsNumeric(Empty) contava a
        ' amostra como "fora dos limites" (0..0), ao contrario do que este cabecalho promete
        If Not NumOk(li) Or Not NumOk(ls) Then
            semLim = semLim + 1
        ElseIf CDbl(x) < CDbl(li) Or CDbl(x) > CDbl(ls) Then
            fora = fora + 1
        Else
            dentro = dentro + 1
        End If
prox:
    Next i

    If dentro + fora + semLim = 0 Then Exit Function
    If fora > 0 Then
        StatusLimitesEQ = "NAO OK (" & fora & " fora dos limites)"
    ElseIf dentro = 0 Then
        ' revisao 07/10/2026: nenhuma amostra com limite (ex.: CAP NAO AVALIADO em IG/NRBC) -- nao e "OK"
        StatusLimitesEQ = "NAO AVALIADO (" & semLim & " sem limite do provedor)"
    ElseIf semLim > 0 Then
        StatusLimitesEQ = "OK (" & dentro & " dentro; " & semLim & " sem limite)"
    Else
        StatusLimitesEQ = "OK (" & dentro & " dentro dos limites)"
    End If
End Function

' ---------------------------------------------------------------------------
' Rastreabilidade: memoria de calculo, para auditoria.
Public Function BiasEQMemoria(ByVal analito As String, ByVal anoRef As Variant, _
                              Optional ByVal provedor As Variant = "", _
                              Optional ByVal rodada As Variant = "") As String
    Dim d As Variant, i As Long, anoVig As Long, s As String, n As Long
    Dim somaAbs As Double, somaSig As Double, novo As Boolean

    On Error GoTo falhou
    d = LerBanco(analito)
    If IsEmpty(d) Then
        BiasEQMemoria = IIf(mEQExiste, "sem rodada utilizavel", "EQA_Base ausente")
        Exit Function
    End If
    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_BIASABS)
    If anoVig = -32768 Then BiasEQMemoria = "sem rodada utilizavel": Exit Function
    novo = ModoNovo(rodada)

    For i = 1 To UBound(d, 1)
        If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
        If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
        If Not AnoEntra(CLng(Val(CStr(d(i, EQ_C_ANO)))), anoVig, novo) Then GoTo prox
        If Not NumOk(d(i, EQ_C_BIAS)) Or Not NumOk(d(i, EQ_C_BIASABS)) Then GoTo prox
        n = n + 1
        somaSig = somaSig + CDbl(d(i, EQ_C_BIAS))
        somaAbs = somaAbs + CDbl(d(i, EQ_C_BIASABS))
        s = s & TxtEQ(d(i, EQ_C_PROVEDOR)) & "/" & TxtEQ(d(i, EQ_C_RODADA)) & _
                "|ano=" & CStr(CLng(Val(CStr(d(i, EQ_C_ANO))))) & _
                "|Xlab=" & Format$(d(i, EQ_C_XLAB), "0.####") & _
                "|Xref=" & Format$(d(i, EQ_C_XREF), "0.####") & _
                "|SDI=" & Format$(d(i, EQ_C_SDI), "0.##") & _
                "|bias=" & Format$(d(i, EQ_C_BIAS), "0.####") & ";"
prox:
    Next i
    If n = 0 Then BiasEQMemoria = "sem rodada utilizavel": Exit Function
    BiasEQMemoria = s & " CONSOLIDADO n=" & n & _
                    " media|bias|=" & Format$(somaAbs / n, "0.######") & _
                    " mediaAssinada=" & Format$(somaSig / n, "0.######")
    Exit Function
falhou:
    BiasEQMemoria = "memoria indisponivel (" & Err.Description & ")"
End Function

' Resumo do filtro em uso, para a propria aba dizer o que esta usando (Estatistica!K5).
' ADR-071: descreve o modo e CONTA, pela mesma regra do BiasEQ, as linhas analiticas (analito
' canonico, |bias| numerico de verdade) e as rodadas que casaram -- texto montado aqui (CStr), nunca
' TEXTO(...;"0,00") na planilha (o defeito pt-BR de AT14). Erro de digitacao nao passa calado:
' texto invalido diz "INVALIDA" e item do conjunto sem dado aparece como "sem dado".
' Revisao 07/10/2026: a contagem e feita ANALITO A ANALITO com o recorte que o BiasEQ daquele analito
' usa -- o ano vigente de cada um nos modos de antes (AnoVigente, o mesmo da coluna G) e as ULTIMAS n
' DELE. Antes o K5 tomava um ano vigente GLOBAL: com a 1a rodada de 2026 chegando so para parte dos
' analitos, dizia "ano 2026 | 1 rodada" enquanto os outros seguiam calculando com 2025. Quando os anos
' diferem o texto diz "ano vigente por analito"; em ULTIMAS n diz quantos analitos tem menos de n.
Public Function ResumoFiltroEQ(ByVal provedor As Variant, ByVal ano As Variant, _
                               ByVal rodada As Variant) As String
    Dim p As String, r As String, a As String, md As Long, okT As Boolean, sTxt As String
    Dim key As String, k As Variant, d As Variant, i As Long, an As String, aRef As Long
    Dim anoVig As Long, aMin As Long, aMax As Long, ano_ As Long, nLin As Long
    Dim rods As Object, rodsAn As Object, anosVig As Object, semDado As String, prov As Variant, nItens As Long
    Dim nAnal As Long, nCurto As Long, ch As String
    On Error GoTo falhou
    If IsObject(provedor) Then prov = provedor.Value Else prov = provedor
    If IsObject(ano) Then ano = ano.Value
    sTxt = TextoRodada(rodada, okT)
    GarantirEQ
    If Len(mEQErro) > 0 Then
        ResumoFiltroEQ = "Bias do EP: " & mEQErro & ": SEM EP | 0 rodada(s), 0 linha(s) analiticas"
        Exit Function
    End If
    key = UCase$(Trim$(TxtEQ(prov))) & "|" & Trim$(TxtEQ(ano)) & "|" & sTxt & "|" & okT
    If Not mResumoCache Is Nothing Then
        If mResumoCache.Exists(key) Then ResumoFiltroEQ = mResumoCache(key): Exit Function
    End If
    md = ModoRodada(rodada)
    p = IIf(Livre(prov), "todos os provedores", Trim$(TxtEQ(prov)))
    aRef = TetoAno(ano)
    If md = MR_INVALIDO Then
        If okT Then
            r = "rodada INVALIDA (" & mRodErro & "): SEM EP"
        Else
            r = "rodada INVALIDA (celula com erro): SEM EP"
        End If
        ResumoFiltroEQ = "Bias do EP: " & p & " | " & r & " | 0 rodada(s), 0 linha(s) analiticas"
        GoTo guardar
    End If
    If md >= MR_CONJUNTO And Livre(prov) Then
        ResumoFiltroEQ = "Bias do EP: provedor nao definido -- nos modos conjunto/ULTIMAS/ACUMULADAS o provedor e " & _
                         "obrigatorio (CAP e Controllab nunca se misturam): SEM EP | 0 rodada(s), 0 linha(s) analiticas"
        GoTo guardar
    End If
    Set rods = CreateObject("Scripting.Dictionary")
    Set anosVig = CreateObject("Scripting.Dictionary")
    aMin = 32767
    aMax = -32768
    For Each k In mEQPorAn.Keys
        d = mEQPorAn(k)
        an = CStr(k)
        anoVig = -32768
        ' modos de antes: o ano vigente DESTE analito -- o mesmo que o BiasEQ dele usa
        If md < MR_CONJUNTO Then anoVig = AnoVigente(d, an, ano, prov, rodada, EQ_C_BIASABS)
        Set rodsAn = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(d, 1)
            If Not Casa(d, i, an, prov, rodada, ano) Then GoTo prox
            If Not NumOk(d(i, EQ_C_ANO)) Or Not NumOk(d(i, EQ_C_BIASABS)) Then GoTo prox
            ano_ = CLng(Val(CStr(d(i, EQ_C_ANO))))
            If md < MR_CONJUNTO Then
                If ano_ <> anoVig Then GoTo prox
            ElseIf md <> MR_CONJUNTO Then
                If ano_ > aRef Then GoTo prox
            End If
            nLin = nLin + 1
            ch = ChaveRodada(d, i)
            rods(ch) = 1
            rodsAn(ch) = 1
            If ano_ < aMin Then aMin = ano_
            If ano_ > aMax Then aMax = ano_
prox:
        Next i
        If rodsAn.Count > 0 Then
            nAnal = nAnal + 1
            If md < MR_CONJUNTO Then anosVig(anoVig) = 1
            If md = MR_ULTIMAS And rodsAn.Count < mRodN Then nCurto = nCurto + 1
        End If
    Next k
    If nLin = 0 Then
        a = ""
    ElseIf md < MR_CONJUNTO And anosVig.Count > 1 Then
        a = "ano vigente por analito (" & aMin & " a " & aMax & ")"
    ElseIf aMin = aMax Then
        a = "ano " & aMax
    Else
        a = "anos " & aMin & " a " & aMax
    End If
    Select Case md
        Case MR_LIVRE
            r = IIf(nLin = 0, "ano vigente ate " & IIf(aRef = 32767, "o mais recente", CStr(aRef)), a) & _
                " | media de todas as rodadas"
        Case MR_ROTULO
            r = IIf(nLin = 0, "ate " & IIf(aRef = 32767, "o ano mais recente", CStr(aRef)), a) & _
                " | rodada " & Trim$(CStr(sTxt))
        Case MR_CONJUNTO
            For Each k In mRodConj.Keys
                nItens = nItens + 1
                If Not rods.Exists(k) Then semDado = semDado & IIf(Len(semDado) > 0, "; ", "") & k
            Next k
            r = "conjunto de " & nItens & " rodada(s) informada(s), " & (nItens - CountSemDado(semDado)) & " com dado" & _
                IIf(Len(semDado) > 0, " -- SEM DADO: " & semDado, "") & IIf(nLin > 0, " | " & a, "")
        Case MR_ACUM
            r = "todas as rodadas acumuladas ate " & IIf(aRef = 32767, "o ano mais recente", CStr(aRef)) & _
                IIf(nLin > 0, " (" & a & ")", "")
        Case MR_ULTIMAS
            r = "ultimas " & mRodN & " rodada(s) de cada analito ate " & IIf(aRef = 32767, "o ano mais recente", CStr(aRef)) & _
                IIf(nCurto > 0, " (" & nCurto & " de " & nAnal & " analito(s) com menos de " & mRodN & " com dado)", "") & _
                IIf(nLin > 0, " (" & a & ")", "")
    End Select
    ResumoFiltroEQ = "Bias do EP: " & p & " | " & r & " | " & rods.Count & " rodada(s), " & nLin & " linha(s) analiticas"
guardar:
    If Not mResumoCache Is Nothing Then mResumoCache(key) = ResumoFiltroEQ
    Exit Function
falhou:
    ResumoFiltroEQ = "Bias do EP: resumo indisponivel (" & Err.Description & ")"
End Function

Private Function CountSemDado(ByVal s As String) As Long
    If Len(s) = 0 Then Exit Function
    CountSemDado = UBound(Split(s, "; ")) + 1
End Function

' ===========================================================================
' ADR-064: VIES PELO CEQ PARA A INCERTEZA DE MEDICAO
'
' O BiasEQ ("ABS") e a media dos |bias| -- serve ao modelo de ERRO TOTAL (ET e
' Sigma, ADR-035), mas NAO e vies no sentido do VIM: com vies verdadeiro zero, a
' media de |bias| ainda da ~0,8 x a dispersao das amostras (media de |X| de uma
' normal = sigma x raiz(2/pi)). Para a incerteza o que vale e o bias COM SINAL de
' cada amostra. ADR-067 (05/10/2026): a incerteza e Nordtest TR 537 (CIQ + CEQ,
' Magnusson 2012): u(bias) = raiz(RMS^2 + u(Cref)^2) ENTRA em uc. A TRIAGEM continua
' como monitoramento: vies relevante se investiga/corrige, alem de entrar em U.
'
' Mesma selecao do BiasEQ: provedor, maior ano <= ano EP (AnoVigente), rodada,
' Uso_Analitico = SIM -- CAP e Controllab nunca se misturam (o provedor e filtro).
' TRIAGEM, nao teste formal: amostras da mesma rodada compartilham calibracao e
' cobrem concentracoes diferentes, entao o EP sai subestimado. Por isso exige
' >= 6 amostras de >= 2 rodadas (abaixo disso: "nao verificavel").
'
'   "MEDIA"    media do bias com sinal (%)          "DP"      DP dos bias (%)
'   "EP"       DP / raiz(m)                         "N"       amostras    "NRODADAS"  rodadas
'   "TRIAGEM"  texto; com bperm (vies permitido %, opcional) diz se e relevante
'   "RMS"      raiz(media dos bias^2)            (Nordtest)
'   "UCREF"    media de 100 x DP grupo / |alvo| / raiz(n labs)   (Nordtest)
'   "UBIAS"    raiz(RMS^2 + UCREF^2)              (Nordtest; so com o criterio da triagem)
'   "NREF"     amostras com u(Cref) (DP do grupo e n de laboratorios)
'   "SITUACAO" "" se u(bias) e estimavel; senao o motivo (para a coluna Situacao)
'   "ANO"      ano vigente usado
' Sem amostra utilizavel devolve SEM_EP -- nunca 0.
Public Function ViesEQ(ByVal analito As String, ByVal anoRef As Variant, ByVal modo As String, _
                       Optional ByVal provedor As Variant = "", _
                       Optional ByVal rodada As Variant = "", _
                       Optional ByVal bperm As Variant = "") As Variant
    ' Amostras que entram (auditoria 04/10/2026):
    '  - avaliadas pelo provedor (Status_Padronizado <> NAO AVALIADO): a nao avaliada costuma ser a
    '    concentracao fora da faixa avaliavel ("See Note");
    '  - resultado e alvo diferentes de zero: perto do limite de deteccao (NRBC, bilirrubina direta)
    '    o bias relativo vira +-100% a 150% e nao descreve o metodo;
    '  - cada CHAVE uma vez: a consolidacao mantem a linha digitada em dobro (e so conta).
    ' Janela: o ano vigente (<= ano EP); se ele nao atinge o criterio (>= 6 amostras de >= 2 rodadas
    ' -- tipico em janeiro, com a 1a rodada do ano), inclui o ano anterior (ate 24 meses).
    Dim d As Variant, i As Long, anoVig As Long, md As String, ano As Long, passo As Long, anoMin As Long
    Dim n As Long, b As Double, soma As Double, somaQ As Double, rods As Object, vistos As Object
    Dim nRef As Long, somaRef As Double, nLab As Variant, sdg As Double, alvo As Double, xlab As Double
    Dim media As Double, dp As Double, ep As Double, rms As Double, ucref As Double, ok As Boolean, ch As String

    md = UCase$(Trim$(modo))
    ' ADR-071 (D3): a incerteza NAO segue a selecao nova de rodadas (conjunto, ULTIMAS, ACUMULADAS
    ' ou texto invalido) -- continua na janela do ADR-067 (ano vigente + anterior, >= 6 amostras de
    ' >= 2 rodadas). TODAS e rotulo unico seguem como antes (ADR-032).
    If ModoRodada(rodada) >= MR_CONJUNTO Or ModoRodada(rodada) = MR_INVALIDO Then rodada = "TODAS"
    ViesEQ = SEM_EP
    If md = "TRIAGEM" Then ViesEQ = "não verificável (sem CEQ)"   ' analito sem nenhuma amostra tambem
    If md = "SITUACAO" Then ViesEQ = "sem CEQ para o analito"
    If Len(Trim$(analito)) = 0 Then Exit Function
    On Error GoTo falhou
    d = LerBanco(analito)
    If IsEmpty(d) Then Exit Function
    anoVig = AnoVigente(d, analito, anoRef, provedor, rodada, EQ_C_BIAS)
    If anoVig = -32768 Then Exit Function

    For passo = 0 To 1
        anoMin = anoVig - passo
        n = 0: soma = 0: somaQ = 0: nRef = 0: somaRef = 0
        Set rods = CreateObject("Scripting.Dictionary")
        Set vistos = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(d, 1)
            If Not Casa(d, i, analito, provedor, rodada, anoRef) Then GoTo prox
            If Not NumOk(d(i, EQ_C_ANO)) Then GoTo prox
            ano = CLng(Val(CStr(d(i, EQ_C_ANO))))
            If ano > anoVig Or ano < anoMin Then GoTo prox
            If Igual(d(i, EQ_C_STATUS), "NAO AVALIADO") Then GoTo prox
            If Not NumOk(d(i, EQ_C_BIAS)) Then GoTo prox
            If Not NumOk(d(i, EQ_C_XLAB)) Or Not NumOk(d(i, EQ_C_XREF)) Then GoTo prox
            xlab = CDbl(d(i, EQ_C_XLAB))
            alvo = CDbl(d(i, EQ_C_XREF))
            If xlab = 0 Or alvo = 0 Then GoTo prox
            ch = UCase$(Trim$(TxtEQ(d(i, EQ_C_CHAVE))))
            If Len(ch) > 0 Then
                If vistos.Exists(ch) Then GoTo prox
                vistos(ch) = 1
            End If
            b = CDbl(d(i, EQ_C_BIAS))
            n = n + 1
            soma = soma + b
            somaQ = somaQ + b * b
            rods(CStr(ano) & "|" & UCase$(Trim$(TxtEQ(d(i, EQ_C_RODADA))))) = 1
            If NumOk(d(i, EQ_C_SDGRUPO)) Then
                sdg = CDbl(d(i, EQ_C_SDGRUPO))
                nLab = mEQA.NLabsPorChave(TxtEQ(d(i, EQ_C_CHAVE)))
                If Not IsEmpty(nLab) And sdg >= 0 Then
                    If CDbl(nLab) > 0 Then
                        somaRef = somaRef + 100# * sdg / Abs(alvo) / Sqr(CDbl(nLab))
                        nRef = nRef + 1
                    End If
                End If
            End If
prox:
        Next i
        ok = (n >= VIES_M_MIN And rods.Count >= VIES_ROD_MIN)
        If ok Then Exit For
    Next passo
    If passo > 1 Then passo = 1
    If md = "ANO" Then
        If passo = 0 Then ViesEQ = anoVig Else ViesEQ = CStr(anoVig - 1) & "-" & CStr(anoVig)
        Exit Function
    End If
    If n = 0 Then
        If md = "SITUACAO" Then ViesEQ = "sem CEQ utilizável para o analito"
        Exit Function
    End If

    media = soma / n
    If n >= 2 Then
        dp = somaQ - n * media * media
        If dp < 0 Then dp = 0
        dp = Sqr(dp / (n - 1))
        ep = dp / Sqr(n)
    End If
    rms = Sqr(somaQ / n)
    If nRef > 0 Then ucref = somaRef / nRef

    Select Case md
        Case "MEDIA":    ViesEQ = media
        Case "DP":       If n >= 2 Then ViesEQ = dp
        Case "EP":       If n >= 2 Then ViesEQ = ep
        Case "N":        ViesEQ = n
        Case "NRODADAS": ViesEQ = rods.Count
        Case "TRIAGEM"
            If Not ok Then
                ViesEQ = "não verificável (" & n & " amostra(s), " & rods.Count & " rodada(s))"
            ElseIf Abs(media) <= 2# * ep Then
                ViesEQ = "não detectável"
            ElseIf IsNumeric(bperm) And Len(Trim$(CStr(bperm))) > 0 Then
                If CDbl(bperm) > 0 And Abs(media) > CDbl(bperm) Then
                    ViesEQ = "RELEVANTE: investigar/corrigir"
                Else
                    ViesEQ = "detectável, dentro do permitido"
                End If
            Else
                ViesEQ = "detectável (relevância não avaliada)"
            End If
        Case "RMS":      ViesEQ = rms
        Case "UCREF":    If nRef > 0 Then ViesEQ = ucref
        Case "NREF":     ViesEQ = nRef
        Case "SITUACAO"
            If Not ok Then
                ViesEQ = "CEQ insuficiente (" & n & " amostra(s), " & rods.Count & " rodada(s); mín. " & _
                         VIES_M_MIN & " de " & VIES_ROD_MIN & ")"
            ElseIf nRef = 0 Then
                ViesEQ = "sem u(Cref): faltam DP do grupo e nº de laboratórios no CEQ"
            Else
                ViesEQ = ""
            End If
        Case "UBIAS"
            If ok And nRef > 0 Then ViesEQ = Sqr(rms * rms + ucref * ucref)
        Case Else:       ViesEQ = CVErr(xlErrValue)
    End Select
    Exit Function
falhou:
    ViesEQ = CVErr(xlErrValue)
End Function
