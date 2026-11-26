Attribute VB_Name = "mAlvosRef"
Option Explicit
' CAMADA: Lotes / Painel -- ADR-102 (25/11/2026)
'
' ============================================================================
'  1. ALVOS DO LABORATORIO POR LOTE E PERIODO (manual)
'     O analista escolhe no rodape do Painel o LOTE (B98), DE (E98) e ATE (G98); o analito e o do Painel
'     (selAnalito). Por nivel: N, Media e DP (amostral, n-1, mEstatistica.CalcularMedia/CalcularDP) dos
'     resultados ATIVOS (PARTICIPA_ESTATISTICA = SIM) do lote e do analito exatos, de TODOS os equipamentos,
'     com resultado numerico e DATA entre as datas (inclusive; vazio = sem limite). Sem exclusao de outlier.
'     So grava no botao "Aplicar Alvos ao Lote", com confirmacao; o valor fica ESTATICO na LotesStore C:H
'     (mLotes.GravarParametroStore), origem "MANUAL ...". Regra do N (Cfg_Integracao ALVO_N_RECOMENDADO e
'     ALVO_N_IDEAL, padrao 20 e 90): < 2 bloqueia; 2..19 so com a ciencia de TEMPORARIO; 20..89 aceitavel;
'     >= 90 ideal. Trava: a chave LOTE|ANALITO existe UMA vez na LotesStore. Auditoria por nivel em AA:AO.
'     Adendo 25/11: o botao pergunta o que transferir (1 = Media e DP, 2 = so a Media, 3 = so o DP; cancelar nao
'     grava); so o campo escolhido e gravado e a escolha fica na origem W ([MEDIA+DP] / [MEDIA] / [DP]) e no Audit_Log.
'  1b. EXCECAO TEMPORARIA (linhas 105-111): Media da bula e DP do lote anterior (ou da bula) digitados. Celulas
'     TRANCADAS (cinza) ate o botao "Liberar digitacao manual"; a liberacao vale para UMA chave
'     LOTE|ANALITO nesta sessao (troca de lote/analito, abrir ou fechar a pasta trancam de novo). "Gravar valores
'     manuais" grava so os campos preenchidos, origem "EXCECAO TEMPORARIA ...", CONDICAO_Nk = EXCECAO_TEMPORARIA.
'     A coluna Origem atual da apuracao fica amarela enquanto a origem W for temporaria OU algum CONDICAO_Nk (AA:AO)
'     tiver marca temporaria. A marca e POR CAMPO (MarcasDaCondicao / CondicaoDasMarcas): uma transferencia parcial
'     so tira a marca do campo transferido (ex.: ACEITAVEL+EXCECAO_TEMPORARIA_DP = Media propria, DP da excecao).
'  2. REFERENCIA DA BULA / UNITY (informativa; NUNCA grava na LotesStore)
'     Grade do Painel (linhas 115-154) = vista da RefBulaStore para o lote da secao: Media, LI e LS por nivel.
'     Todo resultado ATIVO abaixo do LI ou acima do LS vira uma linha na RefBulaObs (chave ID|ANALITO|Nk|LI|LS,
'     PENDENTE / CIENTE / RETIRADA). So observacao NOVA alerta. PENDENTE que deixou de estar fora vira
'     RETIRADA; CIENTE nunca e apagada. O parecer do analista (RefBulaParecer) da ciencia as pendentes do
'     lote. Nada aqui muda status, Westgard, LJ ou liberacao.
' ============================================================================
Private Const ABA_P As String = "Painel"
Private Const ABA_LS As String = "LotesStore"
Private Const ABA_REF As String = "RefBulaStore"
Private Const ABA_OBS As String = "RefBulaObs"
Private Const ABA_PAR As String = "RefBulaParecer"
Private Const MODULO As String = "mAlvosRef"
Private Const TITULO_CX As String = "Alvos e referencia"

' ---- layout do Painel (igual nos dois produtos; N3 so na Hematologia) ----
Private Const R_TIT As Long = 97
Private Const R_ENT As Long = 98
Private Const R_NV1 As Long = 100
Private Const R_REGRA As Long = 103
Private Const R_MAN_TIT As Long = 105           ' excecao temporaria: titulo (A:J) + situacao (K:U) + marca (Z)
Private Const R_MAN1 As Long = 107              ' N1..N3 nas linhas 107..109 (cabecalho na 106)
Private Const R_MAN_JUST As Long = 110          ' justificativa (B110:U110)
Private Const R_REF1 As Long = 115
Private Const N_REF As Long = 40
Private Const R_OBS_TIT As Long = 156
Private Const R_OBS1 As Long = 158
Private Const N_OBS_TELA As Long = 15
Private Const R_OBS_TOT As Long = 173
Private Const R_PAR_TXT As Long = 176
Private Const R_PAR_ULT As Long = 181
Private Const DESLOC_DIG As Long = 13             ' ADR-103: tabela de digitacao da bula = grade + 13 colunas (N114)
Private Const R_DIG_CAB As Long = 114
Private Const C_MAN_MED As Long = 2             ' B107.. Media (bula)
Private Const C_MAN_DP As Long = 3              ' C107.. DP
Private Const C_MAN_ORIG As Long = 4            ' D107.. Origem do DP (lista Lote anterior / Bula)
Private Const C_MAN_LOTE As Long = 5            ' E107.. Lote anterior (E:F mescladas)
Private Const C_MAN_INFO As Long = 7            ' G107.. informacao (G:U mescladas)
Private Const C_MAN_STATUS As Long = 11         ' K105 (K105:U105) situacao da digitacao manual
Private Const C_MAN_JUST As Long = 2            ' B110 (B110:U110)
Private Const C_MARCA As Long = 26              ' Z105 marca da liberacao: LIBERADO|LOTE|ANALITO|quem|quando
Private Const MIN_JUST As Long = 10
Private Const C_LOTE As Long = 2                ' B98 (B98:C98)
Private Const C_DE As Long = 5                  ' E98
Private Const C_ATE As Long = 7                 ' G98
Private Const C_APURADO As Long = 26            ' Z98 analito da ultima apuracao (a linha 100.. se refere a ele)
Private Const C_OBS_STATUS As Long = 7          ' G156 (G156:U156) situacao das observacoes do lote
Private Const C_ULT As Long = 21                ' U

' ---- LotesStore (mLotes) ----
Private Const LS_R0 As Long = 2
Private Const LS_C_CHAVE As Long = 16           ' P
Private Const LS_C_ORIGEM As Long = 23          ' W
Private Const LS_C_TESTE1 As Long = 24          ' X..Z
Private Const LS_C_AUD1 As Long = 27            ' AA..AO: 5 colunas por nivel
Private Const LS_AUD_N As Long = 5

' ---- RefBulaStore ----
Private Const RB_C_LOTE As Long = 1
Private Const RB_C_AN As Long = 2
Private Const RB_C_CHAVE As Long = 3
Private Const RB_C_V1 As Long = 4               ' D..L: N1 Media, LI, LS, N2 ..., N3 ...
Private Const RB_C_FONTE As Long = 13
Private Const RB_C_QUANDO As Long = 14
Private Const RB_C_QUEM As Long = 15
Private Const RB_NCOL As Long = 15

' ---- RefBulaObs ----
Private Const OB_C_CHAVE As Long = 1
Private Const OB_C_ID As Long = 2
Private Const OB_C_LOTE As Long = 3
Private Const OB_C_AN As Long = 4
Private Const OB_C_NV As Long = 5
Private Const OB_C_DATA As Long = 6
Private Const OB_C_RES As Long = 7
Private Const OB_C_LI As Long = 8
Private Const OB_C_LS As Long = 9
Private Const OB_C_SIT As Long = 10
Private Const OB_C_ST As Long = 11
Private Const OB_C_CRIADA As Long = 12
Private Const OB_C_STEM As Long = 13
Private Const OB_C_STPOR As Long = 14
Private Const OB_C_PARECER As Long = 15
Private Const OB_C_EQ As Long = 16
Private Const OB_C_RUN As Long = 17
Private Const OB_NCOL As Long = 17

' ---- RefBulaParecer ----
Private Const PA_C_ID As Long = 1
Private Const PA_C_DATA As Long = 2
Private Const PA_C_QUEM As Long = 3
Private Const PA_C_LOTE As Long = 4
Private Const PA_C_NOBS As Long = 5
Private Const PA_C_TEXTO As Long = 6
Private Const PA_NCOL As Long = 6

Private Const ST_PEND As String = "PENDENTE"
Private Const ST_CIENTE As String = "CIENTE"
Private Const ST_RETIRADA As String = "RETIRADA"
Private Const MIN_PARECER As Long = 10
Private Const ORIG_LOTE_ANT As String = "Lote anterior"
Private Const ORIG_BULA As String = "Bula"
Private Const COND_EXCECAO As String = "EXCECAO_TEMPORARIA"

Private mRespostas As Collection          ' gancho do QA: respostas das perguntas (SIM / NAO / 1 / 2 / 3), na ordem
Private mUltimoAviso As String
Private mUltimaPergunta As String
Private mDentro As Boolean
Private mLiberado As String               ' chave LOTE|ANALITO com a digitacao manual liberada nesta sessao ("" = trancada)
Private mSilencioso As Boolean            ' ADR-103: laco de varios analitos (sem caixa por analito)

' ============================ GANCHOS DO QA ============================
' "SIM|NAO|..." responde as proximas perguntas, na ordem; vazio limpa. Sem resposta guardada e sem gente olhando
' (Application.Interactive = False) a resposta e NAO: nada e gravado.
Public Sub DefinirRespostasTeste(ByVal respostas As String)
    Dim p As Variant
    Set mRespostas = Nothing
    If Len(Trim$(respostas)) = 0 Then Exit Sub
    Set mRespostas = New Collection
    For Each p In Split(respostas, "|")
        mRespostas.Add UCase$(Trim$(CStr(p)))
    Next p
End Sub

Public Function UltimoAviso() As String
    UltimoAviso = mUltimoAviso
End Function

Public Function UltimaPergunta() As String
    UltimaPergunta = mUltimaPergunta
End Function

Private Function Perguntar(ByVal msg As String, ByVal estilo As VbMsgBoxStyle) As Boolean
    Dim r As String
    mUltimaPergunta = msg
    If Not mRespostas Is Nothing Then
        If mRespostas.Count > 0 Then
            r = mRespostas(1)
            mRespostas.Remove 1
            Perguntar = (r = "SIM")
            Exit Function
        End If
    End If
    If Not Application.Interactive Then Exit Function
    Perguntar = (MsgBox(msg, estilo + vbYesNo + vbDefaultButton2, TITULO_CX) = vbYes)
End Function

' Adendo ADR-102: o que transferir. 1 = Media e DP (padrao), 2 = so a Media, 3 = so o DP; 0 = cancelado (Cancelar,
' vazio ou outra resposta). Sem resposta guardada e sem gente olhando: 0 (nada e gravado).
Private Function PerguntarModo(ByVal msg As String) As Long
    Dim r As String
    mUltimaPergunta = msg
    If Not mRespostas Is Nothing Then
        If mRespostas.Count > 0 Then
            r = mRespostas(1)
            mRespostas.Remove 1
            GoTo decidir
        End If
    End If
    If Not Application.Interactive Then Exit Function
    r = Trim$(InputBox(msg, TITULO_CX, "1"))
decidir:
    Select Case r
        Case "1": PerguntarModo = 1
        Case "2": PerguntarModo = 2
        Case "3": PerguntarModo = 3
    End Select
End Function

Private Function ModoTxt(ByVal modo As Long) As String
    Select Case modo
        Case 1: ModoTxt = "MEDIA+DP"
        Case 2: ModoTxt = "MEDIA"
        Case 3: ModoTxt = "DP"
    End Select
End Function

' O campo j (1 = Media, 2 = DP) entra no modo escolhido?
Private Function CampoNoModo(ByVal modo As Long, ByVal j As Long) As Boolean
    CampoNoModo = (modo = 1) Or (modo = 2 And j = 1) Or (modo = 3 And j = 2)
End Function

' Origem temporaria (aviso amarelo na apuracao): excecao manual ou alvo com N abaixo do recomendado.
Public Function OrigemTemporaria(ByVal origem As String) As Boolean
    Dim u As String
    u = UCase$(origem)
    OrigemTemporaria = (InStr(u, "EXCECAO TEMPORARIA") > 0 Or InStr(u, "TEMPORARIO") > 0)
End Function

' Marca temporaria POR CAMPO guardada no CONDICAO_Nk (AA:AO). tM / tD: "" = valor proprio (CIQ) ou sem registro;
' "EXCECAO" = digitado na excecao temporaria; "TEMPORARIO" = CIQ com N abaixo do recomendado.
' Codificacao: EXCECAO_TEMPORARIA / TEMPORARIO = os dois campos; senao tokens unidos por "+": a condicao propria
' (ACEITAVEL / IDEAL) e <marca>_MEDIA / <marca>_DP, ex.: ACEITAVEL+EXCECAO_TEMPORARIA_DP.
Private Sub MarcasDaCondicao(ByVal cond As String, ByRef tM As String, ByRef tD As String)
    Dim p() As String, i As Long, t As String
    tM = "": tD = ""
    cond = UCase$(Trim$(cond))
    If Len(cond) = 0 Then Exit Sub
    If cond = COND_EXCECAO Then tM = "EXCECAO": tD = "EXCECAO": Exit Sub
    If cond = "TEMPORARIO" Then tM = "TEMPORARIO": tD = "TEMPORARIO": Exit Sub
    p = Split(cond, "+")
    For i = 0 To UBound(p)
        t = Trim$(p(i))
        If t = COND_EXCECAO & "_MEDIA" Then tM = "EXCECAO"
        If t = COND_EXCECAO & "_DP" Then tD = "EXCECAO"
        If t = "TEMPORARIO_MEDIA" Then tM = "TEMPORARIO"
        If t = "TEMPORARIO_DP" Then tD = "TEMPORARIO"
    Next i
End Sub

' Condicao propria (ACEITAVEL / IDEAL) contida no CONDICAO_Nk; "" se nao houver.
Private Function CondPropria(ByVal cond As String) As String
    Dim p() As String, i As Long, t As String
    cond = UCase$(Trim$(cond))
    If Len(cond) = 0 Then Exit Function
    p = Split(cond, "+")
    For i = 0 To UBound(p)
        t = Trim$(p(i))
        If t = "ACEITAVEL" Or t = "IDEAL" Then CondPropria = t: Exit Function
    Next i
End Function

Private Function MarcaCond(ByVal t As String) As String
    If t = "EXCECAO" Then MarcaCond = COND_EXCECAO Else MarcaCond = t
End Function

Private Function MarcaOrigem(ByVal t As String) As String
    If t = "EXCECAO" Then MarcaOrigem = "EXCECAO TEMPORARIA" Else MarcaOrigem = t
End Function

' Inverso de MarcasDaCondicao. propria = condicao do(s) campo(s) sem marca (ACEITAVEL / IDEAL / "").
Private Function CondicaoDasMarcas(ByVal tM As String, ByVal tD As String, ByVal propria As String) As String
    Dim s As String
    If tM = tD Then
        If Len(tM) = 0 Then CondicaoDasMarcas = propria Else CondicaoDasMarcas = MarcaCond(tM)
        Exit Function
    End If
    If Len(propria) > 0 And (Len(tM) = 0 Or Len(tD) = 0) Then s = propria
    If Len(tM) > 0 Then s = s & IIf(Len(s) > 0, "+", "") & MarcaCond(tM) & "_MEDIA"
    If Len(tD) > 0 Then s = s & IIf(Len(s) > 0, "+", "") & MarcaCond(tD) & "_DP"
    CondicaoDasMarcas = s
End Function

' Para o QA: "tM|tD|propria" do CONDICAO_Nk dado.
Public Function MarcasTexto(ByVal cond As String) As String
    Dim tM As String, tD As String
    MarcasDaCondicao cond, tM, tD
    MarcasTexto = tM & "|" & tD & "|" & CondPropria(cond)
End Function

' MsgBox so com gente olhando (sob automacao um modal invisivel trava o Excel); o texto fica para o QA.
Private Sub Avisar(ByVal msg As String, ByVal estilo As VbMsgBoxStyle)
    mUltimoAviso = msg
    If mSilencioso Then Exit Sub                ' ADR-103: dentro do laco de varios analitos o aviso e um so, no fim
    If Application.Interactive Then MsgBox msg, estilo, TITULO_CX
End Sub

' ============================ AUXILIARES ============================
Private Function NLV() As Long
    NLV = mEstatistica.NLV
End Function

Private Function Pn() As Worksheet
    Set Pn = ThisWorkbook.Worksheets(ABA_P)
End Function

' A secao foi montada pelo instalar_adr102 (titulo na A97)?
Public Function SecaoInstalada() As Boolean
    On Error Resume Next
    SecaoInstalada = (UCase$(Left$(Trim$(CStr(Pn().Cells(R_TIT, 1).Value)), 5)) = "ALVOS")
End Function

Private Function AbaSeExiste(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set AbaSeExiste = ThisWorkbook.Worksheets(nome)
End Function

' Aba tecnica oculta: criada se faltar (com a estrutura da pasta livre; travada, o erro diz para rodar o instalador).
Private Function Aba(ByVal nome As String) As Worksheet
    Dim ws As Worksheet, cab As Variant, nE As Long
    Set ws = AbaSeExiste(nome)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        nE = Err.Number
        If nE = 0 Then ws.Name = nome
        If nE = 0 Then nE = Err.Number
        If nE = 0 Then ws.Visible = xlSheetHidden
        On Error GoTo 0
        If ws Is Nothing Or nE <> 0 Then
            Err.Raise vbObjectError + 1021, MODULO, "a aba " & nome & " nao existe e nao pode ser criada " & _
                      "(estrutura da pasta protegida). Rode o instalar_adr102."
        End If
    End If
    If Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then
        Select Case nome
            Case ABA_REF
                cab = Array("LOTE", "ANALITO", "CHAVE", "N1_MEDIA", "N1_LI", "N1_LS", "N2_MEDIA", "N2_LI", "N2_LS", _
                            "N3_MEDIA", "N3_LI", "N3_LS", "FONTE", "REGISTRADO_EM", "REGISTRADO_POR")
            Case ABA_OBS
                cab = Array("CHAVE", "ID_REGISTRO", "LOTE", "ANALITO", "NIVEL", "DATA_HORA", "RESULTADO", "LI", "LS", _
                            "SITUACAO", "STATUS", "CRIADA_EM", "STATUS_EM", "STATUS_POR", "ID_PARECER", _
                            "EQUIPAMENTO", "RUN")
            Case ABA_PAR
                cab = Array("ID_PARECER", "DATA_HORA", "USUARIO", "LOTE", "N_OBSERVACOES", "PARECER")
        End Select
        If IsArray(cab) Then
            Dim prot As Boolean
            prot = mSeguranca.LiberarEscritaRapida(ws)
            ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(cab) + 1)).Value = cab
            ws.Rows(1).Font.Bold = True
            mSeguranca.RestaurarProtecao ws, prot
        End If
    End If
    Set Aba = ws
End Function

Private Function UltimaLinha(ByVal ws As Worksheet, ByVal c As Long) As Long
    UltimaLinha = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
End Function

Private Function NumOk(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbString, vbBoolean: Exit Function
    End Select
    NumOk = IsNumeric(v)
End Function

' texto de chave sem depender da lingua do Windows (Str$ usa sempre o ponto)
Private Function NumChave(ByVal v As Variant) As String
    If NumOk(v) Then NumChave = Trim$(Str$(CDbl(v)))
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function NumTxt(ByVal v As Variant) As String
    If NumOk(v) Then NumTxt = Format$(CDbl(v), "0.0###") Else NumTxt = "-"
End Function

Private Function DataTxt(ByVal d As Double) As String
    If d > 0 Then DataTxt = Format$(d, "dd/mm/yyyy") Else DataTxt = "?"
End Function

' Data de uma celula de entrada: 0 = vazia (sem limite); ok = False quando ha algo que nao e data.
Private Function LerData(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = True
    If IsError(v) Then ok = False: Exit Function
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If VarType(v) = vbDate Or NumOk(v) Then
        LerData = Int(CDbl(v))
    ElseIf IsDate(v) Then
        LerData = Int(CDbl(CDate(v)))
    Else
        ok = False
    End If
    If LerData < 1 And ok Then ok = False
End Function

' Login da sessao (DoLogin); sem sessao, o usuario do Windows.
Private Function Quem() As String
    Dim s As String
    On Error Resume Next
    If mSeguranca.SessaoAtiva() Then s = mAuditoria.UsuarioSistema()
    On Error GoTo 0
    If Len(s) = 0 Then
        s = Environ$("USERNAME")
        If Len(s) = 0 Then s = "(desconhecido)"
        s = s & " (Windows)"
    End If
    Quem = s
End Function

Private Function LoteSecao() As String
    LoteSecao = mLotes.LoteCanonico(Txt(Pn().Cells(R_ENT, C_LOTE).Value))
End Function

Private Function AnalitoPainel() As String
    On Error Resume Next
    AnalitoPainel = Txt(ThisWorkbook.Names("selAnalito").RefersToRange.Value)
End Function

Private Function CfgNum(ByVal chave As String, ByVal padrao As Long) As Long
    Dim s As String
    On Error Resume Next
    s = mIntegracao.Cfg(chave)
    On Error GoTo 0
    If Len(s) > 0 And IsNumeric(s) Then CfgNum = CLng(Val(Replace(s, ",", "."))) Else CfgNum = padrao
    If CfgNum < 2 Then CfgNum = padrao
End Function

' Limites da regra do N: Cfg_Integracao ALVO_N_RECOMENDADO / ALVO_N_IDEAL (padrao 20 / 90).
Private Sub LimitesN(ByRef rec As Long, ByRef ideal As Long)
    rec = CfgNum("ALVO_N_RECOMENDADO", 20)
    ideal = CfgNum("ALVO_N_IDEAL", 90)
    If ideal < rec Then ideal = rec
End Sub

' Para o instalador e o QA: "rec|ideal" que valem agora.
Public Function LimitesNTexto() As String
    Dim rec As Long, ideal As Long
    LimitesN rec, ideal
    LimitesNTexto = rec & "|" & ideal
End Function

' BLOQUEADO (< 2) | TEMPORARIO (2..rec-1) | ACEITAVEL (rec..ideal-1) | IDEAL (>= ideal)
Public Function CondicaoN(ByVal n As Long) As String
    Dim rec As Long, ideal As Long
    LimitesN rec, ideal
    If n < 2 Then
        CondicaoN = "BLOQUEADO"
    ElseIf n < rec Then
        CondicaoN = "TEMPORARIO"
    ElseIf n < ideal Then
        CondicaoN = "ACEITAVEL"
    Else
        CondicaoN = "IDEAL"
    End If
End Function

' Condicao do nivel para gravar: a do N, ou DP_ZERO quando N >= 2 e todos os resultados sao iguais (DP = 0).
' DP <= 0 tiraria o nivel do LJ/Westgard (mLotes.ParametrosLote); o nivel e mantido, como no BLOQUEADO.
Private Function CondicaoNivel(ByVal n As Long, ByVal dp As Double) As String
    If n >= 2 And dp <= 0 Then
        CondicaoNivel = "DP_ZERO"
    Else
        CondicaoNivel = CondicaoN(n)
    End If
End Function

' BLOQUEADO ou DP_ZERO: o nivel nao e gravado (Media/DP, origem do nivel, bit TESTE e AA:AO ficam como estao).
Private Function Mantido(ByVal cond As String) As Boolean
    Mantido = (cond = "BLOQUEADO" Or cond = "DP_ZERO")
End Function

Private Function TextoCondicao(ByVal cond As String) As String
    Dim rec As Long, ideal As Long
    LimitesN rec, ideal
    Select Case cond
        Case "BLOQUEADO": TextoCondicao = "BLOQUEADO: N < 2, nao e alterado"
        Case "DP_ZERO": TextoCondicao = "BLOQUEADO: todos os resultados iguais (DP = 0), nao e alterado"
        Case "TEMPORARIO": TextoCondicao = "NÃO RECOMENDADO (N < " & rec & "): só como TEMPORÁRIO"
        Case "ACEITAVEL": TextoCondicao = "ACEITÁVEL (o ideal é N >= " & ideal & ")"
        Case "IDEAL": TextoCondicao = "IDEAL (N >= " & ideal & ")"
    End Select
End Function

' ============================ APURACAO ============================
' N, Media e DP por nivel dos resultados ATIVOS do lote x analito (todos os equipamentos), resultado numerico,
' DATA entre dIni e dFim (dias inteiros, inclusive; 0 = sem limite). dMin/dMax = datas do 1o e do ultimo usado.
Private Sub Apurar(ByVal dados As Variant, ByVal lote As String, ByVal analito As String, _
                   ByVal dIni As Double, ByVal dFim As Double, ByRef n() As Long, ByRef m() As Double, _
                   ByRef s() As Double, ByRef dMin() As Double, ByRef dMax() As Double)
    Dim nl As Long, nv As Long, i As Long, cnt As Long, dd As Double, v() As Double, uL As String, uA As String
    nl = NLV()
    ReDim n(1 To 3): ReDim m(1 To 3): ReDim s(1 To 3): ReDim dMin(1 To 3): ReDim dMax(1 To 3)
    If IsEmpty(dados) Then Exit Sub
    uL = UCase$(Trim$(lote)): uA = UCase$(Trim$(analito))
    ReDim v(1 To UBound(dados, 1))
    For nv = 1 To nl
        cnt = 0
        For i = 1 To UBound(dados, 1)
            If CStr(dados(i, mDados.COL_STATUS)) <> mDados.ST_ATIVO Then GoTo prox
            If UCase$(Trim$(CStr(dados(i, mDados.COL_LOTE)))) <> uL Then GoTo prox
            If UCase$(Trim$(CStr(dados(i, mDados.COL_ANALITO)))) <> uA Then GoTo prox
            If IsError(dados(i, mDados.COL_NIVEL)) Then GoTo prox
            If Val(CStr(dados(i, mDados.COL_NIVEL))) <> nv Then GoTo prox
            If Not NumOk(dados(i, mDados.COL_RESULT)) Then GoTo prox
            dd = 0
            If Not IsError(dados(i, mDados.COL_DATA)) Then
                If IsDate(dados(i, mDados.COL_DATA)) Then dd = Int(CDbl(CDate(dados(i, mDados.COL_DATA))))
            End If
            If dIni > 0 Or dFim > 0 Then
                If dd = 0 Then GoTo prox
                If dIni > 0 And dd < dIni Then GoTo prox
                If dFim > 0 And dd > dFim Then GoTo prox
            End If
            cnt = cnt + 1
            v(cnt) = CDbl(dados(i, mDados.COL_RESULT))
            If dd > 0 Then
                If dMin(nv) = 0 Or dd < dMin(nv) Then dMin(nv) = dd
                If dd > dMax(nv) Then dMax(nv) = dd
            End If
prox:
        Next i
        n(nv) = cnt
        If cnt >= 1 Then m(nv) = mEstatistica.CalcularMedia(v, cnt)
        If cnt >= 2 Then s(nv) = mEstatistica.CalcularDP(v, cnt, m(nv))
    Next nv
End Sub

' Para o QA (oraculo independente): "n1;media1;dp1|n2;..." da apuracao com estes parametros (datas em numero de
' serie; 0 = sem limite). Nao grava nada.
Public Function ApuracaoTexto(ByVal lote As String, ByVal analito As String, ByVal dIni As Double, _
                              ByVal dFim As Double) As String
    Dim n() As Long, m() As Double, s() As Double, d0() As Double, d1() As Double, nv As Long, out As String
    Apurar mDados.CarregarDB(), lote, analito, Int(dIni), Int(dFim), n, m, s, d0, d1
    For nv = 1 To NLV()
        out = out & IIf(nv > 1, "|", "") & n(nv) & ";" & Trim$(Str$(m(nv))) & ";" & Trim$(Str$(s(nv)))
    Next nv
    ApuracaoTexto = out
End Function

' Linha da chave exata LOTE|ANALITO na LotesStore (coluna P) e quantas vezes ela aparece.
Private Function LinhaStore(ByVal lote As String, ByVal analito As String, ByRef vezes As Long) As Long
    Dim ws As Worksheet, ult As Long, v As Variant, i As Long, k As String
    vezes = 0
    Set ws = ThisWorkbook.Worksheets(ABA_LS)
    ult = UltimaLinha(ws, LS_C_CHAVE)
    If ult < LS_R0 Then Exit Function
    k = mLotes.ChaveLote(lote, analito)
    v = ws.Range(ws.Cells(LS_R0, LS_C_CHAVE), ws.Cells(ult + 1, LS_C_CHAVE)).Value
    For i = 1 To UBound(v, 1) - 1
        If UCase$(Txt(v(i, 1))) = k Then
            vezes = vezes + 1
            If LinhaStore = 0 Then LinhaStore = LS_R0 + i - 1
        End If
    Next i
End Function

Private Sub GarantirCabecalhoAuditoria(ByVal ws As Worksheet)
    Dim k As Long, nomes As Variant, j As Long
    nomes = Array("N_USADO_N", "DATA_INI_N", "DATA_FIM_N", "CONDICAO_N", "ATUALIZADO_EM_N")
    For k = 1 To 3
        For j = 0 To LS_AUD_N - 1
            If Len(Txt(ws.Cells(1, LS_C_AUD1 + (k - 1) * LS_AUD_N + j).Value)) = 0 Then
                ws.Cells(1, LS_C_AUD1 + (k - 1) * LS_AUD_N + j).Value = nomes(j) & k
            End If
        Next j
    Next k
End Sub

' ============================ VISTA: ALVOS (linhas 98-103) ============================
Public Sub AtualizarVistaAlvos(Optional ByVal dados As Variant)
    Dim ws As Worksheet, lote As String, an As String, ok1 As Boolean, ok2 As Boolean, dIni As Double, dFim As Double
    Dim n() As Long, m() As Double, s() As Double, d0() As Double, d1() As Double, nv As Long, r As Long
    Dim prot As Boolean, ev As Boolean, out() As Variant, rLs As Long, vezes As Long, ls As Worksheet
    Dim rec As Long, ideal As Long, aviso As String, j As Long, temp As Boolean, orig As Range
    Dim tM As String, tD As String
    If Not SecaoInstalada() Then Exit Sub
    Set ws = Pn()
    lote = LoteSecao()
    an = AnalitoPainel()
    dIni = LerData(ws.Cells(R_ENT, C_DE).Value, ok1)
    dFim = LerData(ws.Cells(R_ENT, C_ATE).Value, ok2)
    ReDim out(1 To NLV(), 1 To 9)
    If Len(lote) = 0 Then
        aviso = "Escolha o lote em B98 para apurar."
    ElseIf Len(an) = 0 Then
        aviso = "Escolha o analito no Painel (C3)."
    ElseIf Not (ok1 And ok2) Then
        aviso = "Datas De/Ate (E98/G98) invalidas."
    ElseIf dIni > 0 And dFim > 0 And dIni > dFim Then
        aviso = "A data De (E98) e posterior a Ate (G98)."
    Else
        If IsMissing(dados) Then dados = mDados.CarregarDB()
        Apurar dados, lote, an, dIni, dFim, n, m, s, d0, d1
        Set ls = ThisWorkbook.Worksheets(ABA_LS)
        rLs = LinhaStore(lote, an, vezes)
        For nv = 1 To NLV()
            out(nv, 1) = "N" & nv
            out(nv, 2) = n(nv)
            If n(nv) >= 1 Then out(nv, 3) = m(nv)
            If n(nv) >= 2 Then
                out(nv, 4) = s(nv)
                If s(nv) > 0 And m(nv) <> 0 Then out(nv, 5) = mEstatistica.CalcularCV(s(nv), m(nv))
            End If
            out(nv, 6) = TextoCondicao(CondicaoNivel(n(nv), s(nv)))
            For j = 0 To 1
                If NumOk(mLotes.ValorDoLote(lote, an, 2 * nv - 1 + j)) Then out(nv, 7 + j) = mLotes.ValorDoLote(lote, an, 2 * nv - 1 + j)
            Next j
            If vezes = 0 Then
                out(nv, 9) = "(sem registro na LotesStore)"
            ElseIf vezes > 1 Then
                out(nv, 9) = "(chave duplicada na LotesStore)"
            Else
                out(nv, 9) = Txt(ls.Cells(rLs, LS_C_ORIGEM).Value)
            End If
        Next nv
        ' temporario: origem W (uma so para a linha) OU marca por campo no CONDICAO_Nk (AA:AO) -- outras gravacoes
        ' (bloco da aba Analitos) reescrevem W sem mexer em AA:AO, e a marca nao pode sumir com isso
        If vezes = 1 Then
            temp = OrigemTemporaria(Txt(ls.Cells(rLs, LS_C_ORIGEM).Value))
            For nv = 1 To NLV()
                MarcasDaCondicao Txt(ls.Cells(rLs, LS_C_AUD1 + (nv - 1) * LS_AUD_N + 3).Value), tM, tD
                If Len(tM) > 0 Or Len(tD) > 0 Then temp = True
            Next nv
            If temp Then
                For nv = 1 To NLV()
                    out(nv, 9) = "temporário: substituir pelos valores do laboratório -- " & out(nv, 9)
                Next nv
            End If
        End If
    End If
    LimitesN rec, ideal
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscritaRapida(ws)
    For nv = 1 To NLV()
        r = R_NV1 + nv - 1
        If Len(aviso) > 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, C_ULT)).ClearContents   ' A:U contem a mesclagem I:U inteira
            ws.Cells(r, 1).Value = "N" & nv
            If nv = 1 Then ws.Cells(r, 6).Value = aviso
        Else
            For j = 1 To 9
                ws.Cells(r, j).Value = out(nv, j)
            Next j
        End If
        ' Origem atual (I:U, mesclada) em amarelo enquanto o alvo for temporario (excecao manual ou N baixo)
        Set orig = ws.Cells(r, 9).MergeArea
        If temp And Len(aviso) = 0 Then
            orig.Interior.Color = RGB(255, 235, 156)
            orig.Font.Color = RGB(156, 87, 0)
            orig.Font.Bold = True
        Else
            orig.Interior.Pattern = xlNone
            orig.Font.ColorIndex = xlAutomatic
            orig.Font.Bold = False
        End If
    Next nv
    ws.Range(ws.Cells(R_NV1, 5), ws.Cells(R_NV1 + NLV() - 1, 5)).NumberFormat = "0.00"
    ws.Cells(R_ENT, C_APURADO).Value = IIf(Len(aviso) > 0, "", an)
    ws.Cells(R_REGRA, 1).Value = "Regra do N: < 2 bloqueia · 2 a " & (rec - 1) & " NÃO RECOMENDADO (grava só como " & _
        "TEMPORÁRIO, com ciência) · " & rec & " a " & (ideal - 1) & " aceitável · >= " & ideal & " ideal.  Conta: " & _
        "resultados ativos (PARTICIPA = SIM) do lote e do analito, todos os equipamentos, DATA entre De e Até " & _
        "(inclusive; vazio = sem limite); DP amostral (n-1), sem exclusão de pontos. Só grava no botão." & _
        IIf(Len(aviso) = 0, "  Apurado para " & an & " em " & Format$(Now, "dd/mm/yyyy hh:nn") & ".", "")
fim:
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub

' ============================ APLICAR ALVOS AO LOTE ============================
' Botao "Aplicar Alvos ao Lote".
Public Sub AplicarAlvosAoLote()
    Dim r As String
    r = AplicarAlvos()
End Sub

' Nucleo do botao (o QA chama este, com DefinirRespostasTeste). "OK|niveis gravados|campos alterados|condicoes",
' "CANCELADO|motivo" ou "ERRO|motivo". Nada e gravado fora do OK.
' ADR-103 (varios analitos): sem argumentos e o botao de UM analito, exatamente como antes. Com anLote (chamado SO
' pelo AplicarAlvosVarios): o analito vem do argumento, o modo de modoLote, a base de dadosLote; nao ha caixa por
' analito (a confirmacao e a ciencia do TEMPORARIO foram feitas uma vez, para todos: temporarioAceito), nem
' InicioUsuario/fim, nem recalculo da tela (feitos uma vez no laco). A apuracao, a trava, a gravacao, a origem, o
' AA:AO e o Audit_Log sao os MESMOS, por analito.
Public Function AplicarAlvos(Optional ByVal anLote As String = "", Optional ByVal modoLote As Long = 0, _
                             Optional ByVal dadosLote As Variant, _
                             Optional ByVal temporarioAceito As Boolean = False) As String
    Dim ws As Worksheet, lote As String, an As String, ok1 As Boolean, ok2 As Boolean, dIni As Double, dFim As Double
    Dim n() As Long, m() As Double, s() As Double, d0() As Double, d1() As Double, nv As Long, dados As Variant
    Dim vezes As Long, rLs As Long, cond(1 To 3) As String, nGrav As Long, temTemp As Boolean, msg As String
    Dim antM As Variant, antS As Variant, rec As Long, ideal As Long, periodo As String, tmp As String
    Dim origem As String, partes As String, agora As Date, quemTxt As String, esp As Variant, nAlt As Long
    Dim ls As Worksheet, prot As Boolean, ev As Boolean, k As Long, conds As String, nE As Long, sE As String
    Dim pIni(1 To 3) As Double, pFim(1 To 3) As Double, temAceit As Boolean, dentroOp As Boolean
    Dim modo As Long, wAnt As String, herda As String, bitT As Long, condAnt As String
    Dim tM As String, tD As String, novaM As String, novaD As String, condNova(1 To 3) As String
    Dim keepM As Boolean, keepD As Boolean, emLote As Boolean
    emLote = (Len(Trim$(anLote)) > 0)
    If Not SecaoInstalada() Then AplicarAlvos = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    Set ws = Pn()
    lote = LoteSecao()
    If emLote Then an = Trim$(anLote) Else an = AnalitoPainel()
    If Len(lote) = 0 Or UCase$(lote) = "TODOS" Then
        AplicarAlvos = Falhar("Escolha o LOTE em B98 (obrigatório). Nada foi gravado.")
        Exit Function
    End If
    If Len(an) = 0 Then AplicarAlvos = Falhar("Escolha o analito no Painel (C3). Nada foi gravado."): Exit Function
    dIni = LerData(ws.Cells(R_ENT, C_DE).Value, ok1)
    dFim = LerData(ws.Cells(R_ENT, C_ATE).Value, ok2)
    If Not (ok1 And ok2) Then
        AplicarAlvos = Falhar("As datas De (E98) e Até (G98) precisam ser datas (ou ficar vazias). Nada foi gravado.")
        Exit Function
    End If
    If dIni > 0 And dFim > 0 And dIni > dFim Then
        AplicarAlvos = Falhar("A data De (E98) é posterior a Até (G98). Nada foi gravado.")
        Exit Function
    End If
    ' trava por lote: a chave exata, uma vez so
    rLs = LinhaStore(lote, an, vezes)
    If vezes = 0 Then
        AplicarAlvos = Falhar("O lote " & lote & " não tem registro de " & an & " na LotesStore (chave " & _
            mLotes.ChaveLote(lote, an) & "). Confira o lote ou cadastre Média/DP no bloco do lote (aba Analitos). " & _
            "Nada foi gravado.")
        Exit Function
    ElseIf vezes > 1 Then
        AplicarAlvos = Falhar("A chave " & mLotes.ChaveLote(lote, an) & " aparece " & vezes & " vezes na LotesStore " & _
            "(duplicada). Corrija a LotesStore antes de aplicar. Nada foi gravado.")
        Exit Function
    End If
    On Error GoTo falha
    If emLote And Not IsMissing(dadosLote) Then dados = dadosLote Else dados = mDados.CarregarDB()
    Apurar dados, lote, an, dIni, dFim, n, m, s, d0, d1
    LimitesN rec, ideal
    For nv = 1 To NLV()
        cond(nv) = CondicaoNivel(n(nv), s(nv))
        If Not Mantido(cond(nv)) Then
            nGrav = nGrav + 1
            If cond(nv) = "TEMPORARIO" Then temTemp = True
            If cond(nv) = "ACEITAVEL" Then temAceit = True
            If dIni > 0 Then pIni(nv) = dIni Else pIni(nv) = d0(nv)
            If dFim > 0 Then pFim(nv) = dFim Else pFim(nv) = d1(nv)
        End If
        conds = conds & IIf(nv > 1, ";", "") & "N" & nv & "=" & cond(nv)
    Next nv
    periodo = IIf(dIni > 0, DataTxt(dIni), "início") & " a " & IIf(dFim > 0, DataTxt(dFim), "hoje")
    If nGrav = 0 Then
        If Not emLote Then AtualizarVistaAlvos dados
        AplicarAlvos = Falhar("Impossível calcular o DP: nenhum nível tem ao menos 2 resultados ativos do lote " & lote & _
            " (" & an & ") no período " & periodo & " com DP maior que zero. Nada foi gravado.")
        Exit Function
    End If
    ' pergunta do adendo: o que transferir (o campo nao escolhido fica como esta)
    msg = "O que transferir para o lote " & lote & " (" & an & ")?" & vbCrLf & vbCrLf & _
          "1 = Média e DP" & vbCrLf & "2 = só a Média (o DP atual fica como está)" & vbCrLf & _
          "3 = só o DP (a Média atual fica como está)" & vbCrLf & vbCrLf & "Digite 1, 2 ou 3 (Cancelar não grava nada)."
    If emLote Then modo = modoLote Else modo = PerguntarModo(msg)
    If modo = 0 Then
        AtualizarVistaAlvos dados
        AplicarAlvos = "CANCELADO|transferencia"
        mUltimoAviso = "Cancelado: nada foi gravado (nenhuma opção 1, 2 ou 3 escolhida)."
        Exit Function
    End If
    ' confirmacao 1: o que muda
    msg = "Aplicar os alvos do laboratório ao lote " & lote & "?" & vbCrLf & vbCrLf & _
          "Analito: " & an & vbCrLf & "Período: " & periodo & " (todos os equipamentos; só resultados ativos)" & vbCrLf & _
          "Transferir: " & Choose(modo, "Média e DP", "SÓ A MÉDIA (o DP fica como está)", _
                                  "SÓ O DP (a Média fica como está)") & vbCrLf
    For nv = 1 To NLV()
        msg = msg & vbCrLf & "N" & nv & ": N = " & n(nv) & " -- " & TextoCondicao(cond(nv))
        If Not Mantido(cond(nv)) Then
            msg = msg & vbCrLf & "    "
            If CampoNoModo(modo, 1) Then msg = msg & " Média " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv - 1)) & _
                                               " -> " & NumTxt(m(nv)) & "    "
            If CampoNoModo(modo, 2) Then msg = msg & " DP " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv)) & _
                                               " -> " & NumTxt(s(nv))
        End If
    Next nv
    msg = msg & vbCrLf & vbCrLf & "Os valores ficam gravados (estáticos) na LotesStore: mudar o lote ou as datas " & _
          "depois não muda o alvo gravado."
    If emLote Then
        ' confirmacao feita uma vez para todos os analitos (AplicarAlvosVarios)
    ElseIf Not Perguntar(msg, vbQuestion) Then
        AtualizarVistaAlvos dados
        AplicarAlvos = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: nada foi gravado."
        Exit Function
    End If
    ' confirmacao 2: N abaixo do recomendado = TEMPORARIO, com ciencia
    If temTemp And emLote Then
        If Not temporarioAceito Then
            AplicarAlvos = "CANCELADO|temporario"
            mUltimoAviso = "Cancelado: nada foi gravado em " & an & " (alvo temporário não confirmado)."
            Exit Function
        End If
    ElseIf temTemp Then
        tmp = ""
        For nv = 1 To NLV()
            If cond(nv) = "TEMPORARIO" Then tmp = tmp & IIf(Len(tmp) > 0, ", ", "") & "N" & nv & " (n = " & n(nv) & ")"
        Next nv
        msg = "ATENÇÃO: N abaixo de " & rec & " em " & tmp & "." & vbCrLf & vbCrLf & _
              "Com menos de " & rec & " resultados o alvo NÃO É RECOMENDADO: a média e o DP ainda são instáveis. " & _
              "O ideal é N >= " & ideal & "." & vbCrLf & vbCrLf & _
              "Declaro que estou ciente de que este alvo é TEMPORÁRIO e deve ser refeito quando houver resultados " & _
              "suficientes. Gravar como TEMPORÁRIO?"
        If Not Perguntar(msg, vbExclamation) Then
            AtualizarVistaAlvos dados
            AplicarAlvos = "CANCELADO|temporario"
            mUltimoAviso = "Cancelado: nada foi gravado (alvo temporário não confirmado)."
            Exit Function
        End If
    End If
    ' ---- gravacao ----
    agora = Now
    quemTxt = Quem()
    For nv = 1 To NLV()
        If Mantido(cond(nv)) Then
            partes = partes & IIf(Len(partes) > 0, " | ", "") & "N" & nv & " mantido (n=" & n(nv) & _
                     IIf(cond(nv) = "DP_ZERO", "; DP=0", "") & ")"
        Else
            partes = partes & IIf(Len(partes) > 0, " | ", "") & "N" & nv & " n=" & n(nv) & " (" & DataTxt(pIni(nv)) & _
                     " a " & DataTxt(pFim(nv)) & ")" & IIf(cond(nv) = "TEMPORARIO", " TEMPORARIO", "")
        End If
    Next nv
    origem = "MANUAL " & Format$(agora, "dd/mm/yyyy hh:nn") & " " & quemTxt & ": " & partes & " [" & ModoTxt(modo) & "]"
    ' a origem W e uma so para a linha: o que NAO foi trocado (campo fora do modo ou nivel mantido) e era temporario
    ' (marca do campo no CONDICAO_Nk anterior, AA:AO) continua temporario -- a marca POR CAMPO e mantida no
    ' CONDICAO_Nk novo e herdada na origem; o aviso amarelo so sai quando tudo o que era temporario for substituido
    Set ls = ThisWorkbook.Worksheets(ABA_LS)
    For nv = 1 To NLV()
        condAnt = Txt(ls.Cells(rLs, LS_C_AUD1 + (nv - 1) * LS_AUD_N + 3).Value)
        MarcasDaCondicao condAnt, tM, tD
        If Mantido(cond(nv)) Then
            condNova(nv) = condAnt                                     ' nivel nao e gravado
        Else
            novaM = tM: novaD = tD
            If cond(nv) = "TEMPORARIO" Then herda = "TEMPORARIO" Else herda = ""
            If CampoNoModo(modo, 1) Then novaM = herda
            If CampoNoModo(modo, 2) Then novaD = herda
            ' condicao propria: a do N quando um campo transferido e proprio; senao a que o campo mantido ja tinha
            If Len(herda) = 0 Then herda = cond(nv) Else herda = CondPropria(condAnt)
            condNova(nv) = CondicaoDasMarcas(novaM, novaD, herda)
        End If
        keepM = (Mantido(cond(nv)) Or Not CampoNoModo(modo, 1)) And Len(tM) > 0
        keepD = (Mantido(cond(nv)) Or Not CampoNoModo(modo, 2)) And Len(tD) > 0
        tmp = ""
        If keepM And keepD And tM = tD Then
            tmp = "Media e DP (" & MarcaOrigem(tM) & ")"
        Else
            If keepM Then tmp = "Media (" & MarcaOrigem(tM) & ")"
            If keepD Then tmp = tmp & IIf(Len(tmp) > 0, ", ", "") & "DP (" & MarcaOrigem(tD) & ")"
        End If
        If Len(tmp) > 0 Then wAnt = wAnt & IIf(Len(wAnt) > 0, ", ", "") & "N" & nv & " " & tmp
    Next nv
    If Len(wAnt) > 0 Then origem = origem & "; mantido do registro anterior: " & wAnt
    If Not emLote Then
        mApp.InicioUsuario "Gravando os alvos do lote " & lote & "..."
        dentroOp = True
    End If
    esp = mLotesBlocos.EspecDoLote(lote, an)
    For nv = 1 To NLV()
        If Not Mantido(cond(nv)) Then
            If CampoNoModo(modo, 1) Then antM = mLotes.GravarParametroStore(lote, an, 2 * nv - 1, m(nv), esp, origem)
            If CampoNoModo(modo, 2) Then antS = mLotes.GravarParametroStore(lote, an, 2 * nv, s(nv), esp, origem)
            If CampoNoModo(modo, 1) And CStr(antM) <> CStr(m(nv)) Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_CALCULADO", MODULO, 0, Empty, "", lote, nv, an, _
                    antM, m(nv), "", "", "Media N" & nv, "Media N" & nv & " do lote " & lote & " = media dos " & n(nv) & _
                    " resultados ativos de " & DataTxt(pIni(nv)) & " a " & DataTxt(pFim(nv)) & " (" & cond(nv) & "; " & _
                    ModoTxt(modo) & "; " & quemTxt & ")."
                nAlt = nAlt + 1
            End If
            If CampoNoModo(modo, 2) And CStr(antS) <> CStr(s(nv)) Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_CALCULADO", MODULO, 0, Empty, "", lote, nv, an, _
                    antS, s(nv), "", "", "DP N" & nv, "DP N" & nv & " do lote " & lote & " = DP amostral (n-1) dos " & n(nv) & _
                    " resultados ativos de " & DataTxt(pIni(nv)) & " a " & DataTxt(pFim(nv)) & " (" & cond(nv) & "; " & _
                    ModoTxt(modo) & "; " & quemTxt & ")."
                nAlt = nAlt + 1
            End If
        End If
    Next nv
    ' origem (W) e bits TESTE (X..Z) mesmo quando o valor era igual (GravarParametroStore so mexe neles quando muda);
    ' auditoria do nivel em AA:AO. Bit TESTE: so o do campo transferido sai (bit 1 = Media, bit 2 = DP, ADR-078)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(ls)
    GarantirCabecalhoAuditoria ls
    ls.Cells(rLs, LS_C_ORIGEM).Value = origem
    For nv = 1 To NLV()
        If Not Mantido(cond(nv)) Then
            bitT = Val(Txt(ls.Cells(rLs, LS_C_TESTE1 + nv - 1).Value))
            If CampoNoModo(modo, 1) Then bitT = bitT And Not 1
            If CampoNoModo(modo, 2) Then bitT = bitT And Not 2
            If bitT > 0 Then ls.Cells(rLs, LS_C_TESTE1 + nv - 1).Value = bitT Else ls.Cells(rLs, LS_C_TESTE1 + nv - 1).ClearContents
            k = LS_C_AUD1 + (nv - 1) * LS_AUD_N
            ls.Cells(rLs, k).Value = n(nv)
            If pIni(nv) > 0 Then ls.Cells(rLs, k + 1).Value = CDate(pIni(nv)) Else ls.Cells(rLs, k + 1).ClearContents
            If pFim(nv) > 0 Then ls.Cells(rLs, k + 2).Value = CDate(pFim(nv)) Else ls.Cells(rLs, k + 2).ClearContents
            ls.Cells(rLs, k + 3).Value = condNova(nv)                 ' marca por campo (o campo mantido fica marcado)
            ls.Cells(rLs, k + 4).Value = agora
        End If
    Next nv
    mSeguranca.RestaurarProtecao ls, prot
    prot = False
    Application.EnableEvents = ev
    mLotes.InvalidarLotes
    If emLote Then
        AplicarAlvos = "OK|" & nGrav & "|" & nAlt & "|" & conds & "|" & ModoTxt(modo)
        Exit Function
    End If
    If StrComp(lote, mLotes.LoteEmTela(), vbTextCompare) = 0 Then mLotes.CarregarParametrosView lote
    mEstatistica.InvalidarParametros
    mEstatistica.RecalcularAnalitoAtual
    mLotes.AtualizarResumoLote
    mLotesBlocos.EscreverValoresDoLote lote
    AtualizarVistaAlvos dados
    mApp.fim
    dentroOp = False
    msg = "Alvos gravados no lote " & lote & " (" & an & "): " & nGrav & " nível(is), " & nAlt & " valor(es) alterado(s)." & _
          vbCrLf & "Transferido: " & Choose(modo, "Média e DP", "só a Média", "só o DP") & "." & _
          IIf(temTemp, vbCrLf & vbCrLf & "Nível(is) com N < " & rec & " gravado(s) como TEMPORÁRIO: refaça quando houver " & _
              "resultados suficientes.", "") & _
          IIf(temAceit, vbCrLf & vbCrLf & "N aceitável; o ideal é N >= " & ideal & ".", "")
    Avisar msg, IIf(temTemp, vbExclamation, vbInformation)
    AplicarAlvos = "OK|" & nGrav & "|" & nAlt & "|" & conds & "|" & ModoTxt(modo)
    Exit Function
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If Not ls Is Nothing Then mSeguranca.RestaurarProtecao ls, prot
    If dentroOp Then
        mApp.fim
    ElseIf Not emLote Then
        Application.EnableEvents = True
    End If
    mLotes.InvalidarLotes
    On Error GoTo 0
    AplicarAlvos = Falhar("Falha ao aplicar os alvos do lote " & lote & IIf(emLote, " (" & an & ")", "") & ": " & sE)
End Function

' ============================ ADR-103: APLICAR ALVOS A VARIOS ANALITOS ============================
' Botao "Aplicar a vários analitos": a janela frmAlvosLote escolhe os analitos (um a um, todos, desmarcar); o lote e o
' periodo sao os da secao (B98, E98, G98), os mesmos para todos. Cada analito e apurado e gravado pelo MESMO nucleo do
' botao de um analito (AplicarAlvos com anLote), com os seus proprios resultados -- nunca a Media/DP de outro.
Public Sub AplicarAlvosVariosUI()
    Dim f As Object, lista As String, r As String
    If Not SecaoInstalada() Then Exit Sub
    On Error GoTo semJanela
    Set f = VBA.UserForms.Add("frmAlvosLote")
    On Error GoTo 0
    f.Preparar CatalogoAnalitos(), "Lote: " & LoteSecao() & "    Período: " & PeriodoSecaoTxt()
    f.Show
    If f.Confirmado Then lista = f.Selecionados
    Unload f
    If Len(lista) = 0 Then Exit Sub
    r = AplicarAlvosVarios(lista)
    Exit Sub
semJanela:
    Avisar "A janela de seleção de analitos (frmAlvosLote) não está instalada. Nada foi gravado.", vbExclamation
End Sub

' Para o QA: a janela sem Show. acoes = "TODOS", "NENHUM", "+Analito", "-Analito" separadas por "|", aplicadas na
' ordem. Devolve "selecionados<TAB>legenda da contagem<TAB>botao Aplicar habilitado (1/0)" (TAB: ha analito com "#",
' ex. NEUT#). Nada e gravado.
Public Function TesteJanelaVarios(ByVal acoes As String) As String
    Dim f As Object, a As Variant, t As String, ok As Boolean
    Set f = VBA.UserForms.Add("frmAlvosLote")
    f.Preparar CatalogoAnalitos(), "teste"
    For Each a In Split(acoes, "|")
        t = Trim$(CStr(a))
        If UCase$(t) = "TODOS" Then
            f.MarcarTodos True
        ElseIf UCase$(t) = "NENHUM" Then
            f.MarcarTodos False
        ElseIf Left$(t, 1) = "+" Or Left$(t, 1) = "-" Then
            ok = f.Marcar(Mid$(t, 2), Left$(t, 1) = "+")
        End If
    Next a
    TesteJanelaVarios = f.SelecaoAtual() & vbTab & f.Controls("lblQtd").Caption & vbTab & IIf(f.Controls("btnAplicar").Enabled, "1", "0")
    Unload f
End Function

' Analitos do catalogo (aba Especificacoes A4:A43), separados por "|", na ordem da aba.
Public Function CatalogoAnalitos() As String
    Dim wa As Worksheet, v As Variant, i As Long, out As String
    Set wa = ThisWorkbook.Worksheets(mLotes.AbaEspec())
    v = wa.Range(wa.Cells(4, 1), wa.Cells(4 + N_REF - 1, 1)).Value
    For i = 1 To UBound(v, 1)
        If Len(Txt(v(i, 1))) > 0 Then out = out & IIf(Len(out) > 0, "|", "") & Txt(v(i, 1))
    Next i
    CatalogoAnalitos = out
End Function

Private Function PeriodoSecaoTxt() As String
    Dim ok1 As Boolean, ok2 As Boolean, dIni As Double, dFim As Double
    dIni = LerData(Pn().Cells(R_ENT, C_DE).Value, ok1)
    dFim = LerData(Pn().Cells(R_ENT, C_ATE).Value, ok2)
    If Not (ok1 And ok2) Then PeriodoSecaoTxt = "(datas inválidas)": Exit Function
    PeriodoSecaoTxt = IIf(dIni > 0, DataTxt(dIni), "início") & " a " & IIf(dFim > 0, DataTxt(dFim), "hoje")
End Function

' Nucleo do botao de varios analitos (o QA chama este com DefinirRespostasTeste "modo|SIM|SIM"). lista = analitos
' separados por "|". Perguntas: o modo (1/2/3, uma vez), a confirmacao (periodo, modo, quantidade, bloqueados) e, se
' houver nivel com N abaixo do recomendado, UMA ciencia de TEMPORARIO para todos (NAO = os analitos com nivel
' temporario ficam de fora, como no botao de um analito). Devolve "OK|gravados|bloqueados|temporarios fora|erros|
' detalhe (analito=resultado;...)", "CANCELADO|motivo" ou "ERRO|motivo".
Public Function AplicarAlvosVarios(ByVal lista As String) As String
    Dim lote As String, ok1 As Boolean, ok2 As Boolean, dIni As Double, dFim As Double, dados As Variant
    Dim nomes() As String, nA As Long, i As Long, nv As Long, vezes As Long, an As String, p As Variant
    Dim n() As Long, m() As Double, s() As Double, d0() As Double, d1() As Double, cond As String
    Dim estado() As String, temTemp() As Boolean, nTemp As Long, nBloq As Long, nOkPrev As Long
    Dim modo As Long, msg As String, periodo As String, rec As Long, ideal As Long, aceita As Boolean
    Dim r As String, nOk As Long, nErr As Long, nFora As Long, det As String, bloqTxt As String, tempTxt As String
    Dim visto As Object, grav As Boolean, dentroOp As Boolean, sE As String
    If Not SecaoInstalada() Then AplicarAlvosVarios = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    lote = LoteSecao()
    If Len(lote) = 0 Or UCase$(lote) = "TODOS" Then
        AplicarAlvosVarios = Falhar("Escolha o LOTE em B98 (obrigatório). Nada foi gravado.")
        Exit Function
    End If
    dIni = LerData(Pn().Cells(R_ENT, C_DE).Value, ok1)
    dFim = LerData(Pn().Cells(R_ENT, C_ATE).Value, ok2)
    If Not (ok1 And ok2) Then
        AplicarAlvosVarios = Falhar("As datas De (E98) e Até (G98) precisam ser datas (ou ficar vazias). Nada foi gravado.")
        Exit Function
    End If
    If dIni > 0 And dFim > 0 And dIni > dFim Then
        AplicarAlvosVarios = Falhar("A data De (E98) é posterior a Até (G98). Nada foi gravado.")
        Exit Function
    End If
    Set visto = CreateObject("Scripting.Dictionary")
    visto.CompareMode = 1
    ReDim nomes(1 To 1)
    For Each p In Split(lista, "|")
        If Len(Trim$(CStr(p))) > 0 Then
            If Not visto.Exists(Trim$(CStr(p))) Then
                visto.Add Trim$(CStr(p)), 1
                nA = nA + 1
                ReDim Preserve nomes(1 To nA)
                nomes(nA) = Trim$(CStr(p))
            End If
        End If
    Next p
    If nA = 0 Then
        AplicarAlvosVarios = "CANCELADO|nenhum analito"
        mUltimoAviso = "Nenhum analito selecionado: nada foi gravado."
        Exit Function
    End If
    On Error GoTo falha
    dados = mDados.CarregarDB()
    LimitesN rec, ideal
    ReDim estado(1 To nA): ReDim temTemp(1 To nA)
    ' pre-apuracao (nao grava): o que cada analito vai ser
    For i = 1 To nA
        an = nomes(i)
        LinhaStore lote, an, vezes
        If vezes = 0 Then
            estado(i) = "SEM_REGISTRO"
        ElseIf vezes > 1 Then
            estado(i) = "CHAVE_DUPLICADA"
        Else
            Apurar dados, lote, an, dIni, dFim, n, m, s, d0, d1
            grav = False
            For nv = 1 To NLV()
                cond = CondicaoNivel(n(nv), s(nv))
                If Not Mantido(cond) Then
                    grav = True
                    If cond = "TEMPORARIO" Then temTemp(i) = True
                End If
            Next nv
            If Not grav Then
                estado(i) = "BLOQUEADO"
            Else
                estado(i) = "GRAVAVEL"
                nOkPrev = nOkPrev + 1
                If temTemp(i) Then
                    nTemp = nTemp + 1
                    tempTxt = tempTxt & IIf(Len(tempTxt) > 0, ", ", "") & an
                End If
            End If
        End If
        If estado(i) <> "GRAVAVEL" Then
            nBloq = nBloq + 1
            bloqTxt = bloqTxt & IIf(Len(bloqTxt) > 0, ", ", "") & an & " (" & LCase$(Replace(estado(i), "_", " ")) & ")"
        End If
    Next i
    periodo = IIf(dIni > 0, DataTxt(dIni), "início") & " a " & IIf(dFim > 0, DataTxt(dFim), "hoje")
    If nOkPrev = 0 Then
        AplicarAlvosVarios = Falhar("Nenhum dos " & nA & " analito(s) selecionado(s) pode ser gravado no lote " & lote & _
            " no período " & periodo & ": " & bloqTxt & ". Nada foi gravado.")
        Exit Function
    End If
    msg = "O que transferir para o lote " & lote & " nos " & nA & " analito(s) selecionado(s)?" & vbCrLf & vbCrLf & _
          "1 = Média e DP" & vbCrLf & "2 = só a Média (o DP atual fica como está)" & vbCrLf & _
          "3 = só o DP (a Média atual fica como está)" & vbCrLf & vbCrLf & "Digite 1, 2 ou 3 (Cancelar não grava nada)."
    modo = PerguntarModo(msg)
    If modo = 0 Then
        AplicarAlvosVarios = "CANCELADO|transferencia"
        mUltimoAviso = "Cancelado: nada foi gravado (nenhuma opção 1, 2 ou 3 escolhida)."
        Exit Function
    End If
    msg = "Aplicar os alvos do laboratório a VÁRIOS analitos do lote " & lote & "?" & vbCrLf & vbCrLf & _
          "Período (o mesmo para todos): " & periodo & " (todos os equipamentos; só resultados ativos)" & vbCrLf & _
          "Transferir: " & Choose(modo, "Média e DP", "SÓ A MÉDIA (o DP fica como está)", _
                                  "SÓ O DP (a Média fica como está)") & vbCrLf & _
          "Analitos selecionados: " & nA & vbCrLf & "Podem ser gravados: " & nOkPrev & _
          IIf(nTemp > 0, " (" & nTemp & " com nível de N < " & rec & ": TEMPORÁRIO)", "") & vbCrLf & _
          IIf(nBloq > 0, "Ficam de fora (nada é alterado): " & nBloq & " -- " & bloqTxt & vbCrLf, "") & vbCrLf & _
          "Cada analito usa SÓ os seus próprios resultados do período (Média e DP calculados um a um). " & _
          "Os analitos não selecionados não mudam. Os valores ficam gravados (estáticos) na LotesStore."
    If Not Perguntar(msg, vbQuestion) Then
        AplicarAlvosVarios = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: nada foi gravado."
        Exit Function
    End If
    If nTemp > 0 Then
        msg = "ATENÇÃO: N abaixo de " & rec & " em " & nTemp & " analito(s): " & tempTxt & "." & vbCrLf & vbCrLf & _
              "Com menos de " & rec & " resultados o alvo NÃO É RECOMENDADO: a média e o DP ainda são instáveis. " & _
              "O ideal é N >= " & ideal & "." & vbCrLf & vbCrLf & _
              "Declaro que estou ciente de que esses alvos são TEMPORÁRIOS e devem ser refeitos quando houver " & _
              "resultados suficientes. Gravar esses analitos como TEMPORÁRIO?" & vbCrLf & vbCrLf & _
              "(Não = esses " & nTemp & " analito(s) ficam de fora; os demais são gravados.)"
        aceita = Perguntar(msg, vbExclamation)
    End If
    mApp.InicioUsuario "Gravando os alvos do lote " & lote & " (" & nOkPrev & " analitos)..."
    dentroOp = True
    mSilencioso = True
    For i = 1 To nA
        If estado(i) = "GRAVAVEL" Then
            If temTemp(i) And Not aceita Then
                nFora = nFora + 1
                r = "CANCELADO|temporario"
            Else
                r = AplicarAlvos(nomes(i), modo, dados, aceita)
                If Left$(r, 3) = "OK|" Then nOk = nOk + 1 Else nErr = nErr + 1
            End If
        Else
            r = estado(i)
        End If
        det = det & IIf(Len(det) > 0, ";", "") & nomes(i) & "=" & Split(r, "|")(0)
        If Left$(r, 9) = "CANCELADO" Then det = det & ":" & Split(r, "|")(1)
    Next i
    mSilencioso = False
    mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_CALCULADO_VARIOS", MODULO, 0, Empty, "", lote, 0, "", _
        "", nOk & " gravado(s)", "", "", "Varios analitos", "Aplicar alvos a varios analitos do lote " & lote & _
        " (" & periodo & "; " & ModoTxt(modo) & "; " & Quem() & "): " & det
    mLotes.InvalidarLotes
    If StrComp(lote, mLotes.LoteEmTela(), vbTextCompare) = 0 Then mLotes.CarregarParametrosView lote
    mEstatistica.InvalidarParametros
    mEstatistica.RecalcularAnalitoAtual
    mLotes.AtualizarResumoLote
    mLotesBlocos.EscreverValoresDoLote lote
    AtualizarVistaAlvos dados
    mApp.fim
    dentroOp = False
    msg = "Alvos gravados no lote " & lote & " (" & periodo & "): " & nOk & " analito(s)." & vbCrLf & _
          "Transferido: " & Choose(modo, "Média e DP", "só a Média", "só o DP") & "." & _
          IIf(nBloq > 0, vbCrLf & "Fora (nada alterado): " & bloqTxt & ".", "") & _
          IIf(nFora > 0, vbCrLf & "Temporários não confirmados (nada alterado): " & nFora & ".", "") & _
          IIf(nErr > 0, vbCrLf & "Com erro (nada alterado nesses): " & nErr & ". Veja o Audit_Log.", "") & _
          IIf(nTemp > 0 And aceita, vbCrLf & vbCrLf & "Analito(s) com N < " & rec & " gravado(s) como TEMPORÁRIO: " & _
              "refaça quando houver resultados suficientes.", "")
    Avisar msg, IIf(nErr > 0 Or (nTemp > 0 And aceita), vbExclamation, vbInformation)
    AplicarAlvosVarios = "OK|" & nOk & "|" & nBloq & "|" & nFora & "|" & nErr & "|" & det
    Exit Function
falha:
    sE = Err.Description
    mSilencioso = False
    On Error Resume Next
    If dentroOp Then mApp.fim
    mLotes.InvalidarLotes
    On Error GoTo 0
    AplicarAlvosVarios = Falhar("Falha ao aplicar os alvos a vários analitos do lote " & lote & ": " & sE & _
        IIf(Len(det) > 0, " (já processados: " & det & ")", ""))
End Function

Private Function Falhar(ByVal msg As String) As String
    Avisar msg, vbExclamation
    Falhar = "ERRO|" & msg
End Function

' ============================ EXCECAO TEMPORARIA: VALORES MANUAIS (linhas 105-111) ============================
' Regra do laboratorio: Media e DP PROPRIOS (CIQ, "Aplicar Alvos ao Lote"). A digitacao e EXCECAO, so no inicio do
' lote sem valores proprios: Media da bula e DP do lote anterior (ou da bula). As celulas ficam TRANCADAS (cinza) ate
' o botao "Liberar digitacao manual"; a liberacao vale para a chave LOTE|ANALITO do momento,
' nesta sessao (mLiberado + marca em Z105). Gravar, trocar o lote/analito, abrir ou fechar a pasta trancam de novo.
Private Function RangeManual(ByVal ws As Worksheet) As Range
    Set RangeManual = Union(ws.Range(ws.Cells(R_MAN1, C_MAN_MED), ws.Cells(R_MAN1 + NLV() - 1, C_MAN_LOTE + 1)), _
                            ws.Range(ws.Cells(R_MAN_JUST, C_MAN_JUST), ws.Cells(R_MAN_JUST, C_ULT)))
End Function

Private Function TextoTrancado() As String
    TextoTrancado = "TRANCADA (padrão). Exceção: só no início do lote, sem valores próprios -- botão " & _
                    """Liberar digitação manual""."
End Function

' Chave LOTE|ANALITO gravada na marca Z105 ("" sem liberacao).
Private Function ChaveDaMarca() As String
    Dim p() As String, s As String
    s = Txt(Pn().Cells(R_MAN_TIT, C_MARCA).Value)
    If Len(s) = 0 Then Exit Function
    p = Split(s & "|||", "|")
    If p(0) = "LIBERADO" Then ChaveDaMarca = UCase$(p(1) & "|" & p(2))
End Function

' Para o QA: "LIBERADA|chave" ou "TRANCADA".
Public Function EstadoManualTexto() As String
    If Len(mLiberado) > 0 Then EstadoManualTexto = "LIBERADA|" & mLiberado Else EstadoManualTexto = "TRANCADA"
End Function

' Tranca (limpa os campos, cinza, Locked) ou libera (amarelo, destravado) o bloco. Locked so muda com a aba
' DESPROTEGIDA (UserInterfaceOnly nao basta): LiberarEscrita / RestaurarProtecao, tambem no erro.
Private Sub EstadoManual(ByVal liberar As Boolean, ByVal statusTxt As String, ByVal marca As String, _
                         Optional ByVal lote As String = "", Optional ByVal an As String = "")
    Dim ws As Worksheet, prot As Boolean, ev As Boolean, nv As Long, r As Long, rg As Range, nE As Long, sE As String
    Dim st As Range
    Set ws = Pn()
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscrita(ws)
    For nv = 1 To NLV()
        r = R_MAN1 + nv - 1
        ws.Range(ws.Cells(r, C_MAN_MED), ws.Cells(r, C_MAN_ORIG)).ClearContents
        ws.Cells(r, C_MAN_LOTE).MergeArea.ClearContents
        ws.Cells(r, C_MAN_INFO).MergeArea.ClearContents
        If liberar And Len(lote) > 0 Then
            ws.Cells(r, C_MAN_INFO).Value = "Atual na LotesStore: Média " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv - 1)) & _
                " · DP " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv)) & ". Preencha só o que vai gravar."
        End If
    Next nv
    ws.Cells(R_MAN_JUST, C_MAN_JUST).MergeArea.ClearContents
    Set rg = RangeManual(ws)
    rg.Locked = Not liberar
    If liberar Then rg.Interior.Color = RGB(255, 242, 204) Else rg.Interior.Color = RGB(217, 217, 217)
    Set st = ws.Cells(R_MAN_TIT, C_MAN_STATUS).MergeArea
    st.Cells(1, 1).Value = statusTxt
    If liberar Then
        st.Interior.Color = RGB(255, 235, 156)
        st.Font.Color = RGB(156, 87, 0)
    Else
        st.Interior.Color = RGB(217, 217, 217)
        st.Font.Color = RGB(64, 64, 64)
    End If
    If Len(marca) > 0 Then ws.Cells(R_MAN_TIT, C_MARCA).Value = marca Else ws.Cells(R_MAN_TIT, C_MARCA).ClearContents
fim:
    nE = Err.Number: sE = Err.Description
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    If nE <> 0 Then Err.Raise nE, MODULO & ".EstadoManual", sE
End Sub

' Tranca de novo (sem gente olhando tambem). motivo vazio = sem linha no Audit_Log (abrir/fechar a pasta). Nunca
' levanta erro.
Public Sub BloquearManual(Optional ByVal motivo As String = "")
    Dim ws As Worksheet, eraLib As Boolean, lk As Variant
    On Error GoTo fim
    If Not SecaoInstalada() Then mLiberado = "": Exit Sub
    Set ws = Pn()
    lk = ws.Cells(R_MAN1, C_MAN_MED).Locked
    eraLib = (Len(mLiberado) > 0) Or (Len(Txt(ws.Cells(R_MAN_TIT, C_MARCA).Value)) > 0)
    If VarType(lk) = vbBoolean Then
        If lk = False Then eraLib = True
    End If
    mLiberado = ""
    If Not eraLib Then Exit Sub
    EstadoManual False, TextoTrancado(), ""
    If Len(motivo) > 0 Then
        mAuditoria.RegistrarLog "ALVO_MANUAL_BLOQUEADO", "ADR-102: digitacao manual de Media/DP trancada de novo (" & motivo & ")."
    End If
fim:
End Sub

' Worksheet_Change do Painel (B3:C3) e SecaoMudou: a liberacao e da chave LOTE|ANALITO do momento; outra chave tranca.
Public Sub ManualContextoMudou()
    Dim k As String
    On Error Resume Next
    If Not SecaoInstalada() Then Exit Sub
    If Len(mLiberado) = 0 And Len(ChaveDaMarca()) = 0 Then Exit Sub
    k = mLotes.ChaveLote(LoteSecao(), AnalitoPainel())
    If mLiberado <> k Or ChaveDaMarca() <> k Then BloquearManual "troca de lote ou analito"
End Sub

' Chave exata LOTE|ANALITO, uma vez so na LotesStore (a mesma trava da apuracao). "" = ok; senao o motivo.
Private Function TravaDoLote(ByVal lote As String, ByVal an As String, ByRef rLs As Long) As String
    Dim vezes As Long
    If Len(lote) = 0 Or UCase$(lote) = "TODOS" Then TravaDoLote = "Escolha o LOTE em B" & R_ENT & " (obrigatório).": Exit Function
    If Len(an) = 0 Then TravaDoLote = "Escolha o analito no Painel (C3).": Exit Function
    rLs = LinhaStore(lote, an, vezes)
    If vezes = 0 Then
        TravaDoLote = "O lote " & lote & " não tem registro de " & an & " na LotesStore (chave " & _
                      mLotes.ChaveLote(lote, an) & "). Confira o lote ou cadastre o lote na aba Analitos."
    ElseIf vezes > 1 Then
        TravaDoLote = "A chave " & mLotes.ChaveLote(lote, an) & " aparece " & vezes & " vezes na LotesStore " & _
                      "(duplicada). Corrija a LotesStore antes."
    End If
End Function

' Botao "Liberar digitacao manual".
Public Sub LiberarDigitacaoManual()
    Dim r As String
    r = LiberarManualCore()
End Sub

' "OK|chave", "CANCELADO|confirmacao" ou "ERRO|motivo".
Public Function LiberarManualCore() As String
    Dim lote As String, an As String, rLs As Long, motivo As String, msg As String, quemTxt As String, k As String
    Dim sE As String
    If Not SecaoInstalada() Then LiberarManualCore = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    lote = LoteSecao()
    an = AnalitoPainel()
    motivo = TravaDoLote(lote, an, rLs)
    If Len(motivo) > 0 Then LiberarManualCore = Falhar(motivo & " A digitação manual continua trancada."): Exit Function
    k = mLotes.ChaveLote(lote, an)
    msg = "EXCEÇÃO TEMPORÁRIA -- digitação manual de Média e DP" & vbCrLf & vbCrLf & _
          "O laboratório trabalha com Média e DP PRÓPRIOS, calculados pelo CIQ (botão ""Aplicar Alvos ao Lote"")." & _
          vbCrLf & vbCrLf & "A digitação manual é exceção: só no início de um lote, enquanto o laboratório ainda não " & _
          "tem os próprios valores -- Média da bula e DP transferido do lote anterior (ou da bula)." & vbCrLf & vbCrLf & _
          "Os valores gravados ficam marcados como EXCEÇÃO TEMPORÁRIA e devem ser substituídos pelos valores do " & _
          "laboratório (""Aplicar Alvos ao Lote"") assim que houver resultados suficientes." & vbCrLf & vbCrLf & _
          "Liberar a digitação para o lote " & lote & " / " & an & "?"
    If Not Perguntar(msg, vbExclamation) Then
        LiberarManualCore = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: a digitação manual continua trancada."
        Exit Function
    End If
    On Error GoTo falha
    quemTxt = Quem()
    EstadoManual True, "LIBERADA (exceção temporária): lote " & lote & " · " & an & " · " & quemTxt & " · " & _
        Format$(Now, "dd/mm/yyyy hh:nn") & ". Gravar, trocar o lote/analito ou fechar a pasta tranca de novo.", _
        "LIBERADO|" & k & "|" & quemTxt & "|" & Format$(Now, "dd/mm/yyyy hh:nn:ss"), lote, an
    mLiberado = k
    mAuditoria.Auditar mAuditoria.CAT_CONFIG, "ALVO_MANUAL_LIBERADO", MODULO, 0, Empty, "", lote, 0, an, Empty, Empty, _
        "", "", "Excecao temporaria", "Digitacao manual de Media/DP liberada no Painel por " & quemTxt & _
        " (excecao temporaria: inicio do lote, sem valores proprios do laboratorio)."
    LiberarManualCore = "OK|" & k
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    BloquearManual
    LiberarManualCore = Falhar("Falha ao liberar a digitação manual: " & sE)
End Function

' Worksheet_Change do Painel no Lote anterior (E107:F1xx): puxa o DP desse lote (mesmo analito e nivel) da LotesStore
' para a coluna DP (editavel) e marca a origem "Lote anterior".
Public Sub ManualEditado(ByVal Target As Range)
    Dim ws As Worksheet, z As Range, nv As Long, r As Long, la As String, an As String, dp As Variant
    Dim prot As Boolean, ev As Boolean, info As String, lote As String
    If mDentro Then Exit Sub
    If Not SecaoInstalada() Then Exit Sub
    Set ws = Pn()
    Set z = Intersect(Target, ws.Range(ws.Cells(R_MAN1, C_MAN_LOTE), ws.Cells(R_MAN1 + NLV() - 1, C_MAN_LOTE + 1)))
    If z Is Nothing Then Exit Sub
    If Len(mLiberado) = 0 Then Exit Sub
    lote = LoteSecao()
    an = AnalitoPainel()
    ' o analito pode ter mudado pelo spinner (B3 sem Change): outra chave nao puxa nada e tranca de novo
    If mLiberado <> mLotes.ChaveLote(lote, an) Then ManualContextoMudou: Exit Sub
    mDentro = True
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscritaRapida(ws)
    For nv = 1 To NLV()
        r = R_MAN1 + nv - 1
        If Not Intersect(z, ws.Range(ws.Cells(r, C_MAN_LOTE), ws.Cells(r, C_MAN_LOTE + 1))) Is Nothing Then
            la = mLotes.LoteCanonico(Txt(ws.Cells(r, C_MAN_LOTE).Value))
            If Len(la) = 0 Then
                info = ""
            Else
                dp = mLotes.ValorDoLote(la, an, 2 * nv)
                If NumOk(dp) Then
                    If CDbl(dp) > 0 Then
                        ws.Cells(r, C_MAN_DP).Value = CDbl(dp)
                        ws.Cells(r, C_MAN_ORIG).Value = ORIG_LOTE_ANT
                        info = "DP do lote anterior " & la & " (LotesStore, " & an & " N" & nv & ") = " & NumTxt(dp) & _
                               " -- pode corrigir."
                    Else
                        info = "O lote " & la & " tem DP " & NumTxt(dp) & " para " & an & " N" & nv & ": digite o DP."
                    End If
                Else
                    info = "O lote " & la & " não tem DP de " & an & " N" & nv & " na LotesStore: digite o DP."
                End If
                If StrComp(la, lote, vbTextCompare) = 0 Then info = info & " ATENÇÃO: é o próprio lote da seção."
            End If
            ws.Cells(r, C_MAN_INFO).MergeArea.ClearContents
            ws.Cells(r, C_MAN_INFO).Value = info
        End If
    Next nv
fim:
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mDentro = False
End Sub

' Botao "Gravar valores manuais".
Public Sub GravarValoresManuais()
    Dim r As String
    r = GravarManualCore()
End Sub

' "OK|campos gravados|valores alterados", "CANCELADO|confirmacao" ou "ERRO|motivo". Nada e gravado fora do OK.
Public Function GravarManualCore() As String
    Dim ws As Worksheet, lote As String, an As String, k As String, rLs As Long, motivo As String, just As String
    Dim nv As Long, r As Long, vM As Variant, vD As Variant, oD As String, la As String, invalidos As String
    Dim temM(1 To 3) As Boolean, temD(1 To 3) As Boolean, valM(1 To 3) As Double, valD(1 To 3) As Double
    Dim fonteD(1 To 3) As String, nCampos As Long, msg As String, partes As String, origem As String, agora As Date
    Dim quemTxt As String, esp As Variant, antes As Variant, nAlt As Long, ls As Worksheet, prot As Boolean
    Dim ev As Boolean, c As Long, bitT As Long, dentroOp As Boolean, sE As String, gravou As Boolean, novo As Double
    Dim dpLa As Variant, notaD(1 To 3) As String, condAnt(1 To 3) As String, tM As String, tD As String
    Dim proprio(1 To 3, 1 To 2) As Boolean, proprios As String, pr As String
    If Not SecaoInstalada() Then GravarManualCore = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    Set ws = Pn()
    lote = LoteSecao()
    an = AnalitoPainel()
    k = mLotes.ChaveLote(lote, an)
    ' 1. so com a digitacao liberada PARA ESTA chave nesta sessao
    If Len(mLiberado) = 0 Then
        BloquearManual
        GravarManualCore = Falhar("A digitação manual está TRANCADA. Clique em ""Liberar digitação manual"" antes. Nada foi gravado.")
        Exit Function
    End If
    If mLiberado <> k Or ChaveDaMarca() <> k Then
        motivo = mLiberado
        BloquearManual "lote ou analito diferente do liberado"
        GravarManualCore = Falhar("A digitação foi liberada para " & motivo & ", não para " & k & _
                                  ". Trancada de novo; libere para o lote/analito certo. Nada foi gravado.")
        Exit Function
    End If
    motivo = TravaDoLote(lote, an, rLs)
    If Len(motivo) > 0 Then GravarManualCore = Falhar(motivo & " Nada foi gravado."): Exit Function
    ' 2. justificativa e valores
    just = Txt(ws.Cells(R_MAN_JUST, C_MAN_JUST).Value)
    If Len(just) < MIN_JUST Then
        GravarManualCore = Falhar("Escreva a justificativa da exceção em B" & R_MAN_JUST & " (ao menos " & MIN_JUST & _
                                  " caracteres: por que o laboratório ainda não tem valores próprios). Nada foi gravado.")
        Exit Function
    End If
    For nv = 1 To NLV()
        r = R_MAN1 + nv - 1
        vM = ws.Cells(r, C_MAN_MED).Value
        vD = ws.Cells(r, C_MAN_DP).Value
        oD = Txt(ws.Cells(r, C_MAN_ORIG).Value)
        la = mLotes.LoteCanonico(Txt(ws.Cells(r, C_MAN_LOTE).Value))
        If IsError(vM) Then
            invalidos = invalidos & vbCrLf & "  N" & nv & ": Média precisa ser um número maior que zero"
        ElseIf Len(Txt(vM)) > 0 Then
            If NumOk(vM) Then
                If CDbl(vM) > 0 Then temM(nv) = True: valM(nv) = CDbl(vM)
            End If
            If Not temM(nv) Then invalidos = invalidos & vbCrLf & "  N" & nv & ": Média precisa ser um número maior que zero"
        End If
        If IsError(vD) Then
            invalidos = invalidos & vbCrLf & "  N" & nv & ": DP precisa ser um número maior que zero"
        ElseIf Len(Txt(vD)) > 0 Then
            If NumOk(vD) Then
                If CDbl(vD) > 0 Then temD(nv) = True: valD(nv) = CDbl(vD)
            End If
            If Not temD(nv) Then invalidos = invalidos & vbCrLf & "  N" & nv & ": DP precisa ser um número maior que zero"
            ' origem do DP: a da lista; vazia = Lote anterior se ele foi preenchido, senao Bula
            If Len(oD) = 0 Then oD = IIf(Len(la) > 0, ORIG_LOTE_ANT, ORIG_BULA)
            If StrComp(oD, ORIG_LOTE_ANT, vbTextCompare) = 0 Then
                If Len(la) = 0 Then
                    invalidos = invalidos & vbCrLf & "  N" & nv & ": origem do DP = Lote anterior, mas o lote anterior está vazio"
                    temD(nv) = False
                ElseIf StrComp(la, lote, vbTextCompare) = 0 Then
                    invalidos = invalidos & vbCrLf & "  N" & nv & ": o lote anterior não pode ser o próprio lote " & lote
                    temD(nv) = False
                Else
                    fonteD(nv) = "LOTE ANTERIOR " & la
                    ' conferencia com a LotesStore no momento de gravar (o DP pode ter sido corrigido a mao)
                    dpLa = mLotes.ValorDoLote(la, an, 2 * nv)
                    If Not NumOk(dpLa) Then
                        notaD(nv) = " (lote " & la & " sem DP de " & an & " N" & nv & " na LotesStore; digitado = " & _
                                    NumTxt(vD) & ")"
                    ElseIf temD(nv) Then
                        If Abs(CDbl(dpLa) - valD(nv)) > 0.000000001 * Abs(CDbl(dpLa)) Then
                            notaD(nv) = " (LotesStore = " & NumTxt(dpLa) & "; digitado = " & NumTxt(valD(nv)) & ")"
                        End If
                    End If
                End If
            ElseIf StrComp(oD, ORIG_BULA, vbTextCompare) = 0 Then
                fonteD(nv) = "BULA"
            Else
                invalidos = invalidos & vbCrLf & "  N" & nv & ": origem do DP precisa ser Lote anterior ou Bula"
                temD(nv) = False
            End If
        End If
        If temM(nv) Then nCampos = nCampos + 1
        If temD(nv) Then nCampos = nCampos + 1
    Next nv
    If Len(invalidos) > 0 Then
        GravarManualCore = Falhar("Valores manuais NÃO gravados (corrija e grave de novo):" & invalidos)
        Exit Function
    End If
    If nCampos = 0 Then
        GravarManualCore = Falhar("Nenhum valor preenchido (Média e/ou DP por nível). Nada foi gravado.")
        Exit Function
    End If
    ' 3. confirmacao
    msg = "Gravar os valores MANUAIS (EXCEÇÃO TEMPORÁRIA) no lote " & lote & " (" & an & ")?" & vbCrLf
    For nv = 1 To NLV()
        If temM(nv) Or temD(nv) Then
            msg = msg & vbCrLf & "N" & nv & ":"
            If temM(nv) Then msg = msg & "  Média " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv - 1)) & " -> " & _
                                  NumTxt(valM(nv)) & " (bula)"
            If temD(nv) Then msg = msg & "  DP " & NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv)) & " -> " & _
                                  NumTxt(valD(nv)) & " (" & LCase$(fonteD(nv)) & ")" & notaD(nv)
        End If
    Next nv
    msg = msg & vbCrLf & vbCrLf & "Justificativa: " & Left$(just, 300) & vbCrLf & vbCrLf & _
          "Os valores ficam marcados como TEMPORÁRIOS (amarelo na apuração) e devem ser substituídos pelos valores do " & _
          "laboratório (""Aplicar Alvos ao Lote"") assim que houver resultados suficientes."
    If Not Perguntar(msg, vbExclamation) Then
        GravarManualCore = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: nada foi gravado (a digitação continua liberada)."
        Exit Function
    End If
    ' 3b. a excecao e so para quem ainda NAO tem valor proprio: campo sem marca temporaria e com condicao propria
    ' (ACEITAVEL / IDEAL no CONDICAO_Nk) pede uma segunda confirmacao explicita
    Set ls = ThisWorkbook.Worksheets(ABA_LS)
    For nv = 1 To NLV()
        condAnt(nv) = Txt(ls.Cells(rLs, LS_C_AUD1 + (nv - 1) * LS_AUD_N + 3).Value)
        MarcasDaCondicao condAnt(nv), tM, tD
        pr = CondPropria(condAnt(nv))
        If Len(pr) > 0 Then
            proprio(nv, 1) = temM(nv) And Len(tM) = 0
            proprio(nv, 2) = temD(nv) And Len(tD) = 0
            If proprio(nv, 1) Then proprios = proprios & vbCrLf & "  N" & nv & " Média " & _
                NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv - 1)) & " (" & pr & ")"
            If proprio(nv, 2) Then proprios = proprios & vbCrLf & "  N" & nv & " DP " & _
                NumTxt(mLotes.ValorDoLote(lote, an, 2 * nv)) & " (" & pr & ")"
        End If
    Next nv
    If Len(proprios) > 0 Then
        msg = "ATENÇÃO: estes valores são PRÓPRIOS do laboratório (calculados pelo CIQ, ""Aplicar Alvos ao Lote""):" & _
              proprios & vbCrLf & vbCrLf & "A exceção temporária é só para o início do lote, SEM valores próprios. " & _
              "Substituir o valor PRÓPRIO do laboratório pelo valor manual mesmo assim?"
        If Not Perguntar(msg, vbExclamation) Then
            GravarManualCore = "CANCELADO|proprio"
            mUltimoAviso = "Cancelado: nada foi gravado (os valores próprios do laboratório ficam; a digitação continua liberada)."
            Exit Function
        End If
    End If
    ' 4. gravacao
    On Error GoTo falha
    agora = Now
    quemTxt = Quem()
    For nv = 1 To NLV()
        If temM(nv) Or temD(nv) Then
            partes = partes & IIf(Len(partes) > 0, " | ", "") & "N" & nv
            If temM(nv) Then partes = partes & " media=BULA"
            If temM(nv) And temD(nv) Then partes = partes & ";"
            If temD(nv) Then partes = partes & " DP=" & fonteD(nv)
        End If
    Next nv
    origem = "EXCECAO TEMPORARIA " & Format$(agora, "dd/mm/yyyy hh:nn") & " " & quemTxt & ": " & partes
    mApp.InicioUsuario "Gravando os valores manuais do lote " & lote & "..."
    dentroOp = True
    esp = mLotesBlocos.EspecDoLote(lote, an)
    For nv = 1 To NLV()
        For c = 1 To 2
            If (c = 1 And temM(nv)) Or (c = 2 And temD(nv)) Then
                If c = 1 Then novo = valM(nv) Else novo = valD(nv)
                antes = mLotes.GravarParametroStore(lote, an, 2 * (nv - 1) + c, novo, esp, origem)
                gravou = True
                If CStr(antes) <> CStr(novo) Then nAlt = nAlt + 1
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "PARAMETRO_LOTE_EXCECAO", MODULO, 0, Empty, "", lote, nv, an, _
                    antes, novo, "", "", IIf(c = 1, "Media N", "DP N") & nv, _
                    IIf(c = 1, "Media N" & nv & " = BULA", "DP N" & nv & " = " & fonteD(nv) & notaD(nv)) & _
                    " (EXCECAO TEMPORARIA, " & quemTxt & ")" & _
                    IIf(proprio(nv, c), "; substituiu valor PROPRIO do laboratorio (" & CondPropria(condAnt(nv)) & _
                        ", confirmado)", "") & ". Justificativa: " & Left$(just, 800)
            End If
        Next c
    Next nv
    ' origem W (mesmo com valor igual), bit TESTE do campo gravado e AA:AO do nivel (CONDICAO = EXCECAO_TEMPORARIA)
    Set ls = ThisWorkbook.Worksheets(ABA_LS)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(ls)
    GarantirCabecalhoAuditoria ls
    ls.Cells(rLs, LS_C_ORIGEM).Value = origem
    For nv = 1 To NLV()
        If temM(nv) Or temD(nv) Then
            bitT = Val(Txt(ls.Cells(rLs, LS_C_TESTE1 + nv - 1).Value))
            If temM(nv) Then bitT = bitT And Not 1
            If temD(nv) Then bitT = bitT And Not 2
            If bitT > 0 Then ls.Cells(rLs, LS_C_TESTE1 + nv - 1).Value = bitT Else ls.Cells(rLs, LS_C_TESTE1 + nv - 1).ClearContents
            c = LS_C_AUD1 + (nv - 1) * LS_AUD_N
            ' marca POR CAMPO: so o campo digitado vira EXCECAO; o outro fica com a marca/condicao que tinha
            MarcasDaCondicao condAnt(nv), tM, tD
            If temM(nv) Then tM = "EXCECAO"
            If temD(nv) Then tD = "EXCECAO"
            ' N e periodo so saem quando os dois campos sao digitados (o campo mantido continua sendo daquele N)
            If tM = "EXCECAO" And tD = "EXCECAO" Then ls.Range(ls.Cells(rLs, c), ls.Cells(rLs, c + 2)).ClearContents
            ls.Cells(rLs, c + 3).Value = CondicaoDasMarcas(tM, tD, CondPropria(condAnt(nv)))
            ls.Cells(rLs, c + 4).Value = agora
        End If
    Next nv
    mSeguranca.RestaurarProtecao ls, prot
    prot = False
    Application.EnableEvents = ev
    mLotes.InvalidarLotes
    If StrComp(lote, mLotes.LoteEmTela(), vbTextCompare) = 0 Then mLotes.CarregarParametrosView lote
    mEstatistica.InvalidarParametros
    mEstatistica.RecalcularAnalitoAtual
    mLotes.AtualizarResumoLote
    mLotesBlocos.EscreverValoresDoLote lote
    AtualizarVistaAlvos
    mApp.fim
    dentroOp = False
    BloquearManual "valores manuais gravados"
    Avisar "Valores manuais gravados no lote " & lote & " (" & an & "): " & nCampos & " campo(s), " & nAlt & _
           " valor(es) alterado(s). EXCEÇÃO TEMPORÁRIA: substitua pelos valores do laboratório (""Aplicar Alvos ao " & _
           "Lote"") assim que houver resultados suficientes. A digitação foi trancada de novo.", vbExclamation
    GravarManualCore = "OK|" & nCampos & "|" & nAlt
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ls Is Nothing Then mSeguranca.RestaurarProtecao ls, prot
    If dentroOp Then mApp.fim Else Application.EnableEvents = True
    mLotes.InvalidarLotes
    BloquearManual "falha ao gravar"
    GravarManualCore = Falhar("Falha ao gravar os valores manuais do lote " & lote & ": " & sE & _
                              IIf(gravou, " (parte pode ter sido gravada: confira a Origem atual na apuração)", ""))
End Function

' ============================ REFERENCIA (RefBulaStore) ============================
Private Function ColRef(ByVal nv As Long, ByVal campo As Long) As Long       ' campo 0 Media, 1 LI, 2 LS (Painel)
    ColRef = 2 + (nv - 1) * 3 + campo
End Function

Private Function ColFonte() As Long
    ColFonte = 2 + NLV() * 3
End Function

' RefBulaStore em memoria: chave -> Array(9 valores N1..N3 Media/LI/LS, fonte, linha)
Private Function LerRef() As Object
    Dim d As Object, ws As Worksheet, ult As Long, v As Variant, i As Long, k As String, j As Long, a() As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set ws = AbaSeExiste(ABA_REF)
    If Not ws Is Nothing Then
        ult = UltimaLinha(ws, RB_C_CHAVE)
        If ult >= 2 Then
            v = ws.Range(ws.Cells(2, 1), ws.Cells(ult + 1, RB_NCOL)).Value
            For i = 1 To UBound(v, 1) - 1
                k = UCase$(Txt(v(i, RB_C_CHAVE)))
                If Len(k) > 0 And Not d.Exists(k) Then
                    ReDim a(0 To 10)
                    For j = 0 To 8
                        a(j) = v(i, RB_C_V1 + j)
                    Next j
                    a(9) = v(i, RB_C_FONTE)
                    a(10) = i + 1
                    d.Add k, a
                End If
            Next i
        End If
    End If
    Set LerRef = d
End Function

' Vista da RefBulaStore para o lote da secao (linhas 115-154) + coluna "Fora da faixa" (pendentes / cientes).
Public Sub AtualizarVistaReferencia()
    Dim ws As Worksheet, lote As String, nomes As Variant, ref As Object, out() As Variant, i As Long, nv As Long
    Dim c As Long, k As String, a As Variant, prot As Boolean, ev As Boolean, pend As Object, cien As Object
    Dim fora() As Variant, nm As String, wa As Worksheet
    If Not SecaoInstalada() Then Exit Sub
    Set ws = Pn()
    lote = LoteSecao()
    Set wa = ThisWorkbook.Worksheets(mLotes.AbaEspec())
    nomes = wa.Range(wa.Cells(4, 1), wa.Cells(4 + N_REF - 1, 1)).Value
    Set ref = LerRef()
    ContarObsPorAnalito lote, pend, cien
    ReDim out(1 To N_REF, 1 To ColFonte())
    ReDim fora(1 To N_REF, 1 To 1)
    For i = 1 To N_REF
        nm = Txt(nomes(i, 1))
        out(i, 1) = nm
        If Len(nm) > 0 And Len(lote) > 0 Then
            k = mLotes.ChaveLote(lote, nm)
            If ref.Exists(k) Then
                a = ref(k)
                For nv = 1 To NLV()
                    For c = 0 To 2
                        If NumOk(a((nv - 1) * 3 + c)) Then out(i, ColRef(nv, c)) = a((nv - 1) * 3 + c)
                    Next c
                Next nv
                out(i, ColFonte()) = Txt(a(9))
            End If
            If pend.Exists(UCase$(nm)) Then fora(i, 1) = pend(UCase$(nm)) & " pendente(s)"
            If cien.Exists(UCase$(nm)) Then fora(i, 1) = IIf(Len(CStr(fora(i, 1))) > 0, fora(i, 1) & " · ", "") & _
                                                         cien(UCase$(nm)) & " ciente(s)"
        End If
    Next i
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscritaRapida(ws)
    ws.Range(ws.Cells(R_REF1, 1), ws.Cells(R_REF1 + N_REF - 1, ColFonte())).Value = out
    ws.Range(ws.Cells(R_REF1, ColFonte() + 1), ws.Cells(R_REF1 + N_REF - 1, ColFonte() + 1)).Value = fora
fim:
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub

' Worksheet_Change do Painel na grade (B115 ate a coluna Fonte, linha 154): grava na RefBulaStore e verifica.
Public Sub GradeEditada(ByVal Target As Range)
    Dim ws As Worksheet, z As Range, c As Range, linhas As Object, r As Variant, lote As String, an As String
    Dim st As Worksheet, ref As Object, k As String, a As Variant, rs As Long, nv As Long, j As Long
    Dim novo(0 To 8) As Variant, ok(1 To 3) As Boolean, invalidos As String, avisos As String, v As Variant
    Dim prot As Boolean, ev As Boolean, mudou As Boolean, campo As Variant, fonte As String, antes As Variant
    Dim nE As Long, sE As String, res As String, novas As String, nAud As Long
    If mDentro Then Exit Sub
    If Not SecaoInstalada() Then Exit Sub
    Set ws = Pn()
    Set z = Intersect(Target, ws.Range(ws.Cells(R_REF1, 2), ws.Cells(R_REF1 + N_REF - 1, ColFonte())))
    If z Is Nothing Then Exit Sub
    lote = LoteSecao()
    If Len(lote) = 0 Then
        AtualizarVistaReferencia
        Avisar "Escolha o lote em B98 antes de digitar os valores de referência. Nada foi gravado.", vbExclamation
        Exit Sub
    End If
    mDentro = True
    On Error GoTo falha
    Set linhas = CreateObject("Scripting.Dictionary")
    For Each c In z.Cells
        If Not linhas.Exists(c.Row) Then linhas.Add c.Row, 1
    Next c
    Set st = Aba(ABA_REF)
    Set ref = LerRef()
    campo = Array("Media", "LI", "LS")
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(st)
    For Each r In linhas.Keys
        an = Txt(ws.Cells(r, 1).Value)
        If Len(an) = 0 Then GoTo proxLinha
        k = mLotes.ChaveLote(lote, an)
        If ref.Exists(k) Then a = ref(k) Else a = Empty
        For j = 0 To 8
            If IsArray(a) Then novo(j) = a(j) Else novo(j) = Empty
            If NumOk(novo(j)) = False Then novo(j) = Empty
        Next j
        For nv = 1 To NLV()
            ok(nv) = True
            For j = 0 To 2
                v = ws.Cells(r, ColRef(nv, j)).Value
                If IsError(v) Then
                    ok(nv) = False
                ElseIf Len(Trim$(CStr(v))) > 0 And Not NumOk(v) Then
                    ok(nv) = False
                End If
            Next j
            If ok(nv) Then
                If NumOk(ws.Cells(r, ColRef(nv, 1)).Value) And NumOk(ws.Cells(r, ColRef(nv, 2)).Value) Then
                    If CDbl(ws.Cells(r, ColRef(nv, 1)).Value) > CDbl(ws.Cells(r, ColRef(nv, 2)).Value) Then ok(nv) = False
                End If
            End If
            If ok(nv) Then
                For j = 0 To 2
                    v = ws.Cells(r, ColRef(nv, j)).Value
                    If NumOk(v) Then novo((nv - 1) * 3 + j) = CDbl(v) Else novo((nv - 1) * 3 + j) = Empty
                Next j
                If NumOk(novo((nv - 1) * 3)) Then
                    If NumOk(novo((nv - 1) * 3 + 1)) Then
                        If novo((nv - 1) * 3) < novo((nv - 1) * 3 + 1) Then avisos = avisos & vbCrLf & "  " & an & " N" & nv & ": Média abaixo do LI"
                    End If
                    If NumOk(novo((nv - 1) * 3 + 2)) Then
                        If novo((nv - 1) * 3) > novo((nv - 1) * 3 + 2) Then avisos = avisos & vbCrLf & "  " & an & " N" & nv & ": Média acima do LS"
                    End If
                End If
            Else
                invalidos = invalidos & vbCrLf & "  " & an & " N" & nv & ": valores não numéricos ou LI maior que LS"
            End If
        Next nv
        fonte = Txt(ws.Cells(r, ColFonte()).Value)
        ' compara com a store
        mudou = Not IsArray(a)
        If IsArray(a) Then
            For j = 0 To 8
                If CStr(a(j)) <> CStr(novo(j)) Then mudou = True
            Next j
            If Txt(a(9)) <> fonte Then mudou = True
        Else
            mudou = False
            For j = 0 To 8
                If Not IsEmpty(novo(j)) Then mudou = True
            Next j
            If Len(fonte) > 0 Then mudou = True
        End If
        If Not mudou Then GoTo proxLinha
        If IsArray(a) Then rs = a(10) Else rs = UltimaLinha(st, RB_C_CHAVE) + 1
        If rs < 2 Then rs = 2
        st.Cells(rs, RB_C_LOTE).NumberFormat = "@"
        st.Cells(rs, RB_C_LOTE).Value = lote
        st.Cells(rs, RB_C_AN).Value = an
        st.Cells(rs, RB_C_CHAVE).Value = k
        For j = 0 To 8
            st.Cells(rs, RB_C_V1 + j).Value = novo(j)
            If IsArray(a) Then antes = a(j) Else antes = Empty
            If CStr(antes) <> CStr(novo(j)) Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "REFERENCIA_BULA_ALTERADA", MODULO, 0, Empty, "", lote, _
                    j \ 3 + 1, an, antes, novo(j), "", "", "Ref " & campo(j Mod 3) & " N" & (j \ 3 + 1), _
                    "Referencia da bula/Unity (" & IIf(Len(fonte) > 0, fonte, "sem fonte") & ") do lote " & lote & _
                    " alterada no Painel por " & Quem() & "."
                nAud = nAud + 1
            End If
        Next j
        st.Cells(rs, RB_C_FONTE).Value = fonte
        st.Cells(rs, RB_C_QUANDO).Value = Now
        st.Cells(rs, RB_C_QUEM).Value = Quem()
        If Not IsArray(a) Then
            ReDim a(0 To 10)
            a(10) = rs
        End If
        For j = 0 To 8
            a(j) = novo(j)
        Next j
        a(9) = fonte
        If ref.Exists(k) Then ref(k) = a Else ref.Add k, a
proxLinha:
    Next r
    mSeguranca.RestaurarProtecao st, prot
    prot = False
    Application.EnableEvents = ev
    AtualizarVistaReferencia                   ' devolve o que a store tem (desfaz o invalido)
    res = VerificarNucleo(Empty, novas)
    AtualizarVistaReferencia
    AtualizarVistaObs
    mDentro = False
    If Len(invalidos) > 0 Then
        Avisar "Valores de referência NÃO gravados (corrija e digite de novo):" & invalidos, vbExclamation
    End If
    If Len(avisos) > 0 Then
        Avisar "Gravado, mas confira (informativo):" & avisos, vbInformation
    End If
    AlertarNovas res, novas
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then mSeguranca.RestaurarProtecao st, prot
    Application.EnableEvents = True
    mDentro = False
    AtualizarVistaReferencia
    Avisar "Falha ao gravar a referência: " & sE, vbCritical
End Sub

' ============================ VERIFICACAO CONTRA A REFERENCIA ============================
' Compara todo resultado ATIVO com a referencia do seu lote x analito x nivel. Atualiza a RefBulaObs (novas
' PENDENTE; PENDENTE que deixou de estar fora = RETIRADA; RETIRADA que voltou = PENDENTE; CIENTE nunca muda).
' Devolve "OK|novas|pendentes|retiradas"; novasTxt = as novas, uma por linha (para o alerta).
Private Function VerificarNucleo(ByVal dados As Variant, ByRef novasTxt As String) As String
    Dim ref As Object, atual As Object, i As Long, k As String, a As Variant, nv As Long, x As Double
    Dim li As Variant, lsv As Variant, sit As String, ko As String, id As String, ws As Worksheet, ult As Long
    Dim v As Variant, ex As Object, st As String, nNovas As Long, nRet As Long, nPend As Long, prot As Boolean
    Dim novos As Collection, item As Variant, r As Long, agora As Date, ev As Boolean, j As Long, mudouV As Boolean
    novasTxt = ""
    Set ref = LerRef()
    Set atual = CreateObject("Scripting.Dictionary")
    atual.CompareMode = 1
    If ref.Count > 0 Then
        If IsEmpty(dados) Then dados = mDados.CarregarDB()
        If Not IsEmpty(dados) Then
            For i = 1 To UBound(dados, 1)
                If CStr(dados(i, mDados.COL_STATUS)) <> mDados.ST_ATIVO Then GoTo prox
                If Not NumOk(dados(i, mDados.COL_RESULT)) Then GoTo prox
                k = mLotes.ChaveLote(CStr(dados(i, mDados.COL_LOTE)), CStr(dados(i, mDados.COL_ANALITO)))
                If Not ref.Exists(k) Then GoTo prox
                If IsError(dados(i, mDados.COL_NIVEL)) Then GoTo prox
                nv = Val(CStr(dados(i, mDados.COL_NIVEL)))
                If nv < 1 Or nv > NLV() Then GoTo prox
                a = ref(k)
                li = a((nv - 1) * 3 + 1): lsv = a((nv - 1) * 3 + 2)
                x = CDbl(dados(i, mDados.COL_RESULT))
                sit = ""
                If NumOk(li) Then
                    If x < CDbl(li) Then sit = "ABAIXO DO LI"
                End If
                If Len(sit) = 0 And NumOk(lsv) Then
                    If x > CDbl(lsv) Then sit = "ACIMA DO LS"
                End If
                If Len(sit) = 0 Then GoTo prox
                id = Txt(dados(i, mDados.COL_ID))
                If Len(id) = 0 Then id = "SEM_ID:" & Txt(dados(i, mDados.COL_RUN)) & "/" & k & "/" & nv & "/" & Trim$(Str$(x))
                ' chave ID|ANALITO|Nk|LI|LS: o ID_REGISTRO e da amostra (o mesmo para todos os analitos -- a inativacao e
                ' por ID + analito, ADR-070); com o analito e o nivel a chave e de UM resultado
                ko = UCase$(id) & "|" & UCase$(Txt(dados(i, mDados.COL_ANALITO))) & "|N" & nv & "|" & NumChave(li) & "|" & NumChave(lsv)
                If Not atual.Exists(ko) Then
                    atual.Add ko, Array(id, Txt(dados(i, mDados.COL_LOTE)), Txt(dados(i, mDados.COL_ANALITO)), nv, _
                                        dados(i, mDados.COL_DATA), x, li, lsv, sit, Txt(dados(i, mDados.COL_EQUIP)), _
                                        Txt(dados(i, mDados.COL_RUN)))
                End If
prox:
            Next i
        End If
    End If
    ' sem referencia nenhuma e sem aba de observacoes: nada a fazer (nao cria aba)
    If atual.Count = 0 And AbaSeExiste(ABA_OBS) Is Nothing Then VerificarNucleo = "OK|0|0|0": Exit Function
    Set ws = Aba(ABA_OBS)
    Set ex = CreateObject("Scripting.Dictionary")
    ex.CompareMode = 1
    ult = UltimaLinha(ws, OB_C_CHAVE)
    agora = Now
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(ws)
    On Error GoTo falha
    If ult >= 2 Then
        v = ws.Range(ws.Cells(2, 1), ws.Cells(ult + 1, OB_NCOL)).Value
        mudouV = False
        For i = 1 To UBound(v, 1) - 1
            k = UCase$(Txt(v(i, OB_C_CHAVE)))
            If Len(k) > 0 Then
                If Not ex.Exists(k) Then ex.Add k, i
                st = UCase$(Txt(v(i, OB_C_ST)))
                If st = ST_PEND And Not atual.Exists(k) Then
                    v(i, OB_C_ST) = ST_RETIRADA: v(i, OB_C_STEM) = agora: v(i, OB_C_STPOR) = "verificacao automatica"
                    nRet = nRet + 1: mudouV = True
                ElseIf st = ST_RETIRADA And atual.Exists(k) Then
                    v(i, OB_C_ST) = ST_PEND: v(i, OB_C_STEM) = agora: v(i, OB_C_STPOR) = "verificacao automatica"
                    nNovas = nNovas + 1: mudouV = True
                    novasTxt = novasTxt & LinhaAlerta(atual(k)) & vbCrLf
                End If
                If UCase$(Txt(v(i, OB_C_ST))) = ST_PEND Then nPend = nPend + 1
            End If
        Next i
        If mudouV Then ws.Range(ws.Cells(2, 1), ws.Cells(ult, OB_NCOL)).Value = SemUltimaLinha(v)
    End If
    Set novos = New Collection
    For Each item In atual.Keys
        If Not ex.Exists(item) Then novos.Add item
    Next item
    If novos.Count > 0 Then
        Dim out() As Variant
        ReDim out(1 To novos.Count, 1 To OB_NCOL)
        r = 0
        For Each item In novos
            r = r + 1
            a = atual(item)
            out(r, OB_C_CHAVE) = item
            out(r, OB_C_ID) = a(0)
            out(r, OB_C_LOTE) = a(1)
            out(r, OB_C_AN) = a(2)
            out(r, OB_C_NV) = a(3)
            out(r, OB_C_DATA) = a(4)
            out(r, OB_C_RES) = a(5)
            out(r, OB_C_LI) = a(6)
            out(r, OB_C_LS) = a(7)
            out(r, OB_C_SIT) = a(8)
            out(r, OB_C_ST) = ST_PEND
            out(r, OB_C_CRIADA) = agora
            out(r, OB_C_EQ) = a(9)
            out(r, OB_C_RUN) = a(10)
            If r <= 10 Then novasTxt = novasTxt & LinhaAlerta(a) & vbCrLf
        Next item
        If ult < 1 Then ult = 1
        ws.Range(ws.Cells(ult + 1, OB_C_ID), ws.Cells(ult + novos.Count, OB_C_LOTE)).NumberFormat = "@"
        ws.Range(ws.Cells(ult + 1, 1), ws.Cells(ult + novos.Count, OB_NCOL)).Value = out
        nNovas = nNovas + novos.Count
        nPend = nPend + novos.Count
    End If
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    If nNovas > 0 Then
        mAuditoria.RegistrarLog "REFERENCIA_OBSERVACAO_NOVA", nNovas & " resultado(s) novo(s) fora da faixa de " & _
            "referencia da bula/Unity (informativo; " & nPend & " pendente(s))."
    End If
    VerificarNucleo = "OK|" & nNovas & "|" & nPend & "|" & nRet
    Exit Function
falha:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    Err.Raise nE, MODULO & ".VerificarNucleo", sE
End Function

Private Function SemUltimaLinha(ByVal v As Variant) As Variant
    Dim o() As Variant, i As Long, j As Long
    ReDim o(1 To UBound(v, 1) - 1, 1 To UBound(v, 2))
    For i = 1 To UBound(v, 1) - 1
        For j = 1 To UBound(v, 2)
            o(i, j) = v(i, j)
        Next j
    Next i
    SemUltimaLinha = o
End Function

Private Function LinhaAlerta(ByVal a As Variant) As String
    Dim d As String
    If IsDate(a(4)) Then d = Format$(CDate(a(4)), "dd/mm/yyyy hh:nn") Else d = Txt(a(4))
    LinhaAlerta = "  lote " & a(1) & " · " & a(2) & " N" & a(3) & " · " & d & " · " & NumTxt(a(5)) & " " & _
                  LCase$(a(8)) & " (LI " & NumTxt(a(6)) & ", LS " & NumTxt(a(7)) & ")"
End Function

Private Sub AlertarNovas(ByVal res As String, ByVal novasTxt As String)
    Dim p() As String
    p = Split(res & "|0|0|0", "|")
    If p(0) <> "OK" Then Exit Sub
    If Val(p(1)) = 0 Then Exit Sub
    Avisar Val(p(1)) & " resultado(s) de controle NOVO(S) fora da faixa de referência da bula/Unity:" & vbCrLf & _
           novasTxt & IIf(Val(p(1)) > 10, "  (...)" & vbCrLf, "") & vbCrLf & _
           "Registrado em OBSERVAÇÕES AUTOMÁTICAS (Painel, linha " & R_OBS_TIT & "). Informativo: não muda status, " & _
           "Westgard, LJ nem liberação. A ciência, a análise e a aprovação da corrida são do analista/supervisor " & _
           "técnico (parecer na linha " & R_PAR_TXT & ").", vbExclamation
End Sub

' Para o QA e para a edicao: verifica e alerta (sem gente olhando, so guarda o texto). "OK|novas|pendentes|retiradas".
Public Function VerificarReferencia() As String
    Dim novas As String, r As String
    On Error GoTo falha
    r = VerificarNucleo(Empty, novas)
    If SecaoInstalada() Then
        AtualizarVistaReferencia
        AtualizarVistaObs
    End If
    AlertarNovas r, novas
    VerificarReferencia = r
    Exit Function
falha:
    VerificarReferencia = "ERRO|" & Err.Description
End Function

' Fim do ATUALIZAR (mIntegracao.AtualizarDadosCore): sem caixa de mensagem. Devolve o texto para o resumo do
' ATUALIZAR ("" sem observacao nova). Nunca levanta erro: a falha vai para o Audit_Log.
Public Function VerificarAoAtualizar() As String
    Dim novas As String, r As String, p() As String, dados As Variant
    On Error GoTo falha
    If AbaSeExiste(ABA_REF) Is Nothing And Not SecaoInstalada() Then Exit Function
    dados = mDados.CarregarDB()
    r = VerificarNucleo(dados, novas)
    If SecaoInstalada() Then
        AtualizarVistaAlvos dados
        AtualizarVistaReferencia
        AtualizarVistaObs
    End If
    p = Split(r & "|0|0|0", "|")
    If Val(p(1)) > 0 Then
        mUltimoAviso = novas
        VerificarAoAtualizar = "Referência bula/Unity: " & Val(p(1)) & " resultado(s) novo(s) fora da faixa (" & _
            Val(p(2)) & " pendente(s) de ciência). Informativo -- veja o Painel, OBSERVAÇÕES AUTOMÁTICAS."
    End If
    Exit Function
falha:
    Dim sE As String
    sE = Err.Description
    On Error Resume Next
    mAuditoria.RegistrarLog "REFERENCIA_VERIFICACAO_FALHOU", "ADR-102: " & sE
    VerificarAoAtualizar = ""
End Function

' ============================ OBSERVACOES (linhas 156-173) ============================
Private Sub ContarObsPorAnalito(ByVal lote As String, ByRef pend As Object, ByRef cien As Object)
    Dim ws As Worksheet, ult As Long, v As Variant, i As Long, an As String, st As String
    Set pend = CreateObject("Scripting.Dictionary"): pend.CompareMode = 1
    Set cien = CreateObject("Scripting.Dictionary"): cien.CompareMode = 1
    If Len(lote) = 0 Then Exit Sub
    Set ws = AbaSeExiste(ABA_OBS)
    If ws Is Nothing Then Exit Sub
    ult = UltimaLinha(ws, OB_C_CHAVE)
    If ult < 2 Then Exit Sub
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult + 1, OB_NCOL)).Value
    For i = 1 To UBound(v, 1) - 1
        If StrComp(Txt(v(i, OB_C_LOTE)), lote, vbTextCompare) = 0 Then
            an = UCase$(Txt(v(i, OB_C_AN)))
            st = UCase$(Txt(v(i, OB_C_ST)))
            If st = ST_PEND Then pend(an) = pend(an) + 1
            If st = ST_CIENTE Then cien(an) = cien(an) + 1
        End If
    Next i
End Sub

Public Sub AtualizarVistaObs()
    Dim ws As Worksheet, o As Worksheet, lote As String, ult As Long, v As Variant, i As Long, j As Long
    Dim idx() As Long, nP As Long, nC As Long, nR As Long, out() As Variant, usado() As Boolean, melhor As Long
    Dim dm As Double, dd As Double, q As Long, prot As Boolean, ev As Boolean, st As String, txtSt As String
    Dim pw As Worksheet, uP As Long, pv As Variant, ultPar As String
    If Not SecaoInstalada() Then Exit Sub
    Set ws = Pn()
    lote = LoteSecao()
    ReDim out(1 To N_OBS_TELA, 1 To 10)
    Set o = AbaSeExiste(ABA_OBS)
    If Not o Is Nothing And Len(lote) > 0 Then
        ult = UltimaLinha(o, OB_C_CHAVE)
        If ult >= 2 Then
            v = o.Range(o.Cells(2, 1), o.Cells(ult + 1, OB_NCOL)).Value
            ReDim idx(1 To UBound(v, 1))
            For i = 1 To UBound(v, 1) - 1
                If StrComp(Txt(v(i, OB_C_LOTE)), lote, vbTextCompare) = 0 Then
                    st = UCase$(Txt(v(i, OB_C_ST)))
                    If st = ST_PEND Then
                        nP = nP + 1: idx(nP) = i
                    ElseIf st = ST_CIENTE Then
                        nC = nC + 1
                    ElseIf st = ST_RETIRADA Then
                        nR = nR + 1
                    End If
                End If
            Next i
            ' as 15 mais recentes (data do resultado)
            If nP > 0 Then
                ReDim usado(1 To nP)
                For q = 1 To IIf(nP < N_OBS_TELA, nP, N_OBS_TELA)
                    melhor = 0: dm = -1
                    For j = 1 To nP
                        If Not usado(j) Then
                            dd = 0
                            If IsDate(v(idx(j), OB_C_DATA)) Then dd = CDbl(CDate(v(idx(j), OB_C_DATA)))
                            If dd > dm Then dm = dd: melhor = j
                        End If
                    Next j
                    usado(melhor) = True
                    i = idx(melhor)
                    If IsDate(v(i, OB_C_DATA)) Then
                        out(q, 1) = Format$(CDate(v(i, OB_C_DATA)), "dd/mm/yyyy hh:nn")
                    Else
                        out(q, 1) = Txt(v(i, OB_C_DATA))
                    End If
                    out(q, 2) = Txt(v(i, OB_C_AN))
                    out(q, 3) = "N" & Txt(v(i, OB_C_NV))
                    out(q, 4) = v(i, OB_C_RES)
                    If NumOk(v(i, OB_C_LI)) Then out(q, 5) = v(i, OB_C_LI)
                    If NumOk(v(i, OB_C_LS)) Then out(q, 6) = v(i, OB_C_LS)
                    out(q, 7) = Txt(v(i, OB_C_SIT))
                    out(q, 8) = Txt(v(i, OB_C_EQ))
                    out(q, 9) = Txt(v(i, OB_C_ID))
                    If IsDate(v(i, OB_C_CRIADA)) Then out(q, 10) = Format$(CDate(v(i, OB_C_CRIADA)), "dd/mm/yyyy hh:nn")
                Next q
            End If
        End If
    End If
    If Len(lote) = 0 Then
        txtSt = "Escolha o lote em B98."
    ElseIf nP = 0 Then
        txtSt = "Lote " & lote & ": nenhuma observação pendente."
    Else
        txtSt = "Lote " & lote & ": " & nP & " observação(ões) PENDENTE(S) de ciência -- registre o parecer abaixo."
    End If
    ' ultimo parecer do lote
    Set pw = AbaSeExiste(ABA_PAR)
    If Not pw Is Nothing And Len(lote) > 0 Then
        uP = UltimaLinha(pw, PA_C_ID)
        If uP >= 2 Then
            pv = pw.Range(pw.Cells(2, 1), pw.Cells(uP + 1, PA_NCOL)).Value
            For i = UBound(pv, 1) - 1 To 1 Step -1
                If StrComp(Txt(pv(i, PA_C_LOTE)), lote, vbTextCompare) = 0 Then
                    ultPar = ""
                    If IsDate(pv(i, PA_C_DATA)) Then ultPar = Format$(CDate(pv(i, PA_C_DATA)), "dd/mm/yyyy hh:nn")
                    ultPar = "Último parecer do lote " & lote & ": " & ultPar & _
                             " · " & Txt(pv(i, PA_C_QUEM)) & " · " & Txt(pv(i, PA_C_NOBS)) & " observação(ões) · " & _
                             Txt(pv(i, PA_C_TEXTO))
                    Exit For
                End If
            Next i
        End If
    End If
    If Len(ultPar) = 0 Then ultPar = "Último parecer: nenhum" & IIf(Len(lote) > 0, " para o lote " & lote, "") & "."
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscritaRapida(ws)
    ws.Range(ws.Cells(R_OBS1, 1), ws.Cells(R_OBS1 + N_OBS_TELA - 1, 10)).Value = out
    ws.Cells(R_OBS_TOT, 1).Value = "Pendentes: " & nP & IIf(nP > N_OBS_TELA, " (mostrando as " & N_OBS_TELA & " mais recentes)", "") & _
        " · cientes: " & nC & " · retiradas (limite corrigido ou resultado inativado): " & nR & _
        ".  Informativo: não muda status, Westgard, LJ nem liberação; a ciência, a análise e a aprovação da corrida " & _
        "são do analista/supervisor técnico."
    ws.Cells(R_OBS_TIT, C_OBS_STATUS).Value = txtSt
    If nP > 0 Then
        ws.Cells(R_OBS_TIT, C_OBS_STATUS).Interior.Color = RGB(255, 199, 206)
        ws.Cells(R_OBS_TIT, C_OBS_STATUS).Font.Color = RGB(156, 0, 6)
    Else
        ws.Cells(R_OBS_TIT, C_OBS_STATUS).Interior.Color = RGB(198, 239, 206)
        ws.Cells(R_OBS_TIT, C_OBS_STATUS).Font.Color = RGB(0, 97, 0)
    End If
    ws.Cells(R_PAR_ULT, 1).Value = ultPar
fim:
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub

' ============================ PARECER TECNICO ============================
' Botao "Registrar parecer".
Public Sub RegistrarParecer()
    Dim r As String
    r = RegistrarParecerCore()
End Sub

' "OK|id|observacoes", "CANCELADO|..." ou "ERRO|...".
Public Function RegistrarParecerCore() As String
    Dim ws As Worksheet, lote As String, texto As String, o As Worksheet, ult As Long, v As Variant, i As Long
    Dim nP As Long, pw As Worksheet, uP As Long, id As String, agora As Date, quemTxt As String
    Dim prot As Boolean, protO As Boolean, protP As Boolean, ev As Boolean, nE As Long, sE As String
    Dim gravado As Boolean
    If Not SecaoInstalada() Then RegistrarParecerCore = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    Set ws = Pn()
    lote = LoteSecao()
    If Len(lote) = 0 Then RegistrarParecerCore = Falhar("Escolha o lote em B98. Nada foi registrado."): Exit Function
    texto = Txt(ws.Cells(R_PAR_TXT, 1).Value)
    If Len(texto) < MIN_PARECER Then
        RegistrarParecerCore = Falhar("Escreva o parecer técnico em A" & R_PAR_TXT & " (razão da ciência e conduta; " & _
                                      "ao menos " & MIN_PARECER & " caracteres). Nada foi registrado.")
        Exit Function
    End If
    Set o = AbaSeExiste(ABA_OBS)
    If Not o Is Nothing Then
        ult = UltimaLinha(o, OB_C_CHAVE)
        If ult >= 2 Then
            v = o.Range(o.Cells(2, 1), o.Cells(ult + 1, OB_NCOL)).Value
            For i = 1 To UBound(v, 1) - 1
                If StrComp(Txt(v(i, OB_C_LOTE)), lote, vbTextCompare) = 0 And UCase$(Txt(v(i, OB_C_ST))) = ST_PEND Then nP = nP + 1
            Next i
        End If
    End If
    If Not Perguntar("Registrar o parecer técnico do lote " & lote & " e dar ciência a " & nP & _
                     " observação(ões) pendente(s)?" & vbCrLf & vbCrLf & "Parecer: " & Left$(texto, 400) & vbCrLf & vbCrLf & _
                     "Fica registrado com o seu usuário (" & Quem() & ") e a data/hora.", vbQuestion) Then
        RegistrarParecerCore = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: nada foi registrado."
        Exit Function
    End If
    On Error GoTo falha
    agora = Now
    quemTxt = Quem()
    Set pw = Aba(ABA_PAR)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscritaRapida(pw)
    uP = UltimaLinha(pw, PA_C_ID)
    If uP < 1 Then uP = 1
    id = "PAR-" & Format$(uP, "00000")
    pw.Cells(uP + 1, PA_C_ID).Value = id
    pw.Cells(uP + 1, PA_C_DATA).Value = agora
    pw.Cells(uP + 1, PA_C_QUEM).Value = quemTxt
    pw.Cells(uP + 1, PA_C_LOTE).NumberFormat = "@"
    pw.Cells(uP + 1, PA_C_LOTE).Value = lote
    pw.Cells(uP + 1, PA_C_NOBS).Value = nP
    pw.Cells(uP + 1, PA_C_TEXTO).Value = texto
    mSeguranca.RestaurarProtecao pw, prot
    prot = False
    gravado = True
    If nP > 0 Then
        protO = mSeguranca.LiberarEscritaRapida(o)
        For i = 1 To UBound(v, 1) - 1
            If StrComp(Txt(v(i, OB_C_LOTE)), lote, vbTextCompare) = 0 And UCase$(Txt(v(i, OB_C_ST))) = ST_PEND Then
                v(i, OB_C_ST) = ST_CIENTE
                v(i, OB_C_STEM) = agora
                v(i, OB_C_STPOR) = quemTxt
                v(i, OB_C_PARECER) = id
            End If
        Next i
        o.Range(o.Cells(2, 1), o.Cells(ult, OB_NCOL)).Value = SemUltimaLinha(v)
        mSeguranca.RestaurarProtecao o, protO
        protO = False
    End If
    mAuditoria.Auditar mAuditoria.CAT_CONFIG, "REFERENCIA_PARECER_REGISTRADO", MODULO, 0, Empty, "", lote, 0, "", _
        "", id, "", "", nP & " observacao(oes) com ciencia", Left$(texto, 1000) & " (" & quemTxt & ")"
    ' A176 e mesclada (A176:U179): limpar a area mesclada inteira (ClearContents numa celula so da erro 1004)
    protP = mSeguranca.LiberarEscritaRapida(ws)
    ws.Cells(R_PAR_TXT, 1).MergeArea.ClearContents
    mSeguranca.RestaurarProtecao ws, protP
    protP = False
    Application.EnableEvents = ev
    AtualizarVistaReferencia
    AtualizarVistaObs
    Avisar "Parecer " & id & " registrado no lote " & lote & ": ciência dada a " & nP & " observação(ões).", vbInformation
    RegistrarParecerCore = "OK|" & id & "|" & nP
    Exit Function
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If Not pw Is Nothing Then mSeguranca.RestaurarProtecao pw, prot
    If Not o Is Nothing Then mSeguranca.RestaurarProtecao o, protO
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, protP
    Application.EnableEvents = True
    If gravado Then
        ' o parecer ja esta na RefBulaParecer: nao induzir o analista a registrar de novo
        RegistrarParecerCore = Falhar("O parecer " & id & " foi registrado, mas a tela não terminou de atualizar (" & sE & _
                                      "). Não registre de novo: o último parecer aparece em A" & R_PAR_ULT & ".")
    Else
        RegistrarParecerCore = Falhar("Falha ao registrar o parecer: " & sE)
    End If
End Function

' ============================ ENTRADA DA SECAO ============================
' Worksheet_Change do Painel em B98:C98, E98, G98 (lote e periodo da secao): tranca a digitacao manual de outra chave; refaz a apuracao, a grade e as
' observacoes do lote, e verifica a referencia (alerta so para observacao nova).
Public Sub SecaoMudou()
    Dim dados As Variant, res As String, novas As String, nE As Long, sE As String
    If mDentro Then Exit Sub
    If Not SecaoInstalada() Then Exit Sub
    ManualContextoMudou                        ' adendo: a liberacao manual e de UMA chave lote|analito
    mDentro = True
    mApp.InicioUsuario "Apurando os alvos do lote " & LoteSecao() & "..."
    On Error GoTo falha
    dados = mDados.CarregarDB()
    AtualizarVistaAlvos dados
    res = VerificarNucleo(dados, novas)
    AtualizarVistaReferencia
    AtualizarVistaObs
    mApp.fim
    mDentro = False
    AlertarNovas res, novas
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    mApp.fim
    mDentro = False
    Avisar "Falha ao atualizar a seção de alvos/referência: " & sE, vbExclamation
End Sub

' ============================ ADR-103: TABELA DE DIGITACAO DA BULA (N114) E TRANSFERENCIA ============================
' A partir de N114 (instalar_adr103): a MESMA estrutura da grade da referencia (A114 em diante: Analito, Media/LI/LS
' por nivel e Fonte), deslocada 13 colunas -- coluna da tabela de digitacao = coluna da grade + 13; as linhas sao as
' mesmas (115..154, o analito espelhado da coluna A). Digitar ali NAO grava nada. O botao "TRANSFERIR VALORES PARA
' REFERÊNCIA INICIAL" copia SO OS VALORES preenchidos (nunca formula, nem formula como texto, nem vinculo) para a grade
' da referencia do lote em B98 e entrega as celulas copiadas ao GradeEditada -- a mesma validacao, gravacao na
' RefBulaStore, trilha e verificacao da digitacao direta na grade. Celula vazia na digitacao nao mexe no destino. A
' tabela de digitacao nao e alterada.

Public Sub TransferirReferenciaInicial()
    Dim r As String
    r = TransferirReferenciaCore()
End Sub

' Nucleo do botao (o QA chama este, com DefinirRespostasTeste). "OK|valores copiados|linhas", "CANCELADO|motivo" ou
' "ERRO|motivo". Nada e gravado fora do OK.
Public Function TransferirReferenciaCore() As String
    Dim ws As Worksheet, lote As String, i As Long, c As Long, nCol As Long, rr As Long, v As Variant, d As Variant
    Dim invalidos As String, nVal As Long, nSobre As Long, lin As Object, alvo As Range, ev As Boolean, prot As Boolean
    Dim an As String, cabG As String, cabD As String, nE As Long, sE As String, msg As String, k As Variant
    Dim li As Variant, ls_ As Variant
    If Not SecaoInstalada() Then TransferirReferenciaCore = "ERRO|secao do ADR-102 nao instalada no Painel": Exit Function
    Set ws = Pn()
    lote = LoteSecao()
    If Len(lote) = 0 Or UCase$(lote) = "TODOS" Then
        TransferirReferenciaCore = Falhar("Escolha o LOTE em B98 antes de transferir: a referência é do lote. Nada foi transferido.")
        Exit Function
    End If
    nCol = ColFonte()
    ' estrutura: cabecalho e analitos iguais aos da grade
    For c = 1 To nCol
        cabG = Txt(ws.Cells(R_DIG_CAB, c).Value)
        cabD = Txt(ws.Cells(R_DIG_CAB, c + DESLOC_DIG).Value)
        If StrComp(cabG, cabD, vbBinaryCompare) <> 0 Then
            TransferirReferenciaCore = Falhar("A tabela de digitação (N114) não tem a mesma estrutura da grade da referência " & _
                "(coluna " & c & ": """ & cabD & """ em vez de """ & cabG & """). Nada foi transferido.")
            Exit Function
        End If
    Next c
    For i = 0 To N_REF - 1
        rr = R_REF1 + i
        If StrComp(Txt(ws.Cells(rr, 1).Value), Txt(ws.Cells(rr, 1 + DESLOC_DIG).Value), vbTextCompare) <> 0 Then
            TransferirReferenciaCore = Falhar("O analito da linha " & rr & " da tabela de digitação não é o mesmo da grade (" & _
                Txt(ws.Cells(rr, 1).Value) & "). Nada foi transferido.")
            Exit Function
        End If
    Next i
    ' valores: numero nas colunas de Media/LI/LS; texto sem "=" na Fonte; LI <= LS no mesmo nivel
    Set lin = CreateObject("Scripting.Dictionary")
    For i = 0 To N_REF - 1
        rr = R_REF1 + i
        an = Txt(ws.Cells(rr, 1).Value)
        For c = 2 To nCol
            v = ws.Cells(rr, c + DESLOC_DIG).Value
            If IsError(v) Then
                invalidos = invalidos & vbCrLf & "  linha " & rr & " (" & an & "), coluna " & Txt(ws.Cells(R_DIG_CAB, c).Value) & ": erro"
            ElseIf Len(Trim$(CStr(v))) > 0 Then
                If c < nCol And Not NumOk(v) Then
                    invalidos = invalidos & vbCrLf & "  " & an & ", " & Txt(ws.Cells(R_DIG_CAB, c).Value) & ": não é número"
                ElseIf c = nCol And Left$(Trim$(CStr(v)), 1) = "=" Then
                    invalidos = invalidos & vbCrLf & "  " & an & ", Fonte: fórmula como texto não é aceita"
                ElseIf Len(an) = 0 Then
                    invalidos = invalidos & vbCrLf & "  linha " & rr & ": valor sem analito"
                Else
                    nVal = nVal + 1
                    If Not lin.Exists(rr) Then lin.Add rr, 1
                    d = ws.Cells(rr, c).Value
                    If Len(Trim$(CStr(d))) > 0 Then
                        If CStr(d) <> CStr(v) Then nSobre = nSobre + 1
                    End If
                End If
            End If
        Next c
        For c = 1 To NLV()                      ' LI <= LS com o valor que ficara na grade
            li = ws.Cells(rr, ColRef(c, 1) + DESLOC_DIG).Value
            If Len(Trim$(CStr(li))) = 0 Then li = ws.Cells(rr, ColRef(c, 1)).Value
            ls_ = ws.Cells(rr, ColRef(c, 2) + DESLOC_DIG).Value
            If Len(Trim$(CStr(ls_))) = 0 Then ls_ = ws.Cells(rr, ColRef(c, 2)).Value
            If NumOk(li) And NumOk(ls_) Then
                If CDbl(li) > CDbl(ls_) Then invalidos = invalidos & vbCrLf & "  " & an & " N" & c & ": LI maior que LS"
            End If
        Next c
    Next i
    If Len(invalidos) > 0 Then
        TransferirReferenciaCore = Falhar("Valores inválidos na tabela de digitação; nada foi transferido:" & invalidos)
        Exit Function
    End If
    If nVal = 0 Then
        TransferirReferenciaCore = Falhar("A tabela de digitação (a partir de N114) está vazia. Nada foi transferido.")
        Exit Function
    End If
    msg = "Transferir " & nVal & " valor(es) de " & lin.Count & " analito(s) para a referência inicial (bula / Unity) " & _
          "do lote " & lote & "?" & vbCrLf & vbCrLf & "Só os valores preenchidos são copiados (nunca fórmulas); célula " & _
          "vazia na digitação não altera o destino; a tabela de digitação fica como está."
    If nSobre > 0 Then msg = msg & vbCrLf & vbCrLf & "ATENÇÃO: " & nSobre & " valor(es) já transferido(s) antes serão " & _
                             "SUBSTITUÍDOS (o anterior fica no Audit_Log)."
    If Not Perguntar(msg, IIf(nSobre > 0, vbExclamation, vbQuestion)) Then
        TransferirReferenciaCore = "CANCELADO|confirmacao"
        mUltimoAviso = "Cancelado: nada foi transferido."
        Exit Function
    End If
    ev = Application.EnableEvents
    On Error GoTo falha
    Application.EnableEvents = False             ' o GradeEditada e chamado UMA vez, com todas as celulas copiadas
    prot = mSeguranca.LiberarEscritaRapida(ws)
    For Each k In lin.Keys
        For c = 2 To nCol
            v = ws.Cells(k, c + DESLOC_DIG).Value
            If Len(Trim$(CStr(v))) > 0 Then
                If c < nCol Then ws.Cells(k, c).Value = CDbl(v) Else ws.Cells(k, c).Value = CStr(v)
                If alvo Is Nothing Then Set alvo = ws.Cells(k, c) Else Set alvo = Union(alvo, ws.Cells(k, c))
            End If
        Next c
    Next k
    mSeguranca.RestaurarProtecao ws, prot
    prot = False
    Application.EnableEvents = ev
    GradeEditada alvo
    Avisar "Transferidos " & nVal & " valor(es) de " & lin.Count & " analito(s) para a referência inicial do lote " & _
           lote & ".", vbInformation
    TransferirReferenciaCore = "OK|" & nVal & "|" & lin.Count
    Exit Function
falha:
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    If prot Then mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    On Error GoTo 0
    TransferirReferenciaCore = Falhar("Falha ao transferir os valores: " & sE)
End Function

' Instalador e QA: refaz todas as vistas da secao (sem alerta). "OK|<verificacao>" ou "ERRO|motivo".
Public Function AtualizarTudo() As String
    Dim dados As Variant, novas As String, r As String
    On Error GoTo falha
    If Not SecaoInstalada() Then AtualizarTudo = "ERRO|secao nao instalada": Exit Function
    dados = mDados.CarregarDB()
    AtualizarVistaAlvos dados
    r = VerificarNucleo(dados, novas)
    AtualizarVistaReferencia
    AtualizarVistaObs
    AtualizarTudo = "OK|" & r
    Exit Function
falha:
    AtualizarTudo = "ERRO|" & Err.Description
End Function
' ADR-102 (entrega 48, achado do gate qa_adr090 V02): o PainelEncaixar muda a altura das linhas da faixa
' dos graficos e o Excel desloca os botoes que ficam abaixo dela (medido: o "Registrar parecer" desceu
' 26 pt e cobriu o "Ultimo parecer" da linha 181). Recoloca cada botao da secao na sua celula, centrado
' na altura da linha, como o instalar_adr102 faz. Chamado pelo PainelEncaixar com a aba liberada.
' So mexe na posicao (nunca no tamanho, texto ou macro); botao ausente e ignorado.
Public Sub AlinharBotoes(ByVal ws As Worksheet)
    Dim nomes As Variant, cels As Variant, i As Long, b As Shape, c As Range, t As Double
    On Error Resume Next
    nomes = Array("adr102_aplicar", "adr102_liberar", "adr102_manual_gravar", "adr102_parecer", "adr103_transferir")
    cels = Array("K98", "A111", "K111", "A180", "N155")
    For i = LBound(nomes) To UBound(nomes)
        Set b = Nothing
        Set b = ws.Shapes(nomes(i))
        If Not b Is Nothing Then
            Set c = ws.Range(cels(i))
            t = c.Top + (c.Height - b.Height) / 2
            If t < c.Top Then t = c.Top
            If Abs(b.Top - t) > 0.5 Then b.Top = t
            If Abs(b.Left - (c.Left + 2)) > 0.5 Then b.Left = c.Left + 2
        End If
    Next i
    ' ADR-103: "Aplicar a vários analitos" logo a direita do "Aplicar Alvos ao Lote", na mesma linha
    Dim a As Shape
    Set a = Nothing: Set b = Nothing
    Set a = ws.Shapes("adr102_aplicar")
    Set b = ws.Shapes("adr103_varios")
    If Not a Is Nothing And Not b Is Nothing Then
        If Abs(b.Top - a.Top) > 0.5 Then b.Top = a.Top
        If Abs(b.Left - (a.Left + a.Width + 6)) > 0.5 Then b.Left = a.Left + a.Width + 6
    End If
End Sub
