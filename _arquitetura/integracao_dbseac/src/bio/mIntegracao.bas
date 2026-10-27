Attribute VB_Name = "mIntegracao"
Option Explicit
' ============================================================================
'  ORQUESTRACAO DA CAMADA DE DADOS (ADR-057)
'
'  O Power Query e o motor de dados: recebe o DB_SEAC, junta os manuais,
'  aplica a inativacao e produz a tblCQ_Final. Este modulo NAO decide nada
'  sobre dado -- so:
'    1. prepara as tabelas de entrada (ID normalizado, carimbo de data/usuario,
'       caixa REGISTRAR - LJ marcada por padrao, ID MAN_nnnn automatico). A aba
'       Inativar (ADR-070) e so ID + ANALITO + caixa + MOTIVO + carimbo: quem
'       confere se o ID e daquele analito e o Power Query (QA E10/E11);
'    2. dispara o refresh das consultas NA ORDEM, de forma SINCRONA
'       (BackgroundQuery = False): cada Refresh so devolve quando a consulta
'       terminou -- ou levanta o erro dela. Nenhuma etapa seguinte roda sobre
'       tabela pela metade, e falha nunca vira "Atualizacao concluida";
'    3. so entao atualiza os consumidores (motor, Westgard, Estatistica, LJ);
'    4. registra no Audit_Log o que mudou de estado (inativado, reativado,
'       plotagem alterada, resultado manual incluido/retirado).
'  ADR-079: com Cfg_Integracao MOTOR_DADOS = VBA o passo 2 e o mEtlSync (mesmas
'  tabelas, calculadas em memoria e gravadas so onde mudou; segundos em vez de
'  minutos). As consultas ficam no arquivo: MOTOR_DADOS = PQ volta ao refresh.
'  Toda execucao (VBA ou PQ, ok ou falha) deixa uma linha no Sync_Log.
'  ADR-081: inativar/reativar pelo grafico sem dialogo aqui (DadosDoPonto, InativarPorId, ReativarPorId) e o
'  ATUALIZAR so das entradas (AtualizarSoEntradas); a janela e do mUI.
'  ADR-082: coluna SUBSTITUI_ID da aba Digitar Resultados -- o manual substitui o resultado do interfaceamento
'  indicado (quem decide e o motor de dados). Aqui: normalizar o ID, preencher a linha a partir do resultado indicado
'  (EntradaMudou), a trilha da substituicao e da edicao do manual (Audit_Log) e o gancho da trilha do Westgard
'  (ADR-083: mEstatistica.AuditarFotosWestgard, todos os lotes x equipamento).
' ============================================================================
Private Const ABA_MANUAL As String = "Digitar Resultados"
Private Const ABA_INAT As String = "Inativar"
Private Const ABA_COMENT As String = "COMENTARIOS_TECNICOS"
Private Const TB_MANUAL As String = "tblResultados_Manuais"
Private Const TB_INAT As String = "tblInativacao_NaoConformes"
Private Const TB_COMENT As String = "tblComentariosTecnicos"
Private Const COL_LJ As String = "REGISTRAR - LJ"
Private Const FOLGA As Long = 100              ' linhas acrescentadas quando a tabela de entrada enche
Private Const TXT_FALHA As String = "Nada foi alterado. Tente de novo ou clique em ATUALIZAR DADOS."   ' ADR-081
Private Const COL_SUBST As String = "SUBSTITUI_ID"                      ' ADR-082: coluna opcional da tblResultados_Manuais

' ordem do refresh = ordem das camadas
Private Function Camadas() As Variant
    Camadas = Array("tblDB_Recebimento", "tblDB_Organizado", "tblCQ_Final", "tblQA_Integracao")
End Function

' ============================================================================
'  BOTAO  ATUALIZAR DADOS
' ============================================================================
Public Sub AtualizarDados()
    Dim volta As Object, resumo As String
    Set volta = mApp.TelaAtual()
    On Error GoTo falha
    resumo = AtualizarDadosCore()
    mApp.VoltarPara volta
    MsgBox resumo, IIf(ContarQA("ERRO") > 0, vbExclamation, vbInformation), "Atualizar dados"
    mLotes.AvisarLotesSemValidade          ' ADR-060: lote novo sem validade -> "vamos inserir agora?"
    Exit Sub
falha:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    mApp.VoltarPara volta
    MsgBox "A atualização NÃO foi concluída." & vbCrLf & vbCrLf & sE & vbCrLf & vbCrLf & _
           "Os gráficos e a Estatística continuam com os dados da atualização anterior. " & _
           "Confira se o DB_SEAC está acessível (aba Cfg_Integracao, CAMINHO_DB_SEAC)." & vbCrLf & vbCrLf & _
           "Fora da rede do laboratório, para trabalhar só com o histórico já recebido: " & _
           "Cfg_Integracao, MODO_FONTE = HISTORICO.", _
           vbCritical, "Atualizar dados"
End Sub

' Nucleo do botao, sem caixa de dialogo (o QA automatizado chama este). Devolve
' o resumo; em falha levanta o erro com a ETAPA que falhou e NAO atualiza os
' consumidores -- grafico e Estatistica ficam com a atualizacao anterior.
' ADR-081: soEntradas = so aplicar as entradas (inativar pelo grafico), sem ler o DB_SEAC. Vale no motor VBA
' (Sincronizar sem a origem); no PQ e o ATUALIZAR normal. NAO e o modo historico: a configuracao MODO_FONTE
' nao muda, o Audit_Log nao recebe ATUALIZACAO_MODO_HISTORICO e o QA nao ganha o A07.
Public Function AtualizarDadosCore(Optional ByVal soEntradas As Boolean = False, _
                                   Optional ByVal analitoAlvo As String = "") As String
    Dim camada As Variant, t0 As Double, tempos As String
    Dim antes As Object, depois As Object, nRec0 As Long, nRec1 As Long, etapa As String, hist As Boolean
    Dim resLotes As String, vba As Boolean, nNovos As Long, mudou As Boolean, semOrigem As Boolean
    Dim tWg As Double, tWgS As Double, resWg As String
    mApp.InicioUsuario "Atualizando dados: preparando as entradas..."
    On Error GoTo falha
    hist = ModoHistorico()
    vba = mEtlSync.MotorVBA()
    semOrigem = hist Or (soEntradas And vba)
    mEtlSync.IniciarLog IIf(vba, "VBA", "PQ"), IIf(hist, "HISTORICO", IIf(semOrigem, "SO_ENTRADAS", "SEAC"))

    etapa = "preparar as tabelas de entrada"
    PrepararEntradas
    nRec0 = Linhas("tblDB_Recebimento")
    Set antes = EstadoAuditavel()

    If vba Then
        etapa = "atualizar os dados (VBA)"
        If semOrigem And Not hist Then
            Application.StatusBar = "Aplicando a inativação: recalculando sem ler o DB_SEAC..."
        Else
            Application.StatusBar = "Atualizando dados: lendo o DB_SEAC e recalculando..."
        End If
        ' ADR-081: no "so entradas" os achados do QA que dependem da origem ficam como estavam (mEtlSync)
        tempos = mEtlSync.Sincronizar(semOrigem, nNovos, mudou, resLotes, semOrigem And Not hist)
    Else
        For Each camada In Camadas()
            ' ADR-060: lote novo entra no cadastro ANTES do QA, que confere o cadastro (I03)
            If CStr(camada) = "tblQA_Integracao" Then
                etapa = "registrar lotes novos"
                resLotes = mLotes.RegistrarLotesRecebidos()
            End If
            etapa = "atualizar " & camada
            Application.StatusBar = "Atualizando dados: " & camada & "..."
            t0 = Timer
            Refrescar CStr(camada)                       ' sincrono; erro sobe e para tudo
            tempos = tempos & IIf(Len(tempos) > 0, " · ", "") & camada & " " & Format$(Timer - t0, "0.0") & "s"
        Next camada
    End If
    nRec1 = Linhas("tblDB_Recebimento")

    etapa = "conferir a DB_CQ_FINAL"
    If Linhas(FONTE_CQ) = 0 Then Err.Raise vbObjectError + 571, , "a tabela Principal - Resultados voltou vazia"

    etapa = "registrar a auditoria"
    Set depois = EstadoAuditavel()
    AuditarMudancas antes, depois

    ' ---- consumidores: so depois de TODAS as consultas terminarem ----
    etapa = "atualizar o motor (Levey-Jennings, Westgard, Estatistica)"
    Application.StatusBar = "Atualizando Levey-Jennings, Westgard e Estatistica..."
    t0 = Timer
    mDados.IniciarSnapshot                         ' ADR-079: uma leitura da tblCQ_Final para o motor inteiro
    mEstatistica.InvalidarCache
    RodarSeExistir "InvalidarCacheEstat"           ' Estatistica por periodo (so Bioquimica)
    ' ADR-081: inativar/reativar nao cria ano novo; a lista de anos e refeita no proximo ATUALIZAR completo
    If Not soEntradas Then mDados.AtualizarListasAno
    mLotes.AtualizarListaLiberacao
    mEstatistica.AtualizarEstatistica
    ' ADR-083 (opcao completa, 27/10/2026): trilha das mudancas de veredicto de Westgard em TODOS os lotes x equipamento
    ' (gancho pedido no ADR-082). A foto dos eventos OFICIAIS fica na aba tecnica Westgard_Fotos; so os lotes cuja
    ' impressao digital mudou (resultado novo, inativacao, reativacao, manual, substituicao, media/DP) sao refeitos, em
    ' memoria -- Eventos_Westgard e o Painel continuam no lote em que estavam. Logo depois de AtualizarEstatistica (o
    ' snapshot do banco ja esta carregado), antes do periodo e da incerteza; vale tambem para o caminho rapido da
    ' inativacao pelo grafico (AtualizarSoEntradas). A foto so muda depois de auditar: se algo falhar aqui ou depois,
    ' o proximo ATUALIZAR compara com a foto antiga e nada se perde. Nunca derruba a atualizacao; o tempo vai no TEMPOS.
    tWg = Timer
    resWg = ""
    On Error Resume Next
    resWg = mEstatistica.AuditarFotosWestgard()
    If Err.Number <> 0 Then resWg = "ERRO|" & Err.Description
    Err.Clear
    On Error GoTo falha
    tWgS = Timer - tWg
    If tWgS < 0 Then tWgS = tWgS + 86400#              ' meia-noite
    RodarSeExistir "RecalcularEstatPeriodo"
    ' ADR-064: incerteza de medicao (CIQ + CEQ). ADR-081: depois de inativar/reativar UM resultado so a incerteza do
    ' analito dele muda (a dos outros depende so dos resultados deles; o CEQ nao muda) -- recalcula so as linhas dele
    mIncerteza.RecalcularIncerteza IIf(soEntradas, analitoAlvo, "")
    mDados.EncerrarSnapshot
    tempos = tempos & " · motor " & Format$(Timer - t0, "0.0") & "s" & _
             " · westgard_fotos " & Format$(tWgS, "0.00") & "s [" & resWg & "]"      ' ADR-083 (dentro do motor)
    mEtlSync.AnotarLog "T_MOTOR", Round(Timer - t0, 2)
    mEtlSync.AnotarLog "QA_ERROS", ContarQA("ERRO")
    mEtlSync.AnotarLog "QA_ALERTAS", ContarQA("ALERTA")
    If Not vba Then
        mEtlSync.AnotarLog "NOVOS", nRec1 - nRec0
        mEtlSync.AnotarLog "TOTAL_RECEBIDO", nRec1
        mEtlSync.AnotarLog "LINHAS_CQ", Linhas(FONTE_CQ)
    End If
    mEtlSync.AnotarLog "TEMPOS", tempos
    mEtlSync.FecharLog IIf(vba And Not mudou, "OK_SEM_MUDANCA", "OK"), ""

    If hist Then mAuditoria.RegistrarLog "ATUALIZACAO_MODO_HISTORICO", _
        "MODO_FONTE = HISTORICO: DB_SEAC nao lido; reprocessado o historico ja recebido (" & nRec1 & " resultados)"

    mApp.fim
    AtualizarDadosCore = IIf(hist, AvisoHistorico() & vbCrLf & vbCrLf, "") & _
           IIf(semOrigem And Not hist, "Só as entradas (inativações, manuais e comentários) foram aplicadas; o DB_SEAC " & _
               "não foi lido -- os resultados novos entram no próximo ATUALIZAR DADOS." & vbCrLf & vbCrLf, "") & _
           "Atualização concluída." & vbCrLf & vbCrLf & _
           "Resultados recebidos: " & Format$(nRec1, "#,##0") & " (novos nesta atualização: " & (nRec1 - nRec0) & ")" & vbCrLf & _
           UltimoResultado() & _
           "Principal - Resultados: " & Format$(Linhas(FONTE_CQ), "#,##0") & " resultados" & vbCrLf & _
           ResumoQA() & ResumoLotes(resLotes) & vbCrLf & vbCrLf & "Tempo: " & tempos
    Exit Function
falha:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    mDados.EncerrarSnapshot
    ' ADR-083 (opcao completa): nada a fazer aqui -- a foto por lote x equipamento (Westgard_Fotos) so muda depois de
    ' auditada; o proximo ATUALIZAR compara com ela
    If vba And Len(mEtlSync.EtapaAtual()) > 0 Then etapa = etapa & " / " & mEtlSync.EtapaAtual()
    mAuditoria.RegistrarLog "ATUALIZACAO_FALHOU", "Etapa: " & etapa & " | Erro " & nE & ": " & sE
    mEtlSync.FecharLog "ERRO", "Etapa: " & etapa & " | Erro " & nE & ": " & sE
    mApp.fim
    On Error GoTo 0
    Err.Raise vbObjectError + 579, "mIntegracao.AtualizarDados", "Etapa: " & etapa & vbCrLf & "Erro " & nE & ": " & sE
End Function

' Entrada para AUTOMACAO (script, agendamento, QA): o mesmo nucleo do botao, mas
' NUNCA abre caixa de dialogo -- erro nao tratado chamado de fora (Application.Run)
' abriria a janela de depuracao do VBA na tela do usuario. Devolve
' "OK|<resumo>" ou "ERRO|<etapa e erro>".
Public Function AtualizarDadosAutomatico() As String
    On Error GoTo falha
    AtualizarDadosAutomatico = "OK|" & AtualizarDadosCore()
    Exit Function
falha:
    AtualizarDadosAutomatico = "ERRO|" & Err.Description
End Function

' ADR-081: o ATUALIZAR depois de inativar/reativar pelo grafico. O mesmo nucleo, sem ler o DB_SEAC (motor VBA;
' no PQ e o ATUALIZAR normal), sem caixa de dialogo. Devolve "OK|<resumo>" ou "ERRO|<etapa e erro>".
' analito = o do resultado inativado/reativado (so a incerteza dele e recalculada); vazio = todos.
Public Function AtualizarSoEntradas(Optional ByVal analito As String = "") As String
    On Error GoTo falha
    AtualizarSoEntradas = "OK|" & AtualizarDadosCore(True, analito)
    Exit Function
falha:
    AtualizarSoEntradas = "ERRO|" & Err.Description
End Function

' ADR-058. MODO_FONTE = HISTORICO: a consulta SEAC_ORIGEM nao le o DB_SEAC e as
' camadas reprocessam so o que ja foi recebido. SEAC (ou vazio) e o modo normal.
' Aceita a grafia acentuada: UCase$("Histórico") = "HISTÓRICO" (mesma armadilha
' do item 7.4 do QUALITY_GATE).
Public Function ModoHistorico() As Boolean
    Dim m As String
    m = UCase$(Trim$(Cfg("MODO_FONTE")))
    ModoHistorico = (m = "HISTORICO" Or m = "HISTÓRICO")
End Function

Private Function AvisoHistorico() As String
    AvisoHistorico = "ATENÇÃO - MODO HISTÓRICO: o DB_SEAC NÃO foi lido. Nenhum resultado novo entrou; " & _
                     "inativações, manuais e comentários foram reprocessados sobre o histórico já recebido." & vbCrLf & _
                     "Na rede do laboratório, volte Cfg_Integracao MODO_FONTE = SEAC."
End Function

' Refresh SINCRONO de uma tabela de consulta. BackgroundQuery = False faz o
' Refresh so devolver o controle quando a consulta terminou; um erro do Power
' Query sobe como erro de VBA. A conferencia de Refreshing e a garantia extra
' de que nada ficou pendente antes de a proxima camada ler esta.
'
' O DB_SEAC de producao e regravado pela tarefa agendada a cada 2 minutos. Ler no meio do Save
' da "nao foi possivel reconhecer a entrada como um documento do Excel valido" (DataFormat.Error)
' ou "arquivo em uso" -- medido no laboratorio em 05/10/2026. So esses erros de LEITURA DA ORIGEM
' sao tentados de novo (ate REFRESH_TENTATIVAS, REFRESH_ESPERA_S entre elas); qualquer outro
' erro sobe na hora, como antes.
Public Sub Refrescar(ByVal nomeTabela As String)
    Const REFRESH_TENTATIVAS As Long = 4
    Const REFRESH_ESPERA_S As Long = 15
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, nE As Long, sE As String, tentativa As Long
    Dim sb As Variant
    Set lo = AcharTabela(nomeTabela)
    If lo Is Nothing Then Err.Raise vbObjectError + 570, "Refrescar", "tabela " & nomeTabela & " nao encontrada"
    Set ws = lo.Parent
    prot = LiberarEscrita(ws)
    sb = Application.StatusBar
    For tentativa = 1 To REFRESH_TENTATIVAS
        nE = 0: sE = ""
        On Error Resume Next
        With lo.QueryTable
            .BackgroundQuery = False
            .Refresh BackgroundQuery:=False
            If Err.Number = 0 Then
                If .Refreshing Then Err.Raise vbObjectError + 572, "Refrescar", nomeTabela & " ainda em atualizacao"
            End If
        End With
        nE = Err.Number: sE = Err.Description
        On Error GoTo 0
        If nE = 0 Then Exit For
        If Not OrigemOcupada(sE) Or tentativa = REFRESH_TENTATIVAS Then Exit For
        Application.StatusBar = "DB_SEAC sendo gravado pela atualizacao automatica -- nova tentativa em " & _
                                REFRESH_ESPERA_S & " s (" & tentativa & "/" & REFRESH_TENTATIVAS - 1 & ")"
        Application.Wait Now + TimeSerial(0, 0, REFRESH_ESPERA_S)
    Next tentativa
    Application.StatusBar = sb
    If tentativa > 1 And nE = 0 Then mAuditoria.RegistrarLog "ORIGEM_OCUPADA", nomeTabela & ": lida na tentativa " & tentativa
    RestaurarProtecao ws, prot
    If nE <> 0 Then Err.Raise nE, "mIntegracao.Refrescar(" & nomeTabela & ")", sE
End Sub

Private Function OrigemOcupada(ByVal msg As String) As Boolean
    Dim m As String
    m = LCase$(msg)
    ' 13/10/2026: lido no meio do Save o Power Query tambem da "[Expression.Error] O valor nao foi
    ' especificado" (medido na entrega: passou na repeticao isolada) -- tratado como origem ocupada
    OrigemOcupada = (InStr(m, "dataformat.error") > 0) Or (InStr(m, "sendo usado") > 0) Or _
                    (InStr(m, "being used") > 0) Or (InStr(m, "em uso") > 0) Or _
                    (InStr(m, "foi especificado") > 0) Or (InStr(m, "was not specified") > 0)
End Function

Public Function AcharTabela(ByVal nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nome Then Set AcharTabela = lo: Exit Function
        Next lo
    Next ws
End Function

' "Resultado mais recente: dd/mm/aaaa hh:mm" -- o usuario ve na hora se os dados estao em dia
Private Function UltimoResultado() As String
    Dim lo As ListObject, v As Variant, i As Long, mx As Double
    On Error GoTo fim
    Set lo = AcharTabela("tblDB_Recebimento")
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.ListColumns("DATA_HORA").DataBodyRange.Value2
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        If VarType(v(i, 1)) = vbDouble Then If v(i, 1) > mx Then mx = v(i, 1)
    Next i
    If mx > 0 Then UltimoResultado = "Resultado mais recente recebido: " & Format$(CDate(mx), "dd/mm/yyyy hh:mm") & vbCrLf
fim:
End Function

Private Function Linhas(ByVal nome As String) As Long
    Dim lo As ListObject
    Set lo = AcharTabela(nome)
    If Not lo Is Nothing Then Linhas = lo.ListRows.Count
End Function

Private Sub RodarSeExistir(ByVal macro As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & macro
    Err.Clear
End Sub

' ============================================================================
'  QA
' ============================================================================
Public Function ContarQA(ByVal severidade As String) As Long
    Dim lo As ListObject
    Set lo = AcharTabela("tblQA_Integracao")
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    ContarQA = Application.WorksheetFunction.CountIf(lo.ListColumns("SEVERIDADE").DataBodyRange, severidade)
End Function

' ADR-060: o que o registro automatico de lotes fez nesta atualizacao.
Private Function ResumoLotes(ByVal r As String) As String
    Dim p As Variant
    If Len(r) = 0 Then Exit Function
    p = Split(r & "|||", "|")
    If p(0) = "ERRO" Then
        ResumoLotes = vbCrLf & "ATENÇÃO: o registro automático de lotes falhou (" & p(1) & "). Os dados foram atualizados."
    ElseIf Val(p(1)) > 0 Then
        ResumoLotes = vbCrLf & "Lotes novos registrados automaticamente: " & Replace(p(2), ";", ", ") & _
                      " (falta a validade)."
    End If
    If p(0) = "OK" And Len(p(3)) > 0 Then
        ResumoLotes = ResumoLotes & vbCrLf & "ATENÇÃO: cadastro de lotes cheio (100). Não couberam: " & Replace(p(3), ";", ", ")
    End If
End Function

Private Function ResumoQA() As String
    Dim e As Long, a As Long
    e = ContarQA("ERRO"): a = ContarQA("ALERTA")
    ResumoQA = "QA da integração: " & e & " erro(s), " & a & " alerta(s), " & ContarQA("INFO") & " informativo(s)."
    If e > 0 Then ResumoQA = ResumoQA & vbCrLf & "Há ERRO DE GOVERNANÇA: veja a aba QA_INTEGRACAO (ex.: inativação sem motivo, " & _
                             "ou ID que não é do analito informado -- essa não foi aplicada)."
End Function

' ============================================================================
'  AUDITORIA (ISO 15189 8.4) -- o que mudou de estado entre duas atualizacoes
' ============================================================================
' ID -> Array(estado, analito, lote, nivel, run, data, equip, resultado, comentario, valor do manual, motivo automatico,
' quem lancou o manual)
' So o que interessa a trilha: resultados INATIVADOS, SUBSTITUIDOS por manual e MANUAIS.
' ADR-070: "comentario" = MOTIVO da inativacao (aba Inativar) + COMENTARIO_TECNICO.
' ADR-082: estado = partes separadas por "|": "INATIVADO/<plotagem>" ou "SUBSTITUIDO" (status); "SUBST:<manual>" no
'   resultado do interfaceamento que um manual substitui (tambem o que continua INATIVADO); "MANUAL:<status>",
'   "VAL:<resultado>" e "DH:<data/hora>" no manual (a edicao do valor/instante de um manual fica na trilha).
'   Array de TAMANHO FIXO (0..11) para todas as linhas: x(9) = valor do manual que substitui, x(10) =
'   MOTIVO_EXCLUSAO_AUTOMATICA, x(11) = "manual lancado por <USUARIO> em <data>" (do manual que substitui, ou do
'   proprio manual: a decisao humana na trilha, ISO 15189 8.4 -- o Audit_Log grava o usuario da sessao que rodou o
'   ATUALIZAR, que pode ser a atualizacao agendada); x(8) continua o parecer da inativacao. O SUBSTITUIDO_MANUAL nao
'   tem RUN: x(4) = o do manual.
Private Function EstadoAuditavel() As Object
    Dim d As Object, lo As ListObject, n As Long, i As Long
    Dim vId, vSt, vPl, vOr, vAn, vLt, vNv, vRun, vDt, vEq, vRes, vCom, vMot, est As String, txt As String
    Dim vRel As Variant, vMex As Variant, manRes As Object, manRun As Object, rel As String, x As Variant
    Dim vUsr As Variant, vReg As Variant, manQuem As Object, uq As Variant, dq As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set EstadoAuditavel = d
    Set lo = AcharTabela(FONTE_CQ)
    If lo Is Nothing Then Exit Function
    n = lo.ListRows.Count
    If n < 2 Then Exit Function
    vId = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value
    vSt = lo.ListColumns("STATUS_ANALITICO").DataBodyRange.Value
    vPl = lo.ListColumns("TIPO_PLOTAGEM_LJ").DataBodyRange.Value
    vOr = lo.ListColumns("ORIGEM_RESULTADO").DataBodyRange.Value
    vAn = lo.ListColumns("ANALITO").DataBodyRange.Value
    vLt = lo.ListColumns("LOTE").DataBodyRange.Value
    vNv = lo.ListColumns("NIVEL").DataBodyRange.Value
    vRun = lo.ListColumns("RUN").DataBodyRange.Value
    vDt = lo.ListColumns("DATA_HORA").DataBodyRange.Value
    vEq = lo.ListColumns("EQUIPAMENTO").DataBodyRange.Value
    vRes = lo.ListColumns("RESULTADO").DataBodyRange.Value
    vCom = lo.ListColumns("COMENTARIO_TECNICO").DataBodyRange.Value
    vMot = Empty: vRel = Empty: vMex = Empty: vUsr = Empty: vReg = Empty
    On Error Resume Next                       ' coluna nova (ADR-070): ausente ate a 1a atualizacao com a consulta nova
    vMot = lo.ListColumns("MOTIVO_INATIVACAO").DataBodyRange.Value
    vRel = lo.ListColumns("ID_RELACIONADO").DataBodyRange.Value
    vMex = lo.ListColumns("MOTIVO_EXCLUSAO_AUTOMATICA").DataBodyRange.Value
    vUsr = lo.ListColumns("USUARIO_MANUAL").DataBodyRange.Value
    vReg = lo.ListColumns("RECEBIDO_EM").DataBodyRange.Value         ' no manual = REGISTRADO_EM (carimbo da digitacao)
    On Error GoTo 0
    ' ADR-082: valor, RUN e quem lancou cada manual (o SUBSTITUIDO_MANUAL fica sem RUN: a trilha registra o do manual)
    Set manRes = CreateObject("Scripting.Dictionary")
    Set manRun = CreateObject("Scripting.Dictionary")
    Set manQuem = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If CStr(vOr(i, 1)) = "MANUAL" Then
            manRes(CStr(vId(i, 1))) = vRes(i, 1)
            manRun(CStr(vId(i, 1))) = vRun(i, 1)
            uq = Empty: dq = Empty
            If IsArray(vUsr) Then uq = vUsr(i, 1)
            If IsArray(vReg) Then dq = vReg(i, 1)
            manQuem(CStr(vId(i, 1))) = TextoQuemLancou(uq, dq)
        End If
    Next i
    For i = 1 To n
        est = ""
        rel = ""
        If IsArray(vRel) Then rel = CStr(vRel(i, 1))
        If CStr(vSt(i, 1)) = "INATIVADO" Then
            est = "INATIVADO/" & CStr(vPl(i, 1))
        ElseIf CStr(vSt(i, 1)) = "SUBSTITUIDO_MANUAL" Then
            est = "SUBSTITUIDO"
        End If
        ' ADR-082: o vinculo com o manual que substitui (ID_RELACIONADO de um interfaceamento que nao e retransmissao)
        If CStr(vOr(i, 1)) = "INTERFACEAMENTO" And Len(rel) > 0 And CStr(vSt(i, 1)) <> "DUPLICIDADE_ORIGEM" Then
            est = est & "|SUBST:" & rel
        End If
        If CStr(vOr(i, 1)) = "MANUAL" Then
            est = est & "|MANUAL:" & CStr(vSt(i, 1)) & "|VAL:" & TextoAudit(vRes(i, 1)) & "|DH:" & TextoAudit(vDt(i, 1))
        End If
        If Len(est) > 0 Then
            txt = CStr(vCom(i, 1))
            If IsArray(vMot) Then
                If Len(CStr(vMot(i, 1))) > 0 Then txt = "MOTIVO: " & CStr(vMot(i, 1)) & IIf(Len(txt) > 0, " | " & txt, "")
            End If
            x = Array(est, vAn(i, 1), vLt(i, 1), vNv(i, 1), vRun(i, 1), vDt(i, 1), vEq(i, 1), vRes(i, 1), txt, Empty, Empty, "")
            If InStr(est, "|SUBST:") > 0 Then
                If manRes.Exists(rel) Then
                    x(9) = manRes(rel)
                    If Len(CStr(x(4))) = 0 Then x(4) = manRun(rel)
                    x(11) = manQuem(rel)
                End If
            End If
            If CStr(vOr(i, 1)) = "MANUAL" Then x(11) = manQuem(CStr(vId(i, 1)))
            If IsArray(vMex) Then x(10) = CStr(vMex(i, 1))
            d(CStr(vId(i, 1))) = x
        End If
    Next i
End Function

' ADR-082: "manual lancado por <USUARIO_MANUAL> em <dd/mm/yyyy hh:nn>" ("" sem os dois). Nunca levanta erro.
Private Function TextoQuemLancou(ByVal usuario As Variant, ByVal quando As Variant) As String
    Dim u As String, s As String
    If Not (IsEmpty(usuario) Or IsNull(usuario) Or IsError(usuario)) Then u = Trim$(CStr(usuario))
    If Len(u) > 0 Then s = " por " & u
    If Not IsError(quando) Then
        If IsDate(quando) Then s = s & " em " & Format$(CDate(quando), "dd/mm/yyyy hh:nn")
    End If
    If Len(s) > 0 Then TextoQuemLancou = "manual lancado" & s
End Function

' Texto estavel de um valor da tblCQ_Final para o estado auditavel (numero, data/hora ou texto; vazio = "")
Private Function TextoAudit(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        TextoAudit = Format$(v, "yyyy-mm-dd hh:nn:ss")
    ElseIf VarType(v) = vbString Then
        TextoAudit = v
    ElseIf IsNumeric(v) Then
        TextoAudit = CStr(CDbl(v))
    Else
        TextoAudit = CStr(v)
    End If
End Function

' ADR-082: a parte "<rotulo>valor" do estado auditavel ("" se nao houver). Toda parte rotulada vem depois de um "|".
Private Function ParteEstado(ByVal est As String, ByVal rotulo As String) As String
    Dim p As Long, q As Long
    p = InStr(1, est, "|" & rotulo, vbBinaryCompare)
    If p = 0 Then Exit Function
    p = p + Len(rotulo) + 1
    q = InStr(p, est, "|", vbBinaryCompare)
    If q = 0 Then ParteEstado = Mid$(est, p) Else ParteEstado = Mid$(est, p, q - p)
End Function

' O status no estado auditavel (antes do 1o "|"): "INATIVADO/X_VERMELHO", "SUBSTITUIDO" ou ""
Private Function StatusEstado(ByVal est As String) As String
    Dim p As Long
    p = InStr(1, est, "|", vbBinaryCompare)
    If p = 0 Then StatusEstado = est Else StatusEstado = Left$(est, p - 1)
End Function

Private Sub JuntarAcao(ByRef acao As String, ByVal nova As String)
    acao = acao & IIf(Len(acao) > 0, "+", "") & nova
End Sub

Private Sub AuditarMudancas(ByVal antes As Object, ByVal depois As Object)
    Dim k As Variant, a As String, b As String, xA As Variant
    On Error Resume Next                       ' a trilha nunca derruba a atualizacao
    For Each k In depois.Keys
        b = depois(k)(0)
        If antes.Exists(k) Then
            a = antes(k)(0): xA = antes(k)
        Else
            a = "": xA = Empty
        End If
        If a <> b Then Registrar CStr(k), a, b, xA, depois(k)
    Next k
    For Each k In antes.Keys
        If Not depois.Exists(k) Then Registrar CStr(k), antes(k)(0), "", antes(k), Empty
    Next k
End Sub

' a/b = estado antes/depois; xA/xB = os dados de antes/depois (Empty = o ID nao estava rastreado daquele lado)
Private Sub Registrar(ByVal id As String, ByVal a As String, ByVal b As String, ByVal xA As Variant, ByVal xB As Variant)
    Dim acao As String, inA As Boolean, inB As Boolean, manA As Boolean, manB As Boolean, x As Variant
    Dim sA As String, sB As String, resAnt As Variant, resNovo As Variant, nota As String, parecer As String
    Dim quem As String
    If IsArray(xB) Then x = xB Else x = xA          ' dados do resultado: os de depois (se saiu da trilha, os de antes)
    quem = CStr(x(11))                             ' ADR-082: quem lancou o manual (vazio fora das acoes do manual)
    inA = Left$(a, 9) = "INATIVADO": inB = Left$(b, 9) = "INATIVADO"
    manA = InStr(a, "|MANUAL:") > 0: manB = InStr(b, "|MANUAL:") > 0
    sA = ParteEstado(a, "SUBST:"): sB = ParteEstado(b, "SUBST:")
    resAnt = x(7): resNovo = x(7)
    If inB And Not inA Then
        acao = "RESULTADO_INATIVADO"
    ElseIf inA And Not inB Then
        acao = "RESULTADO_REATIVADO"
    ElseIf inA And inB Then
        ' ADR-082: so a parte do status (a plotagem); o vinculo e o valor do manual tem acao propria
        If StatusEstado(a) <> StatusEstado(b) Then acao = "PLOTAGEM_LJ_ALTERADA"
    End If
    ' ADR-082: substituicao pelo manual (SUBSTITUI_ID): resAnt/resNovo = o que valia antes e o que vale depois
    If Len(sB) > 0 And Len(sA) = 0 Then
        JuntarAcao acao, "RESULTADO_SUBSTITUIDO_POR_MANUAL"
        resNovo = x(9)
        If Left$(CStr(x(10)), 11) = "substituido" Then nota = CStr(x(10)) Else nota = "substituido pelo manual " & sB
        If Len(quem) > 0 Then nota = nota & " (" & quem & ")"
    ElseIf Len(sA) > 0 And Len(sB) = 0 Then
        JuntarAcao acao, "SUBSTITUICAO_POR_MANUAL_DESFEITA"
        If IsArray(xA) Then resAnt = xA(9)
        nota = "o manual " & sA & " deixou de substituir este resultado (SUBSTITUI_ID apagado ou invalido, ou manual retirado)"
    ElseIf Len(sA) > 0 And sA <> sB Then
        JuntarAcao acao, "SUBSTITUICAO_POR_MANUAL_ALTERADA"
        If IsArray(xA) Then resAnt = xA(9)
        resNovo = x(9)
        nota = "era substituido pelo manual " & sA & "; agora pelo manual " & sB
        If Len(quem) > 0 Then nota = nota & " (" & quem & ")"
    End If
    If manB And Not manA Then
        JuntarAcao acao, "RESULTADO_MANUAL_INCLUIDO"
        If Len(quem) > 0 Then nota = nota & IIf(Len(nota) > 0, " | ", "") & quem
    End If
    If manA And Not manB Then JuntarAcao acao, "RESULTADO_MANUAL_RETIRADO"
    If manA And manB Then
        If ParteEstado(a, "MANUAL:") <> ParteEstado(b, "MANUAL:") And Not (inA Or inB) Then JuntarAcao acao, "RESULTADO_MANUAL_STATUS"
        ' ADR-082: valor ou data/hora do manual editados (o numero que vale no lugar do dado do equipamento)
        If ParteEstado(a, "VAL:") <> ParteEstado(b, "VAL:") Or ParteEstado(a, "DH:") <> ParteEstado(b, "DH:") Then
            JuntarAcao acao, "RESULTADO_MANUAL_ALTERADO"
            If IsArray(xA) Then resAnt = xA(7)
            resNovo = x(7)
            If ParteEstado(a, "DH:") <> ParteEstado(b, "DH:") Then
                nota = nota & IIf(Len(nota) > 0, " | ", "") & "data/hora " & ParteEstado(a, "DH:") & " -> " & ParteEstado(b, "DH:")
            End If
            If Len(quem) > 0 Then nota = nota & IIf(Len(nota) > 0, " | ", "") & quem
        End If
    End If
    If Len(acao) = 0 Then Exit Sub
    parecer = CStr(x(8))
    If Len(nota) > 0 Then parecer = parecer & IIf(Len(parecer) > 0, " | ", "") & nota
    mAuditoria.Auditar mAuditoria.CAT_DADO, acao, "mIntegracao", CLng(Val(CStr(x(4)))), x(5), CStr(x(6)), _
                       CStr(x(2)), CLng(Val(CStr(x(3)))), CStr(x(1)), resAnt, resNovo, a, b, _
                       "ID " & id, parecer
End Sub

' ============================================================================
'  ENTRADAS: preparo em lote antes do refresh e eventos de digitacao
' ============================================================================
' Mesma normalizacao do Power Query (DB_CQ_FINAL.NormId): numero puro = ID do
' interfaceamento deste setor; MAN_7 / man_0007 -> MAN_0007.
Public Function NormalizarId(ByVal v As Variant) As String
    ' Mesma regra do NormId do PQ (DB_CQ_FINAL e QA_INTEGRACAO) -- D14 (QA-ETL-001): sem espaco nem NBSP
    ' (colado de e-mail/web) e numero sem zeros a esquerda ("00123" e "HEM-00123" = HEM-123); prefixo em maiusculas.
    Dim t As String, resto As String, pre As String
    t = UCase$(Replace(Replace(Trim$(CStr(v)), " ", ""), ChrW$(160), ""))
    If Len(t) = 0 Then Exit Function
    pre = UCase$(Trim$(Cfg("PREFIXO_ID"))) & "-"
    If SoDigitos(t) Then
        NormalizarId = pre & SemZeros(t)
    ElseIf Left$(t, Len(pre)) = pre And SoDigitos(Mid$(t, Len(pre) + 1)) Then
        NormalizarId = pre & SemZeros(Mid$(t, Len(pre) + 1))
    ElseIf Left$(t, 4) = "MAN_" Then
        resto = Mid$(t, 5)
        If SoDigitos(resto) Then NormalizarId = "MAN_" & Right$(String$(4, "0") & resto, IIf(Len(resto) > 4, Len(resto), 4)) _
        Else NormalizarId = t
    Else
        NormalizarId = t
    End If
End Function

Private Function SoDigitos(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then Exit Function
    Next i
    SoDigitos = True
End Function

Private Function SemZeros(ByVal d As String) As String
    Do While Len(d) > 1 And Left$(d, 1) = "0"
        d = Mid$(d, 2)
    Loop
    SemZeros = d
End Function

' Valor de tblConfigIntegracao (CHAVE -> VALOR).
Public Function Cfg(ByVal chave As String) As String
    Dim lo As ListObject, v As Variant, i As Long
    Set lo = AcharTabela("tblConfigIntegracao")
    If lo Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(CStr(v(i, 1)))) = UCase$(chave) Then Cfg = Trim$(CStr(v(i, 2))): Exit Function
    Next i
End Function

' Passada em lote por todas as entradas: o mesmo que o evento de digitacao faz,
' para o que tiver sido colado/importado com eventos desligados.
Public Sub PrepararEntradas()
    Dim ev As Boolean, nome As Variant, lo As ListObject, r As Long, ws As Worksheet, prot As Boolean
    Dim nE As Long, sE As String
    ev = Application.EnableEvents
    Application.EnableEvents = False
    For Each nome In Array(TB_INAT, TB_COMENT, TB_MANUAL)
        Set lo = AcharTabela(CStr(nome))
        If Not lo Is Nothing Then
            Set ws = lo.Parent
            prot = LiberarEscrita(ws)
            On Error GoTo fim
            For r = 1 To lo.ListRows.Count
                PrepararLinha lo, r
            Next r
            On Error GoTo 0
            RestaurarProtecao ws, prot
        End If
    Next nome
    Application.EnableEvents = ev
    Exit Sub
fim:
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    Err.Raise nE, "mIntegracao.PrepararEntradas", sE
End Sub

' Uma linha de uma tabela de entrada. Idempotente: so preenche o que falta.
Private Sub PrepararLinha(ByVal lo As ListObject, ByVal r As Long)
    Dim rg As Range, id As String, idN As String, temDado As Boolean, c As Variant
    Set rg = lo.ListRows(r).Range
    Select Case lo.Name
        Case TB_INAT, TB_COMENT
            id = CStr(Celula(lo, r, "ID_REGISTRO").Value)
            If Len(Trim$(id)) = 0 Then Exit Sub
            idN = NormalizarId(id)
            If idN <> id Then Celula(lo, r, "ID_REGISTRO").Value = idN
            If lo.Name = TB_INAT Then
                ' inativacao NOVA (ainda sem carimbo): REGISTRAR - LJ nasce marcado (SIM).
                ' Depois do carimbo, vale o que o usuario marcar/desmarcar.
                If Len(Trim$(CStr(Celula(lo, r, "DATA_INATIVACAO").Value))) = 0 Then Celula(lo, r, COL_LJ).Value = True
                Carimbar Celula(lo, r, "DATA_INATIVACAO"), Celula(lo, r, "USUARIO")
            Else
                If Len(Trim$(CStr(Celula(lo, r, "COMENTARIO_TECNICO").Value))) > 0 Then _
                    Carimbar Celula(lo, r, "DATA"), Celula(lo, r, "USUARIO")
            End If
        Case TB_MANUAL
            For Each c In Array("DATA", "HORA", "LOTE", "NIVEL", "ANALITO", "RESULTADO")
                If Len(Trim$(CStr(Celula(lo, r, CStr(c)).Value))) > 0 Then temDado = True: Exit For
            Next c
            id = CStr(Celula(lo, r, "ID_REGISTRO").Value)
            If Not temDado And Len(Trim$(id)) = 0 Then Exit Sub
            If Len(Trim$(id)) = 0 Then
                Celula(lo, r, "ID_REGISTRO").Value = ProximoIdManual(lo)
            Else
                idN = NormalizarId(id)
                If idN <> id Then Celula(lo, r, "ID_REGISTRO").Value = idN
            End If
            If Len(Trim$(CStr(Celula(lo, r, "EQUIPAMENTO").Value))) = 0 Then
                Celula(lo, r, "EQUIPAMENTO").Value = EquipPadrao()
            End If
            If Len(Trim$(CStr(Celula(lo, r, "MATRIZ").Value))) = 0 Then
                Celula(lo, r, "MATRIZ").Value = Cfg("MATRIZ_PADRAO")
            End If
            ' ADR-082: SUBSTITUI_ID normalizado como o ID da aba Inativar ("54565" -> "HEM-54565"). Tabela ainda sem a
            ' coluna (antes do instalar_adr082): nada
            If TemColuna(lo, COL_SUBST) Then
                id = TextoCel(lo, r, COL_SUBST)
                If Len(id) > 0 Then
                    idN = NormalizarId(id)
                    If idN <> CStr(Celula(lo, r, COL_SUBST).Value) Then Celula(lo, r, COL_SUBST).Value = idN
                End If
            End If
            Carimbar Celula(lo, r, "REGISTRADO_EM"), Celula(lo, r, "USUARIO")
    End Select
End Sub

' A tabela tem a coluna? (ADR-082: SUBSTITUI_ID so existe depois do instalar_adr082)
Private Function TemColuna(ByVal lo As ListObject, ByVal coluna As String) As Boolean
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If c.Name = coluna Then TemColuna = True: Exit Function
    Next c
End Function

Private Function Celula(ByVal lo As ListObject, ByVal r As Long, ByVal coluna As String) As Range
    Set Celula = lo.ListColumns(coluna).DataBodyRange.Cells(r, 1)
End Function

Private Sub LimparSeExistir(ByVal lo As ListObject, ByVal r As Long, ByVal coluna As String)
    On Error Resume Next                       ' tabela ainda no layout antigo: a coluna nao existe
    Celula(lo, r, coluna).ClearContents
End Sub

' ADR-070: grava uma inativacao na aba Inativar como o usuario faria (ID normalizado, ANALITO, REGISTRAR - LJ
' marcado, MOTIVO; data e usuario carimbados). Usada pelo duplo clique no grafico (mUI.InativarPeloGrafico) e
' pelo QA. NAO atualiza os dados. Devolve "OK|<linha da tabela>" ou "ERRO|<motivo>" -- nunca abre dialogo.
Public Function RegistrarInativacao(ByVal id As String, ByVal analito As String, ByVal motivo As String) As String
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, idN As String, r As Long
    Dim v As Variant, i As Long, livre As Long
    On Error GoTo falha
    idN = NormalizarId(id)
    If Len(idN) = 0 Then RegistrarInativacao = "ERRO|ID vazio": Exit Function
    If Len(Trim$(analito)) = 0 Then RegistrarInativacao = "ERRO|analito vazio": Exit Function
    If Len(Trim$(motivo)) = 0 Then RegistrarInativacao = "ERRO|motivo vazio": Exit Function
    Set lo = AcharTabela(TB_INAT)
    If lo Is Nothing Then RegistrarInativacao = "ERRO|tabela " & TB_INAT & " nao encontrada": Exit Function
    Set ws = lo.Parent
    v = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value
    If Not IsArray(v) Then
        Dim u(1 To 1, 1 To 1) As Variant: u(1, 1) = v: v = u
    End If
    For i = 1 To lo.ListRows.Count
        If NormalizarId(v(i, 1)) = idN Then
            RegistrarInativacao = "ERRO|o ID " & idN & " ja esta na aba Inativar (linha " & i & ")": Exit Function
        End If
        If livre = 0 And Len(Trim$(CStr(v(i, 1)))) = 0 Then
            If LinhaInatVazia(lo, i) Then livre = i
        End If
    Next i
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    If livre = 0 Then                          ' tabela cheia: cresce FOLGA linhas (a aba protegida nao deixa)
        lo.Resize lo.Range.Resize(lo.Range.rows.Count + FOLGA)
        livre = lo.ListRows.Count - FOLGA + 1
    End If
    r = livre
    Celula(lo, r, "ID_REGISTRO").Value = idN
    Celula(lo, r, "ANALITO").Value = Trim$(analito)
    Celula(lo, r, "MOTIVO").Value = Trim$(motivo)
    PrepararLinha lo, r                        ' caixa marcada (linha nova) e carimbo de data/usuario
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    RegistrarInativacao = "OK|" & r
    Exit Function
falha:
    Dim sE As String
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    RegistrarInativacao = "ERRO|" & sE
End Function

' ADR-070 (revisao 07/10/2026): linha da aba Inativar TOTALMENTE vazia -- sem ID, analito, motivo, data,
' usuario e com a caixa desmarcada. So ela e reaproveitada pelo RegistrarInativacao. Antes bastavam ID,
' ANALITO e MOTIVO vazios: uma linha com carimbo residual (sobra da migracao, de exclusao em varias areas
' ou de edicao com eventos desligados) era reaproveitada, o Carimbar mantinha a data e o usuario antigos e
' a caixa nao era marcada (o PrepararLinha so marca linha sem data) -- o ponto inativado pelo grafico
' virava NAO_PLOTAR em vez de X vermelho, e a trilha registrava outra data e outro usuario.
Private Function LinhaInatVazia(ByVal lo As ListObject, ByVal r As Long) As Boolean
    Dim c As Variant, v As Variant
    On Error GoTo nao
    For Each c In Array("ID_REGISTRO", "ANALITO", "MOTIVO", "DATA_INATIVACAO", "USUARIO")
        v = Celula(lo, r, CStr(c)).Value
        If IsError(v) Then Exit Function
        If Len(Trim$(CStr(v))) > 0 Then Exit Function
    Next c
    v = Celula(lo, r, COL_LJ).Value
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        If v Then Exit Function
    ElseIf Len(Trim$(CStr(v))) > 0 Then
        Exit Function
    End If
    LinhaInatVazia = True
nao:
End Function

' ============================================================================
'  INATIVAR / REATIVAR PELO GRAFICO -- nucleo SEM dialogo (ADR-081)
' ============================================================================
' A janela do duplo clique (mUI.InativarPontoUI) so pergunta; quem le o ponto, grava a aba Inativar e registra
' no Audit_Log sao estas funcoes. Nenhuma abre dialogo nem atualiza o grafico (quem chama decide) -- o QA as
' chama por automacao com os mesmos argumentos. Devolvem "OK|..." ou "ERRO|...".

' Linha (1..n) do ID normalizado na tblCQ_Final, ou 0. Vale o 1o resultado com o ID: os repetidos da
' DB_CQ_FINAL levam o sufixo ~L<linha> e nao casam.
Private Function LinhaCQ(ByVal lo As ListObject, ByVal idN As String) As Long
    Dim v As Variant, i As Long
    If lo Is Nothing Then Exit Function
    If Len(idN) = 0 Or lo.ListRows.Count = 0 Then Exit Function
    v = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value2
    If Not IsArray(v) Then
        Dim u(1 To 1, 1 To 1) As Variant: u(1, 1) = v: v = u
    End If
    For i = 1 To UBound(v, 1)
        If Not IsError(v(i, 1)) Then
            If UCase$(Trim$(CStr(v(i, 1)))) = idN Then LinhaCQ = i: Exit Function
        End If
    Next i
End Function

' Valor de uma coluna da tblCQ_Final na linha r; Empty se a coluna nao existir ou a celula tiver erro.
Private Function ValorCQ(ByVal lo As ListObject, ByVal r As Long, ByVal coluna As String) As Variant
    Dim v As Variant
    On Error GoTo sai
    v = lo.ListColumns(coluna).DataBodyRange.Cells(r, 1).Value
    If Not IsError(v) Then ValorCQ = v
sai:
End Function

' Texto de uma celula da aba Inativar ("" com erro ou coluna ausente).
Private Function TextoCel(ByVal lo As ListObject, ByVal r As Long, ByVal coluna As String) As String
    Dim v As Variant
    On Error GoTo sai
    v = Celula(lo, r, coluna).Value
    If Not IsError(v) Then TextoCel = Trim$(CStr(v))
sai:
End Function

Private Function EhNum(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    EhNum = IsNumeric(v)
End Function

' ANALITO do resultado na tblCQ_Final ("" se o ID nao existir). E o que o Power Query confere (E10/E11).
Public Function AnalitoDoId(ByVal id As String) As String
    Dim lo As ListObject, rq As Long
    On Error GoTo sai
    Set lo = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(lo, NormalizarId(id))
    If rq > 0 Then AnalitoDoId = Trim$(CStr(ValorCQ(lo, rq, "ANALITO")))
sai:
End Function

' Texto gravado como TEXTO: um motivo que comeca com = + - @ viraria formula (ou erro) na celula.
Private Function TextoParaCelula(ByVal s As String) As String
    If Len(s) > 0 Then
        If InStr(1, "=+-@", Left$(s, 1), vbBinaryCompare) > 0 Then s = "'" & s
    End If
    TextoParaCelula = s
End Function

' Primeira linha da aba Inativar com o ID (0 = nenhuma) e, em livre, a 1a linha TOTALMENTE vazia (0 = cheia).
Private Function LinhaInat(ByVal lo As ListObject, ByVal idN As String, ByRef livre As Long) As Long
    Dim v As Variant, i As Long
    livre = 0
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value
    If Not IsArray(v) Then
        Dim u(1 To 1, 1 To 1) As Variant: u(1, 1) = v: v = u
    End If
    For i = 1 To UBound(v, 1)
        If Not IsError(v(i, 1)) Then
            If LinhaInat = 0 And Len(Trim$(CStr(v(i, 1)))) > 0 Then
                If NormalizarId(v(i, 1)) = idN Then LinhaInat = i
            End If
            If livre = 0 And Len(Trim$(CStr(v(i, 1)))) = 0 Then
                If LinhaInatVazia(lo, i) Then livre = i
            End If
        End If
        If LinhaInat > 0 And livre > 0 Then Exit Function
    Next i
End Function

' Evento da inativacao/reativacao pedida pelo grafico no Audit_Log, com os dados do resultado (a trilha nunca
' derruba a gravacao -- mesma regra do AuditarMudancas). O RESULTADO_INATIVADO/REATIVADO continua vindo da
' atualizacao seguinte (AuditarMudancas), quando a mudanca de fato vale.
Private Sub AuditarPedido(ByVal acao As String, ByVal loQ As ListObject, ByVal rq As Long, ByVal stAnt As String, _
                          ByVal stNovo As String, ByVal motivo As String, ByVal detalhe As String)
    On Error Resume Next
    mAuditoria.Auditar mAuditoria.CAT_DADO, acao, "mIntegracao", CLng(Val(CStr(ValorCQ(loQ, rq, "RUN")))), _
                       ValorCQ(loQ, rq, "DATA_HORA"), CStr(ValorCQ(loQ, rq, "EQUIPAMENTO")), CStr(ValorCQ(loQ, rq, "LOTE")), _
                       CLng(Val(CStr(ValorCQ(loQ, rq, "NIVEL")))), CStr(ValorCQ(loQ, rq, "ANALITO")), _
                       ValorCQ(loQ, rq, "RESULTADO"), ValorCQ(loQ, rq, "RESULTADO"), stAnt, stNovo, motivo, detalhe
End Sub

' O ponto do grafico do Painel: "OK|id|analito|nivel|corrida|dd/mm/yyyy hh:nn|valor|INATIVADO ou ATIVO|lote".
'   nivel = indice do grafico; serie/ponto = o que o evento do grafico passa (ponto = slot da corrida: Eng_Saida
'   linha 2 + ponto). Series "Resultado"/"OK"/"Violacao..." = o ponto do nivel; "Nao conforme (X)" = os X (o
'   k-esimo X do grafico e o slot k da corrida). corrida = corrida do lote do ponto (Eng_Saida BK..); no X, que
'   nao tem numero, "entre k e k+1" (k = parte inteira da posicao no eixo, Eng_Saida BN..).
'   O analito, a data, o valor, o estado e o lote sao os da tblCQ_Final (pelo ID), nao os da tela.
Public Function DadosDoPonto(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim ch As Chart, eng As Worksheet, nm As String, r As Long, k As Long, slot As Long, col As Long
    Dim idv As Variant, cor As Variant, pos As Variant, ehX As Boolean, lo As ListObject, rq As Long, idN As String
    Dim dh As Variant, vl As Variant, txtCor As String, txtDh As String, txtVl As String, st As String
    On Error GoTo falha
    If nivel < 1 Or nivel > mEstatistica.NLV Or ponto < 1 Or ponto > 180 Then
        DadosDoPonto = "ERRO|ponto fora do grafico": Exit Function
    End If
    Set ch = ThisWorkbook.Worksheets("Painel").ChartObjects(nivel).Chart
    If serie < 1 Or serie > ch.SeriesCollection.Count Then DadosDoPonto = "ERRO|serie inexistente": Exit Function
    nm = CStr(ch.SeriesCollection(serie).Name)
    If InStr(1, nm, "conforme", vbTextCompare) > 0 Then
        ehX = True
        For k = 1 To serie
            If InStr(1, CStr(ch.SeriesCollection(k).Name), "conforme", vbTextCompare) > 0 Then slot = slot + 1
        Next k
        If slot < 1 Or slot > 3 Then DadosDoPonto = "ERRO|X fora dos tres da corrida": Exit Function
    ElseIf Not (nm = "Resultado" Or nm = "OK" Or InStr(1, nm, "Viola", vbTextCompare) = 1) Then
        DadosDoPonto = "ERRO|nao e um resultado (linha de limite ou calibracao)": Exit Function
    End If
    Set eng = ThisWorkbook.Worksheets("Eng_Saida")
    If ponto > CLng(Val(CStr(eng.Range("I1").Value))) Then DadosDoPonto = "ERRO|alem das corridas publicadas": Exit Function
    r = 2 + ponto
    If ehX Then
        col = (nivel - 1) * 3 + slot - 1
        idv = eng.Cells(r, mEstatistica.ENG_COL_IDX0 + col).Value
        pos = eng.Cells(r, mEstatistica.ENG_COL_XPOS0 + col).Value
        If EhNum(pos) Then
            k = Int(CDbl(pos))
            If k < 1 Then txtCor = "antes da 1" Else txtCor = "entre " & k & " e " & (k + 1)
        End If
    Else
        idv = eng.Cells(r, mEstatistica.ENG_COL_ID0 + nivel - 1).Value
        cor = eng.Cells(r, mEstatistica.ENG_COL_CORL0 + nivel - 1).Value
        If EhNum(cor) Then txtCor = CStr(CLng(cor))
    End If
    If IsError(idv) Then idv = ""
    idN = NormalizarId(idv)
    If Len(idN) = 0 Then DadosDoPonto = "ERRO|ponto sem ID": Exit Function
    Set lo = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(lo, idN)
    If rq = 0 Then DadosDoPonto = "ERRO|" & TextoIdNaoEncontrado(idN): Exit Function
    dh = ValorCQ(lo, rq, "DATA_HORA")
    If IsDate(dh) Then txtDh = Format$(dh, "dd/mm/yyyy hh:nn")
    vl = ValorCQ(lo, rq, "RESULTADO")
    If EhNum(vl) Then txtVl = CStr(CDbl(vl)) Else txtVl = CStr(vl)
    If UCase$(Trim$(CStr(ValorCQ(lo, rq, "STATUS_ANALITICO")))) = "INATIVADO" Then st = "INATIVADO" Else st = "ATIVO"
    DadosDoPonto = "OK|" & idN & "|" & Trim$(CStr(ValorCQ(lo, rq, "ANALITO"))) & "|" & nivel & "|" & txtCor & "|" & _
                   txtDh & "|" & txtVl & "|" & st & "|" & Trim$(CStr(ValorCQ(lo, rq, "LOTE")))
    Exit Function
falha:
    DadosDoPonto = "ERRO|" & TXT_FALHA & " (detalhe: " & Err.Description & ")"
End Function

' "ID 54565 nao encontrado nos resultados." (ID curto: o numero que aparece na dica e que se digita)
Private Function TextoIdNaoEncontrado(ByVal idN As String) As String
    TextoIdNaoEncontrado = "ID " & mUI.IdCurto(idN) & " não encontrado nos resultados."
End Function

' ADR-082: o ID do resultado do interfaceamento que o manual "id" substitui ("" se nao e manual, nao existe ou nao
' substitui). Le a tblCQ_Final (o que o motor de dados decidiu): manual ATIVO ou INATIVADO com ID_RELACIONADO. Para a
' dica do grafico e a janela de inativar (mUI). Nunca levanta erro.
Public Function SubstituidoPeloManual(ByVal id As String) As String
    Dim lo As ListObject, rq As Long, idN As String, st As String
    On Error GoTo sai
    idN = NormalizarId(id)
    If Left$(idN, 4) <> "MAN_" Then Exit Function
    Set lo = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(lo, idN)
    If rq = 0 Then Exit Function
    If UCase$(Trim$(CStr(ValorCQ(lo, rq, "ORIGEM_RESULTADO")))) <> "MANUAL" Then Exit Function
    st = UCase$(Trim$(CStr(ValorCQ(lo, rq, "STATUS_ANALITICO"))))
    If st <> "ATIVO" And st <> "INATIVADO" Then Exit Function
    SubstituidoPeloManual = Trim$(CStr(ValorCQ(lo, rq, "ID_RELACIONADO")))
sai:
End Function

' ADR-082: o manual que substitui o resultado do interfaceamento "id" ("" se nao e do interfaceamento, nao existe ou
' nenhum manual o substitui). Le a tblCQ_Final: ID_RELACIONADO = MAN_... num resultado do interfaceamento que nao e
' retransmissao. Para a janela de reativar (mUI): reativado, ele continua fora (SUBSTITUIDO_MANUAL). Nunca levanta erro.
Public Function ManualQueSubstitui(ByVal id As String) As String
    Dim lo As ListObject, rq As Long, idN As String, rel As String
    On Error GoTo sai
    idN = NormalizarId(id)
    If Len(idN) = 0 Or Left$(idN, 4) = "MAN_" Then Exit Function
    Set lo = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(lo, idN)
    If rq = 0 Then Exit Function
    If UCase$(Trim$(CStr(ValorCQ(lo, rq, "ORIGEM_RESULTADO")))) <> "INTERFACEAMENTO" Then Exit Function
    If UCase$(Trim$(CStr(ValorCQ(lo, rq, "STATUS_ANALITICO")))) = "DUPLICIDADE_ORIGEM" Then Exit Function
    rel = Trim$(CStr(ValorCQ(lo, rq, "ID_RELACIONADO")))
    If Left$(rel, 4) = "MAN_" Then ManualQueSubstitui = rel
sai:
End Function

' ADR-082: digitou o SUBSTITUI_ID na linha r da aba Digitar Resultados -> preenche SO o que estiver vazio (EQUIPAMENTO,
' MATRIZ, LOTE, NIVEL, ANALITO, DATA e HORA) com os dados do resultado indicado, da Principal - Resultados. Devolve o
' aviso para quem digita ("" = preenchido): ID inexistente ou que nao e do interfaceamento; ou, sem sobrescrever nada,
' os campos da chave que ja estavam preenchidos com outro valor (o padrao do CFG que o PrepararLinha pos no 1o toque,
' outro nivel...) -- sem o aviso, quem digita da esquerda para a direita so veria o E12 "tem outra chave" no ATUALIZAR.
Private Function PreencherDoSubstituido(ByVal lo As ListObject, ByVal r As Long, ByVal idN As String) As String
    Dim loQ As ListObject, rq As Long, dh As Variant, c As Variant, v As Variant, s As Long
    Dim jaTem As String, alvoTxt As String, difs As String
    Set loQ = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(loQ, idN)
    If rq = 0 Then
        PreencherDoSubstituido = TextoIdNaoEncontrado(idN) & " O SUBSTITUI_ID aceita só o ID de um resultado do " & _
            "interfaceamento (o número da dica do gráfico); se o resultado acabou de chegar, clique em ATUALIZAR DADOS " & _
            "antes. Sem um ID válido, nada é substituído."
        Exit Function
    End If
    If UCase$(Trim$(CStr(ValorCQ(loQ, rq, "ORIGEM_RESULTADO")))) <> "INTERFACEAMENTO" Then
        PreencherDoSubstituido = "O ID " & mUI.IdCurto(idN) & " é de um resultado manual: o SUBSTITUI_ID aceita só o ID " & _
            "de um resultado do interfaceamento. Sem um ID válido, nada é substituído."
        Exit Function
    End If
    For Each c In Array("EQUIPAMENTO", "MATRIZ", "LOTE", "NIVEL", "ANALITO")
        v = ValorCQ(loQ, rq, CStr(c))
        jaTem = TextoCel(lo, r, CStr(c))
        If Len(jaTem) = 0 Then
            If Not IsEmpty(v) Then Celula(lo, r, CStr(c)).Value = v
        Else
            alvoTxt = Trim$(CStr(v))
            If jaTem <> alvoTxt Then difs = difs & vbLf & "  " & CStr(c) & ": na linha '" & jaTem & "', no resultado '" & alvoTxt & "'"
        End If
    Next c
    If Len(difs) > 0 Then
        PreencherDoSubstituido = "Estes campos já estavam preenchidos e são diferentes do resultado " & mUI.IdCurto(idN) & _
            ":" & difs & vbLf & vbLf & "Nada foi sobrescrito. Para o manual substituir o resultado, corrija-os (com a " & _
            "chave diferente, o QA acusa E12 ""tem outra chave"" e nada é substituído)."
    End If
    dh = ValorCQ(loQ, rq, "DATA_HORA")
    If IsDate(dh) Then
        ' a data e a hora do proprio resultado (ao segundo): o manual entra na mesma corrida
        If Len(TextoCel(lo, r, "DATA")) = 0 Then Celula(lo, r, "DATA").Value = CDate(Int(CDbl(dh)))
        If Len(TextoCel(lo, r, "HORA")) = 0 Then
            s = CLng(Int((CDbl(dh) - Int(CDbl(dh))) * 86400# + 0.5))
            If s > 86399 Then s = 86399
            Celula(lo, r, "HORA").Value = CDbl(s) / 86400#
        End If
    Else
        ' sem DATA_HORA no interfaceamento (SEM_DATA_HORA): so a data, se houver -- a hora fica para quem digita
        v = ValorCQ(loQ, rq, "DATA")
        If IsDate(v) And Len(TextoCel(lo, r, "DATA")) = 0 Then Celula(lo, r, "DATA").Value = CDate(Int(CDbl(v)))
    End If
End Function

' Grava a inativacao na aba Inativar: ID normalizado, ANALITO DA tblCQ_Final (nunca o da tela -- a conferencia
' E10 do Power Query nao recusa o que o grafico mandou), REGISTRAR - LJ = plotarX, MOTIVO, data e usuario
' carimbados como no evento. Motivo obrigatorio. ID ja na aba com o analito certo = recusa ("ja inativado"); com
' o analito errado ou vazio (inativacao que nao valia: E10/E11) a linha e CORRIGIDA no lugar, com carimbo novo.
' Registra no Audit_Log (INATIVACAO_PEDIDA ou INATIVACAO_CORRIGIDA) com a corrida do lote de antes no detalhe.
' NAO atualiza os dados. "OK|<linha>" ou "OK|<linha>|corrigida"; "ERRO|<motivo>".
Public Function InativarPorId(ByVal id As String, ByVal motivo As String, ByVal plotarX As Boolean, _
                              Optional ByVal corridaAntes As String = "") As String
    Dim loQ As ListObject, lo As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, idN As String
    Dim rq As Long, an As String, r As Long, livre As Long, anAntigo As String, corrigida As Boolean, sE As String
    On Error GoTo falha
    idN = NormalizarId(id)
    motivo = Trim$(motivo)
    If Len(idN) = 0 Then InativarPorId = "ERRO|Informe o ID do resultado.": Exit Function
    If Len(motivo) = 0 Then InativarPorId = "ERRO|Escolha o motivo da inativação.": Exit Function
    Set loQ = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(loQ, idN)
    If rq = 0 Then InativarPorId = "ERRO|" & TextoIdNaoEncontrado(idN): Exit Function
    an = Trim$(CStr(ValorCQ(loQ, rq, "ANALITO")))
    If Len(an) = 0 Then InativarPorId = "ERRO|O resultado ID " & mUI.IdCurto(idN) & " está sem analito nos resultados. " & _
                                        "Nada foi alterado.": Exit Function
    Set lo = AcharTabela(TB_INAT)
    If lo Is Nothing Then InativarPorId = "ERRO|A aba Inativar não foi encontrada. Nada foi alterado.": Exit Function
    Set ws = lo.Parent
    r = LinhaInat(lo, idN, livre)
    If r > 0 Then
        anAntigo = TextoCel(lo, r, "ANALITO")
        If UCase$(anAntigo) = UCase$(an) Then
            InativarPorId = "ERRO|Este resultado já está inativado (aba Inativar, linha " & r & ").": Exit Function
        End If
        corrigida = True
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    If Not corrigida Then
        If livre = 0 Then                      ' tabela cheia: cresce FOLGA linhas (a aba protegida nao deixa)
            lo.Resize lo.Range.Resize(lo.Range.rows.Count + FOLGA)
            livre = lo.ListRows.Count - FOLGA + 1
        End If
        r = livre
        Celula(lo, r, "ID_REGISTRO").Value = idN
    Else
        ' a linha antiga nao valia (analito errado ou vazio): passa a valer agora -- carimbo de agora
        Celula(lo, r, "DATA_INATIVACAO").ClearContents
        Celula(lo, r, "USUARIO").ClearContents
    End If
    Celula(lo, r, "ANALITO").Value = an
    Celula(lo, r, "MOTIVO").Value = TextoParaCelula(motivo)
    PrepararLinha lo, r                        ' ID normalizado, caixa marcada (linha sem carimbo) e carimbo
    Celula(lo, r, COL_LJ).Value = plotarX      ' a escolha da janela vale sobre o padrao
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    AuditarPedido IIf(corrigida, "INATIVACAO_CORRIGIDA", "INATIVACAO_PEDIDA"), loQ, rq, "ATIVO", _
                  "INATIVADO (aba Inativar)", motivo, "ID " & idN & " | corrida do lote antes: " & _
                  IIf(Len(corridaAntes) > 0, corridaAntes, "?") & " | X vermelho no grafico: " & IIf(plotarX, "SIM", "NÃO") & _
                  " | aba Inativar linha " & r & IIf(corrigida, " | linha corrigida: ANALITO era '" & anAntigo & "'", "")
    InativarPorId = "OK|" & r & IIf(corrigida, "|corrigida", "")
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    InativarPorId = "ERRO|" & TXT_FALHA & " (detalhe: " & sE & ")"
End Function

' Reativa: apaga a(s) linha(s) do ID na aba Inativar (ADR-070: apagar a linha = reativar; com o ID repetido, todas
' -- senao a 2a linha passaria a valer). Motivo obrigatorio, registrado no Audit_Log (REATIVACAO_PEDIDA) com o
' motivo da inativacao que sai. NAO atualiza os dados. "OK|<linhas apagadas>" ou "ERRO|<motivo>".
Public Function ReativarPorId(ByVal id As String, ByVal motivo As String) As String
    Dim lo As ListObject, loQ As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, idN As String
    Dim v As Variant, i As Long, quais As String, motAnt As String, c As Variant, rq As Long, sE As String
    Dim achou As Collection, it As Variant
    On Error GoTo falha
    idN = NormalizarId(id)
    motivo = Trim$(motivo)
    If Len(idN) = 0 Then ReativarPorId = "ERRO|Informe o ID do resultado.": Exit Function
    If Len(motivo) = 0 Then ReativarPorId = "ERRO|Escolha o motivo da reativação.": Exit Function
    Set lo = AcharTabela(TB_INAT)
    If lo Is Nothing Then ReativarPorId = "ERRO|A aba Inativar não foi encontrada. Nada foi alterado.": Exit Function
    Set ws = lo.Parent
    Set achou = New Collection
    If lo.ListRows.Count > 0 Then
        v = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value
        If Not IsArray(v) Then
            Dim u(1 To 1, 1 To 1) As Variant: u(1, 1) = v: v = u
        End If
        For i = 1 To UBound(v, 1)
            If Not IsError(v(i, 1)) Then
                If Len(Trim$(CStr(v(i, 1)))) > 0 Then
                    If NormalizarId(v(i, 1)) = idN Then achou.Add i
                End If
            End If
        Next i
    End If
    If achou.Count = 0 Then
        ReativarPorId = "ERRO|O ID " & mUI.IdCurto(idN) & " não está na aba Inativar (não está inativado)."
        Exit Function
    End If
    motAnt = TextoCel(lo, CLng(achou(1)), "MOTIVO")
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = LiberarEscrita(ws)
    For Each it In achou
        For Each c In Array("ID_REGISTRO", "ANALITO", COL_LJ, "MOTIVO", "DATA_INATIVACAO", "USUARIO")
            LimparSeExistir lo, CLng(it), CStr(c)
        Next c
        quais = quais & IIf(Len(quais) > 0, ",", "") & it
    Next it
    RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    Set loQ = AcharTabela(FONTE_CQ)
    rq = LinhaCQ(loQ, idN)
    If rq > 0 Then
        AuditarPedido "REATIVACAO_PEDIDA", loQ, rq, "INATIVADO", "ATIVO (aba Inativar)", motivo, _
                      "ID " & idN & " | aba Inativar linha(s) " & quais & " apagada(s) | motivo da inativacao: " & _
                      IIf(Len(motAnt) > 0, motAnt, "(vazio)")
    Else
        On Error Resume Next
        mAuditoria.Auditar mAuditoria.CAT_DADO, "REATIVACAO_PEDIDA", "mIntegracao", 0, Empty, "", "", 0, "", _
                           Empty, Empty, "INATIVADO", "ATIVO (aba Inativar)", motivo, "ID " & idN & " | aba Inativar linha(s) " & quais
        On Error GoTo falha
    End If
    ReativarPorId = "OK|" & quais
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    ReativarPorId = "ERRO|" & TXT_FALHA & " (detalhe: " & sE & ")"
End Function

Private Sub Carimbar(ByVal cData As Range, ByVal cUsuario As Range)
    If IsEmpty(cData.Value) Or Len(Trim$(CStr(cData.Value))) = 0 Then cData.Value = Now
    If Len(Trim$(CStr(cUsuario.Value))) = 0 Then cUsuario.Value = mAuditoria.UsuarioSistema()
End Sub

Private Function EquipPadrao() As String
    Dim eq As String
    On Error Resume Next
    eq = Trim$(CStr(ThisWorkbook.Names("selEquipamento").RefersToRange.Value))
    On Error GoTo 0
    If Len(eq) = 0 Then eq = Cfg("EQUIPAMENTO_PADRAO")
    EquipPadrao = eq
End Function

' Proximo MAN_nnnn: o maior numero ja usado + 1 (nunca reaproveita um ID).
Public Function ProximoIdManual(ByVal lo As ListObject) As String
    ' D05 (QA-ETL-001): NUNCA reaproveita um MAN_ ja usado. Antes era "maior MAN_ presente + 1": apagar a ultima
    ' linha manual devolvia o numero, e o manual novo herdava a inativacao e o comentario do anterior (nascia
    ' INATIVADO). O maximo agora vem da tabela manual, das tabelas que citam IDs (inativacao, comentarios), da
    ' tblCQ_Final e de uma marca persistente no arquivo (nome oculto qcUltimoIdManual).
    Dim mx As Long, n As Long, nome As Variant, l As ListObject
    mx = MaiorIdManual(lo)
    For Each nome In Array("tblInativacao_NaoConformes", "tblComentariosTecnicos", "tblCQ_Final")
        Set l = Nothing
        On Error Resume Next
        Set l = AcharTabela(CStr(nome))
        On Error GoTo 0
        If Not l Is Nothing Then
            n = MaiorIdManual(l)
            If n > mx Then mx = n
        End If
    Next nome
    n = MarcaIdManual()
    If n > mx Then mx = n
    ProximoIdManual = "MAN_" & Format$(mx + 1, "0000")
    GravarMarcaIdManual mx + 1
End Function

' Maior numero n de "MAN_n" (com ou sem o sufixo ~L<linha> da DB_CQ_FINAL) na coluna ID_REGISTRO da tabela.
Private Function MaiorIdManual(ByVal lo As ListObject) As Long
    Dim v As Variant, i As Long, t As String, n As Long, p As Long
    On Error GoTo sai
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.ListColumns("ID_REGISTRO").DataBodyRange.Value
    If Not IsArray(v) Then
        Dim u(1 To 1, 1 To 1) As Variant: u(1, 1) = v: v = u
    End If
    For i = 1 To UBound(v, 1)
        t = UCase$(Trim$(CStr(v(i, 1))))
        p = InStr(t, "~")
        If p > 0 Then t = Left$(t, p - 1)
        If Left$(t, 4) = "MAN_" Then
            If SoDigitos(Mid$(t, 5)) And Len(t) <= 13 Then
                n = CLng(Mid$(t, 5))
                If n > MaiorIdManual Then MaiorIdManual = n
            End If
        End If
    Next i
sai:
End Function

Private Function MarcaIdManual() As Long
    Dim r As String
    On Error Resume Next
    r = ThisWorkbook.Names("qcUltimoIdManual").RefersTo
    If Len(r) > 1 Then MarcaIdManual = CLng(Val(Mid$(r, 2)))
End Function

Private Sub GravarMarcaIdManual(ByVal n As Long)
    On Error Resume Next
    If n <= MarcaIdManual() Then Exit Sub
    ThisWorkbook.Names.Add Name:="qcUltimoIdManual", RefersTo:="=" & CStr(n), Visible:=False
End Sub

' Chamado pelo Worksheet_Change das abas de entrada.
Public Sub EntradaMudou(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject, cel As Range, ar As Range, r As Long, prot As Boolean, an As String
    Dim idDig As String, celAviso As Range, txtAviso As String, tituloAviso As String, aviso As String
    If mApp.Ocupado() Then Exit Sub
    Set lo = Nothing
    On Error Resume Next
    Select Case ws.Name
        Case ABA_MANUAL: Set lo = ws.ListObjects(TB_MANUAL)
        Case ABA_INAT: Set lo = ws.ListObjects(TB_INAT)
        Case ABA_COMENT: Set lo = ws.ListObjects(TB_COMENT)
    End Select
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    ' ADR-081: a aba Inativar e um caminho de inativar/reativar, como a janela do grafico -- sem permissao
    ' (mSeguranca.PodeInativar: login e papel minimo PAPEL_MIN_INATIVAR) a alteracao e desfeita. So com gente
    ' digitando (Interactive): a automacao (QA, instaladores) escreve na aba sem sessao de login.
    If ws.Name = ABA_INAT And Application.Interactive Then
        If Not mSeguranca.PodeInativar() Then
            Application.EnableEvents = False
            On Error Resume Next
            Application.Undo
            On Error GoTo 0
            Application.EnableEvents = True
            MsgBox mSeguranca.TextoSemPermissaoInativar() & vbCrLf & vbCrLf & _
                   "A alteração na aba Inativar foi desfeita.", vbExclamation, "Inativar"
            Exit Sub
        End If
    End If
    Application.EnableEvents = False
    On Error GoTo fim
    prot = LiberarEscrita(ws)
    ' ADR-070 (revisao): percorre TODAS as areas da selecao -- .Rows de um intervalo com varias areas
    ' (Ctrl+selecao) so cobre a primeira, e as outras linhas apagadas ficavam com o carimbo residual
    For Each ar In Intersect(Target, lo.DataBodyRange).Areas
    For Each cel In ar.Rows
        r = cel.Row - lo.DataBodyRange.Row + 1
        If ws.Name = ABA_INAT Then
            ' apagou o ID = reativacao: limpa o carimbo e a caixa (volta ao padrao)
            If Len(Trim$(CStr(Celula(lo, r, "ID_REGISTRO").Value))) = 0 Then
                Celula(lo, r, COL_LJ).ClearContents
                Celula(lo, r, "DATA_INATIVACAO").ClearContents
                Celula(lo, r, "USUARIO").ClearContents
                ' ADR-070: foi a celula do ID que o usuario apagou -> a linha inteira sai (o analito e o motivo
                ' nao ficam esperando o proximo ID digitado ali). Digitar o analito ANTES do ID nao apaga nada.
                If Not Intersect(Target, Celula(lo, r, "ID_REGISTRO")) Is Nothing Then
                    LimparSeExistir lo, r, "ANALITO"
                    LimparSeExistir lo, r, "MOTIVO"
                End If
            ElseIf Not Intersect(Target, Celula(lo, r, "ID_REGISTRO")) Is Nothing Then
                ' ADR-081: digitou o ID e o ANALITO esta vazio -> o analito DAQUELE resultado, da Principal -
                ' Resultados (quem digita so precisa do numero). So no evento e so quando a celula do ID mudou:
                ' a linha que entra sem analito por outro caminho (PrepararEntradas) continua E11, como no ADR-070
                idDig = NormalizarId(TextoCel(lo, r, "ID_REGISTRO"))
                If LinhaCQ(AcharTabela(FONTE_CQ), idDig) = 0 Then
                    ' ID que nao existe: avisa quem digita (depois do laco), sem apagar o que foi digitado. Manual
                    ' (MAN_) ainda nao atualizado nao esta nos resultados e nao e erro
                    If Application.Interactive And celAviso Is Nothing And Left$(idDig, 4) <> "MAN_" Then
                        Set celAviso = Celula(lo, r, "ID_REGISTRO")
                        txtAviso = TextoIdNaoEncontrado(idDig) & " Confira o número no gráfico (passe o mouse no ponto)."
                        tituloAviso = "Inativar"
                    End If
                ElseIf Len(TextoCel(lo, r, "ANALITO")) = 0 Then
                    an = AnalitoDoId(idDig)
                    If Len(an) > 0 Then Celula(lo, r, "ANALITO").Value = an
                End If
            End If
        ElseIf ws.Name = ABA_MANUAL Then
            ' ADR-082: digitou o SUBSTITUI_ID -> a linha se preenche com os dados do resultado indicado (so o vazio)
            If TemColuna(lo, COL_SUBST) Then
                If Not Intersect(Target, Celula(lo, r, COL_SUBST)) Is Nothing Then
                    idDig = NormalizarId(TextoCel(lo, r, COL_SUBST))
                    If Len(idDig) > 0 Then
                        aviso = PreencherDoSubstituido(lo, r, idDig)
                        If Len(aviso) > 0 And Application.Interactive And celAviso Is Nothing Then
                            Set celAviso = Celula(lo, r, COL_SUBST)
                            txtAviso = aviso
                            tituloAviso = "Digitar Resultados"
                        End If
                    End If
                End If
            End If
        End If
        PrepararLinha lo, r
    Next cel
    Next ar
    ' tabela quase cheia: acrescenta FOLGA linhas (a aba protegida nao deixa a
    ' tabela crescer sozinha)
    If lo.ListRows.Count - (Intersect(Target, lo.DataBodyRange).Row - lo.DataBodyRange.Row + 1) < 10 Then
        lo.Resize lo.Range.Resize(lo.Range.rows.Count + FOLGA)
    End If
fim:
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
    If Not celAviso Is Nothing Then
        On Error Resume Next
        MsgBox txtAviso, vbExclamation, tituloAviso
        celAviso.Select                        ' o ID digitado fica selecionado para corrigir
    End If
End Sub

' ============================================================================
'  NAVEGACAO
' ============================================================================
' Um botao por aba, com o MESMO nome da aba (pedido do usuario):
'   DB ...................... dados recebidos do DB_SEAC
'   Inativar ................ retirar resultados da populacao principal
'   Digitar Resultados ...... lancamento manual (MAN_0001, MAN_0002...)
'   Principal - Resultados .. a tabela final (DB_CQ_FINAL), fonte do LJ e da Estatistica
Public Sub IrDB(): mApp.Ir "DB", "A6": End Sub
Public Sub IrInativar(): IrLinhaLivre ABA_INAT, TB_INAT, "ID_REGISTRO": End Sub
Public Sub IrDigitarResultados(): IrLinhaLivre ABA_MANUAL, TB_MANUAL, "DATA": End Sub
Public Sub IrPrincipal(): mApp.Ir ABA_CQ, "A6": End Sub
Public Sub IrComentarios(): IrLinhaLivre ABA_COMENT, TB_COMENT, "ID_REGISTRO": End Sub
Public Sub IrQA(): mApp.Ir "QA_INTEGRACAO", "A6": End Sub
Public Sub IrOrganizado(): mApp.Ir "DB_ORGANIZADO", "A6": End Sub

' Abre a aba de entrada ja na primeira linha vazia da tabela.
Private Sub IrLinhaLivre(ByVal aba As String, ByVal tabela As String, ByVal coluna As String)
    Dim lo As ListObject, v As Variant, i As Long, alvo As String
    alvo = "A7"
    On Error Resume Next
    Set lo = ThisWorkbook.Sheets(aba).ListObjects(tabela)
    If Not lo Is Nothing Then
        v = lo.ListColumns(coluna).DataBodyRange.Value
        For i = 1 To UBound(v, 1)
            If Len(Trim$(CStr(v(i, 1)))) = 0 Then
                alvo = lo.ListColumns(coluna).DataBodyRange.Cells(i, 1).Address(False, False)
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0
    mApp.Ir aba, alvo
End Sub
