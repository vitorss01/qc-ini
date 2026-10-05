Attribute VB_Name = "mIntegracao"
Option Explicit
' ============================================================================
'  ORQUESTRACAO DA CAMADA DE DADOS (ADR-057)
'
'  O Power Query e o motor de dados: recebe o DB_SEAC, junta os manuais,
'  aplica a inativacao e produz a tblCQ_Final. Este modulo NAO decide nada
'  sobre dado -- so:
'    1. prepara as tabelas de entrada (ID normalizado, carimbo de data/usuario,
'       caixa REGISTRAR - LJ marcada por padrao, ID MAN_nnnn automatico);
'    2. dispara o refresh das consultas NA ORDEM, de forma SINCRONA
'       (BackgroundQuery = False): cada Refresh so devolve quando a consulta
'       terminou -- ou levanta o erro dela. Nenhuma etapa seguinte roda sobre
'       tabela pela metade, e falha nunca vira "Atualizacao concluida";
'    3. so entao atualiza os consumidores (motor, Westgard, Estatistica, LJ);
'    4. registra no Audit_Log o que mudou de estado (inativado, reativado,
'       plotagem alterada, resultado manual incluido/retirado).
' ============================================================================
Private Const ABA_MANUAL As String = "Digitar Resultados"
Private Const ABA_INAT As String = "Inativar"
Private Const ABA_COMENT As String = "COMENTARIOS_TECNICOS"
Private Const TB_MANUAL As String = "tblResultados_Manuais"
Private Const TB_INAT As String = "tblInativacao_NaoConformes"
Private Const TB_COMENT As String = "tblComentariosTecnicos"
Private Const COL_LJ As String = "REGISTRAR - LJ"
Private Const FOLGA As Long = 100              ' linhas acrescentadas quando a tabela de entrada enche

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
Public Function AtualizarDadosCore() As String
    Dim camada As Variant, t0 As Double, tempos As String
    Dim antes As Object, depois As Object, nRec0 As Long, nRec1 As Long, etapa As String, hist As Boolean
    Dim resLotes As String
    mApp.InicioUsuario "Atualizando dados: preparando as entradas..."
    On Error GoTo falha
    hist = ModoHistorico()

    etapa = "preparar as tabelas de entrada"
    PrepararEntradas
    nRec0 = Linhas("tblDB_Recebimento")
    Set antes = EstadoAuditavel()

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
    mEstatistica.InvalidarCache
    RodarSeExistir "InvalidarCacheEstat"           ' Estatistica por periodo (so Bioquimica)
    mDados.AtualizarListasAno
    mLotes.AtualizarListaLiberacao
    mEstatistica.AtualizarEstatistica
    RodarSeExistir "RecalcularEstatPeriodo"
    mIncerteza.RecalcularIncerteza                 ' ADR-064: incerteza de medicao (CIQ + CEQ)
    tempos = tempos & " · motor " & Format$(Timer - t0, "0.0") & "s"

    If hist Then mAuditoria.RegistrarLog "ATUALIZACAO_MODO_HISTORICO", _
        "MODO_FONTE = HISTORICO: DB_SEAC nao lido; reprocessado o historico ja recebido (" & nRec1 & " resultados)"

    mApp.fim
    AtualizarDadosCore = IIf(hist, AvisoHistorico() & vbCrLf & vbCrLf, "") & _
           "Atualização concluída." & vbCrLf & vbCrLf & _
           "Resultados recebidos: " & Format$(nRec1, "#,##0") & " (novos nesta atualização: " & (nRec1 - nRec0) & ")" & vbCrLf & _
           "Principal - Resultados: " & Format$(Linhas(FONTE_CQ), "#,##0") & " resultados" & vbCrLf & _
           ResumoQA() & ResumoLotes(resLotes) & vbCrLf & vbCrLf & "Tempo: " & tempos
    Exit Function
falha:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    On Error Resume Next
    mAuditoria.RegistrarLog "ATUALIZACAO_FALHOU", "Etapa: " & etapa & " | Erro " & nE & ": " & sE
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
Public Sub Refrescar(ByVal nomeTabela As String)
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, nE As Long, sE As String
    Set lo = AcharTabela(nomeTabela)
    If lo Is Nothing Then Err.Raise vbObjectError + 570, "Refrescar", "tabela " & nomeTabela & " nao encontrada"
    Set ws = lo.Parent
    prot = LiberarEscrita(ws)
    On Error GoTo restaura
    With lo.QueryTable
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
        If .Refreshing Then Err.Raise vbObjectError + 572, "Refrescar", nomeTabela & " ainda em atualizacao"
    End With
restaura:
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao ws, prot
    If nE <> 0 Then Err.Raise nE, "mIntegracao.Refrescar(" & nomeTabela & ")", sE
End Sub

Public Function AcharTabela(ByVal nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nome Then Set AcharTabela = lo: Exit Function
        Next lo
    Next ws
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
    If e > 0 Then ResumoQA = ResumoQA & vbCrLf & "Há ERRO DE GOVERNANÇA: veja a aba QA_INTEGRACAO (ex.: inativação sem justificativa técnica)."
End Function

' ============================================================================
'  AUDITORIA (ISO 15189 8.4) -- o que mudou de estado entre duas atualizacoes
' ============================================================================
' ID -> Array(estado, analito, lote, nivel, run, data, equip, resultado, comentario)
' So o que interessa a trilha: resultados INATIVADOS e resultados MANUAIS.
Private Function EstadoAuditavel() As Object
    Dim d As Object, lo As ListObject, n As Long, i As Long
    Dim vId, vSt, vPl, vOr, vAn, vLt, vNv, vRun, vDt, vEq, vRes, vCom, est As String
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
    For i = 1 To n
        est = ""
        If CStr(vSt(i, 1)) = "INATIVADO" Then est = "INATIVADO/" & CStr(vPl(i, 1))
        If CStr(vOr(i, 1)) = "MANUAL" Then est = est & "|MANUAL:" & CStr(vSt(i, 1))
        If Len(est) > 0 Then
            d(CStr(vId(i, 1))) = Array(est, vAn(i, 1), vLt(i, 1), vNv(i, 1), vRun(i, 1), vDt(i, 1), vEq(i, 1), vRes(i, 1), vCom(i, 1))
        End If
    Next i
End Function

Private Sub AuditarMudancas(ByVal antes As Object, ByVal depois As Object)
    Dim k As Variant, a As String, b As String, x As Variant
    On Error Resume Next                       ' a trilha nunca derruba a atualizacao
    For Each k In depois.Keys
        b = depois(k)(0)
        If antes.Exists(k) Then a = antes(k)(0) Else a = ""
        If a <> b Then Registrar CStr(k), a, b, depois(k)
    Next k
    For Each k In antes.Keys
        If Not depois.Exists(k) Then Registrar CStr(k), antes(k)(0), "", antes(k)
    Next k
End Sub

Private Sub Registrar(ByVal id As String, ByVal a As String, ByVal b As String, ByVal x As Variant)
    Dim acao As String, inA As Boolean, inB As Boolean, manA As Boolean, manB As Boolean
    inA = Left$(a, 9) = "INATIVADO": inB = Left$(b, 9) = "INATIVADO"
    manA = InStr(a, "MANUAL:") > 0: manB = InStr(b, "MANUAL:") > 0
    If inB And Not inA Then
        acao = "RESULTADO_INATIVADO"
    ElseIf inA And Not inB Then
        acao = "RESULTADO_REATIVADO"
    ElseIf inA And inB Then
        acao = "PLOTAGEM_LJ_ALTERADA"
    End If
    If manB And Not manA Then acao = acao & IIf(Len(acao) > 0, "+", "") & "RESULTADO_MANUAL_INCLUIDO"
    If manA And Not manB Then acao = acao & IIf(Len(acao) > 0, "+", "") & "RESULTADO_MANUAL_RETIRADO"
    If manA And manB And Not (inA Or inB) Then acao = "RESULTADO_MANUAL_STATUS"
    If Len(acao) = 0 Then Exit Sub
    mAuditoria.Auditar mAuditoria.CAT_DADO, acao, "mIntegracao", CLng(Val(CStr(x(4)))), x(5), CStr(x(6)), _
                       CStr(x(2)), CLng(Val(CStr(x(3)))), CStr(x(1)), x(7), x(7), a, b, _
                       "ID " & id, CStr(x(8))
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
            Carimbar Celula(lo, r, "REGISTRADO_EM"), Celula(lo, r, "USUARIO")
    End Select
End Sub

Private Function Celula(ByVal lo As ListObject, ByVal r As Long, ByVal coluna As String) As Range
    Set Celula = lo.ListColumns(coluna).DataBodyRange.Cells(r, 1)
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
    Dim lo As ListObject, cel As Range, r As Long, prot As Boolean
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
    Application.EnableEvents = False
    On Error GoTo fim
    prot = LiberarEscrita(ws)
    For Each cel In Intersect(Target, lo.DataBodyRange).Rows
        r = cel.Row - lo.DataBodyRange.Row + 1
        If ws.Name = ABA_INAT Then
            ' apagou o ID = reativacao: limpa o carimbo e a caixa (volta ao padrao)
            If Len(Trim$(CStr(Celula(lo, r, "ID_REGISTRO").Value))) = 0 Then
                Celula(lo, r, COL_LJ).ClearContents
                Celula(lo, r, "DATA_INATIVACAO").ClearContents
                Celula(lo, r, "USUARIO").ClearContents
            End If
        End If
        PrepararLinha lo, r
    Next cel
    ' tabela quase cheia: acrescenta FOLGA linhas (a aba protegida nao deixa a
    ' tabela crescer sozinha)
    If lo.ListRows.Count - (Intersect(Target, lo.DataBodyRange).Row - lo.DataBodyRange.Row + 1) < 10 Then
        lo.Resize lo.Range.Resize(lo.Range.rows.Count + FOLGA)
    End If
fim:
    RestaurarProtecao ws, prot
    Application.EnableEvents = True
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
