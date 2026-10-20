Attribute VB_Name = "mEtlSync"
' ============================================================================
'  mEtlSync -- a atualizacao da camada de dados em VBA (ADR-079)
'
'  Substitui o refresh das consultas do Power Query no botao ATUALIZAR quando Cfg_Integracao MOTOR_DADOS = VBA.
'  As consultas continuam no arquivo, intactas: MOTOR_DADOS = PQ volta ao caminho antigo, e um refresh delas
'  confere (reconcilia) o que o VBA gravou -- os dois motores dao a mesma tabela (estudos/equivalencia_*.py).
'
'  Sequencia (nada e publicado antes de TUDO estar calculado e conferido -- staging):
'    1. origem       le a tabela do setor no DB_SEAC (copia local, driver ACE) -- MODO_FONTE = HISTORICO pula;
'    2. recebimento  so os IDs ainda nao recebidos (chave estavel ID_REGISTRO; idempotente: a 2a execucao nao
'                    acrescenta nada); linha recebida nunca e alterada;
'    3. cq           a Principal - Resultados inteira, recalculada em memoria sobre o recebimento NOVO;
'    4. validar      colunas, IDs unicos, nenhuma linha recebida perdida, Principal nao encolhe;
'    5. publicar     grava SO da primeira linha diferente em diante (incremental) -- sem diferenca, nao grava;
'    6. lotes, qa    lote novo no cadastro (ADR-060) e depois a conferencia (QA_INTEGRACAO);
'    7. Sync_Log     uma linha por execucao: tempos por etapa, contagens, datas, erro, versao. Sem credencial e
'                    sem dado clinico (nada de resultado, so contagem e data).
'  Falha em 1-4: nada e gravado. Falha em 5-6: o que ja foi gravado e consistente (recebimento antes da
'  Principal, como as camadas do Power Query) e a proxima atualizacao completa o resto.
'  A tblDB_Organizado (so visualizacao, sem consumidor) nao e mais refeita a cada ATUALIZAR: AtualizarOrganizado.
' ============================================================================
Option Explicit
Option Compare Binary

Public Const VERSAO_ETL As String = "ADR-079/1"
Private Const TB_REC As String = "tblDB_Recebimento"
Private Const TB_QA As String = "tblQA_Integracao"
Private Const TB_LOG As String = "tblSyncLog"
Private Const LOG_MAX As Long = 2000              ' linhas guardadas no Sync_Log (as mais antigas saem)
Private Const ENCOLHE_MAX As Double = 0.9         ' a Principal nao pode cair abaixo de 90% do que era

Private mTempos As String, mT0 As Double, mEtapa As String
Private mLog As Object                            ' campo -> valor da linha do Sync_Log desta execucao

' ---------------------------------------------------------------- configuracao
' MOTOR_DADOS = VBA liga este caminho; PQ (ou ausente) mantem o refresh das consultas.
Public Function MotorVBA() As Boolean
    MotorVBA = (UCase$(Trim$(mIntegracao.Cfg("MOTOR_DADOS"))) = "VBA")
End Function

Public Function EtapaAtual() As String
    EtapaAtual = mEtapa
End Function

Private Sub Marca(ByVal fase As String)
    mTempos = mTempos & IIf(Len(mTempos) > 0, " · ", "") & fase & " " & Format$(Timer - mT0, "0.0") & "s"
    mLog("T_" & UCase$(fase)) = Round(Timer - mT0, 2)
    mT0 = Timer
End Sub

' ---------------------------------------------------------------- a atualizacao
' Devolve os tempos por etapa; nNovos = resultados recebidos agora; mudou = alguma tabela foi regravada;
' resLotes = retorno do mLotes.RegistrarLotesRecebidos. Em falha levanta o erro (EtapaAtual diz onde).
Public Function Sincronizar(ByVal historico As Boolean, ByRef nNovos As Long, ByRef mudou As Boolean, _
                            ByRef resLotes As String) As String
    Dim nomes As Variant, orig As Variant, nOrig As Long, msg As String, lo As ListObject
    Dim rec As Variant, nRec As Long, cabRec As Object, i As Long, ids As Object
    Dim novos As Variant, nSemId As Long, rec2 As Variant, nRec2 As Long, F As Variant, nF As Long, nCQ0 As Long
    Dim q As Variant, mRec As Boolean, mCQ As Boolean, mQA As Boolean, gravRec As Long, gravCQ As Long

    mTempos = "": mT0 = Timer: nNovos = 0: mudou = False: resLotes = ""
    If mLog Is Nothing Then IniciarLog "VBA", IIf(historico, "HISTORICO", "SEAC")
    ' ---- 1. origem
    mEtapa = "ler o DB_SEAC"
    If Not historico Then
        msg = mEtlReceb.LerOrigem(mIntegracao.Cfg("CAMINHO_DB_SEAC"), mIntegracao.Cfg("TABELA_ORIGEM"), nomes, orig, nOrig)
        If Len(msg) > 0 Then Err.Raise vbObjectError + 591, "mEtlSync", msg
        If nOrig = 0 Then Err.Raise vbObjectError + 592, "mEtlSync", _
            "a tabela do setor no DB_SEAC veio VAZIA -- nada foi gravado (confira o DB_SEAC antes de atualizar)"
    End If
    mLog("LIDOS_ORIGEM") = nOrig
    Marca "origem"

    ' ---- 2. recebimento (staging)
    mEtapa = "calcular o recebimento"
    Set lo = mIntegracao.AcharTabela(TB_REC)
    If lo Is Nothing Then Err.Raise vbObjectError + 593, "mEtlSync", "tabela " & TB_REC & " nao encontrada"
    Set cabRec = Cabecalho(lo)
    nRec = lo.ListRows.Count
    If nRec > 0 Then rec = lo.DataBodyRange.Value2
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To nRec
        ids(CStr(rec(i, cabRec("ID_REGISTRO")))) = 1
    Next i
    If nOrig > 0 Then
        novos = mEtlReceb.NovosParaReceber(nomes, orig, nOrig, cabRec, ids, UCase$(Trim$(mIntegracao.Cfg("PREFIXO_ID"))), _
                                           mEtlReceb.Carimbo(), nNovos, nSemId)
    End If
    If nNovos > 0 Then
        rec2 = mEtlReceb.JuntarOrdenado(rec, nRec, novos, nNovos, cabRec)
        nRec2 = nRec + nNovos
    Else
        rec2 = rec: nRec2 = nRec
    End If
    mLog("SEM_ID") = nSemId: mLog("NOVOS") = nNovos: mLog("TOTAL_RECEBIDO") = nRec2
    DatasMinMax rec2, nRec2, cabRec("DATA_HORA")
    Marca "recebimento"

    ' ---- 3. Principal - Resultados (staging, sobre o recebimento novo)
    mEtapa = "calcular a Principal - Resultados"
    nCQ0 = LinhasDe(FONTE_CQ)
    mEtlCQ.UsarRecebimento ComoData(rec2, nRec2, cabRec), cabRec, nRec2
    F = mEtlCQ.CalcularCQFinal()
    If IsEmpty(F) Then nF = 0 Else nF = UBound(F, 1)
    mLog("LINHAS_CQ") = nF
    Marca "cq"

    ' ---- 4. validar ANTES de gravar
    mEtapa = "validar antes de gravar"
    If nRec2 <> nRec + nNovos Then Err.Raise vbObjectError + 594, "mEtlSync", "recebimento: contagem inconsistente"
    If nNovos > 0 Then
        For i = 1 To nNovos
            If ids.Exists(CStr(novos(i, cabRec("ID_REGISTRO")))) Then _
                Err.Raise vbObjectError + 595, "mEtlSync", "recebimento: ID ja recebido voltou como novo"
            ids(CStr(novos(i, cabRec("ID_REGISTRO")))) = 1
        Next i
    End If
    If nRec2 > 0 And nF = 0 Then Err.Raise vbObjectError + 596, "mEtlSync", _
        "a Principal - Resultados voltaria VAZIA com " & nRec2 & " resultados recebidos -- nada foi gravado"
    If nCQ0 > 100 And nF < ENCOLHE_MAX * nCQ0 Then Err.Raise vbObjectError + 597, "mEtlSync", _
        "a Principal - Resultados cairia de " & nCQ0 & " para " & nF & " linhas -- nada foi gravado. " & _
        "Se for esperado, atualize uma vez com Cfg_Integracao MOTOR_DADOS = PQ."
    ConferirEsquema FONTE_CQ, mEtlCQ.NomesColunas()
    ConferirEsquema TB_QA, mEtlQA.NomesColunasQA()
    Marca "validar"

    ' ---- 5. publicar (incremental)
    mEtapa = "gravar o recebimento"
    If nNovos > 0 Then gravRec = Publicar(lo, rec2, nRec2, NomesDe(cabRec), mRec)
    mEtapa = "gravar a Principal - Resultados"
    gravCQ = Publicar(mIntegracao.AcharTabela(FONTE_CQ), F, nF, mEtlCQ.NomesColunas(), mCQ)
    mLog("GRAVADAS") = gravRec + gravCQ
    Marca "gravar"

    ' ---- 6. lotes novos e conferencia
    mEtapa = "registrar lotes novos"
    If mCQ Then resLotes = mLotes.RegistrarLotesRecebidos()
    mEtapa = "conferencia (QA_INTEGRACAO)"
    q = mEtlQA.CalcularQA(F, rec2, nRec2, cabRec, orig, nomes, nOrig, historico)
    Set lo = mIntegracao.AcharTabela(TB_QA)
    If Not lo Is Nothing Then
        If IsEmpty(q) Then
            Publicar lo, q, 0, mEtlQA.NomesColunasQA(), mQA
        Else
            Publicar lo, q, UBound(q, 1), mEtlQA.NomesColunasQA(), mQA
        End If
    End If
    Marca "qa"
    mudou = mRec Or mCQ Or mQA
    mEtapa = ""
    Sincronizar = mTempos
End Function

' Para o instalador: as tabelas publicadas tem as colunas do calculo? Refaz pelo Power Query a que nao tiver.
' Devolve "" ou o que foi refeito; "ERRO|..." em falha (chamada de fora: erro nao tratado abriria o depurador).
Public Function ConferirEsquemas() As String
    On Error GoTo falha
    IniciarLog "VBA", "INSTALADOR"
    mLog("ESQUEMA") = ""
    ConferirEsquema FONTE_CQ, mEtlCQ.NomesColunas()
    ConferirEsquema TB_QA, mEtlQA.NomesColunasQA()
    ConferirEsquemas = mLog("ESQUEMA")
    Set mLog = Nothing
    Exit Function
falha:
    ConferirEsquemas = "ERRO|" & Err.Description
    Set mLog = Nothing
End Function

' A tblDB_Organizado (so visualizacao) pelo Power Query, sob demanda.
Public Sub AtualizarOrganizado()
    mIntegracao.Refrescar "tblDB_Organizado"
End Sub

' ---------------------------------------------------------------- Sync_Log
Public Sub IniciarLog(ByVal motor As String, ByVal modo As String)
    Set mLog = CreateObject("Scripting.Dictionary")
    mLog("INICIO") = mEtlReceb.Carimbo()
    mLog("MOTOR_DADOS") = motor
    mLog("MODO_FONTE") = modo
    mLog("VERSAO") = VERSAO_ETL
    mLog("USUARIO") = Environ$("USERNAME")
End Sub

Public Sub AnotarLog(ByVal campo As String, ByVal valor As Variant)
    If mLog Is Nothing Then Exit Sub
    mLog(campo) = valor
End Sub

' Acrescenta a linha desta execucao ao tblSyncLog (aba Sync_Log). Nunca derruba a atualizacao.
Public Sub FecharLog(ByVal resultado As String, ByVal erro As String)
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, h As Variant, j As Long, n As Long, linha() As Variant
    If mLog Is Nothing Then Exit Sub
    On Error GoTo fim
    mLog("FIM") = mEtlReceb.Carimbo()
    mLog("TOTAL_S") = Round((mLog("FIM") - mLog("INICIO")) * 86400#, 2)
    mLog("RESULTADO") = resultado
    mLog("ERRO") = Left$(erro, 500)
    Set lo = mIntegracao.AcharTabela(TB_LOG)
    If lo Is Nothing Then GoTo fim
    Set ws = lo.Parent
    prot = mSeguranca.LiberarEscrita(ws)
    h = lo.HeaderRowRange.Value
    ReDim linha(1 To 1, 1 To UBound(h, 2))
    For j = 1 To UBound(h, 2)
        If mLog.Exists(CStr(h(1, j))) Then linha(1, j) = mLog(CStr(h(1, j)))
    Next j
    n = lo.ListRows.Count
    If n >= LOG_MAX Then lo.ListRows(1).Range.Delete xlShiftUp: n = n - 1
    lo.ListRows.Add.Range.Value2 = linha
    mSeguranca.RestaurarProtecao ws, prot
fim:
    Set mLog = Nothing
End Sub

' ---------------------------------------------------------------- publicar
' Grava a matriz m (colunas na ordem de nomes; n linhas) na tabela lo: so da primeira linha que difere em diante,
' ajustando o tamanho. mudou = False quando a tabela ja tinha exatamente isso (nada e gravado). Colunas casadas
' pelo NOME do cabecalho; coluna faltando ou sobrando = erro (nada gravado). Devolve as linhas gravadas.
Private Function Publicar(ByVal lo As ListObject, ByRef m As Variant, ByVal n As Long, ByVal nomes As Variant, _
                          ByRef mudou As Boolean) As Long
    Dim idx As Object, h As Variant, nc As Long, j As Long, i As Long, mapa() As Long, atual As Variant, nAt As Long
    Dim r0 As Long, k As Long, out() As Variant, ws As Worksheet, prot As Boolean, temTexto As Boolean, nE As Long, sE As String
    mudou = False
    Set idx = CreateObject("Scripting.Dictionary")
    For j = LBound(nomes) To UBound(nomes)
        idx(CStr(nomes(j))) = j - LBound(nomes) + 1
    Next j
    h = lo.HeaderRowRange.Value
    nc = UBound(h, 2)
    If nc <> idx.Count Then Err.Raise vbObjectError + 598, "mEtlSync", lo.Name & ": a tabela tem " & nc & _
                                      " colunas e o calculo " & idx.Count & " -- nada foi gravado"
    ReDim mapa(1 To nc)
    For j = 1 To nc
        If Not idx.Exists(CStr(h(1, j))) Then Err.Raise vbObjectError + 598, "mEtlSync", _
            lo.Name & ": coluna " & h(1, j) & " sem correspondente no calculo -- nada foi gravado"
        mapa(j) = idx(CStr(h(1, j)))
    Next j
    nAt = lo.ListRows.Count
    If nAt > 0 Then atual = lo.DataBodyRange.Value2
    For i = 1 To IIf(n < nAt, n, nAt)
        For j = 1 To nc
            ' caminho rapido: mesmo tipo (nao erro) compara direto -- texto binario (Option Compare Binary)
            If VarType(atual(i, j)) = VarType(m(i, mapa(j))) And VarType(atual(i, j)) <> vbError Then
                If atual(i, j) <> m(i, mapa(j)) Then r0 = i: Exit For
            ElseIf Not Igual(atual(i, j), m(i, mapa(j))) Then
                r0 = i: Exit For
            End If
        Next j
        If r0 > 0 Then Exit For
    Next i
    If r0 = 0 Then
        If n = nAt Then Exit Function
        r0 = IIf(n < nAt, n + 1, nAt + 1)
    End If
    mudou = True
    atual = Empty
    Set ws = lo.Parent
    prot = mSeguranca.LiberarEscrita(ws)
    On Error GoTo falha
    If n < nAt Then
        lo.DataBodyRange.Rows(n + 1).Resize(nAt - n, nc).ClearContents
        If n = 0 Then lo.DataBodyRange.Delete Else lo.Resize lo.HeaderRowRange.Resize(n + 1, nc)
    ElseIf n > nAt Then
        lo.Resize lo.HeaderRowRange.Resize(n + 1, nc)
    End If
    If n >= r0 Then
        k = n - r0 + 1
        ReDim out(1 To k, 1 To nc)
        For j = 1 To nc
            temTexto = False
            For i = 1 To k
                out(i, j) = m(r0 + i - 1, mapa(j))
                If VarType(out(i, j)) = vbString Then temTexto = True
            Next i
            ' texto fica texto: "8588" nao vira numero, "=..." nao vira formula. Numero gravado continua numero.
            If temTexto Then lo.DataBodyRange.Cells(r0, j).Resize(k, 1).NumberFormat = "@"
        Next j
        lo.DataBodyRange.Rows(r0).Resize(k, nc).Value2 = out
        ' texto acima do limite da gravacao em bloco: celula a celula (sem truncar)
        For i = 1 To k
            For j = 1 To nc
                If VarType(out(i, j)) = vbString Then
                    If Len(out(i, j)) > 8000 Then lo.DataBodyRange.Cells(r0 + i - 1, j).Value2 = out(i, j)
                End If
            Next j
        Next i
        Publicar = k
    End If
    mSeguranca.RestaurarProtecao ws, prot
    Exit Function
falha:
    nE = Err.Number: sE = Err.Description
    mSeguranca.RestaurarProtecao ws, prot
    Err.Raise nE, "mEtlSync.Publicar(" & lo.Name & ")", sE
End Function

' As colunas da tabela sao as do calculo? Uma tabela gravada por uma versao antiga da consulta (ex.: a Hematologia
' de producao em 20/10/2026 tinha 39 colunas na tblCQ_Final, o calculo 40) e refeita UMA vez pelo Power Query, que
' cria e ordena as colunas como a consulta atual manda; continuando diferente = erro, nada gravado.
Private Sub ConferirEsquema(ByVal nomeTab As String, ByVal nomes As Variant)
    Dim lo As ListObject, dif As String
    Set lo = mIntegracao.AcharTabela(nomeTab)
    If lo Is Nothing Then Err.Raise vbObjectError + 599, "mEtlSync", "tabela " & nomeTab & " nao encontrada"
    dif = DiferencaColunas(lo, nomes)
    If Len(dif) = 0 Then Exit Sub
    Application.StatusBar = "Atualizando dados: " & nomeTab & " com colunas antigas -- refeita uma vez pelo Power Query..."
    mIntegracao.Refrescar nomeTab
    mAuditoria.RegistrarLog "ESQUEMA_REFEITO_PQ", nomeTab & " refeita pelo Power Query (" & dif & ")"
    mLog("ESQUEMA") = mLog("ESQUEMA") & IIf(Len(mLog("ESQUEMA")) > 0, " | ", "") & nomeTab & ": " & dif
    dif = DiferencaColunas(mIntegracao.AcharTabela(nomeTab), nomes)
    If Len(dif) > 0 Then Err.Raise vbObjectError + 598, "mEtlSync", nomeTab & ": colunas diferentes do calculo (" & dif & _
                                   ") mesmo depois do Power Query -- nada foi gravado"
End Sub

Private Function DiferencaColunas(ByVal lo As ListObject, ByVal nomes As Variant) As String
    Dim tem As Object, quer As Object, h As Variant, j As Long, k As Variant, falta As String, sobra As String
    Set tem = CreateObject("Scripting.Dictionary"): Set quer = CreateObject("Scripting.Dictionary")
    h = lo.HeaderRowRange.Value
    For j = 1 To UBound(h, 2): tem(CStr(h(1, j))) = 1: Next j
    For j = LBound(nomes) To UBound(nomes): quer(CStr(nomes(j))) = 1: Next j
    For Each k In quer.Keys
        If Not tem.Exists(k) Then falta = falta & IIf(Len(falta) > 0, ", ", "") & k
    Next k
    For Each k In tem.Keys
        If Not quer.Exists(k) Then sobra = sobra & IIf(Len(sobra) > 0, ", ", "") & k
    Next k
    If Len(falta) > 0 Then DiferencaColunas = "faltam: " & falta
    If Len(sobra) > 0 Then DiferencaColunas = DiferencaColunas & IIf(Len(falta) > 0, "; ", "") & "sobram: " & sobra
End Function

' igualdade de celula: branco (vazio ou "") = branco; texto binario; numero/data pelo valor
Private Function Igual(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim ta As Long, tb As Long
    If Branco(a) Then Igual = Branco(b): Exit Function
    If Branco(b) Then Exit Function
    If IsError(a) Or IsError(b) Then Exit Function
    ta = VarType(a): tb = VarType(b)
    If ta = vbString Or tb = vbString Then
        If ta = tb Then Igual = (StrComp(a, b, vbBinaryCompare) = 0)
    ElseIf ta = vbBoolean Or tb = vbBoolean Then
        If ta = tb Then Igual = (a = b)
    Else
        Igual = (CDbl(a) = CDbl(b))
    End If
End Function

Private Function Branco(ByRef v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        Branco = True
    ElseIf VarType(v) = vbString Then
        Branco = (Len(v) = 0)
    End If
End Function

' ---------------------------------------------------------------- utilitarios
Private Function Cabecalho(ByVal lo As ListObject) As Object
    Dim h As Variant, j As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    h = lo.HeaderRowRange.Value
    For j = 1 To UBound(h, 2): d(CStr(h(1, j))) = j: Next j
    Set Cabecalho = d
End Function

Private Function NomesDe(ByVal cab As Object) As Variant
    Dim r() As String, k As Variant
    ReDim r(1 To cab.Count)
    For Each k In cab.Keys
        r(cab(k)) = CStr(k)
    Next k
    NomesDe = r
End Function

Private Function LinhasDe(ByVal nome As String) As Long
    Dim lo As ListObject
    Set lo = mIntegracao.AcharTabela(nome)
    If Not lo Is Nothing Then LinhasDe = lo.ListRows.Count
End Function

' copia do recebimento com as colunas de data como Date -- o mesmo que a leitura .Value que o mEtlCQ faz da tabela
Private Function ComoData(ByRef rec As Variant, ByVal n As Long, ByVal cab As Object) As Variant
    Dim r As Variant, c As Variant, i As Long, j As Long
    If n = 0 Then Exit Function
    r = rec
    For Each c In Array("DATA", "DATA_HORA", "RECEBIDO_EM")
        If cab.Exists(c) Then
            j = cab(c)
            For i = 1 To n
                If VarType(r(i, j)) = vbDouble Then r(i, j) = CDate(r(i, j))
            Next i
        End If
    Next c
    ComoData = r
End Function

Private Sub DatasMinMax(ByRef rec As Variant, ByVal n As Long, ByVal c As Long)
    Dim i As Long, v As Variant, mn As Double, mx As Double
    For i = 1 To n
        v = rec(i, c)
        If VarType(v) = vbDouble Or VarType(v) = vbDate Then
            If mn = 0 Or CDbl(v) < mn Then mn = CDbl(v)
            If CDbl(v) > mx Then mx = CDbl(v)
        End If
    Next i
    If mn > 0 Then mLog("DATA_MIN") = mn: mLog("DATA_MAX") = mx
End Sub
