Attribute VB_Name = "mDados"
Option Explicit
' ===== CAMADA DE DADOS (Fase 2) =====
' DB_Resultados e a UNICA fonte operacional. Formato vertical normalizado.
' Schema: A=RUN | B=Data | C=Nivel | D=Lote | E=Analito | F=Resultado | G=Status
Public Const BANCO As String = "DB_Resultados"
Public Const VIEW As String = "Resultados"
Public Const BANCO_R0 As Long = 4
Public Const COL_RUN As Long = 1
Public Const COL_DATA As Long = 2
Public Const COL_NIVEL As Long = 3
Public Const COL_LOTE As Long = 4
Public Const COL_ANALITO As Long = 5
Public Const COL_RESULT As Long = 6
Public Const COL_STATUS As Long = 7
Public Const ST_ATIVO As String = "Ativo"
Public Const ST_EXCLUIDO As String = "Excluído"

' ===================== IDENTIDADE DA CORRIDA =====================
' O RUN identifica a CORRIDA, nao o resultado. Uma corrida cobre varios
' analitos e os 3 niveis; a chave de um resultado continua sendo
' RUN + Nivel + Analito (ver ChaveReg).
'
' ANTES: NovoRUN devolvia max(RUN encontrado no banco) + 1. O contador era
' DERIVADO do conteudo da tabela, e isso permite REUTILIZACAO: apagando
' linhas de verdade, o maximo cai e o proximo RUN repete um numero ja usado.
' Num Audit_Log append-only isso e irreparavel -- duas corridas distintas com
' a mesma identidade, e nao da para reescrever o log para desambiguar.
' Havia ainda um defeito de lote: numa importacao com DUAS datas, nenhuma
' delas ainda esta no banco, entao ambas recebiam max+1 -- o MESMO RUN para
' corridas diferentes.
'
' AGORA: o RUN vem de um contador PERSISTIDO (proxRUN, na aba Corridas) e a
' corrida e registrada numa tabela propria. O contador nunca e recalculado a
' partir dos dados; so e reparado PARA CIMA, nunca para baixo. Assim, mesmo
' que o contador se perca ou que algum caminho de gravacao escape, o proximo
' RUN nunca repete um numero ja presente no banco ou no registro de corridas.
'
' O RUN e IMUTAVEL: uma vez atribuido, nao muda por exclusao, correcao de data
' ou reordenacao. UpsertResultados nunca reescreve a coluna do RUN.
Public Const CORRIDAS As String = "Corridas"
Public Const CORRIDAS_R0 As Long = 4
Public Const CC_RUN As Long = 1
Public Const CC_DATA As Long = 2
Public Const CC_HORA As Long = 3
Public Const CC_TURNO As Long = 4
Public Const CC_LOTE As Long = 5
Public Const CC_EQUIP As Long = 6
Public Const CC_USER As Long = 7
Public Const CC_CRIADO As Long = 8
Public Const CC_CRIADOPOR As Long = 9

Public Function UltimaLinhaBanco() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANCO)
    UltimaLinhaBanco = ws.Cells(ws.rows.Count, COL_RUN).End(xlUp).Row
    If UltimaLinhaBanco < BANCO_R0 Then UltimaLinhaBanco = BANCO_R0 - 1
End Function

Public Function LoteAtivoCore() As String
    On Error Resume Next
    LoteAtivoCore = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
End Function

' Le o banco inteiro para memoria de uma vez (performance: sem loop celula a celula).
Public Function CarregarDB() As Variant
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lastRow = UltimaLinhaBanco()
    If lastRow < BANCO_R0 Then
        CarregarDB = Empty
    Else
        CarregarDB = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_STATUS)).Value
    End If
End Function

Public Function ChaveReg(ByVal run As Long, ByVal nivel As Long, ByVal analito As String) As String
    ' Chave de unicidade do registro. O Nivel ENTRA na chave: um mesmo RUN abrange
    ' todos os niveis da corrida, entao RUN+Analito sozinho colidiria entre niveis.
    ChaveReg = CStr(run) & "|" & CStr(nivel) & "|" & UCase$(Trim$(analito))
End Function


Public Function UltimaLinhaCorridas() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CORRIDAS)
    UltimaLinhaCorridas = ws.Cells(ws.rows.Count, CC_RUN).End(xlUp).Row
    If UltimaLinhaCorridas < CORRIDAS_R0 Then UltimaLinhaCorridas = CORRIDAS_R0 - 1
End Function

' Maior RUN ja existente, olhando o banco E o registro de corridas.
' E o piso do contador: ele nunca pode ficar abaixo disso.
Private Function MaiorRUNUsado() As Long
    Dim dados As Variant, i As Long, mx As Long, ws As Worksheet, lastRow As Long
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If IsNumeric(dados(i, COL_RUN)) Then
                If CLng(Val(dados(i, COL_RUN))) > mx Then mx = CLng(Val(dados(i, COL_RUN)))
            End If
        Next i
    End If
    Set ws = ThisWorkbook.Sheets(CORRIDAS)
    lastRow = UltimaLinhaCorridas()
    For i = CORRIDAS_R0 To lastRow
        If IsNumeric(ws.Cells(i, CC_RUN).Value) Then
            If CLng(Val(ws.Cells(i, CC_RUN).Value)) > mx Then mx = CLng(Val(ws.Cells(i, CC_RUN).Value))
        End If
    Next i
    MaiorRUNUsado = mx
End Function

' Proximo RUN SEM consumir. Auto-reparo: se o contador estiver atrasado em
' relacao ao que ja existe, ele SOBE. Nunca desce -- e essa assimetria que
' torna a reutilizacao impossivel.
Public Function ProximoRUNPeek() As Long
    Dim r As Range, atual As Long, piso As Long
    Set r = ThisWorkbook.Names("proxRUN").RefersToRange
    atual = CLng(Val(r.Value))
    piso = MaiorRUNUsado() + 1
    If atual < piso Then
        atual = piso
        r.Value = atual
    End If
    ProximoRUNPeek = atual
End Function

' Consome o proximo RUN.
Public Function AlocarRUN() As Long
    Dim n As Long
    n = ProximoRUNPeek()
    ThisWorkbook.Names("proxRUN").RefersToRange.Value = n + 1
    AlocarRUN = n
End Function

' RUN de uma corrida ja registrada. Devolve 0 se ela nao existe.
' Mantida para consultas e compatibilidade. A gravacao nao usa mais
' (data, lote, turno) como chave unica: isso colidia quando havia duas
' corridas no mesmo dia e lote.
Public Function RunDaCorrida(ByVal dt As Date, ByVal loteCore As String, _
                             Optional ByVal turno As String = "") As Long
    Dim ws As Worksheet, i As Long, lastRow As Long
    Set ws = ThisWorkbook.Sheets(CORRIDAS)
    lastRow = UltimaLinhaCorridas()
    For i = CORRIDAS_R0 To lastRow
        If IsDate(ws.Cells(i, CC_DATA).Value) Then
            If CDate(ws.Cells(i, CC_DATA).Value) = dt Then
                If Trim$(CStr(ws.Cells(i, CC_LOTE).Value)) = Trim$(loteCore) Then
                    If Trim$(CStr(ws.Cells(i, CC_TURNO).Value)) = Trim$(turno) Then
                        RunDaCorrida = CLng(Val(ws.Cells(i, CC_RUN).Value))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' Ultima corrida do mesmo dia/lote/turno que AINDA NAO recebeu resultados do
' nivel informado. Assim, N1/N2/N3 gravados separadamente continuam compondo
' uma mesma corrida; repetir um nivel ja gravado cria outro RUN.
Private Function RunDisponivelParaNivel(ByVal dt As Date, ByVal loteCore As String, _
                                        ByVal turno As String, ByVal nivel As Long) As Long
    Dim ws As Worksheet, dados As Variant, usados As Object
    Dim i As Long, lastRow As Long, r As Long
    If nivel < 1 Then Exit Function

    Set usados = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If CLng(Val(CStr(dados(i, COL_NIVEL)))) = nivel Then
                usados(CStr(CLng(Val(CStr(dados(i, COL_RUN)))))) = True
            End If
        Next i
    End If

    Set ws = ThisWorkbook.Sheets(CORRIDAS)
    lastRow = UltimaLinhaCorridas()
    For i = lastRow To CORRIDAS_R0 Step -1
        If IsDate(ws.Cells(i, CC_DATA).Value) Then
            If CDate(ws.Cells(i, CC_DATA).Value) = dt And _
               Trim$(CStr(ws.Cells(i, CC_LOTE).Value)) = Trim$(loteCore) And _
               Trim$(CStr(ws.Cells(i, CC_TURNO).Value)) = Trim$(turno) Then
                r = CLng(Val(CStr(ws.Cells(i, CC_RUN).Value)))
                If r > 0 And Not usados.Exists(CStr(r)) Then
                    RunDisponivelParaNivel = r
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Previsao para exibir enquanto o usuario digita. NAO aloca e NAO registra.
' Usar em qualquer caminho que nao seja gravacao efetiva.
Public Function PreverRUN(ByVal dt As Date, ByVal loteCore As String, _
                          Optional ByVal turno As String = "", _
                          Optional ByVal nivel As Long = 0) As Long
    Dim r As Long
    If nivel > 0 Then r = RunDisponivelParaNivel(dt, loteCore, turno, nivel)
    If r > 0 Then
        PreverRUN = r
    Else
        PreverRUN = ProximoRUNPeek()
    End If
End Function

' Identidade definitiva da corrida. Com nivel informado, reaproveita somente
' uma corrida ainda sem esse nivel. Sem nivel (importacao em lote), sempre
' cria uma corrida nova; o chamador agrupa as linhas do mesmo lote/data e
' chama esta funcao uma vez por grupo.
Public Function ObterOuCriarRUN(ByVal dt As Date, ByVal loteCore As String, _
                                Optional ByVal turno As String = "", _
                                Optional ByVal nivel As Long = 0) As Long
    Dim ws As Worksheet, r As Long, lin As Long
    If nivel > 0 Then r = RunDisponivelParaNivel(dt, loteCore, turno, nivel)
    If r > 0 Then ObterOuCriarRUN = r: Exit Function

    r = AlocarRUN()
    Set ws = ThisWorkbook.Sheets(CORRIDAS)
    lin = UltimaLinhaCorridas() + 1
    If lin < CORRIDAS_R0 Then lin = CORRIDAS_R0
    ws.Cells(lin, CC_RUN).Value = r
    ws.Cells(lin, CC_DATA).Value = dt
    ws.Cells(lin, CC_TURNO).Value = turno
    ws.Cells(lin, CC_LOTE).NumberFormat = "@"
    ws.Cells(lin, CC_LOTE).Value = loteCore
    ws.Cells(lin, CC_USER).Value = Application.UserName
    ws.Cells(lin, CC_CRIADO).Value = Now
    ws.Cells(lin, CC_CRIADOPOR).Value = Environ$("USERNAME")
    ObterOuCriarRUN = r
End Function

' MANTIDA POR COMPATIBILIDADE, com semantica de PREVISAO.
' Qualquer chamador antigo que nao tenha sido migrado passa a nao alocar --
' falha segura. Se ainda assim um RUN previsto for gravado no banco sem passar
' por ObterOuCriarRUN, o auto-reparo de ProximoRUNPeek eleva o contador acima
' dele na proxima chamada, e o numero nao e reutilizado.
Public Function NovoRUN(ByVal dt As Date, ByVal loteCore As String) As Long
    NovoRUN = PreverRUN(dt, loteCore)
End Function

' ===== UPSERT EM LOTE (auditado) =====
' regs: array (1..n, 1..7) ja no schema do banco. Atualiza o que existir
' (mesma chave RUN|Nivel|Analito) e acrescenta o resto - nunca duplica.
' Devolve "novos|atualizados|bloqueados".
'
' ITEM 2.3 DO GATE. A versao anterior forcava Status = Ativo em toda
' atualizacao: reenviar a mesma chave RESSUSCITAVA uma linha excluida, sem
' deixar registro. Agora, linha nao ativa e PRESERVADA e a tentativa e
' auditada. Reverter uma exclusao passa a exigir acao propria, com parecer.
Public Function UpsertResultados(ByRef regs As Variant) As String
    Dim ws As Worksheet, dados As Variant, idx As Object, vistos As Object
    Dim i As Long, c As Long, nRegs As Long, lastRow As Long
    Dim novos As Long, atual As Long, bloq As Long, nAdd As Long
    Dim k As String, lin As Long, stAntes As String, acaoUp As String
    Dim acao() As Integer, linhaAlvo() As Long, antesRes() As Variant
    Dim antesData() As Variant, antesLote() As Variant, antesStatus() As Variant
    Dim addBuf() As Variant, outp() As Variant, dbAntes As Variant
    Dim gravou As Boolean, screenAntes As Boolean, nErr As Long, sErr As String

    If IsEmpty(regs) Then UpsertResultados = "0|0|0": Exit Function
    screenAntes = Application.ScreenUpdating
    On Error GoTo rollback
    nRegs = UBound(regs, 1)
    Set ws = ThisWorkbook.Sheets(BANCO)
    Set idx = CreateObject("Scripting.Dictionary")
    Set vistos = CreateObject("Scripting.Dictionary")
    idx.CompareMode = 1: vistos.CompareMode = 1

    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
                k = ChaveReg(CLng(dados(i, COL_RUN)), CLng(dados(i, COL_NIVEL)), CStr(dados(i, COL_ANALITO)))
                If idx.Exists(k) Then
                    Err.Raise vbObjectError + 531, "mDados.UpsertResultados", _
                              "Banco inconsistente: chave duplicada " & k & ". Nenhum dado foi gravado."
                End If
                idx.Add k, BANCO_R0 + i - 1
            End If
        Next i
    End If

    lastRow = UltimaLinhaBanco()
    ReDim acao(1 To nRegs)
    ReDim linhaAlvo(1 To nRegs)
    ReDim antesRes(1 To nRegs)
    ReDim antesData(1 To nRegs)
    ReDim antesLote(1 To nRegs)
    ReDim antesStatus(1 To nRegs)
    ReDim addBuf(1 To nRegs, 1 To COL_STATUS)

    ' PRE-FLIGHT completo. Nao escreve em nenhuma planilha antes de validar o
    ' lote inteiro, inclusive capacidade e duplicidade dentro da propria carga.
    For i = 1 To nRegs
        If Not IsNumeric(regs(i, COL_RUN)) Or CLng(Val(CStr(regs(i, COL_RUN)))) < 1 Then _
            Err.Raise vbObjectError + 532, "mDados.UpsertResultados", "RUN invalido na linha " & i & "."
        If Not IsDate(regs(i, COL_DATA)) Then _
            Err.Raise vbObjectError + 533, "mDados.UpsertResultados", "Data invalida na linha " & i & "."
        If Not IsNumeric(regs(i, COL_NIVEL)) Or CLng(Val(CStr(regs(i, COL_NIVEL)))) < 1 Then _
            Err.Raise vbObjectError + 534, "mDados.UpsertResultados", "Nivel invalido na linha " & i & "."
        If Len(Trim$(CStr(regs(i, COL_ANALITO)))) = 0 Then _
            Err.Raise vbObjectError + 535, "mDados.UpsertResultados", "Analito vazio na linha " & i & "."

        k = ChaveReg(CLng(regs(i, COL_RUN)), CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)))
        If vistos.Exists(k) Then
            Err.Raise vbObjectError + 536, "mDados.UpsertResultados", _
                      "Carga recusada: chave repetida " & k & ". Nenhum dado foi gravado."
        End If
        vistos.Add k, True

        If idx.Exists(k) Then
            lin = CLng(idx(k))
            linhaAlvo(i) = lin
            stAntes = Trim$(CStr(ws.Cells(lin, COL_STATUS).Value))
            antesRes(i) = ws.Cells(lin, COL_RESULT).Value
            antesData(i) = ws.Cells(lin, COL_DATA).Value
            antesLote(i) = ws.Cells(lin, COL_LOTE).Value
            antesStatus(i) = ws.Cells(lin, COL_STATUS).Value
            If stAntes <> "" And stAntes <> ST_ATIVO Then
                acao(i) = 2: bloq = bloq + 1
            Else
                acao(i) = 1: atual = atual + 1
            End If
        Else
            acao(i) = 0
            nAdd = nAdd + 1: novos = novos + 1
            For c = 1 To COL_STATUS
                addBuf(nAdd, c) = regs(i, c)
            Next c
        End If
    Next i
    ExigirCapacidade nAdd

    Application.ScreenUpdating = False

    ' Tentativas bloqueadas sao eventos reais, mas nao alteram DB_Resultados.
    For i = 1 To nRegs
        If acao(i) = 2 Then
            Auditar CAT_DADO, AC_BLOQUEIO, "mDados", _
                    CLng(regs(i, COL_RUN)), regs(i, COL_DATA), "", CStr(regs(i, COL_LOTE)), _
                    CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)), _
                    antesRes(i), regs(i, COL_RESULT), CStr(antesStatus(i)), CStr(antesStatus(i)), _
                    "Reenvio de registro nao ativo", _
                    "Reenvio recusado: reverter exclusao exige acao propria e justificada."
        End If
    Next i

    ' Snapshot transacional do banco. Se qualquer escrita, auditoria ou ajuste
    ' derivado falhar, o bloco original e restaurado antes de propagar o erro.
    If lastRow >= BANCO_R0 Then
        dbAntes = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_STATUS)).Value
    End If
    gravou = True

    For i = 1 To nRegs
        If acao(i) = 1 Then
            lin = linhaAlvo(i)
            ws.Cells(lin, COL_DATA).Value = regs(i, COL_DATA)
            ws.Cells(lin, COL_LOTE).NumberFormat = "@"
            ws.Cells(lin, COL_LOTE).Value = regs(i, COL_LOTE)
            ws.Cells(lin, COL_RESULT).Value = regs(i, COL_RESULT)
            ws.Cells(lin, COL_STATUS).Value = ST_ATIVO
        End If
    Next i

    If nAdd > 0 Then
        ReDim outp(1 To nAdd, 1 To COL_STATUS)
        For i = 1 To nAdd
            For c = 1 To COL_STATUS
                outp(i, c) = addBuf(i, c)
            Next c
        Next i
        ws.Range(ws.Cells(lastRow + 1, COL_LOTE), ws.Cells(lastRow + nAdd, COL_LOTE)).NumberFormat = "@"
        ws.Range(ws.Cells(lastRow + 1, COL_DATA), ws.Cells(lastRow + nAdd, COL_DATA)).NumberFormatLocal = "dd/mm/aaaa;@"
        ws.Range(ws.Cells(lastRow + 1, COL_RUN), ws.Cells(lastRow + nAdd, COL_STATUS)).Value = outp
    End If

    AtualizarFlagsBanco

    For i = 1 To nRegs
        If acao(i) = 1 Then
            If Trim$(CStr(regs(i, COL_RESULT))) = "" Then acaoUp = AC_APAGADO Else acaoUp = AC_ALTERACAO
            Auditar CAT_DADO, acaoUp, "mDados", _
                    CLng(regs(i, COL_RUN)), regs(i, COL_DATA), "", CStr(regs(i, COL_LOTE)), _
                    CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)), _
                    antesRes(i), regs(i, COL_RESULT), CStr(antesStatus(i)), ST_ATIVO, "", ""
        End If
    Next i
    For i = 1 To nAdd
        Auditar CAT_DADO, AC_INCLUSAO, "mDados", _
                CLng(addBuf(i, COL_RUN)), addBuf(i, COL_DATA), "", CStr(addBuf(i, COL_LOTE)), _
                CLng(addBuf(i, COL_NIVEL)), CStr(addBuf(i, COL_ANALITO)), _
                Empty, addBuf(i, COL_RESULT), "", CStr(addBuf(i, COL_STATUS)), "", ""
    Next i

    Application.ScreenUpdating = screenAntes
    UpsertResultados = CStr(novos) & "|" & CStr(atual) & "|" & CStr(bloq)
    Exit Function

rollback:
    nErr = Err.Number: sErr = Err.Description
    On Error Resume Next
    If gravou Then
        If lastRow >= BANCO_R0 Then
            ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_STATUS)).Value = dbAntes
        End If
        If nAdd > 0 Then
            ws.Range(ws.Cells(lastRow + 1, COL_RUN), ws.Cells(lastRow + nAdd, COL_STATUS)).ClearContents
        End If
        AtualizarFlagsBanco
        Auditar CAT_SIS, "TRANSACAO_REVERTIDA", "mDados", 0, Empty, "", "", 0, "", _
                Empty, Empty, "", "", "Falha no lote", sErr
    End If
    Application.ScreenUpdating = screenAntes
    On Error GoTo 0
    Err.Raise nErr, "mDados.UpsertResultados", sErr
End Function

' Exclusao LOGICA por RUN + Nivel + lista de analitos (Dictionary de nomes em UCase).
' O parecer e opcional na assinatura para nao quebrar chamadores existentes; a
' Sprint NC passa a exigi-lo na interface. Toda exclusao e auditada aqui, dentro
' da camada de dados - nenhum caminho de gravacao escapa (item 3.2).
Public Function ExcluirLogico(ByVal run As Long, ByVal nivel As Long, ByRef alvoS As Object, _
                              Optional ByVal parecer As String = "", _
                              Optional ByVal novoStatus As String = "", _
                              Optional ByVal motivo As String = "", _
                              Optional ByVal origemLog As String = "Resultados") As Long
    Dim ws As Worksheet, dados As Variant, i As Long, n As Long, lin As Long
    Dim stAntes As String, stNovo As String
    Set ws = ThisWorkbook.Sheets(BANCO)
    dados = CarregarDB()
    If IsEmpty(dados) Then ExcluirLogico = 0: Exit Function
    stNovo = Trim$(novoStatus)
    If stNovo = "" Then stNovo = ST_EXCLUIDO
    Application.ScreenUpdating = False
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
                If alvoS.Exists(UCase$(Trim$(CStr(dados(i, COL_ANALITO))))) Then
                    lin = BANCO_R0 + i - 1
                    stAntes = Trim$(CStr(ws.Cells(lin, COL_STATUS).Value))
                    ws.Cells(lin, COL_STATUS).Value = stNovo
                    ' Duas camadas, um identificador so: o Event Store prova a
                    ' integridade, a tabela do banco preserva a estrutura por
                    ' origem. Compartilham o ID para que divergencia entre elas
                    ' seja detectavel.
                    Dim idEv As String
                    idEv = Auditar(CAT_DADO, AC_EXCLUSAO, "mDados", _
                            run, dados(i, COL_DATA), "", CStr(dados(i, COL_LOTE)), _
                            nivel, CStr(dados(i, COL_ANALITO)), _
                            dados(i, COL_RESULT), dados(i, COL_RESULT), _
                            stAntes, stNovo, motivo, parecer)
                    RegistrarLogDB origemLog, idEv, AC_EXCLUSAO, _
                            run, dados(i, COL_DATA), nivel, CStr(dados(i, COL_ANALITO)), _
                            CStr(dados(i, COL_LOTE)), dados(i, COL_RESULT), _
                            stAntes, stNovo, parecer
                    n = n + 1
                End If
            End If
        End If
    Next i

    ' Excluir logicamente PROMOVE a proxima duplicata a "primeira ativa", entao a
    ' flag muda para linhas que nao foram tocadas. Recalcular aqui tambem (ADR-025).
    AtualizarFlagsBanco

    Application.ScreenUpdating = True
    ExcluirLogico = n
End Function

' Lista de RUNs distintos ATIVOS do lote em uso (para os combos dos formularios).
Public Function RunsDoLote(ByVal loteCore As String) As Collection
    Dim dados As Variant, i As Long, seen As Object, c As Collection
    Set c = New Collection: Set seen = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If IsEmpty(dados) Then Set RunsDoLote = c: Exit Function
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
            If NucleoLote(CStr(dados(i, COL_LOTE))) = loteCore Then
                If Not seen.Exists(CStr(dados(i, COL_RUN))) Then
                    seen.Add CStr(dados(i, COL_RUN)), 1
                    c.Add CLng(dados(i, COL_RUN))
                End If
            End If
        End If
    Next i
    Set RunsDoLote = c
End Function

' Nomes dos analitos cadastrados (aba Analitos).
Public Function ListaAnalitos() As Collection
    Dim ws As Worksheet, i As Long, nm As String, c As Collection
    Set c = New Collection
    Set ws = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        nm = Trim$(CStr(ws.Cells(i, 1).Value))
        If nm <> "" Then c.Add nm
    Next i
    Set ListaAnalitos = c
End Function

' Codigos completos de lote disponiveis no registro (Configuracao).
Public Function ListaLotes() As Collection
    Dim rng As Range, cel As Range, c As Collection, v As String
    Set c = New Collection
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    If rng Is Nothing Then Set ListaLotes = c: Exit Function
    For Each cel In rng
        v = Trim$(CStr(cel.Value))
        If v <> "" Then c.Add v
    Next cel
    Set ListaLotes = c
End Function

Public Sub AtualizarBanco()
    Application.Calculate
End Sub

' RegistrarLog SAIU DAQUI e vive em mAuditoria.bas.
'
' Dois procedimentos Public com o mesmo nome em modulos diferentes produzem
' "nome ambiguo" e derrubam a compilacao do PROJETO INTEIRO -- a mesma classe de
' falha do identificador aS, com o mesmo sintoma enganoso: o erro aparece na
' primeira rotina chamada, longe da causa. A trilha de auditoria tem um dono so.

