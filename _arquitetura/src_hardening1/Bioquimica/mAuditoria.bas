Attribute VB_Name = "mAuditoria"
Option Explicit
' ===== TRILHA DE AUDITORIA (itens 3.1 e 3.2) =====
'
' ISO 15189:2022 secoes 8.4.1 a 8.4.3. Pensado tambem para CAP e DAIDS.
'
' NAO E UM ARQUIVO TECNICO -- E UMA BASE DE AUDITORIA. O auditor abre a aba
' Audit_Log e responde tudo com os filtros nativos do Excel, sem conhecer a
' estrutura interna do sistema e sem depender de nenhuma tela em VBA. Por isso
' a aba e uma TABELA do Excel (ListObject "tblAuditoria"): filtros acompanham
' as linhas novas, e Power Query e Power BI consomem pelo nome.
'
' UMA LINHA = UM EVENTO. Nunca se edita evento anterior. Nunca se exclui.
' Append-only.
'
' ENCADEADA POR HASH. Cada linha carrega o SHA-256 da anterior somada ao
' proprio conteudo. A protecao de planilha do Excel nao resiste a fraude
' deliberada -- a senha esta em texto no projeto VBA e a marcacao sai
' descompactando o arquivo. Sem a cadeia, o log seria "confie em mim". Com ela,
' adulteracao vira EVIDENCIA VERIFICAVEL: VerificarIntegridadeLog aponta a
' linha exata onde a cadeia quebrou.
'
' IDENTIDADE. Usuario de registro = LOGIN DO SISTEMA (aba Usuarios, senha em
' hash). Usuario do Windows, do Office e nome da maquina entram como
' CORROBORACAO: vem de variavel de ambiente e qualquer um as altera.
'
' LIMITE DECLARADO. A hora vem do relogio da maquina, que o usuario pode
' mudar. Nao ha solucao dentro do Excel. Fica registrado, nao escondido.
'
' VERSAO DE SCHEMA. Mudar o layout muda o que a cadeia cobre. VersaoSchema
' marca o formato de cada linha para que um verificador futuro saiba como
' recalcular. Cadeias de schemas diferentes nao se misturam.

Public Const AUDIT As String = "Audit_Log"
Public Const AUDIT_TAB As String = "tblAuditoria"
Public Const AUDIT_R0 As Long = 4              ' primeira linha de dado
Public Const AUDIT_SCHEMA As Long = 2          ' versao do layout
Public Const VERSAO_SISTEMA As String = "1.0.0-rc1"

' --- colunas -----------------------------------------------------------------
Public Const AU_ID As Long = 1
Public Const AU_SCHEMA As Long = 2
Public Const AU_TS As Long = 3                 ' timestamp completo
Public Const AU_DATA As Long = 4               ' so a data  (filtro por periodo)
Public Const AU_HORA As Long = 5               ' so a hora  (filtro por faixa)
Public Const AU_CATEG As Long = 6              ' DADO | CONFIGURACAO | SEGURANCA | SISTEMA
Public Const AU_ACAO As Long = 7
Public Const AU_MODULO As Long = 8
Public Const AU_CHAVE As Long = 9              ' RUN|Nivel|Analito
Public Const AU_RUN As Long = 10
Public Const AU_DTCORRIDA As Long = 11
Public Const AU_EQUIP As Long = 12
Public Const AU_LOTE As Long = 13
Public Const AU_NIVEL As Long = 14
Public Const AU_ANALITO As Long = 15
Public Const AU_SEQ As Long = 16               ' 1a, 2a, 3a alteracao daquela chave
Public Const AU_RESANT As Long = 17
Public Const AU_RESNOVO As Long = 18
Public Const AU_DELTA As Long = 19
Public Const AU_DELTAP As Long = 20
Public Const AU_STANT As Long = 21
Public Const AU_STNOVO As Long = 22
Public Const AU_MOTIVO As Long = 23            ' lista fechada -- da para contar
Public Const AU_PARECER As Long = 24           ' texto livre -- da para entender
Public Const AU_USRSIST As Long = 25
Public Const AU_PAPEL As Long = 26
Public Const AU_USROFF As Long = 27
Public Const AU_USRWIN As Long = 28
Public Const AU_MAQUINA As Long = 29
Public Const AU_ARQUIVO As Long = 30
Public Const AU_VERSAO As Long = 31
Public Const AU_HASHANT As Long = 32
Public Const AU_HASH As Long = 33
Public Const AU_NCOL As Long = 33

' --- vocabulario controlado --------------------------------------------------
' Categorias
Public Const CAT_DADO As String = "DADO"
Public Const CAT_CONFIG As String = "CONFIGURACAO"
Public Const CAT_SEG As String = "SEGURANCA"
Public Const CAT_SIS As String = "SISTEMA"

' Acoes. Uma acao por situacao: dois nomes para a mesma coisa fazem o filtro
' perder registro.
Public Const AC_INCLUSAO As String = "RESULTADO_INCLUIDO"
Public Const AC_ALTERACAO As String = "RESULTADO_ALTERADO"
Public Const AC_APAGADO As String = "RESULTADO_APAGADO"
Public Const AC_EXCLUSAO As String = "RESULTADO_EXCLUIDO"
Public Const AC_BLOQUEIO As String = "REENVIO_BLOQUEADO"
Public Const AC_STATUS As String = "STATUS_ALTERADO"

Public Const VAZIO As String = "<VAZIO>"
Public Const PARECER_MIN_PALAVRAS As Long = 5

' Cache do contador de alteracoes por chave.
'
' ProximoSeq varria o log INTEIRO a cada gravacao: custo O(n) por evento e
' O(n^2) ao longo do dia. Com o cache, a varredura acontece uma vez por sessao
' e cada evento seguinte e O(1). Medido: era o maior componente dos 403ms/evento.
Private mSeq As Object

' ===================== CONTEXTO =====================

Public Function UsuarioSistema() As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(ThisWorkbook.Names("currentUser").RefersToRange.Value))
    On Error GoTo 0
    If v = "" Then v = "(sem login)"
    UsuarioSistema = v
End Function

Public Function PapelSistema() As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(ThisWorkbook.Names("currentPapel").RefersToRange.Value))
    On Error GoTo 0
    If v = "" Then v = "(sem papel)"
    PapelSistema = v
End Function

' ===================== PARECER TECNICO =====================

' Conta palavras REAIS: espacos multiplos, tabulacoes e quebras nao inflam a
' contagem, e token sem letra nem digito nao vale como palavra.
Public Function ContarPalavras(ByVal texto As String) As Long
    Dim s As String, partes As Variant, i As Long, n As Long, t As String
    s = Replace$(Replace$(Replace$(texto, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    s = Trim$(s)
    If s = "" Then ContarPalavras = 0: Exit Function
    partes = Split(s, " ")
    For i = LBound(partes) To UBound(partes)
        t = Trim$(CStr(partes(i)))
        If TemAlfanumerico(t) Then n = n + 1
    Next i
    ContarPalavras = n
End Function

Private Function TemAlfanumerico(ByVal t As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(t)
        c = UCase$(Mid$(t, i, 1))
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then
            TemAlfanumerico = True: Exit Function
        End If
    Next i
End Function

Public Function ParecerValido(ByVal parecer As String) As Boolean
    ParecerValido = (ContarPalavras(parecer) >= PARECER_MIN_PALAVRAS)
End Function

' ===================== APOIO =====================

' Ultima linha COM EVENTO.
'
' End(xlUp) sozinho nao serve: um ListObject nasce com uma LINHA DE DADO EM
' BRANCO, e essa linha fantasma nao e vazia para o Excel -- tem formatacao de
' tabela. O resultado era o primeiro evento cair na linha 5 em vez da 4,
' deixando a 4 vazia e quebrando o elo genese da cadeia logo na primeira
' verificacao.
Public Function UltimaLinhaAudit() As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets(AUDIT)
    r = ws.Cells(ws.Rows.Count, AU_ID).End(xlUp).Row
    Do While r >= AUDIT_R0
        If Len(Trim$(CStr(ws.Cells(r, AU_ID).Value))) > 0 Then Exit Do
        r = r - 1
    Loop
    If r < AUDIT_R0 Then r = AUDIT_R0 - 1
    UltimaLinhaAudit = r
End Function

' Rotulo do valor: distingue "nao se aplica" de "foi apagado". Sem isso, uma
' celula limpa e um campo sem valor ficam identicos no log -- e apagar
' resultado e exatamente o que a auditoria mais quer enxergar.
' Rotulo NAO pode devolver String para valor numerico.
'
' Devolvendo String, um resultado 92,0028 virava "92,0028" e, gravado numa
' celula de formato Geral, o Excel reinterpretava a virgula como separador de
' MILHAR: a celula passava a conter 920028. O valor anterior e o novo ficavam
' multiplicados por 10.000 no Audit_Log.
'
' Passava despercebido porque AU_DELTA e gravado como NUMERO e continuava
' certo -- um delta correto ao lado de dois valores errados nao chama atencao.
' Pior: o hash da cadeia e calculado sobre o que se RELE da celula, entao a
' verificacao de integridade APROVAVA o registro corrompido. Trilha que prova
' nao-adulteracao de dado errado e a pior combinacao possivel para auditoria.
'
' ISO 15189 8.4.2 exige o valor original recuperavel apos a emenda.
Private Function Rotulo(ByVal v As Variant) As Variant
    If IsEmpty(v) Or IsNull(v) Then
        Rotulo = ""
    ElseIf VarType(v) = vbDate Then
        Rotulo = CStr(v)
    ElseIf IsNumeric(v) Then
        ' numero permanece NUMERO: sem CStr, sem reinterpretacao de separador
        Rotulo = CDbl(v)
    ElseIf Trim$(CStr(v)) = "" Then
        Rotulo = VAZIO
    Else
        Rotulo = CStr(v)
    End If
End Function

' Quantas vezes esta chave ja apareceu no log. O valor ORIGINAL de um registro
' e sempre o da linha com Seq = 1.
Private Function ProximoSeq(ByVal chave As String) As Long
    Dim ws As Worksheet, i As Long, ult As Long, dados As Variant, k As String
    If chave = "" Then ProximoSeq = 0: Exit Function

    If mSeq Is Nothing Then
        ' Uma varredura por sessao, em BLOCO (nao celula a celula), e o mapa
        ' fica em memoria para todos os eventos seguintes.
        Set mSeq = CreateObject("Scripting.Dictionary")
        mSeq.CompareMode = 1
        Set ws = ThisWorkbook.Sheets(AUDIT)
        ult = UltimaLinhaAudit()
        If ult >= AUDIT_R0 Then
            dados = ws.Range(ws.Cells(AUDIT_R0, AU_CHAVE), ws.Cells(ult, AU_CHAVE)).Value
            If IsArray(dados) Then
                For i = 1 To UBound(dados, 1)
                    k = Trim$(CStr(dados(i, 1)))
                    If k <> "" Then
                        If mSeq.Exists(k) Then mSeq(k) = mSeq(k) + 1 Else mSeq(k) = 1
                    End If
                Next i
            End If
        End If
    End If

    If mSeq.Exists(chave) Then
        mSeq(chave) = mSeq(chave) + 1
    Else
        mSeq(chave) = 1
    End If
    ProximoSeq = mSeq(chave)
End Function

Private Function Payload(ByRef v() As Variant) As String
    Dim i As Long, s As String
    For i = AU_ID To AU_HASHANT
        s = s & "|" & CStr(v(i))
    Next i
    Payload = s
End Function

' ===================== ESCRITA =====================

' Ponto UNICO de gravacao no log. Devolve o ID do evento.
Public Function Auditar(ByVal categoria As String, ByVal acao As String, _
                        ByVal modulo As String, _
                        ByVal run As Long, ByVal dtCorrida As Variant, _
                        ByVal equip As String, ByVal lote As String, _
                        ByVal nivel As Long, ByVal analito As String, _
                        ByVal resAnt As Variant, ByVal resNovo As Variant, _
                        ByVal stAnt As String, ByVal stNovo As String, _
                        ByVal motivo As String, ByVal parecer As String) As String

    Dim ws As Worksheet, lin As Long, v() As Variant, i As Long
    Dim hashAnt As String, id As String, chave As String, agora As Date
    ReDim v(1 To AU_NCOL)

    Set ws = ThisWorkbook.Sheets(AUDIT)
    lin = UltimaLinhaAudit() + 1
    agora = Now

    If lin > AUDIT_R0 Then
        hashAnt = CStr(ws.Cells(lin - 1, AU_HASH).Value)
    Else
        hashAnt = String$(64, "0")                  ' bloco genese
    End If

    If Trim$(analito) <> "" Then
        chave = CStr(run) & "|" & CStr(nivel) & "|" & UCase$(Trim$(analito))
    End If

    id = "AUD-" & Format$(lin - AUDIT_R0 + 1, "000000") & "-" & Format$(agora, "yyyymmddhhnnss")

    v(AU_ID) = id
    v(AU_SCHEMA) = AUDIT_SCHEMA
    v(AU_TS) = agora
    v(AU_DATA) = DateSerial(Year(agora), Month(agora), Day(agora))
    v(AU_HORA) = TimeSerial(Hour(agora), Minute(agora), Second(agora))
    v(AU_CATEG) = categoria
    v(AU_ACAO) = acao
    v(AU_MODULO) = modulo
    v(AU_CHAVE) = chave
    v(AU_RUN) = IIf(run = 0, "", run)
    v(AU_DTCORRIDA) = dtCorrida
    v(AU_EQUIP) = equip
    v(AU_LOTE) = lote
    v(AU_NIVEL) = IIf(nivel = 0, "", nivel)
    v(AU_ANALITO) = analito
    v(AU_SEQ) = IIf(chave = "", "", ProximoSeq(chave))
    v(AU_RESANT) = Rotulo(resAnt)
    v(AU_RESNOVO) = Rotulo(resNovo)

    ' Delta so quando os dois lados sao numero. Comparar texto com numero, ou
    ' calcular diferenca contra vazio, produziria valor sem significado.
    If IsNumeric(resAnt) And IsNumeric(resNovo) And Trim$(CStr(resAnt)) <> "" And Trim$(CStr(resNovo)) <> "" Then
        v(AU_DELTA) = CDbl(resNovo) - CDbl(resAnt)
        If CDbl(resAnt) <> 0 Then
            v(AU_DELTAP) = (CDbl(resNovo) - CDbl(resAnt)) / Abs(CDbl(resAnt))
        Else
            v(AU_DELTAP) = ""
        End If
    Else
        v(AU_DELTA) = ""
        v(AU_DELTAP) = ""
    End If

    v(AU_STANT) = stAnt
    v(AU_STNOVO) = stNovo
    v(AU_MOTIVO) = motivo
    v(AU_PARECER) = parecer
    v(AU_USRSIST) = UsuarioSistema()
    v(AU_PAPEL) = PapelSistema()
    v(AU_USROFF) = Application.UserName
    v(AU_USRWIN) = Environ$("USERNAME")
    v(AU_MAQUINA) = Environ$("COMPUTERNAME")
    v(AU_ARQUIVO) = ThisWorkbook.Name
    v(AU_VERSAO) = VERSAO_SISTEMA
    v(AU_HASHANT) = hashAnt

    ' ADR-046: a janela destrancada tem de fechar tambem no caminho de ERRO.
    ' O Audit_Log guarda a cadeia de hash; deixa-lo aberto porque uma excecao
    ' passou por aqui nao e susto, e perda de evidencia -- e Auditar e chamada
    ' justamente a partir de caminhos de erro.
    '
    ' Reprotecao por ProtegerAudit, e nao pelo RestaurarProtecao generico: o
    ' auditor precisa de AllowFiltering/AllowSorting, que o generico nao
    ' concede. Trocar um pelo outro reduziria a permissao sem ninguem notar.
    Dim protegida As Boolean
    On Error GoTo restauraAudit
    protegida = LiberarEscrita(ws)
    ' Texto ANTES de gravar: um hash de 64 digitos que por acaso seja so
    ' numeros viraria notacao cientifica, e a verificacao acusaria adulteracao
    ' onde nao houve. O mesmo vale para o ID.
    ws.Cells(lin, AU_ID).NumberFormat = "@"
    ws.Cells(lin, AU_HASHANT).NumberFormat = "@"
    ws.Cells(lin, AU_HASH).NumberFormat = "@"

    ' UMA escrita de bloco em vez de 32 escritas de celula. Cada acesso a
    ' Range custa uma travessia COM; em lote, o Excel resolve tudo de uma vez.
    Dim bloco() As Variant
    ReDim bloco(1 To 1, 1 To AU_HASHANT)
    For i = AU_ID To AU_HASHANT
        bloco(1, i) = v(i)
    Next i
    ws.Range(ws.Cells(lin, AU_ID), ws.Cells(lin, AU_HASHANT)).Value = bloco

    ' O HASH COBRE O QUE FICOU GRAVADO, NAO O QUE SE PRETENDIA GRAVAR.
    '
    ' Calcular sobre o array em memoria e depois verificar lendo as celulas
    ' compara duas coisas diferentes: o Excel converte na viagem -- hora vira
    ' fracao de dia, numero vira texto, ponto flutuante perde bit. Qualquer
    ' dessas diferencas acusaria adulteracao onde nao houve, e um verificador
    ' que da alarme falso e pior do que nenhum: ninguem confia nele quando
    ' importa. Por isso grava-se primeiro, LE-SE DE VOLTA, e so entao se calcula.
    Dim g() As Variant, lido As Variant
    ReDim g(1 To AU_NCOL)
    lido = ws.Range(ws.Cells(lin, AU_ID), ws.Cells(lin, AU_HASHANT)).Value
    For i = AU_ID To AU_HASHANT
        g(i) = lido(1, i)
    Next i
    ws.Cells(lin, AU_HASH).Value = SHA256Hex(hashAnt & Payload(g))
    ws.Cells(lin, AU_TS).NumberFormatLocal = "dd/mm/aaaa hh:mm:ss;@"
    ws.Cells(lin, AU_DATA).NumberFormatLocal = "dd/mm/aaaa;@"
    ws.Cells(lin, AU_HORA).NumberFormat = "hh:mm:ss"
    ws.Cells(lin, AU_DTCORRIDA).NumberFormatLocal = "dd/mm/aaaa;@"
    ws.Cells(lin, AU_DELTAP).NumberFormat = "0.0%"
    ExpandirTabela ws, lin
    If protegida Then ProtegerAudit ws
    On Error GoTo 0

    Auditar = id
    Exit Function

restauraAudit:
    Dim nErrA As Long, sErrA As String
    nErrA = Err.Number: sErrA = Err.Description
    On Error Resume Next
    If protegida Then ProtegerAudit ws
    On Error GoTo 0
    Err.Raise nErrA, "mAuditoria.Auditar", sErrA
End Function

' A aba e protegida, mas o auditor PRECISA filtrar e ordenar. Sem estas
' permissoes a tabela vira um bloco morto e o objetivo se perde.
Public Sub ProtegerAudit(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
               DrawingObjects:=False, Contents:=True, Scenarios:=True, _
               AllowFiltering:=True, AllowSorting:=True, AllowFormattingColumns:=True
End Sub

' Mantem o ListObject cobrindo a linha nova: sem isso o filtro nao alcanca os
' eventos recentes -- justamente os que o auditor procura primeiro.
Private Sub ExpandirTabela(ByVal ws As Worksheet, ByVal lin As Long)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(AUDIT_TAB)
    If lo Is Nothing Then Exit Sub
    If lin > lo.Range.Row + lo.Range.Rows.Count - 1 Then
        lo.Resize ws.Range(ws.Cells(AUDIT_R0 - 1, 1), ws.Cells(lin, AU_NCOL))
    End If
End Sub

' Compatibilidade: assinatura antiga continua valida.
Public Sub RegistrarLog(ByVal acao As String, ByVal detalhe As String)
    Auditar CAT_SIS, acao, "", 0, Empty, "", "", 0, "", Empty, Empty, "", "", "", detalhe
End Sub

' ===================== CONSULTA =====================

' Historia completa de um resultado: valor original, cada alteracao, quem, quando
' e por que. E o que responde "qual era o valor antes da emenda?" da ISO 15189
' 8.4.2 sem ninguem precisar garimpar o log.
Public Function HistoricoDoRegistro(ByVal run As Long, ByVal nivel As Long, _
                                    ByVal analito As String) As String
    Dim ws As Worksheet, i As Long, ult As Long, chave As String, r As String, n As Long
    chave = CStr(run) & "|" & CStr(nivel) & "|" & UCase$(Trim$(analito))
    Set ws = ThisWorkbook.Sheets(AUDIT)
    ult = UltimaLinhaAudit()
    For i = AUDIT_R0 To ult
        If CStr(ws.Cells(i, AU_CHAVE).Value) = chave Then
            n = n + 1
            r = r & Format$(ws.Cells(i, AU_TS).Value, "dd/mm/yyyy hh:mm") & "  " & _
                CStr(ws.Cells(i, AU_ACAO).Value) & "  " & _
                CStr(ws.Cells(i, AU_RESANT).Value) & " -> " & CStr(ws.Cells(i, AU_RESNOVO).Value) & _
                "  [" & CStr(ws.Cells(i, AU_USRSIST).Value) & "]" & vbLf
        End If
    Next i
    If n = 0 Then
        HistoricoDoRegistro = "Nenhum evento registrado para " & chave
    Else
        HistoricoDoRegistro = chave & " - " & n & " evento(s):" & vbLf & r
    End If
End Function

' ===================== VERIFICACAO =====================

Public Function VerificarIntegridadeLog() As String
    Dim ws As Worksheet, ult As Long, lin As Long, i As Long
    Dim v() As Variant, esperado As String, gravado As String, hashAnt As String
    ReDim v(1 To AU_NCOL)

    Set ws = ThisWorkbook.Sheets(AUDIT)
    ult = UltimaLinhaAudit()
    If ult < AUDIT_R0 Then VerificarIntegridadeLog = "OK|0": Exit Function

    hashAnt = String$(64, "0")
    For lin = AUDIT_R0 To ult
        For i = 1 To AU_NCOL
            v(i) = ws.Cells(lin, i).Value
        Next i
        If CStr(v(AU_HASHANT)) <> hashAnt Then
            VerificarIntegridadeLog = "QUEBRADO|" & lin & "|elo anterior nao confere"
            Exit Function
        End If
        esperado = SHA256Hex(hashAnt & Payload(v))
        gravado = LCase$(Trim$(CStr(v(AU_HASH))))
        If LCase$(esperado) <> gravado Then
            VerificarIntegridadeLog = "QUEBRADO|" & lin & "|conteudo alterado"
            Exit Function
        End If
        hashAnt = CStr(v(AU_HASH))
    Next lin
    VerificarIntegridadeLog = "OK|" & (ult - AUDIT_R0 + 1)
End Function

' Chamada quando o log e recriado ou quando a sessao precisa reler do zero.
Public Sub InvalidarCacheAuditoria()
    Set mSeq = Nothing
End Sub

Public Sub ConferirAuditoria()
    Dim r As String, p As Variant
    r = VerificarIntegridadeLog()
    p = Split(r, "|")
    If p(0) = "OK" Then
        MsgBox "Trilha de auditoria integra." & vbLf & vbLf & _
               p(1) & " evento(s) verificados." & vbLf & _
               "Nenhuma linha foi alterada ou removida.", vbInformation, "Auditoria"
    Else
        MsgBox "INTEGRIDADE COMPROMETIDA." & vbLf & vbLf & _
               "A cadeia quebra na linha " & p(1) & " do Audit_Log." & vbLf & _
               "Motivo: " & p(2) & vbLf & vbLf & _
               "Isso indica alteracao ou remocao feita fora do sistema.", vbCritical, "Auditoria"
    End If
End Sub
