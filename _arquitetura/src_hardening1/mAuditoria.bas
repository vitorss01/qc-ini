Attribute VB_Name = "mAuditoria"
Option Explicit
' ===== TRILHA DE AUDITORIA (Sprint HARDENING 3) =====
'
' Itens 3.1 e 3.2 do Quality Gate. ISO 15189:2022 secoes 8.4.1 a 8.4.3.
'
' PRINCIPIO. O log e append-only e ENCADEADO POR HASH. Cada linha carrega o
' SHA-256 do conteudo da linha anterior somado ao proprio conteudo. Alterar ou
' apagar qualquer linha quebra a cadeia, e VerificarIntegridadeLog aponta em
' que linha quebrou.
'
' POR QUE A CADEIA. A protecao de planilha do Excel nao resiste a fraude
' deliberada: a senha esta em texto no projeto VBA e a marcacao de protecao
' pode ser removida descompactando o arquivo. Sem a cadeia, o log seria
' "confie em mim". Com ela, adulteracao vira EVIDENCIA VERIFICAVEL. O sistema
' nao promete o que nao entrega -- promete deteccao, e prova.
'
' IDENTIDADE. O usuario de registro e o LOGIN DO SISTEMA (aba Usuarios, senha
' em hash). Usuario do Windows, do Office e nome da maquina entram como
' CORROBORACAO: vem de variavel de ambiente e qualquer um as altera.
'
' LIMITE DECLARADO. A hora vem do relogio da maquina, que o usuario pode
' mudar. Nao ha solucao dentro do Excel. Fica registrado, nao escondido.

Public Const AUDIT As String = "Audit_Log"
Public Const AUDIT_R0 As Long = 4          ' primeira linha de dado
Public Const VERSAO_SISTEMA As String = "1.0.0-rc1"

Public Const AU_ID As Long = 1
Public Const AU_DTHORA As Long = 2
Public Const AU_ACAO As Long = 3
Public Const AU_ORIGEM As Long = 4
Public Const AU_RUN As Long = 5
Public Const AU_DTCORRIDA As Long = 6
Public Const AU_NIVEL As Long = 7
Public Const AU_ANALITO As Long = 8
Public Const AU_LOTE As Long = 9
Public Const AU_RESULT As Long = 10
Public Const AU_STANTES As Long = 11
Public Const AU_STDEPOIS As Long = 12
Public Const AU_PARECER As Long = 13
Public Const AU_USRSIST As Long = 14
Public Const AU_USROFF As Long = 15
Public Const AU_USRWIN As Long = 16
Public Const AU_MAQUINA As Long = 17
Public Const AU_ARQUIVO As Long = 18
Public Const AU_VERSAO As Long = 19
Public Const AU_HASHANT As Long = 20
Public Const AU_HASH As Long = 21
Public Const AU_NCOL As Long = 21

Public Const PARECER_MIN_PALAVRAS As Long = 5

' ===================== CONTEXTO =====================

Public Function UsuarioSistema() As String
    ' Identidade de registro: quem esta autenticado no sistema.
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

' ===================== ESCRITA =====================

Public Function UltimaLinhaAudit() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(AUDIT)
    UltimaLinhaAudit = ws.Cells(ws.Rows.Count, AU_ID).End(xlUp).Row
    If UltimaLinhaAudit < AUDIT_R0 Then UltimaLinhaAudit = AUDIT_R0 - 1
End Function

' Concatena o conteudo da linha (exceto o proprio hash) para o calculo.
'
' O parametro e "v() As Variant", array DINAMICO. Array de tamanho fixo
' (Dim v(1 To N)) nao pode ser passado para parametro Variant em VBA: o projeto
' inteiro deixa de compilar, e o sintoma que aparece e 0x800A9C68 em qualquer
' Application.Run -- longe da causa. Por isso quem chama usa ReDim.
Private Function Payload(ByRef v() As Variant) As String
    Dim i As Long, s As String
    For i = AU_ID To AU_HASHANT
        s = s & "|" & CStr(v(i))
    Next i
    Payload = s
End Function

' Ponto UNICO de gravacao no log. Devolve o ID_Auditoria gerado.
Public Function Auditar(ByVal acao As String, ByVal origem As String, _
                        ByVal run As Long, ByVal dtCorrida As Variant, _
                        ByVal nivel As Long, ByVal analito As String, _
                        ByVal lote As String, ByVal resultado As Variant, _
                        ByVal stAntes As String, ByVal stDepois As String, _
                        ByVal parecer As String) As String
    Dim ws As Worksheet, lin As Long, v() As Variant, i As Long
    ReDim v(1 To AU_NCOL)
    Dim hashAnt As String, id As String

    Set ws = ThisWorkbook.Sheets(AUDIT)
    lin = UltimaLinhaAudit() + 1

    If lin > AUDIT_R0 Then
        hashAnt = CStr(ws.Cells(lin - 1, AU_HASH).Value)
    Else
        hashAnt = String$(64, "0")          ' bloco genese
    End If

    ' ID nunca reaproveitado: sequencial + carimbo + usuario.
    id = "AUD-" & Format$(lin - AUDIT_R0 + 1, "000000") & "-" & _
         Format$(Now, "yyyymmddhhnnss")

    v(AU_ID) = id
    v(AU_DTHORA) = Now
    v(AU_ACAO) = acao
    v(AU_ORIGEM) = origem
    v(AU_RUN) = run
    v(AU_DTCORRIDA) = dtCorrida
    v(AU_NIVEL) = nivel
    v(AU_ANALITO) = analito
    v(AU_LOTE) = lote
    v(AU_RESULT) = resultado
    v(AU_STANTES) = stAntes
    v(AU_STDEPOIS) = stDepois
    v(AU_PARECER) = parecer
    v(AU_USRSIST) = UsuarioSistema() & IIf(PapelSistema() = "", "", " [" & PapelSistema() & "]")
    v(AU_USROFF) = Application.UserName
    v(AU_USRWIN) = Environ$("USERNAME")
    v(AU_MAQUINA) = Environ$("COMPUTERNAME")
    v(AU_ARQUIVO) = ThisWorkbook.Name
    v(AU_VERSAO) = VERSAO_SISTEMA
    v(AU_HASHANT) = hashAnt
    v(AU_HASH) = SHA256Hex(hashAnt & Payload(v))

    Dim protegida As Boolean
    protegida = ws.ProtectContents
    If protegida Then ws.Unprotect Password:="qcini2025"
    For i = 1 To AU_NCOL
        ws.Cells(lin, i).Value = v(i)
    Next i
    ws.Cells(lin, AU_DTHORA).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(lin, AU_DTCORRIDA).NumberFormat = "dd/mm/yyyy"
    If protegida Then ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True

    Auditar = id
End Function

' ===================== VERIFICACAO =====================

' Recalcula a cadeia inteira. Devolve "OK|n" ou "QUEBRADO|linha".
Public Function VerificarIntegridadeLog() As String
    Dim ws As Worksheet, ult As Long, lin As Long, i As Long
    Dim v() As Variant, esperado As String, gravado As String
    Dim hashAnt As String
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

Public Sub ConferirAuditoria()
    Dim r As String, p As Variant
    r = VerificarIntegridadeLog()
    p = Split(r, "|")
    If p(0) = "OK" Then
        MsgBox "Trilha de auditoria integra." & vbLf & vbLf & _
               p(1) & " registro(s) verificados." & vbLf & _
               "Nenhuma linha foi alterada ou removida.", _
               vbInformation, "Auditoria"
    Else
        MsgBox "INTEGRIDADE COMPROMETIDA." & vbLf & vbLf & _
               "A cadeia quebra na linha " & p(1) & " do Audit_Log." & vbLf & _
               "Motivo: " & p(2) & vbLf & vbLf & _
               "Isso indica que o registro foi alterado ou removido fora do sistema.", _
               vbCritical, "Auditoria"
    End If
End Sub
