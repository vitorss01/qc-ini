Attribute VB_Name = "mCalibracao"
Option Explicit
' ============================================================================
'  mCalibracao -- registrar a calibracao pelo ID do resultado (ADR-084)
'
'  Aba Registros (lote em uso; cada lote guardado na RegistrosStore pelo mLotes),
'  linhas 4..203. O usuario digita na coluna K o ID de um resultado de controle
'  (ex.: BIO-365293, ou so o numero) -- o resultado medido logo depois da
'  calibracao. O sistema:
'    1. valida o formato e acha o resultado EXATO (mIntegracao.DadosDoId: a mesma
'       busca do inativar; nunca aproximada);
'    2. recusa ID ja registrado em outra linha (nada e gravado em silencio);
'    3. preenche a linha: B data/hora do resultado, C analito, D "Todos" (a
'       calibracao e do analito no equipamento, vale para todos os niveis), E RUN
'       do resultado, F "Sim", H usuario da sessao, J status, L equipamento;
'    4. registra no Audit_Log (CALIBRACAO_REGISTRADA / CALIBRACAO_RECUSADA).
'
'  O GRAFICO nao muda de regra: o icone "Calibracao" continua na MEDIA DO LOTE do
'  nivel (Calc AX1/AZ1/BT1...), nas corridas do mesmo dia e analito com
'  "Foi calibrado? = Sim" -- a mesma regra da marcacao manual. O ADR-084 so
'  acrescenta o equipamento: com L preenchido, o icone sai so no grafico daquele
'  equipamento; linha sem equipamento (as antigas) vale para todos, como antes.
' ============================================================================
Private Const ABA As String = "Registros"
Private Const R0 As Long = 4
Private Const R1 As Long = 203
Private Const C_DATA As Long = 2
Private Const C_ANALITO As Long = 3
Private Const C_NIVEL As Long = 4
Private Const C_CORRIDA As Long = 5
Private Const C_CALIB As Long = 6
Private Const C_RESP As Long = 8
Private Const C_STATUS As Long = 10
Private Const C_ID As Long = 11
Private Const C_EQUIP As Long = 12

' Nucleo sem dialogo (o QA chama este). linha = a linha da aba onde o ID foi digitado (0 = a 1a linha vazia).
' Devolve "OK|<linha>|<analito>|<equipamento>|<data/hora>" ou "ERRO|<mensagem>". Nada e gravado no ERRO.
Public Function RegistrarCalibracaoPorId(ByVal id As String, Optional ByVal linha As Long = 0) As String
    Dim d As String, p As Variant, ws As Worksheet, r As Long, dup As Long, prot As Boolean, ev As Boolean
    Dim dh As Date, sE As String, quando As String
    On Error GoTo falha
    d = mIntegracao.DadosDoId(id)
    If Left$(d, 3) <> "OK|" Then
        RegistrarCalibracaoPorId = d
        Auditar "CALIBRACAO_RECUSADA", "ID informado: " & Left$(Trim$(CStr(id)), 40) & " | " & Mid$(d, 6)
        Exit Function
    End If
    p = Split(d, "|")                        ' OK|ID|analito|nivel|equip|lote|RUN|data|estado
    Set ws = ThisWorkbook.Worksheets(ABA)
    dup = LinhaDoId(ws, CStr(p(1)), linha)
    If dup > 0 Then
        RegistrarCalibracaoPorId = "ERRO|O ID " & mUI.IdCurto(CStr(p(1))) & " já está registrado como calibração " & _
                                   "(aba Registros, linha " & dup & "). Nada foi gravado."
        Auditar "CALIBRACAO_RECUSADA", "ID " & p(1) & " repetido (ja na linha " & dup & ")"
        Exit Function
    End If
    If linha < R0 Or linha > R1 Then linha = LinhaLivre(ws)
    If linha = 0 Then
        RegistrarCalibracaoPorId = "ERRO|A aba Registros está cheia (200 linhas neste lote). Nada foi gravado."
        Exit Function
    End If
    r = linha
    dh = CDate(CDbl(p(7)))
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    ws.Cells(r, C_DATA).Value = dh
    ws.Cells(r, C_DATA).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(r, C_ANALITO).Value = CStr(p(2))
    ws.Cells(r, C_NIVEL).Value = "Todos"
    ws.Cells(r, C_CORRIDA).Value = CStr(p(6))
    ws.Cells(r, C_CALIB).Value = "Sim"
    ws.Cells(r, C_RESP).Value = mAuditoria.UsuarioSistema()
    ws.Cells(r, C_STATUS).Value = "Calibração registrada pelo ID"
    ws.Cells(r, C_ID).Value = CStr(p(1))
    ws.Cells(r, C_EQUIP).Value = CStr(p(4))
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    quando = Format$(dh, "dd/mm/yyyy hh:mm")
    Auditar "CALIBRACAO_REGISTRADA", "ID " & p(1) & " | " & p(2) & " | " & p(4) & " | RUN " & p(6) & " | " & quando & _
                                     " | lote do resultado " & p(5) & " | aba Registros linha " & r
    RegistrarCalibracaoPorId = "OK|" & r & "|" & p(2) & "|" & p(4) & "|" & quando
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    RegistrarCalibracaoPorId = "ERRO|Não foi possível registrar a calibração (detalhe: " & sE & ")"
End Function

' Chamado pelo Worksheet_Change da aba Registros: um ID digitado (ou colado) na coluna K. ID apagado = nada a fazer
' (a linha fica como esta; para desfazer, o usuario limpa a linha). Recusa = a celula volta vazia, com o motivo.
Public Sub IdDigitado(ByVal Target As Range)
    Dim area As Range, cel As Range, r As String, prot As Boolean, msg As String, ev As Boolean
    On Error GoTo fim
    Set area = Intersect(Target, Target.Worksheet.Range(Target.Worksheet.Cells(R0, C_ID), Target.Worksheet.Cells(R1, C_ID)))
    If area Is Nothing Then Exit Sub
    For Each cel In area.Cells
        If Len(Trim$(CStr(cel.Value))) > 0 Then
            r = RegistrarCalibracaoPorId(CStr(cel.Value), cel.Row)
            If Left$(r, 3) = "OK|" Then
                msg = msg & "Calibração registrada: " & Split(r, "|")(2) & " · " & Split(r, "|")(3) & " · " & _
                      Split(r, "|")(4) & " (ID " & mUI.IdCurto(mIntegracao.NormalizarId(cel.Value)) & ")." & vbCrLf
            Else
                ev = Application.EnableEvents
                Application.EnableEvents = False
                prot = mSeguranca.LiberarEscrita(cel.Worksheet)
                cel.ClearContents
                mSeguranca.RestaurarProtecao cel.Worksheet, prot
                Application.EnableEvents = ev
                If Application.Interactive Then
                    MsgBox Mid$(r, 6), vbExclamation, "Registrar calibração"
                    cel.Select
                End If
                Exit Sub
            End If
        End If
    Next cel
    If Len(msg) > 0 And Application.Interactive Then
        MsgBox msg & vbCrLf & "O ícone de calibração aparece na média do lote, em todos os níveis do gráfico " & _
               "deste analito e equipamento, nas corridas desse dia.", vbInformation, "Registrar calibração"
    End If
fim:
    Application.EnableEvents = True
End Sub

' Linha (R0..R1) que ja tem este ID na coluna K, fora da linha "exceto"; 0 = nenhuma.
Private Function LinhaDoId(ByVal ws As Worksheet, ByVal idN As String, ByVal exceto As Long) As Long
    Dim v As Variant, i As Long
    v = ws.Range(ws.Cells(R0, C_ID), ws.Cells(R1, C_ID)).Value
    For i = 1 To UBound(v, 1)
        If R0 + i - 1 <> exceto Then
            If Not IsError(v(i, 1)) Then
                If Len(Trim$(CStr(v(i, 1)))) > 0 Then
                    If mIntegracao.NormalizarId(v(i, 1)) = idN Then LinhaDoId = R0 + i - 1: Exit Function
                End If
            End If
        End If
    Next i
End Function

' 1a linha sem data, analito, calibracao e ID; 0 = cheia.
Private Function LinhaLivre(ByVal ws As Worksheet) As Long
    Dim v As Variant, i As Long
    v = ws.Range(ws.Cells(R0, C_DATA), ws.Cells(R1, C_EQUIP)).Value
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, 1)) & CStr(v(i, 2)) & CStr(v(i, C_CALIB - 1)) & CStr(v(i, C_ID - 1)))) = 0 Then
            LinhaLivre = R0 + i - 1
            Exit Function
        End If
    Next i
End Function

Private Sub Auditar(ByVal acao As String, ByVal detalhe As String)
    On Error Resume Next
    mAuditoria.RegistrarLog acao, detalhe
End Sub
