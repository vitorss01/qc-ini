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
'  ADR-085 (29/10/2026): a linha tambem mostra o valor do resultado (J) e pode ser
'  ASSINADA -- duplo clique na coluna I (Assinatura): a mesma assinatura da
'  Liberacao (mSeguranca.AssinarCom: login e senha proprios conferidos, identidade
'  do cadastro, imagem), com o texto da identidade em H terminando em "ID <id>".
'  A assinatura fica presa ao ID: sem ID nao assina; trocar ou apagar o ID de uma
'  linha assinada ANULA a assinatura (Audit_Log ASSINATURA_ANULADA).
'
'  ADR-087 (29/10/2026): a aba deixou de ser 200 linhas fixas trocadas por lote
'  (RegistrosStore). Agora e a Tabela do Excel tblRegistros, com TODOS os lotes
'  (coluna M "Lote" = o lote em uso quando a linha foi registrada), que cresce sozinha:
'  ha sempre UMA linha vazia no fim da tabela (GarantirLinhaNova); preenchida, outra
'  e acrescentada (ListRows.Add -- formato, validacoes e a coluna calculada Nº vem da
'  tabela). A aba e protegida, e tabela protegida nao cresce sozinha: quem acrescenta
'  e o VBA (LiberarEscrita). Filtros (linhas 3:4): Lote, Analito, De/Ate, Equipamento,
'  Calibrado?, Status -- listas dos dados reais (AtualizarListasFiltro), so mudam a
'  VISTA (AutoFilter), nunca os registros; "Limpar filtros" volta a vista completa.
'
'  O GRAFICO nao muda de regra: o icone "Calibracao" continua na MEDIA DO LOTE do
'  nivel (Calc AX1/AZ1/BT1...), nas corridas do mesmo dia e analito com
'  "Foi calibrado? = Sim" -- a mesma regra da marcacao manual. O ADR-084 so
'  acrescenta o equipamento: com L preenchido, o icone sai so no grafico daquele
'  equipamento; linha sem equipamento (as antigas) vale para todos, como antes.
' ============================================================================
Private Const ABA As String = "Registros"
Private Const TB As String = "tblRegistros"   ' ADR-087
Private Const C_LOTE As Long = 13               ' ADR-087: M = lote em uso quando a linha foi registrada
Private Const F_LIN As Long = 4                 ' ADR-087: linha dos campos de filtro (rotulos na 3)
' ADR-087: os filtros -- (coluna do campo na linha F_LIN, coluna da tabela filtrada, rotulo)
Private Const F_LOTE As Long = 2, F_ANALITO As Long = 3, F_DE As Long = 5, F_ATE As Long = 6
Private Const F_EQUIP As Long = 7, F_CALIB As Long = 8, F_STATUS As Long = 10, F_INFO As Long = 11
Private Const C_DATA As Long = 2
Private Const C_ANALITO As Long = 3
Private Const C_NIVEL As Long = 4
Private Const C_CORRIDA As Long = 5
Private Const C_CALIB As Long = 6
Private Const C_RESP As Long = 8
Private Const C_ASSIN As Long = 9
Private Const C_STATUS As Long = 10
Private Const C_ID As Long = 11
Private Const C_EQUIP As Long = 12

' ADR-087: a tabela da aba Registros (Nothing antes do instalar_adr087).
Public Function TabelaRegistros() As ListObject
    On Error Resume Next
    Set TabelaRegistros = ThisWorkbook.Worksheets(ABA).ListObjects(TB)
End Function

' ADR-087: 1a e ultima linha de dados da tabela (a tabela cresce: nada de 4..203 fixo).
Private Function R0() As Long
    R0 = TabelaRegistros().HeaderRowRange.Row + 1
End Function

Private Function R1() As Long
    Dim lo As ListObject
    Set lo = TabelaRegistros()
    R1 = lo.Range.Row + lo.Range.Rows.Count - 1
End Function

' Valores de um intervalo sempre como matriz 2D (uma celula so vira 1x1).
Private Function Matriz(ByVal rg As Range) As Variant
    Dim v As Variant, u(1 To 1, 1 To 1) As Variant
    v = rg.Value
    If IsArray(v) Then
        Matriz = v
    Else
        u(1, 1) = v
        Matriz = u
    End If
End Function

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
    p = Split(d, "|")                        ' OK|ID|analito|nivel|equip|lote|RUN|data|estado|resultado
    Set ws = ThisWorkbook.Worksheets(ABA)
    dup = LinhaDoId(ws, CStr(p(1)), linha)
    If dup > 0 Then
        RegistrarCalibracaoPorId = "ERRO|O ID " & mUI.IdCurto(CStr(p(1))) & " já está registrado como calibração " & _
                                   "(aba Registros, linha " & dup & "). Nada foi gravado."
        Auditar "CALIBRACAO_RECUSADA", "ID " & p(1) & " repetido (ja na linha " & dup & ")"
        Exit Function
    End If
    If linha < R0 Or linha > R1 Then linha = LinhaLivre(ws)      ' ADR-087: sem limite -- a tabela cresce
    If linha = 0 Then
        RegistrarCalibracaoPorId = "ERRO|Não foi possível acrescentar uma linha na tabela de registros. Nada foi gravado."
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
    ' ADR-085: linha ja assinada para este mesmo ID (o ID foi redigitado igual): a identidade da assinatura fica
    If InStr(1, CStr(ws.Cells(r, C_RESP).Value), " " & ChrW$(183) & " ID ", vbBinaryCompare) = 0 Then _
        ws.Cells(r, C_RESP).Value = mAuditoria.UsuarioSistema()
    If UBound(p) >= 9 Then                   ' ADR-085: o valor do resultado do ID, sem redigitar
        ws.Cells(r, C_STATUS).Value = "Calibração registrada pelo ID " & ChrW$(183) & " N" & p(3) & " = " & p(9) & _
                                      " " & ChrW$(183) & " lote " & p(5)
    Else
        ws.Cells(r, C_STATUS).Value = "Calibração registrada pelo ID"
    End If
    ws.Cells(r, C_ID).Value = CStr(p(1))
    ws.Cells(r, C_EQUIP).Value = CStr(p(4))
    If Len(Trim$(CStr(ws.Cells(r, C_LOTE).Value))) = 0 Then ws.Cells(r, C_LOTE).Value = "'" & LoteEmUso()   ' ADR-087
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    GarantirLinhaNova                        ' ADR-087: a proxima linha ja fica pronta
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
    If TabelaRegistros() Is Nothing Then Exit Sub
    Set area = Intersect(Target, Target.Worksheet.Range(Target.Worksheet.Cells(R0, C_ID), Target.Worksheet.Cells(R1, C_ID)))
    If area Is Nothing Then Exit Sub
    For Each cel In area.Cells
        AnularAssinaturaSeOutroId cel.Worksheet, cel.Row, CStr(cel.Value)     ' ADR-085: assinatura presa ao ID
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

' ADR-085: assina a linha da aba Registros com login e senha (nucleo sem dialogo -- o QA chama este). Exige o ID do
' resultado na coluna K e que ele exista nos resultados (a mesma busca exata do inativar). A assinatura e a da
' Liberacao: AssinarCom confere a credencial, grava "Assinado eletronicamente" (ou a rubrica do usuario local) em I e a
' identidade do cadastro + data/hora + "ID <id>" em H, desenha a imagem e registra ASSINATURA_REGISTRO no Audit_Log.
' Devolve "OK|<linha>" ou "ERRO|<mensagem>" (nada gravado no ERRO).
Public Function AssinarRegistroCom(ByVal linha As Long, ByVal login As String, ByVal senha As String) As String
    Dim ws As Worksheet, d As String, p As Variant, res As String, prot As Boolean, ev As Boolean
    On Error GoTo falha
    If linha < R0 Or linha > R1 Then AssinarRegistroCom = "ERRO|Linha fora da tabela de registros.": Exit Function
    Set ws = ThisWorkbook.Worksheets(ABA)
    If Len(Trim$(CStr(ws.Cells(linha, C_ID).Value))) = 0 Then
        AssinarRegistroCom = "ERRO|Informe o ID do resultado na coluna K antes de assinar: a assinatura fica ligada a ele."
        Exit Function
    End If
    d = mIntegracao.DadosDoId(CStr(ws.Cells(linha, C_ID).Value))
    If Left$(d, 3) <> "OK|" Then AssinarRegistroCom = d: Exit Function
    p = Split(d, "|")
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    res = mSeguranca.AssinarCom(ws.Cells(linha, C_ASSIN), login, senha, "REGISTRO", ws.Cells(linha, C_RESP), _
                                "ID " & CStr(p(1)))
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    If res <> "OK" Then AssinarRegistroCom = "ERRO|" & res: Exit Function
    AssinarRegistroCom = "OK|" & linha
    Exit Function
falha:
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    AssinarRegistroCom = "ERRO|Não foi possível assinar (detalhe: " & Err.Description & ")"
End Function

' Duplo clique na coluna I da aba Registros: pede login e senha (a mesma janela da Liberacao) e assina a linha.
Public Sub AssinarRegistro(ByVal cel As Range)
    Dim u As String, s As String, r As String
    On Error Resume Next
    If cel.Row < R0 Or cel.Row > R1 Then Exit Sub
    If Len(Trim$(CStr(cel.Worksheet.Cells(cel.Row, C_ID).Value))) = 0 Then
        MsgBox "Informe o ID do resultado na coluna K antes de assinar: a assinatura fica ligada a ele.", _
               vbExclamation, "Assinatura do registro"
        Exit Sub
    End If
    frmAssinar.lblTipo.Caption = "Assinatura eletrônica - Registro (ID " & _
                                 mUI.IdCurto(mIntegracao.NormalizarId(cel.Worksheet.Cells(cel.Row, C_ID).Value)) & ")"
    frmAssinar.txtLogin.Value = ""
    frmAssinar.txtSenha.Value = ""
    frmAssinar.Show
    If Not frmAssinar.Confirmado Then Unload frmAssinar: Exit Sub
    u = Trim$(frmAssinar.txtLogin.Value): s = frmAssinar.txtSenha.Value
    Unload frmAssinar
    r = AssinarRegistroCom(cel.Row, u, s)
    s = ""
    If Left$(r, 3) <> "OK|" Then MsgBox Mid$(r, 6), vbExclamation, "Assinatura do registro"
End Sub

' ADR-085: a linha estava assinada para um ID e o ID da coluna K mudou (ou foi apagado): a assinatura NAO passa para o
' novo ID -- I e H sao limpos, a imagem sai e o Audit_Log registra ASSINATURA_ANULADA (com o ID antigo e o novo).
Private Sub AnularAssinaturaSeOutroId(ByVal ws As Worksheet, ByVal r As Long, ByVal idNovo As String)
    Dim txt As String, p As Long, idAss As String, idN As String, prot As Boolean, ev As Boolean
    On Error Resume Next
    txt = CStr(ws.Cells(r, C_RESP).Value)
    p = InStrRev(txt, " " & ChrW$(183) & " ID ")
    If p = 0 Then Exit Sub                               ' linha nao assinada pelo ADR-085
    idAss = Trim$(Mid$(txt, p + 6))
    idN = mIntegracao.NormalizarId(idNovo)
    If Len(idN) > 0 And idN = mIntegracao.NormalizarId(idAss) Then Exit Sub
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    ws.Cells(r, C_ASSIN).ClearContents
    ws.Cells(r, C_RESP).ClearContents
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mSeguranca.RedesenharAssinaturas ws
    Auditar "ASSINATURA_ANULADA", "aba Registros linha " & r & " | assinatura era do ID " & idAss & _
                                  " | ID agora: " & IIf(Len(idN) > 0, idN, "(apagado)") & " | " & Left$(txt, 200)
End Sub

' Linha (R0..R1) que ja tem este ID na coluna K, fora da linha "exceto"; 0 = nenhuma.
Private Function LinhaDoId(ByVal ws As Worksheet, ByVal idN As String, ByVal exceto As Long) As Long
    Dim v As Variant, i As Long
    v = Matriz(ws.Range(ws.Cells(R0, C_ID), ws.Cells(R1, C_ID)))
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

' 1a linha sem data, analito, calibracao e ID; nenhuma = a tabela ganha uma linha (ADR-087). 0 = nao conseguiu.
Private Function LinhaLivre(ByVal ws As Worksheet) As Long
    Dim v As Variant, i As Long
    v = Matriz(ws.Range(ws.Cells(R0, C_DATA), ws.Cells(R1, C_EQUIP)))
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, 1)) & CStr(v(i, 2)) & CStr(v(i, C_CALIB - 1)) & CStr(v(i, C_ID - 1)))) = 0 Then
            LinhaLivre = R0 + i - 1
            Exit Function
        End If
    Next i
    LinhaLivre = AcrescentarLinha()
End Function

' ============================================================================
'  ADR-087: tabela que cresce sozinha e filtros da vista
' ============================================================================
' Lote em uso (Configuracao C20) -- o lote da linha nova, como na vista por lote de antes.
Private Function LoteEmUso() As String
    On Error Resume Next
    LoteEmUso = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
End Function

' Uma linha nova no fim da tabela (formato, validacoes e a coluna calculada vem da tabela). Devolve a linha; 0 = falhou.
Public Function AcrescentarLinha() As Long
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, lr As ListRow
    On Error GoTo falha
    Set lo = TabelaRegistros()
    If lo Is Nothing Then Exit Function
    Set ws = lo.Parent
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    Set lr = lo.ListRows.Add
    AcrescentarLinha = lr.Range.Row
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    Exit Function
falha:
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    AcrescentarLinha = 0
End Function

' A ultima linha da tabela esta preenchida? Acrescenta uma vazia -- o usuario sempre tem onde digitar o proximo
' registro (tabela em aba protegida nao cresce sozinha com a digitacao). Devolve True se acrescentou.
Public Function GarantirLinhaNova() As Boolean
    Dim lo As ListObject, v As Variant, j As Long, cheia As Boolean
    On Error GoTo sai
    Set lo = TabelaRegistros()
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then GarantirLinhaNova = (AcrescentarLinha() > 0): Exit Function
    v = Matriz(lo.ListRows(lo.ListRows.Count).Range)
    For j = 2 To UBound(v, 2)                    ' a coluna 1 (Nº) e calculada
        If j <> C_LOTE Then
            If Not IsError(v(1, j)) Then
                If Len(Trim$(CStr(v(1, j)))) > 0 Then cheia = True: Exit For
            End If
        End If
    Next j
    If cheia Then GarantirLinhaNova = (AcrescentarLinha() > 0)
sai:
End Function

' Worksheet_Change da aba Registros: linha da tabela digitada (sem ser pelo ID) ganha o Lote em uso, se nao tiver, e
' a proxima linha fica pronta. Celula de filtro (linha F_LIN) aplica os filtros.
Public Sub RegistrosMudou(ByVal Target As Range)
    Dim lo As ListObject, area As Range, linha As Range, ws As Worksheet, prot As Boolean, ev As Boolean, r As Long
    Dim temDado As Boolean, j As Long
    On Error GoTo fim
    Set lo = TabelaRegistros()
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    If Not Intersect(Target, ws.Rows(F_LIN)) Is Nothing Then
        AplicarFiltrosRegistros
        Exit Sub
    End If
    Set area = Intersect(Target, lo.DataBodyRange)
    If area Is Nothing Then Exit Sub
    For Each linha In area.Rows
        r = linha.Row
        temDado = False
        For j = 2 To C_EQUIP
            If Len(Trim$(CStr(ws.Cells(r, j).Value))) > 0 Then temDado = True: Exit For
        Next j
        If temDado And Len(Trim$(CStr(ws.Cells(r, C_LOTE).Value))) = 0 Then
            ev = Application.EnableEvents
            Application.EnableEvents = False
            prot = mSeguranca.LiberarEscrita(ws)
            ws.Cells(r, C_LOTE).Value = "'" & LoteEmUso()
            mSeguranca.RestaurarProtecao ws, prot
            Application.EnableEvents = ev
        End If
    Next linha
    GarantirLinhaNova
fim:
End Sub

' Os filtros escolhidos na linha F_LIN, aplicados pelo AutoFilter da tabela -- so a VISTA muda; nenhum registro e
' alterado, excluido ou inativado. Campo vazio = sem filtro naquela coluna. Mostra em F_INFO quais filtros estao
' ativos e quantos registros aparecem (ou que nenhum corresponde). Devolve "<visiveis>|<total>".
Public Function AplicarFiltrosRegistros() As String
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, ativos As String, sep As String
    Dim vLote As String, vAn As String, vEq As String, vCal As String, vSt As String, vDe As Variant, vAte As Variant
    Dim nVis As Long, nTot As Long, txt As String
    On Error GoTo falha
    Set lo = TabelaRegistros()
    If lo Is Nothing Then Exit Function
    Set ws = lo.Parent
    sep = "  " & ChrW$(183) & "  "
    vLote = Trim$(CStr(ws.Cells(F_LIN, F_LOTE).Value))
    vAn = Trim$(CStr(ws.Cells(F_LIN, F_ANALITO).Value))
    vDe = ws.Cells(F_LIN, F_DE).Value
    vAte = ws.Cells(F_LIN, F_ATE).Value
    vEq = Trim$(CStr(ws.Cells(F_LIN, F_EQUIP).Value))
    vCal = Trim$(CStr(ws.Cells(F_LIN, F_CALIB).Value))
    vSt = Trim$(CStr(ws.Cells(F_LIN, F_STATUS).Value))
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    If lo.ShowAutoFilter = False Then lo.ShowAutoFilter = True
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    If Len(vLote) > 0 Then FiltrarTexto lo, C_LOTE, vLote: ativos = ativos & sep & "Lote " & vLote
    If Len(vAn) > 0 Then FiltrarTexto lo, C_ANALITO, vAn: ativos = ativos & sep & vAn
    If IsDate(vDe) And IsDate(vAte) Then
        lo.Range.AutoFilter Field:=C_DATA, Criteria1:=">=" & CLng(Int(CDbl(CDate(vDe)))), Operator:=xlAnd, _
                            Criteria2:="<" & CLng(Int(CDbl(CDate(vAte)))) + 1
        ativos = ativos & sep & Format$(vDe, "dd/mm/yyyy") & " a " & Format$(vAte, "dd/mm/yyyy")
    ElseIf IsDate(vDe) Then
        lo.Range.AutoFilter Field:=C_DATA, Criteria1:=">=" & CLng(Int(CDbl(CDate(vDe))))
        ativos = ativos & sep & "desde " & Format$(vDe, "dd/mm/yyyy")
    ElseIf IsDate(vAte) Then
        lo.Range.AutoFilter Field:=C_DATA, Criteria1:="<" & CLng(Int(CDbl(CDate(vAte)))) + 1
        ativos = ativos & sep & "até " & Format$(vAte, "dd/mm/yyyy")
    End If
    If Len(vEq) > 0 Then FiltrarTexto lo, C_EQUIP, vEq: ativos = ativos & sep & vEq
    If Len(vCal) > 0 Then FiltrarTexto lo, C_CALIB, vCal: ativos = ativos & sep & "Calibrado: " & vCal
    If Len(vSt) > 0 Then FiltrarTexto lo, C_STATUS, vSt: ativos = ativos & sep & vSt
    MostrarLinhaDeEntrada lo
    ContarRegistros lo, nVis, nTot
    If Len(ativos) = 0 Then
        txt = "Sem filtro" & sep & nTot & " registro(s)"
    ElseIf nVis = 0 Then
        txt = "Nenhum registro com estes filtros" & ativos & "  " & ChrW$(8212) & "  clique em Limpar filtros"
    Else
        txt = "Filtros ativos:" & Mid$(ativos, Len(sep) + 1) & sep & nVis & " de " & nTot & " registro(s)"
        txt = Replace(txt, "Filtros ativos:", "Filtros ativos: ", 1, 1)
    End If
    ws.Cells(F_LIN, F_INFO).Value = txt
    ws.Cells(F_LIN, F_INFO).Font.Color = IIf(Len(ativos) > 0 And nVis = 0, RGB(192, 0, 0), RGB(31, 78, 121))
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    mSeguranca.RedesenharAssinaturas ws         ' a imagem so existe nas linhas visiveis (a vista mudou)
    AplicarFiltrosRegistros = nVis & "|" & nTot
    Exit Function
falha:
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    AplicarFiltrosRegistros = "ERRO|" & Err.Description
End Function

' Botao "Limpar filtros": esvazia os campos de filtro e mostra a tabela inteira.
Public Sub LimparFiltrosRegistros()
    Dim ws As Worksheet, prot As Boolean, ev As Boolean, c As Variant, r As String
    On Error Resume Next
    If TabelaRegistros() Is Nothing Then Exit Sub
    Set ws = TabelaRegistros().Parent
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    For Each c In Array(F_LOTE, F_ANALITO, F_DE, F_ATE, F_EQUIP, F_CALIB, F_STATUS)
        ws.Cells(F_LIN, c).ClearContents
    Next c
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    r = AplicarFiltrosRegistros()
End Sub

' Filtra a vista pelo lote (troca do lote em uso: a vista acompanha, como a vista por lote de antes).
Public Sub FiltrarRegistrosPorLote(ByVal lote As String)
    Dim ws As Worksheet, prot As Boolean, ev As Boolean, r As String
    On Error Resume Next
    If TabelaRegistros() Is Nothing Then Exit Sub
    Set ws = TabelaRegistros().Parent
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    ws.Cells(F_LIN, F_LOTE).Value = "'" & lote
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    AtualizarListasFiltro
    r = AplicarFiltrosRegistros()
End Sub

' Criterio de texto da coluna: o valor escolhido (a lista vem com Trim) e tambem cada grafia da coluna que so difere
' por espaco sobrando -- senao um registro digitado com "Glicose " nunca apareceria no filtro "Glicose".
Private Sub FiltrarTexto(ByVal lo As ListObject, ByVal campo As Long, ByVal valor As String)
    Dim v As Variant, i As Long, d As Object, s As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d.Add valor, True
    If lo.ListRows.Count > 0 Then
        v = Matriz(lo.ListColumns(campo).DataBodyRange)
        For i = 1 To UBound(v, 1)
            If Not IsError(v(i, 1)) Then
                s = CStr(v(i, 1))
                If s <> Trim$(s) And StrComp(Trim$(s), valor, vbTextCompare) = 0 Then
                    If Not d.Exists(s) Then d.Add s, True
                End If
            End If
        Next i
    End If
    If d.Count = 1 Then
        lo.Range.AutoFilter Field:=campo, Criteria1:=valor
    Else
        lo.Range.AutoFilter Field:=campo, Criteria1:=d.Keys, Operator:=xlFilterValues
    End If
End Sub

' A linha vazia do fim (onde se digita o proximo registro) nunca fica escondida pelo filtro: o filtro so esconde
' registros. Ela tem o Lote vazio e nao casaria com nenhum criterio.
Private Sub MostrarLinhaDeEntrada(ByVal lo As ListObject)
    Dim i As Long, j As Long, v As Variant, vazia As Boolean
    If lo.ListRows.Count = 0 Then Exit Sub
    v = Matriz(lo.DataBodyRange)
    For i = UBound(v, 1) To 1 Step -1
        vazia = True
        For j = 2 To UBound(v, 2)                ' a coluna 1 (Nº) e calculada
            If Not IsError(v(i, j)) Then
                If Len(Trim$(CStr(v(i, j)))) > 0 Then vazia = False: Exit For
            End If
        Next j
        If Not vazia Then Exit For
        If lo.DataBodyRange.Rows(i).EntireRow.Hidden Then lo.DataBodyRange.Rows(i).EntireRow.Hidden = False
    Next i
End Sub

Private Sub ContarRegistros(ByVal lo As ListObject, ByRef nVis As Long, ByRef nTot As Long)
    Dim i As Long, v As Variant, vis As Boolean
    nVis = 0: nTot = 0
    If lo.ListRows.Count = 0 Then Exit Sub
    v = Matriz(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, C_DATA)) & CStr(v(i, C_ANALITO)) & CStr(v(i, C_ID)) & CStr(v(i, 7)))) > 0 Then
            nTot = nTot + 1
            vis = Not lo.DataBodyRange.Rows(i).EntireRow.Hidden
            If vis Then nVis = nVis + 1
        End If
    Next i
End Sub

' As listas dos campos de filtro, tiradas dos DADOS da tabela (e, no Lote, tambem dos lotes cadastrados): nada fixo
' no codigo -- lote novo aparece sozinho. Ficam na area tecnica da aba (colunas AA:AF, ocultas) e as validacoes dos
' campos apontam para elas (nomes lstRegLote, lstRegAnalito, lstRegEquip, lstRegStatus).
Public Sub AtualizarListasFiltro()
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, ev As Boolean, v As Variant, i As Long
    Dim dL As Object, dA As Object, dE As Object, dS As Object, cad As Variant
    On Error GoTo fim
    Set lo = TabelaRegistros()
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    Set dL = CreateObject("Scripting.Dictionary"): Set dA = CreateObject("Scripting.Dictionary")
    Set dE = CreateObject("Scripting.Dictionary"): Set dS = CreateObject("Scripting.Dictionary")
    If lo.ListRows.Count > 0 Then
        v = Matriz(lo.DataBodyRange)
        For i = 1 To UBound(v, 1)
            AnotarTexto dL, v(i, C_LOTE)
            AnotarTexto dA, v(i, C_ANALITO)
            AnotarTexto dE, v(i, C_EQUIP)
            AnotarTexto dS, v(i, C_STATUS)
        Next i
    End If
    cad = Matriz(ThisWorkbook.Names("regLoteCol").RefersToRange)
    For i = 1 To UBound(cad, 1)
        AnotarTexto dL, cad(i, 1)
    Next i
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    EscreverLista ws, 27, dL, "lstRegLote"
    EscreverLista ws, 28, dA, "lstRegAnalito"
    EscreverLista ws, 29, dE, "lstRegEquip"
    EscreverLista ws, 30, dS, "lstRegStatus"
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
fim:
End Sub

Private Sub AnotarTexto(ByVal d As Object, ByVal x As Variant)
    Dim s As String
    If IsError(x) Then Exit Sub
    s = Trim$(CStr(x))
    If Len(s) > 0 And Len(s) <= 80 Then
        If Not d.Exists(s) Then d.Add s, True
    End If
End Sub

' Escreve as chaves em ordem alfabetica na coluna c (a partir da linha 2; a 1 tem o titulo) e aponta o nome para elas.
Private Sub EscreverLista(ByVal ws As Worksheet, ByVal c As Long, ByVal d As Object, ByVal nome As String)
    Dim a() As String, i As Long, j As Long, t As String, n As Long, out() As Variant, k As Variant
    n = d.Count
    ws.Range(ws.Cells(2, c), ws.Cells(ws.Rows.Count, c).End(xlUp).Offset(1, 0)).ClearContents
    If n = 0 Then
        ws.Cells(2, c).Value = ""
        ThisWorkbook.Names.Add Name:=nome, RefersTo:="=" & ws.Cells(2, c).Address(True, True, xlA1, True)
        Exit Sub
    End If
    ReDim a(1 To n)
    i = 0
    For Each k In d.Keys
        i = i + 1: a(i) = CStr(k)
    Next k
    For i = 2 To n                                   ' insercao (listas curtas)
        t = a(i): j = i - 1
        Do While j >= 1
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = "'" & a(i)
    Next i
    ws.Range(ws.Cells(2, c), ws.Cells(1 + n, c)).Value = out
    ThisWorkbook.Names.Add Name:=nome, RefersTo:="=" & ws.Range(ws.Cells(2, c), ws.Cells(1 + n, c)).Address(True, True, xlA1, True)
End Sub

Private Sub Auditar(ByVal acao As String, ByVal detalhe As String)
    On Error Resume Next
    mAuditoria.RegistrarLog acao, detalhe
End Sub
