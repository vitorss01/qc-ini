Attribute VB_Name = "frmCorrida"
Attribute VB_Base = "0{A4535387-D350-4AC6-B95D-2C3B8D6062A9}{A1AE3AAD-ADFF-4E1D-B656-9F912DC06D75}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
Private analNames() As String
Private analCount As Long

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, i As Long, n As Long, nm As String
    Dim tmp(1 To 40) As String
    Set ws = ThisWorkbook.Sheets("Analitos")
    n = 0
    For i = 1 To 40
        nm = Trim$(CStr(ws.Cells(3 + i, 1).Value))
        If nm <> "" Then
            n = n + 1
            tmp(n) = nm
        End If
    Next i
    analCount = n
    ReDim analNames(1 To n)
    For i = 1 To n
        analNames(i) = tmp(i)
    Next i

    Me.txtData.Value = Format(Date, "dd/mm/yyyy")
    Me.cboNivel.Clear
    Me.cboNivel.AddItem "1"
    Me.cboNivel.AddItem "2"
    Me.cboNivel.AddItem "3"

    Me.cboNivel.ListIndex = 0

    Dim col As Long, rowi As Long
    Dim lbl As Object, txt As Object
    Dim baseTop As Long, rowH As Long, col1L As Long, col2L As Long, txtW As Long, lblW As Long, ti As Long
    baseTop = 64: rowH = 20: lblW = 88: txtW = 62: col1L = 12: col2L = 202: ti = 10
    For i = 1 To n
        rowi = (i - 1) \ 2
        col = (i - 1) Mod 2
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblV" & i, True)
        lbl.Caption = analNames(i)
        lbl.Top = baseTop + rowi * rowH
        lbl.Left = IIf(col = 0, col1L, col2L)
        lbl.Width = lblW
        lbl.Height = 17
        lbl.Font.Size = 9
        Set txt = Me.Controls.Add("Forms.TextBox.1", "txtV" & i, True)
        txt.Top = baseTop + rowi * rowH
        txt.Left = IIf(col = 0, col1L + lblW + 2, col2L + lblW + 2)
        txt.Width = txtW
        txt.Height = 18
        txt.Font.Size = 9
        txt.TabIndex = ti
        ti = ti + 1
    Next i

    Dim rows As Long, formH As Long
    rows = (n + 1) \ 2
    If rows < 1 Then rows = 1
    formH = baseTop + rows * rowH + 130
    Me.Height = formH
    Me.btnCarregar.Top = baseTop + rows * rowH + 8
    Me.lblInfoCarregar.Top = Me.btnCarregar.Top + 34
    Me.btnSalvar.Top = Me.lblInfoCarregar.Top + 38
    Me.btnCancelar.Top = Me.btnSalvar.Top

    AtualizarSeq
End Sub

Private Sub txtData_Change()
    AtualizarSeq
End Sub

' Seq = numero da corrida. E por DATA (nao por nivel): se a data ja tiver alguma
' linha (de QUALQUER analito/nivel), reaproveita o mesmo Seq; senao, cria um novo.
' (Antes calculava por nivel isoladamente, o que podia gerar Seq colidindo entre
' corridas de datas diferentes quando os niveis eram lancados em sessoes separadas.)
Private Function SeqParaData(ByVal dt As Date) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long, mx As Long
    Set ws = ThisWorkbook.Sheets("Resultados")
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    mx = 0
    For i = 4 To lastRow
        If IsDate(ws.Cells(i, 1).Value) Then
            If CDate(ws.Cells(i, 1).Value) = dt Then
                SeqParaData = ws.Cells(i, 3).Value
                Exit Function
            End If
            If ws.Cells(i, 3).Value > mx Then mx = ws.Cells(i, 3).Value
        End If
    Next i
    SeqParaData = mx + 1
End Function

Private Sub AtualizarSeq()
    Dim dt As Date
    On Error GoTo semData
    dt = CDate(Me.txtData.Value)
    Me.lblSeqInfo.Caption = "Corrida n. " & SeqParaData(dt)
    Exit Sub
semData:
    Me.lblSeqInfo.Caption = ""
End Sub

Private Sub btnCarregar_Click()
    Dim ws As Worksheet, r As Long, lastRow As Long, dt As Date, i As Long, v As Variant
    Dim achou As Boolean, n As Long, lvl As Long
    On Error Resume Next
    dt = CDate(Me.txtData.Value)
    If Err.Number <> 0 Then
        MsgBox "Informe uma data valida antes de carregar.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    If Me.cboNivel.ListIndex = -1 Then
        MsgBox "Selecione o nivel antes de carregar.", vbExclamation
        Exit Sub
    End If
    lvl = CLng(Me.cboNivel.Value)
    Set ws = ThisWorkbook.Sheets("Resultados")
    lastRow = ws.Cells(ws.rows.Count, 11).End(xlUp).Row
    achou = False
    For r = 4 To lastRow
        If IsDate(ws.Cells(r, 11).Value) Then
            If CDate(ws.Cells(r, 11).Value) = dt And ws.Cells(r, 12).Value = lvl Then
                achou = True
                Exit For
            End If
        End If
    Next r
    If Not achou Then
        MsgBox "Nao ha dados interfaceados para " & Format(dt, "dd/mm/yyyy") & " - Nivel " & lvl & ".", vbInformation
        Exit Sub
    End If
    n = 0
    For i = 1 To analCount
        v = ws.Cells(r, 12 + i).Value
        If Trim$(CStr(v)) <> "" And IsNumeric(v) Then
            Me.Controls("txtV" & i).Value = v
            n = n + 1
        End If
    Next i
    MsgBox n & " valor(es) carregado(s) (interfaceamento de " & Format(dt, "dd/mm/yyyy") & ", Nivel " & lvl & ").", vbInformation
End Sub

Private Sub btnSalvar_Click()
    Dim ws As Worksheet, i As Long, r As Long, lvl As Long, sq As Long, dt As Date
    Dim n As Long, v As Variant, lastRowA As Long
    On Error Resume Next
    dt = CDate(Me.txtData.Value)
    If Err.Number <> 0 Then
        MsgBox "Data invalida.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    If Me.cboNivel.ListIndex = -1 Then
        MsgBox "Selecione o nivel.", vbExclamation
        Exit Sub
    End If
    lvl = CLng(Me.cboNivel.Value)
    Set ws = ThisWorkbook.Sheets("Resultados")
    sq = SeqParaData(dt)
    Dim loteAtual As String
    loteAtual = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
    If loteAtual = "" Then
        MsgBox "Nenhum lote em uso selecionado (aba Configuracao).", vbExclamation
        Exit Sub
    End If
    lastRowA = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    n = 0
    r = lastRowA + 1
    For i = 1 To analCount
        v = Me.Controls("txtV" & i).Value
        If Trim$(CStr(v)) <> "" Then
            If Not IsNumeric(v) Then
                MsgBox "Valor invalido em " & analNames(i) & ".", vbExclamation
                Exit Sub
            End If
            ws.Cells(r, 1).Value = dt
            ws.Cells(r, 2).Value = lvl
            ws.Cells(r, 3).Value = sq
            ws.Cells(r, 4).NumberFormat = "@"
            ws.Cells(r, 4).Value = "QC-" & loteAtual & Format(lvl, "00")   ' Lote (código completo)
            ws.Cells(r, 5).Value = analNames(i)
            ws.Cells(r, 6).Value = CDbl(v)
            r = r + 1
            n = n + 1
        End If
    Next i
    If n = 0 Then
        MsgBox "Nenhum valor informado.", vbExclamation
        Exit Sub
    End If
    Application.CalculateFull
    MsgBox n & " resultado(s) salvo(s) na corrida " & sq & " (Nivel " & lvl & ").", vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

