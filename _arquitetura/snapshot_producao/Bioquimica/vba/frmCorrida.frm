VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorrida 
   Caption         =   "Nova Corrida"
   ClientHeight    =   5832
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8385.001
   OleObjectBlob   =   "frmCorrida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCorrida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private analNames() As String
Private analCount As Long

Private Sub UserForm_Initialize()
    Dim c As Collection, i As Long, lote As String
    Set c = ListaAnalitos()
    analCount = c.Count
    If analCount > 0 Then
        ReDim analNames(1 To analCount)
        For i = 1 To analCount
            analNames(i) = c(i)
        Next i
    End If

    Me.txtData.Value = Format(Date, "dd/mm/yyyy")
    Me.cboNivel.Clear
    Me.cboNivel.AddItem "1"
    Me.cboNivel.AddItem "2"

    Me.cboNivel.ListIndex = 0

    Dim lotes As Collection
    Set lotes = ListaLotes()
    Me.cboLote.Clear
    For i = 1 To lotes.Count
        Me.cboLote.AddItem lotes(i)
    Next i
    lote = LoteAtivoCore()
    For i = 0 To Me.cboLote.ListCount - 1
        If Me.cboLote.List(i) = lote Then Me.cboLote.ListIndex = i
    Next i
    If Me.cboLote.ListIndex = -1 And Me.cboLote.ListCount > 0 Then Me.cboLote.ListIndex = 0

    Dim col As Long, rowi As Long, lbl As Object, txt As Object
    Dim baseTop As Long, rowH As Long, col1L As Long, col2L As Long, txtW As Long, lblW As Long, ti As Long
    baseTop = 74: rowH = 20: lblW = 88: txtW = 62: col1L = 12: col2L = 202: ti = 10
    For i = 1 To analCount
        rowi = (i - 1) \ 2
        col = (i - 1) Mod 2
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblV" & i, True)
        lbl.Caption = analNames(i)
        lbl.Top = baseTop + rowi * rowH
        lbl.Left = IIf(col = 0, col1L, col2L)
        lbl.Width = lblW: lbl.Height = 17: lbl.Font.Size = 9
        Set txt = Me.Controls.Add("Forms.TextBox.1", "txtV" & i, True)
        txt.Top = baseTop + rowi * rowH
        txt.Left = IIf(col = 0, col1L + lblW + 2, col2L + lblW + 2)
        txt.Width = txtW: txt.Height = 18: txt.Font.Size = 9
        txt.TabIndex = ti: ti = ti + 1
    Next i

    Dim rows As Long
    rows = (analCount + 1) \ 2
    If rows < 1 Then rows = 1
    Me.Height = baseTop + rows * rowH + 96
    Me.btnSalvar.Top = baseTop + rows * rowH + 14
    Me.btnCancelar.Top = Me.btnSalvar.Top
    Me.lblInfo.Top = Me.btnSalvar.Top - 20
    AtualizarRun
End Sub

Private Sub txtData_Change()
    AtualizarRun
End Sub
Private Sub cboLote_Change()
    AtualizarRun
End Sub

Private Sub AtualizarRun()
    Dim dt As Date, ok As Boolean, lote As String
    dt = ParseData(Me.txtData.Value, ok)
    If Not ok Or Me.cboLote.ListIndex = -1 Then
        Me.lblRun.Caption = ""
        Exit Sub
    End If
    lote = Me.cboLote.Value
    Me.lblRun.Caption = "RUN " & NovoRUN(dt, lote)
End Sub

Private Sub btnSalvar_Click()
    Dim dt As Date, ok As Boolean, lvl As Long, lote As String
    Dim i As Long, n As Long, run As Long, regs() As Variant, res As String, v As Double
    dt = ParseData(Me.txtData.Value, ok)
    If Not ok Then MsgBox "Data inválida.", vbExclamation: Exit Sub
    If Me.cboNivel.ListIndex = -1 Then MsgBox "Selecione o nível.", vbExclamation: Exit Sub
    If Me.cboLote.ListIndex = -1 Then MsgBox "Selecione o lote.", vbExclamation: Exit Sub
    lvl = CLng(Me.cboNivel.Value)
    lote = Me.cboLote.Value
    run = NovoRUN(dt, lote)

    n = 0
    ReDim regs(1 To analCount, 1 To 7)
    For i = 1 To analCount
        Dim s As String
        s = Trim$(CStr(Me.Controls("txtV" & i).Value))
        If s <> "" Then
            v = ParseNum(s, ok)
            If Not ok Then
                MsgBox "Valor inválido em " & analNames(i) & ": """ & s & """", vbExclamation
                Exit Sub
            End If
            n = n + 1
            regs(n, 1) = run
            regs(n, 2) = dt
            regs(n, 3) = lvl
            regs(n, 4) = CodigoLote(lote, lvl)
            regs(n, 5) = analNames(i)
            regs(n, 6) = v
            regs(n, 7) = ST_ATIVO
        End If
    Next i
    If n = 0 Then MsgBox "Nenhum valor informado.", vbExclamation: Exit Sub

    If MsgBox("Gravar " & n & " resultado(s) no RUN " & run & " (Nível " & lvl & ", lote " & lote & ")?", _
              vbQuestion + vbYesNo, "Confirmar") <> vbYes Then Exit Sub

    Dim final_() As Variant, c As Long
    ReDim final_(1 To n, 1 To 7)
    For i = 1 To n
        For c = 1 To 7
            final_(i, c) = regs(i, c)
        Next c
    Next i
    res = UpsertResultados(final_)
    RegistrarLog "LANCAMENTO", "RUN " & run & " N" & lvl
    AtualizarOperacao
    MsgBox "RUN " & run & " gravado." & vbCrLf & _
           "Novos: " & Split(res, "|")(0) & "   ·   Atualizados: " & Split(res, "|")(1), vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


