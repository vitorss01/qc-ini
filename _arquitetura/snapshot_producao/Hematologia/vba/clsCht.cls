VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public WithEvents c As Chart
Attribute c.VB_VarHelpID = -1
Public ValCol As Long
Private Sub C_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    Dim eid As Long, a1 As Long, a2 As Long, msg As String
    c.GetChartElement X, Y, eid, a1, a2
    If eid = 3 And a2 >= 1 Then
        Dim r As Long, dt As Variant, vv As Variant
        r = 3 + a2 - 1
        dt = ThisWorkbook.Sheets("Calc").Cells(r, 3).Value
        vv = ThisWorkbook.Sheets("Calc").Cells(r, ValCol).Value
        msg = "RUN " & a2
        If IsDate(dt) Then msg = msg & "  ·  " & Format(dt, "dd/mm/yyyy")
        If IsNumeric(vv) Then msg = msg & "  ·  valor: " & Format(vv, "0.0000")
    Else
        msg = ""
    End If
    Dim tb As Shape
    Set tb = Nothing
    Set tb = c.Shapes("tip")
    If tb Is Nothing Then
        Set tb = c.Shapes.AddTextbox(1, 6, 2, 260, 15)
        tb.Name = "tip"
        tb.Line.Visible = msoFalse
        tb.Fill.Visible = msoFalse
        tb.TextFrame2.TextRange.Font.Size = 9
        tb.TextFrame2.TextRange.Font.Bold = msoTrue
    End If
    tb.TextFrame2.TextRange.Text = msg
End Sub

