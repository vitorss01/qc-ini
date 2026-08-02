Attribute VB_Name = "Planilha5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim r As Long, nivelAtual As Variant, loteAtual As String
    If Target.Row < 4 Or Target.Column > 7 Then Exit Sub
    r = Target.Row
    Cancel = True
    Application.EnableEvents = False
    If Trim$(CStr(Me.Cells(r, 1).Value)) = "" Then Me.Cells(r, 1).Value = Date
    If Trim$(CStr(Me.Cells(r, 2).Value)) = "" Then Me.Cells(r, 2).Value = LastLevel(r)
    nivelAtual = Me.Cells(r, 2).Value
    If Trim$(CStr(Me.Cells(r, 4).Value)) = "" And IsNumeric(nivelAtual) Then
        On Error Resume Next
        loteAtual = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
        On Error GoTo 0
        If loteAtual <> "" Then
            Me.Cells(r, 4).NumberFormat = "@"
            Me.Cells(r, 4).Value = "QC-" & loteAtual & Format(CLng(nivelAtual), "00")
        End If
    End If
    Me.Cells(r, 5).Select
    Application.EnableEvents = True
End Sub
Private Function LastLevel(ByVal r As Long) As Variant
    Dim i As Long
    For i = r - 1 To 4 Step -1
        If Trim$(CStr(Me.Cells(i, 2).Value)) <> "" Then
            LastLevel = Me.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    LastLevel = 1
End Function

