VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Row < 4 Then Exit Sub
    If Target.Column <> 3 And Target.Column <> 5 Then Exit Sub
    If Trim$(CStr(Me.Cells(Target.Row, 2).Value)) = "" Then Exit Sub
    Cancel = True
    If Target.Column = 3 Then
        Assinar Target, "INICIADO"
    Else
        Assinar Target, "FINALIZADO"
    End If
End Sub

