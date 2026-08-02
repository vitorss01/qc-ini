Attribute VB_Name = "Planilha12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

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

