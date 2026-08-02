Attribute VB_Name = "Planilha2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Dim sel As Range
    Set sel = ThisWorkbook.Names("loteAtivo").RefersToRange
    If sel Is Nothing Then Exit Sub
    If Not Intersect(Target, sel) Is Nothing Then
        Application.EnableEvents = False
        TrocarLote
        Application.EnableEvents = True
    End If
End Sub

