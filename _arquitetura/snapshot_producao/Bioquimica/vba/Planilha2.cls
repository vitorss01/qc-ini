VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

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

