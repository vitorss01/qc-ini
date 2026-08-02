Attribute VB_Name = "Planilha7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("B3,G3,G4,M3,M4,N3,N4")) Is Nothing Then
        Application.EnableEvents = False
        AtualizarEixos
        Application.EnableEvents = True
    End If
End Sub
Private Sub Worksheet_Activate()
    AtualizarEixos
    HookCharts
End Sub

