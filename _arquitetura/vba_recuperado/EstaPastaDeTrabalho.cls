Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Workbook_Open()
    LockApp
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    Application.EnableEvents = False
    FlushLoteAtual
    Application.EnableEvents = True
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    SystemLook False
End Sub

