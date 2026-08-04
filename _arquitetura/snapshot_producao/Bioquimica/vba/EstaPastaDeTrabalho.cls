VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

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

