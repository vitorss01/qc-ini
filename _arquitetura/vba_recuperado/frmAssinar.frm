Attribute VB_Name = "frmAssinar"
Attribute VB_Base = "0{37FD3759-16FB-4F37-B58B-81959FE7001E}{D2DD1DD9-9213-4101-93D1-8918A2538D8B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Confirmado As Boolean
Private Sub btnOK_Click()
    Confirmado = True
    Me.Hide
End Sub
Private Sub btnCancel_Click()
    Confirmado = False
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirmado = False
End Sub

