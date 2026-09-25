Attribute VB_Name = "Planilha3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

' ADR-049: Media/DP digitados aqui pertencem ao lote em tela (loteParam) e sao
' gravados NA HORA na LotesStore, com rastro no Audit_Log.
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("E4:J43")) Is Nothing Then Exit Sub
    mLotes.ParametroEditado Target
End Sub
