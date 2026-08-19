VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo falha
    LockApp
    Exit Sub
falha:
    MsgBox "A pasta nao conseguiu iniciar em modo protegido. Feche o arquivo e acione o responsavel tecnico.", _
           vbCritical, "QC_INI - falha segura"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim eventosAntes As Boolean, recDB As String, recCalc As String
    Dim p As Variant, nErr As Long, sErr As String
    eventosAntes = Application.EnableEvents
    On Error GoTo falha
    Application.EnableEvents = False

    FlushLoteAtual
    mEstatistica.InvalidarCache
    AtualizarCalc
    AtualizarPainelEng
    AtualizarEstatisticaAba
    RegistrarEventosWestgard
    AtualizarBIData

    recDB = ReconciliarBancoBI()
    p = Split(recDB, "|")
    If UBound(p) < 1 Or CLng(Val(CStr(p(1)))) <> 0 Then
        Err.Raise vbObjectError + 570, "QC_INI.BeforeSave", "BI_Data diverge do banco: " & recDB
    End If

    recCalc = ReconciliarComCalc()
    p = Split(recCalc, "|")
    If UBound(p) >= 1 Then
        If CLng(Val(CStr(p(0)))) > 0 And CLng(Val(CStr(p(1)))) <> 0 Then
            Err.Raise vbObjectError + 571, "QC_INI.BeforeSave", "BI_Data diverge do motor: " & recCalc
        End If
    End If

    Application.CalculateFull
    ProtegerEstrutura
    Application.EnableEvents = eventosAntes
    Exit Sub
falha:
    nErr = Err.Number: sErr = Err.Description
    Cancel = True
    Application.EnableEvents = eventosAntes
    MsgBox "O arquivo NAO foi salvo porque a atualizacao controlada falhou." & vbCrLf & vbCrLf & _
           sErr, vbCritical, "QC_INI - salvamento bloqueado"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    SystemLook False
End Sub
