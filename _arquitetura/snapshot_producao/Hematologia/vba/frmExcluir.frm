VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExcluir 
   Caption         =   "Excluir resultados (exclusão lógica)"
   ClientHeight    =   7440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6990
   OleObjectBlob   =   "frmExcluir.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExcluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cboNivel.Clear
    Me.cboNivel.AddItem "1"
    Me.cboNivel.AddItem "2"
    Me.cboNivel.AddItem "3"

    Me.cboNivel.ListIndex = 0
    CarregarRuns
    CarregarAnalitos
End Sub

Private Sub cboNivel_Change()
    CarregarAnalitos
End Sub

Private Sub CarregarAnalitos()
    Dim c As Collection, i As Long
    Me.lstAnalitos.Clear
    Set c = ListaAnalitos()
    For i = 1 To c.Count
        Me.lstAnalitos.AddItem c(i)
    Next i
    Me.lblInfo.Caption = c.Count & " analito(s) do nível selecionado."
End Sub

Private Sub CarregarRuns()
    Dim runs As Collection, i As Long
    Me.cboRun.Clear
    Set runs = RunsDoLote(LoteAtivoCore())
    For i = 1 To runs.Count
        Me.cboRun.AddItem CStr(runs(i))
    Next i
    If Me.cboRun.ListCount > 0 Then Me.cboRun.ListIndex = Me.cboRun.ListCount - 1
End Sub

Private Sub btnTodos_Click()
    Dim i As Long
    For i = 0 To Me.lstAnalitos.ListCount - 1
        Me.lstAnalitos.Selected(i) = True
    Next i
End Sub

Private Sub btnLimpar_Click()
    Dim i As Long
    For i = 0 To Me.lstAnalitos.ListCount - 1
        Me.lstAnalitos.Selected(i) = False
    Next i
End Sub

Private Sub btnConfirmar_Click()
    Dim i As Long, alvoS As Object, n As Long, lvl As Long, run As Long
    If Me.cboRun.ListIndex = -1 Then MsgBox "Selecione a corrida (RUN).", vbExclamation: Exit Sub
    If Me.cboNivel.ListIndex = -1 Then MsgBox "Selecione o nível.", vbExclamation: Exit Sub
    Set alvoS = CreateObject("Scripting.Dictionary")
    For i = 0 To Me.lstAnalitos.ListCount - 1
        If Me.lstAnalitos.Selected(i) Then alvoS(UCase$(Trim$(Me.lstAnalitos.List(i)))) = 1
    Next i
    If alvoS.Count = 0 Then MsgBox "Marque pelo menos um analito.", vbExclamation: Exit Sub
    lvl = CLng(Me.cboNivel.Value)
    run = CLng(Me.cboRun.Value)

    If MsgBox("Excluir logicamente " & alvoS.Count & " analito(s) do RUN " & run & _
              " (Nível " & lvl & ")?" & vbCrLf & vbCrLf & _
              "Os registros NÃO são apagados — apenas marcados como Excluído.", _
              vbExclamation + vbYesNo, "Confirmar exclusão") <> vbYes Then Exit Sub

    n = ExcluirLogico(run, lvl, alvoS)
    RegistrarLog "EXCLUSAO", "RUN " & run & " N" & lvl & " (" & n & ")"
    AtualizarOperacao
    MsgBox n & " registro(s) marcado(s) como Excluído.", vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


