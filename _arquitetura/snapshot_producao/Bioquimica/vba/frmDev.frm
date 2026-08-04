VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDev 
   Caption         =   "Acesso restrito"
   ClientHeight    =   2040
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4785
   OleObjectBlob   =   "frmDev.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

