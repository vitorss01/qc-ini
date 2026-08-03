Attribute VB_Name = "mOperacao"
Option Explicit
' ===== CAMADA DE OPERACAO (Fase 2) =====
' A aba Resultados e apenas VISUALIZACAO: reflete DB_Resultados em 3 blocos por nivel.
' Blocos: N1 = A:G | N2 = K:Q | N3 = U:AA   (colunas I:J, R:T e AB+ ficam livres p/ botoes)
Public Const VIEW_R0 As Long = 4
Public Const VIEW_ROWS As Long = 3000
Public Const VIEW_NLV As Long = 3

Public Function ColunaBloco(ByVal t As Long) As Long
    ColunaBloco = 1 + t * 10          ' t=0 -> A(1), t=1 -> K(11), t=2 -> U(21)
End Function

' Reconstroi a view inteira em LOTE (uma escrita por bloco).
Public Sub AtualizarViewResultados()
    Dim ws As Worksheet, dados As Variant, lote As String
    Dim t As Long, i As Long, n As Long, c0 As Long
    Dim buf() As Variant, cont() As Long
    Set ws = ThisWorkbook.Sheets(VIEW)
    lote = LoteAtivoCore()
    Application.ScreenUpdating = False
    On Error GoTo fim

    For t = 0 To VIEW_NLV - 1
        c0 = ColunaBloco(t)
        ws.Range(ws.Cells(VIEW_R0, c0), ws.Cells(VIEW_R0 + VIEW_ROWS - 1, c0 + 5)).ClearContents
    Next t

    dados = CarregarDB()
    If IsEmpty(dados) Then GoTo fim

    ReDim buf(0 To VIEW_NLV - 1, 1 To VIEW_ROWS, 1 To 6)
    ReDim cont(0 To VIEW_NLV - 1)

    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If Mid$(CStr(dados(i, COL_LOTE)), 4, 6) = lote Then
                t = CLng(dados(i, COL_NIVEL)) - 1
                If t >= 0 And t <= VIEW_NLV - 1 Then
                    If cont(t) < VIEW_ROWS Then
                        cont(t) = cont(t) + 1
                        n = cont(t)
                        buf(t, n, 1) = dados(i, COL_RUN)
                        buf(t, n, 2) = dados(i, COL_DATA)
                        buf(t, n, 3) = dados(i, COL_LOTE)
                        buf(t, n, 4) = dados(i, COL_ANALITO)
                        buf(t, n, 5) = dados(i, COL_RESULT)
                        buf(t, n, 6) = dados(i, COL_STATUS)
                    End If
                End If
            End If
        End If
    Next i

    For t = 0 To VIEW_NLV - 1
        If cont(t) > 0 Then
            Dim outp() As Variant, r As Long, c As Long
            ReDim outp(1 To cont(t), 1 To 6)
            For r = 1 To cont(t)
                For c = 1 To 6
                    outp(r, c) = buf(t, r, c)
                Next c
            Next r
            c0 = ColunaBloco(t)
            ws.Range(ws.Cells(VIEW_R0, c0), ws.Cells(VIEW_R0 + cont(t) - 1, c0 + 5)).Value = outp
        End If
    Next t
fim:
    Application.ScreenUpdating = True
End Sub

' Cadeia unica chamada apos qualquer gravacao/exclusao.
Public Sub AtualizarOperacao()
    Application.ScreenUpdating = False
    AtualizarBanco
    AtualizarViewResultados
    AtualizarEstatistica
    AtualizarPainel
    Application.ScreenUpdating = True
End Sub

Public Sub AbrirFormMassa()
    frmMassa.Show
End Sub

Public Sub AbrirFormExcluir()
    frmExcluir.Show
End Sub

