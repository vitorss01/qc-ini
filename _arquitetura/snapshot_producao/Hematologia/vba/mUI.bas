Attribute VB_Name = "mUI"
Option Explicit
' CAMADA: Interface — aparencia de sistema, formularios, eixos e graficos.
Private gHooks As Collection


Public Sub SystemLook(ByVal onx As Boolean)
    Dim wnd As Window
    On Error Resume Next
    Application.DisplayFormulaBar = Not onx
    Application.DisplayStatusBar = Not onx
    If onx Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    End If
    For Each wnd In ThisWorkbook.Windows
        wnd.DisplayHeadings = Not onx
        wnd.DisplayGridlines = False
    Next wnd
    On Error GoTo 0
End Sub

Public Sub AbrirFormCorrida()
    frmCorrida.Show
End Sub

Public Sub AtualizarEixos()
    Dim ws As Worksheet, i As Long, mn As Variant, mx As Variant, ax As Axis
    Set ws = ThisWorkbook.Sheets("Painel")
    Application.ScreenUpdating = False
    For i = 1 To ws.ChartObjects.Count
        On Error Resume Next
        mn = ThisWorkbook.Names("axmin" & i).RefersToRange.Value
        mx = ThisWorkbook.Names("axmax" & i).RefersToRange.Value
        Set ax = ws.ChartObjects(i).Chart.Axes(xlValue)
        If IsNumeric(mn) And IsNumeric(mx) Then
            If mx > mn Then
                ax.MinimumScale = mn
                ax.MaximumScale = mx
            End If
        Else
            ax.MinimumScaleIsAuto = True
            ax.MaximumScaleIsAuto = True
        End If
        ' eixo X (RUN) — remove o espaco em branco a esquerda
        Dim xn As Variant, xx As Variant, axx As Axis
        xn = ThisWorkbook.Names("xrunmin").RefersToRange.Value
        xx = ThisWorkbook.Names("xrunmax").RefersToRange.Value
        Set axx = ws.ChartObjects(i).Chart.Axes(xlCategory)
        If IsNumeric(xn) And IsNumeric(xx) Then
            If xx > xn Then
                axx.MinimumScale = xn
                axx.MaximumScale = xx
            End If
        Else
            axx.MinimumScaleIsAuto = True
            axx.MaximumScaleIsAuto = True
        End If
        On Error GoTo 0
    Next i
    Application.ScreenUpdating = True
End Sub

Public Sub HookCharts()
    On Error Resume Next
    Dim co As ChartObject, h As clsCht
    Set gHooks = New Collection
    For Each co In ThisWorkbook.Sheets("Painel").ChartObjects
        Set h = New clsCht
        Set h.c = co.Chart
        h.ValCol = 6 + (co.Index - 1) * 22
        gHooks.Add h
    Next co
End Sub


' ===================== ORQUESTRACAO (Etapa 7) =====================
' Rotinas de responsabilidade unica; a cadeia completa e AtualizarTudo.
Public Sub AtualizarResultados()
    Application.Calculate
End Sub

Public Sub AtualizarGraficos()
    AtualizarEixos
    HookCharts
End Sub

Public Sub AtualizarPainel()
    AtualizarPainelEng
    On Error Resume Next
    ThisWorkbook.Sheets("Painel").Calculate
    AtualizarGraficos
End Sub

Public Sub AtualizarTudo()
    Application.ScreenUpdating = False
    AtualizarBanco
    AtualizarResultados
    AtualizarEstatistica
    AtualizarPainel
    Application.ScreenUpdating = True
End Sub

