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
        ' ADR-051: visual de aplicativo -- a barra de navegacao de cada tela
        ' substitui as guias (o Inicio tem o botao para mostra-las de volta)
        wnd.DisplayWorkbookTabs = Not onx
    Next wnd
    On Error GoTo 0
End Sub

Public Sub AbrirFormCorrida()
    frmCorrida.Show
End Sub

Public Sub AtualizarEixos()
    Dim ws As Worksheet, i As Long, mn As Variant, mx As Variant, ax As Axis, telaAntes As Boolean
    Set ws = ThisWorkbook.Sheets("Painel")
    telaAntes = Application.ScreenUpdating       ' ADR-050: devolve o estado de quem chamou
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
    ' Religar a tela aqui fazia o Painel piscar no meio de toda operacao maior
    ' (troca de lote, importacao): quem desliga e quem religa (mApp).
    Application.ScreenUpdating = telaAntes
End Sub

' ADR-054: o grafico ocupa a largura VISIVEL da janela, em qualquer zoom.
'
' O usuario aumenta ou diminui o zoom para ver mais coluna; o grafico
' acompanha, sem sobrar faixa branca a direita nem passar da tela. Sem isso a
' largura era fixa: em 70% sobrava tela, em 130% o grafico saia do campo de
' visao e o usuario tinha que rolar para ver o fim da serie.
'
' UsableWidth vem em pontos de TELA (ja com o zoom aplicado); a largura de um
' objeto da planilha e em pontos de PLANILHA. Dai a divisao pelo zoom.
Public Sub AjustarGraficos()
    Dim ws As Worksheet, w As Window, co As ChartObject, larg As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Painel")
    Set w = ActiveWindow
    If w Is Nothing Then Exit Sub
    If w.ActiveSheet Is Nothing Then Exit Sub
    If Not (w.ActiveSheet.Parent Is ThisWorkbook) Then Exit Sub
    If w.ActiveSheet.Name <> "Painel" Then Exit Sub
    If CDbl(w.Zoom) > 0 Then larg = w.UsableWidth * 100# / CDbl(w.Zoom)
    If larg <= 0 Then larg = w.VisibleRange.Width
    larg = larg - 6
    If larg < 400 Then larg = 400
    If larg > 3000 Then larg = 3000
    For Each co In ws.ChartObjects
        If Abs(co.Width - larg) > 2# Or co.Left <> 0 Then
            co.Left = 0
            co.Width = larg
        End If
    Next co
End Sub

' ADR-055: a legenda tinha 14 entradas numa caixa de largura FIXA de 80 pt --
' os rotulos saiam cortados ("Viol...", "Rep..."). Seis delas sao as linhas de
' limite (-3s..+3s), que se leem pela POSICAO no grafico, e "Repeticao"
' aparecia tres vezes (uma por replica). Ficam seis. Reposicionar a legenda
' tira a largura fixa: o Excel a redimensiona para o que os rotulos pedem.
Public Sub AjustarLegendas()
    Dim co As ChartObject, ch As Chart, i As Long, n As Long
    Dim nome As String, vistos As String
    On Error Resume Next
    For Each co In ThisWorkbook.Sheets("Painel").ChartObjects
        Set ch = co.Chart
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionRight
        ch.Legend.Font.Name = "Segoe UI"
        ch.Legend.Font.Size = 9
        n = ch.SeriesCollection.Count
        ' so age na legenda INTEIRA: passar de novo por uma ja limpa apagaria
        ' as entradas que sobraram (o indice e por entrada, nao por serie)
        If ch.Legend.LegendEntries.Count = n And n > 0 Then
            vistos = "|"
            For i = n To 1 Step -1
                nome = ch.SeriesCollection(i).Name
                If InStr(1, "|-3s|-2s|-1s|+1s|+2s|+3s|", "|" & nome & "|") > 0 Then
                    ch.Legend.LegendEntries(i).Delete
                ElseIf InStr(1, vistos, "|" & nome & "|") > 0 Then
                    ch.Legend.LegendEntries(i).Delete     ' nome repetido
                Else
                    vistos = vistos & nome & "|"
                End If
            Next i
        End If
    Next co
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
    AjustarGraficos
    AjustarLegendas
    HookCharts
End Sub

Public Sub AtualizarPainel()
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

