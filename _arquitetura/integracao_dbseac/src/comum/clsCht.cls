Option Explicit
' ============================================================================
'  clsCht -- eventos de um grafico de Levey-Jennings do Painel (ADR-070)
'
'  mUI.HookCharts liga cada ChartObject do Painel a uma instancia (indice do
'  grafico = nivel). Ao passar o mouse ou clicar num ponto, a caixa "tip" no
'  canto do grafico mostra de QUEM e o ponto -- ID (o numero que se digita na
'  aba Inativar), RUN real, data/hora e valor --, texto montado por
'  mUI.DicaPonto a partir do Eng_Saida (o que o motor publicou). Duplo clique
'  num ponto: mUI.InativarPeloGrafico (pede o motivo e grava a linha na aba
'  Inativar; a atualizacao continua sendo do usuario).
'  Ate o ADR-069 esta classe so existia dentro dos .xlsm e lia a aba Calc:
'  mostrava "RUN <indice do ponto>", data e valor, sem o ID.
' ============================================================================
Public WithEvents c As Chart
Public Nivel As Long
Private mUltimo As String                  ' ultimo elemento sob o mouse: nao reescreve a caixa a cada pixel

Private Sub c_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    Apontar x, y, False
End Sub

Private Sub c_MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    Apontar x, y, True                   ' clique: mostra mesmo que o mouse nao tenha saido do ponto
End Sub

Private Sub c_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
    If ElementID <> xlSeries Or Arg2 < 1 Then Exit Sub
    Cancel = True                          ' nada de painel "Formatar serie"
    mUI.InativarPeloGrafico Nivel, Arg1, Arg2
End Sub

Private Sub Apontar(ByVal x As Long, ByVal y As Long, ByVal sempre As Boolean)
    Dim eid As Long, a1 As Long, a2 As Long, chave As String, msg As String
    On Error Resume Next
    c.GetChartElement x, y, eid, a1, a2
    chave = eid & "|" & a1 & "|" & a2
    If chave = mUltimo And Not sempre Then Exit Sub
    mUltimo = chave
    If eid = xlSeries And a2 >= 1 Then msg = mUI.DicaPonto(Nivel, a1, a2)
    Mostrar msg
End Sub

' Caixa "tip" no canto superior esquerdo do grafico. Mostrar a dica nao "suja" a pasta (nao provoca
' "Deseja salvar?" ao fechar) -- mesma regra da geometria de tela do ADR-061.
Private Sub Mostrar(ByVal msg As String)
    Dim tb As Shape, salvo As Boolean
    On Error Resume Next
    salvo = ThisWorkbook.Saved
    Set tb = Nothing
    Set tb = c.Shapes("tip")
    If tb Is Nothing Then
        If Len(msg) = 0 Then Exit Sub
        Set tb = c.Shapes.AddTextbox(1, 6, 2, 480, 16)
        tb.Name = "tip"
        tb.Line.Visible = msoFalse
        tb.Fill.Visible = msoFalse
        tb.TextFrame2.WordWrap = msoFalse
        tb.TextFrame2.TextRange.Font.Size = 9
        tb.TextFrame2.TextRange.Font.Bold = msoTrue
    End If
    If tb.Width < 470 Then tb.Width = 480  ' a caixa das versoes anteriores tinha 260 pt: o ID nao cabia
    If CStr(tb.TextFrame2.TextRange.Text) <> msg Then tb.TextFrame2.TextRange.Text = msg
    If salvo Then ThisWorkbook.Saved = True
End Sub
