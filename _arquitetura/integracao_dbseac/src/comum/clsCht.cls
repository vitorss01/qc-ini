Option Explicit
' ============================================================================
'  clsCht -- eventos de um grafico de Levey-Jennings do Painel (ADR-070)
'
'  mUI.HookCharts liga cada ChartObject do Painel a uma instancia (indice do
'  grafico = nivel). Ao passar o mouse ou clicar num ponto, a caixa "tip" no
'  canto do grafico mostra de QUEM e o ponto -- ID (o numero que se digita na
'  aba Inativar), RUN real, data/hora e valor --, texto montado por
'  mUI.DicaPonto a partir do Eng_Saida (o que o motor publicou). Duplo clique
'  ou clique direito num ponto: mUI.InativarPeloGrafico (ADR-081: janela
'  Inativar/Reativar, grava e ja atualiza o grafico; a confirmacao aparece
'  nesta mesma caixa por alguns segundos -- mUI.AvisoGrafico).
'  Ate o ADR-069 esta classe so existia dentro dos .xlsm e lia a aba Calc:
'  mostrava "RUN <indice do ponto>", data e valor, sem o ID.
' ============================================================================
Public WithEvents c As Chart
Public Nivel As Long
Private mUltimo As String                  ' geracao|elemento sob o mouse: nao reescreve a caixa a cada pixel
Private mX As Long, mY As Long             ' ADR-081: onde o botao foi apertado (o BeforeRightClick nao diz)

Private Sub c_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    mX = x: mY = y
    Apontar x, y, False
End Sub

Private Sub c_MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    Apontar x, y, True                   ' clique: mostra mesmo que o mouse nao tenha saido do ponto
End Sub

Private Sub c_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    mX = x: mY = y
End Sub

Private Sub c_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
    On Error Resume Next                   ' evento de grafico: erro nunca abre o depurador
    If ElementID <> xlSeries Or Arg2 < 1 Then Exit Sub
    Cancel = True                          ' nada de painel "Formatar serie"
    mUI.InativarPeloGrafico Nivel, Arg1, Arg2
End Sub

' ADR-081: clique direito num PONTO DE DADOS = a mesma janela do duplo clique (no lugar do menu de contexto).
' Fora de um resultado (eixo, legenda, linha de limite, area vazia) o menu do Excel continua.
Private Sub c_BeforeRightClick(Cancel As Boolean)
    Dim eid As Long, a1 As Long, a2 As Long
    On Error Resume Next
    c.GetChartElement mX, mY, eid, a1, a2
    If eid <> xlSeries Or a2 < 1 Then Exit Sub
    If Left$(mIntegracao.DadosDoPonto(Nivel, a1, a2), 3) <> "OK|" Then Exit Sub
    Cancel = True
    mUI.InativarPeloGrafico Nivel, a1, a2
End Sub

Private Sub Apontar(ByVal x As Long, ByVal y As Long, ByVal sempre As Boolean)
    Dim eid As Long, a1 As Long, a2 As Long, chave As String, msg As String
    On Error Resume Next
    c.GetChartElement x, y, eid, a1, a2
    ' a geracao muda a cada publicacao do motor (mUI.LimparDicas): o mesmo (serie, ponto) depois de trocar
    ' periodo/lote/analito e OUTRA corrida -- a dica e reescrita mesmo com o mouse parado no ponto
    chave = mUI.GeracaoDica() & "|" & eid & "|" & a1 & "|" & a2
    ' ADR-081: o "aguarde" e a confirmacao do inativar/reativar ficam, mesmo com o mouse andando
    If mUI.TipAviso(c, Nivel) Then
        mUltimo = ""                       ' vencido o aviso, o proximo movimento reescreve a dica
        Exit Sub
    End If
    If chave = mUltimo And Not sempre Then Exit Sub
    mUltimo = chave
    If eid = xlSeries And a2 >= 1 Then msg = mUI.DicaPonto(Nivel, a1, a2)
    Mostrar msg
End Sub

' Caixa "tip" no canto superior esquerdo do grafico. ADR-081: quem escreve e o mUI.EscreverTip (o mesmo do
' "aguarde" e da confirmacao do inativar); texto vazio = a dica de uso "Duplo clique num ponto...". Nao suja a pasta.
Private Sub Mostrar(ByVal msg As String)
    On Error Resume Next
    mUI.EscreverTip c, msg
End Sub
