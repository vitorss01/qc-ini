Attribute VB_Name = "mUI"
Option Explicit
' CAMADA: Interface — aparencia de sistema, formularios, eixos e graficos.
Private gHooks As Collection
' ADR-061: graficos do Painel x zoom. Declaracao de modulo SO aqui, antes do 1o procedimento.
Private Const GRAF_VIGIA_SEG As Long = 1          ' 0 = vigia desligado (rollback sem tirar codigo)
' ADR-063: todos os graficos a vista (encaixe do Painel). Medidas em pontos da planilha.
Private Const PAINEL_COL_FIM As String = "U"        ' cabecalho do Painel: A..U nos dois (ADR-054)
Private Const PAINEL_FAIXA_FIM As Long = 38         ' ultima linha da faixa dos graficos (ADR-054)
Private Const GRAF_ALT_MIN As Double = 130          ' altura minima legivel de um grafico
Private Const GRAF_ALT_MAX As Double = 420          ' teto (monitor em pe nao vira grafico de 1 m)
Private Const GRAF_VAO As Double = 8                ' espaco entre um grafico e o seguinte
Private Const PAINEL_FOLGA As Double = 4            ' abaixo da legenda do lote
Private Const PAINEL_ZOOM_MIN As Long = 40
Private Const PAINEL_ZOOM_MAX As Long = 150
Private mVigiaUH As Double                        ' CACHE: ultima UsableHeight vista
Private mCabZoom As Double                        ' CACHE dos cabecalhos (CabecalhosTela): zoom medido
Private mCabLx As Double                          '   largura dos cabecalhos de linha (pt de tela)
Private mCabLy As Double                          '   altura dos cabecalhos de coluna (pt de tela)
#If VBA7 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If
Private mVigiaProx As Date                        ' hora do tique na fila; 0 = nenhum conhecido
Private mVigiaZoom As Double                      ' CACHE (nao estado): ultimo Zoom visto
Private mVigiaUW As Double                        ' CACHE: ultima UsableWidth vista
Private Const GRAF_VIGIA_FOLGA As Long = 1        ' LatestTime = hora + folga: tique vencido expira
Private mVigiaUltimo As Date                      ' hora do ultimo tique agendado (o tique nao zera)
Private mFechando As Boolean                      ' BeforeClose em curso: tique nao reagenda


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

' ADR-054 (regra inalterada): largura = largura VISIVEL da janela, em qualquer zoom.
' UsableWidth vem em pontos de TELA; a largura do objeto e em pontos de PLANILHA.
' ADR-061:
'  - decide ANTES de escrever: sem diferenca nao toca em nada (o vigia chama isto);
'  - NAO mexe em ScreenUpdating: religar a tela forca redesenho SINCRONO e o estado
'    da tela e de quem chamou (ADR-050);
'  - geometria de tela nao suja a pasta (nao provoca "Deseja salvar?");
'  - aba reprotegida pela interface (DrawingObjects=True): ADR-046;
'  - inclui os graficos ocultos pelo foco: voltam com a largura certa.
' Nunca ZOrder/BringToFront: AtualizarEixos e HookCharts casam grafico e nivel pelo INDICE.
Public Sub AjustarGraficos()
    Dim ws As Worksheet, w As Window, co As ChartObject, lx As Double, ly As Double
    Dim larg As Double, z As Double, precisa As Boolean, salvo As Boolean, prot As Boolean
    On Error Resume Next
    If Not GrafNoPainel(w) Then Exit Sub
    If Application.WindowState = xlMinimized Then Exit Sub     ' minimizado nao mede tela
    If w.WindowState = xlMinimized Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Painel")
    z = CDbl(w.Zoom)
    CabecalhosTela w, lx, ly                                    ' ADR-063: a area util os inclui
    If z > 0 Then larg = (w.UsableWidth - lx) * 100# / z
    If larg <= 0 Then larg = w.VisibleRange.Width
    larg = larg - 6
    If larg < 400 Then larg = 400
    If larg > 3000 Then larg = 3000
    For Each co In ws.ChartObjects
        If Abs(co.Width - larg) > 2# Or Abs(co.Left) > 0.5 Then precisa = True
    Next co
    If Not precisa Then Exit Sub                                ' zero escrita, zero redesenho
    salvo = ThisWorkbook.Saved
    If ws.ProtectDrawingObjects Then prot = mSeguranca.LiberarEscrita(ws)
    For Each co In ws.ChartObjects
        If Abs(co.Width - larg) > 2# Or Abs(co.Left) > 0.5 Then
            co.Left = 0
            co.Width = larg
        End If
    Next co
    mSeguranca.RestaurarProtecao ws, prot
    If salvo Then ThisWorkbook.Saved = True
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

' ADR-057: a cadeia completa e o botao ATUALIZAR DADOS (Power Query na ordem,
' sincrono, e so depois motor, Westgard, Estatistica e graficos).
Public Sub AtualizarTudo()
    mIntegracao.AtualizarDados
End Sub

' A janela ATIVA mostra o Painel DESTA pasta? (Bio e Hema podem estar abertas juntas)
Private Function GrafNoPainel(ByRef w As Window) As Boolean
    On Error GoTo nao
    Set w = Application.ActiveWindow
    If w Is Nothing Then Exit Function
    If Not (w.ActiveSheet.Parent Is ThisWorkbook) Then Exit Function
    GrafNoPainel = (w.ActiveSheet.Name = "Painel")
nao:
End Function

' ===================== VIGIA DO ZOOM (ADR-061, etapa A) =====================
' O Excel nao tem evento de zoom, e o zoom real e Ctrl+roda/pinca (faixa e barra de
' status ocultas, SystemLook). Com o Painel DESTA pasta na frente, um tique OnTime por
' segundo le Zoom e UsableWidth (~35 us) e, so se mudaram, chama AjustarGraficos.
'  - so com gente olhando: Interactive (convencao de mLotes.Avisar) e EnableEvents (sem
'    eventos o BeforeClose nao cancela o tique e o Excel REABRE o arquivo -- medido);
'  - so EVENTOS ligam; o tique so se reagenda com o Painel desta pasta na frente;
'  - alvo qualificado pela pasta: Bio e Hema abertas juntas nao se cruzam;
'  - hora sempre em segundo inteiro: perdida a variavel (reset do VBA), o cancelamento
'    a reconstroi varrendo a vizinhanca de Now;
'  - BeforeClose, Workbook_Deactivate, Worksheet_Deactivate e BeforeSave cancelam.
Private Function GrafVigiaAlvo() As String
    GrafVigiaAlvo = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!mUI.GrafVigiaTique"
End Function

Private Function GrafSegundo(ByVal t As Date) As Date
    ' mesma entrada, mesmo Double: cancelar um OnTime exige a hora EXATA
    GrafSegundo = DateSerial(Year(t), Month(t), Day(t)) + TimeSerial(Hour(t), Minute(t), Second(t))
End Function

Private Function GrafVigiaPode() As Boolean
    Dim w As Window
    If GRAF_VIGIA_SEG <= 0 Then Exit Function
    On Error GoTo nao
    If Not Application.Interactive Then Exit Function      ' automacao: nunca
    If Not Application.EnableEvents Then Exit Function     ' sem eventos nao ha quem cancele
    GrafVigiaPode = GrafNoPainel(w)
nao:
End Function

Private Sub GrafVigiaAgendar()
    Dim t As Date
    On Error Resume Next
    t = GrafSegundo(DateAdd("s", GRAF_VIGIA_SEG, Now))
    ' LatestTime: se o Excel estiver ocupado alem da folga, este tique NAO roda (a corrente morre
    ' e o proximo clique/ativacao a religa). E o que permite ao fechamento drenar a fila.
    Application.OnTime EarliestTime:=t, Procedure:=GrafVigiaAlvo(), _
                       LatestTime:=DateAdd("s", GRAF_VIGIA_FOLGA, t), Schedule:=True
    If Err.Number = 0 Then
        mVigiaProx = t
        mVigiaUltimo = t
    Else
        mVigiaProx = 0
    End If
End Sub

' Idempotente: no clique custa uma leitura de variavel.
' BeforeClose em curso (GrafVigiaEncerrar). O "Salvar" da pergunta do Excel ao fechar passa pelo
' BeforeSave DEPOIS do BeforeClose: ele nao pode religar o vigia (tique agendado numa pasta que fecha
' REABRE o arquivo -- ADR-063) nem refazer a vista (ADR-065).
Public Function FechandoPasta() As Boolean
    FechandoPasta = mFechando
End Function

Public Sub GrafVigiaLigar()
    If mVigiaProx <> 0 Then
        If mVigiaProx > Now - TimeSerial(0, 0, 5) Then Exit Sub   ' ha tique vivo na fila
    End If
    If Not GrafVigiaPode() Then Exit Sub
    mFechando = False                     ' fechamento cancelado e o usuario seguiu no Painel
    GrafVigiaDesligar                     ' orfao de reset ou tique perdido: nunca duas correntes
    mVigiaZoom = 0: mVigiaUW = 0: mVigiaUH = 0   ' a primeira leitura so registra
    GrafVigiaAgendar
End Sub

Public Sub GrafVigiaDesligar()
    Dim k As Long, agora As Date, alvo As String
    On Error Resume Next
    alvo = GrafVigiaAlvo()
    If mVigiaProx <> 0 Then Application.OnTime EarliestTime:=mVigiaProx, Procedure:=alvo, Schedule:=False
    agora = Now                           ' variavel perdida: a hora e um segundo inteiro perto de agora
    For k = -3 To GRAF_VIGIA_SEG + 1
        Application.OnTime EarliestTime:=GrafSegundo(DateAdd("s", k, agora)), Procedure:=alvo, Schedule:=False
    Next k
    Err.Clear
    mVigiaProx = 0
End Sub

' Fechamento da pasta (Workbook_BeforeClose). Medido em 03/10/2026: um tique que ja VENCEU
' (hora <= agora) nao se cancela -- Schedule:=False devolve sucesso e o Excel o roda assim que
' fica livre; com a pasta fechada, REABRE o arquivo para isso. Com LatestTime = hora + folga,
' basta manter o Excel ocupado ate a folga do ultimo tique passar: ele expira sem rodar.
' So espera se o vigia agendou algo ha pouco (no maximo folga + 1 s).
' ADR-063 (03/10/2026): quando o fechamento vem de AUTOMACAO (wb.Close), o Excel ignora o
' Application.Wait (volta na hora) e o cancelamento do tique agendado no proprio Close nao tem
' efeito -- o tique vencia e REABRIA o arquivo ~5 s depois (registro instrumentado). Pelo X do
' usuario o Wait funcionava. Laco ocupado vale nos dois caminhos: com o VBA ocupado o Excel nao
' roda OnTime, e o tique passa do LatestTime sem rodar. Gasta no maximo ~3 s, so se houve tique
' agendado ha pouco. mFechando: tique que ainda chegar nao reagenda (GrafVigiaLigar o desfaz, se o
' fechamento for cancelado e o usuario continuar no Painel).
Public Sub GrafVigiaEncerrar()
    Dim limite As Date
    mFechando = True
    GrafVigiaDesligar
    If mVigiaUltimo <> 0 Then
        limite = DateAdd("s", GRAF_VIGIA_FOLGA + 1, mVigiaUltimo)
        If limite > Now And limite <= DateAdd("s", GRAF_VIGIA_SEG + GRAF_VIGIA_FOLGA + 2, Now) Then
            Do While Now < limite              ' sem DoEvents: o Excel fica ocupado de verdade
            Loop
        End If
    End If
    mVigiaUltimo = 0
End Sub

Public Sub GrafVigiaTique()               ' alvo do OnTime
    Dim w As Window, z As Double, uw As Double, uh As Double
    mVigiaProx = 0                        ' este tique saiu da fila
    If mFechando Then Exit Sub            ' pasta fechando: nao reagenda
    If Not GrafVigiaPode() Then Exit Sub  ' NAO reagenda: a corrente morre sozinha
    On Error Resume Next
    If (Not mApp.Ocupado()) And (Application.CutCopyMode = 0) Then
        Set w = Application.ActiveWindow
        z = CDbl(w.Zoom)
        uw = w.UsableWidth                ' em pontos de TELA: nao muda com o zoom
        uh = w.UsableHeight
        If Err.Number = 0 Then
            If mVigiaUW = 0 Or mVigiaUH = 0 Then
                ' 1a leitura desta corrente: so registra. Encaixar aqui desfaria o zoom
                ' que o usuario escolheu (a corrente religa a cada clique)
                mVigiaZoom = z: mVigiaUW = uw: mVigiaUH = uh
                AjustarGraficos
            ElseIf Abs(uw - mVigiaUW) > 0.5 Or Abs(uh - mVigiaUH) > 0.5 Then
                PainelEncaixar            ' ADR-063: a janela mudou de tamanho -> todos a vista
            ElseIf z <> mVigiaZoom Then
                mVigiaZoom = z            ' zoom do usuario: vale; so a largura acompanha
                AjustarGraficos           ' escreve so se passar de 2 pt
            End If
        End If
    End If                                ' ocupado/copiando: tenta no proximo
    Err.Clear
    On Error GoTo 0
    GrafVigiaAgendar
End Sub

Public Function GrafVigiaEstado() As String   ' diagnostico para as provas: "0|..." = parado
    GrafVigiaEstado = CStr(CDbl(mVigiaProx)) & "|" & CStr(mVigiaZoom) & "|" & CStr(mVigiaUW)
End Function

' ===================== TODOS OS GRAFICOS A VISTA (ADR-063) =====================
' Pedido do usuario (03/10/2026): "sao 2 graficos de Levey-Jennings na Bioquimica e 3 na
' Hematologia -- nao pode exibir apenas um na tela; precisa garantir boa visualizacao de todos
' os graficos da aba Painel". Substitui o foco por nivel do ADR-061, que ocultava os outros
' niveis. Medido nesta maquina (1536x864, zoom 100%): o N3 da Hematologia e o N2 da Bioquimica
' ficavam abaixo da borda da janela e o cabecalho A:U passava da direita.
' PainelEncaixar -- ao ENTRAR no Painel, quando a janela muda de tamanho e no botao
' "Ver todos os graficos":
'  1. zoom que cabe o cabecalho inteiro (A:U) e TODOS os graficos com altura minima legivel;
'  2. a faixa dos graficos (linha do 1o grafico ate PAINEL_FAIXA_FIM) acompanha a altura da
'     janela e os graficos a dividem em partes iguais: o mais altos possivel, todos inteiros na
'     tela. A legenda do lote (linha seguinte) e o bloco Sigma descem junto -- nada e coberto;
'  3. rolagem em A1.
' Zoom manual depois disso continua valendo (o vigia so acerta a largura). Nenhum grafico e
' ocultado, nunca. Nao "suja" a pasta (devolve Saved).
Public Function PainelEncaixar() As String
    ' "OK|zoom|altura|graficos" ou o motivo de nao ter encaixado (diagnostico das provas)
    Dim ws As Worksheet, w As Window, co As ChartObject, primeiro As ChartObject
    Dim n As Long, r0 As Long, nLin As Long, tentativa As Long, z As Long
    Dim topo As Double, larg As Double, legenda As Double, uw As Double, uh As Double
    Dim zW As Double, zH As Double, fundo As Double, fundoReal As Double
    Dim hLin As Double, passo As Double, alt As Double, rh As Variant, faixa As Range
    Dim lx As Double, ly As Double, passada As Long
    Dim salvo As Boolean, prot As Boolean, tela As Boolean, encaixou As Boolean
    If Not GrafNoPainel(w) Then PainelEncaixar = "FORA_DO_PAINEL": Exit Function
    If Not Application.Visible Then PainelEncaixar = "OCULTO": Exit Function
    If Application.WindowState = xlMinimized Then PainelEncaixar = "MINIMIZADO": Exit Function
    If w.WindowState = xlMinimized Then PainelEncaixar = "MINIMIZADO": Exit Function
    Set ws = ThisWorkbook.Worksheets("Painel")
    n = ws.ChartObjects.Count
    If n = 0 Then PainelEncaixar = "SEM_GRAFICOS": Exit Function
    For Each co In ws.ChartObjects
        If primeiro Is Nothing Then
            Set primeiro = co
        ElseIf co.Top < primeiro.Top Then
            Set primeiro = co
        End If
    Next co
    topo = primeiro.Top                           ' o topo do 1o grafico nao se mexe
    r0 = primeiro.TopLeftCell.Row
    If r0 > PAINEL_FAIXA_FIM Then PainelEncaixar = "LAYOUT|graficos abaixo da faixa": Exit Function
    larg = ws.Range(PAINEL_COL_FIM & "1").Left + ws.Range(PAINEL_COL_FIM & "1").Width
    legenda = ws.Rows(PAINEL_FAIXA_FIM + 1).RowHeight + PAINEL_FOLGA
    uw = w.UsableWidth                            ' pontos de TELA (iguais em qualquer zoom); ja
    uh = w.UsableHeight                           ' sem as barras de rolagem, COM os cabecalhos
    If uw <= 0 Or uh <= 0 Or larg <= 0 Then PainelEncaixar = "SEM_MEDIDA": Exit Function

    On Error GoTo falhou
    tela = Application.ScreenUpdating
    Application.ScreenUpdating = False
    salvo = ThisWorkbook.Saved
    prot = mSeguranca.LiberarEscritaRapida(ws)
    ' os cabecalhos mudam de tamanho com o zoom: mede no zoom atual, escolhe, mede de novo no
    ' escolhido e refaz a conta ate o zoom parar de mudar (converge em 1 ou 2 passadas)
    CabecalhosTela w, lx, ly
    For passada = 1 To 4
        zW = 100# * (uw - lx) / larg
        zH = 100# * (uh - ly) / (topo + n * (GRAF_ALT_MIN + GRAF_VAO) + legenda)
        If zW < zH Then z = Int(zW) Else z = Int(zH)
        If z < PAINEL_ZOOM_MIN Then z = PAINEL_ZOOM_MIN
        If z > PAINEL_ZOOM_MAX Then z = PAINEL_ZOOM_MAX
        If CLng(w.Zoom) = z Then Exit For
        w.Zoom = z
        CabecalhosTela w, lx, ly
    Next passada
    z = CLng(w.Zoom)
    fundo = (uh - ly) * 100# / z - legenda        ' a faixa acaba aqui: a legenda do lote fica na tela
    If (fundo - topo) / n - GRAF_VAO > GRAF_ALT_MAX Then fundo = topo + n * (GRAF_ALT_MAX + GRAF_VAO)
    If (fundo - topo) / n - GRAF_VAO < GRAF_ALT_MIN Then fundo = topo + n * (GRAF_ALT_MIN + GRAF_VAO)
    Set faixa = ws.Range(ws.Rows(r0), ws.Rows(PAINEL_FAIXA_FIM))
    nLin = PAINEL_FAIXA_FIM - r0 + 1
    For tentativa = 1 To 2                        ' o Excel arredonda a altura da linha ao pixel
        hLin = (fundo - ws.Rows(r0).Top) / nLin
        If hLin > 409 Then hLin = 409
        If hLin < 1 Then hLin = 1
        rh = faixa.RowHeight                      ' Null = linhas com alturas diferentes
        If IsNull(rh) Then
            faixa.RowHeight = hLin
        ElseIf Abs(rh - hLin) > 0.1 Then
            faixa.RowHeight = hLin
        End If
        fundoReal = ws.Rows(PAINEL_FAIXA_FIM).Top + ws.Rows(PAINEL_FAIXA_FIM).Height
        If fundoReal <= fundo + 0.5 Then Exit For
        fundo = fundo - (fundoReal - fundo) - 1   ' arredondou para cima: tira o excesso
    Next tentativa
    passo = (fundoReal - topo) / n
    alt = passo - GRAF_VAO
    For Each co In ws.ChartObjects                ' indice = nivel (AtualizarEixos, HookCharts)
        If Not co.Visible Then co.Visible = True
        If Abs(co.Top - (topo + (co.Index - 1) * passo)) > 0.5 Then co.Top = topo + (co.Index - 1) * passo
        If Abs(co.Height - alt) > 0.5 Then co.Height = alt
    Next co
    w.ScrollRow = 1
    w.ScrollColumn = 1
    mVigiaZoom = z: mVigiaUW = uw: mVigiaUH = uh  ' o vigia nao refaz o que acabou de ser feito
    AjustarGraficos                               ' largura = largura visivel no zoom novo
    PainelEncaixar = "OK|" & z & "|" & Round(alt, 1) & "|" & n
    encaixou = True
sai:
    mSeguranca.RestaurarProtecao ws, prot
    If salvo Then ThisWorkbook.Saved = True
    Application.ScreenUpdating = tela
    If encaixou Then                              ' de novo com a tela ligada: ajuste adiado do Excel
        If w.ScrollRow <> 1 Then w.ScrollRow = 1  ' nao pode esconder a linha 1 (titulo e navegacao)
        If w.ScrollColumn <> 1 Then w.ScrollColumn = 1
    End If
    Exit Function
falhou:
    PainelEncaixar = "ERRO|" & Err.Number & " " & Err.Description
    Resume sai
End Function

' Pixels por ponto da tela (DPI do Windows / 72). O Excel devolve PointsToScreenPixels em pixels
' fisicos; a 125% de escala sao 120 dpi -- 1 pt = 1,667 px.
Private Function PxPorPt() As Double
    #If VBA7 Then
        Dim dc As LongPtr
    #Else
        Dim dc As Long
    #End If
    Dim dpi As Long
    On Error Resume Next
    dc = GetDC(0)
    dpi = GetDeviceCaps(dc, 88)                   ' LOGPIXELSX
    ReleaseDC 0, dc
    If dpi <= 0 Then dpi = 96
    PxPorPt = dpi / 72#
End Function

' Cabecalhos de linha (lx) e de coluna (ly) no zoom ATUAL, em pontos de tela. UsableWidth e
' UsableHeight os INCLUEM (medido em 03/10/2026: so as barras de rolagem ficam de fora) -- sem
' descontar, a borda direita do grafico (legenda) e o rodape do Painel ficavam escondidos.
' Mede ligando/desligando os cabecalhos com a tela congelada; guarda por zoom (no clique, zero troca).
Private Sub CabecalhosTela(ByVal w As Window, ByRef lx As Double, ByRef ly As Double)
    Dim x1 As Long, y1 As Long, x0 As Long, y0 As Long, k As Double, tela As Boolean, salvo As Boolean
    Dim desliguei As Boolean
    lx = 0: ly = 0
    tela = Application.ScreenUpdating
    On Error GoTo sai
    If Not w.DisplayHeadings Then Exit Sub
    If mCabZoom = w.Zoom Then
        lx = mCabLx: ly = mCabLy
        Exit Sub
    End If
    salvo = ThisWorkbook.Saved
    Application.ScreenUpdating = False
    x1 = w.PointsToScreenPixelsX(0): y1 = w.PointsToScreenPixelsY(0)
    desliguei = True
    w.DisplayHeadings = False
    x0 = w.PointsToScreenPixelsX(0): y0 = w.PointsToScreenPixelsY(0)
    w.DisplayHeadings = True
    desliguei = False
    Application.ScreenUpdating = tela
    If salvo Then ThisWorkbook.Saved = True
    k = PxPorPt()
    lx = (x1 - x0) / k: ly = (y1 - y0) / k
    If lx < 0 Or lx > 150 Then lx = 0             ' medida absurda: nao desconta
    If ly < 0 Or ly > 150 Then ly = 0
    mCabZoom = w.Zoom: mCabLx = lx: mCabLy = ly
    Exit Sub
sai:
    lx = 0: ly = 0
    On Error Resume Next
    If desliguei Then w.DisplayHeadings = True       ' nunca deixa o cabecalho do usuario sumido
    Application.ScreenUpdating = tela
End Sub

' Botao "Ver todos os graficos" (btnGrafTodos): volta ao encaixe depois de um zoom manual.
Public Sub PainelVerTodos()
    Dim r As String
    r = PainelEncaixar()
    If Left$(r, 5) = "ERRO|" Then
        MsgBox "Nao consegui ajustar o Painel a janela: " & Mid$(r, 6), vbExclamation, "Painel"
    End If
End Sub

' ADR-062: o spinner ia de 1 a 40 e o cadastro tem menos analitos -- acima do ultimo, o
' Painel caia numa posicao vazia ("0.0" no lugar do nome). O limite acompanha o cadastro.
Public Sub SpinnerLimites()
    Dim ws As Worksheet, n As Long, sp As Object, prot As Boolean, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Painel")
    n = Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets("Analitos").Range("A4:A43"))
    If n < 1 Then n = 1
    Set sp = ws.Spinners("Spinner 1")
    If sp Is Nothing Then Exit Sub
    If sp.Max <> n Then sp.Max = n
    If sp.Min <> 1 Then sp.Min = 1
    If ws.Range("B3").Value > n Or ws.Range("B3").Value < 1 Then
        ev = Application.EnableEvents
        Application.EnableEvents = False
        prot = mSeguranca.LiberarEscritaRapida(ws)
        ws.Range("B3").Value = IIf(ws.Range("B3").Value > n, n, 1)
        mSeguranca.RestaurarProtecao ws, prot
        Application.EnableEvents = ev
    End If
End Sub
