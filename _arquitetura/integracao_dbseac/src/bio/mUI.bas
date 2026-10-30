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
Private mCabRuim As Boolean                       ' ultima medida invalida: o vigia mede de novo no proximo tique
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
Private mGeracaoDica As Long                      ' ADR-070 (revisao): muda a cada publicacao do motor (LimparDicas)
Private mFechando As Boolean                      ' BeforeClose em curso: tique nao reagenda
' ADR-081: janela Inativar/Reativar do grafico. Resposta do QA no lugar da janela (DefinirRespostaTeste; vale uma
' vez); o aviso da caixa "tip" (progresso enquanto atualiza, confirmacao por AVISO_SEG segundos -- o mouse continua
' sobre o grafico depois da janela); e o bloqueio de cliques que ficaram na fila (BLOQUEIO_SEG depois de terminar).
Private mRespostaTeste As String
Private mAvisoTxt As String
Private mAvisoNivel As Long
Private mAvisoAte As Date
Private mAvisoEstilo As Long
Private mUltimoProgresso As String
Private mOcupadoUI As Boolean
Private mFimT As Double                           ' Timer do fim da ultima inativacao/reativacao (0 = nenhuma)
Private Const AVISO_SEG As Long = 10
Private Const BLOQUEIO_SEG As Double = 2
Private Const TIP_DICA As Long = 0                ' dica do ponto: 9 pt negrito
Private Const TIP_USO As Long = 1                 ' dica de uso (caixa sem ponto): 9 pt cinza
Private Const TIP_PROGRESSO As Long = 2           ' "aguarde": 10 pt negrito azul
Private Const TIP_OK As Long = 3                  ' confirmacao: 10 pt negrito verde
Private Const DICA_USO As String = "Duplo clique num ponto: não conforme, inativar ou comentário técnico"
Private Const TXT_FALHA As String = "Nada foi alterado. Tente de novo ou clique em ATUALIZAR DADOS."


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
        ' ADR-078: zera a cada grafico -- nome ausente nao herda o eixo do grafico anterior
        mn = Empty: mx = Empty: Set ax = Nothing
        mn = ThisWorkbook.Names("axmin" & i).RefersToRange.Value
        mx = ThisWorkbook.Names("axmax" & i).RefersToRange.Value
        Set ax = ws.ChartObjects(i).Chart.Axes(xlValue)
        If Not ax Is Nothing Then Escala ax, mn, mx
        ' eixo X (corrida do lote) — remove o espaco em branco a esquerda. ADR-078: sem corrida, moldura fixa 0..21 (Calc)
        Dim xn As Variant, xx As Variant, axx As Axis
        xn = Empty: xx = Empty: Set axx = Nothing
        ' ADR-080: cada nivel tem a sua sequencia de corridas do lote (xrunmin1, xrunmax1, ...); sem eles, o eixo comum
        xn = ThisWorkbook.Names("xrunmin" & i).RefersToRange.Value
        xx = ThisWorkbook.Names("xrunmax" & i).RefersToRange.Value
        If Not (EhNumero(xn) And EhNumero(xx)) Then
            xn = ThisWorkbook.Names("xrunmin").RefersToRange.Value
            xx = ThisWorkbook.Names("xrunmax").RefersToRange.Value
        End If
        Set axx = ws.ChartObjects(i).Chart.Axes(xlCategory)
        If Not axx Is Nothing Then
            Escala axx, xn, xx
            If EhNumero(xn) And EhNumero(xx) Then
                If CDbl(xx) - CDbl(xn) <= 12 Then axx.MajorUnit = 1 Else axx.MajorUnitIsAuto = True
            End If
        End If
        On Error GoTo 0
    Next i
    ' Religar a tela aqui fazia o Painel piscar no meio de toda operacao maior
    ' (troca de lote, importacao): quem desliga e quem religa (mApp).
    Application.ScreenUpdating = telaAntes
End Sub

' ADR-078: grava minimo e maximo na ordem que o Excel aceita. Escrever um minimo acima do maximo ATUAL falhava em
' silencio (On Error Resume Next) e o eixo ficava com o maximo novo e o minimo velho -- grafico achatado.
' Sem os dois numeros (ou com mx <= mn), eixo automatico.
Private Sub Escala(ByVal ax As Axis, ByVal mn As Variant, ByVal mx As Variant)
    On Error Resume Next
    If EhNumero(mn) And EhNumero(mx) Then
        If CDbl(mx) > CDbl(mn) Then
            If CDbl(mn) >= ax.MaximumScale Then
                ax.MaximumScale = CDbl(mx)
                ax.MinimumScale = CDbl(mn)
            Else
                ax.MinimumScale = CDbl(mn)
                ax.MaximumScale = CDbl(mx)
            End If
            Exit Sub
        End If
    End If
    ax.MinimumScaleIsAuto = True
    ax.MaximumScaleIsAuto = True
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

' Liga cada grafico do Painel a uma clsCht (dica do ponto e duplo clique). Indice do grafico = nivel.
Public Sub HookCharts()
    On Error Resume Next
    Dim co As ChartObject, h As clsCht
    Set gHooks = New Collection
    For Each co In ThisWorkbook.Sheets("Painel").ChartObjects
        Set h = New clsCht
        Set h.c = co.Chart
        h.Nivel = co.Index                        ' ADR-070: a dica le Eng_Saida pelo nivel (antes: coluna do Calc)
        gHooks.Add h
    Next co
End Sub

' ADR-085: o duplo clique so e interceptado (nada de "Formatar" do Excel) com cada grafico ligado a uma clsCht. Perdida
' a colecao (reset do VBA, erro nao tratado em outra rotina), religa: chamada a cada clique no Painel (GrafVigiaLigar)
' e a cada tique do vigia (1 s, so com o Painel na frente). Custa uma contagem.
Public Sub GarantirHooks()
    On Error Resume Next
    If gHooks Is Nothing Then
        HookCharts
    ElseIf gHooks.Count <> ThisWorkbook.Sheets("Painel").ChartObjects.Count Then
        HookCharts
    End If
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
    GarantirHooks                         ' ADR-085: duplo clique sempre interceptado
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
    mApp.GarantirEventos                  ' ADR-078: ANTES do GrafVigiaPode (que exige eventos ligados)
    If Not GrafVigiaPode() Then Exit Sub  ' NAO reagenda: a corrente morre sozinha
    GarantirHooks                         ' ADR-085: religa os graficos se o VBA perdeu a colecao
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
            If mCabRuim Then              ' cabecalho medido no meio do redesenho: refaz no proximo tique
                If mVigiaZoom = z Then mVigiaZoom = 0
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
    Dim lx As Double, ly As Double, passada As Long, ajuste As Long
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
    ' o Excel desenha cada coluna em pixels inteiros: em zoom baixo a soma arredondada passa da conta
    ' e a coluna U fica cortada (tela pequena do laboratorio, 05/10/2026). Confere na tela: a coluna
    ' seguinte a U tem de aparecer, nem que em parte -- senao desce o zoom de 1 em 1
    w.ScrollColumn = 1
    Do While z > PAINEL_ZOOM_MIN
        If w.VisibleRange.Column + w.VisibleRange.Columns.Count - 1 > ws.Range(PAINEL_COL_FIM & "1").Column Then Exit Do
        z = z - 1
        w.Zoom = z
    Loop
    If CLng(w.Zoom) <> z Then w.Zoom = z
    CabecalhosTela w, lx, ly
    fundo = (uh - ly) * 100# / z - legenda        ' a faixa acaba aqui: a legenda do lote fica na tela
    If (fundo - topo) / n - GRAF_VAO > GRAF_ALT_MAX Then fundo = topo + n * (GRAF_ALT_MAX + GRAF_VAO)
    If (fundo - topo) / n - GRAF_VAO < GRAF_ALT_MIN Then fundo = topo + n * (GRAF_ALT_MIN + GRAF_VAO)
    Set faixa = ws.Range(ws.Rows(r0), ws.Rows(PAINEL_FAIXA_FIM))
    nLin = PAINEL_FAIXA_FIM - r0 + 1
    w.ScrollRow = 1
    For ajuste = 1 To 4                           ' confere na tela: a linha depois da legenda aparece
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
        If w.VisibleRange.Row + w.VisibleRange.Rows.Count - 1 > PAINEL_FAIXA_FIM + 1 Then Exit For
        If (fundo - topo) / n - GRAF_VAO - 4# < GRAF_ALT_MIN Then Exit For
        fundo = fundo - 4#                        ' zoom baixo: as linhas arredondadas somam mais que a conta
    Next ajuste
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
    ' cabecalho LIGADO tem largura > 0. Zero logo depois de trocar o zoom = o Excel ainda nao
    ' redesenhou (medido no laboratorio, 05/10/2026: guardado no cache, os graficos ficavam
    ' ~15 pt largos demais naquele zoom ate fechar a pasta). Nao guarda; o vigia mede de novo
    If lx <= 0 Or lx > 150 Or ly <= 0 Or ly > 150 Then
        lx = 0: ly = 0
        mCabRuim = True
        mCabZoom = 0
        Exit Sub
    End If
    mCabRuim = False
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
    n = Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets(mLotes.AbaEspec()).Range("A4:A43"))
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

' ===================== DICA DO LEVEY-JENNINGS (ADR-070) =====================
' O usuario passa o mouse (ou clica) num ponto do grafico do Painel e ve DE QUEM e o ponto: o ID do
' resultado -- o numero que se digita na aba Inativar --, o RUN real, a data/hora e o valor. Antes a caixa
' lia a aba Calc e mostrava "RUN <indice do ponto>", sem ID.
'   nivel = indice do grafico (1 = N1); serie = indice da serie no grafico; ponto = indice do ponto, que e o
'   slot da corrida (Calc e Eng_Saida: linha 2 + ponto). Funcao publica e sem efeito colateral: o QA a chama
'   por automacao (testes/qa_inativar.py) com o mesmo (nivel, serie, ponto) que o Excel passa no evento.
' Series do grafico: "Resultado"/"OK"/"Violacao" = o ponto do nivel; "Nao conforme (X)" (tres, uma por X da
' corrida) = os X vermelhos; linhas de limite e "Calibracao" = so a corrida. Devolve "" quando nao ha ponto.
Public Function DicaPonto(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim ch As Chart, eng As Worksheet, nm As String, r As Long, k As Long, slot As Long, col As Long
    Dim v As Variant, idv As Variant, dh As Variant, runv As Variant
    On Error GoTo sai
    If nivel < 1 Or nivel > mEstatistica.NLV Or ponto < 1 Or ponto > 182 Then Exit Function
    Set ch = ThisWorkbook.Worksheets("Painel").ChartObjects(nivel).Chart
    If serie < 1 Or serie > ch.SeriesCollection.Count Then Exit Function
    nm = CStr(ch.SeriesCollection(serie).Name)
    ' Linhas de limite (ADR-086): series de 182 pontos = as corridas 1..180 (degrau do modo TODOS, cada uma com o alvo do
    ' lote do proprio ponto) + as 2 bordas do eixo do ADR-078 (181 e 182, o alvo do lote do Painel no modo normal)
    Select Case nm
        Case "-3s", "-2s", "-1s", "Média", "+1s", "+2s", "+3s"
            v = ch.SeriesCollection(serie).Values
            If IsArray(v) Then
                If ponto <= UBound(v) - LBound(v) + 1 Then
                    If EhNumero(v(LBound(v) + ponto - 1)) Then
                        If ponto > 180 Or UBound(v) - LBound(v) + 1 <= 2 Then
                            DicaPonto = "Alvo do lote " & mLotes.LotePainel()
                        Else
                            DicaPonto = "Alvo do lote deste ponto"
                        End If
                        DicaPonto = DicaPonto & " " & ChrW$(183) & " " & nm & " = " & CStr(CDbl(v(LBound(v) + ponto - 1)))
                    End If
                End If
            End If
            Exit Function
    End Select
    If ponto > 180 Then Exit Function
    Set eng = ThisWorkbook.Worksheets("Eng_Saida")
    If ponto > CLng(Val(CStr(eng.Range("I1").Value))) Then Exit Function      ' alem das corridas publicadas
    r = 2 + ponto
    runv = eng.Cells(r, 2).Value
    If Len(CStr(runv)) = 0 Then Exit Function
    If InStr(1, nm, "conforme", vbTextCompare) > 0 Then
        For k = 1 To serie                        ' qual dos tres X: a ordem das series "Nao conforme (X)"
            If InStr(1, CStr(ch.SeriesCollection(k).Name), "conforme", vbTextCompare) > 0 Then slot = slot + 1
        Next k
        If slot < 1 Or slot > 3 Then Exit Function
        col = (nivel - 1) * 3 + slot - 1
        v = eng.Cells(r, mEstatistica.ENG_COL_X0 + col).Value
        idv = eng.Cells(r, mEstatistica.ENG_COL_IDX0 + col).Value
        dh = eng.Cells(r, mEstatistica.ENG_COL_DHX0 + col).Value
        If Not EhNumero(v) Then Exit Function
        DicaPonto = TextoDica(idv, True, runv, dh, v, CorridaDaDica(eng.Cells(r, mEstatistica.ENG_COL_XPOS0 + col).Value, True))
    ElseIf nm = "Resultado" Or nm = "OK" Or InStr(1, nm, "Viola", vbTextCompare) = 1 Then
        v = eng.Cells(r, mEstatistica.ENG_COL_VALOR0 + nivel - 1).Value
        idv = eng.Cells(r, mEstatistica.ENG_COL_ID0 + nivel - 1).Value
        dh = eng.Cells(r, mEstatistica.ENG_COL_DH0 + nivel - 1).Value
        If Not EhNumero(v) Then Exit Function
        DicaPonto = TextoDica(idv, False, runv, dh, v, CorridaDaDica(eng.Cells(r, mEstatistica.ENG_COL_CORL0 + nivel - 1).Value, False))
    Else
        ' linha de limite ou calibracao: nao e um resultado, so a corrida
        dh = eng.Cells(r, 29).Value                 ' Eng_Saida AC: data da corrida
        DicaPonto = "RUN " & CStr(runv) & IIf(IsDate(dh), " " & ChrW$(183) & " " & Format$(dh, "dd/mm/yyyy"), "")
    End If
sai:
End Function

Private Function EhNumero(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    EhNumero = IsNumeric(v)
End Function

' ADR-081: a corrida do lote vem PRIMEIRO -- e o numero do eixo X. "Corrida 3" (ponto ativo) ou "Entre as corridas 2
' e 3" (X vermelho: inativado, sem numero). Vazio quando a coluna nao foi publicada (arquivo sem o ADR-080/081).
Private Function CorridaDaDica(ByVal pos As Variant, ByVal ehX As Boolean) As String
    Dim k As Long
    If Not EhNumero(pos) Then Exit Function
    If ehX Then
        k = Int(CDbl(pos))
        If k < 1 Then
            CorridaDaDica = "Antes da corrida 1"
        Else
            CorridaDaDica = "Entre as corridas " & k & " e " & (k + 1)
        End If
    Else
        CorridaDaDica = "Corrida " & CLng(pos)
    End If
End Function

' "Corrida 3 · ID 314216 · RUN 25103103 · 25/10/2025 14:32 · 1,234"
' (X: "Entre as corridas 2 e 3 · ID 314216 · INATIVADO · RUN ...").
' ADR-082: manual que substitui um resultado do interfaceamento: "ID MAN_0007 (substitui 54565)" -- DENTRO da parte do
' ID (nenhuma parte nova separada por " · ").
Private Function TextoDica(ByVal idv As Variant, ByVal inativado As Boolean, ByVal runv As Variant, _
                           ByVal dh As Variant, ByVal v As Variant, Optional ByVal corrida As String = "") As String
    Dim s As String, sep As String
    sep = " " & ChrW$(183) & " "
    If Len(corrida) > 0 Then s = corrida & sep
    s = s & "ID " & IdCurto(CStr(idv)) & TextoSubstitui(idv)
    If inativado Then s = s & sep & "INATIVADO"
    s = s & sep & "RUN " & CStr(runv)
    If IsDate(dh) Then s = s & sep & Format$(dh, "dd/mm/yyyy hh:mm")
    TextoDica = s & sep & CStr(CDbl(v))
End Function

' ADR-082: " (substitui 54565)" quando o ponto e um manual que substitui um resultado do interfaceamento; senao "".
' So consulta a Principal - Resultados para ID de manual (a dica roda a cada movimento do mouse). Nunca levanta erro.
Private Function TextoSubstitui(ByVal idv As Variant) As String
    Dim t As String
    On Error GoTo sai
    If UCase$(Left$(Trim$(CStr(idv)), 4)) <> "MAN_" Then Exit Function
    t = mIntegracao.SubstituidoPeloManual(CStr(idv))
    If Len(t) > 0 Then TextoSubstitui = " (substitui " & IdCurto(t) & ")"
sai:
End Function

' ADR-082: o texto de efeito da janela de inativar quando o ponto e um manual que substitui um resultado do
' interfaceamento: o efeito de sempre (sai da media, do DP e do Westgard; nao e apagado) + inativar o manual NAO traz o
' automatico de volta; "" nos outros pontos (fica o texto padrao). Curto: o lblEfeito tem 342 x 27 pt, fonte 8 (duas
' linhas; o padrao tem 127 caracteres) -- o QA confere que cabe (<= 150). Publica: o QA confere o texto.
Public Function EfeitoInativar(ByVal id As String) As String
    Dim t As String
    On Error GoTo sai
    t = mIntegracao.SubstituidoPeloManual(id)
    If Len(t) > 0 Then EfeitoInativar = "O manual sai da média, do DP e do Westgard (não é apagado), mas NÃO traz o " & _
        "automático " & IdCurto(t) & " de volta: para isso, apague o SUBSTITUI_ID."
sai:
End Function

' ADR-082: o texto de efeito da janela de REATIVAR quando o ponto (X vermelho) e um resultado do interfaceamento que um
' manual substitui: reativado, ele vira SUBSTITUIDO_MANUAL e continua fora (o texto padrao diria "volta a media");
' "" nos outros pontos. Mesmo limite de tamanho do EfeitoInativar. Publica: o QA confere o texto.
Public Function EfeitoReativar(ByVal id As String) As String
    Dim t As String
    On Error GoTo sai
    t = mIntegracao.ManualQueSubstitui(id)
    If Len(t) > 0 Then EfeitoReativar = "Substituído pelo manual " & t & ": reativado, continua fora da média, do DP e " & _
        "do gráfico. Para voltar a valer, apague o SUBSTITUI_ID do manual."
sai:
End Function

' O numero que o usuario digita na aba Inativar: "BIO-314216" -> "314216". Manual (MAN_0001) fica inteiro.
Public Function IdCurto(ByVal id As String) As String
    Dim p As Long, resto As String
    id = Trim$(id)
    p = InStrRev(id, "-")
    If p > 0 Then
        resto = Mid$(id, p + 1)
        If Len(resto) > 0 And resto Like String$(Len(resto), "#") Then IdCurto = resto: Exit Function
    End If
    IdCurto = id
End Function

' ADR-081: duplo clique (ou clique direito) num ponto do LJ (clsCht). Mantida com o nome do ADR-070 para o
' evento; o trabalho e da InativarPontoUI. Chamada por evento de grafico: erro nunca sobe (abriria o depurador).
Public Sub InativarPeloGrafico(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long)
    Dim r As String
    On Error Resume Next
    r = InativarPontoUI(nivel, serie, ponto)
End Sub

' ADR-085: todo duplo clique / clique direito num grafico do Painel chega aqui (o clsCht ja cancelou o Excel: nada de
' "Formatar" nem menu de contexto). Ponto de RESULTADO (normal ou X vermelho) = a janela de registrar; qualquer outro
' lugar (eixo, legenda, titulo, linha de limite, calibracao, area vazia) = so a dica de uso na caixa do grafico, por
' alguns segundos, sem MsgBox. Publica e sem dialogo fora da janela: o QA a chama com os mesmos argumentos do evento.
' Devolve "JANELA|<o que a InativarPontoUI devolveu>" ou "DICA|<texto mostrado>". Nunca levanta erro.
Public Function CliqueNoGrafico(ByVal nivel As Long, ByVal elemento As Long, ByVal serie As Long, _
                                ByVal ponto As Long) As String
    Dim msg As String
    On Error Resume Next                          ' DadosDoPonto e InativarPontoUI nao levantam erro
    If elemento = xlSeries And ponto >= 1 Then
        If Left$(mIntegracao.DadosDoPonto(nivel, serie, ponto), 3) = "OK|" Then
            CliqueNoGrafico = "JANELA|" & InativarPontoUI(nivel, serie, ponto)
            Exit Function
        End If
    End If
    msg = "Dê o duplo clique SOBRE um ponto de resultado para registrar não conforme, inativar ou comentar."
    DefinirAviso nivel, msg, TIP_USO, 4
    EscreverTip ChartDoNivel(nivel), msg, TIP_USO
    CliqueNoGrafico = "DICA|" & msg
End Function

' ADR-085: a dica do ponto (DicaPonto, inalterada) e, no ponto com violacao, uma 2a linha com as regras de Westgard
' que o motor marcou nele ("Westgard: 1-3s + R-4s"). E o texto da caixa do grafico (clsCht). "" sem ponto.
Public Function DicaCompleta(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim s As String, rg As String
    On Error GoTo sai
    s = DicaPonto(nivel, serie, ponto)
    If Len(s) = 0 Then Exit Function
    rg = RegrasNoPonto(nivel, serie, ponto)
    If Len(rg) > 0 Then s = s & vbLf & "Westgard: " & rg
    DicaCompleta = s
sai:
End Function

' ADR-085: a identificacao do ponto exatamente como a janela a mostra (FichaDoPonto), sem abrir a janela -- para o
' QA conferir contra o registro de origem. "" quando nao e ponto de resultado.
Public Function FichaPonto(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim d As String
    On Error GoTo sai
    d = mIntegracao.DadosDoPonto(nivel, serie, ponto)
    If Left$(d, 3) = "OK|" Then FichaPonto = FichaDoPonto(nivel, serie, ponto, Split(d, "|"))
sai:
End Function

' ADR-085: as regras de Westgard do ponto do grafico (so os pontos de resultado ATIVO: o X vermelho nao entra em regra).
' Le as marcas que o motor publicou no Eng_Saida para o (nivel, corrida) -- as mesmas que pintam o ponto de vermelho --
' pelo slot da corrida, nunca pela posicao na tela. "" = sem violacao (ou nao e ponto de resultado).
Public Function RegrasNoPonto(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim ch As Chart, nm As String
    On Error GoTo sai
    Set ch = ChartDoNivel(nivel)
    If ch Is Nothing Then Exit Function
    If serie < 1 Or serie > ch.SeriesCollection.Count Then Exit Function
    nm = CStr(ch.SeriesCollection(serie).Name)
    If nm = "Resultado" Or nm = "OK" Or InStr(1, nm, "Viola", vbTextCompare) = 1 Then
        RegrasNoPonto = mEstatistica.RegrasDoPonto(nivel, ponto)
    End If
sai:
End Function

' ADR-081: inativar tem de ser MUITO simples -- duplo clique no ponto, escolher o motivo, confirmar; o grafico ja
' volta atualizado (sem passar pela aba Inativar nem pelo ATUALIZAR DADOS). No X vermelho, a mesma janela reativa.
' Passos: permissao (mSeguranca.PodeInativar); o ponto (mIntegracao.DadosDoPonto); a janela frmInativar (ou a
' resposta do teste); "aguarde" na caixa "tip" (redesenhada antes de a tela congelar); InativarPorId/ReativarPorId;
' AtualizarSoEntradas (sem ler o DB_SEAC); a confirmacao com o lugar do ponto na caixa e na barra de status -- sem
' MsgBox no caminho feliz. Sem o gancho do teste e com Application.Interactive = False (automacao) a janela NAO abre:
' "ERRO|sem interface" (nunca trava um script). Clique durante a operacao ou ate BLOQUEIO_SEG depois de terminar
' (clique que ficou na fila do Excel) = "IGNORADO|...": nao abre a janela de outro ponto.
' A janela e chamada por nome (VBA.UserForms.Add): o projeto compila mesmo sem ela instalada.
' Devolve "OK|<confirmacao>|<segundos>", "CANCELADO|nada mudou", "IGNORADO|<por que>" ou "ERRO|<mensagem>".
Public Function InativarPontoUI(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long) As String
    Dim resp As String, r As String, dif As Double
    On Error GoTo falha
    resp = mRespostaTeste
    mRespostaTeste = ""                           ' o gancho vale uma vez (tambem quando o clique e ignorado)
    If mOcupadoUI Then InativarPontoUI = "IGNORADO|outra inativação em andamento": Exit Function
    If mFimT > 0 Then
        dif = Timer - mFimT
        If dif < 0 Then dif = dif + 86400#        ' passou da meia-noite
        If dif < BLOQUEIO_SEG Then InativarPontoUI = "IGNORADO|clique repetido logo depois de terminar": Exit Function
    End If
    mOcupadoUI = True
    r = InativarPontoNucleo(nivel, serie, ponto, resp)
    mOcupadoUI = False
    If Left$(r, 3) = "OK|" Then
        mFimT = Timer
    ElseIf mAvisoEstilo = TIP_PROGRESSO Then
        EncerrarAviso                             ' parou no meio: o "aguarde" nao fica na tela
    End If
    InativarPontoUI = r
    Exit Function
falha:
    r = TXT_FALHA & " (detalhe: " & Err.Description & ")"
    mOcupadoUI = False
    On Error Resume Next
    EncerrarAviso
    InativarPontoUI = "ERRO|" & r
End Function

' ADR-085 (29/10/2026, pedido do usuario): a janela tem acoes EXPLICITAS, e o texto livre digitado nela e a UNICA
' digitacao. Ponto ATIVO: NC (marcar como NAO CONFORME -- fica no grafico como X vermelho), INATIVAR (sai do grafico)
' ou COMENTAR (so o comentario; o resultado continua ativo). X vermelho: REATIVAR, TIRAR_X (o X sai do grafico; continua
' inativado) ou COMENTAR. O texto vale como MOTIVO na aba Inativar (quando o estado muda) e como COMENTARIO TECNICO do
' ID na aba COMENTARIOS_TECNICOS (mIntegracao.GravarComentarioTecnico: uma linha por ID, sem duplicar; a edicao leva o
' texto anterior ao Audit_Log). REATIVAR nao reescreve o comentario: so marca INATIVADO = NAO na linha do ID (o motivo
' da reativacao vai ao Audit_Log, como antes). Ordem com reversao: o comentario e gravado primeiro; se a gravacao do
' estado falhar, ele volta exatamente ao que era (DesfazerComentario) -- nada fica pela metade.
' COMENTAR nao muda grafico nem estatistica: o comentario e aplicado so na linha do ID da Principal - Resultados
' (AplicarComentarioNaBase, o mesmo texto que o ATUALIZAR publicaria), sem recalcular o resto.
' Gancho do QA (DefinirRespostaTeste): "ACAO|<NC|INATIVAR|COMENTAR|REATIVAR|TIRAR_X>|texto", "CANCELAR" ou o formato do
' ADR-081 "CONFIRMAR|texto|1" (1 = NC com X; 0 = INATIVAR sem X; no X vermelho = REATIVAR).
Private Function InativarPontoNucleo(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long, _
                                     ByVal resp As String) As String
    Dim d As String, p As Variant, reativar As Boolean, conf As Boolean, motivo As String, acao As String
    Dim plotarX As Boolean, r As String, t0 As Double, msg As String, idc As String, f As Object
    Dim titulo As String, detalhe As String, k1 As Long, k2 As Long, teste As Boolean, efeito As String
    Dim rc As String, linC As Long, antes As Variant, comentAtual As String, verbo As String
    On Error GoTo falha
    teste = (Len(resp) > 0)
    If Not mSeguranca.PodeInativar() Then
        msg = mSeguranca.TextoSemPermissaoInativar()
        Application.StatusBar = msg
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, "Registrar no resultado"
        InativarPontoNucleo = "ERRO|" & msg
        Exit Function
    End If
    d = mIntegracao.DadosDoPonto(nivel, serie, ponto)
    If Left$(d, 3) <> "OK|" Then InativarPontoNucleo = d: Exit Function    ' linha de limite, calibracao...
    p = Split(d, "|")
    If UBound(p) < 8 Then InativarPontoNucleo = "ERRO|" & TXT_FALHA & " (detalhe: dados do ponto incompletos)": Exit Function
    reativar = (p(7) = "INATIVADO")               ' o X vermelho: a mesma janela reativa, tira o X ou comenta
    idc = IdCurto(CStr(p(1)))
    If reativar Then
        titulo = "Resultado ID " & idc & " " & ChrW$(8212) & " NÃO CONFORME (X vermelho)"
    Else
        titulo = "Resultado ID " & idc
    End If
    detalhe = FichaDoPonto(nivel, serie, ponto, p)
    ' ADR-082: manual que substitui um automatico -- o texto de efeito avisa que inativar nao o traz de volta;
    ' X de um automatico substituido -- reativado, continua fora (vira SUBSTITUIDO_MANUAL)
    If reativar Then efeito = EfeitoReativar(CStr(p(1))) Else efeito = EfeitoInativar(CStr(p(1)))
    comentAtual = mIntegracao.ComentarioTecnicoDoId(CStr(p(1)))
    If teste Then
        If UCase$(Left$(resp, 5)) = "ACAO|" Then
            k1 = InStr(6, resp, "|")
            If k1 = 0 Then k1 = Len(resp) + 1
            acao = UCase$(Trim$(Mid$(resp, 6, k1 - 6)))
            motivo = Mid$(resp, k1 + 1)
            conf = True
        Else                                      ' ADR-081: "CONFIRMAR|motivo|1" ou "CANCELAR"
            k1 = InStr(resp, "|")
            If k1 = 0 Then k1 = Len(resp) + 1
            conf = (UCase$(Left$(resp, k1 - 1)) = "CONFIRMAR")
            k2 = InStrRev(resp, "|")
            If k2 > k1 Then
                motivo = Mid$(resp, k1 + 1, k2 - k1 - 1)
                plotarX = (Trim$(Mid$(resp, k2 + 1)) <> "0")
            Else
                motivo = Mid$(resp, k1 + 1)
                plotarX = True
            End If
            If reativar Then
                acao = "REATIVAR"
            ElseIf plotarX Then
                acao = "NC"
            Else
                acao = "INATIVAR"
            End If
        End If
    ElseIf Not Application.Interactive Then
        InativarPontoNucleo = "ERRO|sem interface"
        Exit Function
    Else
        Set f = VBA.UserForms.Add("frmInativar")
        f.Preparar reativar, titulo, detalhe, comentAtual
        If Len(efeito) > 0 Then TrocarEfeito f, efeito
        f.Show
        conf = f.Confirmado
        motivo = f.Motivo
        acao = AcaoDaJanela(f)
        Unload f
        Set f = Nothing
    End If
    If Not conf Then InativarPontoNucleo = "CANCELADO|nada mudou": Exit Function
    ' a acao tem de servir ao estado do ponto (o gancho do QA pode pedir qualquer uma)
    Select Case acao
        Case "NC", "INATIVAR"
            If reativar Then InativarPontoNucleo = "ERRO|O resultado ID " & idc & " já está inativado (X vermelho): " & _
                "escolha reativar, tirar o X do gráfico ou só comentar. Nada foi alterado.": Exit Function
        Case "REATIVAR", "TIRAR_X"
            If Not reativar Then InativarPontoNucleo = "ERRO|O resultado ID " & idc & " está ativo: escolha não " & _
                "conforme, inativar ou só comentar. Nada foi alterado.": Exit Function
        Case "COMENTAR"
        Case Else
            InativarPontoNucleo = "ERRO|Escolha uma ação (não conforme, inativar ou comentário). Nada foi alterado."
            Exit Function
    End Select
    If Len(Trim$(Replace(Replace(Replace(motivo, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Then
        InativarPontoNucleo = "ERRO|Escreva o comentário técnico (texto livre). Nada foi alterado."
        Exit Function
    End If
    If acao = "COMENTAR" Then
        InativarPontoNucleo = ComentarNucleo(nivel, CStr(p(1)), idc, motivo, reativar, teste)
        Exit Function
    End If
    plotarX = (acao = "NC")
    t0 = Timer
    Select Case acao
        Case "REATIVAR": verbo = "Reativando"
        Case "TIRAR_X": verbo = "Tirando do gráfico"
        Case Else: verbo = "Inativando"
    End Select
    ' o usuario precisa ver que comecou: a atualizacao congela a tela por alguns segundos
    msg = ChrW$(&H23F3) & " " & verbo & " o ID " & idc & IIf(acao = "NC", " como não conforme", "") & ChrW$(&H2026) & _
          " aguarde (cerca de " & IIf(mEstatistica.NLV >= 3, 25, 15) & " s)"
    MostrarProgresso nivel, msg
    ' ADR-084: o indicador central (frmCarregando) -- comentario + estado + as etapas reais do AtualizarDadosCore
    mCarregando.Abrir verbo & " o resultado ID " & idc, 2 + mIntegracao.ETAPAS_ATUALIZAR
    mCarregando.Etapa "Gravando o comentário técnico"
    If acao = "REATIVAR" Then
        rc = mIntegracao.GravarComentarioTecnico(CStr(p(1)), "", "NÃO", linC, antes)
    Else
        rc = mIntegracao.GravarComentarioTecnico(CStr(p(1)), motivo, "SIM", linC, antes)
    End If
    If Left$(rc, 3) <> "OK|" Then
        mCarregando.Fechar
        msg = Mid$(rc, InStr(rc, "|") + 1)
        Application.StatusBar = False
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, titulo
        InativarPontoNucleo = rc
        Exit Function
    End If
    mCarregando.Etapa "Gravando na aba Inativar"
    Select Case acao
        Case "REATIVAR": r = mIntegracao.ReativarPorId(CStr(p(1)), motivo)
        Case "TIRAR_X": r = mIntegracao.TrocarPlotagemPorId(CStr(p(1)), motivo, False)
        Case Else: r = mIntegracao.InativarPorId(CStr(p(1)), motivo, plotarX, CStr(p(4)))
    End Select
    If Left$(r, 3) <> "OK|" Then
        mIntegracao.DesfazerComentario linC, antes            ' nada pela metade: o comentario volta ao que era
        mCarregando.Fechar
        msg = Mid$(r, InStr(r, "|") + 1)
        Application.StatusBar = False
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, titulo
        InativarPontoNucleo = r
        Exit Function
    End If
    r = mIntegracao.AtualizarSoEntradas(CStr(p(2)))     ' ADR-081: so a incerteza do analito do ponto
    mCarregando.Fechar
    If Left$(r, 3) <> "OK|" Then
        msg = "Gravado (aba Inativar e comentário técnico), mas a atualização não terminou. Clique em ATUALIZAR DADOS " & _
              "para aplicar." & vbCrLf & "(detalhe: " & Mid$(r, InStr(r, "|") + 1) & ")"
        Application.StatusBar = False
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, titulo
        InativarPontoNucleo = "ERRO|" & msg
        Exit Function
    End If
    If acao = "TIRAR_X" Then
        msg = ChrW$(&H2713) & " ID " & idc & " continua inativado " & ChrW$(8212) & " o X saiu do gráfico (para " & _
              "reativar: aba Inativar)"
    Else
        msg = TextoConfirmacao(nivel, CStr(p(1)), idc, reativar, plotarX)
    End If
    MostrarAvisoGrafico nivel, msg
    Application.StatusBar = msg
    InativarPontoNucleo = "OK|" & msg & "|" & Format$(Timer - t0, "0.0")
    Exit Function
falha:
    msg = TXT_FALHA & vbCrLf & "(detalhe: " & Err.Description & ")"
    On Error Resume Next
    mCarregando.Fechar
    If Not f Is Nothing Then Unload f
    Application.StatusBar = False
    If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, "Registrar no resultado"
    InativarPontoNucleo = "ERRO|" & msg
End Function

' ADR-085: so o comentario tecnico -- grafico e estatistica nao mudam, nada e recalculado. Grava na aba
' COMENTARIOS_TECNICOS (uma linha por ID) e aplica o texto na linha do ID da Principal - Resultados; se a aplicacao
' falhar, o comentario volta ao que era. Confirmacao na caixa do grafico, como as outras acoes.
Private Function ComentarNucleo(ByVal nivel As Long, ByVal id As String, ByVal idc As String, ByVal texto As String, _
                                ByVal inativado As Boolean, ByVal teste As Boolean) As String
    Dim rc As String, r As String, linC As Long, antes As Variant, msg As String, t0 As Double
    On Error GoTo falha
    t0 = Timer
    rc = mIntegracao.GravarComentarioTecnico(id, texto, IIf(inativado, "SIM", "NÃO"), linC, antes)
    If Left$(rc, 3) <> "OK|" Then
        msg = Mid$(rc, InStr(rc, "|") + 1)
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, "Comentário técnico"
        ComentarNucleo = rc
        Exit Function
    End If
    r = mIntegracao.AplicarComentarioNaBase(id)
    If Left$(r, 3) <> "OK|" Then
        mIntegracao.DesfazerComentario linC, antes
        msg = Mid$(r, InStr(r, "|") + 1)
        If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, "Comentário técnico"
        ComentarNucleo = r
        Exit Function
    End If
    If Split(rc, "|")(2) = "IGUAL" Then
        msg = ChrW$(&H2713) & " ID " & idc & ": o comentário técnico já era este " & ChrW$(8212) & " nada mudou"
    Else
        msg = ChrW$(&H2713) & " ID " & idc & ": comentário técnico " & IIf(Split(rc, "|")(2) = "ALTERADO", "atualizado", _
              "registrado") & " (aba COMENTARIOS_TECNICOS)"
    End If
    MostrarAvisoGrafico nivel, msg
    Application.StatusBar = msg
    ComentarNucleo = "OK|" & msg & "|" & Format$(Timer - t0, "0.0")
    Exit Function
falha:
    msg = TXT_FALHA & vbCrLf & "(detalhe: " & Err.Description & ")"
    On Error Resume Next
    If linC > 0 Then mIntegracao.DesfazerComentario linC, antes
    If Application.Interactive And Not teste Then MsgBox msg, vbExclamation, "Comentário técnico"
    ComentarNucleo = "ERRO|" & msg
End Function

' ADR-085: a acao escolhida na janela (frmInativar.Acao). Janela de versao anterior (sem a propriedade): o X marcado
' era NC, desmarcado INATIVAR; no X vermelho, REATIVAR.
Private Function AcaoDaJanela(ByVal f As Object) As String
    On Error Resume Next
    AcaoDaJanela = UCase$(CStr(f.Acao))
    If Err.Number <> 0 Then
        Err.Clear
        If f.PlotarX Then AcaoDaJanela = "NC" Else AcaoDaJanela = "INATIVAR"
    End If
End Function

' ADR-085: a identificacao do ponto na janela, toda do registro de origem (Principal - Resultados, pelo ID) e do que o
' motor publicou para a corrida (Eng_Saida) -- nada digitado, nada lido da posicao na tela:
'   ID 314216 · Proteína total · DIMENSION 2 · Nível 1 · Lote 8977
'   Corrida 12 (RUN 26102803) · Data 28/10/2026 · Hora 14:32 · Valor 7,1
'   Gráfico N1 · série "Violação" · Westgard: 1-3s + R-4s
'   Estado atual: ATIVO (entra na média, no DP e no Westgard)
Private Function FichaDoPonto(ByVal nivel As Long, ByVal serie As Long, ByVal ponto As Long, ByVal p As Variant) As String
    Dim sep As String, q As Variant, eq As String, runv As String, s As String, rg As String, nm As String
    Dim dt As String, hr As String
    On Error Resume Next
    sep = " " & ChrW$(183) & " "
    q = Split(mIntegracao.DadosDoId(CStr(p(1))), "|")         ' OK|ID|analito|nivel|equip|lote|RUN|data|estado
    If UBound(q) >= 6 Then
        If q(0) = "OK" Then eq = CStr(q(4)): runv = CStr(q(6))
    End If
    dt = Trim$(CStr(p(5)))
    If InStr(dt, " ") > 0 Then hr = Mid$(dt, InStr(dt, " ") + 1): dt = Left$(dt, InStr(dt, " ") - 1)
    s = "ID " & IdCurto(CStr(p(1))) & sep & p(2)
    If Len(eq) > 0 Then s = s & sep & eq
    s = s & sep & "Nível " & p(3) & sep & "Lote " & p(8) & vbLf
    s = s & TextoCorrida(CStr(p(4)))
    If Len(runv) > 0 Then s = s & " (RUN " & runv & ")"
    s = s & sep & "Data " & dt
    If Len(hr) > 0 Then s = s & sep & "Hora " & hr
    s = s & sep & "Valor " & p(6) & vbLf
    nm = CStr(ChartDoNivel(nivel).SeriesCollection(serie).Name)
    s = s & "Gráfico N" & nivel & sep & "série """ & nm & """"
    If p(7) <> "INATIVADO" Then
        rg = RegrasNoPonto(nivel, serie, ponto)
        s = s & sep & "Westgard: " & IIf(Len(rg) > 0, rg, "sem violação")
    End If
    s = s & vbLf
    If p(7) = "INATIVADO" Then
        s = s & "Estado atual: NÃO CONFORME (inativado, X vermelho) " & ChrW$(8212) & " fora da média, do DP e do Westgard"
    Else
        s = s & "Estado atual: ATIVO " & ChrW$(8212) & " entra na média, no DP e no Westgard"
    End If
    FichaDoPonto = s
End Function

' ADR-082/085: o texto de efeito especial (manual que substitui um automatico) vai para a janela, que o mostra quando a
' acao que ele descreve for escolhida. Janela de outra versao: troca o rotulo direto, como antes.
Private Sub TrocarEfeito(ByVal f As Object, ByVal efeito As String)
    On Error Resume Next
    f.DefinirEfeitoEspecial efeito
    If Err.Number <> 0 Then
        Err.Clear
        f.Controls("lblEfeito").Caption = efeito
    End If
End Sub

' A corrida do campo do DadosDoPonto no texto da dica: "3" -> "Corrida 3"; "entre 2 e 3" -> "Entre as corridas 2 e 3".
Private Function TextoCorrida(ByVal c As String) As String
    If Left$(c, 6) = "entre " Then
        TextoCorrida = "Entre as corridas " & Mid$(c, 7)
    ElseIf c = "antes da 1" Then
        TextoCorrida = "Antes da corrida 1"
    ElseIf Len(c) > 0 Then
        TextoCorrida = "Corrida " & c
    Else
        TextoCorrida = "Corrida ?"
    End If
End Function

' "aguarde" na caixa do grafico e na barra de status, redesenhado AGORA (ScreenUpdating + DoEvents so neste
' instante): o AtualizarSoEntradas congela a tela e, sem isso, o usuario acha que travou. Clique que chegar no
' DoEvents encontra mOcupadoUI e e ignorado. O LimparDicas da republicacao mantem este texto (AvisoGrafico).
Private Sub MostrarProgresso(ByVal nivel As Long, ByVal msg As String)
    Dim tela As Boolean
    On Error Resume Next
    mUltimoProgresso = msg
    DefinirAviso nivel, msg, TIP_PROGRESSO, 180
    EscreverTip ChartDoNivel(nivel), msg, TIP_PROGRESSO
    Application.StatusBar = msg
    tela = Application.ScreenUpdating
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = tela
End Sub

' Para o QA (U14): o ultimo texto de progresso mostrado.
Public Function UltimoAvisoProgresso() As String
    UltimoAvisoProgresso = mUltimoProgresso
End Function

' Confirmacao dizendo onde o ponto foi parar, lida DEPOIS da atualizacao (o eixo do Eng_Saida ja e o novo):
'   "<visto> ID 54565 inativado -- X vermelho entre as corridas 2 e 3; as seguintes foram renumeradas"
'   sem X: "... inativado -- saiu do grafico; ..."; reativar: "... reativado -- voltou a ser a corrida 3".
Private Function TextoConfirmacao(ByVal nivel As Long, ByVal idN As String, ByVal idc As String, _
                                  ByVal reativar As Boolean, ByVal plotarX As Boolean) As String
    Dim ini As String, onde As String, d As String, p As Variant, s As Long, pt As Long, tr As String
    On Error Resume Next
    ini = ChrW$(&H2713) & " ID " & idc
    tr = " " & ChrW$(&H2014) & " "
    If LocalizarNoGrafico(nivel, idN, s, pt) Then
        d = mIntegracao.DadosDoPonto(nivel, s, pt)
        If Left$(d, 3) = "OK|" Then
            p = Split(d, "|")
            If UBound(p) >= 4 Then onde = CStr(p(4))
        End If
    End If
    If reativar Then
        If Len(onde) > 0 And IsNumeric(onde) Then
            TextoConfirmacao = ini & " reativado" & tr & "voltou a ser a corrida " & onde
        Else
            TextoConfirmacao = ini & " reativado"
        End If
    ElseIf Not plotarX Then
        TextoConfirmacao = ini & " inativado" & tr & "saiu do gráfico; as seguintes foram renumeradas"
    ElseIf Left$(onde, 6) = "entre " Then
        TextoConfirmacao = ini & " não conforme (inativado)" & tr & "X vermelho entre as corridas " & Mid$(onde, 7) & _
                           "; as seguintes foram renumeradas"
    ElseIf onde = "antes da 1" Then
        TextoConfirmacao = ini & " não conforme (inativado)" & tr & "X vermelho antes da corrida 1; as seguintes " & _
                           "foram renumeradas"
    Else
        TextoConfirmacao = ini & " não conforme (inativado)" & tr & "as corridas seguintes foram renumeradas"
    End If
End Function

' Serie e ponto do resultado no grafico do nivel (X vermelho ou ponto normal), pelo ID publicado no Eng_Saida.
Private Function LocalizarNoGrafico(ByVal nivel As Long, ByVal idN As String, ByRef serie As Long, _
                                    ByRef ponto As Long) As Boolean
    Dim ch As Chart, eng As Worksheet, n As Long, r As Long, s As Long, k As Long, nm As String
    Dim sNormal As Long, sX(1 To 3) As Long, nx As Long
    On Error GoTo sai
    Set ch = ChartDoNivel(nivel)
    If ch Is Nothing Then Exit Function
    For k = 1 To ch.SeriesCollection.Count
        nm = CStr(ch.SeriesCollection(k).Name)
        If InStr(1, nm, "conforme", vbTextCompare) > 0 Then
            If nx < 3 Then
                nx = nx + 1
                sX(nx) = k
            End If
        ElseIf sNormal = 0 And (nm = "Resultado" Or nm = "OK" Or InStr(1, nm, "Viola", vbTextCompare) = 1) Then
            sNormal = k
        End If
    Next k
    Set eng = ThisWorkbook.Worksheets("Eng_Saida")
    n = CLng(Val(CStr(eng.Range("I1").Value)))
    If n > 180 Then n = 180
    For r = 1 To n
        For s = 1 To nx
            If IdIgual(eng.Cells(2 + r, mEstatistica.ENG_COL_IDX0 + (nivel - 1) * 3 + s - 1).Value, idN) Then
                serie = sX(s): ponto = r: LocalizarNoGrafico = True
                Exit Function
            End If
        Next s
        If sNormal > 0 Then
            If IdIgual(eng.Cells(2 + r, mEstatistica.ENG_COL_ID0 + nivel - 1).Value, idN) Then
                serie = sNormal: ponto = r: LocalizarNoGrafico = True
                Exit Function
            End If
        End If
    Next r
sai:
End Function

Private Function IdIgual(ByVal v As Variant, ByVal idN As String) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    IdIgual = (mIntegracao.NormalizarId(v) = idN)
End Function

Private Function ChartDoNivel(ByVal nivel As Long) As Chart
    On Error Resume Next
    Set ChartDoNivel = ThisWorkbook.Worksheets("Painel").ChartObjects(nivel).Chart
End Function

' ADR-081: gancho do QA -- "CONFIRMAR|motivo|1" (1 = X vermelho, 0 = nao plotar) ou "CANCELAR" no lugar da janela,
' na PROXIMA chamada da InativarPontoUI (vale uma vez: nada fica armado para o usuario).
Public Sub DefinirRespostaTeste(ByVal resposta As String)
    mRespostaTeste = resposta
End Sub

Private Sub DefinirAviso(ByVal nivel As Long, ByVal msg As String, ByVal estilo As Long, ByVal segundos As Long)
    mAvisoTxt = msg
    mAvisoNivel = nivel
    mAvisoEstilo = estilo
    mAvisoAte = Now + TimeSerial(0, 0, segundos)
End Sub

' Tira o aviso (progresso que parou no meio): a caixa volta a dica de uso.
Private Sub EncerrarAviso()
    Dim nv As Long
    On Error Resume Next
    nv = mAvisoNivel
    mAvisoTxt = ""
    mAvisoEstilo = TIP_DICA
    If nv > 0 Then EscreverTip ChartDoNivel(nv), ""
End Sub

' ADR-081: confirmacao na caixa "tip" do grafico do nivel, em verde, por AVISO_SEG segundos: o mouse continua
' sobre o grafico depois da janela, e o proximo movimento reescreveria a dica na hora (o clsCht consulta TipAviso).
Public Sub MostrarAvisoGrafico(ByVal nivel As Long, ByVal msg As String)
    DefinirAviso nivel, msg, TIP_OK, AVISO_SEG
    EscreverTip ChartDoNivel(nivel), msg, TIP_OK
End Sub

' O aviso (progresso ou confirmacao) ainda vale para o grafico deste nivel? ("" = nao)
Public Function AvisoGrafico(ByVal nivel As Long) As String
    If nivel <> mAvisoNivel Or Len(mAvisoTxt) = 0 Then Exit Function
    If Now > mAvisoAte Then
        mAvisoTxt = ""
        Exit Function
    End If
    AvisoGrafico = mAvisoTxt
End Function

' Para o clsCht: com aviso valendo neste grafico, escreve o aviso (no estilo dele) e devolve True.
Public Function TipAviso(ByVal ch As Chart, ByVal nivel As Long) As Boolean
    If Len(AvisoGrafico(nivel)) = 0 Then Exit Function
    EscreverTip ch, mAvisoTxt, mAvisoEstilo
    TipAviso = True
End Function

Public Function DicaDeUso() As String
    DicaDeUso = DICA_USO
End Function

' Caixa "tip" no canto do grafico (a da dica do ponto). Texto vazio = a dica de uso, em cinza: quem nunca usou
' descobre o duplo clique. Estilos: TIP_DICA, TIP_USO, TIP_PROGRESSO, TIP_OK. A caixa cresce com o texto (a
' confirmacao e mais longa que a dica). So escreve se o texto mudou; nao suja a pasta (mesma regra do ADR-061).
Public Sub EscreverTip(ByVal ch As Chart, ByVal msg As String, Optional ByVal estilo As Long = 0)
    Dim tb As Shape, salvo As Boolean
    On Error Resume Next
    If ch Is Nothing Then Exit Sub
    If Len(msg) = 0 Then
        msg = DICA_USO
        estilo = TIP_USO
    End If
    salvo = ThisWorkbook.Saved
    Set tb = Nothing
    Set tb = ch.Shapes("tip")
    If tb Is Nothing Then
        Set tb = ch.Shapes.AddTextbox(1, 6, 2, 480, 16)
        tb.Name = "tip"
        tb.Line.Visible = msoFalse
        tb.Fill.Visible = msoFalse
        tb.TextFrame2.WordWrap = msoFalse
    End If
    If tb.TextFrame2.AutoSize <> msoAutoSizeShapeToFitText Then tb.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    If CStr(tb.TextFrame2.TextRange.Text) <> msg Then
        tb.TextFrame2.TextRange.Text = msg
        With tb.TextFrame2.TextRange.Font
            Select Case estilo
                Case TIP_OK
                    .Size = 10
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = RGB(0, 128, 0)
                Case TIP_PROGRESSO
                    .Size = 10
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = RGB(31, 78, 121)
                Case TIP_USO
                    .Size = 9
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = RGB(118, 118, 118)
                Case Else
                    .Size = 9
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
            End Select
        End With
    End If
    If salvo Then ThisWorkbook.Saved = True
End Sub

' ADR-070 (revisao 07/10/2026): o motor republicou o Eng_Saida (troca de periodo, analito, lote ou
' equipamento) -- o slot sob o mouse pode ter virado OUTRA corrida, e a caixa "tip" continuava com o ID
' antigo (o clsCht so reescreve quando o elemento sob o mouse muda). Apaga o texto da caixa em todos os
' graficos do Painel e muda a geracao das dicas: o proximo movimento do mouse reescreve, mesmo parado no
' mesmo ponto. Chamada pelo AtualizarCalc a cada publicacao. Nao suja a pasta (como a propria dica).
Public Function GeracaoDica() As Long
    GeracaoDica = mGeracaoDica
End Function

Public Sub LimparDicas()
    Dim co As ChartObject
    On Error Resume Next
    mGeracaoDica = mGeracaoDica + 1
    ' ADR-081: durante uma inativacao pelo grafico o "aguarde" continua na caixa daquele nivel; fora dela o aviso sai
    ' (a republicacao pode ser de outro analito) e a caixa vazia mostra a dica de uso (EscreverTip)
    If Not mOcupadoUI Then mAvisoTxt = ""
    For Each co In ThisWorkbook.Worksheets("Painel").ChartObjects
        If Len(AvisoGrafico(co.Index)) > 0 Then
            EscreverTip co.Chart, mAvisoTxt, mAvisoEstilo
        Else
            EscreverTip co.Chart, ""
        End If
    Next co
End Sub

' ADR-070 (revisao): o InputBox do VBA e um dialogo ANSI (cp1252 no Windows pt-BR). Simbolos fora dele
' (>= U+2265, sigma U+03C3, Delta U+0394 ...) chegam como "?" (ou trocados por "melhor aproximacao") e o
' MOTIVO -- justificativa oficial da inativacao -- seria gravado diferente do que o tecnico escreveu, sem
' aviso. Duas defesas, sem UserForm (limitacao registrada no ADR-070):
'  1. o texto MOSTRADO no dialogo (analito, dica) passa por TextoAnsi: os simbolos conhecidos viram texto
'     visivel (">=", "sigma", "delta"...) e o resto vira "[?]" -- nunca um "?" mudo;
'  2. se o motivo digitado contem "?", pergunta antes de gravar: o "?" pode ser um simbolo perdido.
Public Function TextoAnsi(ByVal s As String) As String
    Dim i As Long, c As String, w As Long, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        w = AscW(c) And &HFFFF&
        If w < 256 Then
            out = out & c
        ElseIf InStr(1, CP1252_EXTRA, c, vbBinaryCompare) > 0 Then
            out = out & c                       ' "euro", aspas curvas, travessao...: existem no cp1252
        Else
            Select Case w
                Case &H2265: out = out & ">="
                Case &H2264: out = out & "<="
                Case &H2260: out = out & "<>"
                Case &H3C3, &H3A3: out = out & "sigma"
                Case &H394, &H3B4: out = out & "delta"
                Case &H3BC: out = out & "u"
                Case &H2248: out = out & "~"
                Case &H2192: out = out & "->"
                Case Else: out = out & "[?]"
            End Select
        End If
    Next i
    TextoAnsi = out
End Function

Private Function CP1252_EXTRA() As String
    ' os caracteres de 0x80-0x9F do cp1252 que no Unicode ficam acima de 255
    CP1252_EXTRA = ChrW$(&H20AC) & ChrW$(&H201A) & ChrW$(&H192) & ChrW$(&H201E) & ChrW$(&H2026) & _
                   ChrW$(&H2020) & ChrW$(&H2021) & ChrW$(&H2C6) & ChrW$(&H2030) & ChrW$(&H160) & _
                   ChrW$(&H2039) & ChrW$(&H152) & ChrW$(&H17D) & ChrW$(&H2018) & ChrW$(&H2019) & _
                   ChrW$(&H201C) & ChrW$(&H201D) & ChrW$(&H2022) & ChrW$(&H2013) & ChrW$(&H2014) & _
                   ChrW$(&H2DC) & ChrW$(&H2122) & ChrW$(&H161) & ChrW$(&H203A) & ChrW$(&H153) & _
                   ChrW$(&H17E) & ChrW$(&H178) & ChrW$(&HB7)
End Function
