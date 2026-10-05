' Painel (Bioquimica) -- ADR-049/050/054. O MESMO mapa de celulas da Hematologia:
'   B3 / C3 ... analito: B3 pelo spinner, C3 escolhendo o nome na lista
'   H3 ........ lote EM ANALISE (lista de lotes cadastrados)
'   G3 / G4 ... periodo De/Ate (as datas que o Calc realmente filtra)
'   K3 / K4 ... Ano / Periodo: atalho que PREENCHE G3/G4
'   M4 ........ provedor do controle externo deste analito (so a Bioquimica)
'   L4 ........ EQUIPAMENTO em analise (selEquipamento; so a Bioquimica) -- ADR-057
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("K3:K4")) Is Nothing Then
        mApp.AplicarFiltroAnoMes
        mEstatistica.FiltroMudou
    End If
    If Not Intersect(Target, Me.Range("H3")) Is Nothing Then
        ' ADR-049: lote em analise. Carrega media/DP/limites DESSE lote.
        mLotes.TrocarLoteAnalise
    End If
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
        ' ADR-050: escolher o analito pelo NOME (lista), sem 30 cliques no spinner
        mApp.AnalitoEscolhido Me.Range("C3")
    End If
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        ' analito: refaz o motor so se ele ficou para tras (o spinner ja refaz)
        mEstatistica.PainelMudou
    End If
    If Not Intersect(Target, Me.Range("G3,G4")) Is Nothing Then
        ' periodo: a janela de corridas do motor depende do fim do periodo
        mEstatistica.FiltroMudou
    End If
    If Not Intersect(Target, Me.Range("L4")) Is Nothing Then
        ' ADR-057: DIMENSION 1 / DIMENSION 2 -- cada equipamento e uma serie propria
        mEstatistica.EquipMudou
    End If
    If Not Intersect(Target, Me.Range("M4")) Is Nothing Then
        Application.EnableEvents = False
        mEstatistica.GravarProvedorAnalito Me.Range("C3").Value, CStr(Me.Range("M4").Value)
        Me.Range("M4").Formula = "=IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($C$3,Analitos!$A$4:$A$43,0)),""CAP"")"
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Activate()
    mUI.SpinnerLimites                    ' ADR-062
    mEstatistica.AquecerMotor             ' ADR-062: quente sai na hora
    mEstatistica.GarantirMotorDoPainel
    AtualizarEixos
    mUI.PainelEncaixar                    ' ADR-063: todos os graficos inteiros na tela
    mUI.AjustarGraficos
    HookCharts
    mUI.GrafVigiaLigar                    ' ADR-061: graficos acompanham o zoom sem clique
End Sub

' ADR-061: sair do Painel (navegacao; logout -- LockApp ativa o Login) para o vigia.
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    mUI.GrafVigiaDesligar
End Sub

' ADR-054/061: o clique ainda reajusta na hora; o vigia (mUI) cobre o zoom sem clique.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    mUI.AjustarGraficos
    mUI.GrafVigiaLigar                    ' ADR-061: religa se a corrente morreu (reset, fechamento cancelado)
End Sub
