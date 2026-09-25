Attribute VB_Name = "Planilha7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' Painel (Bioquimica) -- ADR-049/050/054. O MESMO mapa de celulas da Hematologia:
'   B3 / C3 ... analito: B3 pelo spinner, C3 escolhendo o nome na lista
'   H3 ........ lote EM ANALISE (lista de lotes cadastrados)
'   G3 / G4 ... periodo De/Ate (as datas que o Calc realmente filtra)
'   K3 / K4 ... Ano / Periodo: atalho que PREENCHE G3/G4
'   M4 ........ provedor do controle externo deste analito (so a Bioquimica)
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
    If Not Intersect(Target, Me.Range("M4")) Is Nothing Then
        Application.EnableEvents = False
        mEstatistica.GravarProvedorAnalito Me.Range("C3").Value, CStr(Me.Range("M4").Value)
        Me.Range("M4").Formula = "=IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($C$3,Analitos!$A$4:$A$43,0)),""CAP"")"
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Activate()
    mEstatistica.GarantirMotorDoPainel
    AtualizarEixos
    mUI.AjustarGraficos
    HookCharts
End Sub

' ADR-054: o grafico acompanha a largura visivel. O Excel nao avisa quando o
' zoom muda; o primeiro clique depois do zoom reajusta (e custa dois numeros).
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    mUI.AjustarGraficos
End Sub
