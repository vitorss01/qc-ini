VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Painel. Antes do Sprint HARDENING 1, trocar o analito ou o filtro so chamava
' AtualizarEixos: o Calc recalculava tudo sozinho por formula, inclusive as
' regras de Westgard, e ficava correto.
'
' Depois do Marco 2 as regras passaram a vir de Eng_Saida, que so o motor
' preenche. Sem recalcular aqui, o painel exibiria as regras da ULTIMA execucao
' do motor -- de outro analito ou de outro periodo. A chave ANALITO|RUN impede
' que isso vire veredicto falso (o MATCH falha e a celula fica vazia), e esta
' chamada e o que faz o dado aparecer de fato.
'
' B3        = seletor de analito
' G3,G4     = filtro de data (De / Ate)
' M3,M4,N3,N4 = seletores de trimestre
' Todos entram no que o motor publica: o analito muda a serie inteira, e o
' filtro muda a coluna Filtro e a estatistica do Painel.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("B3,G3,G4,M3,M4,N3,N4")) Is Nothing Then
        Application.EnableEvents = False
        RecalcularAnalitoAtual
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Activate()
    AtualizarEixos
    HookCharts
End Sub
