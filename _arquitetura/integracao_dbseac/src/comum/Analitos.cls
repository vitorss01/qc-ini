Option Explicit
' Aba Analitos (ADR-074) -- UM BLOCO POR LOTE: Media/DP de cada nivel (E:J das linhas de analito, gravados
' na hora na LotesStore, com trilha), limites de controle, anos de referencia CLIA/VB (linha ANO DE
' REFERENCIA de cada bloco) e as especificacoes, ET e Sigma do lote. Ver mLotesBlocos.
Private Sub Worksheet_Change(ByVal Target As Range)
    mLotesBlocos.BlocoEditado Target
End Sub

' Lote cadastrado depois da ultima visita ganha o seu bloco (barato quando nao ha nada a fazer).
Private Sub Worksheet_Activate()
    On Error Resume Next
    mLotesBlocos.SincronizarBlocos
End Sub
