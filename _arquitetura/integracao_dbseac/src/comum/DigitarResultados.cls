' Digitar Resultados -- ADR-057: aba de ENTRADA da camada de dados. O evento so prepara a linha (ID normalizado,
' carimbo, ID MAN_nnnn); quem decide o que o dado e, e o motor (DB_CQ_FINAL).
' ADR-098: a selecao guarda o retrato dos IDs da aba (mIntegracao.EntradaSelecionada) -- e ele que diz, no Change
' seguinte, se o ID de uma linha CHEGOU agora (recorte/colar) ou so foi tocado na propria linha.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    mIntegracao.EntradaMudou Me, Target
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    mIntegracao.EntradaSelecionada Me
End Sub
