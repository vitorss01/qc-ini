' Abas EQA.CAP_Dados e EQA.Controllab_Dados (o MESMO modulo nas duas) -- ADR-089 (OC-0028, OC-0032).
'   Edicao direta na tabela (quem tem a aba liberada): marca a edicao pendente (Z4, oculto) -- o Z3 avisa "HA EDICAO
'   NAO CONSOLIDADA" ate o Consolidar ou o ATUALIZAR DADOS -- e registra CEQ_ALTERADO com o antes e o depois.
'   O "antes" e o valor guardado quando a celula foi selecionada. A janela do CEQ grava com os eventos desligados
'   (ela mesma consolida e registra a trilha).
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    mCEQForm.GuardarSelecaoCEQ Target
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    mCEQForm.EdicaoDiretaCEQ Target
End Sub
