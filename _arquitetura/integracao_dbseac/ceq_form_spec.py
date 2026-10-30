# -*- coding: utf-8 -*-
"""ADR-087: a janela frmCEQ (registro do Controle Externo) -- controles e propriedades, para instalar_frm_inativar.criar_form.

Medidas em multiplos de 3 pt. Duas colunas de grupos (1 Rodada | 2 Analito e amostra; 3 Resultado | 4 Avaliacao e
fonte), a faixa de mensagem (lblMsg) entre os campos e a lista, a lista dos registros da rodada e os botoes no fim.
Nenhum controle se sobrepoe a outro (conferido pelo testes/qa_adr087.py: retangulos e texto medido com a fonte real).
"""
import os

AQUI = os.path.dirname(os.path.abspath(__file__))
FONTE_CODIGO = os.path.join(AQUI, 'src', 'comum', 'frmCEQ.frm.txt')
UI = 'Segoe UI'
CINZA = 0x5A5A5A
AZUL = 0x79481F                                   # BGR de RGB(31, 72, 121)

FORM = {'Caption': 'Controle Externo da Qualidade (CEQ) — registrar resultado', 'Width': 606, 'Height': 486,
        'StartUpPosition': 1}


def _lbl(nome, cap, left, top, width, height=15, bold=False, size=9, cor=None, wrap=False):
    p = {'Left': left, 'Top': top, 'Width': width, 'Height': height, 'Caption': cap, 'Font.Name': UI,
         'Font.Size': size, 'Font.Bold': bold, 'WordWrap': wrap}
    if cor is not None:
        p['ForeColor'] = cor
    return ('Forms.Label.1', nome, p)


def _txt(nome, left, top, width, tab, maxlen=0):
    p = {'Left': left, 'Top': top, 'Width': width, 'Height': 18, 'Font.Name': UI, 'Font.Size': 9, 'TabIndex': tab}
    if maxlen:
        p['MaxLength'] = maxlen
    return ('Forms.TextBox.1', nome, p)


def _cbo(nome, left, top, width, tab, lista_fixa=False):
    p = {'Left': left, 'Top': top, 'Width': width, 'Height': 18, 'Font.Name': UI, 'Font.Size': 9, 'TabIndex': tab,
         # texto livre sem completar: "Colesterol" (novo) viraria "Colesterol HDL" (existente) ao sair do campo
         'Style': 2 if lista_fixa else 0, 'MatchEntry': 1 if lista_fixa else 2}
    return ('Forms.ComboBox.1', nome, p)


# colunas: esquerda 12..285 (rotulo 90 + campo 180); pares do resultado; direita 306..588 (rotulo 90 + campo 189)
CONTROLES = [
    _lbl('lblTitulo', 'Registrar resultado do Controle Externo (CEQ)', 12, 9, 576, 21, bold=True, size=11, cor=AZUL),
    _lbl('lblSub', 'Campos com * são obrigatórios. O resultado vai para a aba do provedor e a EQA_Base é '
                   'consolidada em seguida.', 12, 30, 576, 15, size=8, cor=CINZA),
    _lbl('lblG1', '1 · RODADA', 12, 51, 273, 15, bold=True, cor=AZUL),
    _lbl('lblG2', '2 · ANALITO E AMOSTRA', 306, 51, 282, 15, bold=True, cor=AZUL),
    _lbl('lblProvedor', 'Provedor *', 12, 72, 90),
    _cbo('cboProvedor', 105, 69, 180, 0, lista_fixa=True),
    _lbl('lblRodada', 'Rodada / Survey *', 12, 96, 90),
    _cbo('cboRodada', 105, 93, 180, 1),
    _lbl('lblAno', 'Ano *', 12, 120, 90),
    _txt('txtAno', 105, 117, 60, 2, 4),
    _lbl('lblAnalito', 'Analito *', 306, 72, 90),
    _cbo('cboAnalito', 399, 69, 189, 3),
    _lbl('lblAmostra', 'Amostra *', 306, 96, 90),
    _txt('txtAmostra', 399, 93, 120, 4, 40),
    _lbl('lblUnidade', 'Unidade', 306, 120, 90),
    _txt('txtUnidade', 399, 117, 120, 5, 30),
    _lbl('lblG3', '3 · RESULTADO', 12, 147, 273, 15, bold=True, cor=AZUL),
    _lbl('lblG4', '4 · AVALIAÇÃO E FONTE', 306, 147, 282, 15, bold=True, cor=AZUL),
    _lbl('lblResultado', 'Seu resultado *', 12, 168, 78),
    _txt('txtResultado', 93, 165, 54, 6),
    _lbl('lblAlvo', 'Valor-alvo', 153, 168, 69),
    _txt('txtAlvo', 225, 165, 60, 7),
    _lbl('lblSD', 'SD', 12, 192, 78),
    _txt('txtSD', 93, 189, 54, 8),
    _lbl('lblSDI', 'SDI', 153, 192, 69),
    _txt('txtSDI', 225, 189, 60, 9),
    _lbl('lblLimInf', 'Limite inferior', 12, 216, 78),
    _txt('txtLimInf', 93, 213, 54, 10),
    _lbl('lblLimSup', 'Limite superior', 153, 216, 69),
    _txt('txtLimSup', 225, 213, 60, 11),
    _lbl('lblNLabs', 'Nº laboratórios', 12, 240, 78),
    _txt('txtNLabs', 93, 237, 54, 12),
    _lbl('lblAvaliacao', 'Avaliação *', 306, 168, 90),
    _cbo('cboAvaliacao', 399, 165, 189, 13),
    _lbl('lblPagina', 'Página do PDF', 306, 192, 90),
    _txt('txtPagina', 399, 189, 60, 14),
    _lbl('lblArquivo', 'Arquivo fonte', 306, 216, 90),
    _txt('txtArquivo', 399, 213, 189, 15, 255),
    _lbl('lblMsg', '', 12, 267, 576, 27, size=9, wrap=True),
    _lbl('lblLista', 'REGISTROS DESTA RODADA', 12, 300, 273, 15, bold=True, cor=AZUL),
    _lbl('lblQtd', '', 306, 300, 282, 15, size=8, cor=CINZA),
    # IntegralHeight antes da altura (com ele ligado o MSForms arredonda a altura para linhas inteiras); a lista e so
    # de consulta: fica fora da ordem de tabulacao (o ultimo TabIndex e reordenado pelo MSForms)
    ('Forms.ListBox.1', 'lstRegs', {'Left': 12, 'Top': 318, 'Width': 576, 'IntegralHeight': False, 'Height': 93,
                                    'ColumnCount': 4, 'ColumnWidths': '200 pt;110 pt;110 pt;140 pt', 'Font.Name': UI,
                                    'Font.Size': 9, 'TabStop': False}),
    ('Forms.CommandButton.1', 'btnSalvar', {'Left': 318, 'Top': 420, 'Width': 168, 'Height': 27,
                                            'Caption': 'Salvar e próximo', 'Default': True, 'Font.Name': UI,
                                            'Font.Size': 9, 'Font.Bold': True, 'TabIndex': 16}),
    ('Forms.CommandButton.1', 'btnFechar', {'Left': 492, 'Top': 420, 'Width': 96, 'Height': 27, 'Caption': 'Fechar',
                                            'Cancel': True, 'Font.Name': UI, 'Font.Size': 9, 'TabIndex': 17}),
]
