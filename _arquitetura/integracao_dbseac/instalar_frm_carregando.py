# -*- coding: utf-8 -*-
"""instalar_frm_carregando.py -- ADR-084: a janela frmCarregando (indicador das operacoes demoradas).

Uso:  python instalar_frm_carregando.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Mesmo esquema do frmInativar (instalar_frm_inativar.criar_form): controles e propriedades declarados aqui, codigo do
form em src/comum/frmCarregando.frm.txt, sem .frx versionado. Idempotente (0 mudancas com o form igual). Quem abre,
avanca e fecha a janela e o modulo mCarregando (codigo_atual.py); sem a janela o projeto compila e a operacao
segue so com a barra de status.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import instalar_integracao as ii  # noqa: E402
import instalar_frm_inativar as fi  # noqa: E402

NOME = 'frmCarregando'
FONTE = os.path.join(AQUI, 'src', 'comum', 'frmCarregando.frm.txt')
FONTE_UI = 'Segoe UI'
# BGR (o MSForms le OLE_COLOR = 0x00BBGGRR)
BRANCO = 0xFFFFFF
AZUL = 0x794E1F                        # RGB(31, 78, 121): o azul dos status do Painel (ADR-078)
TRILHO = 0xEBE5DD                      # RGB(221, 229, 235): cinza-azulado claro
CINZA = 0x606060

# Janela pequena e sem borda de redimensionar; medidas em multiplos de 3 pt (o MSForms arredonda ao pixel).
FORM = {'Caption': 'QC_INI', 'Width': 330, 'Height': 132, 'StartUpPosition': 1, 'BackColor': BRANCO}

CONTROLES = [
    ('Forms.Label.1', 'lblTitulo', {
        'Left': 15, 'Top': 12, 'Width': 294, 'Height': 18, 'Caption': 'Processando', 'BackStyle': 0,
        'ForeColor': AZUL, 'Font.Name': FONTE_UI, 'Font.Size': 11, 'Font.Bold': True, 'WordWrap': False}),
    ('Forms.Label.1', 'lblEtapa', {
        'Left': 15, 'Top': 33, 'Width': 294, 'Height': 27, 'Caption': '', 'BackStyle': 0, 'WordWrap': True,
        'Font.Name': FONTE_UI, 'Font.Size': 9}),
    # trilho e barra: dois rotulos sem texto (a barra por cima do trilho)
    ('Forms.Label.1', 'lblTrilho', {
        'Left': 15, 'Top': 66, 'Width': 294, 'Height': 6, 'Caption': '', 'BackStyle': 1, 'BackColor': TRILHO}),
    ('Forms.Label.1', 'lblBarra', {
        'Left': 15, 'Top': 66, 'Width': 0, 'Height': 6, 'Caption': '', 'BackStyle': 1, 'BackColor': AZUL}),
    ('Forms.Label.1', 'lblPasso', {
        'Left': 15, 'Top': 78, 'Width': 294, 'Height': 15, 'Caption': '', 'BackStyle': 0, 'ForeColor': CINZA,
        'Font.Name': FONTE_UI, 'Font.Size': 8}),
]


def criar_form(wb, log=print):
    return fi.criar_form(wb, log, nome=NOME, fonte=FONTE, form=FORM, controles=CONTROLES)


def main(produto, caminho, salvar=True):
    import xlh
    import codigo_atual
    from instalar_seguranca_usuarios import foto_protecao, devolver_protecao
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    ii.log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ii.log('1. janela frmCarregando (ADR-084)')
        n = criar_form(wb, ii.log)
        ii.log('2. VBA (todo o codigo atual das fontes: codigo_atual.py; compila o form junto)')
        codigo_atual.instalar(ex, wb, produto, ii.log)
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            ii.log(f'salvo ({n} mudanca(s) no form)')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
