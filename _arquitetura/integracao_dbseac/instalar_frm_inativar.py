# -*- coding: utf-8 -*-
"""instalar_frm_inativar.py -- ADR-081: a janela frmInativar (Inativar / Reativar resultado pelo grafico).

Uso:  python instalar_frm_inativar.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

O UserForm e criado AQUI, sem .frx versionado: os controles e as propriedades estao em FORM/CONTROLES, o codigo
do form em src/comum/frmInativar.frm.txt. criar_form(wb, log) e idempotente: com o form igual ao daqui (mesmos
controles, mesmas propriedades, codigo identico a fonte) nao toca em nada e devolve 0 mudancas; senao refaz os
controles e o codigo.

A ordem com o codigo_atual nao importa: o mUI abre a janela por nome (VBA.UserForms.Add("frmInativar")), o projeto
compila sem ela, e sem ela o duplo clique responde "nao foi possivel concluir" em vez de abrir a janela.

main(): instala o form, instala todo o VBA atual das fontes (codigo_atual.py, que EXIGE compilacao -- o codigo do
form entra nessa compilacao), devolve a protecao de cada aba e a da estrutura como estavam. Falha = nada salvo.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import instalar_integracao as ii  # noqa: E402

NOME = 'frmInativar'
FONTE = os.path.join(AQUI, 'src', 'comum', 'frmInativar.frm.txt')
VBEXT_CT_MSFORM = 3
FONTE_UI = 'Segoe UI'
CINZA = 0x606060                       # BGR = RGB (cinza neutro)

# Medidas em multiplos de 3 pt (4 px a 96 dpi): o MSForms arredonda ao pixel, e a conferencia de idempotencia compara
# o que leu de volta. ADR-085: janela maior (ficha do resultado + 3 acoes + comentario), ainda nao redimensionavel.
FORM = {'Caption': 'Registrar no resultado', 'Width': 438, 'Height': 420, 'StartUpPosition': 1}   # 1 = centro do Excel

# (ProgID, nome, propriedades). 'Font.X' = propriedade da fonte do controle.
# ADR-084 (28/10/2026): cboMotivo + txtComplemento deram lugar a UM campo de texto livre (txtMotivo).
# ADR-085 (29/10/2026): chkX + lblAvisoX deram lugar a TRES acoes explicitas (optNC / optInat / optComent, mesmo
# GroupName, nenhuma marcada de inicio); lblDetalhe virou a ficha do resultado (4 linhas) e lblAcao o titulo das acoes.
CONTROLES = [
    ('Forms.Label.1', 'lblTitulo', {
        'Left': 12, 'Top': 9, 'Width': 408, 'Height': 21, 'Caption': 'Resultado',
        'Font.Name': FONTE_UI, 'Font.Size': 11, 'Font.Bold': True, 'WordWrap': False}),
    ('Forms.Label.1', 'lblDetalhe', {
        'Left': 12, 'Top': 33, 'Width': 408, 'Height': 63, 'Caption': '',
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'WordWrap': True}),
    ('Forms.Label.1', 'lblAcao', {
        'Left': 12, 'Top': 102, 'Width': 408, 'Height': 15, 'Caption': 'O que fazer com este resultado?',
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'Font.Bold': True}),
    ('Forms.OptionButton.1', 'optNC', {
        'Left': 18, 'Top': 120, 'Width': 402, 'Height': 27, 'Caption': 'A', 'GroupName': 'acao', 'WordWrap': True,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 0}),
    ('Forms.OptionButton.1', 'optInat', {
        'Left': 18, 'Top': 150, 'Width': 402, 'Height': 27, 'Caption': 'B', 'GroupName': 'acao', 'WordWrap': True,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 1}),
    ('Forms.OptionButton.1', 'optComent', {
        'Left': 18, 'Top': 180, 'Width': 402, 'Height': 18, 'Caption': 'C', 'GroupName': 'acao', 'WordWrap': True,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 2}),
    ('Forms.Label.1', 'lblEfeito', {
        'Left': 12, 'Top': 204, 'Width': 408, 'Height': 30, 'Caption': '', 'WordWrap': True, 'ForeColor': CINZA,
        'Font.Name': FONTE_UI, 'Font.Size': 8}),
    ('Forms.Label.1', 'lblMotivo', {
        'Left': 12, 'Top': 240, 'Width': 408, 'Height': 15, 'Caption': 'Comentário técnico (obrigatório)',
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'Font.Bold': True}),
    # ADR-084: texto livre, varias linhas (Enter quebra a linha; Tab vai ao proximo controle). Sem lista de frases.
    ('Forms.TextBox.1', 'txtMotivo', {
        'Left': 12, 'Top': 255, 'Width': 408, 'Height': 90, 'MultiLine': True, 'WordWrap': True,
        'EnterKeyBehavior': True, 'TabKeyBehavior': False, 'ScrollBars': 2, 'MaxLength': 1000,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 3}),
    ('Forms.CommandButton.1', 'btnOK', {
        'Left': 231, 'Top': 354, 'Width': 93, 'Height': 27, 'Caption': 'Salvar', 'Default': False,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'Font.Bold': True, 'TabIndex': 4}),
    ('Forms.CommandButton.1', 'btnCancelar', {
        'Left': 327, 'Top': 354, 'Width': 93, 'Height': 27, 'Caption': 'Cancelar', 'Cancel': True,   # Esc
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 5}),
]
NUMERICAS = {'Left', 'Top', 'Width', 'Height', 'Font.Size'}


def _ler(obj, prop):
    for p in prop.split('.')[:-1]:
        obj = getattr(obj, p)
    return getattr(obj, prop.split('.')[-1])


def _gravar(obj, prop, valor):
    for p in prop.split('.')[:-1]:
        obj = getattr(obj, p)
    setattr(obj, prop.split('.')[-1], valor)


def _igual(prop, atual, quer):
    if prop in NUMERICAS:
        try:
            return abs(float(atual) - float(quer)) <= 1.0
        except (TypeError, ValueError):
            return False
    if isinstance(quer, bool):
        return bool(atual) == quer
    if isinstance(quer, int):
        try:
            return int(atual) == quer
        except (TypeError, ValueError):
            return False
    return str(atual or '') == str(quer)


def _normalizar(t):
    import codigo_atual
    return codigo_atual.normalizar(t).lower()


def _diferencas(comp, texto, form=None, controles=None):
    """O que o form no arquivo tem de diferente do daqui (lista vazia = igual)."""
    form = FORM if form is None else form
    controles = CONTROLES if controles is None else controles
    dif = []
    # 22/10/2026: num arquivo REABERTO, Properties do UserForm so responde depois que o designer e tocado -- sem isto
    # Caption/Width/Height/StartUpPosition davam erro ('?') e o form era refeito a cada execucao (nunca 0 mudancas)
    try:
        comp.Designer
    except Exception:                                  # noqa: BLE001
        pass
    for k, v in form.items():
        try:
            if not _igual(k, comp.Properties(k).Value, v):
                dif.append(f'form.{k}')
        except Exception:                              # noqa: BLE001
            dif.append(f'form.{k}?')
    ctls = {c.Name: c for c in comp.Designer.Controls}
    if sorted(ctls) != sorted(n for _, n, _ in controles):
        dif.append(f'controles {sorted(ctls)}')
        return dif
    for _, n, props in controles:
        for k, v in props.items():
            try:
                if not _igual(k, _ler(ctls[n], k), v):
                    dif.append(f'{n}.{k}')
            except Exception:                          # noqa: BLE001
                dif.append(f'{n}.{k}?')
    cm = comp.CodeModule
    atual = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ''
    # VBA e insensivel a maiusculas e o VBE reescreve a caixa de identificador ja conhecido
    if _normalizar(atual) != _normalizar(texto):
        dif.append('codigo')
    return dif


def criar_form(wb, log=print, nome=None, fonte=None, form=None, controles=None):
    """Cria/atualiza o UserForm (padrao: frmInativar). Idempotente. Devolve o numero de mudancas (0 = ja estava igual).
    ADR-084: nome/fonte/form/controles servem a outras janelas declarativas (frmCarregando)."""
    NOME_ = NOME if nome is None else nome
    form = FORM if form is None else form
    controles = CONTROLES if controles is None else controles
    texto = ii.codigo(FONTE if fonte is None else fonte)
    vbp = wb.VBProject
    nomes = [c.Name for c in vbp.VBComponents]
    if NOME_ in nomes:
        comp = vbp.VBComponents(NOME_)
        if comp.Type != VBEXT_CT_MSFORM:
            raise SystemExit(f'{NOME_} existe e nao e um UserForm (tipo {comp.Type}) -- nada foi feito')
        dif = _diferencas(comp, texto, form, controles)
        if not dif:
            log(f'   {NOME_}: igual ao instalador (0 mudancas)')
            return 0
        log(f'   {NOME_}: refeito ({", ".join(dif)})')
    else:
        comp = vbp.VBComponents.Add(VBEXT_CT_MSFORM)
        comp.Name = NOME_
        log(f'   {NOME_}: criado')
    # codigo fora antes de mexer nos controles: evento orfao de controle removido nao chega a existir
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    for k, v in form.items():
        comp.Properties(k).Value = v
    des = comp.Designer
    for n in [c.Name for c in des.Controls]:
        des.Controls.Remove(n)
    for progid, n, props in controles:
        ctl = des.Controls.Add(progid, n, True)
        for k, v in props.items():
            _gravar(ctl, k, v)
    cm.AddFromString(texto)
    dif = _diferencas(comp, texto, form, controles)
    if dif:
        raise SystemExit(f'{NOME_} ficou diferente do instalador: {dif} -- nada foi salvo')
    return 1


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
        ii.log('1. janela frmInativar (ADR-081)')
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
