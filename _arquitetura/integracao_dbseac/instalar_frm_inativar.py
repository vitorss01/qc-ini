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
LARANJA = 0x0050C0                     # BGR: RGB(192, 80, 0) -- aviso do X desmarcado

# Janela <= 380 x 300 pt, nao redimensionavel (UserForm nao tem borda de redimensionar). Medidas em multiplos de
# 3 pt (4 px a 96 dpi): o MSForms arredonda ao pixel, e a conferencia de idempotencia compara o que leu de volta.
FORM = {'Caption': 'Inativar resultado', 'Width': 372, 'Height': 297, 'StartUpPosition': 1}   # 1 = centro do Excel

# (ProgID, nome, propriedades). 'Font.X' = propriedade da fonte do controle. TabIndex: motivo primeiro.
CONTROLES = [
    ('Forms.Label.1', 'lblTitulo', {
        'Left': 12, 'Top': 9, 'Width': 342, 'Height': 21, 'Caption': 'Inativar o resultado',
        'Font.Name': FONTE_UI, 'Font.Size': 11, 'Font.Bold': True, 'WordWrap': False}),
    ('Forms.Label.1', 'lblDetalhe', {
        'Left': 12, 'Top': 33, 'Width': 342, 'Height': 27, 'Caption': '',
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'WordWrap': True}),
    ('Forms.Label.1', 'lblMotivo', {
        'Left': 12, 'Top': 63, 'Width': 342, 'Height': 15, 'Caption': 'Motivo',
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'Font.Bold': True}),
    ('Forms.ComboBox.1', 'cboMotivo', {
        'Left': 12, 'Top': 78, 'Width': 342, 'Height': 21, 'Style': 2,          # 2 = so a lista (sem digitar)
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'ListRows': 8, 'TabIndex': 0}),
    ('Forms.Label.1', 'lblComplemento', {
        'Left': 12, 'Top': 105, 'Width': 342, 'Height': 15, 'Caption': 'Complemento (opcional)',
        'Font.Name': FONTE_UI, 'Font.Size': 9}),
    ('Forms.TextBox.1', 'txtComplemento', {
        'Left': 12, 'Top': 120, 'Width': 342, 'Height': 33, 'MultiLine': True, 'WordWrap': True,
        'EnterKeyBehavior': False, 'ScrollBars': 2, 'MaxLength': 300,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 1}),
    ('Forms.CheckBox.1', 'chkX', {
        'Left': 12, 'Top': 159, 'Width': 342, 'Height': 18, 'Caption': 'Mostrar como X vermelho no gráfico',
        'Value': True, 'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 2}),
    # ADR-081 (UX): aparece so com a caixa do X desmarcada (chkX_Change)
    ('Forms.Label.1', 'lblAvisoX', {
        'Left': 12, 'Top': 177, 'Width': 342, 'Height': 24, 'Caption': '', 'WordWrap': True, 'ForeColor': LARANJA,
        'Visible': False, 'Font.Name': FONTE_UI, 'Font.Size': 8, 'Font.Bold': True}),
    ('Forms.Label.1', 'lblEfeito', {
        'Left': 12, 'Top': 204, 'Width': 342, 'Height': 27, 'Caption': '', 'WordWrap': True, 'ForeColor': CINZA,
        'Font.Name': FONTE_UI, 'Font.Size': 8}),
    ('Forms.CommandButton.1', 'btnOK', {
        'Left': 162, 'Top': 237, 'Width': 93, 'Height': 27, 'Caption': 'Inativar', 'Default': False,
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'Font.Bold': True, 'TabIndex': 3}),
    ('Forms.CommandButton.1', 'btnCancelar', {
        'Left': 261, 'Top': 237, 'Width': 93, 'Height': 27, 'Caption': 'Cancelar', 'Cancel': True,   # Esc
        'Font.Name': FONTE_UI, 'Font.Size': 9, 'TabIndex': 4}),
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


def _diferencas(comp, texto):
    """O que o form no arquivo tem de diferente do daqui (lista vazia = igual)."""
    dif = []
    for k, v in FORM.items():
        try:
            if not _igual(k, comp.Properties(k).Value, v):
                dif.append(f'form.{k}')
        except Exception:                              # noqa: BLE001
            dif.append(f'form.{k}?')
    ctls = {c.Name: c for c in comp.Designer.Controls}
    if sorted(ctls) != sorted(n for _, n, _ in CONTROLES):
        dif.append(f'controles {sorted(ctls)}')
        return dif
    for _, n, props in CONTROLES:
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


def criar_form(wb, log=print):
    """Cria/atualiza o UserForm frmInativar. Idempotente. Devolve o numero de mudancas (0 = ja estava igual)."""
    texto = ii.codigo(FONTE)
    vbp = wb.VBProject
    nomes = [c.Name for c in vbp.VBComponents]
    if NOME in nomes:
        comp = vbp.VBComponents(NOME)
        if comp.Type != VBEXT_CT_MSFORM:
            raise SystemExit(f'{NOME} existe e nao e um UserForm (tipo {comp.Type}) -- nada foi feito')
        dif = _diferencas(comp, texto)
        if not dif:
            log(f'   {NOME}: igual ao instalador (0 mudancas)')
            return 0
        log(f'   {NOME}: refeito ({", ".join(dif[:8])}{" ..." if len(dif) > 8 else ""})')
    else:
        comp = vbp.VBComponents.Add(VBEXT_CT_MSFORM)
        comp.Name = NOME
        log(f'   {NOME}: criado')
    # codigo fora antes de mexer nos controles: evento orfao de controle removido nao chega a existir
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    for k, v in FORM.items():
        comp.Properties(k).Value = v
    des = comp.Designer
    for n in [c.Name for c in des.Controls]:
        des.Controls.Remove(n)
    for progid, n, props in CONTROLES:
        ctl = des.Controls.Add(progid, n, True)
        for k, v in props.items():
            _gravar(ctl, k, v)
    cm.AddFromString(texto)
    dif = _diferencas(comp, texto)
    if dif:
        raise SystemExit(f'{NOME} ficou diferente do instalador: {dif} -- nada foi salvo')
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
