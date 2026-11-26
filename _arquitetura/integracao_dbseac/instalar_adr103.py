# -*- coding: utf-8 -*-
"""instalar_adr103.py -- ADR-103 (entrega 49, 26/11/2026): alvos a varios analitos (janela frmAlvosLote).

Uso:  python instalar_adr103.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

A regra institucional do Sigma (ADR-103 §1) e instalada pelo instalar_adr101 (dono das cores/classe) e pelo
codigo_atual (mQualidade). O botao "Aplicar a vários analitos" e da secao do ADR-102 (instalar_adr102.BOTOES: a
secao recusa forma estranha nas linhas 97-183). Aqui: 1. janela frmAlvosLote (declarativa, instalar_frm_inativar);
2. VBA atual (compila a janela junto); 3. fumaca da janela e do botao; 4. protecao e Save. Idempotente; qualquer
falha para ANTES de salvar.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))

MUDANCAS = []
UI = 'Segoe UI'
CINZA = 0x5A5A5A
FONTE_FORM = os.path.join(AQUI, 'src', 'comum', 'frmAlvosLote.frm.txt')
FORM = {'Caption': 'Aplicar alvos a vários analitos', 'Width': 336, 'Height': 444, 'StartUpPosition': 1}
CONTROLES = [
    ('Forms.Label.1', 'lblInfo', {'Left': 12, 'Top': 9, 'Width': 306, 'Height': 30, 'Caption': '', 'Font.Name': UI,
                                  'Font.Size': 9, 'Font.Bold': True, 'WordWrap': True}),
    ('Forms.Label.1', 'lblInstr', {'Left': 12, 'Top': 42, 'Width': 306, 'Height': 30, 'Font.Name': UI, 'Font.Size': 8,
                                   'ForeColor': CINZA, 'WordWrap': True,
                                   'Caption': 'Marque os analitos (caixa em cada linha). O lote e o período são os da '
                                              'seção do Painel (B98, E98 e G98), os mesmos para todos.'}),
    ('Forms.ListBox.1', 'lstAnalitos', {'Left': 12, 'Top': 75, 'Width': 306, 'IntegralHeight': False, 'Height': 240,
                                        'MultiSelect': 1, 'ListStyle': 1, 'Font.Name': UI, 'Font.Size': 9,
                                        'TabIndex': 0}),
    ('Forms.CommandButton.1', 'btnTodos', {'Left': 12, 'Top': 321, 'Width': 120, 'Height': 24,
                                           'Caption': 'Selecionar todos', 'Font.Name': UI, 'Font.Size': 9,
                                           'TabIndex': 1}),
    ('Forms.CommandButton.1', 'btnNenhum', {'Left': 138, 'Top': 321, 'Width': 120, 'Height': 24,
                                            'Caption': 'Desmarcar todos', 'Font.Name': UI, 'Font.Size': 9,
                                            'TabIndex': 2}),
    ('Forms.Label.1', 'lblQtd', {'Left': 12, 'Top': 351, 'Width': 306, 'Height': 15, 'Caption': '', 'Font.Name': UI,
                                 'Font.Size': 8, 'ForeColor': CINZA}),
    ('Forms.CommandButton.1', 'btnAplicar', {'Left': 126, 'Top': 375, 'Width': 96, 'Height': 27, 'Caption': 'Aplicar',
                                             'Default': True, 'Font.Name': UI, 'Font.Size': 9, 'Font.Bold': True,
                                             'TabIndex': 3}),
    ('Forms.CommandButton.1', 'btnCancelar', {'Left': 228, 'Top': 375, 'Width': 90, 'Height': 27,
                                              'Caption': 'Cancelar', 'Cancel': True, 'Font.Name': UI, 'Font.Size': 9,
                                              'TabIndex': 4}),
]


def log(m):
    print(m, flush=True)


def fumaca_lote(ex, wb):
    nome = wb.Name
    cat = str(ex.run(f"'{nome}'!mAlvosRef.CatalogoAnalitos", teto=60)).split('|')
    if len(cat) < 2:
        raise SystemExit(f'fumaca: catalogo com {len(cat)} analito(s) -- nada foi salvo')
    t = str(ex.run(f"'{nome}'!mAlvosRef.TesteJanelaVarios", 'TODOS', teto=60)).split(chr(9))
    if t[0].split('|') != cat or t[2] != '1':
        raise SystemExit(f'fumaca: "Selecionar todos" nao marcou o catalogo inteiro: {t[1]} -- nada foi salvo')
    t = str(ex.run(f"'{nome}'!mAlvosRef.TesteJanelaVarios", f'TODOS|-{cat[0]}', teto=60)).split(chr(9))
    if t[0].split('|') != cat[1:]:
        raise SystemExit(f'fumaca: desmarcar um analito falhou: {t[1]} -- nada foi salvo')
    t = str(ex.run(f"'{nome}'!mAlvosRef.TesteJanelaVarios", f'NENHUM|+{cat[1]}', teto=60)).split(chr(9))
    if t[0] != cat[1] or t[2] != '1':
        raise SystemExit(f'fumaca: marcar um analito falhou: {t} -- nada foi salvo')
    t = str(ex.run(f"'{nome}'!mAlvosRef.TesteJanelaVarios", 'TODOS|NENHUM', teto=60)).split(chr(9))
    if t[0] != '' or t[2] != '0':
        raise SystemExit(f'fumaca: "Desmarcar todos" falhou ou Aplicar habilitado sem selecao: {t} -- nada foi salvo')
    p = wb.Worksheets('Painel')
    b = p.Shapes('adr103_varios')
    a = p.Shapes('adr102_aplicar')
    if not str(b.OnAction).endswith('mAlvosRef.AplicarAlvosVariosUI'):
        raise SystemExit(f'fumaca: botao adr103_varios sem a macro ({b.OnAction}) -- nada foi salvo')
    if abs(float(b.Top) - float(a.Top)) > 0.5 or float(b.Left) < float(a.Left) + float(a.Width):
        raise SystemExit('fumaca: botao adr103_varios fora do lugar (ao lado do Aplicar) -- nada foi salvo')
    log(f'   ok: janela (catalogo {len(cat)} analitos; todos/desmarcar/marcar) e botao ao lado do Aplicar')


def layout_topo(wb, nlv):
    """ADR-103 §3 (pedido do usuario): topo do Painel mais limpo, sem mudar formula nem regra.
    a) C4:D4 (informativo "Analito n de m · Lançamentos no lote") sai: conteudo e mesclagem;
    b) H4:I4 mescladas, centralizadas nos dois eixos (a formula continua em H4);
    c) status por nivel (J6:J9) RECORTADO para a coluna tecnica AB, oculta (mesmas linhas: as referencias relativas, o
       Q4 e as regras de cor vao junto; nome stNivel); a Classe Sigma (K6:K9) RECORTADA para J6:J9 e cada linha mesclada
       J:K (o texto mais longo, "Excelente — Classe mundial", cabe; nenhuma largura muda: o filtro K3:K4 e o encaixe
       dos graficos ficam como estavam -- trocar as larguras cortava o K4, qa_adr090 V02)."""
    import painel_layout as pl
    p = wb.Worksheets('Painel')
    ult = 6 + nlv
    # a) C4:D4
    f = str(p.Range('C4').Formula or '')
    if f.startswith('="Analito "'):
        p.Range('C4:D4').UnMerge()
        p.Range('C4:D4').ClearContents()
        MUDANCAS.append('Painel C4:D4 removida (informativo)')
    elif f not in ('',) or str(p.Range('D4').Formula or '') != '':
        raise SystemExit(f'Painel C4:D4 inesperada: {f[:80]!r} -- nada foi salvo')
    elif p.Range('C4').MergeCells:
        p.Range('C4').MergeArea.UnMerge()
    # b) H4:I4
    if not (p.Range('H4').MergeCells and p.Range('H4').MergeArea.Address == '$H$4:$I$4'):
        if str(p.Range('I4').Formula or '') != '':
            raise SystemExit('Painel I4 ocupada: H4:I4 nao pode ser mesclada -- nada foi salvo')
        p.Range('H4:I4').Merge()
        MUDANCAS.append('Painel H4:I4 mescladas')
    h = p.Range('H4:I4')
    h.HorizontalAlignment = -4108                       # centro
    h.VerticalAlignment = -4108
    h.WrapText = True
    # c) status -> AB (oculta); classe -> J
    C = pl.COL_STATUS_NOVA
    if not pl.layout_novo(p):
        if str(p.Range('K6').Value or '') != pl.CAB_CLASSE:
            raise SystemExit('Painel K6 sem "Classe Sigma" (instalar_adr101 nao rodou antes) -- nada foi salvo')
        if 'SEM MÉD/DP' not in str(p.Range('J7').Formula):
            raise SystemExit('Painel J7 sem o status por nivel (instalar_adr078) -- nada foi salvo')
        alvo = p.Range(f'{C}6:{C}{ult}')
        if any(c.Formula not in (None, '') for c in alvo.Cells):
            raise SystemExit(f'Painel {alvo.Address} ocupada -- nada foi salvo')
        p.Range(f'J6:J{ult}').Cut(p.Range(f'{C}6'))
        p.Range(f'K6:K{ult}').Cut(p.Range('J6'))
        MUDANCAS.append(f'Painel: status por nível J6:J{ult} -> {C} (oculta); Classe Sigma K6:K{ult} -> J6:J{ult}')
    for r in range(6, ult + 1):                         # classe em J:K mesclados (K vazio depois do recorte)
        m = p.Range(f'J{r}:K{r}')
        if not (p.Range(f'J{r}').MergeCells and p.Range(f'J{r}').MergeArea.Address == f'$J${r}:$K${r}'):
            if str(p.Range(f'K{r}').Formula or '') != '':
                raise SystemExit(f'Painel K{r} ocupada: a Classe Sigma nao pode ocupar J{r}:K{r} -- nada foi salvo')
            m.Merge()
            MUDANCAS.append(f'Painel J{r}:K{r} mescladas')
        m.HorizontalAlignment = -4108
    if not p.Columns(C).Hidden:
        p.Columns(C).Hidden = True
        MUDANCAS.append(f'Painel coluna {C} oculta')
    ref = f'=Painel!${C}$7:${C}${ult}'
    try:
        atual = str(wb.Names('stNivel').RefersTo)
    except Exception:                                   # noqa: BLE001
        atual = ''
    if atual.replace(' ', '') != ref:
        if atual:
            wb.Names('stNivel').Delete()
        wb.Names.Add('stNivel', ref)                    # posicional curto: Names.Add(..., True) corrompe (ADR-086/087)
        MUDANCAS.append(f'nome stNivel {ref}')
    q4 = str(p.Range('Q4').Formula)
    if f'${C}$7:${C}${ult}' not in q4 or '#REF' in q4.upper():
        raise SystemExit(f'Painel Q4 nao aponta o status em {C}: {q4[:120]!r} -- nada foi salvo')
    for r in range(7, ult + 1):
        if 'Estatística!$M$' not in str(p.Range(f'J{r}').Formula):
            raise SystemExit(f'Painel J{r} sem a Classe Sigma -- nada foi salvo')
    ruins = [c.Address for c in p.Range('A1:AB12').Cells if '#REF' in str(c.Formula).upper()]
    if ruins:
        raise SystemExit(f'Painel com #REF! depois do layout: {ruins[:6]} -- nada foi salvo')


R_REF_CAB, R_REF1, N_REF, DESLOC = 114, 115, 40, 13      # grade da referencia (ADR-102) e deslocamento da digitacao
NOTA_DIG = ('Digitação de todos os valores de uma vez: preencha a tabela à direita (a partir de N114, mesma estrutura '
            'da grade) e clique em TRANSFERIR VALORES PARA REFERÊNCIA INICIAL. Só os valores preenchidos vão para a '
            'grade do lote em B98 (nunca fórmulas); a tabela de digitação não é apagada.')


def tabela_digitacao(wb, nlv):
    """ADR-103 §4: a partir de N114, a MESMA estrutura da grade da referencia bula/Unity (A114: Analito, Media/LI/LS por
    nivel, Fonte -- sem a coluna calculada "Fora da faixa"), coluna = grade + 13; analito espelhado da coluna A; celulas
    de valor desbloqueadas. Idempotente; os valores ja digitados sao preservados (o instalar_adr102 os guarda na
    reinstalacao da secao)."""
    p = wb.Worksheets('Painel')
    ncol = 1 + nlv * 3 + 1                                # Analito + Media/LI/LS por nivel + Fonte
    c0 = 1 + DESLOC                                       # N
    # cabecalho igual ao da grade (texto e formato)
    p.Range(p.Cells(R_REF_CAB, 1), p.Cells(R_REF_CAB, ncol)).Copy()
    p.Cells(R_REF_CAB, c0).PasteSpecial(-4122)            # xlPasteFormats
    wb.Application.CutCopyMode = False
    for c in range(1, ncol + 1):
        t = p.Cells(R_REF_CAB, c).Value
        if p.Cells(R_REF_CAB, c + DESLOC).Value != t:
            p.Cells(R_REF_CAB, c + DESLOC).Value = t
    # corpo: formato da grade; analito = espelho da A; valores desbloqueados
    valores = p.Range(p.Cells(R_REF1, c0 + 1), p.Cells(R_REF1 + N_REF - 1, c0 + ncol - 1)).Value
    p.Range(p.Cells(R_REF1, 1), p.Cells(R_REF1 + N_REF - 1, ncol)).Copy()
    p.Cells(R_REF1, c0).PasteSpecial(-4122)
    wb.Application.CutCopyMode = False
    p.Range(p.Cells(R_REF1, c0 + 1), p.Cells(R_REF1 + N_REF - 1, c0 + ncol - 1)).Value = valores   # formato nao apaga
    for r in range(R_REF1, R_REF1 + N_REF):
        f = f'=IF($A{r}="","",$A{r})'
        if str(p.Cells(r, c0).Formula) != f:
            p.Cells(r, c0).Formula = f
    corpo = p.Range(p.Cells(R_REF1, c0 + 1), p.Cells(R_REF1 + N_REF - 1, c0 + ncol - 1))
    corpo.Locked = False
    corpo.FormatConditions.Delete()                       # as cores de "fora da faixa" sao da grade, nao da digitacao
    p.Range(p.Cells(R_REF1, c0), p.Cells(R_REF1 + N_REF - 1, c0)).Locked = True
    # nota na linha do botao (a esquerda dele)
    r_bot = R_REF1 + N_REF
    nota = p.Range(f'A{r_bot}:M{r_bot}')
    if not nota.MergeCells:
        nota.Merge()
    if str(p.Range(f'A{r_bot}').Value or '') != NOTA_DIG:
        p.Range(f'A{r_bot}').Value = NOTA_DIG
        MUDANCAS.append(f'Painel tabela de digitação da bula a partir de N{R_REF_CAB} e nota em A{r_bot}')
    nota.WrapText = True
    nota.Font.Size = 8
    nota.Font.Italic = True
    nota.Font.Color = 0x5A5A5A
    nota.VerticalAlignment = -4108


def fumaca_bula(ex, wb, nlv):
    p = wb.Worksheets('Painel')
    ncol = 1 + nlv * 3 + 1
    cab = [p.Cells(R_REF_CAB, c).Value for c in range(1, ncol + 1)]
    dig = [p.Cells(R_REF_CAB, c + DESLOC).Value for c in range(1, ncol + 1)]
    if cab != dig:
        raise SystemExit(f'fumaca: cabecalho da digitacao {dig} != grade {cab} -- nada foi salvo')
    ex.xl.Calculate()
    an = [p.Cells(r, 1).Value for r in range(R_REF1, R_REF1 + N_REF)]
    an_d = [p.Cells(r, 1 + DESLOC).Value for r in range(R_REF1, R_REF1 + N_REF)]
    if [a or '' for a in an] != [a or '' for a in an_d]:
        raise SystemExit('fumaca: analitos da digitacao != grade -- nada foi salvo')
    if p.Cells(R_REF1, 2 + DESLOC).Locked or not p.Cells(R_REF1, 1 + DESLOC).Locked:
        raise SystemExit('fumaca: digitacao com celulas de valor travadas ou analito destravado -- nada foi salvo')
    b = p.Shapes('adr103_transferir')
    if not str(b.OnAction).endswith('mAlvosRef.TransferirReferenciaInicial'):
        raise SystemExit('fumaca: botao TRANSFERIR sem a macro -- nada foi salvo')
    log(f'   ok: tabela de digitação N{R_REF_CAB}:{p.Cells(R_REF1 + N_REF - 1, ncol + DESLOC).Address.replace("$", "")} '
        f'e botão TRANSFERIR')


def main(produto, caminho, salvar=True):
    import xlh
    import instalar_integracao as ii
    import codigo_atual
    import instalar_frm_inativar
    from instalar_seguranca_usuarios import foto_protecao, devolver_protecao
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.EnableEvents = False
        log('1. janela frmAlvosLote')
        n = instalar_frm_inativar.criar_form(wb, log, nome='frmAlvosLote', fonte=FONTE_FORM, form=FORM,
                                             controles=CONTROLES)
        if n:
            MUDANCAS.append('frmAlvosLote')
        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py; compila a janela junto)')
        codigo_atual.instalar(ex, wb, produto, log)
        log('3. topo do Painel (C4:D4, H4:I4, status -> AB oculta, Classe Sigma -> J)')
        layout_topo(wb, 2 if produto.startswith('Bio') else 3)
        log('3b. tabela de digitação da bula (N114) e nota do botão TRANSFERIR')
        tabela_digitacao(wb, 2 if produto.startswith('Bio') else 3)
        # as larguras de J/K mudaram (o K98 do "Aplicar" anda): botoes da secao de volta a sua celula, COM a aba ativa
        # (armadilha do gate da entrega 48: posicao gravada com a aba inativa e recalculada ao ativar)
        import instalar_adr102
        instalar_adr102.alinhar_com_aba_ativa(wb)
        log('4. fumaca')
        try:
            wb.Worksheets('Painel').Shapes('adr103_varios')
            wb.Worksheets('Painel').Shapes('adr103_transferir')
        except Exception:                               # noqa: BLE001
            raise SystemExit('Painel sem os botoes do ADR-103: rodar o instalar_adr102 antes -- nada foi salvo')
        ex.xl.Calculation = -4105
        fumaca_lote(ex, wb)
        fumaca_bula(ex, wb, 2 if produto.startswith('Bio') else 3)
        ex.xl.EnableEvents = False
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
