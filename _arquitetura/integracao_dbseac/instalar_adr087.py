# -*- coding: utf-8 -*-
"""instalar_adr087.py -- ADR-087 (29/10/2026): modernizacao das telas -- parte A: aba Registros como Tabela do Excel.

Uso:  python instalar_adr087.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. todo o VBA atual das fontes (codigo_atual.py: mCalibracao com a tabela que cresce e os filtros, Registros.cls,
     mLotes sem a troca da aba Registros por lote, mSeguranca redesenhando as assinaturas pelas linhas da tabela);
  2. Registros -> Tabela do Excel tblRegistros (A..M, coluna M "Lote"), SO na 1a vez: os registros da vista (lote
     carregado) e os de cada lote da RegistrosStore viram linhas da tabela, com o lote; CONFERE que cada registro
     (todas as celulas) chegou igual e que nenhum se perdeu -- senao PARA sem salvar. A RegistrosStore fica intacta
     como arquivo historico (nota em A1);
  3. linhas 3:4 = filtros da vista (Lote, Analito, De, Ate, Equipamento, Calibrado?, Status) + botao Limpar filtros
     + a linha de status; listas tiradas dos dados (colunas AA:AD ocultas); validacoes;
  4. nomes regData/regAnalito/regNivel/regCalib/regEquip/regID/regView/regLote -> colunas da tabela (crescem junto);
  5. Calc: o icone de calibracao conta so as linhas do LOTE EM USO (regLote = loteAtivo) -- exatamente o que a vista
     por lote mostrava antes (a regra do grafico nao muda);
  6. devolve a protecao de cada aba e a da estrutura como estavam. Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import visual_telas  # noqa: E402
import codigo_atual  # noqa: E402
import instalar_adr078 as a78  # noqa: E402
import instalar_frm_inativar  # noqa: E402
import ceq_form_spec  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
TB = 'tblRegistros'
CAB = ['Nº', 'Data', 'Analito', 'Nível', 'Corrida (Seq)', 'Foi calibrado?', 'Observação', 'Responsável',
       'Assinatura (duplo clique)', 'Status', 'ID do resultado', 'Equipamento', 'Lote']
NLB = 200                                   # linhas por lote na RegistrosStore (mLotes.NLB_LIBER)
H0 = 6                                      # linha do cabecalho da tabela (1 titulo, 2 instrucao, 3:4 filtros, 5 folga)
FILTROS = [(2, 'Lote', '=lstRegLote'), (3, 'Analito', '=lstRegAnalito'), (5, 'De (data)', None), (6, 'Até (data)', None),
           (7, 'Equipamento', '=lstRegEquip'), (8, 'Calibrado?', 'Sim,Não'), (10, 'Status', '=lstRegStatus')]
A2 = ('Uma linha por registro (todos os lotes). Calibração: digite na coluna "ID do resultado" o ID do resultado medido '
      'logo depois da calibração -- data, analito, RUN, equipamento e valor são preenchidos. Assinar: duplo clique em '
      '"Assinatura". A próxima linha aparece sozinha. Os filtros abaixo só mudam a vista.')
AZUL, CINZA_CLARO, CINZA = 0x79481F, 0xF7F3F0, 0x666666        # BGR


def log(m):
    print(m, flush=True)


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return
            n.Delete()
            break
    wb.Names.Add(nm, ref)
    MUDANCAS.append(f'nome {nm}')


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def vazia(linha):
    return not any(txt(x) for x in linha)


def coletar(wb):
    """Registros a migrar: (lote, [11 valores B..L]) -- a vista do lote carregado e os blocos da store dos outros."""
    ws, st = wb.Worksheets('Registros'), wb.Worksheets('RegistrosStore')
    lotes = [txt(r[0]) for r in wb.Names('regLoteCol').RefersToRange.Value]
    carregado = txt(wb.Names('loteCarregado').RefersToRange.Value)
    vista = ws.Range('B4:L203').Value
    out = []
    for i, lote in enumerate(lotes):
        if not lote:
            continue
        if lote == carregado:
            linhas = vista
        else:
            r0 = 2 + i * NLB
            linhas = st.Range(st.Cells(r0, 1), st.Cells(r0 + NLB - 1, 11)).Value
        for ln in linhas:
            if not vazia(ln):
                out.append((lote, list(ln)))
    if carregado and carregado not in lotes:
        for ln in vista:
            if not vazia(ln):
                out.append((carregado, list(ln)))
    return out, carregado


def converter(wb, nlv):
    ws = wb.Worksheets('Registros')
    if TB in [lo.Name for lo in ws.ListObjects]:
        return False
    cab3 = [txt(v) for v in ws.Range('A3:L3').Value[0]]
    if (cab3[0], cab3[1], cab3[2], cab3[5], cab3[9], cab3[10], cab3[11]) != \
            ('Nº', 'Data', 'Analito', 'Foi calibrado?', 'Status', 'ID do resultado', 'Equipamento'):
        raise SystemExit(f'Registros!A3:L3 inesperado: {cab3} -- nada foi feito')
    regs, carregado = coletar(wb)
    fmt_data = ws.Range('B4').NumberFormat
    valid = {}
    for col in ('D', 'F', 'J'):
        try:
            valid[col] = str(ws.Range(f'{col}4').Validation.Formula1)
        except Exception:                              # noqa: BLE001
            valid[col] = None
    # assinaturas desenhadas da vista (refeitas pelo VBA depois)
    for sh in list(ws.Shapes):
        if str(sh.Name).startswith('qcAssin_'):
            sh.Delete()
    ws.Range('A3:M203').Validation.Delete()
    ws.Range('A3:M203').Clear()
    ws.Rows('3:5').Insert()                            # 3:4 filtros, 5 folga -> cabecalho na 6
    n = max(1, len(regs)) + 1                          # + a linha vazia para o proximo registro
    ws.Range(ws.Cells(H0, 1), ws.Cells(H0, 13)).Value = tuple(CAB)
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(H0, 1), ws.Cells(H0 + n, 13)), None, 1)   # xlSrcRange, xlYes
    lo.Name = TB
    lo.TableStyle = 'TableStyleLight9'
    lo.ShowAutoFilterDropDown = False
    if regs:
        dados = [tuple(v) + (lote,) for lote, v in regs]
        ws.Range(ws.Cells(H0 + 1, 13), ws.Cells(H0 + len(regs), 13)).NumberFormat = '@'
        ws.Range(ws.Cells(H0 + 1, 11), ws.Cells(H0 + len(regs), 11)).NumberFormat = '@'
        ws.Range(ws.Cells(H0 + 1, 2), ws.Cells(H0 + len(regs), 13)).Value = tuple(dados)
    lo.ListColumns(2).DataBodyRange.NumberFormat = fmt_data if fmt_data and fmt_data != 'General' else 'dd/mm/yyyy hh:mm'
    lo.ListColumns(11).DataBodyRange.NumberFormat = '@'
    lo.ListColumns(13).DataBodyRange.NumberFormat = '@'
    lo.ListColumns(1).DataBodyRange.Formula = f'=ROW()-ROW({TB}[[#Headers],[Nº]])'
    for col, idx in (('D', 4), ('F', 6), ('J', 10)):
        if valid.get(col):
            v = lo.ListColumns(idx).DataBodyRange.Validation
            v.Delete()
            v.Add(3, 1, 1, valid[col])                 # xlValidateList
            v.IgnoreBlank = True
    for k, w in enumerate([5, 15, 18, 8, 12, 12, 32, 30, 16, 22, 18, 16, 10], start=1):
        ws.Columns(k).ColumnWidth = w
    lo.Range.WrapText = False
    lo.ListColumns(7).DataBodyRange.WrapText = True    # observacao longa quebra a linha (nao corta)
    lo.ListColumns(8).DataBodyRange.WrapText = True    # identidade da assinatura
    lo.HeaderRowRange.WrapText = True
    lo.DataBodyRange.Locked = False
    lo.ListColumns(1).DataBodyRange.Locked = True      # Nº calculado
    # conferencia: todo registro chegou igual
    back = ws.Range(ws.Cells(H0 + 1, 2), ws.Cells(H0 + max(1, len(regs)), 13)).Value
    ruins = []
    for k, (lote, v) in enumerate(regs):
        got = back[k]
        if txt(got[11]) != lote or [txt(x) for x in got[:11]] != [txt(x) for x in v]:
            ruins.append((k, lote, v[:4], got[:4]))
    if ruins:
        raise SystemExit(f'migracao da aba Registros: {len(ruins)} registro(s) diferentes -- nada foi salvo: {ruins[:3]}')
    st = wb.Worksheets('RegistrosStore')
    st.Range('M1').Value = ('ADR-087: congelada -- os registros foram migrados para a tabela tblRegistros (aba Registros). '
                            'Arquivo histórico; não é mais gravada.')
    MUDANCAS.append(f'Registros -> {TB} ({len(regs)} registro(s) de {len({l for l, _ in regs})} lote(s); '
                    f'lote carregado {carregado})')
    return True


def filtros(wb):
    ws = wb.Worksheets('Registros')
    lo = ws.ListObjects(TB)
    if str(ws.Range('A3').Value or '') != 'FILTROS':
        ws.Range('A3').Value = 'FILTROS'
        MUDANCAS.append('Registros filtros (rotulos)')
    for col, rot, lst in FILTROS:
        c_rot, c_val = ws.Cells(3, col), ws.Cells(4, col)
        if str(c_rot.Value or '') != rot:
            c_rot.Value = rot
            MUDANCAS.append(f'Registros filtro {rot}')
        v = c_val.Validation
        try:
            f1 = str(v.Formula1)
        except Exception:                              # noqa: BLE001
            f1 = ''
        if lst and f1 != lst:
            v.Delete()
            v.Add(3, 1, 1, lst)                        # xlValidateList
            v.IgnoreBlank = True
            v.ShowError = True
            MUDANCAS.append(f'Registros validacao {rot}')
        elif not lst and col in (5, 6) and f1 == '':
            v.Delete()
            v.Add(4, 1, 5, '1/1/2000')                 # xlValidateDate, xlGreater... (so datas)
            v.ErrorMessage = 'Digite uma data (dd/mm/aaaa).'
            c_val.NumberFormat = 'dd/mm/yyyy'
            MUDANCAS.append(f'Registros validacao {rot}')
        c_val.Locked = False
    # aparencia: rotulos pequenos cinza, campos com borda e fundo claro
    rot = ws.Range('A3:L3')
    rot.Font.Size = 8
    rot.Font.Color = CINZA
    rot.Font.Bold = True
    ws.Range('A3').Font.Color = AZUL
    for col, _, _ in FILTROS:
        c = ws.Cells(4, col)
        c.Interior.Color = CINZA_CLARO
        c.Borders.LineStyle = 1
        c.Borders.Color = 0xBFBFBF
        c.Font.Size = 10
    info = ws.Cells(4, 11)
    info.Font.Size = 9
    info.Font.Italic = True
    ws.Range('K4:M4').HorizontalAlignment = 1
    # botao Limpar filtros (na linha 3, colunas K:L) -- tamanho pelo rotulo
    nomes = [str(s.Name) for s in ws.Shapes]
    if 'btnLimparFiltrosReg' not in nomes:
        c = ws.Range('K3')
        b = ws.Shapes.AddShape(5, c.Left + 2, c.Top + 1, 120, max(16, c.Height - 2))
        b.Name = 'btnLimparFiltrosReg'
        b.TextFrame2.TextRange.Text = 'Limpar filtros'
        b.TextFrame2.TextRange.Font.Size = 9
        b.TextFrame2.TextRange.Font.Bold = True
        b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = AZUL
        b.TextFrame2.MarginTop = 0
        b.TextFrame2.MarginBottom = 0
        b.TextFrame2.WordWrap = 0
        b.TextFrame2.AutoSize = 1                      # o botao cresce com o rotulo (nada cortado)
        b.Fill.ForeColor.RGB = 0xFFFFFF
        b.Line.ForeColor.RGB = AZUL
        b.OnAction = 'mCalibracao.LimparFiltrosRegistros'
        b.Placement = 2                                # xlMove
        MUDANCAS.append('Registros botao Limpar filtros')
    if ws.Rows(3).RowHeight < 18:
        ws.Rows(3).RowHeight = 18
    if ws.Rows(4).RowHeight < 18:
        ws.Rows(4).RowHeight = 18
    ws.Rows(5).RowHeight = 6
    # area tecnica das listas (AA:AD) oculta
    for k, t in zip(range(27, 31), ('lista Lote', 'lista Analito', 'lista Equipamento', 'lista Status')):
        if str(ws.Cells(1, k).Value or '') != t:
            ws.Cells(1, k).Value = t
    if not ws.Range('AA:AD').EntireColumn.Hidden:
        ws.Range('AA:AD').EntireColumn.Hidden = True
        MUDANCAS.append('Registros listas ocultas')
    # titulo e instrucao
    if str(ws.Range('A1').Value or '') != 'REGISTROS · OBSERVAÇÕES · CALIBRAÇÃO':
        ws.Range('A1').Value = 'REGISTROS · OBSERVAÇÕES · CALIBRAÇÃO'
        MUDANCAS.append('Registros A1')
    if str(ws.Range('A2').Formula or '') != A2:
        ws.Range('A2').Value = A2
        MUDANCAS.append('Registros A2')
    ws.Range('A2').WrapText = True
    ws.Range('A2').Font.Size = 9
    ws.Range('A2').Font.Color = CINZA
    if ws.Rows(2).RowHeight < 26:
        ws.Rows(2).RowHeight = 26
    # cabecalho, filtros e titulo sempre a vista. No arquivo blindado a aba fica oculta (so a Login aparece): mostra
    # so durante o congelamento (SplitRow, sem Select) e devolve como estava
    vis, ativa = ws.Visible, wb.ActiveSheet
    ws.Visible = -1
    ws.Activate()
    win = wb.Windows(1)
    win.FreezePanes = False
    win.ScrollRow, win.ScrollColumn = 1, 1
    win.SplitColumn, win.SplitRow = 0, H0
    win.FreezePanes = True
    if vis != -1:
        ativa.Activate()
        ws.Visible = vis
    return lo


def nomes_reg(wb):
    for nm, col in (('regData', 'Data'), ('regAnalito', 'Analito'), ('regNivel', 'Nível'), ('regCalib', 'Foi calibrado?'),
                    ('regEquip', 'Equipamento'), ('regID', 'ID do resultado'), ('regLote', 'Lote')):
        nome(wb, nm, f'={TB}[{col}]')
    nome(wb, 'regView', f'={TB}[[Data]:[Equipamento]]')


RX_CAL = re.compile(r'regEquip,TRIM\(""&engEquip\)\)')


def calc(wb):
    c = wb.Worksheets('Calc')
    linha3 = c.Range('A3:EZ3').Formula[0]
    alvo = [k + 1 for k, f in enumerate(linha3) if 'regCalib' in str(f or '')]
    if not alvo:
        raise SystemExit('Calc: nenhuma coluna com regCalib (icone de calibracao) -- nada foi feito')
    feitas = []
    for k in alvo:
        f3 = str(c.Cells(3, k).Formula)
        if 'regLote' in f3:
            continue
        if f3.count('regEquip,') != 2:
            raise SystemExit(f'Calc {a78.col(k)}3: formula de calibracao inesperada (ADR-084?) -- nada foi feito: {f3[:160]}')
        novo = f3.replace('regEquip,"")', 'regEquip,"",regLote,loteAtivo)')
        novo = RX_CAL.sub('regEquip,TRIM(""&engEquip),regLote,loteAtivo)', novo)
        if novo.count('regLote,loteAtivo') != 2:
            raise SystemExit(f'Calc {a78.col(k)}3: nao foi possivel acrescentar o lote: {f3[:160]}')
        c.Range(c.Cells(3, k), c.Cells(182, k)).Formula = novo
        feitas.append(a78.col(k))
    if feitas:
        MUDANCAS.append('Calc calibracao so do lote em uso ' + ','.join(feitas))


def botao(ws, nome_b, texto, celula, macro, largura=None):
    """Botao padrao (retangulo arredondado, contorno azul, texto azul em negrito). O tamanho acompanha o rotulo
    (AutoSize): nada cortado. Ja existe = so confere o texto e a macro."""
    nomes = [str(s.Name) for s in ws.Shapes]
    c = ws.Range(celula)
    if nome_b in nomes:
        b = ws.Shapes(nome_b)
        if str(b.TextFrame2.TextRange.Text) == texto and str(b.OnAction).endswith(macro):
            return
        b.TextFrame2.TextRange.Text = texto
        b.OnAction = macro
    else:
        b = ws.Shapes.AddShape(5, c.Left + 2, c.Top + 2, largura or 180, 24)
        b.Name = nome_b
        b.TextFrame2.TextRange.Text = texto
        b.OnAction = macro
    tr = b.TextFrame2.TextRange
    tr.Font.Size = 10
    tr.Font.Bold = True
    tr.Font.Name = 'Segoe UI'
    tr.Font.Fill.ForeColor.RGB = AZUL
    b.TextFrame2.WordWrap = 0
    b.TextFrame2.AutoSize = 1
    b.TextFrame2.MarginLeft = 8
    b.TextFrame2.MarginRight = 8
    b.Fill.ForeColor.RGB = 0xFFFFFF
    b.Line.ForeColor.RGB = AZUL
    b.Line.Weight = 1.25
    b.Placement = 2
    MUDANCAS.append(f'{ws.Name} botao {nome_b}')


def ceq(wb):
    n = instalar_frm_inativar.criar_form(wb, log, nome='frmCEQ', fonte=ceq_form_spec.FONTE_CODIGO,
                                         form=ceq_form_spec.FORM, controles=ceq_form_spec.CONTROLES)
    if n:
        MUDANCAS.append('frmCEQ')
    for aba in ('EQA.CAP_Resumo', 'EQA.Controllab_Resumo'):
        ws = wb.Worksheets(aba)
        if ws.Rows(3).RowHeight < 28:
            ws.Rows(3).RowHeight = 28
        botao(ws, 'btnRegistrarCEQ', 'Registrar resultado do CEQ', 'A3', 'mCEQForm.AbrirCEQ')


def main(produto, caminho, salvar=True):
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
        nlv = 2 if produto.startswith('Bio') else 3
        log('0. janela frmCEQ (registro do Controle Externo)')
        ceq_n = instalar_frm_inativar.criar_form(wb, log, nome='frmCEQ', fonte=ceq_form_spec.FONTE_CODIGO,
                                                 form=ceq_form_spec.FORM, controles=ceq_form_spec.CONTROLES)
        if ceq_n:
            MUDANCAS.append('frmCEQ')
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        ex.xl.EnableEvents = False
        conv = converter(wb, nlv)
        log(f'2. Registros: tabela {TB} ' + ('criada (registros migrados e conferidos)' if conv else 'ja existia'))
        wb.Application.Run(f"'{wb.Name}'!mCalibracao.AtualizarListasFiltro")      # os nomes lstReg* antes das validacoes
        filtros(wb)
        log('3. filtros da vista (linhas 3:4), botao Limpar filtros, listas ocultas')
        nomes_reg(wb)
        log('4. nomes reg* -> colunas da tabela')
        calc(wb)
        log('5. Calc: icone de calibracao so do lote em uso (regLote)')
        for aba in ('EQA.CAP_Resumo', 'EQA.Controllab_Resumo'):
            ws = wb.Worksheets(aba)
            if ws.Rows(3).RowHeight < 28:
                ws.Rows(3).RowHeight = 28
            botao(ws, 'btnRegistrarCEQ', 'Registrar resultado do CEQ', 'A3', 'mCEQForm.AbrirCEQ')
        log('6. CEQ: botao Registrar resultado nos resumos do CAP e do Controllab')
        MUDANCAS.extend(visual_telas.aplicar(wb, log))
        log('7. telas: nada cortado nem sobreposto (guiado pela auditoria visual)')
        # eventos continuam desligados ate o fim: com eles ligados o Workbook_BeforeSave cancela o Save do Excel e grava
        # pelo proprio caminho (TrancarParaGravar), engolindo o erro -- o log dizia "salvo" e o arquivo nao mudava
        ex.xl.Calculation = -4105
        wb.Application.Run(f"'{wb.Name}'!mCalibracao.AtualizarListasFiltro")
        wb.Application.Run(f"'{wb.Name}'!mCalibracao.GarantirLinhaNova")
        # aplica a vista e redesenha as assinaturas das linhas visiveis: a conversao apagou as imagens (qcAssin_) e
        # nada mais as refaria ate a proxima assinatura ou troca de lote
        r = str(wb.Application.Run(f"'{wb.Name}'!mCalibracao.AplicarFiltrosRegistros"))
        if r.startswith('ERRO'):
            raise SystemExit(f'Registros: filtros/assinaturas: {r} -- nada foi salvo')
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
