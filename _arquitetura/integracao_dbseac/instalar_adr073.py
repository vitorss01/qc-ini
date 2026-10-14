# -*- coding: utf-8 -*-
"""instalar_adr073.py -- ADR-073: aba Analitos como CENTRAL DE LIMITES POR LOTE, com rastreabilidade.

Uso:  python instalar_adr073.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

O defeito (relato do usuario, 14/10/2026): no Painel so o primeiro lote mostrava media e limites
(-3DP..+3DP) e formatacao; os outros nao. Causa: so o 1o lote tinha Media/DP na LotesStore (os outros
8 da Hematologia e ~20 da Bioquimica nunca foram cadastrados -- lote novo nasce vazio, de proposito,
para nao herdar limites de outro material) e editar outro lote exigia trocar o lote no Painel antes.

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py: mLotes com seletor/retrato/preenchimento,
     modulo da aba Analitos, LotesStore visivel a todos); EXIGE compilacao;
  2. Analitos, linha 1: titulo encurtado, "LOTE EM EDICAO" + seletor (nome loteEdicao, mesma lista do
     Painel H3) e o resumo do registro do lote (nome loteResumo = mLotes.RegistroLote(loteParam));
     botao "Preencher vazios pelo CIQ" (mLotes.PreencherVaziosPeloCIQ); a nota da linha 2 aponta o seletor;
  3. LotesStore: colunas Q..W (fonte da especificacao, ano da referencia, ETp, CVTp, registrado em, por,
     origem), cabecalho, larguras, linha 1 congelada; linhas JA com Media/DP e sem retrato recebem o da
     especificacao vigente, com a origem dizendo que o retrato foi tirado na instalacao;
  4. FUMACA: nomes, seletor = lote em tela, resumo legivel, colunas, retrato nas linhas com valor;
  5. devolve a protecao de cada aba e a da estrutura como estavam.
Nao cadastra nem calcula Media/DP de lote nenhum: isso e decisao do laboratorio (botao ou digitacao).
"""
import datetime
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

CAB_NOVOS = ['ESPEC_FONTE', 'ESPEC_ANO', 'ETP_%', 'CVTP_%', 'REGISTRADO_EM', 'REGISTRADO_POR', 'ORIGEM']
C_Q = 17
BTN = 'btnPreencherCIQ'


def log(*a):
    print(*a, flush=True)


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def achar(ws, linha, prefixos, ate=60):
    for c in range(1, ate):
        s = ' '.join(str(ws.Cells(linha, c).Value or '').upper().split())
        if any(s.startswith(p) for p in prefixos):
            return c
    return 0


def nome(wb, nm, ref):
    alvo = ref.replace("'", '').replace('$', '').upper()
    for n in list(wb.Names):
        if n.Name == nm:
            if str(n.RefersTo).replace("'", '').replace('$', '').upper() == alvo:
                return                                         # ja aponta para o lugar certo
            n.Delete()
            break
    wb.Names.Add(nm, ref)


def cabecalho_analitos(wb):
    a = wb.Worksheets(ii.aba_espec(wb))
    p = wb.Worksheets('Painel')
    fim = a.Range('A1').MergeArea.Column + a.Range('A1').MergeArea.Columns.Count - 1
    if str(a.Range('I1').Value or '').upper().startswith('LOTE EM EDI'):
        fim = max(fim, a.Range('K1').MergeArea.Column + a.Range('K1').MergeArea.Columns.Count - 1)
    fim = max(fim, 18)
    a.Range(a.Cells(1, 1), a.Cells(1, fim)).UnMerge()
    a.Range('A1:H1').Merge()
    a.Range('I1').Value = 'LOTE EM EDIÇÃO ▸'
    a.Range('I1').Font.Bold = True
    a.Range('I1').HorizontalAlignment = -4152                 # direita
    sel = a.Range('J1')
    sel.NumberFormat = '@'
    sel.Locked = False
    sel.Font.Bold = True
    sel.Interior.Color = 0xCCFFFF                              # amarelo claro (BGR): "entrada"
    sel.HorizontalAlignment = -4108
    v = sel.Validation
    try:
        f1 = p.Range('H3').Validation.Formula1
    except Exception:
        f1 = '=lstLotes'
    v.Delete()
    v.Add(3, 1, 1, f1)
    v.InputTitle = 'Lote em edição'
    v.InputMessage = 'Escolha o lote: as colunas Média/DP (E:J) passam a ser as dele. É o mesmo lote do Painel (H3).'
    nome(wb, 'loteEdicao', f"='{a.Name}'!$J$1")
    sel.Value = str(wb.Names('loteParam').RefersToRange.Value or '')
    res = a.Range(a.Cells(1, 11), a.Cells(1, fim))
    res.Merge()
    res.Formula = '=mLotes.RegistroLote(loteParam)'
    res.Font.Size = 9
    res.Font.Italic = True
    res.WrapText = True
    res.HorizontalAlignment = -4131
    res.VerticalAlignment = -4108
    nome(wb, 'loteResumo', f"='{a.Name}'!$K$1:${col(fim)}$1")
    if a.Rows(1).RowHeight < 30:
        a.Rows(1).RowHeight = 30
    # nota da linha 2 aponta o seletor
    f = str(a.Range('A2').Formula)
    novo = f.replace('selecione-o no Painel (H3)', 'escolha-o em J1 (LOTE EM EDIÇÃO) ou no Painel (H3)')
    if novo != f:
        a.Range('A2').Formula = novo
    # botao
    for s in list(a.Shapes):
        if s.Name == BTN:
            s.Delete()
    c0 = a.Cells(1, fim + 1)
    b = a.Shapes.AddShape(5, c0.Left + 4, c0.Top + 2, 210, a.Rows(1).RowHeight - 4)   # retangulo arredondado
    b.Name = BTN
    b.TextFrame2.TextRange.Text = 'Preencher Média/DP vazios pelo CIQ (20 primeiros)'
    b.TextFrame2.TextRange.Font.Size = 9
    b.TextFrame2.TextRange.Font.Bold = True
    b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    b.Fill.ForeColor.RGB = 0x7A5A1F                            # azul escuro (BGR)
    b.Line.Visible = False
    b.Placement = 2                                            # move, nao redimensiona
    b.OnAction = 'mLotes.PreencherVaziosPeloCIQ'
    return fim


def lotesstore(wb):
    ws = wb.Worksheets('LotesStore')
    ws.Range(ws.Cells(1, 16), ws.Cells(1, 16)).Copy()
    ws.Range(ws.Cells(1, C_Q), ws.Cells(1, C_Q + len(CAB_NOVOS) - 1)).PasteSpecial(-4122)
    for j, h in enumerate(CAB_NOVOS):
        ws.Cells(1, C_Q + j).Value = h
        ws.Columns(C_Q + j).ColumnWidth = 22 if h in ('ORIGEM', 'REGISTRADO_EM') else 13
    ws.Columns(C_Q + 4).NumberFormat = 'dd/mm/yyyy hh:mm'
    # retrato nas linhas que ja tem Media/DP e ainda nao tem (anteriores ao ADR-073)
    a = wb.Worksheets(ii.aba_espec(wb))
    cF = achar(a, 3, ['ESPQ FONTE', 'ETP FONTE'])
    cE = achar(a, 3, ['ETP EM USO', 'ETP FINAL'])
    cC = achar(a, 3, ['CVTP%', 'CVTP %'])
    cCl, cVb = achar(a, 3, ['LIM CV CLIA']), achar(a, 3, ['LIM CV VB'])
    ano = '-'
    for j in range(1, 40):
        if 'ESPECIFICA' in str(a.Cells(2, j).Value or '').upper() and 'DO ANO' in str(a.Cells(2, j).Value or '').upper():
            for k in range(j + 1, j + 4):
                if str(a.Cells(2, k).Value or '').strip():
                    ano = a.Cells(2, k).Value
                    break
    linhas_an = {str(a.Cells(r, 1).Value or '').strip().upper(): r for r in range(4, 44) if a.Cells(r, 1).Value}
    ult = ws.Cells(ws.Rows.Count, 16).End(-4162).Row
    hoje = datetime.date.today().strftime('%d/%m/%Y')
    feitas = 0
    for r in range(2, ult + 1):
        vals = ws.Range(ws.Cells(r, 3), ws.Cells(r, 8)).Value[0]
        if not any(isinstance(x, (int, float)) for x in vals):
            continue
        if str(ws.Cells(r, C_Q).Value or '').strip():
            continue
        ra = linhas_an.get(str(ws.Cells(r, 15).Value or '').strip().upper())
        if not ra:
            continue
        fonte = str(a.Cells(ra, cF).Value or '-') if cF else '-'
        etp = a.Cells(ra, cE).Value if cE else '-'
        if cC:
            cv = a.Cells(ra, cC).Value
        elif fonte.upper() == 'CLIA' and cCl:
            cv = a.Cells(ra, cCl).Value
        elif fonte.upper() == 'VB' and cVb:
            cv = a.Cells(ra, cVb).Value
        else:
            cv = '-'
        num = lambda x: x if isinstance(x, (int, float)) and x > 0 else '-'
        ws.Range(ws.Cells(r, C_Q), ws.Cells(r, C_Q + 6)).Value = (
            (fonte or '-', ano, num(etp), num(cv), None, '-',
             f'registrado antes do ADR-073 (retrato da especificacao tirado na instalacao, {hoje})'),)
        feitas += 1
    vis = ws.Visible
    try:
        ws.Visible = -1                                        # congelar exige a aba visivel
        ws.Activate()
        wb.Application.ActiveWindow.FreezePanes = False
        ws.Range('A2').Select()
        wb.Application.ActiveWindow.FreezePanes = True
        ws.Range('A1').Select()
    finally:
        ws.Visible = vis
    return feitas


def fumaca(wb, fim):
    ruins = []
    a = wb.Worksheets(ii.aba_espec(wb))
    for n in ('loteEdicao', 'loteResumo', 'loteParam'):
        try:
            wb.Names(n).RefersToRange
        except Exception:
            ruins.append(('nome ausente', n))
    if str(a.Range('J1').Value or '') != str(wb.Names('loteParam').RefersToRange.Value or ''):
        ruins.append(('seletor != loteParam', a.Range('J1').Value, wb.Names('loteParam').RefersToRange.Value))
    res = str(a.Range('K1').Text)
    if not res.startswith('Lote ') and not res.startswith('Escolha'):
        ruins.append(('resumo', res[:120]))
    if a.Range('J1').Locked:
        ruins.append(('seletor bloqueado',))
    if not any(s.Name == BTN for s in a.Shapes):
        ruins.append(('botao ausente',))
    ws = wb.Worksheets('LotesStore')
    if [ws.Cells(1, C_Q + j).Value for j in range(len(CAB_NOVOS))] != CAB_NOVOS:
        ruins.append(('cabecalho LotesStore',))
    ult = ws.Cells(ws.Rows.Count, 16).End(-4162).Row
    for r in range(2, ult + 1):
        vals = ws.Range(ws.Cells(r, 3), ws.Cells(r, 8)).Value[0]
        if any(isinstance(x, (int, float)) for x in vals) and not str(ws.Cells(r, C_Q).Value or '').strip():
            ruins.append(('linha com valor sem retrato', r, ws.Cells(r, 16).Value))
            break
    if ruins:
        raise SystemExit(f'fumaca ADR-073 falhou: {ruins[:6]}')
    return res


def main(produto, caminho, salvar=True):
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
        ex.xl.Calculation = -4135

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        fim = cabecalho_analitos(wb)
        log(f'2. Analitos: seletor J1 (loteEdicao), resumo K1:{col(fim)}1 (loteResumo), botao {BTN}')
        n = lotesstore(wb)
        log(f'3. LotesStore: colunas Q..W; retrato tirado agora em {n} linha(s) que ja tinham Media/DP')

        ex.xl.Calculation = -4105
        ex.xl.Calculate()
        ex.esperar()
        res = fumaca(wb, fim)
        log(f'4. fumaca OK: {res[:140]}')
        try:
            wb.Worksheets(ii.aba_espec(wb)).Activate()
            wb.Worksheets(ii.aba_espec(wb)).Range('A1').Select()
        except Exception:
            pass

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
