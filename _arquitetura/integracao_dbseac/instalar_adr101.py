# -*- coding: utf-8 -*-
"""instalar_adr101.py -- ADR-101 (23/11/2026): cor e classe do Sigma (6 faixas), Sigma negativo = 0, legenda no Painel.

Uso:  python instalar_adr101.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Decisoes do usuario (23/11/2026):
  - o Sigma colorido e o do Painel (I7:I9), o MESMO da Estatistica (L): o Painel le da Estatistica;
  - 6 faixas: >=6 Classe mundial (verde-escuro) | 5-6 Excelente (verde) | 4-5 Bom (verde-claro) | 3-4 Marginal
    (amarelo) | 2-3 Ruim (vermelho-claro) | <2 Inaceitavel (vermelho-escuro);
  - pinta o numero E o texto da classe (Estatistica L e M; Painel I e K "Classe Sigma");
  - Sigma negativo (|bias| > ETp) = 0, vermelho-escuro (Estatistica L = MAX(0; ...));
  - legenda no Painel (V5:X8, acima dos graficos) com a regra de CQ sugerida de cada faixa (da tblPlanoQC_Sigma: informativa; as regras
    que rodam nao mudam).
Idempotente. Passos: 1. VBA atual; 2. tblPlanoQC_Sigma: a faixa "<3 Desempenho inadequado" vira "2-3 Ruim" + "<2
Inaceitavel" (mesmas regras); 3. Estatistica L com MAX(0;...); 4. Painel K7:K9 = classe; 5. formatacao condicional;
6. legenda; 7. fumaca; 8. protecao e Save. Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []


def rgb(r, g, b):
    return r + g * 256 + b * 65536


# (minimo, maximo, classe, fundo, fonte)
FAIXAS = [(6.0, None, 'Classe mundial', rgb(0, 97, 0), rgb(255, 255, 255)),
          (5.0, 6.0, 'Excelente', rgb(0, 176, 80), rgb(255, 255, 255)),
          (4.0, 5.0, 'Bom', rgb(198, 239, 206), rgb(0, 97, 0)),
          (3.0, 4.0, 'Marginal', rgb(255, 235, 156), rgb(156, 87, 0)),
          (2.0, 3.0, 'Ruim', rgb(255, 199, 206), rgb(156, 0, 6)),
          (None, 2.0, 'Inaceitável', rgb(192, 0, 0), rgb(255, 255, 255))]
VELHA = 'Desempenho inadequado'
RX_L = re.compile(r'IFERROR\((\(\$K(\d+)-ABS\(\$G\2\)\)/mIncerteza\.IncertezaCIQ\([^()]*\)),""\)')
TITULO_LEG = 'Sigma — classe e regras de CQ sugeridas'


def log(m):
    print(m, flush=True)


def tabela_plano(wb):
    lo = pqlib.tabela(wb, 'tblPlanoQC_Sigma')
    if lo is None:
        raise SystemExit('tblPlanoQC_Sigma nao existe -- nada foi salvo')
    ws = lo.Parent
    cab = [str(c.Name) for c in lo.ListColumns]
    linhas = [list(r) for r in lo.DataBodyRange.Value]
    classes = [str(r[cab.index('Classificacao')] or '') for r in linhas]
    if 'Ruim' in classes and 'Inaceitável' in classes:
        return lo, cab
    if VELHA not in classes:
        raise SystemExit(f'tblPlanoQC_Sigma sem a faixa "{VELHA}": {classes} -- nada foi salvo')
    i = classes.index(VELHA)
    base = linhas[i]
    ws.Unprotect(ii.SENHA)
    r_ = lo.ListRows(i + 1).Range
    r_.Cells(1, cab.index('Sigma_Min') + 1).Value = 2.0
    r_.Cells(1, cab.index('Classificacao') + 1).Value = 'Ruim'
    nova = lo.ListRows.Add()
    for j, v in enumerate(base):
        nova.Range.Cells(1, j + 1).Value = v
    nova.Range.Cells(1, cab.index('Sigma_Min') + 1).Value = -999.0
    nova.Range.Cells(1, cab.index('Sigma_Max') + 1).Value = 2.0
    nova.Range.Cells(1, cab.index('Classificacao') + 1).Value = 'Inaceitável'
    MUDANCAS.append('tblPlanoQC_Sigma: "<3 Desempenho inadequado" -> "2-3 Ruim" + "<2 Inaceitável"')
    return lo, cab


def sigma_piso_zero(wb):
    e = wb.Worksheets('Estatística')
    n = 0
    for r in range(14, 134):
        c = e.Cells(r, 12)
        f = str(c.Formula)
        if not f.startswith('=') or 'MAX(0,' in f:
            continue
        if not RX_L.search(f):
            raise SystemExit(f'Estatística!L{r} inesperada: {f[:200]!r} -- nada foi salvo')
        c.Formula = RX_L.sub(lambda m: f'IFERROR(MAX(0,{m.group(1)}),"")', f, count=1)
        n += 1
    if n:
        MUDANCAS.append(f'Estatística L: Sigma com piso 0 ({n} linha(s))')


def local(ws, f_en):
    """Formula em ingles -> texto na lingua do Excel (FormatConditions.Add usa a lingua da interface)."""
    tmp = ws.Range('ZZ1000')
    tmp.Formula = f_en
    s = str(tmp.FormulaLocal)
    tmp.ClearContents()
    return s


def cores(ws, alvo, ref):
    """6 regras 'expressao' em alvo, pela celula numerica ref (relativa a 1a celula do alvo)."""
    fcs = ws.Range(alvo).FormatConditions
    fcs.Delete()
    for lo_, hi_, _, fundo, fonte in FAIXAS:
        cond = [f'ISNUMBER({ref})']
        if lo_ is not None:
            cond.append(f'{ref}>={lo_:g}')
        if hi_ is not None:
            cond.append(f'{ref}<{hi_:g}')
        fc = fcs.Add(2, None, local(ws, '=AND(' + ','.join(cond) + ')'))     # xlExpression; posicional (ADR-092)
        fc.Interior.Color = fundo
        fc.Font.Color = fonte
        fc.Font.Bold = True
        fc.StopIfTrue = True


def painel(wb, nlv, lo, cab):
    p = wb.Worksheets('Painel')
    ult = 6 + nlv
    # K = classe do Sigma (a mesma da Estatistica M)
    k6 = str(p.Range('K6').Value or '')
    if k6 not in ('', 'Classe Sigma') or any(str(p.Range(f'K{r}').Formula or '') not in ('',) and
                                              'Estatística!$M$' not in str(p.Range(f'K{r}').Formula)
                                              for r in range(7, ult + 1)):
        raise SystemExit('Painel K6:K9 ocupada -- nada foi salvo')
    if k6 != 'Classe Sigma':
        p.Range('K6').Value = 'Classe Sigma'
        p.Range('J6').Copy()
        p.Range('K6').PasteSpecial(-4122)                  # xlPasteFormats
        p.Application.CutCopyMode = False
        MUDANCAS.append('Painel K6 "Classe Sigma"')
    for r in range(7, ult + 1):
        f = (f'=IF($F{r}="-","",IFERROR(INDEX(Estatística!$M$14:$M$133,MATCH(selAnalito&"|"&{r - 6},'
             f'Estatística!$AB$14:$AB$133,0)),""))')
        if str(p.Range(f'K{r}').Formula).replace(' ', '') != f.replace(' ', ''):
            p.Range(f'K{r}').Formula = f
            MUDANCAS.append(f'Painel K{r}')
    p.Range(f'K7:K{ult}').ShrinkToFit = True
    p.Range(f'K7:K{ult}').HorizontalAlignment = -4108
    cores(p, f'I7:I{ult}', '$I7')
    cores(p, f'K7:K{ult}', '$I7')
    # legenda nas linhas 5 a 8, a direita da tabela do Westgard (V..X): os graficos ocupam a LARGURA DA JANELA a partir
    # da linha 9 (Bio) / 10 (Hema) -- mUI.AjustarGraficos --, entao abaixo da linha 8 a legenda ficaria por baixo deles
    # (qa_adr090 V02, 23/11/2026). 6 faixas em 2 linhas x 3 colunas: "faixa · classe · regras sugeridas"
    if str(p.Range('V5').Value or '') == TITULO_LEG:
        p.Range('V5:X12').Clear()                        # legenda de instalacao anterior (V5:W12)
    v = p.Range('V5:X8')
    ocup = [c.Address for c in v.Cells if c.Formula not in (None, '')]
    if ocup:
        raise SystemExit(f'Painel V5:X8 ocupada ({ocup[:4]}) -- nada foi salvo')
    for co in p.ChartObjects():
        if co.TopLeftCell.Row <= 8:
            raise SystemExit(f'Painel: grafico {co.Name} comeca na linha {co.TopLeftCell.Row} (legenda V5:X8) -- nada foi salvo')
    regras = {}
    for row in lo.DataBodyRange.Value:
        regras[str(row[cab.index('Classificacao')])] = str(row[cab.index('Regras')] or '')
    p.Range('V5').Value = TITULO_LEG + ' (informativo; as regras que rodam são as do Westgard configurado)'
    p.Range('V5').Font.Bold = True
    p.Range('V5').Font.Size = 8
    for i, (lo_, hi_, cls, fundo, fonte) in enumerate(FAIXAS):
        faixa = (f'≥ {lo_:g}' if hi_ is None else f'< {hi_:g}' if lo_ is None else f'{lo_:g} a {hi_:g}')
        c = p.Cells(6 + i // 3, 22 + i % 3)
        c.Value = f'{faixa} · {cls} · {regras.get(cls, "")}'
        c.Interior.Color = fundo
        c.Font.Color = fonte
        c.Font.Bold = True
        c.Font.Size = 7
        c.ShrinkToFit = True
        c.WrapText = False
    for col in (22, 23, 24):
        if p.Columns(col).ColumnWidth < 30:
            p.Columns(col).ColumnWidth = 34
    MUDANCAS.append(f'Painel I/K cores e legenda {v.Address}')


def estatistica(wb):
    e = wb.Worksheets('Estatística')
    cores(e, 'L14:L133', '$L14')
    cores(e, 'M14:M133', '$L14')


def fumaca(ex, wb, nlv):
    ex.xl.Calculate()
    e, p = wb.Worksheets('Estatística'), wb.Worksheets('Painel')
    neg = [r for r in range(14, 134) if isinstance(e.Cells(r, 12).Value, float) and e.Cells(r, 12).Value < 0]
    if neg:
        raise SystemExit(f'fumaca: Sigma negativo na Estatística L{neg[:5]} -- nada foi salvo')
    dif = []
    for r in range(7, 7 + nlv):
        i_, k_ = p.Range(f'I{r}').Value, str(p.Range(f'K{r}').Value or '')
        if isinstance(i_, float):
            esp = next(c for lo_, hi_, c, _, _ in FAIXAS if (lo_ is None or i_ >= lo_) and (hi_ is None or i_ < hi_))
            if k_ != esp:
                dif.append((r, i_, k_, esp))
    if dif:
        raise SystemExit(f'fumaca: classe do Painel K != faixa do Sigma {dif} -- nada foi salvo')
    if p.Range('I7').FormatConditions.Count != 6 or e.Range('L14').FormatConditions.Count != 6:
        raise SystemExit('fumaca: formatacao condicional do Sigma nao ficou com 6 faixas -- nada foi salvo')
    log(f'   ok: Painel I7 = {p.Range("I7").Text!r} {p.Range("K7").Text!r}')


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    nlv = 2 if produto.startswith('Bio') else 3
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
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('2. tblPlanoQC_Sigma (6 faixas)')
        lo, cab = tabela_plano(wb)
        log('3. Estatistica L: Sigma com piso 0')
        sigma_piso_zero(wb)
        log('4-6. Painel (classe, cores, legenda) e Estatistica (cores)')
        painel(wb, nlv, lo, cab)
        estatistica(wb)
        ex.xl.Calculation = -4105
        log('7. fumaca')
        fumaca(ex, wb, nlv)
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
