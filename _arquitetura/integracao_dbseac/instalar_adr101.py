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
ADR-103 (26/11/2026): a classe, o status e as cores seguem a REGRA INSTITUCIONAL E IMUTAVEL da Metrica Sigma (5 faixas,
operadores literais; sigma = 4 Marginal, sigma = 5 Bom) -- FAIXAS. A tblPlanoQC_Sigma (faixas do plano de CQ) so e lida
para a regra de CQ sugerida na legenda. Tambem: Estatistica A98:A102 (Bio, resumo), F3 (Hema, legenda) e a cor de
Painel!O4. A fumaca confere mQualidade.ClassificarSigma/StatusSigma contra o oraculo nas fronteiras.
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
import painel_layout  # noqa: E402  (ADR-103: status/classe do topo do Painel)

MUDANCAS = []


def rgb(r, g, b):
    return r + g * 256 + b * 65536


# ADR-103 (26/11/2026): REGRA INSTITUCIONAL E IMUTAVEL DA METRICA SIGMA -- literal, nao alterar.
# (faixa, condicoes sobre {r}, classificacao, status, fundo, fonte, sigma representativo da faixa -- so para achar a regra
# de CQ sugerida na tblPlanoQC_Sigma). Cores: as do ADR-101 (sigma < 3 = a de "Inaceitavel"); nenhuma cor nova.
FAIXAS = [('σ ≥ 6', ['{r}>=6'], 'Excelente — Classe mundial', 'Excelente',
           rgb(0, 97, 0), rgb(255, 255, 255), 6.5),
          ('5 < σ < 6', ['{r}>5', '{r}<6'], 'Excelente', 'Alto desempenho',
           rgb(0, 176, 80), rgb(255, 255, 255), 5.5),
          ('4 < σ ≤ 5', ['{r}>4', '{r}<=5'], 'Bom', 'Desempenho adequado',
           rgb(198, 239, 206), rgb(0, 97, 0), 4.5),
          ('3 ≤ σ ≤ 4', ['{r}>=3', '{r}<=4'], 'Marginal', 'Atenção',
           rgb(255, 235, 156), rgb(156, 87, 0), 3.5),
          ('σ < 3', ['{r}<3'], 'Insatisfatório', 'método mais instável ou tendencioso',
           rgb(192, 0, 0), rgb(255, 255, 255), 2.5)]


def classe_oficial(s):
    """Oraculo Python da regra institucional (os mesmos operadores literais). Devolve a tupla da faixa."""
    if s >= 6:
        return FAIXAS[0]
    if 5 < s < 6:
        return FAIXAS[1]
    if 4 < s <= 5:
        return FAIXAS[2]
    if 3 <= s <= 4:
        return FAIXAS[3]
    return FAIXAS[4]


RESUMO_VELHO = ['Classe mundial', 'Excelente', 'Bom', 'Marginal', 'Desempenho inadequado']
LEGENDA_HEMA = ('Sigma: σ ≥ 6 Excelente — Classe mundial · 5 < σ < 6 Excelente · 4 < σ ≤ 5 Bom · 3 ≤ σ ≤ 4 Marginal · '
                'σ < 3 Insatisfatório · CV% deve ficar ≤ limites CLIA/VB.')
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


def regra_sugerida(lo, cab, s):
    """ADR-103: regra de CQ sugerida (coluna Regras do plano) para o Sigma s -- Sigma_Min <= s < Sigma_Max do plano."""
    for row in lo.DataBodyRange.Value:
        a, b = row[cab.index('Sigma_Min')], row[cab.index('Sigma_Max')]
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and a <= s < b:
            return str(row[cab.index('Regras')] or '')
    return ''


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
    """ADR-103: 5 regras 'expressao' (faixas institucionais, operadores literais) em alvo, pela celula numerica ref."""
    fcs = ws.Range(alvo).FormatConditions
    fcs.Delete()
    for _, conds, _, _, fundo, fonte, _ in FAIXAS:
        cond = [f'ISNUMBER({ref})'] + [c.format(r=ref) for c in conds]
        fc = fcs.Add(2, None, local(ws, '=AND(' + ','.join(cond) + ')'))     # xlExpression; posicional (ADR-092)
        fc.Interior.Color = fundo
        fc.Font.Color = fonte
        fc.Font.Bold = True
        fc.StopIfTrue = True


def painel(wb, nlv, lo, cab):
    p = wb.Worksheets('Painel')
    ult = 6 + nlv
    # K = classe do Sigma (a mesma da Estatistica M); ADR-103: no layout novo a classe esta em J (painel_layout)
    K = painel_layout.col_classe(p)
    k6 = str(p.Range(f'{K}6').Value or '')
    if k6 not in ('', 'Classe Sigma') or any(str(p.Range(f'{K}{r}').Formula or '') not in ('',) and
                                              'Estatística!$M$' not in str(p.Range(f'{K}{r}').Formula)
                                              for r in range(7, ult + 1)):
        raise SystemExit(f'Painel {K}6:{K}9 ocupada -- nada foi salvo')
    if k6 != 'Classe Sigma':
        p.Range(f'{K}6').Value = 'Classe Sigma'
        p.Range('J6').Copy()
        p.Range(f'{K}6').PasteSpecial(-4122)               # xlPasteFormats
        p.Application.CutCopyMode = False
        MUDANCAS.append(f'Painel {K}6 "Classe Sigma"')
    for r in range(7, ult + 1):
        f = (f'=IF($F{r}="-","",IFERROR(INDEX(Estatística!$M$14:$M$133,MATCH(selAnalito&"|"&{r - 6},'
             f'Estatística!$AB$14:$AB$133,0)),""))')
        if str(p.Range(f'{K}{r}').Formula).replace(' ', '') != f.replace(' ', ''):
            p.Range(f'{K}{r}').Formula = f
            MUDANCAS.append(f'Painel {K}{r}')
    p.Range(f'{K}7:{K}{ult}').ShrinkToFit = True
    p.Range(f'{K}7:{K}{ult}').HorizontalAlignment = -4108
    cores(p, f'I7:I{ult}', '$I7')
    cores(p, f'{K}7:{K}{ult}', '$I7')
    # legenda nas linhas 5 a 8, a direita da tabela do Westgard (V..X): os graficos ocupam a LARGURA DA JANELA a partir
    # da linha 9 (Bio) / 10 (Hema) -- mUI.AjustarGraficos --, entao abaixo da linha 8 a legenda ficaria por baixo deles
    # (qa_adr090 V02, 23/11/2026). 6 faixas em 2 linhas x 3 colunas: "faixa · classe · regras sugeridas"
    if str(p.Range('V5').Value or '').startswith(TITULO_LEG):   # 24/11: a celula leva o titulo + ' (informativo...)' -- reinstalacao
        p.Range('V5:X12').Clear()                        # legenda de instalacao anterior (V5:W12)
    v = p.Range('V5:X8')
    ocup = [c.Address for c in v.Cells if c.Formula not in (None, '')]
    if ocup:
        raise SystemExit(f'Painel V5:X8 ocupada ({ocup[:4]}) -- nada foi salvo')
    for co in p.ChartObjects():
        if co.TopLeftCell.Row <= 8:
            raise SystemExit(f'Painel: grafico {co.Name} comeca na linha {co.TopLeftCell.Row} (legenda V5:X8) -- nada foi salvo')
    p.Range('V5').Value = TITULO_LEG + ' (informativo; as regras que rodam são as do Westgard configurado)'
    p.Range('V5').Font.Bold = True
    p.Range('V5').Font.Size = 8
    # ADR-103: "faixa · classificacao · status · regra de CQ sugerida" (textos literais da regra institucional)
    for i, (faixa, _, cls, st, fundo, fonte, rep_) in enumerate(FAIXAS):
        c = p.Cells(6 + i // 3, 22 + i % 3)
        c.Value = f'{faixa} · {cls} · {st} · {regra_sugerida(lo, cab, rep_)}'
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


def estatistica(wb, nlv):
    e = wb.Worksheets('Estatística')
    cores(e, 'L14:L133', '$L14')
    cores(e, 'M14:M133', '$L14')
    # ADR-103: o rotulo mais longo ("Excelente — Classe mundial") passa ~3 % da largura da coluna M (qa_adr090 V02)
    if e.Range('M14').ShrinkToFit is not True:
        e.Range('M14:M133').ShrinkToFit = True
        MUDANCAS.append('Estatística M14:M133 reduzir para caber')
    if nlv == 2:                                  # Bio: quadro de resumo A98:A102 (COUNTIF da coluna M pelo rotulo)
        atual = [str(e.Range(f'A{98 + i}').Value or '') for i in range(5)]
        novo = [f[2] for f in FAIXAS]
        if atual == RESUMO_VELHO:
            for i, t in enumerate(novo):
                e.Range(f'A{98 + i}').Value = t
            MUDANCAS.append('Estatística A98:A102: rótulos da regra institucional')
        elif atual != novo:
            raise SystemExit(f'Estatística A98:A102 inesperada: {atual} -- nada foi salvo')
    else:                                         # Hema: legenda F3
        f3 = str(e.Range('F3').Value or '')
        if f3 != LEGENDA_HEMA:
            if not f3.startswith('Sigma:'):
                raise SystemExit(f'Estatística F3 inesperada: {f3[:80]!r} -- nada foi salvo')
            e.Range('F3').Value = LEGENDA_HEMA
            MUDANCAS.append('Estatística F3: legenda da regra institucional')


def cor_o4(wb):
    """ADR-103: a cor do Sigma do plano (Painel O4: regras do instalar_adr068, verde com O4>=4 e vermelho com O4<3) segue a
    faixa institucional -- verde so com sigma > 4 (sigma = 4 e Marginal). Mesmas cores; so o operador."""
    p = wb.Worksheets('Painel')
    fcs = p.Range('O4').FormatConditions
    for i in range(1, fcs.Count + 1):
        fc = fcs(i)
        if fc.Type != 2:                              # xlExpression
            continue
        f = str(fc.Formula1)
        if '$O$4>=4' in f.replace(' ', ''):
            fc.Modify(2, None, f.replace(' ', '').replace('$O$4>=4', '$O$4>4'))
            MUDANCAS.append('Painel O4: verde só com σ > 4')


def fumaca(ex, wb, nlv):
    ex.xl.Calculate()
    e, p = wb.Worksheets('Estatística'), wb.Worksheets('Painel')
    neg = [r for r in range(14, 134) if isinstance(e.Cells(r, 12).Value, float) and e.Cells(r, 12).Value < 0]
    if neg:
        raise SystemExit(f'fumaca: Sigma negativo na Estatística L{neg[:5]} -- nada foi salvo')
    dif = []
    for r in range(7, 7 + nlv):
        i_, k_ = p.Range(f'I{r}').Value, str(p.Range(f'{painel_layout.col_classe(p)}{r}').Value or '')
        if isinstance(i_, float):
            esp = classe_oficial(i_)[2]
            if k_ != esp:
                dif.append((r, i_, k_, esp))
    if dif:
        raise SystemExit(f'fumaca: classe do Painel K != faixa do Sigma {dif} -- nada foi salvo')
    if p.Range('I7').FormatConditions.Count != 5 or e.Range('L14').FormatConditions.Count != 5:
        raise SystemExit('fumaca: formatacao condicional do Sigma nao ficou com 5 faixas -- nada foi salvo')
    # ADR-103: a funcao VBA instalada (unica escada do produto) contra o oraculo, nas fronteiras e logo acima/abaixo
    pontos = [6.01, 6.0, 5.999999, 5.99, 5.000001, 5.0, 4.999999, 4.99, 4.000001, 4.0, 3.999999, 3.99, 3.000001, 3.0,
              2.999999, 2.99, 0.0, -1.0]
    ruins = []
    for s_ in pontos:
        c_ = str(ex.run(f"'{wb.Name}'!mQualidade.ClassificarSigma", s_))
        t_ = str(ex.run(f"'{wb.Name}'!mQualidade.StatusSigma", s_))
        f_ = classe_oficial(s_)
        if (c_, t_) != (f_[2], f_[3]):
            ruins.append((s_, c_, t_, f_[2], f_[3]))
    if ruins:
        raise SystemExit(f'fumaca: ClassificarSigma/StatusSigma != regra institucional {ruins} -- nada foi salvo')
    for vazio in ('', 'abc'):
        if str(ex.run(f"'{wb.Name}'!mQualidade.ClassificarSigma", vazio)) != '':
            raise SystemExit(f'fumaca: ClassificarSigma({vazio!r}) deveria ser vazio -- nada foi salvo')
    log(f'   ok: regra institucional conferida em {len(pontos)} pontos de fronteira')
    log(f'   ok: Painel I7 = {p.Range("I7").Text!r} {p.Range(painel_layout.col_classe(p) + "7").Text!r}')


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
        log('2. tblPlanoQC_Sigma (so leitura: regra de CQ sugerida)')
        lo, cab = tabela_plano(wb)
        log('3. Estatistica L: Sigma com piso 0')
        sigma_piso_zero(wb)
        log('4-6. Painel (classe, cores, legenda) e Estatistica (cores)')
        painel(wb, nlv, lo, cab)
        estatistica(wb, nlv)
        cor_o4(wb)
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
