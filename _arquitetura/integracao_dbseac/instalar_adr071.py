# -*- coding: utf-8 -*-
"""instalar_adr071.py -- ADR-071: selecao de rodadas do CEQ (1 rodada, conjunto, ULTIMAS n, ACUMULADAS) e
Sigma com CV POOLED do CIQ numa janela de N meses (Bioquimica e Hematologia).

Uso:  python instalar_adr071.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Roda DEPOIS do instalar_adr068.py (exige o "-" do ADR-068 em Estatistica!L e Painel!I7). Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py: mCEQ le o texto novo de eqRodada; mIncerteza com
     cache por janela e Sigma_Faixa no recalculo; mDados gera as listas; mEQA atualiza as listas); compila;
  2. listas so com rodadas EXISTENTES (mDados.AtualizarListasAno -> Configuracao!AF:AJ) e os nomes
     lstAnosCAP/CTL, lstRodadasCAP/CTL/EQA (INDEX/CONT.VALORES, nunca DESLOC); lstAnosCEQ na Hematologia;
  3. Estatistica linha 11 (K11:Q11): Sigma_Meses (padrao 6, validacao 1-24; valor do usuario preservado) e
     Sigma_Ini = EDATE(MU_Fim,-Sigma_Meses)+1 -- ancora no ultimo resultado (MU_Fim), sem HOJE()/DESLOC;
  4. Estatistica!L (Sigma) = (ETp - |Bias|) / mIncerteza.IncertezaCIQ(...,"CV",Sigma_Ini,MU_Fim[,exclusoes]),
     com o "-" do ADR-068; nome Sigma_Faixa (sujo no RecalcularIncerteza); rotulo L13 diz a janela;
     H (ET%) continua com o CV do lote (F) -- decisao D6;
  5. Painel!I7:I8 (Bio) / I7:I9 (Hema) leem o Sigma de Estatistica!L pela chave AB (padrao de G7), com o "-";
     Cfg_PlanoQC!B1 = MIN(I7:I9) acompanha sem mudar (D2); rotulo I6;
  6. CEQ: Hematologia R4 sem vazio; N4/P4 com listas do provedor (estilo Aviso: o conjunto e digitado);
     Bioquimica P4 com a uniao CAP+Controllab e N4 com lstAnosCEQ; K5 = ResumoFiltroEQ (sem TEXTO/SOMARPRODUTO);
  7. FUMACA: janela, L = (ETp - |G|)/CV pooled pela propria funcao, Painel I = L, K5, listas = rodadas
     existentes (sem simulacao), validacoes; nenhum erro;
  8. devolve a protecao de cada aba e a da estrutura como estavam.
"""
import datetime as dt
import calendar
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402
from instalar_adr064 import LIN_TAB, ultima_linha_tabela, selecionar_a1, formato_data  # noqa: E402

MESES_PADRAO = 6
FIXOS = ['TODAS', 'ACUMULADAS', 'ULTIMAS 3', 'ULTIMAS 6']
# Configuracao: colunas das listas geradas por mDados.AtualizarListasAno (ADR-071)
LISTAS = {'lstAnosCAP': 'AF', 'lstAnosCTL': 'AG', 'lstRodadasCAP': 'AH', 'lstRodadasCTL': 'AI', 'lstRodadasEQA': 'AJ'}
# Estatistica linha 11: rotulo, meses, "de", inicio, "a", fim, nota
C_ROT, C_MESES, C_DE, C_INI, C_A, C_FIM, C_NOTA = 'K', 'L', 'M', 'N', 'O', 'P', 'Q'
XL_LIST, XL_WHOLE, XL_BETWEEN = 3, 1, 1
XL_STOP, XL_WARNING = 1, 2


def log(*a):
    print(*a, flush=True)


def traduzir(ws, formula):
    """Validacao e Names.Add leem a formula no idioma da instalacao: traduz por uma celula auxiliar."""
    aux = ws.Range('ZZ1000')
    antes = aux.Formula
    aux.Formula = formula
    loc = aux.FormulaLocal
    aux.Formula = antes
    return loc


def nome_local(wb, ws_aux, n, formula):
    loc = traduzir(ws_aux, formula)
    try:
        wb.Names(n).Delete()
    except Exception:
        pass
    wb.Names.Add(n, loc)
    return wb.Names(n).RefersTo


def edate_mais_1(serial, meses):
    """EDATE(serial, -meses) + 1, em serial do Excel."""
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))
    m0 = d.month - 1 - meses
    y, m = d.year + m0 // 12, m0 % 12 + 1
    dd = min(d.day, calendar.monthrange(y, m)[1])
    return float((dt.date(y, m, dd) - dt.date(1899, 12, 30)).days + 1)


def serial(v):
    if hasattr(v, 'toordinal'):
        return float(v.toordinal() - 693594)
    return float(v)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


# ---------------------------------------------------------------------------------------------- 2. listas
def listas(ex, wb, bio):
    run = lambda m, *x: ex.run("'" + wb.Name + "'!" + m, *x, teto=600)
    run('mDados.AtualizarListasAno')
    e = wb.Worksheets('Estatística')
    feitos = []
    for n, c in LISTAS.items():
        nome_local(wb, e, n, f"=Configuração!${c}$2:INDEX(Configuração!${c}$2:${c}$101,"
                             f"MAX(1,COUNTA(Configuração!${c}$2:${c}$101)))")
        feitos.append(n)
    if not any(x.Name == 'lstAnosCEQ' for x in wb.Names):
        nome_local(wb, e, 'lstAnosCEQ', "=Configuração!$AA$2:INDEX(Configuração!$AA$2:$AA$50,"
                                        "MAX(1,COUNTA(Configuração!$AA$2:$AA$50)))")
        feitos.append('lstAnosCEQ (criado)')
    return feitos


# ---------------------------------------------------------------------------------------------- 3. janela
def janela(wb, e):
    tem = any(x.Name == 'Sigma_Meses' for x in wb.Names)
    if not tem:
        for c in (C_ROT, C_MESES, C_DE, C_INI, C_A, C_FIM, C_NOTA):
            cel = e.Range(f'{c}11')
            if cel.MergeCells or str(cel.Formula or '').strip():
                raise SystemExit(f'Estatistica!{c}11 ocupada ({cel.Formula!r}): nada alterado')
    elif wb.Names('Sigma_Meses').RefersToRange.Address != f'${C_MESES}$11':
        raise SystemExit(f"Sigma_Meses aponta para {wb.Names('Sigma_Meses').RefersTo}: esperado Estatistica!{C_MESES}11")
    m = e.Range(f'{C_MESES}11')
    v = m.Value
    if not (num(v) and 1 <= v <= 24 and float(v).is_integer()):     # preserva a escolha do usuario
        m.Value = MESES_PADRAO
    m.Locked = False
    m.NumberFormatLocal = '0'
    m.HorizontalAlignment = -4108
    m.Font.Bold = True
    m.Interior.Color = 13434879                    # amarelo claro: celula de entrada
    m.Validation.Delete()
    m.Validation.Add(XL_WHOLE, XL_STOP, XL_BETWEEN, '1', '24')
    m.Validation.IgnoreBlank = True
    m.Validation.InputTitle = 'Janela do Sigma (meses)'
    m.Validation.InputMessage = ('CV pooled do CIQ (lotes n >= 20, agrupados pelos graus de liberdade) nos ultimos N '
                                 'meses ate o ultimo resultado (MU_Fim). 1 a 24; padrao 6.')
    m.Validation.ErrorMessage = 'Meses: numero inteiro de 1 a 24.'
    ii.nome(wb, 'Sigma_Meses', f"='Estatística'!${C_MESES}$11")
    e.Range(f'{C_ROT}11').Value = 'Sigma: CV pooled, meses'
    e.Range(f'{C_ROT}11').Font.Bold = True
    e.Range(f'{C_ROT}11').HorizontalAlignment = -4152
    e.Range(f'{C_DE}11').Value = 'de'
    e.Range(f'{C_A}11').Value = 'a'
    for c in (C_DE, C_A):
        e.Range(f'{c}11').HorizontalAlignment = -4108
    fd = formato_data(e.Application)
    ini = e.Range(f'{C_INI}11')
    ini.NumberFormatLocal = fd
    ini.Formula = '=IF(AND(ISNUMBER(MU_Fim),N(Sigma_Meses)>=1),EDATE(MU_Fim,-Sigma_Meses)+1,"")'
    ii.nome(wb, 'Sigma_Ini', f"='Estatística'!${C_INI}$11")
    fim = e.Range(f'{C_FIM}11')
    fim.NumberFormatLocal = fd
    fim.Formula = '=IF(ISNUMBER(MU_Fim),MU_Fim,"")'
    e.Range(f'{C_NOTA}11').Value = ('Sigma (L) = (ETp - |Bias|) / CV pooled desta janela; ET % (H) continua com o CV do '
                                    'lote (F). Sem meses ou sem dado: Sigma vazio (nunca o historico inteiro).')
    e.Range(f'{C_NOTA}11').Font.Size = 9
    e.Range(f'{C_NOTA}11').Font.Italic = True
    return 'ja definida' if tem else 'criada'


# ---------------------------------------------------------------------------------------------- 4. coluna L
def formula_l(r, exc):
    return (f'=IF($K{r}="-","-",IF(OR(NOT(ISNUMBER(Sigma_Ini)),NOT(ISNUMBER(MU_Fim)),NOT(ISNUMBER($G{r})),'
            f'NOT(ISNUMBER($K{r}))),"",IFERROR(($K{r}-ABS($G{r}))/'
            f'mIncerteza.IncertezaCIQ($A{r},$B{r},"CV",Sigma_Ini,MU_Fim{exc}),"")))')


def coluna_l(wb, e, ult, exc):
    f = str(e.Range(f'L{LIN_TAB}').Formula)
    novo = formula_l(LIN_TAB, exc)
    if f == novo:
        estado = 'ja instalada'
    else:
        if not f.startswith(f'=IF($K{LIN_TAB}="-","-",'):
            raise SystemExit(f'Estatistica!L{LIN_TAB} sem o "-" do ADR-068 (rode instalar_adr068.py antes): {f[:160]}')
        antiga = f'/$F{LIN_TAB}' in f
        nossa = 'mIncerteza.IncertezaCIQ(' in f and 'Sigma_Ini' in f
        if not (antiga or nossa):
            raise SystemExit(f'Estatistica!L{LIN_TAB}: formula inesperada, nada alterado: {f[:160]}')
        e.Range(f'L{LIN_TAB}:L{ult}').Formula = novo           # relativo: o Excel propaga linha a linha
        estado = 'instalada'
    ii.nome(wb, 'Sigma_Faixa', f"='Estatística'!$L${LIN_TAB}:$L${ult}")
    e.Range('L13').Formula = '=IF(N(Sigma_Meses)>=1,"SIGMA (CV pooled "&Sigma_Meses&" m)","SIGMA (defina os meses)")'
    e.Range('L13').WrapText = True
    return estado


# ---------------------------------------------------------------------------------------------- 5. Painel
def formula_i(r, k, ult):
    return (f"=IF($F{r}=\"-\",\"-\",IFERROR(INDEX('Estatística'!$L${LIN_TAB}:$L${ult},"
            f"MATCH(selAnalito&\"|\"&{k},'Estatística'!$AB${LIN_TAB}:$AB${ult},0)),\"\"))")


def painel(wb, nlv, ult):
    p = wb.Worksheets('Painel')
    feitos = []
    for k in range(1, nlv + 1):
        r = 6 + k
        f = str(p.Range(f'I{r}').Formula)
        novo = formula_i(r, k, ult)
        if f.replace("'Estatística'!", 'Estatística!') == novo.replace("'Estatística'!", 'Estatística!'):
            continue                                      # o Excel guarda sem as aspas
        if not f.startswith(f'=IF($F{r}="-","-",'):
            raise SystemExit(f'Painel!I{r} sem o "-" do ADR-068 (rode instalar_adr068.py antes): {f[:120]}')
        if not (('ABS(' in f and f'$E{r}' in f) or ('!$L$' in f and 'selAnalito' in f)):
            raise SystemExit(f'Painel!I{r}: formula inesperada, nada alterado: {f[:120]}')
        p.Range(f'I{r}').Formula = novo
        feitos.append(f'I{r}')
    p.Range('I6').Formula = '=IF(N(Sigma_Meses)>=1,"Sigma (CVp "&Sigma_Meses&" m)","Sigma")'
    c = wb.Worksheets('Cfg_PlanoQC')
    b1 = str(c.Range('B1').Formula)
    if f'Painel!$I$7:$I${6 + nlv}' not in b1:
        raise SystemExit(f'Cfg_PlanoQC!B1 inesperada (sigmaDoPlano deveria ser MIN(Painel!I7:I{6 + nlv})): {b1[:120]}')
    return feitos


# ---------------------------------------------------------------------------------------------- 6. CEQ
MSG_P4 = ('TODAS = ano vigente (<= Ano EP) | ACUMULADAS = todas ate o Ano EP | ULTIMAS n = as n mais recentes | '
          'conjunto: ano|rodada; ano|rodada (pode atravessar anos). Modos novos exigem o provedor. K5 confere.')


def validar(cel, formula, estilo, titulo, msg, ignorar_vazio=True):
    v = cel.Validation
    v.Delete()
    v.Add(XL_LIST, estilo, XL_BETWEEN, formula)
    v.IgnoreBlank = ignorar_vazio
    v.InCellDropdown = True
    if titulo:
        v.InputTitle = titulo
        v.InputMessage = msg
        v.ErrorTitle = titulo
        v.ErrorMessage = 'Texto fora da lista: confira em K5 se foi lido (conjunto = ano|rodada separados por ;).'
    v.ShowError = True


def ceq(wb, e, bio):
    feitos = []
    sep = e.Application.International[5 - 1]                 # xlListSeparator
    if not bio:
        r4 = wb.Names('eqProvedor').RefersToRange
        if r4.Address != '$R$4':
            raise SystemExit(f"Hematologia: eqProvedor em {r4.Address} (esperado R4; rode instalar_adr064.py antes)")
        if str(r4.Value or '').strip() == '':
            r4.Value = 'CAP'
        validar(r4, 'CAP' + sep + 'Controllab', XL_STOP, 'Programa do CEQ',
                'CAP ou Controllab -- obrigatorio: os programas nunca se misturam.', ignorar_vazio=False)
        r4.Validation.ErrorMessage = 'Escolha CAP ou Controllab (o provedor e obrigatorio).'
        feitos.append('R4 sem vazio')
        validar(e.Range('N4'), traduzir(e, '=IF(eqProvedor="Controllab",lstAnosCTL,lstAnosCAP)'), XL_WARNING,
                'Ano EP (teto)', 'Maior ano do CEQ usado em TODAS/ACUMULADAS/ULTIMAS n. Lista: anos com rodada do provedor.')
        validar(e.Range('P4'), traduzir(e, '=IF(eqProvedor="Controllab",lstRodadasCTL,lstRodadasCAP)'), XL_WARNING,
                'Rodada(s) do CEQ', MSG_P4)
        feitos.append('N4/P4 com listas do provedor')
    else:
        validar(e.Range('N4'), traduzir(e, '=lstAnosCEQ'), XL_WARNING, 'Ano EP (teto)',
                'Maior ano do CEQ usado em TODAS/ACUMULADAS/ULTIMAS n (formula: o ano mais recente).')
        validar(e.Range('P4'), traduzir(e, '=lstRodadasEQA'), XL_WARNING, 'Rodada(s) do CEQ', MSG_P4)
        feitos.append('N4/P4 com listas (provedor por analito: Analitos!AR)')
    for a in ('N4', 'P4'):
        e.Range(a).Locked = False
    if str(e.Range('P4').Value or '').strip() == '':
        e.Range('P4').Value = 'TODAS'
    e.Range('K5').Formula = '=ResumoFiltroEQ(eqProvedor,eqAnoEP,eqRodada)'
    feitos.append('K5')
    return feitos


# ---------------------------------------------------------------------------------------------- 7. fumaca
def rodadas_existentes(wb, provedor):
    """A regra de 'rodada existente' lida direto da EQA_Base (outra implementacao): Uso <> NAO, canonico, |bias|."""
    b = wb.Worksheets('EQA_Base')
    ult = b.Cells(b.Rows.Count, 1).End(-4162).Row
    out = {}
    if ult < 2:
        return []
    for r in b.Range(b.Cells(2, 1), b.Cells(ult, 20)).Value:
        if provedor and str(r[0] or '').strip().upper() != provedor.upper():
            continue
        if str(r[19] or '').strip().upper() == 'NAO' or not str(r[4] or '').strip():
            continue
        if not num(r[16]) or not num(r[1]):
            continue
        rot = str(int(r[2])) if isinstance(r[2], float) and r[2].is_integer() else str(r[2] or '').strip().upper()
        out[f'{int(r[1])}|{rot}'] = (int(r[1]), rot)

    def chave(it):
        a, rt = it
        return (a, (0, int(rt), '') if rt.isdigit() else (1, 0, rt))
    return [k for k, _ in sorted(out.items(), key=lambda kv: chave(kv[1]), reverse=True)]


def fumaca(ex, wb, e, ult, bio, nlv):
    run = lambda m, *x: ex.run("'" + wb.Name + "'!" + m, *x, teto=600)
    ruins = []
    meses, fim, ini = e.Range(f'{C_MESES}11').Value, e.Range('AI11').Value, e.Range(f'{C_INI}11').Value
    if not (num(meses) and fim not in (None, '') and ini not in (None, '')
            and abs(serial(ini) - edate_mais_1(serial(fim), int(meses))) < 1e-9):
        ruins.append(('janela', meses, fim, ini))
    exc_vazia = True
    if bio:
        exc_vazia = all(v in (None, '') for row in e.Range('Estat_Exclusoes').Value for v in row)
    cont = {'numero': 0, 'vazio': 0, '-': 0, 'conferidos': 0}
    for r in range(LIN_TAB, ult + 1):
        an = e.Range(f'A{r}').Value
        if an in (None, '', 0):
            continue
        k, g, lv = e.Range(f'K{r}').Value, e.Range(f'G{r}').Value, e.Range(f'L{r}').Value
        if str(e.Range(f'L{r}').Text).startswith('#'):
            ruins.append((r, an, 'erro em L'))
            continue
        if k == '-':
            cont['-'] += 1
            if lv != '-':
                ruins.append((r, an, 'sem ETp mas L sem "-"', lv))
            continue
        if not num(lv):
            cont['vazio'] += 1
            if lv not in ('', None):
                ruins.append((r, an, 'L inesperado', lv))
            continue
        cont['numero'] += 1
        if not (num(k) and num(g)):
            ruins.append((r, an, 'L numerico sem ETp/bias', k, g))
            continue
        if exc_vazia:
            nv = e.Range(f'B{r}').Value
            cv = run('mIncerteza.IncertezaCIQ', str(an), nv, 'CV', serial(ini), serial(fim))
            if not num(cv) or abs(lv - (k - abs(g)) / cv) > 1e-9 * max(1.0, abs(lv)):
                ruins.append((r, an, 'L != (ETp-|G|)/CV pooled', lv, cv))
            cont['conferidos'] += 1
    # Painel I = L do analito em tela
    p = wb.Worksheets('Painel')
    sel = str(wb.Names('selAnalito').RefersToRange.Value)
    chaves = {str(e.Range(f'AB{r}').Value): r for r in range(LIN_TAB, ult + 1)}
    for k_ in range(1, nlv + 1):
        r = chaves.get(f'{sel}|{k_}')
        pv = p.Range(f'I{6 + k_}').Value
        ev = e.Range(f'L{r}').Value if r else ''
        if pv != ev and not (num(pv) and num(ev) and abs(pv - ev) < 1e-12):
            ruins.append(('Painel I', 6 + k_, pv, ev))
    k5 = str(e.Range('K5').Text)
    if 'linha(s)' not in k5 or 'INVALIDA' in k5 or k5.startswith('#'):
        ruins.append(('K5', k5[:120]))
    # listas = rodadas existentes, sem simulacao
    cfg = wb.Worksheets('Configuração')
    for nome_, prov in (('lstRodadasCAP', 'CAP'), ('lstRodadasCTL', 'Controllab'), ('lstRodadasEQA', '')):
        col = LISTAS[nome_]
        vals = [cfg.Range(f'{col}{i}').Value for i in range(2, 102)]
        vals = [str(v) for v in vals if v not in (None, '')]
        esperado = FIXOS + rodadas_existentes(wb, prov)
        if [v.upper() for v in vals] != [x.upper() for x in esperado]:
            ruins.append((nome_, vals[:10], esperado[:10]))
    v = e.Range('P4').Validation
    try:
        tipo, estilo = v.Type, v.AlertStyle
    except Exception:
        tipo = estilo = None
    if tipo != XL_LIST or estilo != XL_WARNING:
        ruins.append(('validacao P4', tipo, estilo))
    if ruins:
        raise SystemExit(f'fumaca ADR-071 falhou: {ruins[:6]}')
    return cont, k5


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    bio = produto.startswith('Bio')
    nlv = 2 if bio else 3
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

        e = wb.Worksheets('Estatística')
        ult = ultima_linha_tabela(e)
        exc = ',Estat_Exclusoes' if any(n.Name == 'Estat_Exclusoes' for n in wb.Names) else ''
        log(f'2. listas (so rodadas existentes) e nomes: {listas(ex, wb, bio)}')
        log(f'3. janela do Sigma (Estatistica {C_ROT}11:{C_NOTA}11): {janela(wb, e)}')
        log(f'4. Estatistica!L{LIN_TAB}:L{ult} (Sigma com CV pooled{" e exclusoes" if exc else ""}): {coluna_l(wb, e, ult, exc)}')
        log(f'5. Painel I7:I{6 + nlv} le Estatistica!L: {painel(wb, nlv, ult) or "ja instalado"}')
        log(f'6. CEQ: {ceq(wb, e, bio)}')

        log('7. fumaca')
        ex.xl.Calculation = -4105
        ex.run("'" + wb.Name + "'!mIncerteza.RecalcularIncerteza", teto=600)
        ex.esperar()
        cont, k5 = fumaca(ex, wb, e, ult, bio, nlv)
        log(f'   janela {e.Range(f"{C_INI}11").Text} a {e.Range(f"{C_FIM}11").Text} ({e.Range(f"{C_MESES}11").Value:.0f} m); '
            f'Sigma: {cont}')
        log(f'   K5: {k5}')
        selecionar_a1(e)

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
