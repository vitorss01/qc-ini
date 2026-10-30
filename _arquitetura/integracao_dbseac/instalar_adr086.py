# -*- coding: utf-8 -*-
"""instalar_adr086.py -- ADR-086 (29/10/2026): Painel H3 vazio = TODOS os lotes do periodo.

Uso:  python instalar_adr086.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente (a 2a execucao registra 0 mudancas). Passos:
  1. todo o VBA atual das fontes (codigo_atual.py: mEstatistica com o modo TODOS -- chave RUN|LOTE, alvo do lote de cada
     ponto, Westgard lote a lote, Eng_Saida BW..CB com a Media/DP de cada ponto e AD1 = engTodos; mLotes.TodosOsLotes);
  2. nomes engTodos (Eng_Saida!AD1), engAlvoMN<i>/engAlvoSN<i> (Media/DP de cada ponto do nivel i);
  3. Calc: 7 colunas novas por nivel (cabecalho "T<i> m3".."T<i> p3", na 1a area livre a direita) com a linha de limite
     EM DEGRAU de cada ponto (Media + k*DP do lote do ponto, so no modo TODOS; NA() no modo normal);
     linhas 184:185 das colunas de limite (ADR-078) = NA() no modo TODOS; moldura do eixo Y (min/max) tambem pelas
     colunas novas no modo TODOS;
  4. graficos do Painel: as 7 series de limite de cada nivel leem (Calc <corrida>3:182, <corrida>184:185) x
     (Calc <T>3:182, <limite>184:185) -- no modo normal so as 2 bordas tem numero (o desenho do ADR-078 nao muda);
  5. Painel H4 diz o modo TODOS;
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
import codigo_atual  # noqa: E402
import instalar_adr078 as a78  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MUDANCAS = []
LIM, SUF, KS = a78.LIM, a78.SUF, a78.KS
ENG_ALVO0 = 75                                  # Eng_Saida BW (mEstatistica.ENG_COL_ALVO0)
H4_TODOS = '"▲ H3 vazio = TODOS os lotes do período (cada ponto com o alvo do seu lote)"'


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
    # 29/10/2026: '=...=TRUE' virou "='TRUE'" no Excel em portugues e o arquivo salvo nao abria: o nome tem de voltar
    # como foi pedido (fora as aspas de nome de aba)
    def sem_aba(s):
        return re.sub(r"'([^']+)'!", r'\1!', s).replace('$', '').upper()
    if sem_aba(str(wb.Names(nm).RefersTo)) != sem_aba(ref):
        raise SystemExit(f'nome {nm}: o Excel gravou {wb.Names(nm).RefersTo!r} em vez de {ref!r} -- nada foi salvo')
    MUDANCAS.append(f'nome {nm}')


def setf(rng, f, oque):
    if a78.norm(rng.Formula) != a78.norm(f):
        rng.Formula = f
        MUDANCAS.append(oque)


RX_SER = re.compile(r'^=SERIES\((.*),Calc!\$([A-Z]+)\$184:\$[A-Z]+\$185,Calc!\$([A-Z]+)\$184:\$[A-Z]+\$185,(\d+)\)$')
RX_SER86 = re.compile(r'^=SERIES\((.*),\(Calc!\$([A-Z]+)\$3:\$[A-Z]+\$182,Calc!\$[A-Z]+\$184:\$[A-Z]+\$185\),'
                      r'\(Calc!\$([A-Z]+)\$3:\$[A-Z]+\$182,Calc!\$([A-Z]+)\$184:\$[A-Z]+\$185\),(\d+)\)$')


def series_limite(wb, i):
    """{nome: (xcol, ycol184, formula)} das 7 series de limite do grafico i (padrao ADR-078 ou ja ADR-086)."""
    ch = wb.Worksheets('Painel').ChartObjects(i).Chart
    out = {}
    for s_i in range(1, ch.SeriesCollection().Count + 1):
        s = ch.SeriesCollection(s_i)
        nm = str(s.Name)
        if nm not in LIM:
            continue
        f = str(s.Formula)
        m = RX_SER.match(f)
        if m:
            out[nm] = (m.group(2), m.group(3), s, m.group(1), m.group(4), None)
            continue
        m = RX_SER86.match(f)
        if m:
            out[nm] = (m.group(2), m.group(4), s, m.group(1), m.group(5), m.group(3))
            continue
        raise SystemExit(f'grafico {i}, serie {nm}: formula fora do padrao: {f}')
    if set(out) != set(LIM):
        raise SystemExit(f'grafico {i}: series de limite ausentes: {set(LIM) - set(out)}')
    return out


def colunas_t(c, nlv):
    """Colunas das linhas em degrau: as que ja tem o cabecalho T1 m3.. ou a 1a area livre (com uma coluna de folga)."""
    cab = [str(v or '').strip() for v in c.Range('A2:ZZ2').Value[0]]
    if 'T1 m3' in cab:
        c0 = cab.index('T1 m3') + 1
    else:
        ult = max(c.UsedRange.Column + c.UsedRange.Columns.Count - 1, 1)
        c0 = ult + 2
        bloco = c.Range(c.Cells(1, c0), c.Cells(185, c0 + 7 * nlv - 1))
        if c.Application.WorksheetFunction.CountA(bloco) > 0:
            raise SystemExit(f'Calc: area {a78.col(c0)}..{a78.col(c0 + 7 * nlv - 1)} nao esta vazia -- nada foi feito')
    return {i: [c0 + (i - 1) * 7 + k for k in range(7)] for i in range(1, nlv + 1)}


def calc(wb, nlv, m):
    c = wb.Worksheets('Calc')
    e = wb.Worksheets('Eng_Saida')
    for a, ok in (('AC1', ('', 'todos os lotes:')),):
        if str(e.Range(a).Value or '') not in ok:
            raise SystemExit(f'Eng_Saida!{a} ocupada: {e.Range(a).Value!r}')
    # N(...)<>0: AD1 vazio (motor ainda nao rodou desde a instalacao) vale FALSO -- IF(vazio) dava #VALOR! no Painel H4.
    # Sem literal TRUE: no Excel em portugues o Names.Add gravou ='TRUE' e o arquivo salvo nao abria mais
    nome(wb, 'engTodos', '=N(Eng_Saida!$AD$1)<>0')
    for i in range(1, nlv + 1):
        cm, cs = ENG_ALVO0 + 2 * (i - 1), ENG_ALVO0 + 2 * (i - 1) + 1
        nome(wb, f'engAlvoMN{i}', f'=Eng_Saida!${a78.col(cm)}$3:${a78.col(cm)}$182')
        nome(wb, f'engAlvoSN{i}', f'=Eng_Saida!${a78.col(cs)}$3:${a78.col(cs)}$182')
    tcols = colunas_t(c, nlv)
    a3 = c.Range('A3').Value
    if not isinstance(a3, (int, float)) or int(a3) != 1:
        raise SystemExit(f'Calc!A3 = {a3!r}: esperado o slot 1 (INDEX(engRUN,$A3))')
    xcols = {}
    for i in range(1, nlv + 1):
        ser = series_limite(wb, i)
        xs = {v[0] for v in ser.values()}
        if len(xs) != 1:
            raise SystemExit(f'grafico {i}: as series de limite tem X diferentes: {xs}')
        xcols[i] = xs.pop()
        lims = m['niveis'][i - 1]['cols']
        for k_i, (cc, k) in enumerate(zip(tcols[i], KS)):
            if str(c.Cells(2, cc).Value or '') != f'T{i} {SUF[k_i]}':
                c.Cells(2, cc).Value = f'T{i} {SUF[k_i]}'
                MUDANCAS.append(f'Calc cabecalho T{i} {SUF[k_i]}')
            y = f'INDEX(engAlvoMN{i},$A3)' if k == 0 else f'INDEX(engAlvoMN{i},$A3)+({k})*INDEX(engAlvoSN{i},$A3)'
            f = (f'=IF(AND(engTodos,ISNUMBER(${xcols[i]}3)),IF(ISNUMBER(INDEX(engAlvoSN{i},$A3)),{y},NA()),NA())')
            if (a78.norm(c.Cells(3, cc).Formula) != a78.norm(f)
                    or str(c.Cells(3, cc).FormulaR1C1) != str(c.Cells(182, cc).FormulaR1C1)):
                c.Range(c.Cells(3, cc), c.Cells(182, cc)).Formula = f        # relativo: $A3 -> $A4 ...
                MUDANCAS.append(f'Calc {a78.col(cc)}3:182 (T{i} {SUF[k_i]})')
            # 184:185 das colunas do ADR-078: NA() no modo TODOS
            for r in (184, 185):
                old = str(c.Cells(r, lims[k_i]).Formula)
                if 'engTodos' not in old:
                    if not old.startswith('=IF(OR('):
                        raise SystemExit(f'Calc {a78.col(lims[k_i])}{r} inesperada: {old[:80]}')
                    c.Cells(r, lims[k_i]).Formula = '=IF(engTodos,NA(),' + old[1:] + ')'
                    MUDANCAS.append(f'Calc {a78.col(lims[k_i])}{r} modo TODOS')
    return tcols, xcols


def eixo_y(wb, nlv, m, tcols):
    """Moldura do eixo Y (celulas da linha 1 com -3.3*DP / +3.3*DP do nivel): no modo TODOS, pelas linhas em degrau."""
    c = wb.Worksheets('Calc')
    lin1 = [str(c.Cells(1, k).Formula or '') for k in range(1, 200)]
    for i in range(1, nlv + 1):
        M = m['niveis'][i - 1]['M']
        mn = [k + 1 for k, f in enumerate(lin1) if f'{M}-3.3*' in f.replace(' ', '')]
        mx = [k + 1 for k, f in enumerate(lin1) if f'{M}+3.3*' in f.replace(' ', '')]
        if len(mn) != 1 or len(mx) != 1:
            raise SystemExit(f'Calc linha 1: moldura do eixo Y do N{i} nao encontrada (min {mn}, max {mx})')
        t3, tp3 = a78.col(tcols[i][0]), a78.col(tcols[i][6])
        for k, agg in ((mn[0], 5), (mx[0], 4)):
            old = lin1[k - 1]
            if 'engTodos' in old:                     # ja instalado (a formula comeca com =IF(AND(engTodos,...)
                continue
            faixa = f'(_xlfn.AGGREGATE(4,6,${tp3}$3:${tp3}$182)-_xlfn.AGGREGATE(5,6,${t3}$3:${t3}$182))'
            if agg == 5:
                todos = f'_xlfn.AGGREGATE(5,6,${t3}$3:${t3}$182)-0.05*{faixa}'
            else:
                todos = f'_xlfn.AGGREGATE(4,6,${tp3}$3:${tp3}$182)+0.05*{faixa}'
            # os dados do nivel (pontos e X): os mesmos intervalos que a formula antiga ja usa
            dados = re.findall(r'\$[A-Z]+\$3:\$[A-Z]+\$182', old)
            termos = [f'IF(COUNT({d})=0,{todos},_xlfn.AGGREGATE({agg},6,{d}))' for d in dados]
            fn = 'MIN' if agg == 5 else 'MAX'
            novo = (f'=IF(AND(engTodos,COUNT(${t3}$3:${t3}$182)>0),IFERROR({fn}({todos}'
                    + ''.join(',' + x for x in termos) + '),""),' + old[1:] + ')')
            c.Cells(1, k).Formula = novo
            MUDANCAS.append(f'Calc {a78.col(k)}1 eixo Y modo TODOS')


def graficos(wb, nlv, m, tcols):
    for i in range(1, nlv + 1):
        ser = series_limite(wb, i)
        lims = m['niveis'][i - 1]['cols']
        for nm, (xc, y184, s, nome_ser, ordem, ty) in ser.items():
            k_i = LIM.index(nm)
            esperado184 = a78.col(lims[k_i])
            if y184 != esperado184:
                raise SystemExit(f'grafico {i}, serie {nm}: Y em {y184}, esperado {esperado184}')
            tc = a78.col(tcols[i][k_i])
            novo = (f'=SERIES({nome_ser},(Calc!${xc}$3:${xc}$182,Calc!${xc}$184:${xc}$185),'
                    f'(Calc!${tc}$3:${tc}$182,Calc!${y184}$184:${y184}$185),{ordem})')
            if str(s.Formula) != novo:
                s.Formula = novo
                MUDANCAS.append(f'grafico {i} serie {nm}')


def painel(wb):
    p = wb.Worksheets('Painel')
    old = str(p.Range('H4').Formula)
    if 'engTodos' in old:
        return
    if not old.startswith('=IF(engAviso'):
        raise SystemExit(f'Painel!H4 inesperada: {old[:80]}')
    p.Range('H4').Formula = f'=IF(engTodos,{H4_TODOS},' + old[1:] + ')'
    MUDANCAS.append('Painel H4')


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
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        m = a78.mapa(wb, nlv)
        tcols, xcols = calc(wb, nlv, m)
        log(f'2/3. nomes engTodos/engAlvo*; Calc linhas em degrau {[a78.col(v[0]) for v in tcols.values()]} '
            f'(X {xcols}); 184:185 NA no modo TODOS')
        eixo_y(wb, nlv, m, tcols)
        log('3b. moldura do eixo Y no modo TODOS')
        graficos(wb, nlv, m, tcols)
        log('4. series de limite: corridas 3:182 + bordas 184:185')
        painel(wb)
        log('5. Painel H4')
        ex.xl.Calculation = -4105
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS[:12]) + (' ...' if len(MUDANCAS) > 12 else '')
                                          if MUDANCAS else ''))
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
