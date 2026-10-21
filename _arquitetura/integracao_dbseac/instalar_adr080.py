# -*- coding: utf-8 -*-
"""instalar_adr080.py -- ADR-080: o eixo X do grafico de cada nivel e a CORRIDA DO LOTE (1, 2, 3... por nivel).

Uso:  python instalar_adr080.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Pedido do usuario (21/10/2026): cada nivel com as suas corridas em sequencia, definidas por cada passagem do controle
(inclusive a noite, no mesmo dia), comecando do 1 no primeiro resultado daquele lote. Antes, o X era a POSICAO na
janela das 180 corridas, comum aos niveis: corrida so de N2 deixava buraco no N1, e com filtro de periodo o primeiro
numero nao era o do lote.

Idempotente (a 2a execucao registra 0 mudancas). So:
  1. instala todo o VBA atual das fontes (codigo_atual.py) -- o motor publica a corrida do lote de cada nivel em
     Eng_Saida BK..BM (mEstatistica.ENG_COL_CORL0) e o mUI.AtualizarEixos usa xrunmin<i>/xrunmax<i>; EXIGE compilacao;
  2. nomes engCorN1..3 (Eng_Saida BK3:BK182 ...);
  3. Calc, colunas CR..CT (N1..N3; livres nos dois produtos -- conferido): linhas 3..182 = corrida do lote do ponto
     daquele nivel (NA() sem ponto ou fora do periodo); 184/185 = bordas do eixo (sem corrida: 0..21, como o ADR-078);
     nomes xrunmin<i>/xrunmax<i>;
  4. graficos do Painel: no grafico do nivel i, o X de TODAS as series troca Calc!E (posicao) pela coluna do nivel --
     so a referencia de X (tipo, cor, legenda e Y ficam como estao);
  5. FUMACA no analito em tela: em cada nivel os X plotados sao inteiros, crescentes e sem buraco; devolve a protecao.
O RUN (yymmddkk, a corrida do DIA) nao muda: continua a chave do Westgard entre niveis e da dica do grafico.
"""
import math
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

COL_ENG = {1: 'BK', 2: 'BL', 3: 'BM'}          # Eng_Saida (mEstatistica.ENG_COL_CORL0 = 63 = BK)
COL_CALC = {1: 'CR', 2: 'CS', 3: 'CT'}         # Calc
R0, R1, RX0, RX1 = 3, 182, 184, 185
MUDANCAS = []


def log(*a):
    print(*a, flush=True)


def norm(f):
    return str(f or '').replace('_xlfn.', '').replace(' ', '').upper()


def setf(rng, f, oque):
    if norm(rng.Formula) != norm(f):
        rng.Formula = f
        MUDANCAS.append(oque)


def nome(wb, n, ref, oque):
    try:
        atual = str(wb.Names(n).RefersTo)
    except Exception:
        atual = None
    if norm(atual) != norm(ref):
        if atual is not None:
            wb.Names(n).RefersTo = ref
        else:
            wb.Names.Add(n, ref)
        MUDANCAS.append(oque)


def conferir_livres(wb, nlv):
    """As colunas novas tem de estar vazias ou ja serem do ADR-080 (nada sobrescrito em silencio)."""
    e, c = wb.Worksheets('Eng_Saida'), wb.Worksheets('Calc')
    for i in range(1, nlv + 1):
        k = COL_ENG[i]
        for row in e.Range(f'{k}{R0}:{k}{R1}').Value:
            v = row[0]
            if v not in (None, '') and not (isinstance(v, float) and v == int(v) and v >= 1):
                raise SystemExit(f'Eng_Saida {k}: conteudo que nao e corrida do lote ({v!r}) -- nada foi alterado')
        k = COL_CALC[i]
        for r in [1, 2] + list(range(R0, R1 + 1, 30)) + [R1, 183, RX0, RX1]:
            f = str(c.Range(f'{k}{r}').Formula or '')
            if f and 'ENGCORN' not in norm(f) and 'AGGREGATE' not in norm(f) and not f.startswith('N') \
                    and not f.startswith('CORRIDA'):
                raise SystemExit(f'Calc!{k}{r} ocupada ({f[:60]}) -- nada foi alterado')


def calc(wb, nlv):
    c = wb.Worksheets('Calc')
    for i in range(1, nlv + 1):
        k = COL_CALC[i]
        nome(wb, f'engCorN{i}', f"=Eng_Saida!${COL_ENG[i]}${R0}:${COL_ENG[i]}${R1}", f'nome engCorN{i}')
        if str(c.Range(f'{k}2').Value or '') != f'N{i} corrida lote':
            c.Range(f'{k}2').Value = f'N{i} corrida lote'
            MUDANCAS.append(f'Calc {k}2')
        novo = [f'=IF(OR($B{r}="",$D{r}<>1),NA(),IF(ISNUMBER(INDEX(engCorN{i},$A{r})),INDEX(engCorN{i},$A{r}),NA()))'
                for r in range(R0, R1 + 1)]
        rg = c.Range(f'{k}{R0}:{k}{R1}')
        atual = [norm(x[0]) for x in rg.Formula]
        if atual != [norm(x) for x in novo]:
            rg.Formula = tuple((x,) for x in novo)
            MUDANCAS.append(f'Calc {k}{R0}:{k}{R1}')
        if str(c.Range(f'{k}183').Value or '') != f'CORRIDA DO LOTE N{i} (ADR-080): bordas do eixo':
            c.Range(f'{k}183').Value = f'CORRIDA DO LOTE N{i} (ADR-080): bordas do eixo'
            MUDANCAS.append(f'Calc {k}183')
        faixa = f'${k}${R0}:${k}${R1}'
        setf(c.Range(f'{k}{RX0}'), f'=IF(COUNT({faixa})=0,0,AGGREGATE(5,6,{faixa})-0.6)', f'Calc {k}{RX0}')
        setf(c.Range(f'{k}{RX1}'), f'=IF(COUNT({faixa})=0,21,AGGREGATE(4,6,{faixa})+2.2)', f'Calc {k}{RX1}')
        nome(wb, f'xrunmin{i}', f'=Calc!${k}${RX0}', f'nome xrunmin{i}')
        nome(wb, f'xrunmax{i}', f'=Calc!${k}${RX1}', f'nome xrunmax{i}')


RX_X = re.compile(r'^=SERIES\((.*?),(Calc!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+),(Calc!\$([A-Z]+)\$\d+:\$[A-Z]+\$\d+),(\d+)\)$')


def graficos(wb, nlv):
    p, c = wb.Worksheets('Painel'), wb.Worksheets('Calc')
    cab = [str(v or '').strip() for v in c.Range('A2:DZ2').Value[0]]
    for i in range(1, nlv + 1):
        ch = p.ChartObjects(i).Chart
        k = COL_CALC[i]
        for s_i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(s_i)
            f = str(s.Formula)
            mt = RX_X.match(f)
            if not mt:
                raise SystemExit(f'grafico {i}, serie {s.Name}: formula fora do padrao: {f}')
            ycol = mt.group(4)
            hy = cab[ord_col(ycol) - 1] if ord_col(ycol) <= len(cab) else ''
            if not hy.startswith(f'N{i} '):
                raise SystemExit(f'grafico {i}, serie {s.Name}: Y em {ycol} ("{hy}"), nao e do nivel {i}')
            x = mt.group(2)
            if x in (f'Calc!$E${R0}:$E${R1}', f'Calc!${k}${R0}:${k}${R1}'):
                nx = f'Calc!${k}${R0}:${k}${R1}'
            elif x in (f'Calc!$E${RX0}:$E${RX1}', f'Calc!${k}${RX0}:${k}${RX1}'):
                nx = f'Calc!${k}${RX0}:${k}${RX1}'
            else:
                raise SystemExit(f'grafico {i}, serie {s.Name}: X em {x} (esperado Calc!E)')
            novo = f'=SERIES({mt.group(1)},{nx},{mt.group(3)},{mt.group(5)})'
            if novo != f:
                s.Formula = novo
                MUDANCAS.append(f'grafico {i} serie {s.Name}')


def ord_col(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def fumaca(wb, nlv):
    """Em cada nivel: os X dos pontos plotados (resultado ou X vermelho) sao inteiros, crescentes, sem buraco."""
    c = wb.Worksheets('Calc')
    cab = [str(v or '').strip() for v in c.Range('A2:DZ2').Value[0]]
    out = []
    for i in range(1, nlv + 1):
        xs = [r[0] for r in c.Range(f'{COL_CALC[i]}{R0}:{COL_CALC[i]}{R1}').Value]
        yl = [r[0] for r in c.Range(c.Cells(R0, cab.index(f'N{i} line') + 1), c.Cells(R1, cab.index(f'N{i} line') + 1)).Value]
        xx = [[r[j] for r in c.Range(c.Cells(R0, cab.index(f'N{i} x1') + 1),
                                     c.Cells(R1, cab.index(f'N{i} x3') + 1)).Value] for j in range(3)]
        pts = [xs[r] for r in range(len(xs)) if num(yl[r]) or any(num(xx[j][r]) for j in range(3))]
        if any(not num(x) or x != math.floor(x) for x in pts):
            raise SystemExit(f'fumaca N{i}: ponto plotado sem corrida do lote: {pts[:10]}')
        if any(b - a != 1 for a, b in zip(pts, pts[1:])):
            raise SystemExit(f'fumaca N{i}: corridas fora de sequencia: {pts[:20]}')
        out.append((i, len(pts), pts[0] if pts else None, pts[-1] if pts else None,
                    c.Range(f'{COL_CALC[i]}{RX0}').Value, c.Range(f'{COL_CALC[i]}{RX1}').Value))
    return out


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:
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
        ex.xl.Calculation = -4135
        nlv = 2 if produto.startswith('Bio') else 3
        conferir_livres(wb, nlv)
        import codigo_atual
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        calc(wb, nlv)
        log(f'2-3. nomes engCorN1..{nlv}; Calc {COL_CALC[1]}..{COL_CALC[nlv]}: corrida do lote por nivel e bordas do eixo '
            f'(xrunmin1..{nlv}/xrunmax1..{nlv})')
        graficos(wb, nlv)
        log('4. graficos: o X de cada nivel e a sua corrida do lote')
        ex.xl.Calculation = -4105
        ex.xl.EnableEvents = True
        ex.xl.Calculate()
        ex.esperar()
        for mac in ('mEstatistica.InvalidarCache', 'mEstatistica.RecalcularAnalitoAtual', 'mUI.AtualizarEixos'):
            ex.run(f"'{wb.Name}'!{mac}", teto=600)
        ex.xl.Calculate()
        ex.esperar()
        ex.xl.EnableEvents = False
        res = fumaca(wb, nlv)
        log('5. fumaca OK (nivel, pontos, 1a, ultima, borda esq., borda dir.): ' + str(res))
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS[:12]) + (' ...' if len(MUDANCAS) > 12 else '') if MUDANCAS else ''))
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
