# -*- coding: utf-8 -*-
"""suite_etapa1.py -- bateria de testes da Etapa 1 (multi-lote, ADR-049).

Uso:
  python suite_etapa1.py base  <etapa1.xlsm> <base.xlsm>   monta a base limpa de teste
  python suite_etapa1.py todos <base.xlsm> [T1 T2 ...]      roda os testes

A BASE e uma copia do arquivo JA COM A ETAPA 1, com os dados operacionais
zerados e o modulo mTesteEtapa1 instalado. Cada teste abre uma copia NOVA da
base (independencia: um teste nao herda estado do outro).

Tudo que o teste faz, ele faz pelo caminho do usuario:
  cadastrar lote ........ Configuracao!C26:C125
  lote em uso ........... Configuracao!C20 -> mLotes.TrocarLote
  lote em analise ....... Painel!H3 (ADR-054) -> mLotes.TrocarLoteAnalise
  media/DP .............. Analitos!E:H -> mLotes.ParametroEditado (o Worksheet_Change)
  resultados ............ aba Importar -> mImportar.ExecutarImportacao

E confere pelo que o usuario VE e pelo que o motor DECIDE:
  Calc (corridas, valores, z, linhas de limite), as SERIES DOS GRAFICOS lidas do
  proprio objeto grafico, a escala do eixo, o Painel (n, status, faixa O3),
  Eng_Saida, Eventos_Westgard, Estatistica, BI_Data e a reconciliacao Excel x BI.
"""
import datetime as dt
import json
import math
import os
import shutil
import sys
import time
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402

AQUI = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(os.environ['TMP'], 'qcw', 'suite')
os.makedirs(TMP, exist_ok=True)
NA = -2146826246          # #N/D devolvido pelo COM
TOL = 1e-7

RESULTADOS = []           # (teste, verificacao, ok, detalhe)


# ============================================================================
# utilidades
# ============================================================================
def num(v):
    return isinstance(v, float) or (isinstance(v, int) and not isinstance(v, bool) and v > -2146000000)


def norm(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip() if v is not None else ''


def igual(a, b, tol=TOL):
    return num(a) and num(b) and abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


class Falha(Exception):
    pass


def chk(teste, nome, ok, detalhe=''):
    ok = bool(ok)
    RESULTADOS.append((teste, nome, ok, str(detalhe)[:400]))
    marca = 'PASS' if ok else 'FAIL'
    print(f'   [{marca}] {nome}' + ('' if ok else f'  -> {str(detalhe)[:300]}'), flush=True)
    return ok


def dias_uteis(inicio, n):
    d, out = inicio, []
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


CALMO = [0.3, -0.6, 0.9, -0.2, 0.5, -1.1, 0.7, -0.4, 1.2, -0.8]


def serie_z(n, plantar=None):
    """z por corrida: sequencia alternada que NAO dispara regra, mais plantios."""
    z = [CALMO[k % len(CALMO)] for k in range(n)]
    for k, v in (plantar or {}).items():
        z[k] = v
    return z


# ============================================================================
# sessao: uma copia da base aberta no Excel
# ============================================================================
class Sessao:
    def __init__(self, base, nome):
        self.caminho = os.path.join(TMP, f'{nome}.xlsm')
        shutil.copyfile(base, self.caminho)
        self.ex = xlh.Excel()
        self.wb = self.ex.abrir(self.caminho)
        self.xl = self.ex.xl
        sh = {s.Name: s for s in self.wb.Worksheets}
        self.sh = sh
        self.cfg, self.pai, self.calc = sh['Configuração'], sh['Painel'], sh['Calc']
        self.ana, self.eng, self.est = sh['Analitos'], sh['Eng_Saida'], sh['Estatística']
        self.imp, self.ls = sh['Importar'], sh['LotesStore']
        self.nabas0 = self.wb.Worksheets.Count
        self.analitos = [self.ana.Cells(r, 1).Value for r in range(4, 44)]
        # mapa analito -> coluna da aba Importar (linha oculta de mapeamento)
        self.col_imp = {}
        c = 5
        while self.imp.Cells(4, c).Value not in (None, ''):
            nm = self.imp.Cells(3, c).Value
            if nm:
                self.col_imp[str(nm).strip()] = c
            c += 1

    def run(self, m, *a, teto=600):
        r = self.ex.run(m, *a, teto=teto)
        self.ex.esperar()
        return r

    def fechar(self):
        self.ex.fechar()

    # --- acoes do usuario ---------------------------------------------------
    def cadastrar(self, *lotes):
        for l in lotes:
            self.run('mTesteEtapa1.T_Cadastrar', str(l))

    def em_uso(self, lote):
        self.run('mTesteEtapa1.T_LoteEmUso', str(lote))

    def selecionar(self, lote):
        self.run('mTesteEtapa1.T_Selecionar', str(lote))

    def analito(self, nome):
        self.run('mTesteEtapa1.T_Analito', nome)

    def parametros(self, lote, analito, m1, s1, m2, s2):
        """Seleciona o lote no Painel e digita media/DP na aba Analitos."""
        self.selecionar(lote)
        self.run('mTesteEtapa1.T_Parametros', analito, m1, s1, m2, s2)

    def importar(self, linhas):
        """linhas: [(data, nivel, lote, {analito: valor})] -> aba Importar, lotes de 150."""
        for i in range(0, len(linhas), 150):
            bloco = linhas[i:i + 150]
            ncol = max(self.col_imp.values())
            area = []
            for (d, nv, lote, vals) in bloco:
                row = [None] * (ncol - 1)          # colunas B..ncol
                row[0] = d.strftime('%d/%m/%Y')
                row[1] = nv
                row[2] = str(lote)
                for an, v in vals.items():
                    row[self.col_imp[an] - 2] = str(v).replace('.', ',')
                area.append(row)
            self.imp.Range(self.imp.Cells(5, 2), self.imp.Cells(4 + len(area), ncol)).NumberFormat = '@'
            self.imp.Range(self.imp.Cells(5, 2), self.imp.Cells(4 + len(area), ncol)).Value = area
            r = self.run('mImportar.ExecutarImportacao', True, teto=900)
            if not str(r).startswith('OK'):
                erros = [self.imp.Cells(5 + k, ncol + 2).Value for k in range(5)]
                raise Falha(f'importacao recusada: {r} {erros}')

    def serie(self, lote, analito, datas, p, z1, z2):
        """Monta linhas de importacao: N1 = m1 + z*s1, N2 = m2 + z2*s2."""
        m1, s1, m2, s2 = p
        out = []
        for d, a, b in zip(datas, z1, z2):
            if a is not None:
                out.append((d, 1, lote, {analito: round(m1 + a * s1, 6)}))
            if b is not None:
                out.append((d, 2, lote, {analito: round(m2 + b * s2, 6)}))
        return out

    def filtro(self, de=None, ate=None):
        self.pai.Range('G3').Value = de if de else ''
        self.pai.Range('G4').Value = ate if ate else ''
        self.run('mEstatistica.RecalcularAnalitoAtual')

    # --- leitura ------------------------------------------------------------
    def db(self):
        ws = self.sh['DB_Resultados']
        ult = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if ult < 4:
            return []
        v = ws.Range(f'A4:G{ult}').Value
        out = []
        for r in v:
            if r[0] is None:
                continue
            lote = str(r[3])[3:-2]
            out.append({'run': int(r[0]), 'data': r[1], 'nivel': int(r[2]), 'lote': lote,
                        'analito': r[4], 'valor': r[5], 'status': r[6]})
        return out

    def ler(self):
        self.xl.Calculate()
        self.ex.esperar()
        c = self.calc
        L = {}
        L['par'] = c.Range('AX1:BA1').Value[0]
        L['BH1'] = c.Range('BH1').Value
        L['calc'] = c.Range('A3:AW182').Value
        L['eng1'] = self.eng.Range('A1:Q1').Value[0]
        # ADR-054: o Painel mudou de desenho -- lote em H3, faixa de status
        # abaixo do grafico (A39) e a coluna Status do nivel em J (indice 9).
        L['painel'] = self.pai.Range('A7:U8').Value
        L['O3'] = self.pai.Range('A39').Value
        L['E3'] = self.pai.Range('H3').Value
        L['E3txt'] = self.pai.Range('H3').Text
        L['C4'] = self.pai.Range('C4').Value
        L['loteParam'] = self.cfg.Range('G2').Value
        L['H4'] = self.est.Range('H4').Value
        L['de'] = self.pai.Range('G3').Value
        L['ate'] = self.pai.Range('G4').Value
        L['eventos'] = self.run('mTesteEtapa1.T_Eventos')
        self.run('AtualizarEixos')
        graf = []
        for co in self.pai.ChartObjects():
            ch = co.Chart
            series = {}
            for s in ch.SeriesCollection():
                series.setdefault(s.Name, []).append(list(s.Values))
            ax = ch.Axes(2)
            graf.append({'series': series, 'ymin': ax.MinimumScale, 'ymax': ax.MaximumScale})
        L['graf'] = graf
        return L

    def estat_linha(self, analito, nivel):
        idx = self.analitos.index(analito)
        r = 14 + idx * 2 + (nivel - 1)
        return self.est.Range(f'A{r}:F{r}').Value[0]


# ============================================================================
# conferencia central: "tudo que o usuario ve e o motor decide e DESTE lote"
# ============================================================================
COLS = {  # indice 0-based dentro de Calc!A:AW
    'B': 1, 'C': 2, 'D': 3, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'P': 15,
    'Q': 16, 'T': 19, 'W': 22, 'AB': 27, 'AC': 28, 'AF': 31, 'AG': 32, 'AL': 37, 'AM': 38, 'AP': 41, 'AS': 44}


def conferir(T, s, lote, analito, p, L=None, periodo=None, rotulo=''):
    """p = (m1, s1, m2, s2) esperados, ou None para lote SEM parametro."""
    L = L or s.ler()
    pre = f'{rotulo}lote {lote}: '
    lote = str(lote)
    db = [r for r in s.db() if r['lote'] == lote and r['analito'] == analito and r['status'] == 'Ativo']
    dia = lambda d: d.replace(tzinfo=None).date() if hasattr(d, 'date') else d
    de = dia(L['de']) if hasattr(L.get('de'), 'year') else None
    ate = dia(L['ate']) if hasattr(L.get('ate'), 'year') else None
    corridas = sorted({(dia(r['data']), r['run']) for r in db})
    if ate:
        corridas = [c for c in corridas if c[0] <= ate]
    runs_lote = {rn for _, rn in corridas[-180:]}                 # janela do motor
    runs_todas = {r['run'] for r in db}
    runs_per = {rn for d, rn in corridas if (not de or d >= de) and (not ate or d <= ate)}
    val = {(r['run'], r['nivel']): r['valor'] for r in db}
    ok_all = True

    ok_all &= chk(T, pre + 'Painel H3 / Estatistica H4 / Eng_Saida E1 = lote',
                  norm(L['E3']) == lote and norm(L['H4']) == lote and norm(L['eng1'][4]) == lote,
                  f"H3={L['E3']} H4={L['H4']} eng={L['eng1'][4]}")
    ok_all &= chk(T, pre + 'tela Analitos (loteParam) = lote', norm(L['loteParam']) == lote, L['loteParam'])

    if p is None:
        ok_all &= chk(T, pre + 'SEM parametro -> AX1:BA1 vazios', all(x in (None, '') for x in L['par']), L['par'])
    else:
        ok_all &= chk(T, pre + f'media/DP recuperados {p}',
                      all(igual(a, b) for a, b in zip(L['par'], p)), L['par'])

    linhas = [r for r in L['calc'] if r[COLS['B']] not in (None, '')]
    runs_calc = [int(r[COLS['B']]) for r in linhas]
    ok_all &= chk(T, pre + f'corridas no grafico = corridas do lote na janela ({len(runs_lote)} de {len(runs_todas)})',
                  set(runs_calc) == runs_lote and len(runs_calc) == len(set(runs_calc)),
                  f'calc={len(runs_calc)} lote={len(runs_lote)} estranhas={sorted(set(runs_calc) - runs_lote)[:5]}')
    datas = [r[COLS['C']] for r in linhas]
    ok_all &= chk(T, pre + 'ordem cronologica no eixo', datas == sorted(datas), '')

    erros_v, erros_z, erros_lim, erros_ver, zmax = [], [], [], [], 0.0
    for r in linhas:
        run = int(r[COLS['B']])
        for nv, cv, cz, cst, cm3, cmed, cp3, cflag in ((1, 'F', 'G', 'P', 'Q', 'T', 'W', 'K'),
                                                         (2, 'AB', 'AC', 'AL', 'AM', 'AP', 'AS', 'AG')):
            v = r[COLS[cv]]
            if (run, nv) in val:
                if not igual(v, val[(run, nv)]):
                    erros_v.append((run, nv, v, val[(run, nv)]))
            elif v not in (None, ''):
                erros_v.append((run, nv, v, 'nao existe no lote'))
            if p is None or not num(v):
                if p is None and r[COLS[cz]] not in (None, ''):
                    erros_z.append((run, nv, 'z sem parametro', r[COLS[cz]]))
                if p is None and v not in (None, '') and r[COLS[cst]] not in ('SEM PARAMETROS',):
                    erros_ver.append((run, nv, 'veredicto sem parametro', r[COLS[cst]]))
                continue
            m, sd = (p[0], p[1]) if nv == 1 else (p[2], p[3])
            z = (float(v) - m) / sd
            zmax = max(zmax, abs(z))
            if not igual(r[COLS[cz]], z, 1e-6):
                erros_z.append((run, nv, r[COLS[cz]], z))
            if r[COLS['D']] == 1:
                for col, esp in ((cm3, m - 3 * sd), (cmed, m), (cp3, m + 3 * sd)):
                    if not igual(r[COLS[col]], esp):
                        erros_lim.append((run, nv, col, r[COLS[col]], esp))
            # 1_3s e regra individual: tem de bater com |z| > 3
            if int(r[COLS[cflag]] or 0) != (1 if abs(z) > 3 else 0):
                erros_ver.append((run, nv, '1_3s', r[COLS[cflag]], round(z, 3)))
            if r[COLS[cst]] not in ('OK', 'REJEITADO'):
                erros_ver.append((run, nv, 'veredicto', r[COLS[cst]]))
    ok_all &= chk(T, pre + 'valores plotados = valores do lote no banco', not erros_v, erros_v[:3])
    ok_all &= chk(T, pre + 'z calculado com a media/DP DESTE lote', not erros_z, erros_z[:3])
    ok_all &= chk(T, pre + 'linhas -3s/media/+3s = parametros do lote', not erros_lim, erros_lim[:3])
    ok_all &= chk(T, pre + 'veredicto do motor coerente com o lote (1_3s = |z|>3)', not erros_ver, erros_ver[:3])
    if p is not None:
        ok_all &= chk(T, pre + f'nenhum z absurdo (referencia cruzada) max|z|={zmax:.2f}', zmax < 10, zmax)

    # --- o GRAFICO, lido do objeto grafico --------------------------------
    for gi, nv in ((0, 1), (1, 2)):
        g = L['graf'][gi]
        med = [x for x in g['series']['Média'][0] if num(x)]
        m3 = [x for x in g['series']['-3s'][0] if num(x)]
        p3 = [x for x in g['series']['+3s'][0] if num(x)]
        okp = [x for x in g['series']['OK'][0] if num(x)]
        vio = [x for x in g['series']['Violação'][0] if num(x)]
        if p is None:
            ok_all &= chk(T, pre + f'grafico N{nv}: sem linhas de limite e sem ponto classificado',
                          not med and not m3 and not p3 and not okp and not vio,
                          f'med={len(med)} ok={len(okp)} vio={len(vio)}')
            continue
        m, sd = (p[0], p[1]) if nv == 1 else (p[2], p[3])
        ok_all &= chk(T, pre + f'grafico N{nv}: series Media/-3s/+3s = {m} / {m - 3 * sd} / {m + 3 * sd}',
                      med and all(igual(x, m) for x in med) and all(igual(x, m - 3 * sd) for x in m3)
                      and all(igual(x, m + 3 * sd) for x in p3),
                      f'med={set(round(x, 6) for x in med)} m3={set(round(x, 6) for x in m3)}')
        vis = [float(r[COLS['F' if nv == 1 else 'AB']]) for r in linhas
               if r[COLS['D']] == 1 and num(r[COLS['F' if nv == 1 else 'AB']])]
        lo = min([m - 3.3 * sd] + vis)
        hi = max([m + 3.3 * sd] + vis)
        ok_all &= chk(T, pre + f'grafico N{nv}: eixo Y cobre os limites do lote [{lo:.4g}; {hi:.4g}]',
                      g['ymin'] <= m - 3 * sd and g['ymax'] >= m + 3 * sd
                      and g['ymin'] >= lo - 0.05 * sd - 1e-9 and g['ymax'] <= hi + 0.05 * sd + 1e-9,
                      f"ymin={g['ymin']} ymax={g['ymax']}")

    # --- Painel, eventos, estatistica -------------------------------------
    n1 = sum(1 for r in linhas if r[COLS['D']] == 1 and num(r[COLS['F']]))
    n1_esp = sum(1 for rn in runs_per & runs_lote if (rn, 1) in val)
    ok_all &= chk(T, pre + f'Painel n N1 = {n1_esp} (corridas do lote no periodo)',
                  L['painel'][0][1] == n1 == n1_esp, (L['painel'][0][1], n1, n1_esp))
    st = L['painel'][0][9]
    if p is None:
        ok_all &= chk(T, pre + 'Painel J7 avisa SEM MEDIA/DP (nao diz "OK")',
                      str(st).startswith('SEM MÉDIA/DP'), st)
        ok_all &= chk(T, pre + 'faixa avisa lote sem media/DP', str(L['O3']).startswith('⛔ LOTE ' + lote), L['O3'])
    else:
        ok_all &= chk(T, pre + 'faixa mostra o lote e o alvo', str(L['O3']).startswith('Lote ' + lote + ' · alvo'),
                      L['O3'])
    ev = str(L['eventos']).split('|')
    # o campo Niveis de evento multinivel e "1,2": so o pedaco com ':' abre um evento
    ev_runs = [int(float(x.split(':')[0])) for x in ev[2].split(',') if ':' in x] if len(ev) > 2 else []
    ok_all &= chk(T, pre + 'Eventos_Westgard: lote do cabecalho e todas as corridas sao DESTE lote',
                  norm(ev[0]) == lote and all(rn in runs_todas for rn in ev_runs),
                  f'cab={ev[0]} fora={[x for x in ev_runs if x not in runs_todas][:5]}')
    if p is None:
        ok_all &= chk(T, pre + 'sem parametro -> nenhum evento de Westgard', len(ev_runs) == 0, ev[1] if len(ev) > 1 else ev)
    e1 = s.estat_linha(analito, 1)
    v1 = [val[(rn, 1)] for rn in runs_per if (rn, 1) in val]
    ok_all &= chk(T, pre + f'Estatistica n/media N1 = deste lote (n={len(v1)})',
                  e1[2] == len(v1) and (not v1 or igual(e1[3], sum(v1) / len(v1), 1e-9)),
                  f'estat n={e1[2]} media={e1[3]}')
    return ok_all, L


def conferir_store(T, s, esperado):
    """esperado: {(lote, analito): (m1,s1,m2,s2)} -- a store inteira tem de bater."""
    ult = s.ls.Cells(s.ls.Rows.Count, 1).End(-4162).Row
    v = s.ls.Range(f'A2:P{max(ult, 2)}').Value
    achado = {}
    for r in v:
        if r[15]:
            achado[str(r[15])] = tuple(r[2:6])
    erros = []
    for (lote, an), p in esperado.items():
        k = f'{str(lote).upper()}|{an.upper()}'
        if k not in achado:
            erros.append((k, 'ausente'))
        elif not all(igual(a, b) for a, b in zip(achado[k], p)):
            erros.append((k, achado[k], p))
    extras = set(achado) - {f'{str(l).upper()}|{a.upper()}' for (l, a) in esperado}
    chk(T, f'LotesStore: {len(esperado)} chaves exatas, nenhuma alterada/extra', not erros and not extras,
        f'{erros[:3]} extras={list(extras)[:3]}')
    r = s.run('mLotes.ConferirLotesStore')
    chk(T, f'LotesStore integra (sem chave duplicada/orfa): {r}', str(r).startswith('ok'), r)


# ============================================================================
# TESTES
# ============================================================================
A = ('9001', (100.0, 5.0, 50.0, 2.0))
B = ('9002', (200.0, 10.0, 150.0, 6.0))
C = ('9003', (1000.0, 40.0, 500.0, 20.0))
AN = 'Lactato'
AN2 = 'Glicose'


def preparar(s, lotes_params, n=25, inicio=dt.date(2027, 1, 4), mesmo_periodo=False, plantar=None):
    """Cadastra, digita parametros e importa uma serie por lote (meses diferentes)."""
    s.cadastrar(*[l for l, _ in lotes_params])
    s.em_uso(lotes_params[0][0])
    for i, (l, p) in enumerate(lotes_params):
        if p is not None:
            s.parametros(l, AN, *p)
    linhas = []
    for i, (l, p) in enumerate(lotes_params):
        ini = inicio if mesmo_periodo else inicio + dt.timedelta(days=45 * i)
        datas = dias_uteis(ini, n)
        pp = p or (100.0, 5.0, 50.0, 2.0)
        z1 = serie_z(n, (plantar or {}).get(l))
        z2 = [-x * 0.5 for x in serie_z(n)]
        linhas += s.serie(l, AN, datas, pp, z1, z2)
    s.importar(linhas)


def T1(base):
    T = 'T1 um lote'
    s = Sessao(base, 'T1')
    try:
        preparar(s, [A], plantar={'9001': {10: 3.5}})
        s.selecionar(A[0]); s.analito(AN)
        ok, L = conferir(T, s, A[0], AN, A[1])
        chk(T, 'lista de lotes do Painel tem 1 lote', s.wb.Names('lstLotes').RefersToRange.Rows.Count == 1, s.wb.Names('lstLotes').RefersToRange.Rows.Count)
        chk(T, 'H3 exibe "Lote 9001"', L['E3txt'] == 'Lote 9001', L['E3txt'])
        chk(T, 'tela Analitos mostra 100 / 5 / 50 / 2',
            all(igual(a, b) for a, b in zip(s.ana.Range('E4:H4').Value[0], A[1])), s.ana.Range('E4:H4').Value)
        k = [r for r in L['calc'] if r[COLS['B']] not in (None, '') and r[COLS['K']] == 1]
        chk(T, 'violacao plantada (z=+3,5) detectada como 1_3s em exatamente 1 corrida', len(k) == 1, len(k))
        chk(T, 'Painel J7 = REJEITADO', str(L['painel'][0][9]) == 'REJEITADO', L['painel'][0][9])
        conferir_store(T, s, {(A[0], AN): A[1]})
    finally:
        s.fechar()


def T2(base):
    T = 'T2 dois lotes'
    s = Sessao(base, 'T2')
    try:
        preparar(s, [A, B], plantar={'9002': {5: 3.5}})
        s.analito(AN)
        s.selecionar(A[0]); conferir(T, s, A[0], AN, A[1], rotulo='[A] ')
        s.selecionar(B[0]); ok, L = conferir(T, s, B[0], AN, B[1], rotulo='[B] ')
        k = [r for r in L['calc'] if r[COLS['B']] not in (None, '') and r[COLS['K']] == 1]
        chk(T, '[B] valor 235 (z=+3,5 no lote B) e 1_3s so no lote B', len(k) == 1 and igual(k[0][COLS['F']], 235.0), k[:1])
        s.selecionar(A[0]); ok, L = conferir(T, s, A[0], AN, A[1], rotulo='[A de novo] ')
        chk(T, '[A] 235 nao aparece no grafico do lote A',
            not any(igual(r[COLS['F']], 235.0) for r in L['calc']), '')
        conferir_store(T, s, {(A[0], AN): A[1], (B[0], AN): B[1]})
    finally:
        s.fechar()


def T3(base):
    T = 'T3 troca repetida'
    s = Sessao(base, 'T3')
    try:
        preparar(s, [A, B, C])
        s.analito(AN)
        loja = {(A[0], AN): A[1], (B[0], AN): B[1], (C[0], AN): C[1]}
        ref = dict([A, B, C])
        for i, l in enumerate(['9001', '9002', '9003', '9001', '9003', '9002']):
            s.selecionar(l)
            conferir(T, s, l, AN, ref[l], rotulo=f'[{i + 1}] ')
            chk(T, f'[{i + 1}] tela Analitos = parametros de {l}',
                all(igual(a, b) for a, b in zip(s.ana.Range('E4:H4').Value[0], ref[l])), s.ana.Range('E4:H4').Value)
        conferir_store(T, s, loja)
        # troca com o spinner de analito no meio: Glicose (sem parametro em nenhum lote) e volta
        s.analito(AN2)
        L = s.ler()
        chk(T, 'Glicose (nao configurada) -> sem media/DP, sem veredicto', all(x in (None, '') for x in L['par']), L['par'])
        s.analito(AN)
        conferir(T, s, '9002', AN, B[1], rotulo='[volta ao Lactato] ')
    finally:
        s.fechar()


def T4(base):
    T = 'T4 limites muito diferentes'
    s = Sessao(base, 'T4')
    D = ('7001', (5.0, 0.1, 2.0, 0.05))
    E = ('7002', (5000.0, 250.0, 3000.0, 150.0))
    try:
        preparar(s, [D, E])
        s.analito(AN)
        for l, p in (D, E, D, E):
            s.selecionar(l)
            ok, L = conferir(T, s, l, AN, p, rotulo=f'[{l}] ')
            g = L['graf'][0]
            chk(T, f'[{l}] escala do eixo N1 na ordem do lote ({g["ymin"]:.4g} .. {g["ymax"]:.4g})',
                g['ymax'] - g['ymin'] < 20 * p[1], (g['ymin'], g['ymax']))
    finally:
        s.fechar()


def T5(base):
    T = 'T5 resultados de lotes diferentes'
    s = Sessao(base, 'T5')
    try:
        # MESMAS datas nos dois lotes: corrida paralela de troca de lote
        preparar(s, [A, B], mesmo_periodo=True)
        # segundo analito com parametros proprios em cada lote
        s.parametros(A[0], AN2, 90.0, 3.0, 250.0, 8.0)
        s.parametros(B[0], AN2, 95.0, 4.0, 260.0, 9.0)
        datas = dias_uteis(dt.date(2027, 1, 4), 25)
        s.importar(s.serie(A[0], AN2, datas, (90.0, 3.0, 250.0, 8.0), serie_z(25), serie_z(25))
                   + s.serie(B[0], AN2, datas, (95.0, 4.0, 260.0, 9.0), serie_z(25), serie_z(25)))
        db = s.db()
        por_chave = {}
        for r in db:
            por_chave.setdefault((r['data'], r['lote']), set()).add(r['run'])
        chk(T, 'cada (data, lote) tem UM RUN', all(len(v) == 1 for v in por_chave.values()), '')
        runs_a = {r['run'] for r in db if r['lote'] == A[0]}
        runs_b = {r['run'] for r in db if r['lote'] == B[0]}
        chk(T, 'mesmo dia, lotes diferentes -> RUNs diferentes (nenhum RUN compartilhado)',
            not (runs_a & runs_b), sorted(runs_a & runs_b)[:5])
        chk(T, 'lote de cada resultado vem do codigo QC-<lote><nivel> gravado no banco',
            all(r['lote'] in (A[0], B[0]) for r in db), {r['lote'] for r in db})
        for l, p, an in ((A[0], A[1], AN), (B[0], B[1], AN), (A[0], (90.0, 3.0, 250.0, 8.0), AN2),
                         (B[0], (95.0, 4.0, 260.0, 9.0), AN2)):
            s.selecionar(l); s.analito(an)
            conferir(T, s, l, an, p, rotulo=f'[{an}] ')
        # BI: tabela fato por resultado -- cada linha com o alvo do SEU lote
        s.run('mBI.AtualizarBIData', teto=900)
        bi = s.sh['BI_Data']
        ult = bi.Cells(bi.Rows.Count, 1).End(-4162).Row
        cab = bi.Range(bi.Cells(1, 1), bi.Cells(1, 40)).Value[0]
        v = bi.Range(bi.Cells(2, 1), bi.Cells(ult, 40)).Value
        esp = {(A[0], AN.upper()): A[1], (B[0], AN.upper()): B[1],
               (A[0], AN2.upper()): (90.0, 3.0, 250.0, 8.0), (B[0], AN2.upper()): (95.0, 4.0, 260.0, 9.0)}
        erros, n = [], 0
        for r in v:
            idr = str(r[0]).split('|')
            if len(idr) < 4:
                continue
            lote, nv, an = idr[0], int(idr[2]), idr[3]
            if (lote, an) not in esp:
                continue
            n += 1
            p = esp[(lote, an)]
            m, sd = (p[0], p[1]) if nv == 1 else (p[2], p[3])
            if not (igual(r[18], m) and igual(r[19], sd) and igual(r[20], (float(r[15]) - m) / sd, 1e-6)):
                erros.append((r[0], r[18], r[19], r[20]))
        chk(T, f'BI_Data: {n} linhas, alvo e z de cada uma = do SEU lote', n == 4 * 25 * 2 and not erros,
            f'n={n} {erros[:3]}')
        for l, an in ((A[0], AN), (B[0], AN2)):
            s.selecionar(l); s.analito(an)
            r = s.run('mBI.ReconciliarComCalc')
            comp, div = str(r).split('|')[:2]
            chk(T, f'sentinela Excel x BI ({l} {an}): {r}', int(comp) > 0 and int(div) == 0, r)
    finally:
        s.fechar()


def T6(base):
    T = 'T6 lote antigo + lote novo'
    s = Sessao(base, 'T6')
    L1 = ('2601', (4.70, 0.12, 1.20, 0.05))
    L2 = ('2602', (4.90, 0.15, 1.35, 0.06))
    try:
        s.cadastrar(L1[0])
        s.em_uso(L1[0])
        s.parametros(L1[0], AN, *L1[1])
        d1 = dias_uteis(dt.date(2026, 1, 5), 40)
        s.importar(s.serie(L1[0], AN, d1, L1[1], serie_z(40, {20: 3.4, 30: -3.2}), [-x * 0.5 for x in serie_z(40)]))
        s.analito(AN); s.selecionar(L1[0])
        _, antes = conferir(T, s, L1[0], AN, L1[1], rotulo='[antes] ')
        store_antes = s.ls.Range('A2:P200').Value
        db_antes = [r for r in s.db() if r['lote'] == L1[0]]
        s.run('mBI.AtualizarBIData', teto=900)
        bi_antes = [r for r in s.sh['BI_Data'].Range('A2:AK400').Value if str(r[0]).startswith(L1[0] + '|')]

        # --- chega o lote novo: cadastro, vira lote em uso, parametros, resultados
        s.cadastrar(L2[0])
        s.em_uso(L2[0])
        L = s.ler()
        chk(T, 'lote novo NAO herda media/DP do antigo (tela Analitos vazia)',
            all(x in (None, '') for x in s.ana.Range('E4:H4').Value[0]), s.ana.Range('E4:H4').Value)
        chk(T, 'lote novo, antes de configurar: Painel avisa SEM MEDIA/DP', str(L['O3']).startswith('⛔ LOTE 2602'), L['O3'])
        chk(T, 'trocar o lote em uso levou o Painel junto (H3 = 2602)', norm(L['E3']) == '2602', L['E3'])
        s.parametros(L2[0], AN, *L2[1])
        d2 = dias_uteis(dt.date(2026, 3, 2), 30)
        s.importar(s.serie(L2[0], AN, d2, L2[1], serie_z(30), [-x * 0.5 for x in serie_z(30)]))
        s.selecionar(L2[0]); conferir(T, s, L2[0], AN, L2[1], rotulo='[novo] ')

        # --- volta ao antigo: tudo identico ao de antes
        s.selecionar(L1[0])
        _, depois = conferir(T, s, L1[0], AN, L1[1], rotulo='[antigo depois] ')
        chk(T, 'lote antigo: Calc (corridas, valores, z, limites, veredictos) identico ao de antes',
            antes['calc'] == depois['calc'], '')
        chk(T, 'lote antigo: series dos graficos identicas', antes['graf'] == depois['graf'], '')
        chk(T, 'lote antigo: Painel (n, media, DP, CV, violacoes) identico', antes['painel'] == depois['painel'], '')
        store_depois = s.ls.Range('A2:P200').Value
        linhas_l1 = lambda st: [r for r in st if norm(r[0]) == L1[0]]
        chk(T, 'LotesStore: linhas do lote antigo intactas', linhas_l1(store_antes) == linhas_l1(store_depois), '')
        chk(T, 'banco: resultados do lote antigo intactos',
            db_antes == [r for r in s.db() if r['lote'] == L1[0]], '')
        s.run('mBI.AtualizarBIData', teto=900)
        bi_depois = [r for r in s.sh['BI_Data'].Range('A2:AK400').Value if str(r[0]).startswith(L1[0] + '|')]
        comuns = lambda rows: [tuple(r[:37]) for r in rows]
        chk(T, 'BI_Data: linhas do lote antigo identicas (alvo, z, Westgard)',
            comuns(bi_antes) == comuns(bi_depois), f'{len(bi_antes)} x {len(bi_depois)}')
        conferir_store(T, s, {(L1[0], AN): L1[1], (L2[0], AN): L2[1]})
    finally:
        s.fechar()


def T7(base):
    T = 'T7 dez lotes'
    s = Sessao(base, 'T7')
    lotes = [(f'00{i}' if i < 10 else '010', (100.0 * i, 5.0 * i, 50.0 * i, 2.0 * i)) for i in range(1, 11)]
    lotes[6] = ('L-007A', lotes[6][1])          # codigo alfanumerico no meio
    try:
        preparar(s, lotes, n=15, inicio=dt.date(2025, 1, 6))
        s.analito(AN)
        chk(T, 'lista de lotes do Painel tem os 10', s.wb.Names('lstLotes').RefersToRange.Rows.Count == 10, s.wb.Names('lstLotes').RefersToRange.Rows.Count)
        for l, p in lotes:
            s.selecionar(l)
            conferir(T, s, l, AN, p, rotulo=f'[{l}] ')
        conferir_store(T, s, {(l, AN): p for l, p in lotes})
        chk(T, 'nenhuma estrutura nova por lote (mesmas abas, mesmos graficos)',
            s.wb.Worksheets.Count == s.nabas0 and s.pai.ChartObjects().Count == 2,
            (s.wb.Worksheets.Count, s.pai.ChartObjects().Count))
    finally:
        s.fechar()


def T8(base):
    T = 'T8 lote sem configuracao'
    s = Sessao(base, 'T8')
    try:
        X = ('9099', None)
        preparar(s, [A, X])
        s.analito(AN)
        s.selecionar('9099')
        conferir(T, s, '9099', AN, None, rotulo='[nada] ')
        # parcial: so N1
        s.run('mTesteEtapa1.T_Parametros', AN, 120.0, 6.0, '', '')
        L = s.ler()
        chk(T, '[so N1] N1 avaliado, N2 nao', igual(L['par'][0], 120.0) and L['par'][2] in (None, ''), L['par'])
        chk(T, '[so N1] J8 avisa SEM MEDIA/DP no N2', str(L['painel'][1][9]).startswith('SEM MÉDIA/DP'), L['painel'][1][9])
        n2 = [r[COLS['AL']] for r in L['calc'] if num(r[COLS['AB']])]
        chk(T, '[so N1] veredicto N2 = SEM PARAMETROS em todas as corridas', n2 and all(x == 'SEM PARAMETROS' for x in n2),
            set(n2))
        for rot, m, sd in (('DP = 0', 120.0, 0.0), ('DP negativo', 120.0, -6.0), ('media texto', 'abc', 6.0)):
            s.run('mTesteEtapa1.T_Parametros', AN, m, sd, '', '')
            L = s.ler()
            chk(T, f'[{rot}] tratado como SEM parametro (nao plota limite, nao avalia)',
                L['par'][0] in (None, '') and str(L['painel'][0][9]).startswith('SEM MÉDIA/DP'), (L['par'], L['painel'][0][9]))
        # o lote configurado continua intacto
        s.selecionar(A[0]); conferir(T, s, A[0], AN, A[1], rotulo='[A intacto] ')
        # lote cadastrado, sem resultado e sem parametro
        s.cadastrar('9098'); s.selecionar('9098')
        L = s.ler()
        chk(T, '[sem dados] grafico vazio e aviso na faixa', not [r for r in L['calc'] if r[1] not in (None, '')]
            and 'Nenhuma corrida deste lote' in str(L['O3']), L['O3'])
    finally:
        s.fechar()


def T9(base):
    T = 'T9 alteracao de parametros'
    s = Sessao(base, 'T9')
    try:
        preparar(s, [A, B])
        s.analito(AN)
        s.selecionar(A[0])
        _, L0 = conferir(T, s, A[0], AN, A[1], rotulo='[antes] ')
        aud = s.sh['Audit_Log']
        u0 = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        novo = (102.0, 5.0, 50.0, 2.5)
        s.run('mTesteEtapa1.T_Parametros', AN, *novo)
        _, L1 = conferir(T, s, A[0], AN, novo, rotulo='[depois] ')
        chk(T, 'efeito imediato: z de TODO o historico do lote refeito com 102/5 (nenhum z antigo sobra)',
            L0['calc'] != L1['calc'], '')
        u1 = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        linhas = aud.Range(f'A{u0 + 1}:X{u1}').Value if u1 > u0 else []
        acoes = [(r[6], r[12], r[13], r[14], r[16], r[17], r[22]) for r in linhas]
        chk(T, f'Audit_Log: {len(acoes)} registros PARAMETRO_LOTE_ALTERADO com antes/depois',
            len(acoes) == 2 and all(a[0] == 'PARAMETRO_LOTE_ALTERADO' and norm(a[1]) == A[0] for a in acoes)
            and {a[6] for a in acoes} == {'Media N1', 'DP N2'}, acoes)
        s.selecionar(B[0]); conferir(T, s, B[0], AN, B[1], rotulo='[B nao mudou] ')
        conferir_store(T, s, {(A[0], AN): novo, (B[0], AN): B[1]})
        # desfazer volta exatamente ao estado anterior
        s.selecionar(A[0])
        s.run('mTesteEtapa1.T_Parametros', AN, *A[1])
        _, L2 = conferir(T, s, A[0], AN, A[1], rotulo='[desfeito] ')
        chk(T, 'desfazer a alteracao reproduz o grafico original', L2['calc'] == L0['calc'], '')
    finally:
        s.fechar()


def T_eventos_painel(base):
    """O fio da tomada: mudar Painel!H3 COM eventos ligados troca o lote (Worksheet_Change)."""
    T = 'T0 ligacao do seletor'
    s = Sessao(base, 'T0')
    try:
        preparar(s, [A, B])
        s.analito(AN)
        s.xl.EnableEvents = True
        s.pai.Range('H3').Value = B[0]
        s.ex.esperar()
        time.sleep(1)
        s.xl.EnableEvents = False
        conferir(T, s, B[0], AN, B[1], rotulo='[H3 digitado] ')
        s.xl.EnableEvents = True
        s.ana.Range('E4').Value = 205.0          # digitar na aba Analitos com eventos
        s.ex.esperar()
        s.xl.EnableEvents = False
        chk(T, 'digitar media na aba Analitos grava na LotesStore na hora',
            igual(s.run('mLotes.MediaDoLote', B[0], AN, 1), 205.0), s.run('mLotes.MediaDoLote', B[0], AN, 1))
        chk(T, 'e o lote A nao foi tocado', igual(s.run('mLotes.MediaDoLote', A[0], AN, 1), 100.0), '')
    finally:
        s.fechar()


TESTES = {'T0': T_eventos_painel, 'T1': T1, 'T2': T2, 'T3': T3, 'T4': T4, 'T5': T5,
          'T6': T6, 'T7': T7, 'T8': T8, 'T9': T9}


# ============================================================================
# base limpa
# ============================================================================
def montar_base(origem, base):
    shutil.copyfile(origem, base)
    ex = xlh.Excel()
    wb = ex.abrir(os.path.abspath(base))
    try:
        vbp = wb.VBProject
        with open(os.path.join(AQUI, 'mTesteEtapa1.bas'), encoding='utf-8') as f:
            cod = '\r\n'.join(l for l in f.read().split('\n') if not l.startswith('Attribute '))
        comp = vbp.VBComponents.Add(1)
        comp.Name = 'mTesteEtapa1'
        comp.CodeModule.AddFromString(cod)
        for s in wb.Worksheets:
            if s.ProtectContents:
                s.Unprotect('qcini2025')
        ws = wb.Worksheets.Add()
        ws.Name = 'T_Entrada'
        pai = wb.Sheets('Painel')
        pai.Range('G3').Value = ''
        pai.Range('G4').Value = ''
        pai.Range('I3').Value = ''
        pai.Range('M3:N4').Value = False     # caixinhas de trimestre: base sem filtro
        ex.run('mTesteEtapa1.T_LimparBase', teto=300)
        ex.esperar()
        wb.Save()
        print('base pronta:', base)
    finally:
        ex.fechar()


def main():
    modo = sys.argv[1]
    if modo == 'base':
        montar_base(sys.argv[2], sys.argv[3])
        return
    base = os.path.abspath(sys.argv[2])
    quais = sys.argv[3:] or list(TESTES)
    t00 = time.time()
    for q in quais:
        print(f'\n=== {q} ===', flush=True)
        t0 = time.time()
        try:
            TESTES[q](base)
        except Exception as e:
            print(traceback.format_exc(), flush=True)
            chk(q, 'EXCECAO', False, f'{e.__class__.__name__}: {e}')
        print(f'   ({time.time() - t0:.0f}s)', flush=True)
    npass = sum(1 for r in RESULTADOS if r[2])
    print(f'\nTOTAL: {npass}/{len(RESULTADOS)} PASS em {time.time() - t00:.0f}s')
    saida = os.path.join(AQUI, '..', 'resultados', f'suite_{"_".join(quais)}_{time.strftime("%Y%m%d_%H%M%S")}.json')
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(RESULTADOS, f, ensure_ascii=False, indent=1)
    print('resultado:', os.path.abspath(saida))


if __name__ == '__main__':
    main()
