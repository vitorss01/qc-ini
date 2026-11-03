# -*- coding: utf-8 -*-
"""F4 do Usuario Simulado -- periodos, lotes, filtros, graficos (Levey-Jennings / Westgard) e usabilidade geral.

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    python -u f4_periodos_lotes.py <copia.xlsm> <cq.pkl> <saida.json>

    <copia.xlsm>  COPIA da producao (shutil.copy2 para %TEMP%\\qcw\\auditoria\\<ciclo>\\). Nada e salvo nela.
    <cq.pkl>      tblCQ_Final da MESMA copia lida por openpyxl (pickle {'cab','rows'}), feita antes de abrir o
                  Excel (ver o bloco "como gerar o cq.pkl" abaixo). O script confere que o numero de linhas e os
                  IDs batem com a tabela lida pelo COM antes de usar.
    <saida.json>  tudo o que foi lido e calculado, por teste.

Como gerar o cq.pkl (somente leitura, sem Excel):
    wb = openpyxl.load_workbook(copia, read_only=True, data_only=True); ws = wb['Principal - Resultados']
    cabecalho na linha 6; uma dict por linha -> pickle.dump({'cab': cab, 'rows': rows}).

O que faz (sempre pelo caminho do usuario: as celulas de filtro do Painel com os eventos ligados -- H3 lote,
K3/K4 ano/periodo, G3/G4 De/Ate, C3 analito pela lista, L4 equipamento; e as celulas de filtro da Estatistica):
  P  periodos: mes a mes (todos os meses com dado) e trimestres pelo atalho Ano/Periodo; semestre, ano e "Tudo";
     confere G3/G4, o n/media/DP/CV do Painel (engPainel) contra o calculo independente (statistics) dos dados BRUTOS
     do periodo, e mostra a diferenca para a "media dos meses". Confere o corte das 180 corridas (engCortadas).
  E  aba Estatistica: visoes MENSAL/TRIMESTRAL (B4/D4/F4) com o lote H4, 62 linhas analito x nivel x independente.
  L  lotes: lote A -> lote B -> H3 vazio (ADR-086), conferindo que o MOTOR (Eng_Saida: IDs, n, estatistica) usa
     so o filtrado; ajuste automatico do filtro ao trocar para um lote sem ponto (ADR-077/078).
  C  combinacoes: equipamentos, niveis e analitos diferentes; analito sem CEQ (ET do Painel com "SEM EP").
  G  Levey-Jennings: pontos do grafico x Eng_Saida x tblCQ_Final; corrida do lote (ADR-080) recalculada do historico;
     linhas de limite (Media + k*DP do lote, ADR-078/086); Westgard independente (ADR-083: mesmo nivel na sequencia
     de ativos; entre niveis pelo RUN) x marcas do motor x contagens do Painel; tabela x grafico.
  U  usabilidade: tempos de cada troca de filtro, mensagens (H4, A39), listas de opcoes.

Calculo independente: Python puro (statistics/math) sobre as linhas brutas da tblCQ_Final (PARTICIPA_ESTATISTICA =
SIM, analito, equipamento, lote, nivel, DATA_HORA truncada ao dia dentro de De/Ate, RESULTADO numerico). Nao le
nenhuma formula do QC_INI. A Media/DP do alvo do lote (entrada do usuario na aba Analitos/LotesStore) e lida pelas
funcoes mLotes.MediaDoLote / DPDoLote (o cadastro, nao um calculo).

Marcacao: nenhum dado e gravado nesta frente (so filtros). Senha: importada da suite de testes, nunca escrita aqui.
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01.
"""
import collections
import datetime as dt
import json
import math
import os
import pickle
import statistics
import sys
import threading
import time
import traceback

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
from qa_final import QA  # noqa: E402
from qa_etl_recebimento import SENHA  # noqa: E402

EPOCH = dt.datetime(1899, 12, 30)
OUT = {'testes': [], 'erros': [], 'tempos': {}, 'mensagens': []}
LIM = ['-3s', '-2s', '-1s', 'Média', '+1s', '+2s', '+3s']
KS = [-3, -2, -1, 0, 1, 2, 3]
MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
TOL = 1e-9


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def ingenuo(v):
    if isinstance(v, dt.datetime):
        return dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second, v.microsecond)
    return v


def dia(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if num(v) and v > 0:
        return (EPOCH + dt.timedelta(days=int(v))).date()
    return None


def serie(d):
    return float((dt.datetime(d.year, d.month, d.day) - EPOCH).days)


def rel(a, b):
    if not (num(a) and num(b)):
        return None
    m = max(abs(a), abs(b), 1e-300)
    return abs(a - b) / m


def j(v):
    return json.loads(json.dumps(v, ensure_ascii=False, default=str))


def reg(tid, desc, estado, evid):
    OUT['testes'].append({'id': tid, 'descricao': desc, 'estado': estado, 'evidencia': j(evid)})
    log(estado, tid, desc, json.dumps(j(evid), ensure_ascii=False)[:700])


def stats(vals):
    n = len(vals)
    if n == 0:
        return {'n': 0, 'media': None, 'dp': None, 'cv': None}
    m = statistics.fmean(vals)
    sd = statistics.stdev(vals) if n > 1 else None
    cv = sd / m * 100 if (sd is not None and m != 0) else None
    return {'n': n, 'media': m, 'dp': sd, 'cv': cv}


class Watch:
    """Watchdog para escrita em celula com eventos: um modal invisivel travaria o COM para sempre."""

    def __init__(self, ex, teto):
        self.ex, self.teto, self.ev = ex, teto, threading.Event()

    def __enter__(self):
        def cao():
            if not self.ev.wait(self.teto):
                log('WATCHDOG: passou de', self.teto, 's -- encerrando o Excel desta sessao (PID', self.ex.pid, ')')
                self.ex.matar()
        threading.Thread(target=cao, daemon=True).start()
        return self

    def __exit__(self, *a):
        self.ev.set()


# ====================================================================== dados brutos (independente)
class Bruto:
    def __init__(self, rows):
        self.rows = rows
        self.elig = []
        self.xv = []
        for r in rows:
            dh = r.get('DATA_HORA')
            if not isinstance(dh, dt.datetime):
                continue
            base = {'id': txt(r['ID_REGISTRO']), 'an': txt(r['ANALITO']), 'eq': txt(r['EQUIPAMENTO']).upper(),
                    'lote': txt(r['LOTE']), 'nv': int(float(r['NIVEL'])) if num(r['NIVEL']) or txt(r['NIVEL']).isdigit() else None,
                    'dh': ingenuo(dh), 'run': txt(r['RUN']), 'val': r['RESULTADO']}
            if txt(r['PARTICIPA_ESTATISTICA']).upper() == 'SIM' and num(r['RESULTADO']):
                self.elig.append(base)
            elif txt(r.get('TIPO_PLOTAGEM_LJ')).upper() == 'X_VERMELHO' and num(r['RESULTADO']):
                self.xv.append(base)

    def sel(self, an, eq, lote, nv=None, d0=None, d1=None):
        out = []
        for e in self.elig:
            if e['an'] != an or e['eq'] != eq.upper():
                continue
            if lote and e['lote'] != lote:
                continue
            if nv is not None and e['nv'] != nv:
                continue
            d = e['dh'].date()
            if d0 and d < d0:
                continue
            if d1 and d > d1:
                continue
            out.append(e)
        return out

    def stats_periodo(self, an, eq, lote, nv, d0, d1):
        return stats([float(e['val']) for e in self.sel(an, eq, lote, nv, d0, d1)])

    def corridas(self, an, eq, lote, todos=False):
        """Eixo de corridas do analito/equipamento/lote (todos=True: RUN|LOTE de todos os lotes), em ordem
        cronologica. Cada corrida: chave, data (1a linha), {nivel: (valor, id, lote)} dos elegiveis, so_x."""
        cor = collections.OrderedDict()
        for fonte, eleg in ((self.elig, True), (self.xv, False)):
            for e in fonte:
                if e['an'] != an or e['eq'] != eq.upper():
                    continue
                if not todos and e['lote'] != lote:
                    continue
                k = e['run'] + ('|' + e['lote'] if todos else '')
                c = cor.setdefault(k, {'k': k, 'run': e['run'], 'dh': e['dh'], 'v': {}, 'dup': [], 'eleg': False,
                                       'lote': e['lote']})
                if e['dh'] < c['dh']:
                    c['dh'] = e['dh']
                if eleg:
                    c['eleg'] = True
                    if e['nv'] in c['v']:
                        c['dup'].append(e['id'])
                    else:
                        c['v'][e['nv']] = (float(e['val']), e['id'], e['lote'])
        return sorted(cor.values(), key=lambda c: (c['dh'], c['run']))


def corrida_do_lote(eixo_lote, nlv=2):
    """ADR-080/081: posicao 1,2,3... entre os resultados ATIVOS do nivel, no historico inteiro do lote."""
    out = {}
    cont = [0] * nlv
    for c in eixo_lote:
        for t in range(nlv):
            if (t + 1) in c['v']:
                cont[t] += 1
                out[c['v'][t + 1][1]] = cont[t]
    return out


def janela(eixo, fim, nk=180):
    """As nk corridas ELEGIVEIS mais recentes com data <= fim. Corridas so-X dentro do intervalo entram enquanto
    houver posicao livre (nk - elegiveis); faltando, saem as so-X mais antigas (definicao do D10/QA-ETL-001,
    conferida no texto do AtualizarCalc -- o X nunca tira lugar de resultado elegivel).
    Revisao do 2o passe (ciclo01): antes, com a janela cheia, as so-X ficavam dentro."""
    ate = [c for c in eixo if fim is None or c['dh'].date() <= fim]
    eleg_idx = [i for i, c in enumerate(ate) if c['eleg']]
    i0 = eleg_idx[-nk] if len(eleg_idx) >= nk else 0
    n_eleg = min(len(eleg_idx), nk)
    sox = [i for i in range(i0, len(ate)) if not ate[i]['eleg']]
    pular = set(sox[:max(0, len(sox) - (nk - n_eleg))])
    jan = [ate[i] for i in range(i0, len(ate)) if i not in pular]
    return jan, len([c for c in ate[:i0] if c['eleg']])


def westgard(eixo, alvo, nlv=2):
    """Westgard independente (Bio, N=2) pela definicao do ADR-083 / ADR-041:
       1_3s |z|>3; 2_2s WITHIN_RUN (os 2 niveis do RUN, mesmo lado, >2); 2_2s ACROSS_RUN_SAME_LEVEL (2 observacoes
       ATIVAS seguidas do nivel, mesmo lado, >2; marca a 2a); R_4s |z1-z2|>4 no RUN; 4_1s N2_R2 (2 RUNs completos
       seguidos, 4 obs >1 do mesmo lado; marca os 2 niveis do 2o); 8x N2_R4 (4 RUNs completos, 8 obs do mesmo lado
       da media). alvo = {nivel: (M, S)}. Devolve {(chave, nivel): set(regras)}."""
    n = len(eixo)
    z = [[None] * n for _ in range(nlv)]
    for i, c in enumerate(eixo):
        for t in range(nlv):
            if (t + 1) in c['v']:
                lt = c['v'][t + 1][2]
                M, S = alvo.get((lt, t + 1), (None, None))
                if num(M) and num(S) and S > 0:
                    z[t][i] = (c['v'][t + 1][0] - M) / S
    fl = collections.defaultdict(set)

    def comp(i):
        return all(z[t][i] is not None for t in range(nlv))
    for i in range(n):
        for t in range(nlv):
            if z[t][i] is not None and abs(z[t][i]) > 3:
                fl[(i, t)].add('1_3s')
        if comp(i):
            if (z[0][i] > 2 and z[1][i] > 2) or (z[0][i] < -2 and z[1][i] < -2):
                fl[(i, 0)].add('2_2s'); fl[(i, 1)].add('2_2s')
            if abs(z[0][i] - z[1][i]) > 4:
                fl[(i, 0)].add('R_4s'); fl[(i, 1)].add('R_4s')
        if i >= 1 and comp(i) and comp(i - 1):
            obs = [z[t][k] for k in (i - 1, i) for t in range(nlv)]
            if all(o > 1 for o in obs) or all(o < -1 for o in obs):
                fl[(i, 0)].add('4_1s'); fl[(i, 1)].add('4_1s')
        if i >= 3 and all(comp(k) for k in range(i - 3, i + 1)):
            obs = [z[t][k] for k in range(i - 3, i + 1) for t in range(nlv)]
            if all(o > 0 for o in obs) or all(o < 0 for o in obs):
                fl[(i, 0)].add('8x'); fl[(i, 1)].add('8x')
    for t in range(nlv):
        seq = [i for i in range(n) if z[t][i] is not None]
        for a, b in zip(seq, seq[1:]):
            if (z[t][a] > 2 and z[t][b] > 2) or (z[t][a] < -2 and z[t][b] < -2):
                fl[(b, t)].add('2_2s')
    return fl, z


# ====================================================================== leitura do QC_INI (copia)
class Tela:
    def __init__(self, q, bruto):
        self.q, self.b = q, bruto
        self.xl = q.ex.xl
        self.p = q.ws('Painel')
        self.e = q.ws('Eng_Saida')
        self.est = q.ws('Estatística')
        self.cad = [txt(x[0]) for x in q.espec().Range('A4:A43').Value]
        self.alvo_cache = {}

    # ---------------------------------------------------- acoes do usuario
    def muda(self, cel, valor, teto=600, aba=None):
        ws = aba or self.p
        self.xl.EnableEvents = True
        t0 = time.time()
        with Watch(self.q.ex, teto):
            if valor is None:
                ws.Range(cel).MergeArea.ClearContents()
            else:
                ws.Range(cel).Value = valor
            self.q.ex.esperar()
        dtc = round(time.time() - t0, 2)
        OUT['tempos'].setdefault(cel, []).append(dtc)
        return dtc

    def estado(self):
        p, e = self.p, self.e
        d = {}
        for a in ('B3', 'C3', 'C4', 'G3', 'G4', 'H3', 'H4', 'K3', 'K4', 'L4', 'M4', 'O4', 'Q4', 'T4', 'A39'):
            v = p.Range(a).Value
            d[a] = ingenuo(v) if isinstance(v, dt.datetime) else v
        d['B7:T8'] = [[ingenuo(x) if isinstance(x, dt.datetime) else x for x in l] for l in p.Range('A7:T8').Value]
        d['engPainel'] = [list(l) for l in self.q.nome('engPainel').Value]
        for a in ('C1', 'E1', 'G1', 'I1', 'K1', 'M1', 'O1', 'S1', 'W1', 'Y1', 'AA1', 'AD1'):
            v = e.Range(a).Value
            d['eng_' + a] = ingenuo(v) if isinstance(v, dt.datetime) else v
        return d

    def eng(self):
        e = self.e
        n = int(e.Range('I1').Value or 0)
        out = []
        if not n:
            return out
        blk = e.Range(e.Cells(3, 1), e.Cells(2 + n, 80)).Value
        for k, r in enumerate(blk):
            out.append({'slot': k + 1, 'run': txt(r[1]), 'fil': r[23],
                        'val': [r[24 + t] for t in range(2)],
                        'id': [txt(r[38 + t]) for t in range(2)],
                        'dh': [ingenuo(r[50 + t]) if isinstance(r[50 + t], dt.datetime) else r[50 + t] for t in range(2)],
                        'cor': [r[62 + t] for t in range(2)],
                        'alvo': [(r[74 + 2 * t], r[75 + 2 * t]) for t in range(2)],
                        'flags': [list(r[2 + t * 7:2 + t * 7 + 7]) for t in range(2)]})
        return out

    def alvo(self, lote, an, nv):
        k = (lote, an, nv)
        if k not in self.alvo_cache:
            m = self.q.run('mLotes.MediaDoLote', lote, an, nv, teto=60)
            s = self.q.run('mLotes.DPDoLote', lote, an, nv, teto=60)
            self.alvo_cache[k] = (m, s)
        return self.alvo_cache[k]

    def series(self, i):
        ch = self.p.ChartObjects(i).Chart
        out = {}
        for s_i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(s_i)
            try:
                xs, ys = list(s.XValues or []), list(s.Values or [])
            except Exception:
                xs, ys = [], []
            out[f'{s_i}:{s.Name}'] = [(x, y) for x, y in zip(xs, ys) if num(x) and num(y)]
        eixo = ch.Axes(1)
        eixoy = ch.Axes(2)
        info = {'titulo': ch.ChartTitle.Text if ch.HasTitle else None,
                'x_min': eixo.MinimumScale, 'x_max': eixo.MaximumScale,
                'y_min': eixoy.MinimumScale, 'y_max': eixoy.MaximumScale}
        return out, info

    def estat_linhas(self):
        v = self.est.Range('A14:F93').Value
        out = {}
        for r in v:
            an = txt(r[0])
            if an and num(r[1]):
                out[(an, int(r[1]))] = list(r[2:6])
        filtro = {}
        for a in ('B4', 'D4', 'F4', 'H4', 'B5', 'D5', 'H5', 'J5'):
            x = self.est.Range(a).Value
            filtro[a] = ingenuo(x) if isinstance(x, dt.datetime) else x
        for nm in ('Estat_Ini_Efetiva', 'Estat_Fim_Efetiva', 'Estat_Lote'):
            try:
                x = self.q.wb.Names(nm).RefersToRange.Value
                filtro[nm] = ingenuo(x) if isinstance(x, dt.datetime) else x
            except Exception as ex:
                filtro[nm] = f'erro {type(ex).__name__}'
        return out, filtro

    def validacao(self, ws, cel):
        try:
            v = ws.Range(cel).Validation
            return {'tipo': v.Type, 'f1': v.Formula1}
        except Exception as ex:
            return {'erro': type(ex).__name__}

    def lista(self, ws, cel):
        """Opcoes da lista de validacao (formula de intervalo ou lista literal)."""
        v = self.validacao(ws, cel)
        f = v.get('f1') or ''
        if not f:
            return v, []
        if f.startswith('='):
            try:
                rng = self.q.wb.Application.Evaluate(f[1:]) if '!' in f or f[1:] in [n.Name for n in self.q.wb.Names] else ws.Range(f[1:])
                vals = rng.Value
                if isinstance(vals, tuple):
                    return v, [txt(x[0]) if isinstance(x, tuple) else txt(x) for x in vals if (x[0] if isinstance(x, tuple) else x) not in (None, '')]
                return v, [txt(vals)]
            except Exception as ex:
                return v, ['erro ' + type(ex).__name__ + ' ' + str(ex)[:80]]
        sep = ';' if ';' in f else ','
        return v, [x.strip() for x in f.split(sep)]


# ====================================================================== comparacoes
def cmp_painel(tela, est, an, eq, lote, d0, d1, tid, desc, extra=None):
    """engPainel (n, media, DP, CV) x independente dos dados brutos do periodo."""
    linhas = []
    pior = 0.0
    ok = True
    for t in range(2):
        ind = tela.b.stats_periodo(an, eq, lote, t + 1, d0, d1)
        ep = est['engPainel'][t]
        qc = {'n': ep[1], 'media': ep[2], 'dp': ep[3], 'cv': ep[4]}
        difs = {}
        for k in ('n', 'media', 'dp', 'cv'):
            a, b = qc[k], ind[k]
            if b is None:
                if a not in (None, ''):
                    ok = False
                difs[k] = None
                continue
            r = rel(a, b)
            difs[k] = r
            if r is None or r > TOL:
                ok = False
            else:
                pior = max(pior, r)
        linhas.append({'nivel': t + 1, 'qc': qc, 'independente': ind, 'dif_rel': difs})
    evid = {'analito': an, 'equipamento': eq, 'lote': lote, 'de': d0, 'ate': d1, 'G3': est['G3'], 'G4': est['G4'],
            'H4': est['H4'], 'engCortadas_M1': est['eng_M1'], 'janela_I1': est['eng_I1'], 'linhas': linhas,
            'maior_dif_rel': pior}
    if extra:
        evid.update(extra)
    reg(tid, desc, 'APROVADO' if ok else 'REPROVADO', evid)
    return ok, linhas


def cmp_eng_escopo(tela, an, eq, lote, d0, d1, todos=False):
    """Todo ponto do Eng_Saida com filtro = 1 pertence ao analito/equipamento/lote/nivel/periodo, e todo resultado
    elegivel do periodo esta la (no slot do seu RUN)."""
    eng = tela.eng()
    cq = {e['id']: e for e in tela.b.elig}
    fora, faltam = [], []
    vistos = set()
    for x in eng:
        if x['fil'] != 1:
            continue
        for t in range(2):
            i = x['id'][t]
            if not i:
                continue
            vistos.add(i)
            e = cq.get(i)
            if (e is None or e['an'] != an or e['eq'] != eq.upper() or (not todos and e['lote'] != lote)
                    or e['nv'] != t + 1 or e['dh'].date() < d0 or e['dh'].date() > d1):
                fora.append({'id': i, 'nivel': t + 1, 'cq': e and {k: e[k] for k in ('an', 'eq', 'lote', 'nv', 'dh')}})
            elif not (num(x['val'][t]) and abs(x['val'][t] - e['val']) < 1e-12):
                fora.append({'id': i, 'valor_eng': x['val'][t], 'valor_cq': e['val']})
    for t in range(2):
        for e in tela.b.sel(an, eq, None if todos else lote, t + 1, d0, d1):
            if e['id'] not in vistos:
                faltam.append({'id': e['id'], 'nivel': t + 1, 'lote': e['lote'], 'dh': e['dh']})
    return eng, fora, faltam


# ====================================================================== roteiro
def executar(copia, pkl, saida):
    d = pickle.load(open(pkl, 'rb'))
    b = Bruto(d['rows'])
    q = QA('Bioquimica', copia)
    OUT['excel_pid'] = q.ex.pid
    try:
        t = Tela(q, b)
        xl = t.xl
        # conferencia: o pickle e a tabela da copia sao a mesma coisa
        lo = q.lo('tblCQ_Final')
        n_com = lo.ListRows.Count
        ids_com = [txt(x[0]) for x in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        ids_pkl = [txt(r['ID_REGISTRO']) for r in d['rows']]
        reg('T00', 'Pré-condição: tblCQ_Final do COM = leitura openpyxl da mesma cópia (base do cálculo independente)',
            'APROVADO' if ids_com == ids_pkl else 'REPROVADO', {'linhas_com': n_com, 'linhas_pkl': len(ids_pkl)})
        q.run('mSeguranca.ReprotectAll', teto=120)       # como o usuario encontra o Painel (protegido)
        prot = t.p.ProtectContents
        OUT['painel_protegido'] = prot
        # listas
        v_k3, l_k3 = t.lista(t.p, 'K3')
        v_k4, l_k4 = t.lista(t.p, 'K4')
        v_h3, l_h3 = t.lista(t.p, 'H3')
        v_l4, l_l4 = t.lista(t.p, 'L4')
        v_c3, l_c3 = t.lista(t.p, 'C3')
        OUT['listas'] = {'K3': [v_k3, l_k3], 'K4': [v_k4, l_k4], 'H3': [v_h3, l_h3], 'L4': [v_l4, l_l4], 'C3': [v_c3, l_c3[:40]]}
        log('K4', l_k4)
        log('H3', l_h3)
        OUT['estado_inicial'] = j(t.estado())

        def rot(nome, um=None):
            nome_l = nome.lower()
            for lab in l_k4:
                if lab.lower() == nome_l or lab.lower().startswith(nome_l + ' '):   # rotulos reais: "Ano inteiro", "Tudo (sem filtro)"
                    return lab
            return um

        mes_lab = {}
        for lab in l_k4:
            k = lab.lower()[:3]
            if k in MESES and len(lab) <= 9:
                mes_lab[MESES.index(k) + 1] = lab
        tri_lab = {}
        sem_lab = {}
        for lab in l_k4:
            L = lab.upper().replace('º', '').replace('°', '')
            for k in range(1, 5):
                if L in (f'T{k}', f'{k}T', f'{k} TRI', f'{k}TRI', f'TRI {k}', f'{k} TRIMESTRE', f'{k}O TRI'):
                    tri_lab[k] = lab
            for k in range(1, 3):
                if L in (f'S{k}', f'{k}S', f'{k} SEM', f'{k}SEM', f'SEM {k}', f'{k} SEMESTRE'):
                    sem_lab[k] = lab
        OUT['rotulos'] = {'mes': mes_lab, 'tri': tri_lab, 'sem': sem_lab}

        AN, EQ = 'Glicose', 'DIMENSION 1'
        # ------------------------------------------------ cenario base: Glicose (rotina estavel), DIMENSION 1
        try:
            idx = t.cad.index(AN) + 1
            t.muda('C3', AN)                       # escolher pelo nome (lista), como o usuario
            t.muda('L4', EQ)
            OUT['base'] = j(t.estado())
        except Exception as ex:
            OUT['erros'].append(f'base: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ P: periodos mes a mes e trimestres
        def lote_do_mes(ano, mes):
            c = collections.Counter(e['lote'] for e in b.sel(AN, EQ, None)
                                    if e['dh'].year == ano and e['dh'].month == mes)
            return c.most_common(1)[0][0] if c else None

        meses = sorted({(e['dh'].year, e['dh'].month) for e in b.sel(AN, EQ, None)})
        res_mes = {}
        for (ano, mes) in meses:
            lt = lote_do_mes(ano, mes)
            try:
                if txt(t.p.Range('H3').Value) != lt:
                    t.muda('H3', "'" + lt)
                if t.p.Range('K3').Value != ano:
                    t.muda('K3', ano)
                dtc = t.muda('K4', mes_lab.get(mes, MESES[mes - 1].capitalize()))
                est = t.estado()
                d0 = dt.date(ano, mes, 1)
                d1 = (dt.date(ano + (mes == 12), mes % 12 + 1, 1) - dt.timedelta(days=1))
                g_ok = dia(est['G3']) == d0 and dia(est['G4']) == d1
                ok, linhas = cmp_painel(t, est, AN, EQ, lt, d0, d1, f'P-M{ano}{mes:02d}',
                                        f'Período mês {mes:02d}/{ano} (Ano/Período) — Glicose DIM 1 lote {lt}: Painel = dados brutos do mês',
                                        {'G3G4_corretos': g_ok, 'tempo_s': dtc, 'status_J7_J8': [est['B7:T8'][0][9], est['B7:T8'][1][9]]})
                if not g_ok:
                    reg(f'P-G{ano}{mes:02d}', f'G3/G4 do mês {mes:02d}/{ano}', 'REPROVADO', {'G3': est['G3'], 'G4': est['G4'], 'esperado': [d0, d1]})
                res_mes[(ano, mes)] = {'lote': lt, 'linhas': linhas}
            except Exception as ex:
                OUT['erros'].append(f'P mes {ano}-{mes}: {ex!r} {traceback.format_exc()[-800:]}')
                reg(f'P-M{ano}{mes:02d}', f'Período mês {mes:02d}/{ano}', 'BLOQUEADO', {'erro': repr(ex)})

        # trimestres
        tris = [(2025, 4, '8974'), (2026, 1, '8974'), (2026, 2, '8974'), (2026, 3, '8977'), (2026, 4, '8977')]
        for (ano, k, lt) in tris:
            try:
                if txt(t.p.Range('H3').Value) != lt:
                    t.muda('H3', "'" + lt)
                if t.p.Range('K3').Value != ano:
                    t.muda('K3', ano)
                dtc = t.muda('K4', tri_lab.get(k, f'T{k}'))
                est = t.estado()
                d0 = dt.date(ano, 3 * k - 2, 1)
                d1 = dt.date(ano + (k == 4), (3 * k) % 12 + 1, 1) - dt.timedelta(days=1)
                # alternativa errada: media dos meses (para mostrar que o trimestre NAO e isso)
                alt = []
                for t_ in range(2):
                    mm = [statistics.fmean([float(e['val']) for e in b.sel(AN, EQ, lt, t_ + 1, dt.date(ano, m, 1),
                          dt.date(ano + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))])
                          for m in range(3 * k - 2, 3 * k + 1) if b.sel(AN, EQ, lt, t_ + 1, dt.date(ano, m, 1),
                          dt.date(ano + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))]
                    cvs = []
                    for m in range(3 * k - 2, 3 * k + 1):
                        s = b.stats_periodo(AN, EQ, lt, t_ + 1, dt.date(ano, m, 1),
                                            dt.date(ano + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))
                        if s['cv'] is not None:
                            cvs.append(s['cv'])
                    alt.append({'nivel': t_ + 1, 'media_das_medias_mensais': statistics.fmean(mm) if mm else None,
                                'media_dos_CV_mensais': statistics.fmean(cvs) if cvs else None})
                g_ok = dia(est['G3']) == d0 and dia(est['G4']) == d1
                cmp_painel(t, est, AN, EQ, lt, d0, d1, f'P-T{ano}{k}',
                           f'Trimestre T{k}/{ano} — Glicose DIM 1 lote {lt}: Painel = dados brutos do trimestre (não média dos meses)',
                           {'G3G4_corretos': g_ok, 'tempo_s': dtc, 'alternativa_media_dos_meses': alt})
            except Exception as ex:
                OUT['erros'].append(f'P tri {ano}-{k}: {ex!r} {traceback.format_exc()[-800:]}')
                reg(f'P-T{ano}{k}', f'Trimestre T{k}/{ano}', 'BLOQUEADO', {'erro': repr(ex)})

        # semestre / ano / tudo com o lote 8974 (> 180 corridas) -> corte da janela
        for (ano, lab, d0, d1) in [(2026, sem_lab.get(1), dt.date(2026, 1, 1), dt.date(2026, 6, 30)),
                                   (2026, rot('Ano'), dt.date(2026, 1, 1), dt.date(2026, 12, 31)),
                                   (2026, rot('Tudo'), None, None)]:
            if not lab:
                continue
            try:
                if txt(t.p.Range('H3').Value) != '8974':
                    t.muda('H3', "'8974")
                if t.p.Range('K3').Value != ano:
                    t.muda('K3', ano)
                dtc = t.muda('K4', lab)
                est = t.estado()
                dd0, dd1 = (d0 or dt.date(1900, 1, 1)), (d1 or dt.date(2100, 1, 1))
                eixo = b.corridas(AN, EQ, '8974')
                jan, cort = janela(eixo, d1)
                # o que o Painel deveria mostrar se usasse SO a janela de 180 corridas
                win = []
                for t_ in range(2):
                    vals = [c['v'][t_ + 1][0] for c in jan if (t_ + 1) in c['v'] and dd0 <= c['dh'].date() <= dd1]
                    win.append(stats(vals))
                cort_periodo = len([c for c in eixo[:len(eixo) - len(jan)] if c['eleg'] and dd0 <= c['dh'].date() <= dd1])
                cmp_painel(t, est, AN, EQ, '8974', dd0, dd1, f'P-L{lab}',
                           f'Período "{lab}" {ano} — lote 8974 (mais de 180 corridas): Painel = todos os dados brutos do período',
                           {'tempo_s': dtc, 'corridas_periodo_fora_da_janela_independente': cort_periodo,
                            'estatistica_so_da_janela_180': win, 'A39': est['A39']})
            except Exception as ex:
                OUT['erros'].append(f'P {lab}: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ E: aba Estatistica (visoes do periodo)
        try:
            ws = t.est
            v_b4, l_b4 = t.lista(ws, 'B4')
            v_h4, l_h4 = t.lista(ws, 'H4')
            OUT['listas']['Est_B4'] = [v_b4, l_b4]
            OUT['listas']['Est_H4'] = [v_h4, l_h4[:40]]
            OUT['est_inicial'] = j(t.estat_linhas()[1])
            # cabecalho da aba (formulas e valores) e formulas da 1a linha de dados: para saber o que cada celula faz
            cab = {}
            for r_ in range(1, 14):
                for c_ in range(1, 33):
                    cel = ws.Cells(r_, c_)
                    fo = cel.Formula
                    if fo not in (None, ''):
                        v_ = cel.Value
                        cab[cel.Address.replace('$', '')] = [str(fo)[:300], ingenuo(v_) if isinstance(v_, dt.datetime) else v_]
            OUT['est_cabecalho'] = j(cab)
            OUT['est_linha14'] = j({ws.Cells(14, c_).Address.replace('$', ''): str(ws.Cells(14, c_).Formula)[:300] for c_ in range(1, 21)})
            if not l_b4:
                l_b4 = ['MENSAL', 'TRIMESTRAL']          # sem lista de validacao em B4: os nomes do mEstatPeriodo
                OUT['listas']['Est_B4_fallback'] = l_b4
            casos_e = []
            for lab in l_b4:
                L = lab.upper()
                if L.startswith('MENS'):
                    casos_e += [(lab, 2026, 3, '8974', dt.date(2026, 3, 1), dt.date(2026, 3, 31)),
                                (lab, 2026, 9, '8977', dt.date(2026, 9, 1), dt.date(2026, 9, 30))]
                if L.startswith('TRIM'):
                    casos_e += [(lab, 2026, 1, '8974', dt.date(2026, 1, 1), dt.date(2026, 3, 31)),
                                (lab, 2026, 3, '8977', dt.date(2026, 7, 1), dt.date(2026, 9, 30))]
            ws.Unprotect(SENHA)
            for (lab, ano, parte, lt, d0, d1) in casos_e:
                t.muda('B4', lab, aba=ws)
                t.muda('D4', ano, aba=ws)
                v_f4, l_f4 = t.lista(ws, 'F4')
                alvo_f4 = parte
                if l_f4 and not any(txt(x) == str(parte) for x in l_f4):
                    # rotulo textual (ex.: "Mar", "T1")
                    if lab.upper().startswith('MENS'):
                        alvo_f4 = next((x for x in l_f4 if x.lower()[:3] == MESES[parte - 1]), parte)
                    else:
                        alvo_f4 = next((x for x in l_f4 if str(parte) in x), parte)
                t.muda('F4', alvo_f4, aba=ws)
                t.muda('H4', "'" + lt, aba=ws)
                xl.Calculate()
                q.ex.esperar()
                linhas, filtro = t.estat_linhas()
                difs, pior, comp = [], 0.0, 0
                for (an, nv), qcv in linhas.items():
                    ind = b.stats_periodo(an, EQ, lt, nv, d0, d1)
                    if ind['n'] == 0 and all(a in (None, '', 0) for a in qcv):
                        continue
                    for k, a in zip(('n', 'media', 'dp', 'cv'), qcv):
                        bb = ind[k]
                        if bb is None and a in (None, '', 0) and k != 'n':
                            continue
                        if bb is None and k != 'n':
                            difs.append({'an': an, 'nv': nv, 'k': k, 'qc': a, 'ind': bb})
                            continue
                        r = rel(a, bb if bb is not None else 0)
                        comp += 1
                        if r is None or r > TOL:
                            difs.append({'an': an, 'nv': nv, 'k': k, 'qc': a, 'ind': bb})
                        else:
                            pior = max(pior, r)
                reg(f'E-{lab[:4]}-{ano}{parte:02d}-{lt}',
                    f'Aba Estatística, visão {lab} parte {alvo_f4}/{ano}, lote {lt}, {EQ}: 62 linhas × n/média/DP/CV = dados brutos',
                    'APROVADO' if not difs and comp > 0 else 'REPROVADO',
                    {'filtro': filtro, 'opcoes_F4': l_f4[:20], 'comparacoes': comp, 'divergencias': difs[:12],
                     'n_divergencias': len(difs), 'maior_dif_rel': pior,
                     'janela_esperada': [d0, d1]})
        except Exception as ex:
            OUT['erros'].append(f'E: {ex!r} {traceback.format_exc()[-1200:]}')

        # ------------------------------------------------ L: lote A -> lote B -> H3 vazio (periodo que cruza os lotes)
        try:
            t.muda('L4', EQ)
            t.muda('G3', serie(dt.date(2026, 6, 1)))
            t.muda('G4', serie(dt.date(2026, 8, 31)))
            d0, d1 = dt.date(2026, 6, 1), dt.date(2026, 8, 31)
            for lt in ('8974', '8977', None):
                dtc = t.muda('H3', None if lt is None else "'" + lt)
                est = t.estado()
                eng, fora, faltam = cmp_eng_escopo(t, AN, EQ, lt, d0, d1, todos=lt is None)
                lotes_pts = collections.Counter()
                cq = {e['id']: e for e in b.elig}
                for x in eng:
                    if x['fil'] == 1:
                        for t_ in range(2):
                            if x['id'][t_] in cq:
                                lotes_pts[cq[x['id'][t_]]['lote']] += 1
                if lt is None:
                    ind = [len(b.sel(AN, EQ, None, t_ + 1, d0, d1)) for t_ in range(2)]
                    qc_n = [est['engPainel'][t_][1] for t_ in range(2)]
                    vazios = all(est['engPainel'][t_][k] in (None, '') for t_ in range(2) for k in (2, 3, 4, 6, 7, 8))
                    ok = (est['H3'] in (None, '') and est['eng_AD1'] is True and qc_n == ind and not fora and not faltam
                          and vazios)
                    reg('L-TODOS', 'H3 vazio (ADR-086): todos os lotes do período; o motor usa só analito/equipamento/nível/período; '
                        'média/DP/CV/Bias/ET/Sigma vazios; n = todos os lotes',
                        'APROVADO' if ok else 'REPROVADO',
                        {'H3': est['H3'], 'H4': est['H4'], 'engTodos': est['eng_AD1'], 'n_qc': qc_n, 'n_independente': ind,
                         'pontos_por_lote': dict(lotes_pts), 'fora_do_escopo': fora[:6], 'faltando': faltam[:6],
                         'linha_painel': est['B7:T8'], 'tempo_s': dtc})
                    # alvo de cada ponto = o do seu lote
                    y2 = []
                    for x in eng:
                        for t_ in range(2):
                            i = x['id'][t_]
                            if i and i in cq:
                                m_, s_ = t.alvo(cq[i]['lote'], AN, t_ + 1)
                                g = x['alvo'][t_]
                                if num(m_) and num(s_) and s_ > 0:
                                    if not (num(g[0]) and num(g[1]) and abs(g[0] - m_) < 1e-9 and abs(g[1] - s_) < 1e-9):
                                        y2.append({'id': i, 'lote': cq[i]['lote'], 'eng': g, 'cadastro': (m_, s_)})
                    reg('L-TODOS-ALVO', 'H3 vazio: cada ponto usa a Média/DP do SEU lote (Eng_Saida BW..)',
                        'APROVADO' if not y2 else 'REPROVADO', {'divergencias': y2[:8]})
                    OUT['todos_eng'] = j(eng)
                    OUT['todos_estado'] = j(est)
                else:
                    ok, _ = cmp_painel(t, est, AN, EQ, lt, d0, d1, f'L-{lt}',
                                       f'Lote {lt} em H3, período 01/06–31/08/2026 (cruza a troca de lote): estatística só do lote',
                                       {'tempo_s': dtc, 'pontos_por_lote_no_motor': dict(lotes_pts)})
                    reg(f'L-{lt}-MOTOR', f'Lote {lt}: o MOTOR (Eng_Saida) só tem pontos do lote/analito/equipamento/nível/período, e todos eles',
                        'APROVADO' if not fora and not faltam else 'REPROVADO',
                        {'fora_do_escopo': fora[:6], 'faltando': faltam[:6], 'pontos_por_lote': dict(lotes_pts)})
                    # Estatistica (aba) segue o lote do Painel?
                    linhas, filtro = t.estat_linhas()
                    OUT.setdefault('est_apos_lote', {})[lt] = j({'filtro': filtro, 'glicose': [linhas.get((AN, 1)), linhas.get((AN, 2))]})
        except Exception as ex:
            OUT['erros'].append(f'L: {ex!r} {traceback.format_exc()[-1200:]}')

        # ajuste automatico: Out/2026 e lote 8974 (sem ponto de Glicose DIM 1 em outubro)
        try:
            t.muda('H3', "'8977")
            t.muda('K3', 2026)
            t.muda('K4', mes_lab.get(10, 'Out'))
            antes = t.estado()
            dtc = t.muda('H3', "'8974")
            dep = t.estado()
            dtc2 = t.muda('H3', "'8977")
            volta = t.estado()
            reg('L-AJUSTE', 'Trocar para um lote sem ponto no filtro (8974 em out/2026): o Painel ajusta o período e avisa (ADR-077/078); '
                'voltar ao 8977',
                'APROVADO' if (dep['G3'] != antes['G3'] or dep['G4'] != antes['G4']) and txt(dep['H4']) else 'REPROVADO',
                {'antes': {k: antes[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'depois_8974': {k: dep[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39', 'C3', 'L4')},
                 'volta_8977': {k: volta[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'tempos_s': [dtc, dtc2], 'n_depois': [dep['engPainel'][0][1], dep['engPainel'][1][1]]})
        except Exception as ex:
            OUT['erros'].append(f'L-AJUSTE: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ C: equipamentos, analitos, niveis
        try:
            t.muda('H3', "'8977")
            t.muda('K3', 2026)
            t.muda('K4', mes_lab.get(9, 'Set'))
            d0, d1 = dt.date(2026, 9, 1), dt.date(2026, 9, 30)
            for an, eq in [('Glicose', 'DIMENSION 2'), ('Ureia', 'DIMENSION 2'), ('Proteína total', 'DIMENSION 1'),
                           ('Creatinina', 'DIMENSION 1'), ('Colesterol total', 'DIMENSION 2')]:
                if txt(t.p.Range('L4').Value).upper() != eq:
                    t.muda('L4', eq)
                dtc = t.muda('C3', an)
                est = t.estado()
                ok, _ = cmp_painel(t, est, an, eq, '8977', d0, d1, f'C-{an[:6]}-{eq[-1]}',
                                   f'{an} / {eq} / lote 8977 / set-2026: Painel = dados brutos',
                                   {'tempo_s': dtc, 'C3': est['C3'], 'L4': est['L4'], 'eng_analito': est['eng_C1'], 'eng_equip': est['eng_S1']})
                eng, fora, faltam = cmp_eng_escopo(t, an, eq, '8977', d0, d1)
                reg(f'C-{an[:6]}-{eq[-1]}-MOTOR', f'{an} / {eq}: pontos do motor = só desse analito/equipamento/lote, nível no lugar certo',
                    'APROVADO' if not fora and not faltam else 'REPROVADO', {'fora': fora[:6], 'faltam': faltam[:6]})
            # analito sem CEQ (HbA1c / PCR): ET do Painel
            for an in ('Hemoglobina glicada', 'PCR', 'Proteína C reativa'):
                if an in t.cad:
                    t.muda('C3', an)
                    est = t.estado()
                    OUT.setdefault('sem_ceq', []).append(j({'analito': an, 'linha': est['B7:T8'], 'engPainel': est['engPainel'], 'H3': est['H3'], 'L4': est['L4'], 'G3': est['G3'], 'G4': est['G4']}))
        except Exception as ex:
            OUT['erros'].append(f'C: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ G: Levey-Jennings e Westgard
        cenarios = [('Glicose', 'DIMENSION 1', '8977', 2026, 10), ('Glicose', 'DIMENSION 1', '8977', 2026, None),
                    ('Ureia', 'DIMENSION 2', '8977', 2026, 9), ('Glicose', 'DIMENSION 1', '8974', 2026, 3)]
        for (an, eq, lt, ano, mes) in cenarios:
            tag = f'{an[:6]}-{eq[-1]}-{lt}-{mes or "T3"}'
            try:
                if txt(t.p.Range('L4').Value).upper() != eq:
                    t.muda('L4', eq)
                if txt(t.p.Range('C3').Value) != an:
                    t.muda('C3', an)
                if txt(t.p.Range('H3').Value) != lt:
                    t.muda('H3', "'" + lt)
                if t.p.Range('K3').Value != ano:
                    t.muda('K3', ano)
                if mes:
                    t.muda('K4', mes_lab.get(mes))
                    d0 = dt.date(ano, mes, 1)
                    d1 = dt.date(ano + (mes == 12), mes % 12 + 1, 1) - dt.timedelta(days=1)
                else:
                    t.muda('K4', tri_lab.get(3, 'T3'))
                    d0, d1 = dt.date(ano, 7, 1), dt.date(ano, 9, 30)
                xl.Calculate()
                q.ex.esperar()
                est = t.estado()
                eng = t.eng()
                eixo_lote = b.corridas(an, eq, lt)
                corl = corrida_do_lote(eixo_lote)
                jan, cort = janela(eixo_lote, d1)
                alvo = {(lt, nv): t.alvo(lt, an, nv) for nv in (1, 2)}
                # eixo do Westgard = so as corridas com resultado elegivel, compactadas (AtualizarCalc: pDeE)
                jw = [c for c in jan if c['eleg']]
                fl, z = westgard(jw, alvo)
                # 1) eixo: RUNs do motor = janela independente
                runs_eng = [x['run'] for x in eng]
                runs_ind = [c['run'] for c in jan]
                reg(f'G-EIXO-{tag}', f'{an}/{eq}/lote {lt}/{d0:%m}-{d1:%m}/{ano}: corridas do motor (Eng_Saida) = janela independente',
                    'APROVADO' if runs_eng == runs_ind else 'REPROVADO',
                    {'n_eng': len(runs_eng), 'n_ind': len(runs_ind), 'primeiras_eng': runs_eng[:3], 'primeiras_ind': runs_ind[:3],
                     'ultimas_eng': runs_eng[-3:], 'ultimas_ind': runs_ind[-3:], 'cortadas_ind': cort, 'engM1': est['eng_M1'],
                     'duplicados_mesmo_RUN_nivel': [c['dup'] for c in jan if c['dup']][:5]})
                # 2) pontos do grafico x Eng x independente; corrida do lote
                for gi, nv in ((1, 1), (2, 2)):
                    ser, info = t.series(gi)
                    pts_ind = []
                    for c in jan:
                        if nv in c['v'] and d0 <= c['dh'].date() <= d1:
                            pts_ind.append((corl[c['v'][nv][1]], c['v'][nv][0], c['v'][nv][1]))
                    alvo_xy = set((float(x), round(y, 9)) for x, y, _ in pts_ind)
                    melhor, mk = None, -1
                    for nm, xy in ser.items():
                        s_ = set((float(x), round(y, 9)) for x, y in xy)
                        k = len(s_ & alvo_xy)
                        if k > mk:
                            melhor, mk = nm, k
                    xy_dados = ser.get(melhor, [])
                    s_dados = set((float(x), round(y, 9)) for x, y in xy_dados)
                    sobra = sorted(s_dados - alvo_xy)[:6]
                    falta = sorted(alvo_xy - s_dados)[:6]
                    xs = [x for x, _, _ in pts_ind]
                    seq_ok = all(b_ - a_ == 1 for a_, b_ in zip(xs, xs[1:]))
                    reg(f'G-PONTOS-{tag}-N{nv}', f'LJ N{nv} ({an}/{eq}/{lt}): pontos do gráfico = resultados ativos do período; '
                        'X = corrida do lote (ADR-080) recalculada do histórico do lote',
                        'APROVADO' if not sobra and not falta and pts_ind and seq_ok else 'REPROVADO',
                        {'serie_dados': melhor, 'n_grafico': len(xy_dados), 'n_independente': len(pts_ind),
                         'x_primeiro_ultimo': [xs[:1], xs[-1:]], 'x_sem_buraco': seq_ok, 'sobra_no_grafico': sobra,
                         'faltam_no_grafico': falta, 'eixo': info,
                         'n_painel': est['engPainel'][nv - 1][1]})
                    # limites
                    M, S = alvo[(lt, nv)]
                    lim_r = {}
                    lim_ok = True
                    for nm, xy in ser.items():
                        nome = nm.split(':', 1)[1]
                        if nome in LIM:
                            k = KS[LIM.index(nome)]
                            ys = sorted(set(round(y, 9) for _, y in xy))
                            esperado = round(M + k * S, 9) if num(M) and num(S) else None
                            ok_l = esperado is not None and ys == [esperado]
                            xr = [min(x for x, _ in xy), max(x for x, _ in xy)] if xy else None
                            lim_r[nome] = {'y': ys, 'esperado': esperado, 'x': xr}
                            if not ok_l:
                                lim_ok = False
                    cobre = all(v['x'] and xs and v['x'][0] <= min(xs) and v['x'][1] >= max(xs) for v in lim_r.values()) if lim_r else False
                    reg(f'G-LIMITES-{tag}-N{nv}', f'LJ N{nv}: 7 linhas de limite = Média + k·DP do lote {lt} (cadastro) e cobrem os pontos',
                        'APROVADO' if lim_ok and len(lim_r) == 7 and cobre else 'REPROVADO',
                        {'media_dp_cadastro': [M, S], 'linhas': lim_r, 'cobrem_os_pontos': cobre})
                    # tabela x grafico
                    ys_g = [y for _, y in xy_dados]
                    s_g = stats(ys_g)
                    ep = est['engPainel'][nv - 1]
                    tg_ok = (s_g['n'] == ep[1] and (s_g['n'] == 0 or rel(s_g['media'], ep[2]) <= TOL)
                             and (s_g['n'] < 2 or rel(s_g['dp'], ep[3]) <= TOL))
                    reg(f'G-TABxGRAF-{tag}-N{nv}', f'Tabela do Painel (n, média, DP do N{nv}) = pontos desenhados no gráfico',
                        'APROVADO' if tg_ok else 'REPROVADO', {'grafico': s_g, 'tabela': ep[1:5]})
                # 3) Westgard: marcas do motor x independente (todos os pontos da janela; ponto = (RUN, nivel))
                nomes = ['1_3s', '2_2s', 'R_4s', '4_1s', '8x']
                pos = {c['run']: i for i, c in enumerate(jw)}
                dif = []
                comparados = 0
                cont_ind = [[0] * 5 for _ in range(2)]
                for x in eng:
                    i = pos.get(x['run'])
                    if i is None:
                        continue
                    for t_ in range(2):
                        if not x['id'][t_]:
                            continue
                        comparados += 1
                        motor = {nomes[k] for k in range(5) if num(x['flags'][t_][k]) and x['flags'][t_][k] == 1}
                        ind = fl.get((i, t_), set())
                        if motor != ind:
                            dif.append({'run': x['run'], 'nivel': t_ + 1, 'motor': sorted(motor), 'independente': sorted(ind),
                                        'z': [None if z[0][i] is None else round(z[0][i], 3), None if z[1][i] is None else round(z[1][i], 3)]})
                        if x['fil'] == 1:
                            for k in range(5):
                                if nomes[k] in ind:
                                    cont_ind[t_][k] += 1
                reg(f'G-WESTGARD-{tag}', f'Westgard ({an}/{eq}/{lt}): marcas por ponto do motor = recálculo independente (ADR-083) em toda a janela',
                    'APROVADO' if not dif and comparados else 'REPROVADO',
                    {'pontos_comparados': comparados, 'divergencias': dif[:10], 'n_divergencias': len(dif),
                     'alvo': {f'{k[0]}|N{k[1]}': v for k, v in alvo.items()}})
                cont_qc = [[est['B7:T8'][t_][12 + k] for k in range(5)] for t_ in range(2)]
                status_ind = ['REJEITADO' if sum(cont_ind[t_]) > 0 else 'OK' for t_ in range(2)]
                reg(f'G-CONTAGEM-{tag}', 'Painel "Westgard — violações no período" (M7:Q8) e status (J7:J8) = contagem independente no período',
                    'APROVADO' if cont_qc == cont_ind else 'REPROVADO',
                    {'painel': cont_qc, 'independente': cont_ind, 'status_painel': [est['B7:T8'][0][9], est['B7:T8'][1][9]],
                     'status_independente_sem_regra_1_2s': status_ind,
                     'ultima_violacao_S7_S8': [est['B7:T8'][0][18], est['B7:T8'][1][18]], 'periodo': [d0, d1]})
                OUT.setdefault('lj', {})[tag] = j({'estado': est, 'eng': eng[-40:]})
            except Exception as ex:
                OUT['erros'].append(f'G {tag}: {ex!r} {traceback.format_exc()[-1200:]}')
                reg(f'G-{tag}', f'Levey-Jennings {tag}', 'BLOQUEADO', {'erro': repr(ex)})

        OUT['estado_final'] = j(t.estado())
    finally:
        q.fechar(False)
        log('Excel encerrado (PID', OUT.get('excel_pid'), ')')
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(j(OUT), f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    executar(sys.argv[1], sys.argv[2], sys.argv[3])
