# -*- coding: utf-8 -*-
"""F4 do Usuario Simulado na HEMATOLOGIA -- periodos, lotes, filtros, graficos (Levey-Jennings / Westgard N=3) e
usabilidade. Reaproveita o calculo independente e a leitura do f4_periodos_lotes.py (Bioquimica), com o que muda na
Hematologia: 3 niveis, equipamento unico (sem L4), Painel A7:T9, Estatistica por ANO (B3/D3) e lote B4 = loteAnalise,
e o Westgard de tres niveis (1_3s, 2of3_2s WITHIN_RUN, R_4s, 3_1s N3_R1, 6x N3_R2; complementares SAME_LEVEL nao
marcam porque DetectorAtivo = 0 na tabela DETECTORES).

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    python -u f4_hema.py <copia.xlsm> <cq.pkl> <pasta_saida>

    <copia.xlsm>  COPIA da producao (shutil.copy2 para %TEMP%\\qcw\\auditoria\\<ciclo>\\), caminho ABSOLUTO.
    <cq.pkl>      tblCQ_Final da MESMA copia lida por openpyxl (pickle {'cab','rows'}), feita antes de abrir o Excel
                  (cabecalho na linha 6 da aba 'Principal - Resultados'). T00 confere os IDs contra o COM.
    <pasta_saida> recebe f4h.json (tudo o que foi lido e calculado) e as imagens exportadas pelo proprio Excel.

Blocos (sempre pelas celulas de filtro do Painel com os eventos ligados: C3 analito, H3 lote, K3/K4 ano/periodo,
G3/G4 De/Ate; na Estatistica, B3/D3 Ano De/Ate, que sao as celulas desbloqueadas da aba):
  P  periodos: todos os meses com dado (lote mais frequente do mes e, nos meses de troca, tambem o outro lote),
     trimestres (e a "media dos meses" para mostrar que o trimestre NAO e isso), semestre, "Ano inteiro", "Tudo".
  E  aba Estatistica: segue o lote do Painel? segue Ano De/Ate? (n/media/DP/CV de 28 analitos x 3 niveis contra os
     dados brutos do lote e do ano). Se ficar para tras, mEstatistica.AtualizarEstatistica (passo INTERNO, dito no
     relatorio) mostra se o calculo em si esta certo.
  L  lotes: A -> B -> H3 vazio (ADR-086) num periodo com tres lotes; periodo com lotes sobrepostos (mesmo RUN em dois
     lotes); ajuste automatico do filtro ao trocar de lote (ADR-077/078) nos dois sentidos.
  C  combinacoes: outros analitos (o equipamento e unico na Hematologia); ET do Painel em analito sem CEQ.
  G  Levey-Jennings: pontos x Eng_Saida x tblCQ_Final, corrida do lote (ADR-080), linhas de limite (ADR-078; degraus
     no modo TODOS, ADR-086), Westgard N=3 independente (ADR-083) x marcas do motor x contagens M7:Q9, tabela x grafico.
  U  usabilidade: tempos, mensagens (H4, A39, C4), K4 depois de digitar De/Ate, botoes, imagens, auditar_aba.

Calculo independente: Python puro (statistics/math) sobre as linhas brutas (PARTICIPA_ESTATISTICA = SIM, analito,
equipamento, lote, nivel, dia de DATA_HORA). Nenhuma formula do QC_INI e lida para calcular. A Media/DP do alvo do lote
e entrada do usuario (cadastro), lida por mLotes.MediaDoLote / DPDoLote.
Nenhum dado e gravado (so filtros); a copia nunca e salva. Senha: importada da suite de testes, nunca escrita aqui.
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01.
"""
import collections
import datetime as dt
import json
import os
import pickle
import statistics
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import f4_periodos_lotes as f4  # noqa: E402
from f4_periodos_lotes import (QA, SENHA, Bruto, Watch, stats, janela, corrida_do_lote, txt, num, ingenuo, dia,  # noqa: E402,F401
                               serie, rel, j, reg, OUT, log, TOL, LIM, KS)
import auditoria_abas as aa  # noqa: E402

NLV = 3
EQ = 'XN1000'
REGRAS = ['1_3s', '2of3_2s', 'R_4s', '3_1s', '6x']      # ordem dos campos 1..5 do bloco do nivel no Eng_Saida
D_MIN, D_MAX = dt.date(1900, 1, 1), dt.date(2100, 1, 1)


def fim_mes(a, m):
    return dt.date(a + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)


# ====================================================================== Westgard independente (N = 3)
def westgard_hema(eixo, alvo):
    """Hematologia (ADR-042/083, tabela DETECTORES): 1_3s |z|>3; 2of3_2s WITHIN_RUN (>= 2 niveis presentes; 2 ou mais
    do mesmo lado alem de 2s; marca os que passam); R_4s (todo par de niveis do RUN com |z1-z2|>4; marca os dois);
    3_1s N3_R1 (RUN completo, os 3 alem do mesmo 1s); 6x N3_R2 (2 RUNs completos seguidos no eixo, 6 obs do mesmo lado
    da media; marca os 3 niveis do 2o). Complementares SAME_LEVEL: Ativo = 0, nao marcam. Alerta 1_2s: |z|>2.
    eixo = corridas elegiveis compactadas (pDeE). alvo = {(lote, nivel): (M, S)}."""
    n = len(eixo)
    z = [[None] * n for _ in range(NLV)]
    for i, c in enumerate(eixo):
        for t in range(NLV):
            if (t + 1) in c['v']:
                val, _, lt = c['v'][t + 1]
                M, S = alvo.get((lt, t + 1), (None, None))
                if num(M) and num(S) and S > 0:
                    z[t][i] = (val - M) / S
    fl = collections.defaultdict(set)
    a12 = set()

    def comp(i):
        return all(z[t][i] is not None for t in range(NLV))
    for i in range(n):
        pres = [t for t in range(NLV) if z[t][i] is not None]
        for t in pres:
            if abs(z[t][i]) > 2:
                a12.add((i, t))
            if abs(z[t][i]) > 3:
                fl[(i, t)].add('1_3s')
        if len(pres) >= 2:
            for a_ in range(len(pres)):
                for b_ in range(a_ + 1, len(pres)):
                    t1, t2 = pres[a_], pres[b_]
                    if abs(z[t1][i] - z[t2][i]) > 4:
                        fl[(i, t1)].add('R_4s')
                        fl[(i, t2)].add('R_4s')
            ac = [t for t in pres if z[t][i] > 2]
            ab = [t for t in pres if z[t][i] < -2]
            if len(ac) >= 2:
                for t in ac:
                    fl[(i, t)].add('2of3_2s')
            elif len(ab) >= 2:
                for t in ab:
                    fl[(i, t)].add('2of3_2s')
        if comp(i):
            if all(z[t][i] > 1 for t in range(NLV)) or all(z[t][i] < -1 for t in range(NLV)):
                for t in range(NLV):
                    fl[(i, t)].add('3_1s')
        if i >= 1 and comp(i) and comp(i - 1):
            obs = [z[t][k] for k in (i - 1, i) for t in range(NLV)]
            if all(o > 0 for o in obs) or all(o < 0 for o in obs):
                for t in range(NLV):
                    fl[(i, t)].add('6x')
    return fl, z, a12


# ====================================================================== leitura do QC_INI (copia)
class TelaH(f4.Tela):
    def estado(self):
        p, e = self.p, self.e
        d = {}
        for a in ('B3', 'C3', 'C4', 'G3', 'G4', 'H3', 'H4', 'K3', 'K4', 'O3', 'O4', 'Q4', 'T4', 'A39'):
            v = p.Range(a).Value
            d[a] = ingenuo(v) if isinstance(v, dt.datetime) else v
        d['A7:T9'] = [[ingenuo(x) if isinstance(x, dt.datetime) else x for x in l] for l in p.Range('A7:T9').Value]
        d['engPainel'] = [list(l) for l in self.q.nome('engPainel').Value]
        for a in ('C1', 'E1', 'G1', 'I1', 'K1', 'M1', 'O1', 'Q1', 'S1', 'W1', 'Y1', 'AA1', 'AD1'):
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
                        'val': [r[24 + t] for t in range(NLV)],
                        'id': [txt(r[38 + t]) for t in range(NLV)],
                        'dh': [ingenuo(r[50 + t]) if isinstance(r[50 + t], dt.datetime) else r[50 + t] for t in range(NLV)],
                        'cor': [r[62 + t] for t in range(NLV)],
                        'alvo': [(r[74 + 2 * t], r[75 + 2 * t]) for t in range(NLV)],
                        'flags': [list(r[2 + t * 7:2 + t * 7 + 7]) for t in range(NLV)]})
        return out

    def estat_linhas(self):
        v = self.est.Range('A14:F133').Value
        out = {}
        for r in v:
            an = txt(r[0])
            if an and num(r[1]):
                out[(an, int(r[1]))] = list(r[2:6])
        filtro = {}
        for a in ('B3', 'D3', 'B4', 'K5', 'N4', 'P4', 'R4'):
            x = self.est.Range(a).Value
            filtro[a] = ingenuo(x) if isinstance(x, dt.datetime) else x
        return out, filtro

    def formas(self, ws):
        out = []
        for s in ws.Shapes:
            try:
                acao = s.OnAction
            except Exception:
                acao = None
            try:
                texto = s.TextFrame2.TextRange.Text if s.HasTextFrame else None
            except Exception:
                texto = None
            if acao or texto:
                out.append({'nome': s.Name, 'texto': (texto or '')[:80], 'acao': acao, 'visivel': bool(s.Visible),
                            'top': round(s.Top, 1), 'left': round(s.Left, 1)})
        return out


def mesmo(a, b):
    if b is None:
        return a in (None, '')
    r = rel(a, b)
    return r is not None and r <= TOL


def cmp_painel(tela, est, an, lote, d0, d1, tid, desc, extra=None, todos=False):
    """engPainel (n, media, DP, CV) x independente dos dados brutos do periodo (3 niveis). No modo TODOS so o n
    (media/DP/CV/Bias/ET/Sigma ficam vazios por decisao do ADR-086)."""
    linhas, pior, ok = [], 0.0, True
    for t in range(NLV):
        ind = tela.b.stats_periodo(an, EQ, None if todos else lote, t + 1, d0, d1)
        ep = est['engPainel'][t]
        qc = {'n': ep[1], 'media': ep[2], 'dp': ep[3], 'cv': ep[4]}
        difs = {}
        for k in (('n',) if todos else ('n', 'media', 'dp', 'cv')):
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
    evid = {'analito': an, 'lote': lote, 'de': d0, 'ate': d1, 'G3': est['G3'], 'G4': est['G4'], 'K4': est['K4'],
            'H4': est['H4'], 'engCortadas_M1': est['eng_M1'], 'janela_I1': est['eng_I1'], 'linhas': linhas,
            'maior_dif_rel': pior}
    if extra:
        evid.update(extra)
    reg(tid, desc, 'APROVADO' if ok else 'REPROVADO', evid)
    return ok, linhas


def cmp_eng_escopo(tela, an, lote, d0, d1, todos=False):
    eng = tela.eng()
    cq = {e['id']: e for e in tela.b.elig}
    fora, faltam, vistos = [], [], set()
    for x in eng:
        if x['fil'] != 1:
            continue
        for t in range(NLV):
            i = x['id'][t]
            if not i:
                continue
            vistos.add(i)
            e = cq.get(i)
            if (e is None or e['an'] != an or e['eq'] != EQ or (not todos and e['lote'] != lote)
                    or e['nv'] != t + 1 or e['dh'].date() < d0 or e['dh'].date() > d1):
                fora.append({'id': i, 'nivel': t + 1, 'cq': e and {k: e[k] for k in ('an', 'eq', 'lote', 'nv', 'dh')}})
            elif not (num(x['val'][t]) and abs(x['val'][t] - e['val']) < 1e-12):
                fora.append({'id': i, 'valor_eng': x['val'][t], 'valor_cq': e['val']})
    for t in range(NLV):
        for e in tela.b.sel(an, EQ, None if todos else lote, t + 1, d0, d1):
            if e['id'] not in vistos:
                faltam.append({'id': e['id'], 'nivel': t + 1, 'lote': e['lote'], 'dh': e['dh']})
    return eng, fora, faltam


def cmp_estat(tela, b, lote, a0, a1, tid, desc, extra=None, estado='auto'):
    """Estatistica C:F (n, media, DP, CV de cada analito x nivel) x independente (lote, ANO de DATA_HORA em a0..a1).
    Devolve tambem de qual lote os numeros parecem ser, quando divergem (para dizer se a aba ficou para tras)."""
    linhas, filtro = tela.estat_linhas()
    d0, d1 = dt.date(a0, 1, 1), dt.date(a1, 12, 31)
    difs, comp, pior = [], 0, 0.0
    for (an, nv), qcv in linhas.items():
        ind = b.stats_periodo(an, EQ, lote, nv, d0, d1)
        for k, a in zip(('n', 'media', 'dp', 'cv'), qcv):
            bb = ind[k]
            if k == 'n':
                bb = bb or 0
            if bb is None:
                if a not in (None, ''):
                    difs.append({'an': an, 'nv': nv, 'k': k, 'qc': a, 'ind': bb})
                continue
            comp += 1
            r = rel(a, bb) if num(a) else None
            if r is None or r > TOL:
                difs.append({'an': an, 'nv': nv, 'k': k, 'qc': a, 'ind': bb})
            else:
                pior = max(pior, r)
    wbc = [linhas.get(('WBC', k)) for k in (1, 2, 3)]
    evid = {'filtro': filtro, 'lote_esperado': lote, 'anos': [a0, a1], 'comparacoes': comp,
            'n_divergencias': len(difs), 'divergencias': difs[:9], 'maior_dif_rel': pior, 'WBC_N1_N3': wbc}
    if extra:
        evid.update(extra)
    st = ('APROVADO' if not difs and comp > 0 else 'REPROVADO') if estado == 'auto' else estado
    reg(tid, desc, st, evid)
    return not difs, linhas


# ====================================================================== roteiro
def executar(copia, pkl, pasta):
    os.makedirs(pasta, exist_ok=True)
    d = pickle.load(open(pkl, 'rb'))
    b = Bruto(d['rows'])
    q = QA('Hematologia', copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = 'Hematologia'
    try:
        t = TelaH(q, b)
        xl = t.xl
        lo = q.lo('tblCQ_Final')
        ids_com = [txt(x[0]) for x in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        ids_pkl = [txt(r['ID_REGISTRO']) for r in d['rows']]
        reg('T00', 'Pré-condição: tblCQ_Final do COM = leitura openpyxl da mesma cópia (base do cálculo independente)',
            'APROVADO' if ids_com == ids_pkl else 'REPROVADO', {'linhas_com': len(ids_com), 'linhas_pkl': len(ids_pkl)})
        q.run('mSeguranca.ReprotectAll', teto=120)
        OUT['painel_protegido'] = bool(t.p.ProtectContents)
        OUT['estat_protegida'] = bool(t.est.ProtectContents)
        listas = {}
        for c_ in ('C3', 'H3', 'K3', 'K4'):
            listas[c_] = t.lista(t.p, c_)
        OUT['listas'] = j(listas)
        l_k4 = listas['K4'][1]
        OUT['celulas'] = {}
        for ws_, cels in ((t.p, ('B3', 'C3', 'G3', 'G4', 'H3', 'K3', 'K4')), (t.est, ('B3', 'D3', 'B4', 'N4', 'P4', 'R4', 'L11'))):
            for c_ in cels:
                r_ = ws_.Range(c_)
                OUT['celulas'][f'{ws_.Name}!{c_}'] = {'formula': str(r_.Formula)[:120], 'bloqueada': bool(r_.Locked)}
        OUT['formas_painel'] = t.formas(t.p)
        OUT['formas_estat'] = t.formas(t.est)
        OUT['estado_inicial'] = j(t.estado())
        OUT['estat_inicial'] = j(t.estat_linhas()[1])
        lote_uso = txt(q.nome('loteAtivo').Value)
        OUT['lote_em_uso'] = lote_uso

        AN = 'WBC'
        t.muda('C3', AN)
        OUT['base'] = j(t.estado())

        def lotes_mes(an, a, m):
            return collections.Counter(e['lote'] for e in b.sel(an, EQ, None) if e['dh'].year == a and e['dh'].month == m)

        def poe(cel, valor, como=None):
            atual = t.p.Range(cel).Value
            if cel == 'H3':
                if txt(atual) != txt(valor):
                    return t.muda('H3', None if valor is None else "'" + valor)
                return 0.0
            if atual != valor:
                return t.muda(cel, valor)
            return 0.0

        # ------------------------------------------------ P: periodos
        meses = sorted({(e['dh'].year, e['dh'].month) for e in b.sel(AN, EQ, None)})
        MES_LAB = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        for (ano, mes) in meses:
            cnt = lotes_mes(AN, ano, mes)
            for k_, (lt, nlt) in enumerate(cnt.most_common()):
                try:
                    poe('H3', lt)
                    poe('K3', ano)
                    dtc = t.muda('K4', MES_LAB[mes - 1])
                    est = t.estado()
                    d0, d1 = dt.date(ano, mes, 1), fim_mes(ano, mes)
                    g_ok = dia(est['G3']) == d0 and dia(est['G4']) == d1
                    cmp_painel(t, est, AN, lt, d0, d1, f'P-M{ano}{mes:02d}-{lt}',
                               f'Mês {mes:02d}/{ano} (K3/K4) — WBC lote {lt}' + (' (2º lote do mês)' if k_ else '') +
                               ': Painel = dados brutos do mês e do lote',
                               {'G3G4_corretos': g_ok, 'tempo_s': dtc, 'lotes_no_mes': dict(cnt), 'A39': est['A39'],
                                'status_J': [est['A7:T9'][k][9] for k in range(3)]})
                    if not g_ok:
                        reg(f'P-G{ano}{mes:02d}', f'G3/G4 do mês {mes:02d}/{ano}', 'REPROVADO',
                            {'G3': est['G3'], 'G4': est['G4'], 'esperado': [d0, d1]})
                except Exception as ex:
                    OUT['erros'].append(f'P {ano}-{mes} {lt}: {ex!r} {traceback.format_exc()[-800:]}')
                    reg(f'P-M{ano}{mes:02d}-{lt}', f'Mês {mes:02d}/{ano} lote {lt}', 'BLOQUEADO', {'erro': repr(ex)})

        # trimestres: lote mais frequente do trimestre
        tris = sorted({(e['dh'].year, (e['dh'].month - 1) // 3 + 1) for e in b.sel(AN, EQ, None)})
        for (ano, k) in tris:
            c = collections.Counter(e['lote'] for e in b.sel(AN, EQ, None, None, dt.date(ano, 3 * k - 2, 1), fim_mes(ano, 3 * k)))
            lt = c.most_common(1)[0][0]
            try:
                poe('H3', lt)
                poe('K3', ano)
                dtc = t.muda('K4', f'T{k}')
                est = t.estado()
                d0, d1 = dt.date(ano, 3 * k - 2, 1), fim_mes(ano, 3 * k)
                alt = []
                for t_ in range(NLV):
                    mm, cvs = [], []
                    for m in range(3 * k - 2, 3 * k + 1):
                        s = b.stats_periodo(AN, EQ, lt, t_ + 1, dt.date(ano, m, 1), fim_mes(ano, m))
                        if s['n']:
                            mm.append(s['media'])
                        if s['cv'] is not None:
                            cvs.append(s['cv'])
                    alt.append({'nivel': t_ + 1, 'media_das_medias_mensais': statistics.fmean(mm) if mm else None,
                                'media_dos_CV_mensais': statistics.fmean(cvs) if cvs else None, 'meses_com_dado': len(mm)})
                g_ok = dia(est['G3']) == d0 and dia(est['G4']) == d1
                cmp_painel(t, est, AN, lt, d0, d1, f'P-T{ano}{k}',
                           f'Trimestre T{k}/{ano} — WBC lote {lt}: Painel = dados brutos do trimestre (não média dos meses)',
                           {'G3G4_corretos': g_ok, 'tempo_s': dtc, 'lotes_no_trimestre': dict(c),
                            'alternativa_media_dos_meses': alt})
            except Exception as ex:
                OUT['erros'].append(f'P tri {ano}-{k}: {ex!r} {traceback.format_exc()[-800:]}')
                reg(f'P-T{ano}{k}', f'Trimestre T{k}/{ano}', 'BLOQUEADO', {'erro': repr(ex)})

        # semestres, ano inteiro e tudo
        longos = [(2026, 'S1', dt.date(2026, 1, 1), dt.date(2026, 6, 30)),
                  (2026, 'S2', dt.date(2026, 7, 1), dt.date(2026, 12, 31)),
                  (2025, 'Ano inteiro', dt.date(2025, 1, 1), dt.date(2025, 12, 31)),
                  (2026, 'Ano inteiro', dt.date(2026, 1, 1), dt.date(2026, 12, 31)),
                  (2026, 'Tudo (sem filtro)', None, None)]
        for (ano, lab, d0, d1) in longos:
            dd0, dd1 = d0 or D_MIN, d1 or D_MAX
            c = collections.Counter(e['lote'] for e in b.sel(AN, EQ, None, None, dd0, dd1))
            lt = c.most_common(1)[0][0]
            try:
                poe('H3', lt)
                poe('K3', ano)
                dtc = t.muda('K4', lab)
                est = t.estado()
                cmp_painel(t, est, AN, lt, dd0, dd1, f'P-L{lab[:4]}{ano}',
                           f'Período "{lab}" {ano} — WBC lote {lt}: Painel = todos os dados brutos do período e do lote',
                           {'tempo_s': dtc, 'lotes_no_periodo': dict(c), 'A39': est['A39']})
            except Exception as ex:
                OUT['erros'].append(f'P {lab}: {ex!r} {traceback.format_exc()[-800:]}')

        # TODOS os lotes num periodo longo: a janela de 180 corridas (ADR-086 decisao 6) e o aviso
        for (ano, lab, d0, d1) in [(2026, 'Ano inteiro', dt.date(2026, 1, 1), dt.date(2026, 12, 31)),
                                   (2026, 'Tudo (sem filtro)', None, None)]:
            dd0, dd1 = d0 or D_MIN, d1 or D_MAX
            try:
                poe('K3', ano)
                t.muda('K4', lab)
                dtc = poe('H3', None)
                est = t.estado()
                eixo = b.corridas(AN, EQ, None, todos=True)
                jan, cort = janela(eixo, d1)
                win = [len([c_ for c_ in jan if (t_ + 1) in c_['v'] and dd0 <= c_['dh'].date() <= dd1]) for t_ in range(NLV)]
                todas = [len(b.sel(AN, EQ, None, t_ + 1, dd0, dd1)) for t_ in range(NLV)]
                n_qc = [est['engPainel'][k][1] for k in range(NLV)]
                corr_per = len([c_ for c_ in eixo if dd0 <= c_['dh'].date() <= dd1])
                reg(f'P-TODOS-{lab[:4]}', f'H3 vazio + "{lab}" {ano} (WBC): n do Painel = todos os resultados do período? '
                    f'(o modo TODOS usa no máximo as 180 corridas mais recentes, ADR-086 dec. 6)',
                    'APROVADO' if n_qc == todas else 'REPROVADO',
                    {'n_qc': n_qc, 'n_todos_os_dados_do_periodo': todas, 'n_so_da_janela_180': win,
                     'corridas_RUN_LOTE_no_periodo': corr_per, 'corridas_fora_da_janela': cort, 'engM1': est['eng_M1'],
                     'H4': est['H4'], 'A39': est['A39'], 'linha_painel': est['A7:T9'], 'tempo_s': dtc,
                     'lotes_no_periodo_O1': est['eng_O1']})
            except Exception as ex:
                OUT['erros'].append(f'P TODOS {lab}: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ E: aba Estatistica acompanha o Painel?
        try:
            poe('H3', '522611')
            poe('K3', 2025)
            t.muda('K4', 'Tudo (sem filtro)')
            yb3, yd3 = t.est.Range('B3').Value, t.est.Range('D3').Value
            a0, a1 = int(yb3), int(yd3)
            cmp_estat(t, b, '522611', a0, max(a0, a1), 'E-BASE-522611',
                      f'Estatística no estado de abertura (lote 522611, Ano {a0}–{a1}): n/média/DP/CV = dados brutos')
            dtc = poe('H3', '619611')
            ok, _ = cmp_estat(t, b, '619611', a0, max(a0, a1), 'E-LOTE-619611',
                              f'Usuário troca o lote no Painel (H3 = 619611): a Estatística (B4 refletido) recalcula para o lote novo '
                              f'(Ano {a0}–{a1})?', {'tempo_s': dtc,
                                                   'se_fosse_o_lote_anterior_522611': [b.stats_periodo('WBC', EQ, '522611', k, dt.date(a0, 1, 1), dt.date(a1, 12, 31)) for k in (1, 2, 3)]})
            dtc1 = t.muda('B3', 2026, aba=t.est)
            dtc2 = t.muda('D3', 2026, aba=t.est)
            xl.Calculate()
            q.ex.esperar()
            ok2, _ = cmp_estat(t, b, '619611', 2026, 2026, 'E-ANO-2026-619611',
                               'Usuário muda Ano De/Até da Estatística para 2026/2026 (lote 619611): a aba recalcula?',
                               {'tempos_s': [dtc1, dtc2]})
            # passo INTERNO (nao e o caminho do usuario): o motor completo da Estatistica
            t0 = time.time()
            q.run('mEstatistica.AtualizarEstatistica', teto=900)
            q.ex.esperar()
            cmp_estat(t, b, '619611', 2026, 2026, 'E-INTERNO-619611',
                      'Depois de mEstatistica.AtualizarEstatistica (passo interno do ATUALIZAR): Estatística = dados brutos de 619611/2026',
                      {'tempo_s': round(time.time() - t0, 2)})
            # e o Painel continua coerente com a aba nas colunas que le dela (G Bias, I Sigma)
            est = t.estado()
            OUT['painel_apos_estat_interno'] = j(est['A7:T9'])
            # ano invertido (Ate < De): a regra do codigo diz Ate = De
            t.muda('B3', 2026, aba=t.est)
            t.muda('D3', 2025, aba=t.est)
            q.run('mEstatistica.AtualizarEstatistica', teto=900)
            cmp_estat(t, b, '619611', 2026, 2026, 'E-ANO-INVERTIDO',
                      'Ano Até (2025) < Ano De (2026): vale só o Ano De (regra do EstatBasica); a tela avisa?')
            OUT['estat_ano_invertido_tela'] = j(t.estat_linhas()[1])
            t.muda('D3', 2026, aba=t.est)
        except Exception as ex:
            OUT['erros'].append(f'E: {ex!r} {traceback.format_exc()[-1500:]}')

        # ------------------------------------------------ L: lote A -> lote B -> H3 vazio (fev-mar/2026: tres lotes)
        try:
            poe('K3', 2026)
            t.muda('G3', serie(dt.date(2026, 2, 1)))
            t.muda('G4', serie(dt.date(2026, 3, 31)))
            d0, d1 = dt.date(2026, 2, 1), dt.date(2026, 3, 31)
            OUT['K4_apos_digitar_G3G4'] = t.p.Range('K4').Value
            cq = {e['id']: e for e in b.elig}
            for lt in ('536411', '602911', None):
                dtc = poe('H3', lt)
                est = t.estado()
                eng, fora, faltam = cmp_eng_escopo(t, AN, lt, d0, d1, todos=lt is None)
                lotes_pts = collections.Counter()
                for x in eng:
                    if x['fil'] == 1:
                        for t_ in range(NLV):
                            if x['id'][t_] in cq:
                                lotes_pts[cq[x['id'][t_]]['lote']] += 1
                linhas_e, filtro_e = t.estat_linhas()
                est_info = {'Estat_B4': filtro_e['B4'], 'Estat_WBC': [linhas_e.get(('WBC', k)) for k in (1, 2, 3)]}
                if lt is None:
                    ind = [len(b.sel(AN, EQ, None, t_ + 1, d0, d1)) for t_ in range(NLV)]
                    qc_n = [est['engPainel'][t_][1] for t_ in range(NLV)]
                    vazios = all(est['engPainel'][t_][k] in (None, '') for t_ in range(NLV) for k in (2, 3, 4, 6, 7, 8))
                    tela_vazia = all(est['A7:T9'][t_][k] in (None, '') for t_ in range(NLV) for k in (2, 3, 4, 6, 7, 8))
                    ok = (est['H3'] in (None, '') and est['eng_AD1'] is True and qc_n == ind and not fora and not faltam and vazios)
                    reg('L-TODOS', 'H3 vazio (ADR-086), WBC 01/02–31/03/2026: motor com todos os lotes do período; n = todos os '
                        'lotes; média/DP/CV/Bias/ET/Sigma vazios no motor',
                        'APROVADO' if ok else 'REPROVADO',
                        {'H4': est['H4'], 'engTodos': est['eng_AD1'], 'n_qc': qc_n, 'n_independente': ind,
                         'pontos_por_lote': dict(lotes_pts), 'fora_do_escopo': fora[:6], 'faltando': faltam[:6],
                         'engPainel_vazio': vazios, 'tela_G_H_I_vazia': tela_vazia, 'linha_painel': est['A7:T9'],
                         'A39': est['A39'], 'C4': est['C4'], 'O1': est['eng_O1'], 'tempo_s': dtc, **est_info})
                    y2 = []
                    for x in eng:
                        for t_ in range(NLV):
                            i = x['id'][t_]
                            if i and i in cq:
                                m_, s_ = t.alvo(cq[i]['lote'], AN, t_ + 1)
                                g = x['alvo'][t_]
                                if num(m_) and num(s_) and s_ > 0:
                                    if not (num(g[0]) and num(g[1]) and abs(g[0] - m_) < 1e-9 and abs(g[1] - s_) < 1e-9):
                                        y2.append({'id': i, 'lote': cq[i]['lote'], 'eng': g, 'cadastro': (m_, s_)})
                    reg('L-TODOS-ALVO', 'H3 vazio: cada ponto com a Média/DP do SEU lote (Eng_Saida BW..CB)',
                        'APROVADO' if not y2 else 'REPROVADO', {'divergencias': y2[:8]})
                else:
                    cmp_painel(t, est, AN, lt, d0, d1, f'L-{lt}',
                               f'Lote {lt} em H3, 01/02–31/03/2026 (período com 3 lotes): estatística só do lote',
                               {'tempo_s': dtc, 'pontos_por_lote_no_motor': dict(lotes_pts), **est_info})
                    reg(f'L-{lt}-MOTOR', f'Lote {lt}: o MOTOR (Eng_Saida) só tem pontos do lote/analito/nível/período, e todos eles',
                        'APROVADO' if not fora and not faltam else 'REPROVADO',
                        {'fora_do_escopo': fora[:6], 'faltando': faltam[:6], 'pontos_por_lote': dict(lotes_pts)})
        except Exception as ex:
            OUT['erros'].append(f'L: {ex!r} {traceback.format_exc()[-1500:]}')

        # periodo com lotes SOBREPOSTOS (528211 e 533811 em 22-23/12/2025): RUN|LOTE no modo TODOS
        try:
            t.muda('G3', serie(dt.date(2025, 12, 15)))
            t.muda('G4', serie(dt.date(2025, 12, 31)))
            d0, d1 = dt.date(2025, 12, 15), dt.date(2025, 12, 31)
            dtc = poe('H3', None)
            est = t.estado()
            eng, fora, faltam = cmp_eng_escopo(t, AN, None, d0, d1, todos=True)
            cq = {e['id']: e for e in b.elig}
            runs_dois = collections.defaultdict(set)
            for e in b.sel(AN, EQ, None, None, d0, d1):
                runs_dois[e['run']].add(e['lote'])
            comp2 = {r: sorted(v) for r, v in runs_dois.items() if len(v) > 1}
            slots_run = collections.Counter(x['run'] for x in eng if x['fil'] == 1)
            ind = [len(b.sel(AN, EQ, None, t_ + 1, d0, d1)) for t_ in range(NLV)]
            qc_n = [est['engPainel'][t_][1] for t_ in range(NLV)]
            reg('L-SOBREPOSTOS', 'H3 vazio, WBC 15–31/12/2025 (528211 e 533811 em paralelo): o mesmo RUN em dois lotes vira '
                'duas corridas (RUN|LOTE) e nenhum resultado some',
                'APROVADO' if (not fora and not faltam and qc_n == ind and all(slots_run[r] == len(v) for r, v in comp2.items())) else 'REPROVADO',
                {'runs_em_dois_lotes': comp2, 'slots_no_motor_desses_runs': {r: slots_run[r] for r in comp2},
                 'n_qc': qc_n, 'n_independente': ind, 'fora': fora[:5], 'faltam': faltam[:5], 'tempo_s': dtc})
            # corrida do lote atravessando os lotes (eixo X do modo TODOS)
            eixo_t = b.corridas(AN, EQ, None, todos=True)
            corl = corrida_do_lote(eixo_t, NLV)
            dif = []
            for x in eng:
                for t_ in range(NLV):
                    i = x['id'][t_]
                    if i and corl.get(i) != x['cor'][t_]:
                        dif.append({'id': i, 'nivel': t_ + 1, 'motor': x['cor'][t_], 'independente': corl.get(i)})
            reg('L-SOBREPOSTOS-X', 'Modo TODOS: X = corrida 1, 2, 3... por nível em ordem cronológica atravessando os lotes (ADR-086)',
                'APROVADO' if not dif else 'REPROVADO', {'divergencias': dif[:8], 'n_div': len(dif)})
            OUT['sobrepostos_eng'] = j([x for x in eng if x['run'] in comp2])
        except Exception as ex:
            OUT['erros'].append(f'L-SOBREP: {ex!r} {traceback.format_exc()[-1200:]}')

        # ajuste automatico do filtro ao trocar de lote (ADR-077/078), nos dois sentidos
        try:
            poe('H3', '619611')
            poe('K3', 2026)
            t.muda('K4', 'Set')
            a1_ = t.estado()
            dtc1 = poe('H3', lote_uso)                 # lote em uso nao tem ponto em set/2026
            a2_ = t.estado()
            dtc2 = poe('H3', '619611')
            a3_ = t.estado()
            reg('L-AJUSTE-1', f'Lote corrente 619611 em Set/2026 → lote em uso {lote_uso} (sem ponto no filtro) → volta a 619611',
                'APROVADO' if (a2_['G3'] != a1_['G3'] or a2_['G4'] != a1_['G4']) and txt(a2_['H4']) else 'REPROVADO',
                {'antes': {k: a1_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'no_lote_em_uso': {k: a2_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'volta': {k: a3_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')}, 'tempos_s': [dtc1, dtc2],
                 'n_no_lote_em_uso': [a2_['engPainel'][k][1] for k in range(NLV)]})
            poe('H3', lote_uso)
            poe('K3', 2025)
            t.muda('K4', 'Out')
            b1_ = t.estado()
            dtc1 = poe('H3', '619611')                 # 619611 nao tem ponto em out/2025
            b2_ = t.estado()
            dtc2 = poe('H3', lote_uso)
            b3_ = t.estado()
            volta_ok = b3_['K4'] == b1_['K4'] and b3_['G3'] == b1_['G3'] and b3_['G4'] == b1_['G4']
            reg('L-AJUSTE-2', f'Lote em uso {lote_uso} em Out/2025 → 619611 (sem ponto) → volta ao lote em uso: o filtro de antes volta '
                '(GUIA §4b / ADR-078)', 'APROVADO' if volta_ok else 'REPROVADO',
                {'antes': {k: b1_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'no_619611': {k: b2_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')},
                 'volta': {k: b3_[k] for k in ('G3', 'G4', 'K3', 'K4', 'H4', 'A39')}, 'tempos_s': [dtc1, dtc2]})
        except Exception as ex:
            OUT['erros'].append(f'L-AJUSTE: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ C: outros analitos (lote 619611, set/2026)
        try:
            poe('H3', '619611')
            poe('K3', 2026)
            t.muda('K4', 'Set')
            d0, d1 = dt.date(2026, 9, 1), dt.date(2026, 9, 30)
            for an in ('RBC', 'HGB', 'PLT', 'MCV', 'NEUT%', 'IG%', 'RET%', 'RDW-SD'):
                if an not in t.cad:
                    continue
                dtc = t.muda('C3', an)
                est = t.estado()
                cmp_painel(t, est, an, '619611', d0, d1, f'C-{an}', f'{an} / lote 619611 / set-2026: Painel = dados brutos',
                           {'tempo_s': dtc, 'C3': est['C3'], 'eng_analito': est['eng_C1'],
                            'linha_painel_F_a_J': [est['A7:T9'][k][5:10] for k in range(NLV)]})
                eng, fora, faltam = cmp_eng_escopo(t, an, '619611', d0, d1)
                reg(f'C-{an}-MOTOR', f'{an}: pontos do motor = só desse analito/lote, nível no lugar certo',
                    'APROVADO' if not fora and not faltam else 'REPROVADO', {'fora': fora[:6], 'faltam': faltam[:6]})
                if an in ('RET%', 'RDW-SD', 'IG%'):
                    OUT.setdefault('sem_ceq', []).append(j({'analito': an, 'linha': est['A7:T9'], 'engPainel': est['engPainel']}))
        except Exception as ex:
            OUT['erros'].append(f'C: {ex!r} {traceback.format_exc()[-800:]}')

        # ------------------------------------------------ G: Levey-Jennings e Westgard
        cenarios = [('WBC', '619611', 2026, 'Set'), ('WBC', '614011', 2026, 'T3'), ('PLT', '608511', 2026, 'Mai'),
                    ('HGB', '533811', 2026, 'Jan'), ('RBC', '522611', 2025, 'Tudo (sem filtro)')]
        for (an, lt, ano, per) in cenarios:
            tag = f'{an}-{lt}-{per[:4]}'
            try:
                if txt(t.p.Range('C3').Value) != an:
                    t.muda('C3', an)
                poe('H3', lt)
                poe('K3', ano)
                t.muda('K4', per)
                xl.Calculate()
                q.ex.esperar()
                est = t.estado()
                d0 = dia(est['G3']) or D_MIN
                d1 = dia(est['G4']) or D_MAX
                eng = t.eng()
                eixo_lote = b.corridas(an, EQ, lt)
                corl = corrida_do_lote(eixo_lote, NLV)
                jan, cort = janela(eixo_lote, None if d1 == D_MAX else d1)
                alvo = {(lt, nv): t.alvo(lt, an, nv) for nv in (1, 2, 3)}
                jw = [c for c in jan if c['eleg']]
                fl, z, a12 = westgard_hema(jw, alvo)
                runs_eng = [x['run'] for x in eng]
                runs_ind = [c['run'] for c in jan]
                reg(f'G-EIXO-{tag}', f'{an}/lote {lt}/{per} {ano}: corridas do motor (Eng_Saida) = janela independente',
                    'APROVADO' if runs_eng == runs_ind else 'REPROVADO',
                    {'n_eng': len(runs_eng), 'n_ind': len(runs_ind), 'primeiras': [runs_eng[:3], runs_ind[:3]],
                     'ultimas': [runs_eng[-3:], runs_ind[-3:]], 'cortadas_ind': cort, 'engM1': est['eng_M1'],
                     'duplicados_mesmo_RUN_nivel': [c['dup'] for c in jan if c['dup']][:5], 'periodo': [d0, d1]})
                for nv in (1, 2, 3):
                    ser, info = t.series(nv)
                    pts_ind = [(corl[c['v'][nv][1]], c['v'][nv][0], c['v'][nv][1]) for c in jan
                               if nv in c['v'] and d0 <= c['dh'].date() <= d1]
                    alvo_xy = set((float(x), round(y, 9)) for x, y, _ in pts_ind)
                    melhor, mk = None, -1
                    for nm, xy in ser.items():
                        k = len(set((float(x), round(y, 9)) for x, y in xy) & alvo_xy)
                        if k > mk:
                            melhor, mk = nm, k
                    xy_dados = ser.get(melhor, [])
                    s_dados = set((float(x), round(y, 9)) for x, y in xy_dados)
                    sobra, falta = sorted(s_dados - alvo_xy)[:6], sorted(alvo_xy - s_dados)[:6]
                    xs = [x for x, _, _ in pts_ind]
                    seq_ok = all(b_ - a_ == 1 for a_, b_ in zip(xs, xs[1:]))
                    eng_cor_ok = all(x['cor'][nv - 1] == corl.get(x['id'][nv - 1]) for x in eng if x['id'][nv - 1])
                    reg(f'G-PONTOS-{tag}-N{nv}', f'LJ N{nv} ({an}/{lt}/{per}): pontos do gráfico = resultados ativos do período; '
                        'X = corrida do lote (ADR-080) recalculada do histórico do lote',
                        'APROVADO' if not sobra and not falta and pts_ind and seq_ok and eng_cor_ok else 'REPROVADO',
                        {'serie_dados': melhor, 'n_grafico': len(xy_dados), 'n_independente': len(pts_ind),
                         'x_primeiro_ultimo': [xs[:1], xs[-1:]], 'x_sem_buraco': seq_ok, 'eng_BK_BM_ok': eng_cor_ok,
                         'sobra_no_grafico': sobra, 'faltam_no_grafico': falta, 'eixo': info,
                         'n_painel': est['engPainel'][nv - 1][1]})
                    M, S = alvo[(lt, nv)]
                    lim_r, lim_ok = {}, True
                    for nm, xy in ser.items():
                        nome = nm.split(':', 1)[1]
                        if nome in LIM:
                            k = KS[LIM.index(nome)]
                            ys = sorted(set(round(y, 9) for _, y in xy))
                            esperado = round(M + k * S, 9) if num(M) and num(S) else None
                            ok_l = esperado is not None and ys == [esperado]
                            xr = [min(x for x, _ in xy), max(x for x, _ in xy)] if xy else None
                            lim_r[nome] = {'y': ys, 'esperado': esperado, 'x': xr}
                            lim_ok = lim_ok and ok_l
                    cobre = all(v['x'] and xs and v['x'][0] <= min(xs) and v['x'][1] >= max(xs) for v in lim_r.values()) if lim_r else False
                    reg(f'G-LIMITES-{tag}-N{nv}', f'LJ N{nv}: 7 linhas de limite = Média + k·DP do lote {lt} (cadastro) e cobrem os pontos',
                        'APROVADO' if lim_ok and len(lim_r) == 7 and cobre else 'REPROVADO',
                        {'media_dp_cadastro': [M, S], 'linhas': lim_r, 'cobrem_os_pontos': cobre})
                    s_g = stats([y for _, y in xy_dados])
                    ep = est['engPainel'][nv - 1]
                    tg_ok = (s_g['n'] == ep[1] and (s_g['n'] == 0 or rel(s_g['media'], ep[2]) <= TOL)
                             and (s_g['n'] < 2 or rel(s_g['dp'], ep[3]) <= TOL))
                    reg(f'G-TABxGRAF-{tag}-N{nv}', f'Tabela do Painel (n, média, DP do N{nv}) = pontos desenhados no gráfico',
                        'APROVADO' if tg_ok else 'REPROVADO', {'grafico': s_g, 'tabela': ep[1:5]})
                # Westgard: marcas por ponto
                pos = {c['run']: i for i, c in enumerate(jw)}
                dif, comparados = [], 0
                cont_ind = [[0] * 5 for _ in range(NLV)]
                rej_ind = [False] * NLV
                for x in eng:
                    i = pos.get(x['run'])
                    if i is None:
                        continue
                    for t_ in range(NLV):
                        if not x['id'][t_]:
                            continue
                        comparados += 1
                        motor = {REGRAS[k] for k in range(5) if num(x['flags'][t_][k]) and x['flags'][t_][k] == 1}
                        ind = fl.get((i, t_), set())
                        a12m = num(x['flags'][t_][5]) and x['flags'][t_][5] == 1
                        stm = x['flags'][t_][6]
                        st_i = 'REJEITADO' if ind else 'OK'
                        if motor != ind or a12m != ((i, t_) in a12) or stm != st_i:
                            dif.append({'run': x['run'], 'nivel': t_ + 1, 'motor': sorted(motor), 'independente': sorted(ind),
                                        'a12': [a12m, (i, t_) in a12], 'status': [stm, st_i],
                                        'z': [None if z[k][i] is None else round(z[k][i], 3) for k in range(NLV)]})
                        if x['fil'] == 1:
                            for k in range(5):
                                if REGRAS[k] in ind:
                                    cont_ind[t_][k] += 1
                            if ind:
                                rej_ind[t_] = True
                reg(f'G-WESTGARD-{tag}', f'Westgard N=3 ({an}/{lt}/{per}): marcas por ponto do motor (5 regras, alerta 1_2s e '
                    'veredito) = recálculo independente em toda a janela',
                    'APROVADO' if not dif and comparados else 'REPROVADO',
                    {'pontos_comparados': comparados, 'divergencias': dif[:10], 'n_divergencias': len(dif),
                     'alvo': {f'{k[0]}|N{k[1]}': v for k, v in alvo.items()}})
                cont_qc = [[est['engPainel'][t_][12 + k] for k in range(5)] for t_ in range(NLV)]
                tela_qc = [[est['A7:T9'][t_][c_] for c_ in (12, 15, 16, 13, 14)] for t_ in range(NLV)]   # M=1-3S P=2of3 Q=R4S N=3-1S O=6X
                st_qc = [est['engPainel'][t_][9] for t_ in range(NLV)]
                st_i = ['REJEITADO' if rej_ind[t_] else ('OK' if est['engPainel'][t_][1] else '') for t_ in range(NLV)]
                reg(f'G-CONTAGEM-{tag}', 'Painel "WESTGARD — violações no período" (M7:Q9, ordem 1_3s/2of3/R4S/3-1S/6X) e status = '
                    'contagem independente no período',
                    'APROVADO' if cont_qc == cont_ind and tela_qc == cont_ind and st_qc == st_i else 'REPROVADO',
                    {'engPainel_13a17': cont_qc, 'tela_reordenada': tela_qc, 'independente': cont_ind,
                     'status_motor': st_qc, 'status_independente': st_i, 'status_tela_J': [est['A7:T9'][k][9] for k in range(NLV)],
                     'ultima_violacao_S': [est['A7:T9'][k][18] for k in range(NLV)], 'periodo': [d0, d1]})
                OUT.setdefault('lj', {})[tag] = j({'estado': est, 'eng': eng[-30:]})
            except Exception as ex:
                OUT['erros'].append(f'G {tag}: {ex!r} {traceback.format_exc()[-1500:]}')
                reg(f'G-{tag}', f'Levey-Jennings {tag}', 'BLOQUEADO', {'erro': repr(ex)})

        # LJ no modo TODOS (WBC, fev-mar/2026): degraus de limite, X atravessando lotes, Westgard lote a lote
        try:
            an = 'WBC'
            if txt(t.p.Range('C3').Value) != an:
                t.muda('C3', an)
            t.muda('G3', serie(dt.date(2026, 2, 1)))
            t.muda('G4', serie(dt.date(2026, 3, 31)))
            d0, d1 = dt.date(2026, 2, 1), dt.date(2026, 3, 31)
            poe('H3', None)
            xl.Calculate()
            q.ex.esperar()
            est = t.estado()
            eng = t.eng()
            cq = {e['id']: e for e in b.elig}
            eixo_t = b.corridas(an, EQ, None, todos=True)
            corl = corrida_do_lote(eixo_t, NLV)
            jan, cort = janela(eixo_t, d1)
            jw = [c for c in jan if c['eleg']]
            lotes_j = sorted({c['lote'] for c in jw})
            alvo = {(lt, nv): t.alvo(lt, an, nv) for lt in lotes_j for nv in (1, 2, 3)}
            fl_t, a12_t, z_t = {}, set(), {}
            for lt in lotes_j:
                sub = [c for c in jw if c['lote'] == lt]
                fl, z, a12 = westgard_hema(sub, alvo)
                for (i, t_), v in fl.items():
                    fl_t[(sub[i]['k'], t_)] = v
                for (i, t_) in a12:
                    a12_t.add((sub[i]['k'], t_))
                for i, c in enumerate(sub):
                    z_t[c['k']] = [None if z[k][i] is None else round(z[k][i], 3) for k in range(NLV)]

            def chave(x):
                for t_ in range(NLV):
                    if x['id'][t_] in cq:
                        return x['run'] + '|' + cq[x['id'][t_]]['lote']
                return x['run'] + '|?'
            keys_eng = [chave(x) for x in eng]
            keys_ind = [c['k'] for c in jan]
            reg('G-EIXO-TODOS', 'Modo TODOS (WBC fev–mar/2026): corridas do motor = janela independente das chaves RUN|LOTE',
                'APROVADO' if keys_eng == keys_ind else 'REPROVADO',
                {'n_eng': len(keys_eng), 'n_ind': len(keys_ind), 'primeiras': [keys_eng[:3], keys_ind[:3]],
                 'ultimas': [keys_eng[-3:], keys_ind[-3:]], 'lotes_na_janela': lotes_j, 'cortadas': cort})
            dif, comparados = [], 0
            cont_ind = [[0] * 5 for _ in range(NLV)]
            for x in eng:
                k_ = chave(x)
                for t_ in range(NLV):
                    if not x['id'][t_]:
                        continue
                    comparados += 1
                    motor = {REGRAS[k] for k in range(5) if num(x['flags'][t_][k]) and x['flags'][t_][k] == 1}
                    ind = fl_t.get((k_, t_), set())
                    if motor != ind:
                        dif.append({'chave': k_, 'nivel': t_ + 1, 'motor': sorted(motor), 'independente': sorted(ind), 'z': z_t.get(k_)})
                    if x['fil'] == 1:
                        for k in range(5):
                            if REGRAS[k] in ind:
                                cont_ind[t_][k] += 1
            reg('G-WESTGARD-TODOS', 'Modo TODOS: Westgard lote a lote (cada lote só com os seus pontos, ADR-086) = motor',
                'APROVADO' if not dif else 'REPROVADO',
                {'pontos_comparados': comparados, 'n_divergencias': len(dif), 'divergencias': dif[:10]})
            cont_qc = [[est['engPainel'][t_][12 + k] for k in range(5)] for t_ in range(NLV)]
            reg('G-CONTAGEM-TODOS', 'Modo TODOS: contagens M7:Q9 = contagem independente lote a lote no período',
                'APROVADO' if cont_qc == cont_ind else 'REPROVADO', {'painel': cont_qc, 'independente': cont_ind})
            for nv in (1, 2, 3):
                ser, info = t.series(nv)
                pts = [(corl[c['v'][nv][1]], c['v'][nv][0], c['lote']) for c in jan if nv in c['v'] and d0 <= c['dh'].date() <= d1]
                alvo_xy = set((float(x), round(y, 9)) for x, y, _ in pts)
                melhor, mk = None, -1
                for nm, xy in ser.items():
                    k = len(set((float(x), round(y, 9)) for x, y in xy) & alvo_xy)
                    if k > mk:
                        melhor, mk = nm, k
                s_dados = set((float(x), round(y, 9)) for x, y in ser.get(melhor, []))
                xs = [x for x, _, _ in pts]
                lim_div = []
                for nm, xy in ser.items():
                    nome = nm.split(':', 1)[1]
                    if nome in LIM:
                        k = KS[LIM.index(nome)]
                        esp = {float(x): round(alvo[(lt, nv)][0] + k * alvo[(lt, nv)][1], 9) for x, _, lt in pts
                               if num(alvo[(lt, nv)][0]) and num(alvo[(lt, nv)][1])}
                        got = {float(x): round(y, 9) for x, y in xy}
                        for x_, y_ in esp.items():
                            if got.get(x_) != y_:
                                lim_div.append({'linha': nome, 'x': x_, 'grafico': got.get(x_), 'esperado': y_})
                        extra = sorted(set(got) - set(esp))
                        if extra:
                            lim_div.append({'linha': nome, 'x_sem_ponto_com_valor': extra[:4]})
                lotes_n = collections.Counter(lt for _, _, lt in pts)
                reg(f'G-TODOS-N{nv}', f'Modo TODOS N{nv}: pontos = ativos do período de todos os lotes; X sem buraco atravessando '
                    'os lotes; linhas de limite em degrau = Média + k·DP do lote de cada ponto',
                    'APROVADO' if s_dados == alvo_xy and all(b_ - a_ == 1 for a_, b_ in zip(xs, xs[1:])) and not lim_div else 'REPROVADO',
                    {'n_grafico': len(s_dados), 'n_ind': len(alvo_xy), 'x': [xs[:1], xs[-1:]], 'pontos_por_lote': dict(lotes_n),
                     'sobra': sorted(s_dados - alvo_xy)[:5], 'falta': sorted(alvo_xy - s_dados)[:5],
                     'limites_divergentes': lim_div[:8], 'eixo': info})
            OUT['todos_lj_estado'] = j(est)
        except Exception as ex:
            OUT['erros'].append(f'G-TODOS: {ex!r} {traceback.format_exc()[-1500:]}')

        # ------------------------------------------------ U: imagens e auditoria visual (exportadas pelo proprio Excel)
        try:
            t.p.Unprotect(SENHA)                      # passo tecnico do PageSetup da exportacao (nada e gravado)
            OUT['img_todos'] = aa.exportar_png(t.p, pasta, 45, 30)
            os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, 'Painel_todos_fevmar.png'))
            q.run('mSeguranca.ReprotectAll', teto=120)
            poe('H3', '619611')
            poe('K3', 2026)
            t.muda('K4', 'Set')
            OUT['estado_set_619611'] = j(t.estado())
            t.p.Unprotect(SENHA)
            OUT['achados_painel'] = aa.auditar_aba(t.p)
            OUT['img_set'] = aa.exportar_png(t.p, pasta, 45, 30)
            os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, 'Painel_set_619611.png'))
            ws = t.est
            ws.Unprotect(SENHA)
            OUT['achados_estatistica'] = aa.auditar_aba(ws)
            OUT['img_estatistica'] = aa.exportar_png(ws, pasta, 60, 30)
        except Exception as ex:
            OUT['erros'].append(f'U: {ex!r} {traceback.format_exc()[-1200:]}')
        OUT['estado_final'] = j(t.estado())
    finally:
        q.fechar(False)
        log('Excel encerrado (PID', OUT.get('excel_pid'), ')')
        with open(os.path.join(pasta, 'f4h.json'), 'w', encoding='utf-8') as f:
            json.dump(j(OUT), f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    executar(sys.argv[1], sys.argv[2], sys.argv[3])
