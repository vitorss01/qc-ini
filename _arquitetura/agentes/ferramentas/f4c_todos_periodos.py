# -*- coding: utf-8 -*-
"""F4 do Usuario Simulado -- 3o passe (ciclo 2026-10-31_ciclo01, reexecucao de 01/11): o que os dois primeiros
roteiros (f4_periodos_lotes.py e f4b_complemento.py) nao cobriam.

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    python -u f4c_todos_periodos.py <copia.xlsm> <cq.pkl> <pasta_saida>

    <copia.xlsm>  COPIA da producao. Nada e salvo nela.
    <cq.pkl>      tblCQ_Final da MESMA copia lida por openpyxl (ver f4_periodos_lotes.py).

O que faz (sempre pelas celulas de filtro do Painel com os eventos ligados, como o usuario):
  Q  periodos em OUTRO analito/equipamento: Ureia / DIMENSION 2 (todos os meses, trimestres e semestres) e
     Hemoglobina glicada / DIMENSION 1 (3 lotes em sequencia, com meses de sobreposicao: dez/2025 8588+8589 e
     jun/2026 8589+8590; o lote minoritario do mes tambem e testado).
  W  H3 vazio (ADR-086) com conferencia ponto a ponto do grafico e do Westgard:
       - X = corrida 1,2,3... por nivel, cronologica, atravessando os lotes (historico inteiro, chave RUN|LOTE);
       - linha de limite em degrau = Media + k*DP do lote do PROPRIO ponto;
       - Westgard lote a lote dentro da janela de 180 corridas RUN|LOTE (definicao do ADR-086) e, a parte, contra
         a visao do lote (normal), separando as divergencias de borda (ate o 10o ponto do lote na janela);
       - contagens M:Q do Painel e n por nivel.
     Cenarios: Glicose DIM 1 jun-ago/2026 (8974 -> 8977, + 8353); Hemoglobina glicada DIM 1 jun/2026 (8589 e 8590
     simultaneos); Ureia DIM 1 dez/2025 (linhas do lote de urina 6353 com MATRIZ = SORO).
  H  lista de lotes do H3 (lstLotes) e o que acontece ao escolher um lote que nao e do analito (Glicose + 8588).
  V  imagens do Painel exportadas pelo PROPRIO Excel da copia (PDF -> PNG) dos cenarios W.
Calculo independente: Python puro (statistics) sobre a tblCQ_Final bruta; nada de formula do QC_INI. A Media/DP do
alvo vem do cadastro (mLotes.MediaDoLote / DPDoLote). Senha: importada da suite de testes, nunca escrita aqui.
Autor: AGENTE_USUARIO_SIMULADO.
"""
import collections
import datetime as dt
import json
import os
import pickle
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import f4_periodos_lotes as f4  # noqa: E402
from f4_periodos_lotes import (QA, Bruto, Tela, cmp_painel, cmp_eng_escopo, janela, stats, j, reg, OUT, log,  # noqa: E402
                               corrida_do_lote, westgard, num, txt, LIM, KS, rel, TOL)
import auditoria_abas as aa  # noqa: E402
from qa_etl_recebimento import SENHA  # noqa: E402

MES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
NOMES = ['1_3s', '2_2s', 'R_4s', '4_1s', '8x']


def fim_mes(a, m):
    return dt.date(a + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)


def ir(t, an, eq, lote, ano, k4):
    """Caminho do usuario: equipamento, analito, lote, ano, periodo (nessa ordem, so o que muda)."""
    if txt(t.p.Range('L4').Value).upper() != eq:
        t.muda('L4', eq)
    if txt(t.p.Range('C3').Value) != an:
        t.muda('C3', an)
    if lote is not None and txt(t.p.Range('H3').Value) != lote:
        t.muda('H3', "'" + lote)
    if t.p.Range('K3').Value != ano:
        t.muda('K3', ano)
    return t.muda('K4', k4)


def periodos(t, b, an, eq, tag, extra_lotes=()):
    meses = sorted({(e['dh'].year, e['dh'].month) for e in b.sel(an, eq, None)})
    for (a, m) in meses:
        c = collections.Counter(e['lote'] for e in b.sel(an, eq, None) if e['dh'].year == a and e['dh'].month == m)
        lotes = [c.most_common(1)[0][0]] + [lt for lt, n in c.items() if lt in extra_lotes and lt != c.most_common(1)[0][0]]
        for lt in lotes:
            try:
                dtc = ir(t, an, eq, lt, a, MES[m - 1])
                est = t.estado()
                d0, d1 = dt.date(a, m, 1), fim_mes(a, m)
                g_ok = f4.dia(est['G3']) == d0 and f4.dia(est['G4']) == d1
                cmp_painel(t, est, an, eq, lt, d0, d1, f'Q-{tag}-M{a}{m:02d}-{lt}',
                           f'{an} / {eq} / lote {lt} / {m:02d}-{a}: Painel = dados brutos do mês',
                           {'G3G4_corretos': g_ok, 'tempo_s': dtc, 'lotes_no_mes': dict(c), 'K4': est['K4']})
            except Exception as ex:
                OUT['erros'].append(f'Q {tag} {a}-{m} {lt}: {ex!r} {traceback.format_exc()[-600:]}')
                reg(f'Q-{tag}-M{a}{m:02d}-{lt}', f'{an} {m:02d}/{a} lote {lt}', 'BLOQUEADO', {'erro': repr(ex)})
    # trimestres e semestres (lote majoritario do intervalo)
    for (a, rot, m0, m1) in [(2025, 'T4', 10, 12), (2026, 'T1', 1, 3), (2026, 'T2', 4, 6), (2026, 'T3', 7, 9),
                             (2026, 'T4', 10, 12), (2026, 'S1', 1, 6), (2026, 'S2', 7, 12)]:
        d0, d1 = dt.date(a, m0, 1), fim_mes(a, m1)
        c = collections.Counter(e['lote'] for e in b.sel(an, eq, None, None, d0, d1))
        if not c:
            continue
        lt = c.most_common(1)[0][0]
        try:
            dtc = ir(t, an, eq, lt, a, rot)
            est = t.estado()
            ok, linhas = cmp_painel(t, est, an, eq, lt, d0, d1, f'Q-{tag}-{rot}{a}-{lt}',
                                    f'{an} / {eq} / lote {lt} / {rot}-{a}: Painel = dados brutos do intervalo (não média dos meses)',
                                    {'tempo_s': dtc, 'lotes_no_intervalo': dict(c), 'K4': est['K4'],
                                     'G3G4_corretos': f4.dia(est['G3']) == d0 and f4.dia(est['G4']) == d1})
            # media das medias mensais, para mostrar que o Painel nao a usa
            mm = []
            for t_ in range(2):
                ms = []
                for m in range(m0, m1 + 1):
                    s = b.stats_periodo(an, eq, lt, t_ + 1, dt.date(a, m, 1), fim_mes(a, m))
                    if s['n']:
                        ms.append(s['media'])
                mm.append(sum(ms) / len(ms) if ms else None)
            OUT.setdefault('media_dos_meses', {})[f'{tag}-{rot}{a}-{lt}'] = j({'painel': [est['engPainel'][k][2] for k in range(2)],
                                                                             'media_das_medias_mensais': mm})
        except Exception as ex:
            OUT['erros'].append(f'Q {tag} {rot}{a}: {ex!r} {traceback.format_exc()[-600:]}')


def janela_todos(eixo_todos, fim, nk=180):
    return janela(eixo_todos, fim, nk)


def todos(t, b, an, eq, d0, d1, lote_ini, tag, pasta, rows):
    """H3 vazio: pontos, X, degraus, Westgard lote a lote e contagens."""
    a, m = d0.year, d0.month
    ir(t, an, eq, lote_ini, a, MES[m - 1])
    t.muda('G3', f4.serie(d0))
    t.muda('G4', f4.serie(d1))
    dtc = t.muda('H3', None)
    t.xl.Calculate()
    t.q.ex.esperar()
    est = t.estado()
    eng = t.eng()
    cq = {e['id']: e for e in b.elig}
    mat = {txt(r['ID_REGISTRO']): txt(r.get('MATRIZ')) for r in rows}
    # 1) escopo + n
    eng_, fora, faltam = cmp_eng_escopo(t, an, eq, None, d0, d1, todos=True)
    ind_n = [len(b.sel(an, eq, None, k + 1, d0, d1)) for k in range(2)]
    qc_n = [est['engPainel'][k][1] for k in range(2)]
    vazios = all(est['engPainel'][k][c] in (None, '') for k in range(2) for c in (2, 3, 4))
    pts = collections.Counter()
    for e in b.elig:
        if e['an'] == an and e['eq'] == eq.upper() and d0 <= e['dh'].date() <= d1:
            pts[f"{e['lote']}|N{e['nv']}|{mat.get(e['id'])}"] += 1
    reg(f'W-{tag}-ESCOPO', f'H3 vazio {an}/{eq} {d0:%d/%m/%y}–{d1:%d/%m/%y}: motor com todos os lotes do período, n = independente, '
        'média/DP/CV vazios', 'APROVADO' if (not fora and not faltam and qc_n == ind_n and vazios and est['eng_AD1']) else 'REPROVADO',
        {'engTodos_AD1': est['eng_AD1'], 'H4': est['H4'], 'A39': est['A39'], 'n_qc': qc_n, 'n_independente': ind_n,
         'estat_vazias': vazios, 'pontos_por_lote_nivel_matriz': dict(pts), 'fora': fora[:5], 'faltam': faltam[:5],
         'linha_painel': est['B7:T8'], 'tempo_s': dtc, 'eng_O1': est['eng_O1']})
    # 2) eixo e janela independentes (RUN|LOTE)
    eixo_all = b.corridas(an, eq, None, todos=True)
    xcor = corrida_do_lote(eixo_all)                      # X atravessando os lotes (historico inteiro)
    jan, cort = janela(eixo_all, d1)
    n_eng = len(eng)
    alvo = {}
    for c in jan:
        for nv, v in c['v'].items():
            lt = v[2]
            if (lt, nv) not in alvo:
                alvo[(lt, nv)] = t.alvo(lt, an, nv)
    # 3) pontos do grafico x independente; degraus
    for gi, nv in ((1, 1), (2, 2)):
        ser, info = t.series(gi)
        pts_ind = [(xcor[c['v'][nv][1]], c['v'][nv][0], c['v'][nv][2]) for c in jan
                   if nv in c['v'] and d0 <= c['dh'].date() <= d1]
        alvo_xy = set((float(x), round(y, 9)) for x, y, _ in pts_ind)
        melhor, mk = None, -1
        for nm, xy in ser.items():
            k = len(set((float(x), round(y, 9)) for x, y in xy) & alvo_xy)
            if k > mk:
                melhor, mk = nm, k
        s_dados = set((float(x), round(y, 9)) for x, y in ser.get(melhor, []))
        sobra, falta = sorted(s_dados - alvo_xy)[:6], sorted(alvo_xy - s_dados)[:6]
        reg(f'W-{tag}-PONTOS-N{nv}', f'H3 vazio {an}: pontos do LJ N{nv} = ativos do período; X = corrida por nível atravessando os lotes',
            'APROVADO' if pts_ind and not sobra and not falta else 'REPROVADO',
            {'serie': melhor, 'n_grafico': len(s_dados), 'n_independente': len(pts_ind),
             'x_primeiro_ultimo': [min(x for x, _, _ in pts_ind) if pts_ind else None, max(x for x, _, _ in pts_ind) if pts_ind else None],
             'sobra': sobra, 'falta': falta, 'eixo': info})
        # degraus
        div = []
        nchk = 0
        por_nome = {nm.split(':', 1)[1]: dict((float(x), y) for x, y in xy) for nm, xy in ser.items()}
        for x, y, lt in pts_ind:
            M, S = alvo[(lt, nv)]
            for nome, k in zip(LIM, KS):
                yy = por_nome.get(nome, {}).get(float(x))
                esp = M + k * S if num(M) and num(S) else None
                nchk += 1
                if esp is None or yy is None or abs(yy - esp) > 1e-9 * max(1, abs(esp)):
                    div.append({'x': x, 'lote': lt, 'linha': nome, 'grafico': yy, 'esperado': esp})
        reg(f'W-{tag}-DEGRAUS-N{nv}', f'H3 vazio {an}: as 7 linhas de limite em cada ponto N{nv} = Média + k·DP do lote do PRÓPRIO ponto',
            'APROVADO' if nchk and not div else 'REPROVADO', {'conferencias': nchk, 'divergencias': div[:8], 'n_div': len(div)})
    # 4) Westgard: (a) lote a lote dentro da janela TODOS (ADR-086); (b) visao do lote (janela do lote)
    jw = [c for c in jan if c['eleg']]
    fl_a = {}
    pos_lote = {}
    for lt in sorted({v[2] for c in jw for v in c['v'].values()}):
        seq = [c for c in jw if any(v[2] == lt for v in c['v'].values())]
        f_, _ = westgard(seq, alvo)
        for i, c in enumerate(seq):
            for nv in (1, 2):
                if nv in c['v'] and c['v'][nv][2] == lt:
                    fl_a[c['v'][nv][1]] = f_.get((i, nv - 1), set())
                    pos_lote[c['v'][nv][1]] = i + 1
    fl_b = {}
    for lt in {v[2] for c in jw for v in c['v'].values()}:
        eixo_l = b.corridas(an, eq, lt)
        jl, _ = janela(eixo_l, d1)
        seq = [c for c in jl if c['eleg']]
        al = {(lt, nv): alvo.get((lt, nv)) or t.alvo(lt, an, nv) for nv in (1, 2)}
        f_, _ = westgard(seq, al)
        for i, c in enumerate(seq):
            for nv in (1, 2):
                if nv in c['v']:
                    fl_b[c['v'][nv][1]] = f_.get((i, nv - 1), set())
    dif_a, dif_b, dif_b_borda = [], [], []
    comp = 0
    cont_ind = [[0] * 5 for _ in range(2)]
    for x in eng:
        for t_ in range(2):
            i = x['id'][t_]
            if not i or i not in cq:
                continue
            comp += 1
            motor = {NOMES[k] for k in range(5) if num(x['flags'][t_][k]) and x['flags'][t_][k] == 1}
            a_ = fl_a.get(i, set())
            if motor != a_:
                dif_a.append({'id': i, 'lote': cq[i]['lote'], 'nivel': t_ + 1, 'motor': sorted(motor), 'independente': sorted(a_)})
            b_ = fl_b.get(i, set())
            if motor != b_:
                item = {'id': i, 'lote': cq[i]['lote'], 'nivel': t_ + 1, 'motor': sorted(motor), 'visao_do_lote': sorted(b_),
                        'posicao_no_lote_na_janela': pos_lote.get(i)}
                (dif_b_borda if (pos_lote.get(i) or 99) <= 10 else dif_b).append(item)
            if x['fil'] == 1:
                for k in range(5):
                    if NOMES[k] in a_:
                        cont_ind[t_][k] += 1
    reg(f'W-{tag}-WESTGARD', f'H3 vazio {an}: marcas de Westgard por ponto = recálculo independente lote a lote na janela de 180 corridas RUN|LOTE (ADR-086)',
        'APROVADO' if comp and not dif_a else 'REPROVADO',
        {'pontos_comparados': comp, 'n_div': len(dif_a), 'divergencias': dif_a[:10], 'corridas_eng': n_eng, 'corridas_janela_ind': len(jan),
         'alvos': {f'{k[0]}|N{k[1]}': v for k, v in alvo.items()}})
    reg(f'W-{tag}-WESTGARD-VISAO-LOTE', f'H3 vazio {an}: veredito por ponto = o da visão do lote (fora da borda de 10 pontos)',
        'APROVADO' if comp and not dif_b else 'REPROVADO',
        {'n_div_fora_da_borda': len(dif_b), 'divergencias': dif_b[:10], 'n_div_na_borda': len(dif_b_borda), 'borda': dif_b_borda[:6]})
    cont_qc = [[est['B7:T8'][t_][12 + k] for k in range(5)] for t_ in range(2)]
    reg(f'W-{tag}-CONTAGEM', f'H3 vazio {an}: contagens M:Q do Painel = contagem independente no período',
        'APROVADO' if cont_qc == cont_ind else 'REPROVADO',
        {'painel': cont_qc, 'independente': cont_ind, 'status_J': [est['B7:T8'][0][9], est['B7:T8'][1][9]]})
    try:
        t.p.Unprotect(SENHA)                 # passo tecnico: PageSetup da exportacao (nada e gravado)
        png = aa.exportar_png(t.p, pasta, 45, 30)
        os.replace(os.path.join(pasta, 'Painel.png'), os.path.join(pasta, f'Painel_todos_{tag}.png'))
        OUT.setdefault('imagens', []).append(f'Painel_todos_{tag}.png')
        t.q.run('mSeguranca.ReprotectAll', teto=120)
    except Exception as ex:
        OUT.setdefault('imagens', []).append(f'{tag}: {ex!r}')
    OUT.setdefault('todos', {})[tag] = j({'estado': est, 'eng_ult': eng[-12:]})


def executar(copia, pkl, pasta):
    os.makedirs(pasta, exist_ok=True)
    d = pickle.load(open(pkl, 'rb'))
    b = Bruto(d['rows'])
    q = QA('Bioquimica', copia)
    OUT['excel_pid'] = q.ex.pid
    try:
        t = Tela(q, b)
        q.run('mSeguranca.ReprotectAll', teto=120)
        OUT['estado_inicial'] = j(t.estado())
        # ---------------- Q: periodos em outro analito / equipamento
        periodos(t, b, 'Ureia', 'DIMENSION 2', 'UR2')
        periodos(t, b, 'Hemoglobina glicada', 'DIMENSION 1', 'HBA1', extra_lotes=('8588', '8589', '8590'))
        # ---------------- W: H3 vazio ponto a ponto
        for (an, eq, d0, d1, lt0, tag) in [
                ('Glicose', 'DIMENSION 1', dt.date(2026, 6, 1), dt.date(2026, 8, 31), '8974', 'GLI1-JUNAGO'),
                ('Hemoglobina glicada', 'DIMENSION 1', dt.date(2026, 6, 1), dt.date(2026, 6, 30), '8589', 'HBA1-JUN'),
                ('Ureia', 'DIMENSION 1', dt.date(2025, 12, 1), dt.date(2025, 12, 31), '8974', 'UR1-DEZ')]:
            try:
                todos(t, b, an, eq, d0, d1, lt0, tag, pasta, d['rows'])
            except Exception as ex:
                OUT['erros'].append(f'W {tag}: {ex!r} {traceback.format_exc()[-1200:]}')
                reg(f'W-{tag}', f'H3 vazio {tag}', 'BLOQUEADO', {'erro': repr(ex)})
        # ---------------- H: lista de lotes e lote de outro analito
        try:
            ir(t, 'Glicose', 'DIMENSION 1', '8977', 2026, 'Set')
            v, lst = t.lista(t.p, 'H3')
            antes = t.estado()
            dtc = t.muda('H3', "'8588")             # lote da Hemoglobina glicada
            dep = t.estado()
            lotes_gli = sorted({e['lote'] for e in b.sel('Glicose', 'DIMENSION 1', None)})
            reg('H-LISTA', 'Lista de lotes do H3: mostra só os lotes do analito/equipamento?',
                'APROVADO' if set(lst) <= set(lotes_gli) else 'REPROVADO',
                {'validacao': v, 'lista_H3': lst, 'lotes_com_resultado_de_Glicose_DIM1': lotes_gli,
                 'fora_do_analito': [x for x in lst if x not in lotes_gli]})
            reg('H-OUTRO', 'Escolher em H3 um lote de outro analito (Glicose + 8588, lote da Hemoglobina glicada): o que o Painel faz',
                'APROVADO' if txt(dep['H4']) or txt(dep['A39']) else 'REPROVADO',
                {'antes': {k: antes[k] for k in ('H3', 'H4', 'G3', 'G4', 'K4', 'A39')},
                 'depois': {k: dep[k] for k in ('H3', 'H4', 'G3', 'G4', 'K4', 'A39', 'C3', 'C4')},
                 'n': [dep['engPainel'][0][1], dep['engPainel'][1][1]], 'linha': dep['B7:T8'], 'tempo_s': dtc,
                 'alvo_8588_glicose': [t.alvo('8588', 'Glicose', 1), t.alvo('8588', 'Glicose', 2)]})
            t.muda('H3', "'8977")
        except Exception as ex:
            OUT['erros'].append(f'H: {ex!r} {traceback.format_exc()[-800:]}')
        OUT['estado_final'] = j(t.estado())
    finally:
        q.fechar(False)
        log('Excel encerrado (PID', OUT.get('excel_pid'), ')')
        with open(os.path.join(pasta, 'f4c.json'), 'w', encoding='utf-8') as f:
            json.dump(j(OUT), f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    executar(sys.argv[1], sys.argv[2], sys.argv[3])
