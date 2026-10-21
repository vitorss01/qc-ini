# -*- coding: utf-8 -*-
"""qa_corrida_lote.py -- ADR-080/081: o eixo X do grafico de cada nivel e a CORRIDA DO LOTE (1, 2, 3... por nivel),
contada so sobre os resultados ATIVOS; o X vermelho fica ENTRE as corridas vizinhas (sem numero).

Uso: python qa_corrida_lote.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

Oraculo independente do motor: na Principal - Resultados, os resultados ATIVOS (participa da estatistica) de um nivel,
do lote e do equipamento do grafico, ordenados por data/hora: a posicao (1, 2, 3...) e a corrida do lote. Cada X
vermelho (inativado com REGISTRAR - LJ) fica em k + j/(m+1): k = ativos antes dele, j = sua ordem entre os m X do
mesmo intervalo. O que o motor publicou (Eng_Saida BK..BM e BN..BV) e o que o Calc plota (CR..CT, CU..DC) tem de bater.

  L01  varredura dos analitos do lote em tela: Eng_Saida = oraculo (corridas e posicao dos X), em todos os niveis;
  L02  em cada nivel os ativos plotados sao 1, 2, 3... sem buraco e cada X fica estritamente entre inteiros;
  L03  outros lotes do mesmo analito: cada lote recomeca do 1;
  L04  N1 digitado as 23:50 num dia que ja tem corridas -> corrida SEGUINTE do N1, os outros niveis nao mudam;
  L05  (T01) inativar pela ID o resultado da corrida 3 (com X): o N1 continua 1..n-1 sem lacuna, a antiga 4 vira 3,
       o X fica em 2,5, os outros niveis nao mudam;
  L06  (T02) reativar (apagar a linha da aba Inativar): a numeracao volta EXATAMENTE a de antes;
  L07  (T03) inativar sem X (REGISTRAR - LJ desmarcado): renumera e nao desenha X;
  L08  (T04/T05) inativar o 1o e o ultimo do lote: a nova 1 e a antiga 2; X em 0,5 e depois da ultima; o eixo cobre;
  L09  (T06) dois resultados seguidos inativados (corridas 2 e 3): X em 1,33 e 1,67, a antiga 4 vira 2;
  L10  (T17) linha da aba Inativar com o ANALITO errado: nada e inativado, nada renumera;
  L11  (T20) ATUALIZAR duas vezes seguidas: Eng_Saida identico.
Nada salvo. Sai com 1 se algum falhar.
"""
import datetime as dt
import json
import math
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA  # noqa: E402

COL_ENG = ['BK', 'BL', 'BM']
COL_CALC = ['CR', 'CS', 'CT']
ENG_XPOS0, CALC_XPOS0 = 66, 99
TB = 'tblInativacao_NaoConformes'
RES = []


def reg(nome, ok, ev):
    RES.append({'teste': nome, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': ev})
    print(('PASS ' if ok else 'FAIL ') + nome + ' -- ' + json.dumps(ev, ensure_ascii=False, default=str)[:700], flush=True)


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def txt(v):
    return '' if v is None else str(v).strip()


def iint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def eng(q, nlv):
    e = q.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    runs = [iint(r[0]) for r in e.Range('B3:B182').Value][:n]
    cor = [[iint(r[0]) if num(r[0]) else None for r in e.Range(f'{COL_ENG[t]}3:{COL_ENG[t]}182').Value][:n]
           for t in range(nlv)]
    xpos = []
    for t in range(nlv):
        bloco = e.Range(f'{col(ENG_XPOS0 + t * 3)}3:{col(ENG_XPOS0 + t * 3 + 2)}182').Value[:n]
        xpos.append([sorted(round(v, 9) for v in linha if num(v)) for linha in bloco])
    return {'analito': txt(e.Range('C1').Value), 'lote': txt(e.Range('E1').Value), 'equip': txt(e.Range('S1').Value),
            'n': n, 'runs': runs, 'cor': cor, 'xpos': xpos, 'cortadas': e.Range('M1').Value}


def oraculo(fin, analito, lote, equip, nlv):
    """nivel -> ({RUN: corrida do lote dos ATIVOS}, {RUN: [posicoes dos X]})."""
    out = []
    for t in range(1, nlv + 1):
        base = [r for r in fin if txt(r['ANALITO']).upper() == analito.upper() and txt(r['LOTE']) == lote
                and (not equip or txt(r['EQUIPAMENTO']).upper() == equip.upper()) and iint(r['NIVEL']) == t
                and num(r['RESULTADO']) and iint(r['RUN']) is not None]
        ativos = sorted((r for r in base if r['PARTICIPA_ESTATISTICA'] == 'SIM'),
                        key=lambda r: (r['DATA_HORA'], iint(r['RUN']), str(r['ID_REGISTRO'])))
        xs = sorted((r for r in base if r['PARTICIPA_ESTATISTICA'] != 'SIM' and r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'),
                    key=lambda r: (r['DATA_HORA'], iint(r['RUN']), str(r['ID_REGISTRO'])))
        m = {}
        for r in ativos:
            if iint(r['RUN']) not in m:
                m[iint(r['RUN'])] = len(m) + 1
        tempos = [(r['DATA_HORA'], iint(r['RUN'])) for r in ativos]
        grupos = {}
        for r in xs:
            k = sum(1 for tt in tempos if tt < (r['DATA_HORA'], iint(r['RUN'])))
            grupos.setdefault(k, []).append(r)
        px = {}
        for k, lst in grupos.items():
            for j, r in enumerate(lst, 1):
                px.setdefault(iint(r['RUN']), []).append(round(k + j / (len(lst) + 1), 9))
        out.append((m, {k: sorted(v) for k, v in px.items()}))
    return out


def comparar(e, orc, nlv):
    dif = []
    for t in range(nlv):
        m, px = orc[t]
        for i, run in enumerate(e['runs']):
            if e['cor'][t][i] != m.get(run):
                dif.append({'nivel': t + 1, 'RUN': run, 'motor': e['cor'][t][i], 'oraculo': m.get(run)})
            if e['xpos'][t][i] != px.get(run, []):
                dif.append({'nivel': t + 1, 'RUN': run, 'X_motor': e['xpos'][t][i], 'X_oraculo': px.get(run, [])})
    return dif


def plotados(q, nlv):
    """nivel -> (X dos ativos plotados, posicoes dos X vermelhos plotados), na ordem das linhas do Calc."""
    c = q.ws('Calc')
    cab = [txt(v) for v in c.Range('A2:DZ2').Value[0]]
    out = []
    for i in range(1, nlv + 1):
        xs = [r[0] for r in c.Range(f'{COL_CALC[i - 1]}3:{COL_CALC[i - 1]}182').Value]
        yl = [r[0] for r in c.Range(c.Cells(3, cab.index(f'N{i} line') + 1), c.Cells(182, cab.index(f'N{i} line') + 1)).Value]
        xv = c.Range(c.Cells(3, cab.index(f'N{i} x1') + 1), c.Cells(182, cab.index(f'N{i} x3') + 1)).Value
        c0 = CALC_XPOS0 + (i - 1) * 3
        xp = c.Range(c.Cells(3, c0), c.Cells(182, c0 + 2)).Value
        atv = [xs[r] for r in range(180) if num(yl[r])]
        xx = [xp[r][j] for r in range(180) for j in range(3) if num(xv[r][j])]
        out.append((atv, xx))
    return out


def seq_ok(xs):
    return all(num(x) and x == math.floor(x) for x in xs) and all(b - a == 1 for a, b in zip(xs, xs[1:]))


def x_ok(xx):
    return all(num(p) and p != math.floor(p) for p in xx)


def executar(produto, caminho):
    nlv = 2 if produto.startswith('Bio') else 3
    q = QA(produto, caminho)
    try:
        p = q.ws('Painel')
        nomes_cat = [txt(r[0]) for r in q.espec().Range('A4:A43').Value]

        def ir_analito(an):
            p.Range('B3').Value = nomes_cat.index(an) + 1
            q.run('mEstatistica.PainelMudou', teto=300)
            q.ex.xl.Calculate()

        def trocar(lote):
            p.Range('H3').Value = "'" + lote
            q.ex.esperar()
            q.ex.xl.Calculate()
            q.run('mEstatistica.RecalcularAnalitoAtual', teto=300)
            q.ex.xl.Calculate()

        def inativar(id_, analito, lj=True, motivo='QA ADR-081'):
            q.ex.xl.EnableEvents = True
            q.escrever_linha(TB, {'ID_REGISTRO': id_, 'ANALITO': analito, 'REGISTRAR - LJ': lj, 'MOTIVO': motivo})

        def reativar(id_):
            q.ex.xl.EnableEvents = True
            q.limpar_linha(TB, 'ID_REGISTRO', id_)

        def estado(an):
            q.atualizar()
            f = q.ler('tblCQ_Final')
            ir_analito(an)
            e = eng(q, nlv)
            return f, e, oraculo(f, e['analito'], e['lote'], e['equip'], nlv)

        q.atualizar()
        fin = q.ler('tblCQ_Final')
        e0 = eng(q, nlv)
        lote0 = e0['lote']

        # ---- L01 + L02: varredura dos analitos do lote em tela
        difs, seqs, vistos, nX, t0 = [], [], 0, 0, time.time()
        for an in [a for a in nomes_cat if a]:
            if not any(txt(r['ANALITO']).upper() == an.upper() and txt(r['LOTE']) == lote0 for r in fin):
                continue
            ir_analito(an)
            e = eng(q, nlv)
            if e['n'] == 0:
                continue
            vistos += 1
            d = comparar(e, oraculo(fin, e['analito'], e['lote'], e['equip'], nlv), nlv)
            if d:
                difs.append({'analito': an, 'n_dif': len(d), 'amostra': d[:3]})
            for t, (atv, xx) in enumerate(plotados(q, nlv), 1):
                nX += len(xx)
                if (atv and not seq_ok(atv)) or not x_ok(xx):
                    seqs.append({'analito': an, 'nivel': t, 'ativos': atv[:12], 'X': xx[:6]})
        reg('L01 Corrida do lote (só ATIVOS) e posição dos X publicadas pelo motor = oráculo, em todos os analitos do lote '
            'e todos os níveis', vistos > 0 and not difs,
            {'lote': lote0, 'analitos_conferidos': vistos, 'divergencias': difs[:5], 's': round(time.time() - t0, 1)})
        reg('L02 No gráfico de cada nível os ativos são 1, 2, 3... sem buraco e cada X vermelho fica entre duas corridas',
            vistos > 0 and not seqs, {'fora': seqs[:5], 'X_vermelhos_conferidos': nX})

        # ---- L03: outros lotes recomecam do 1
        an = e0['analito']
        ir_analito(an)
        lotes = sorted({txt(r['LOTE']) for r in fin if txt(r['ANALITO']).upper() == an.upper()
                        and r['PARTICIPA_ESTATISTICA'] == 'SIM' and txt(r['LOTE'])} - {lote0})
        ev3, ruins3 = [], []
        for L in lotes[:3]:
            trocar(L)
            e = eng(q, nlv)
            if e['lote'] != L or e['n'] == 0:
                ev3.append({'lote': L, 'motor_lote': e['lote'], 'n': e['n']})
                continue
            d = comparar(e, oraculo(fin, e['analito'], L, e['equip'], nlv), nlv)
            primeiros = [min([c for c in e['cor'][t] if c], default=None) for t in range(nlv)]
            inicio = e['cortadas'] in (0, None, '')
            ev3.append({'lote': L, 'corridas': e['n'], 'primeira_por_nivel': primeiros, 'janela_desde_o_inicio': inicio,
                        'divergencias': len(d)})
            if d or (inicio and any(x not in (1, None) for x in primeiros)):
                ruins3.append(L)
        trocar(lote0)
        reg('L03 Cada lote recomeça do 1 e a sequência bate com o oráculo', any('corridas' in x for x in ev3) and not ruins3,
            {'analito': an, 'lotes': ev3, 'ruins': ruins3})

        # ---- preparacao dos cenarios de inativacao: o N1 do analito em tela, no lote em uso
        ir_analito(an)
        e1 = eng(q, nlv)
        orc1 = oraculo(fin, e1['analito'], e1['lote'], e1['equip'], nlv)
        ativos_n1 = sorted(((c_, run) for run, c_ in orc1[0][0].items()))      # (corrida, RUN)
        an_cq = e1['analito']

        def id_da_corrida(f, corrida, nivel=1, orc=None):
            m = (orc or orc1)[nivel - 1][0]
            run = next(r_ for r_, c_ in m.items() if c_ == corrida)
            return next(r['ID_REGISTRO'] for r in f if iint(r['RUN']) == run and iint(r['NIVEL']) == nivel
                        and txt(r['ANALITO']).upper() == an_cq.upper() and txt(r['LOTE']) == lote0
                        and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                        and (not e1['equip'] or txt(r['EQUIPAMENTO']).upper() == e1['equip'].upper()))

        n1 = len(ativos_n1)
        cor_base = [list(x) for x in e1['cor']]

        # ---- L05 (T01): inativar a corrida 3 com X
        id3 = id_da_corrida(fin, 3)
        id4 = id_da_corrida(fin, 4)
        inativar(id3, an_cq, True)
        f5, e5, o5 = estado(an)
        n1_5 = sorted(o5[0][0].values())
        nova3 = next((r_ for r_, c_ in o5[0][0].items() if c_ == 3), None)
        run4 = next(r_ for r_, c_ in orc1[0][0].items() if c_ == 4)
        xs5 = [p_ for lst in o5[0][1].values() for p_ in lst]
        d5 = comparar(e5, o5, nlv)
        outros5 = all(o5[t][0] == orc1[t][0] for t in range(1, nlv))
        atv5, xx5 = plotados(q, nlv)[0]
        reg('L05 (T01) Inativar pela ID a corrida 3 (com X): N1 continua 1..n−1 sem lacuna, a antiga 4 vira 3, X em 2,5, '
            'outros níveis iguais, motor = oráculo',
            n1_5 == list(range(1, n1)) and nova3 == run4 and 2.5 in xs5 and outros5 and not d5 and seq_ok(atv5) and x_ok(xx5),
            {'ID_inativado': id3, 'ID_que_virou_corrida_3': id4, 'N1_antes': n1, 'N1_depois': len(n1_5),
             'X_posicoes': xs5[:6], 'outros_niveis_iguais': outros5, 'divergencias': d5[:3]})

        # ---- L06 (T02): reativar volta exatamente
        reativar(id3)
        f6, e6, o6 = estado(an)
        reg('L06 (T02) Reativar (apagar a linha da aba Inativar): numeração do motor volta EXATAMENTE à de antes',
            e6['cor'] == cor_base and o6[0][0] == orc1[0][0], {'ID': id3, 'igual_ao_inicio': e6['cor'] == cor_base})

        # ---- L07 (T03): sem X
        inativar(id3, an_cq, False)
        f7, e7, o7 = estado(an)
        n1_7 = sorted(o7[0][0].values())
        run3 = next(r_ for r_, c_ in orc1[0][0].items() if c_ == 3)
        sem_x = run3 not in o7[0][1] and (run3 not in e7['runs'] or not e7['xpos'][0][e7['runs'].index(run3)])
        d7 = comparar(e7, o7, nlv)
        reg('L07 (T03) Inativar sem X (REGISTRAR - LJ desmarcado): renumera (1..n−1) e não desenha X',
            n1_7 == list(range(1, n1)) and sem_x and not d7, {'ID': id3, 'N1_depois': len(n1_7), 'sem_X': sem_x,
                                                            'divergencias': d7[:3]})
        reativar(id3)

        # ---- L08 (T04/T05): primeiro e ultimo do lote
        id1 = id_da_corrida(fin, 1)
        idn = id_da_corrida(fin, n1)
        inativar(id1, an_cq, True)
        inativar(idn, an_cq, True)
        f8, e8, o8 = estado(an)
        n1_8 = sorted(o8[0][0].values())
        run2 = next(r_ for r_, c_ in orc1[0][0].items() if c_ == 2)
        xs8 = sorted(p_ for lst in o8[0][1].values() for p_ in lst)
        bx = (q.nome('xrunmin1').Value, q.nome('xrunmax1').Value)
        _, xx8 = plotados(q, nlv)[0]
        cobre = all(bx[0] <= p_ <= bx[1] for p_ in xx8)
        d8 = comparar(e8, o8, nlv)
        reg('L08 (T04/T05) Inativar o 1º e o último do lote: a nova corrida 1 é a antiga 2; X em 0,5 e depois da última; '
            'o eixo enquadra os X',
            n1_8 == list(range(1, n1 - 1)) and o8[0][0].get(run2) == 1 and 0.5 in xs8
            and (n1 - 2) + 0.5 in xs8 and cobre and not d8,
            {'IDs': [id1, idn], 'X': xs8[:6], 'bordas_N1': bx, 'X_plotados_no_periodo': xx8[:6], 'divergencias': d8[:3]})
        reativar(id1)
        reativar(idn)

        # ---- L09 (T06): duas seguidas (2 e 3)
        id2 = id_da_corrida(fin, 2)
        inativar(id2, an_cq, True)
        inativar(id3, an_cq, True)
        f9, e9, o9 = estado(an)
        xs9 = sorted(p_ for lst in o9[0][1].values() for p_ in lst)
        d9 = comparar(e9, o9, nlv)
        reg('L09 (T06) Duas corridas seguidas inativadas (2 e 3): X em 1,33 e 1,67, a antiga 4 vira 2',
            round(1 + 1 / 3, 9) in xs9 and round(1 + 2 / 3, 9) in xs9 and o9[0][0].get(run4) == 2 and not d9,
            {'IDs': [id2, id3], 'X': xs9[:6], 'antiga_4_agora': o9[0][0].get(run4), 'divergencias': d9[:3]})
        reativar(id2)
        reativar(id3)

        # ---- L10 (T17): analito errado na aba Inativar -> nada muda
        outro = next(a for a in nomes_cat if a and a.upper() != an_cq.upper())
        inativar(id3, outro, True)
        f10, e10, o10 = estado(an)
        st3 = next(r['STATUS_ANALITICO'] for r in f10 if r['ID_REGISTRO'] == id3)
        e10_qa = [r for r in q.ler('tblQA_Integracao') if r['CODIGO'] == 'E10' and r['ID_REGISTRO'] == id3]
        reg('L10 (T17) Linha da aba Inativar com ANALITO errado: o resultado continua ATIVO, nada renumera, QA E10',
            st3 == 'ATIVO' and e10['cor'] == cor_base and bool(e10_qa),
            {'ID': id3, 'analito_digitado': outro, 'status': st3, 'numeracao_igual': e10['cor'] == cor_base,
             'E10': bool(e10_qa)})
        reativar(id3)

        # ---- L11 (T20): idempotencia
        fa, ea, _ = estado(an)
        fb, eb, _ = estado(an)
        reg('L11 (T20) ATUALIZAR duas vezes seguidas: Eng_Saida (corridas e posição dos X) idêntico',
            ea['cor'] == eb['cor'] and ea['xpos'] == eb['xpos'] and ea['cor'] == cor_base, {'igual': ea['cor'] == eb['cor']})

        # ---- L04: N1 digitado as 23:50 (por ultimo: deixa um manual na copia)
        alvo = [r for r in fb if txt(r['ANALITO']).upper() == an.upper() and txt(r['LOTE']) == lote0
                and r['PARTICIPA_ESTATISTICA'] == 'SIM' and isinstance(r['DATA_HORA'], dt.datetime)
                and (not e1['equip'] or txt(r['EQUIPAMENTO']).upper() == e1['equip'].upper())]
        ult = max(alvo, key=lambda r: r['DATA_HORA'])
        dia = ult['DATA_HORA'].date()
        n1v = sorted(r['RESULTADO'] for r in alvo if iint(r['NIVEL']) == 1)
        quando = dt.datetime.combine(dia, dt.time(23, 50))
        q.ex.xl.EnableEvents = True
        q.escrever_linha('tblResultados_Manuais', {
            'DATA': dt.datetime.combine(dia, dt.time()), 'HORA': (23 * 60 + 50) / 1440.0, 'LOTE': lote0, 'NIVEL': 1,
            'ANALITO': an, 'RESULTADO': n1v[len(n1v) // 2], 'EQUIPAMENTO': ult['EQUIPAMENTO'],
            'MOTIVO': 'QA ADR-080 corrida da noite'})
        f4, e4, o4 = estado(an)
        man = [r for r in f4 if r['ORIGEM_RESULTADO'] == 'MANUAL' and txt(r['ANALITO']).upper() == an.upper()
               and isinstance(r['DATA_HORA'], dt.datetime) and r['DATA_HORA'] == quando]
        run_m = iint(man[0]['RUN']) if man else None
        cor_m = o4[0][0].get(run_m)
        d4 = comparar(e4, o4, nlv)
        outros4 = all(o4[t][0] == orc1[t][0] for t in range(1, nlv))
        reg('L04 N1 digitado às 23:50 num dia que já tem corridas: vira a corrida SEGUINTE do N1 (última + 1); os outros '
            'níveis não mudam',
            bool(man) and man[0]['PARTICIPA_ESTATISTICA'] == 'SIM' and cor_m == n1 + 1 and outros4 and not d4,
            {'manual': man[0]['ID_REGISTRO'] if man else None, 'dia': dia, 'RUN_do_dia': run_m, 'corrida_do_lote': cor_m,
             'ultima_N1_antes': n1, 'outros_niveis_iguais': outros4, 'divergencias': d4[:3]})
    finally:
        q.fechar(salvar=False)


def main(produto, caminho, saida):
    try:
        executar(produto, caminho)
    except Exception as ex:
        import traceback
        reg('EXECUCAO', False, {'erro': repr(ex)[:300], 'traceback': traceback.format_exc()[-1200:]})
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    n_ok = sum(1 for r in RES if r['resultado'] == 'PASS')
    print(f'\n=== {produto} corrida do lote (ADR-080/081): {n_ok} PASS / {len(RES) - n_ok} FAIL ===', flush=True)
    sys.exit(0 if RES and n_ok == len(RES) else 1)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
