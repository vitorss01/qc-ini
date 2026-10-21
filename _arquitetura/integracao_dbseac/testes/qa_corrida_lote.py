# -*- coding: utf-8 -*-
"""qa_corrida_lote.py -- ADR-080: o eixo X do grafico de cada nivel e a CORRIDA DO LOTE (1, 2, 3... por nivel).

Uso: python qa_corrida_lote.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

Oraculo independente do motor: na Principal - Resultados, os resultados PLOTADOS (participa da estatistica, ou X
vermelho) de um nivel, do lote e do equipamento do grafico, ordenados por data/hora; a posicao (1, 2, 3...) e a
corrida do lote esperada. O que o motor publicou (Eng_Saida BK..BM) e o que o Calc plota (CR..CT) tem de bater.

  L01  varredura dos analitos do lote em tela: Eng_Saida = oraculo, corrida por corrida, em todos os niveis;
  L02  em cada nivel os X plotados no grafico sao inteiros, crescentes e sem buraco (1, 2, 3...), e o eixo os enquadra;
  L03  outros lotes do mesmo analito: cada lote recomeca do 1 (o 1o resultado do lote e a corrida 1);
  L04  o caso do usuario: N1 digitado as 23:50 num dia que ja tem corridas -> vira a corrida SEGUINTE do N1 (ultima+1)
       e nenhuma corrida do outro nivel muda de numero;
  L05  os niveis sao independentes: corrida so de N2 nao deixa buraco no N1 (nem o contrario).
Nada salvo. Sai com 1 se algum falhar.
"""
import collections
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
RES = []


def reg(nome, ok, ev):
    RES.append({'teste': nome, 'resultado': 'PASS' if ok else 'FAIL', 'evidencia': ev})
    print(('PASS ' if ok else 'FAIL ') + nome + ' -- ' + json.dumps(ev, ensure_ascii=False, default=str)[:700], flush=True)


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
    return {'analito': txt(e.Range('C1').Value), 'lote': txt(e.Range('E1').Value), 'equip': txt(e.Range('S1').Value),
            'n': n, 'runs': runs, 'cor': cor, 'cortadas': e.Range('M1').Value}


def oraculo(fin, analito, lote, equip, nlv):
    """nivel -> {RUN: corrida do lote}, pela ordem de data/hora dos resultados plotados daquele nivel no lote."""
    out = []
    for t in range(1, nlv + 1):
        linhas = [r for r in fin if txt(r['ANALITO']).upper() == analito.upper() and txt(r['LOTE']) == lote
                  and (not equip or txt(r['EQUIPAMENTO']).upper() == equip.upper()) and iint(r['NIVEL']) == t
                  and num(r['RESULTADO']) and iint(r['RUN']) is not None
                  and (r['PARTICIPA_ESTATISTICA'] == 'SIM' or r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO')]
        linhas.sort(key=lambda r: (r['DATA_HORA'], str(r['ID_REGISTRO'])))
        m = {}
        for r in linhas:
            if iint(r['RUN']) not in m:
                m[iint(r['RUN'])] = len(m) + 1
        out.append(m)
    return out


def comparar(e, orc, nlv):
    dif = []
    for t in range(nlv):
        for i, run in enumerate(e['runs']):
            if e['cor'][t][i] != orc[t].get(run):
                dif.append({'nivel': t + 1, 'RUN': run, 'motor': e['cor'][t][i], 'oraculo': orc[t].get(run)})
    return dif


def plotados(q, nlv):
    """nivel -> lista dos X (Calc CR..CT) dos pontos plotados (resultado ou X vermelho), na ordem das linhas."""
    c = q.ws('Calc')
    cab = [txt(v) for v in c.Range('A2:DZ2').Value[0]]
    out = []
    for i in range(1, nlv + 1):
        xs = [r[0] for r in c.Range(f'{COL_CALC[i - 1]}3:{COL_CALC[i - 1]}182').Value]
        yl = [r[0] for r in c.Range(c.Cells(3, cab.index(f'N{i} line') + 1), c.Cells(182, cab.index(f'N{i} line') + 1)).Value]
        xx = c.Range(c.Cells(3, cab.index(f'N{i} x1') + 1), c.Cells(182, cab.index(f'N{i} x3') + 1)).Value
        out.append([xs[r] for r in range(180) if num(yl[r]) or any(num(v) for v in xx[r])])
    return out


def seq_ok(xs):
    return all(num(x) and x == math.floor(x) for x in xs) and all(b - a == 1 for a, b in zip(xs, xs[1:]))


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

        q.atualizar()
        fin = q.ler('tblCQ_Final')
        e0 = eng(q, nlv)
        lote0 = e0['lote']

        # ---- L01 + L02: varredura dos analitos do lote em tela
        difs, seqs, vistos, t0 = [], [], 0, time.time()
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
            for t, xs in enumerate(plotados(q, nlv), 1):
                if xs and not seq_ok(xs):
                    seqs.append({'analito': an, 'nivel': t, 'x': xs[:15]})
        reg('L01 Corrida do lote publicada pelo motor = oráculo (ordem de data/hora dos plotados do nível no lote), '
            'em todos os analitos do lote e todos os níveis',
            vistos > 0 and not difs, {'lote': lote0, 'analitos_conferidos': vistos, 'divergencias': difs[:5],
                                       's': round(time.time() - t0, 1)})
        reg('L02 No gráfico de cada nível os X são inteiros, crescentes e sem buraco (1, 2, 3...)',
            vistos > 0 and not seqs, {'fora_de_sequencia': seqs[:5]})

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
        reg('L03 Cada lote recomeça do 1: no outro lote a 1ª corrida de cada nível é 1 e a sequência bate com o oráculo',
            any('corridas' in x for x in ev3) and not ruins3, {'analito': an, 'lotes': ev3, 'ruins': ruins3})

        # ---- L04 + L05: o caso do usuario (N1 as 23:50 num dia que ja tem corridas)
        ir_analito(an)
        e1 = eng(q, nlv)
        orc1 = oraculo(fin, e1['analito'], e1['lote'], e1['equip'], nlv)
        alvo = [r for r in fin if txt(r['ANALITO']).upper() == an.upper() and txt(r['LOTE']) == lote0
                and r['PARTICIPA_ESTATISTICA'] == 'SIM' and isinstance(r['DATA_HORA'], dt.datetime)
                and (not e1['equip'] or txt(r['EQUIPAMENTO']).upper() == e1['equip'].upper())]
        ult = max(alvo, key=lambda r: r['DATA_HORA'])
        dia = ult['DATA_HORA'].date()
        n1 = sorted(r['RESULTADO'] for r in alvo if iint(r['NIVEL']) == 1)
        quando = dt.datetime.combine(dia, dt.time(23, 50))
        q.ex.xl.EnableEvents = True
        q.escrever_linha('tblResultados_Manuais', {
            'DATA': dt.datetime.combine(dia, dt.time()), 'HORA': (23 * 60 + 50) / 1440.0, 'LOTE': lote0, 'NIVEL': 1,
            'ANALITO': an, 'RESULTADO': n1[len(n1) // 2], 'EQUIPAMENTO': ult['EQUIPAMENTO'],
            'MOTIVO': 'QA ADR-080 corrida da noite'})
        q.atualizar()
        fin2 = q.ler('tblCQ_Final')
        man = [r for r in fin2 if r['ORIGEM_RESULTADO'] == 'MANUAL' and txt(r['ANALITO']).upper() == an.upper()
               and isinstance(r['DATA_HORA'], dt.datetime) and r['DATA_HORA'] == quando]
        ir_analito(an)
        e2 = eng(q, nlv)
        orc2 = oraculo(fin2, e2['analito'], e2['lote'], e2['equip'], nlv)
        run_m = iint(man[0]['RUN']) if man else None
        cor_m = orc2[0].get(run_m)
        n1_antes = max(orc1[0].values(), default=0)
        outros_iguais = all(orc2[t] == orc1[t] for t in range(1, nlv))
        d2 = comparar(e2, orc2, nlv)
        idx = e2['runs'].index(run_m) if run_m in e2['runs'] else None
        motor_m = e2['cor'][0][idx] if idx is not None else None
        reg('L04 N1 digitado às 23:50 num dia que já tem corridas: vira a corrida SEGUINTE do N1 (última + 1) e é '
            'isso que o motor publica',
            bool(man) and man[0]['PARTICIPA_ESTATISTICA'] == 'SIM' and cor_m == n1_antes + 1 and motor_m == cor_m and not d2,
            {'manual': man[0]['ID_REGISTRO'] if man else None, 'dia': dia, 'RUN_do_dia': run_m,
             'corrida_do_lote': cor_m, 'ultima_N1_antes': n1_antes, 'motor': motor_m, 'divergencias': d2[:3]})
        seqs2 = [{'nivel': t, 'x': xs[-8:]} for t, xs in enumerate(plotados(q, nlv), 1) if xs and not seq_ok(xs)]
        reg('L05 Níveis independentes: a corrida nova do N1 não muda nenhuma corrida dos outros níveis e nenhum '
            'gráfico fica com buraco',
            outros_iguais and not seqs2, {'outros_niveis_inalterados': outros_iguais, 'fora_de_sequencia': seqs2})
    finally:
        q.fechar(salvar=False)


def main(produto, caminho, saida):
    try:
        executar(produto, caminho)
    except Exception as ex:
        reg('EXECUCAO', False, {'erro': repr(ex)[:500]})
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    n_ok = sum(1 for r in RES if r['resultado'] == 'PASS')
    print(f'\n=== {produto} corrida do lote (ADR-080): {n_ok} PASS / {len(RES) - n_ok} FAIL ===', flush=True)
    sys.exit(0 if RES and n_ok == len(RES) else 1)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
