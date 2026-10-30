# -*- coding: utf-8 -*-
"""QA do ADR-086 (29/10/2026): Painel H3 vazio = TODOS os lotes do periodo. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_adr086.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Cenario: o analito/equipamento com resultados ativos em DOIS lotes (o mais recente e o anterior), periodo cobrindo os dois.
  Y01 H3 apagado (como o usuario): sem erro, H3 continua vazio, engTodos = VERDADEIRO; o grafico tem pontos dos dois
      lotes, todos ativos do analito/equipamento ate o fim do periodo
  Y02 cada ponto com o alvo (Media/DP) do SEU lote (Eng_Saida BW..)
  Y03 Westgard lote a lote: as marcas de cada ponto = as da visao do proprio lote (H3 = o lote), nos pontos com 10
      anteriores do mesmo lote na janela -- nenhuma regra atravessa a troca de lote
  Y04 corrida 1, 2, 3... por nivel, sem buraco, em ordem cronologica, atravessando os lotes
  Y05 linhas de limite: no modo TODOS em degrau (um par por ponto, y = Media + k*DP do lote do ponto; bordas NA); com
      um lote em H3, exatamente as 2 bordas do ADR-078
  Y06 Painel: H4 diz TODOS; media/DP/CV vazios (nao juntam lotes); n e veredito presentes
  Y07 a ficha do duplo clique mostra o lote do PROPRIO ponto
  Y08 abrir o arquivo (SincronizarLotesAoAbrir) nao preenche H3; o motor reconhece a troca de modo (GarantirMotorDoPainel)
  Y09 tempos medidos do recalculo nos dois modos (sem meta inventada)
  Y10 de volta a um lote: modo normal identico ao de antes (mesmo Eng_Saida)
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import collections
import datetime as dt
import json
import os
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

LIM = ['-3s', '-2s', '-1s', 'Média', '+1s', '+2s', '+3s']
KS = [-3, -2, -1, 0, 1, 2, 3]
C_ID0, C_CORL0, C_ALVO0, EF0, NEF = 39, 63, 75, 3, 7


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def serie(d):
    return float((d.replace(tzinfo=None) - dt.datetime(1899, 12, 30)).days)


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    tempos = {}
    try:
        xl = q.ex.xl
        nlv = q.nlv
        p, c, e = q.ws('Painel'), q.ws('Calc'), q.ws('Eng_Saida')
        q.run('mSeguranca.ReprotectAll')
        rows = q.ler('tblCQ_Final')
        cad = [txt(x[0]) for x in q.espec().Range('A4:A43').Value]
        # cenario: (analito, equipamento) com N1 ativo em >= 2 lotes; os 2 lotes mais recentes
        g = collections.defaultdict(lambda: collections.defaultdict(list))
        for r in rows:
            if r['STATUS_ANALITICO'] == 'ATIVO' and txt(r['NIVEL']) == '1' and txt(r['ANALITO']) in cad and r['DATA_HORA']:
                g[(txt(r['ANALITO']), txt(r['EQUIPAMENTO']))][txt(r['LOTE'])].append(r['DATA_HORA'])
        cand = []
        eq_hema = None if q.bio else txt(q.nome('selEquipamento').Value)    # Hematologia: equipamento fixo (Cfg B11)
        for (an, eq), lts in g.items():
            if eq_hema and eq.upper() != eq_hema.upper():
                continue
            ok = {l: v for l, v in lts.items() if len(v) >= 12}
            if len(ok) >= 2:
                ult = sorted(ok, key=lambda l: max(ok[l]))[-2:]
                cand.append((min(len(ok[l]) for l in ult), an, eq, ult, ok))
        if not cand:
            raise RuntimeError('pré-condição: nenhum analito/equipamento com N1 ativo em 2 lotes (>= 12 cada)')
        cand.sort(reverse=True)
        _, an, eq, (l_old, l_new), ok = cand[0]
        d_ini = min(ok[l_old])
        d_fim = max(ok[l_new])

        def tela(lote_h3):
            q.run('mLotes.EsquecerFiltroRotina', teto=60)      # ADR-078: nenhum filtro de rotina guardado mexe no periodo
            xl.EnableEvents = False
            p.Unprotect(ii.SENHA)
            p.Range('B3').Value = cad.index(an) + 1
            if q.bio:
                p.Range('L4').Value = eq
            p.Range('G3').Value = serie(d_ini)
            p.Range('G4').Value = serie(d_fim)
            xl.EnableEvents = True
            t0 = time.time()
            if lote_h3:
                p.Range('H3').Value = "'" + lote_h3            # como o usuario: eventos ligados
            else:
                p.Range('H3').MergeArea.ClearContents()   # H3 e mesclada (como o Delete do usuario)
            q.ex.esperar()
            q.run('mEstatistica.RecalcularAnalitoAtual', teto=600)
            xl.Calculate()
            return round(time.time() - t0, 2)

        def eng():
            n = int(e.Range('I1').Value or 0)
            out = []
            if not n:
                return out
            blk = e.Range(e.Cells(3, 1), e.Cells(2 + n, 80)).Value
            for k, r in enumerate(blk):
                out.append({'slot': k + 1, 'run': r[1], 'id': [txt(r[C_ID0 - 1 + t]) for t in range(nlv)],
                            'cor': [r[C_CORL0 - 1 + t] for t in range(nlv)],
                            'alvo': [(r[C_ALVO0 - 1 + 2 * t], r[C_ALVO0 + 2 * t]) for t in range(nlv)],
                            'flags': [tuple(r[EF0 - 1 + t * NEF:EF0 - 1 + t * NEF + 6]) for t in range(nlv)]})
            return out

        def flags_por_id(lst):
            return {x['id'][t]: x['flags'][t] for x in lst for t in range(nlv) if x['id'][t]}

        def pares_limite(i):
            ch = p.ChartObjects(i).Chart
            out = {}
            for s_i in range(1, ch.SeriesCollection().Count + 1):
                s = ch.SeriesCollection(s_i)
                if str(s.Name) in LIM:
                    out[str(s.Name)] = [(x, y) for x, y in zip(s.XValues, s.Values) if num(x) and num(y)]
            return out

        cq = {txt(r['ID_REGISTRO']).upper(): r for r in rows}
        tempos['lote_novo_s'] = tela(l_new)
        vA = eng()
        parA = pares_limite(1)
        tela(l_old)
        vB = eng()
        tempos['todos_s'] = tela('')
        vT = eng()
        h3 = p.Range('H3').Value
        todos = e.Range('AD1').Value
        lotes_t = collections.Counter(txt(cq[x['id'][0].upper()]['LOTE']) for x in vT if x['id'][0])
        fora = [x['id'][0] for x in vT if x['id'][0] and (cq[x['id'][0].upper()]['STATUS_ANALITICO'] != 'ATIVO'
                or txt(cq[x['id'][0].upper()]['ANALITO']) != an
                or ((q.bio or eq_hema) and txt(cq[x['id'][0].upper()]['EQUIPAMENTO']).upper() != eq.upper())
                or cq[x['id'][0].upper()]['DATA_HORA'] > d_fim)]
        reg('Y01 H3 apagado como o usuário: sem erro, H3 continua vazio e o modo é TODOS; o gráfico tem pontos dos dois '
            'lotes, só resultados ativos do analito/equipamento até o fim do período',
            h3 in (None, '') and todos is True and lotes_t.get(l_old, 0) > 0 and lotes_t.get(l_new, 0) > 0 and not fora,
            {'analito': an, 'equipamento': eq, 'lotes': [l_old, l_new], 'pontos_N1_por_lote': dict(lotes_t),
             'H3': h3, 'engTodos': todos, 'fora_do_escopo': fora[:5]})

        # Y02 alvo do lote de cada ponto
        y2 = []
        alvo_c = {}
        for x in vT:
            for t in range(nlv):
                if not x['id'][t]:
                    continue
                lt = txt(cq[x['id'][t].upper()]['LOTE'])
                if (lt, t) not in alvo_c:
                    alvo_c[(lt, t)] = (q.run('mLotes.MediaDoLote', lt, an, t + 1), q.run('mLotes.DPDoLote', lt, an, t + 1))
                m_, s_ = alvo_c[(lt, t)]
                got = x['alvo'][t]
                valido = num(m_) and num(s_) and s_ > 0
                okp = ((valido and num(got[0]) and num(got[1]) and abs(got[0] - m_) < 1e-9 and abs(got[1] - s_) < 1e-9)
                       or (not valido and got[0] in ('', None) and got[1] in ('', None)))
                if not okp:
                    y2.append({'id': x['id'][t], 'lote': lt, 'nivel': t + 1, 'eng': got, 'lote_md': (m_, s_)})
        reg('Y02 Cada ponto com o alvo (Média/DP) do SEU lote no modo TODOS (Eng_Saida BW..)',
            not y2 and len(alvo_c) >= 2, {'divergencias': y2[:8], 'alvos': {f'{k[0]} N{k[1] + 1}': v for k, v in alvo_c.items()}})

        # Y03 Westgard lote a lote
        fA, fB, fT = flags_por_id(vA), flags_por_id(vB), flags_por_id(vT)
        seq = collections.defaultdict(list)
        for x in vT:
            for t in range(nlv):
                if x['id'][t]:
                    seq[(txt(cq[x['id'][t].upper()]['LOTE']), t)].append(x['id'][t])
        comparados, difs = 0, []
        for (lt, t), ids in seq.items():
            ref = fA if lt == l_new else fB if lt == l_old else None
            if ref is None:
                continue
            for idx in ids[10:]:
                if idx in ref:
                    comparados += 1
                    if tuple(ref[idx]) != tuple(fT[idx]):
                        difs.append({'id': idx, 'lote': lt, 'nivel': t + 1, 'visao_do_lote': ref[idx], 'todos': fT[idx]})
        reg('Y03 Westgard lote a lote: as marcas de cada ponto no modo TODOS = as da visão do próprio lote (pontos com 10 '
            'anteriores do mesmo lote) -- nenhuma regra atravessa a troca de lote',
            comparados > 0 and not difs, {'comparados': comparados, 'diferencas': difs[:8]})

        # Y04 corrida atravessando os lotes
        y4 = {}
        for t in range(nlv):
            cs = [x['cor'][t] for x in vT if num(x['cor'][t])]
            y4[f'N{t + 1}'] = (len(cs), cs[:3], cs[-3:])
            if not cs or any(abs(b - a - 1) > 1e-9 for a, b in zip(cs, cs[1:])):
                y4['ruim'] = t + 1
        dts = [cq[x['id'][0].upper()]['DATA_HORA'] for x in vT if x['id'][0]]
        reg('Y04 Corrida 1, 2, 3... por nível, sem buraco, em ordem cronológica, atravessando os lotes',
            'ruim' not in y4 and dts == sorted(dts), {'corridas': y4})

        # Y05 linhas em degrau x bordas
        parT = pares_limite(1)
        # so os pontos que o GRAFICO mostra (uma serie de pontos com y numerico): corrida antes do inicio do periodo
        # entra no Eng_Saida (contexto do Westgard) mas nao e desenhada -- nem ela nem a sua linha de limite
        ch1 = p.ChartObjects(1).Chart
        vistos = set()
        for s_i in range(1, ch1.SeriesCollection().Count + 1):
            s = ch1.SeriesCollection(s_i)
            if str(s.Name) in LIM:
                continue
            try:
                vistos |= {x for x, y in zip(s.XValues, s.Values) if num(x) and num(y)}
            except Exception:                              # noqa: BLE001
                pass
        ptsN1 = {x['cor'][0]: x['alvo'][0] for x in vT if num(x['cor'][0]) and x['cor'][0] in vistos}
        y5 = []
        for nm, k in zip(LIM, KS):
            pr = parT.get(nm, [])
            if len(pr) != len(ptsN1):
                xs = {xv for xv, _ in pr}
                falta = sorted(x for x in ptsN1 if x not in xs)[:6]
                s_ = next(s for s in (p.ChartObjects(1).Chart.SeriesCollection(j) for j in
                                      range(1, p.ChartObjects(1).Chart.SeriesCollection().Count + 1)) if str(s.Name) == nm)
                y5.append((nm, 'pares', len(pr), len(ptsN1), 'faltam corridas', falta,
                           'serie', str(s_.Formula)[:120], 'n_valores', len(s_.Values), len(s_.XValues)))
                continue
            for xv, yv in pr:
                m_, s_ = ptsN1.get(xv, (None, None))
                if not (num(m_) and abs(yv - (m_ + k * s_)) <= 1e-9 * max(1, abs(m_))):
                    y5.append((nm, xv, yv, m_, s_))
                    break
        normal_ok = all(len(v) == 2 for v in parA.values()) and len(parA) == 7
        reg('Y05 Linhas de limite: no modo TODOS em degrau (um par por ponto, Média + k·DP do lote do ponto; bordas NA); '
            'com um lote em H3, exatamente as 2 bordas (ADR-078)',
            not y5 and normal_ok, {'problemas_todos': y5[:6], 'pares_modo_normal': {k: len(v) for k, v in parA.items()},
                                   'pontos_N1': len(ptsN1)})

        # Y06 Painel
        h4 = txt(p.Range('H4').Value)
        ind = [[p.Range(f'{col}{6 + t}').Value for col in 'BCDE'] for t in range(1, nlv + 1)]
        reg('Y06 Painel no modo TODOS: H4 diz TODOS; média/DP/CV vazios (não juntam lotes); n presente',
            'TODOS' in h4 and all(num(r[0]) and r[0] > 0 and r[1] in (None, '') and r[2] in (None, '') and r[3] in (None, '')
                                  for r in ind[:1]),
            {'H4': h4, 'B..E por nivel': ind})

        # Y07 ficha do ponto do lote antigo
        nomes = [str(s.Name) for s in p.ChartObjects(1).Chart.SeriesCollection()]
        s_res = nomes.index('Resultado') + 1
        px = next(x for x in vT if x['id'][0] and txt(cq[x['id'][0].upper()]['LOTE']) == l_old)
        ficha = str(q.run('mUI.FichaPonto', 1, s_res, px['slot']))
        reg('Y07 A ficha do duplo clique mostra o lote do PRÓPRIO ponto no modo TODOS',
            f'Lote {l_old}' in ficha and f'ID {px["id"][0].split("-")[-1]}' in ficha, {'ficha': ficha})

        # Y08 abrir nao preenche H3; troca de modo reconhecida
        q.run('mLotes.SincronizarLotesAoAbrir', teto=120)
        h3b = p.Range('H3').Value
        fresco = q.run('mEstatistica.GarantirMotorDoPainel', teto=300)
        xl.EnableEvents = False
        p.Unprotect(ii.SENHA)
        p.Range('H3').Value = "'" + l_new                       # sem evento: o motor fica para tras
        xl.EnableEvents = True
        atras = q.run('mEstatistica.GarantirMotorDoPainel', teto=600)
        reg('Y08 Abrir o arquivo não preenche H3 (vazio = TODOS fica salvo); o motor reconhece a troca de modo',
            h3b in (None, '') and fresco is False and atras is True, {'H3': h3b, 'fresco': fresco, 'atrasado': atras})

        # Y10 de volta a um lote
        tempos['lote_novo_de_novo_s'] = tela(l_new)
        vA2 = eng()
        reg('Y10 De volta a um lote em H3: modo normal, Eng_Saida igual ao de antes (mesmos pontos, corridas, alvos e '
            'marcas) e as 2 bordas das linhas de limite',
            vA2 == vA and e.Range('AD1').Value is False and all(len(v) == 2 for v in pares_limite(1).values()),
            {'slots': len(vA2)})
        reg('Y09 Tempos medidos (s) do recálculo -- lote único x TODOS (registrados, sem meta inventada)', True,
            {'tempos_s': tempos, 'slots_lote': len(vA), 'slots_todos': len(vT)})
    except Exception as ex_:                         # noqa: BLE001
        reg('Y99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-086 (H3 vazio = todos os lotes): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
