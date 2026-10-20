# -*- coding: utf-8 -*-
"""ATUALIZAR DADOS com MOTOR_DADOS = VBA (ADR-079) numa COPIA ja instalada (instalar_adr079.py): N execucoes
seguidas (tempo de cada uma e a linha do Sync_Log) e, no fim, a RECONCILIACAO -- o Power Query refaz a tblCQ_Final e
a tblQA_Integracao e as duas tem de sair iguais, celula a celula, ao que o VBA publicou. Nada salvo.

Uso: python teste_sync.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json> [--n 3] [--sem-reconciliar]
"""
import json
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA  # noqa: E402


def valores(lo):
    cab = [c.Name for c in lo.ListColumns]
    if lo.ListRows.Count == 0:
        return cab, []
    v = lo.DataBodyRange.Value2
    return cab, [list(r) for r in v]


def norm(x):
    if x in ('', None):
        return None
    if isinstance(x, float):
        return round(x, 9)
    return x


def comparar(a, b):
    """a, b = (cabecalho, linhas). Por nome de coluna; devolve (n_dif, amostras)."""
    ca, la = a
    cb, lb = b
    if sorted(ca) != sorted(cb):
        return -1, [f'colunas diferentes: {set(ca) ^ set(cb)}']
    ib = [cb.index(c) for c in ca]
    n, am = 0, []
    if len(la) != len(lb):
        am.append(f'linhas {len(la)} x {len(lb)}')
        n += abs(len(la) - len(lb))
    for i, (ra, rb) in enumerate(zip(la, lb)):
        for j, c in enumerate(ca):
            x, y = norm(ra[j]), norm(rb[ib[j]])
            if x != y and not (isinstance(x, float) and isinstance(y, float) and abs(x - y) < 1e-8):
                n += 1
                if len(am) < 10:
                    am.append(f'linha {i + 1} {c}: {x!r} x {y!r}')
    return n, am


def main(produto, arq, saida, n=3, reconciliar=True):
    res = {'produto': produto, 'execucoes': []}
    q = QA(produto, arq)
    try:
        res['motor'] = q.cfg('MOTOR_DADOS')
        for k in range(n):
            r, s = q.atualizar()
            tempo = [x for x in r.split('\n') if x.startswith('Tempo:')]
            novos = [x for x in r.split('\n') if 'novos nesta' in x]
            ult = [x for x in r.split('\n') if 'mais recente' in x]
            res['execucoes'].append({'s': round(s, 2), 'tempo': tempo[0] if tempo else '', 'novos': novos[0] if novos else '',
                                     'mais_recente': ult[0] if ult else ''})
            print(f'execucao {k + 1}: {s:.1f}s | {tempo[0] if tempo else r[:300]} | {novos[0] if novos else ""}', flush=True)
        cab, linhas = valores(q.lo('tblSyncLog'))
        res['sync_log'] = [dict(zip(cab, l)) for l in linhas[-n:]]
        for l in res['sync_log']:
            print('  log:', {k: l[k] for k in ('RESULTADO', 'NOVOS', 'TOTAL_RECEBIDO', 'LINHAS_CQ', 'GRAVADAS', 'T_ORIGEM',
                                               'T_RECEBIMENTO', 'T_CQ', 'T_GRAVAR', 'T_QA', 'T_MOTOR', 'TOTAL_S')}, flush=True)
        if reconciliar:
            vba_cq, vba_qa = valores(q.lo('tblCQ_Final')), valores(q.lo('tblQA_Integracao'))
            t0 = time.time()
            q.run('mIntegracao.Refrescar', 'tblCQ_Final', teto=1500)
            q.run('mIntegracao.Refrescar', 'tblQA_Integracao', teto=1500)
            res['reconciliar_pq_s'] = round(time.time() - t0, 1)
            pq_cq, pq_qa = valores(q.lo('tblCQ_Final')), valores(q.lo('tblQA_Integracao'))
            res['cq'] = comparar(vba_cq, pq_cq)
            # QA: dentro de um codigo o PQ desempata de forma instavel -- conjunto
            ca, la = vba_qa
            cb, lb = pq_qa
            ib = [cb.index(c) for c in ca] if sorted(ca) == sorted(cb) else None
            if ib is None:
                res['qa'] = 'colunas diferentes'
            else:
                sa = sorted(map(lambda r: tuple(map(lambda v: str(norm(v)), r)), la))
                sb = sorted(map(lambda r: tuple(str(norm(r[j])) for j in ib), lb))
                res['qa'] = {'vba': len(sa), 'pq': len(sb), 'iguais': sa == sb}
            print('reconciliacao CQ (diferencas, amostras):', res['cq'], flush=True)
            print('reconciliacao QA:', res['qa'], f"(PQ {res['reconciliar_pq_s']}s)", flush=True)
    finally:
        q.fechar(salvar=False)
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1, default=str)


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], os.path.abspath(a[1]), a[2], int(a[a.index('--n') + 1]) if '--n' in a else 3,
         '--sem-reconciliar' not in a)
