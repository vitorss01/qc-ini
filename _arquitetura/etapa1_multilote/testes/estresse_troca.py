# -*- coding: utf-8 -*-
"""estresse_troca.py -- troca de lote/analito repetida, passo a passo registrado.

Uso: python estresse_troca.py <base.xlsm> <n_ciclos> [original]

Prepara 3 lotes (como o T3) e troca de lote e de analito n vezes, gravando
cada passo num log em disco ANTES de executa-lo -- se o Excel cair, a ultima
linha diz onde. Mede o tempo de cada troca.
"""
import datetime as dt
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import suite_etapa1 as S  # noqa: E402

LOG = os.path.join(os.environ['TMP'], 'qcw', 'estresse.log')


def log(s):
    with open(LOG, 'a', encoding='utf-8') as f:
        f.write(time.strftime('%H:%M:%S ') + s + '\n')
    print(s, flush=True)


def main(base, n):
    open(LOG, 'w').close()
    s = S.Sessao(base, 'estresse')
    try:
        log('preparar')
        S.preparar(s, [S.A, S.B, S.C])
        s.analito(S.AN)
        seq = ['9001', '9002', '9003', '9001', '9003', '9002']
        ans = ['Lactato', 'Glicose', 'Lactato']
        tempos = []
        for c in range(n):
            for l in seq:
                log(f'ciclo {c} selecionar {l}')
                t0 = time.time(); s.selecionar(l); tempos.append(time.time() - t0)
            for a in ans:
                log(f'ciclo {c} analito {a}')
                s.analito(a)
        L = s.ler()
        log(f'OK {n} ciclos; troca de lote media {sum(tempos)/len(tempos):.2f}s max {max(tempos):.2f}s; par={L["par"]}')
    finally:
        s.fechar()


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), int(sys.argv[2]))
