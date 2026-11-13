# -*- coding: utf-8 -*-
"""rodar_suites.py -- roda suites (as do entregar.py ou uma lista) sobre um arquivo JA INSTALADO, uma por vez, cada uma
numa copia nova; nao para na primeira falha (lista todas). Para iterar sem refazer a instalacao inteira.

Uso: python rodar_suites.py <produto> <instalado.xlsm> <pasta_saida> [--original <producao_original.xlsm>]
                            [--suites qa_a.py,qa_b.py] [--a-partir qa_x.py]
"""
import os
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
import entregar  # noqa: E402


def main():
    a = sys.argv[1:]
    produto, alvo, saida = a[0], os.path.abspath(a[1]), os.path.abspath(a[2])
    original = os.path.abspath(a[a.index('--original') + 1]) if '--original' in a else None
    suites = a[a.index('--suites') + 1].split(',') if '--suites' in a else list(entregar.SUITES)
    if '--a-partir' in a:
        suites = suites[suites.index(a[a.index('--a-partir') + 1]):]
    os.makedirs(saida, exist_ok=True)
    sha0 = entregar.sha(alvo)
    arq = os.path.basename(alvo)
    falhas = []
    for s in suites:
        nome = s.replace('.py', '')
        copia = os.path.join(saida, 'teste_' + nome, arq)
        os.makedirs(os.path.dirname(copia), exist_ok=True)
        shutil.copy2(alvo, copia)
        t0 = time.time()
        ok, fim = entregar.rodar(os.path.join(saida, nome + '.log'), AQUI, s, produto, copia,
                                 os.path.join(saida, nome + '.json'), original=original)
        print(f'{time.strftime("%H:%M")} {s:24s} {"OK" if ok else "FALHOU"} ({time.time() - t0:.0f}s)  '
              f'{" | ".join(fim)[:400]}', flush=True)
        if not ok:
            falhas.append(s)
    print(f'=== {produto}: {len(suites) - len(falhas)} OK / {len(falhas)} FALHOU {falhas}', flush=True)
    if entregar.sha(alvo) != sha0:
        print('ATENCAO: o arquivo instalado mudou durante os testes', flush=True)


if __name__ == '__main__':
    main()
