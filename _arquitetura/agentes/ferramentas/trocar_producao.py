# -*- coding: utf-8 -*-
"""trocar_producao.py -- termina uma entrega que parou numa suite e foi RETOMADA na mesma copia instalada
(testes/rodar_suites.py --a-partir ...): faz a troca da producao com as mesmas checagens do entregar.py.

Uso: python trocar_producao.py <produto> <pasta_da_entrega> <sha16_instalado> <log_suites_1> [<log_suites_2> ...]

Confere, nesta ordem (qualquer falha = a producao NAO e tocada):
  1. cada log de suites (o do entregar.py e o do rodar_suites.py) cobre, juntos, TODAS as SUITES do entregar.py com OK,
     e nenhum tem FALHOU;
  2. o instalado (<pasta>/<produto>/QC_x.xlsm) tem o sha256 que a instalacao imprimiu (nao mudou durante os testes);
  3. a copia da producao original (<pasta>/<produto>/producao_original) = a producao de agora (ninguem salvou a producao
     durante as horas de teste: trocar apagaria o que foi lancado);
  4. backup da producao com sha conferido; copia; sha da producao = instalado (senao devolve o backup).
Depois copia o relatorio de Westgard (ADR-083), como o entregar.py.
"""
import os
import re
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'integracao_dbseac'))
import entregar  # noqa: E402


def main(a):
    produto, pasta, sha16, logs = a[0], os.path.abspath(a[1]), a[2].lower(), a[3:]
    arq = entregar.ARQ[produto]
    trab = os.path.join(pasta, produto)
    alvo = os.path.join(trab, arq)
    original = os.path.join(trab, 'producao_original', arq)
    prod = os.path.join(entregar.RAIZ, arq)
    ok_suites = set()
    for lg in logs:
        # o log do entregar.py tem os dois produtos ("## Bioquimica" / "## Hematologia"): so vale a secao deste; o do
        # rodar_suites.py (sem secoes) e de um produto so, o do argumento
        secao = None
        for linha in open(lg, encoding='utf-8-sig', errors='replace'):
            h = re.match(r'##\s+(\w+)', linha)
            if h:
                secao = h.group(1)
                continue
            if secao not in (None, produto):
                continue
            # o FALHOU do entregar.py e substituido pelo OK da retomada; suite sem nenhum OK = nao troca
            m = re.search(r'\b(qa_\w+\.py)\s+(OK|FALHOU)\b', linha)
            if m and m.group(2) == 'OK':
                ok_suites.add(m.group(1))
    faltam = [s for s in entregar.SUITES if s not in ok_suites]
    if faltam:
        print('NAO TROCADA: suites sem OK:', faltam)
        return 1
    if not entregar.sha(alvo).startswith(sha16):
        print('NAO TROCADA: o instalado mudou (sha diferente do impresso na instalacao)')
        return 1
    if entregar.sha(prod) != entregar.sha(original):
        print('NAO TROCADA: a producao foi gravada depois do inicio da entrega (alguem salvou). Rodar de novo.')
        return 1
    bk = os.path.join(entregar.RAIZ, '_backup_pre_integracao_' + time.strftime('%Y-%m-%d_%H%M'))
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(prod, os.path.join(bk, arq))
    if entregar.sha(os.path.join(bk, arq)) != entregar.sha(prod):
        print('NAO TROCADA: backup divergente')
        return 1
    shutil.copy2(alvo, prod)
    if not entregar.sha(prod).startswith(sha16):
        shutil.copy2(os.path.join(bk, arq), prod)
        print('copia para a producao divergiu: producao devolvida do backup')
        return 1
    try:
        rel = entregar.relatorio_westgard(trab, prod)
    except OSError as e:
        rel = [f'NAO copiado ({e}); ver {trab}']
    print(f'{produto}: ENTREGUE sha256 {sha16} (backup em {os.path.basename(bk)})' + (f'; Westgard: {", ".join(rel)}' if rel else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
