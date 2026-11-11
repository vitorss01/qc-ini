# -*- coding: utf-8 -*-
"""regressao_reteste03.py -- Tester/QA, ciclo 2026-11-11_reteste03 (VALIDAR ADR-090/091). Regressao sobre COPIAS da
producao atual (ja entregue e ja ATUALIZADA em 11/11): cada suite recebe uma copia nova (shutil.copy2), roda no seu
proprio Excel (uma por vez) e grava log/JSON na pasta do ciclo. qa_rodadas_cv recebe QC_PROD_ORIGINAL = outra copia da
mesma producao (o R00 vira auto-comparacao: so prova que a suite nao muda nada; os R01..R12 sao os que contam).
Uso: python -u regressao_reteste03.py <Bioquimica|Hematologia> <suite1.py,...> [--equivalencia]
"""
import os
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

PROD, SUITES = sys.argv[1], [s for s in sys.argv[2].split(',') if s]
EQUIV = '--equivalencia' in sys.argv
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-11_reteste03', 'regressao', PROD)
os.makedirs(PASTA, exist_ok=True)
src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
h0 = C.sha(src)
orig = os.path.join(PASTA, f'original_QC_{PROD}.xlsm')
if not os.path.exists(orig):
    shutil.copy2(src, orig)
env = dict(os.environ, PYTHONIOENCODING='utf-8', QC_PROD_ORIGINAL=orig)
base = os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac')
testes = os.path.join(base, 'testes')
print('sha_prod', h0, flush=True)


def rodar(nome, cwd, args):
    d = os.path.join(PASTA, 'copia_' + nome)
    os.makedirs(d, exist_ok=True)
    cp = os.path.join(d, f'QC_{PROD}.xlsm')
    shutil.copy2(src, cp)
    log = os.path.join(PASTA, nome + '.log')
    t0 = time.time()
    with open(log, 'w', encoding='utf-8') as f:
        r = subprocess.run([sys.executable, '-u', '-W', 'ignore'] + [a.replace('{COPIA}', cp) for a in args],
                           cwd=cwd, env=env, stdout=f, stderr=subprocess.STDOUT, text=True, encoding='utf-8',
                           errors='replace')
    fim = [l.rstrip() for l in open(log, encoding='utf-8', errors='replace')
           if l.startswith(('===', 'FAIL', 'Traceback', 'ADR-0', 'status PQ', '{"linhas"'))]
    print(nome, 'rc', r.returncode, round(time.time() - t0), 's |', ' | '.join(fim)[:900], flush=True)


for s in SUITES:
    rodar(s[:-3], testes, [s, PROD, '{COPIA}', os.path.join(PASTA, s[:-3] + '.json')])
if EQUIV:
    rodar('equivalencia_cq', os.path.join(base, 'estudos'),
          ['equivalencia_cq.py', PROD, '{COPIA}', os.path.join(PASTA, 'equivalencia_cq.json'), '--n', '1', '--fixtures'])
print('sha_prod_depois', C.sha(src), 'igual', C.sha(src) == h0, flush=True)
