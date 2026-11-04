# -*- coding: utf-8 -*-
"""regressao_pos_atualizar.py -- Tester/QA, reteste01 (ADR-088). A entrega 35 rodou as 28 suites sobre o arquivo
instalado SEM o 1o ATUALIZAR (a tblCQ_Final so muda nele). Aqui: copia da producao -> AtualizarSoEntradas -> salva SO a
copia -> roda as suites do fluxo vizinho sobre ela, uma por vez (cada suite abre e fecha o seu Excel).
Uso: python -u regressao_pos_atualizar.py <Bioquimica|Hematologia> <suite1.py,suite2.py,...> [original_pre_instalacao]
"""
import os
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

PROD, SUITES = sys.argv[1], sys.argv[2].split(',')
ORIG = sys.argv[3] if len(sys.argv) > 3 else None
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste01', 'regressao_' + PROD)
os.makedirs(PASTA, exist_ok=True)
src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
base = os.path.join(PASTA, f'QC_{PROD}.xlsm')
h0 = C.sha(src)
if not os.path.exists(base):
    shutil.copy2(src, base)
    qa = C.qa_final.QA(PROD, base)
    print('EXCEL_PID', qa.ex.pid, flush=True)
    try:
        t0 = time.time()
        print('ATUALIZAR', str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))[:200], round(time.time() - t0, 1),
              flush=True)
        n = sum(1 for r in C.final(qa) if r['STATUS_ANALITICO'] == 'IMPLAUSIVEL')
        print('IMPLAUSIVEL', n, flush=True)
        # eventos desligados: com eles o Workbook_BeforeSave cancela o Save em silencio (como o instalar_adr088 faz)
        qa.ex.xl.EnableEvents = False
        antes = os.path.getmtime(base)
        qa.ex.wb.Save()                 # salva SO a copia em %TEMP%
        if os.path.getmtime(base) == antes:
            raise SystemExit('Save nao chegou a copia')
        qa.ex.xl.EnableEvents = True
        if PROD == 'Bioquimica':        # OC-0009 crit. 2 (depois do Save; nada mais e gravado)
            est = qa.ws('Estatística')
            est.Unprotect(C.senha_planilha())
            print('B98:B103 CAP', [est.Range(f'B{i}').Value for i in range(98, 104)], flush=True)
            C.por(qa, est, 'L4', 'Controllab', local=False)
            g = [est.Cells(i, 7).Value for i in range(14, 94) if est.Cells(i, 1).Value not in (None, '', 0)]
            print('B103 Controllab', est.Range('B103').Value, 'indep SEM EP', sum(1 for x in g if x == 'SEM EP'),
                  'linhas com analito', len(g), flush=True)
    finally:
        qa.fechar()
print('sha_prod', h0, 'base_pos_atualizar', C.sha(base), flush=True)
env = dict(os.environ, PYTHONIOENCODING='utf-8')
if ORIG:
    env['QC_PROD_ORIGINAL'] = ORIG
testes = os.path.join(C.REPO, '_arquitetura', 'integracao_dbseac', 'testes')
for s in SUITES:
    d = os.path.join(PASTA, 'teste_' + s[:-3])
    os.makedirs(d, exist_ok=True)
    cp = os.path.join(d, f'QC_{PROD}.xlsm')
    shutil.copy2(base, cp)
    log = os.path.join(PASTA, s[:-3] + '.log')
    t0 = time.time()
    with open(log, 'w', encoding='utf-8') as f:
        r = subprocess.run([sys.executable, '-u', '-W', 'ignore', s, PROD, cp, os.path.join(PASTA, s[:-3] + '.json')],
                           cwd=testes, env=env, stdout=f, stderr=subprocess.STDOUT, text=True, encoding='utf-8',
                           errors='replace')
    fim = [l.rstrip() for l in open(log, encoding='utf-8', errors='replace') if l.startswith(('===', 'FAIL', 'Traceback'))]
    print(s, 'rc', r.returncode, round(time.time() - t0), 's |', ' | '.join(fim)[:600], flush=True)
print('sha_prod_depois', C.sha(src), flush=True)
