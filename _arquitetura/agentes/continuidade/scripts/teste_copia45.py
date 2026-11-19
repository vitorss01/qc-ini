# -*- coding: utf-8 -*-
"""Pre-teste da entrega 44 numa COPIA da producao: so os instaladores tocados + suites afetadas. Nada na producao."""
import hashlib
import os
import shutil
import subprocess
import sys
import time

QC = sys.argv[1]
PASTA = sys.argv[2]
PRODUTOS = sys.argv[3:] or ['Bioquimica', 'Hematologia']
D = os.path.join(QC, '_arquitetura', 'integracao_dbseac')
ARQ = {'Hematologia': 'QC_Hematologia.xlsm', 'Bioquimica': 'QC_Bioquimica.xlsm'}
PASSOS = [['instalar_adr092.py', '{p}', '{a}']]
SUITES = ['qa_adr097.py', 'qa_manual_substitui.py', 'qa_adr095.py', 'qa_adr090.py', 'qa_casos_extremos.py',
          'qa_inativar.py', 'qa_etl.py', 'qa_equiv_qa_adr094.py']


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def rodar(log, cwd, args, original):
    env = dict(os.environ, PYTHONIOENCODING='utf-8', QC_PROD_ORIGINAL=original)
    t0 = time.time()
    with open(log, 'w', encoding='utf-8') as f:
        r = subprocess.run([sys.executable, '-u', '-W', 'ignore', *args], cwd=cwd, env=env, stdout=f,
                           stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace')
    with open(log, encoding='utf-8', errors='replace') as f:
        fim = [l.rstrip() for l in f if l.startswith(('===', 'FAIL', 'INCONCLUSIVO', 'Traceback', 'VBA NAO', 'SystemExit'))
               or 'nada foi salvo' in l]
    return r.returncode == 0, fim, round((time.time() - t0) / 60, 1)


for p in PRODUTOS:
    prod = os.path.join(QC, ARQ[p])
    trab = os.path.join(PASTA, p)
    os.makedirs(trab, exist_ok=True)
    alvo = os.path.join(trab, ARQ[p])
    orig = os.path.join(trab, 'original', ARQ[p])
    os.makedirs(os.path.dirname(orig), exist_ok=True)
    shutil.copy2(prod, alvo)
    shutil.copy2(prod, orig)
    s0 = sha(prod)
    print(f'## {p} producao {s0[:16]}', flush=True)
    ok_tudo = True
    for i, passo in enumerate(PASSOS):
        args = [x.format(p=p, a=alvo) for x in passo]
        ok, fim, mn = rodar(os.path.join(trab, f'{i:02d}_{passo[0]}.log'), D, args, orig)
        print(f'   {passo[0]:28s} {"OK" if ok else "FALHOU"} {mn} min  {" | ".join(fim)[:300]}', flush=True)
        if not ok:
            ok_tudo = False
            break
    if ok_tudo:
        for s in SUITES:
            copia = os.path.join(trab, 'teste_' + s[:-3], ARQ[p])
            os.makedirs(os.path.dirname(copia), exist_ok=True)
            shutil.copy2(alvo, copia)
            ok, fim, mn = rodar(os.path.join(trab, s[:-3] + '.log'), os.path.join(D, 'testes'),
                                [s, p, copia, os.path.join(trab, s[:-3] + '.json')], orig)
            print(f'   {s:28s} {"OK" if ok else "FALHOU"} {mn} min  {" | ".join(fim)[:400]}', flush=True)
            try:
                os.remove(copia)
            except OSError:
                pass
    print(f'   producao intacta: {sha(prod) == s0}', flush=True)
print('FIM', flush=True)
