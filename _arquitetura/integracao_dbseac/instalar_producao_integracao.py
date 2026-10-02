# -*- coding: utf-8 -*-
"""instalar_producao_integracao.py -- ENTREGA do ADR-057 nos arquivos de PRODUCAO.

Uso:  python instalar_producao_integracao.py [Hematologia|Bioquimica ...]

Para cada produto:
  1. copia o .xlsm para _backup_pre_integracao_<data>/ e confere byte a byte (SHA-256);
  2. gera o VBA a partir do PROPRIO arquivo (patch_vba.py) -- nada de fonte velha;
  3. instala (instalar_integracao.py): camada de dados, VBA, Calc, Estatistica,
     Registros, remocao do legado, navegacao, atualizacao completa, X vermelho;
  4. se qualquer passo falhar, devolve o arquivo do backup e para.

Os mesmos scripts validados em copia por testes/qa_final.py.
"""
import hashlib
import os
import shutil
import subprocess
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..', '..'))
ARQ = {'Hematologia': 'QC_Hematologia.xlsm', 'Bioquimica': 'QC_Bioquimica.xlsm'}


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def rodar(*args):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    r = subprocess.run([sys.executable, '-u', '-W', 'ignore', *args], cwd=AQUI, env=env,
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    print(r.stdout[-6000:], flush=True)
    if r.returncode != 0:
        print(r.stderr[-4000:], flush=True)
    return r.returncode == 0


def main(produtos):
    bk = os.path.join(RAIZ, '_backup_pre_integracao_' + time.strftime('%Y-%m-%d_%H%M'))
    os.makedirs(bk, exist_ok=True)
    for produto in produtos:
        p = os.path.join(RAIZ, ARQ[produto])
        d = os.path.join(bk, ARQ[produto])
        shutil.copy2(p, d)
        if sha(p) != sha(d):
            raise SystemExit(f'backup divergente: {p}')
        print(f'backup: {d}  sha256={sha(d)[:16]}', flush=True)
        ok = rodar('patch_vba.py', produto, p) and rodar('instalar_integracao.py', produto, p)
        if not ok:
            shutil.copy2(d, p)
            raise SystemExit(f'*** {produto}: FALHOU -- arquivo de producao devolvido do backup ***')
        print(f'{produto}: instalado. sha256={sha(p)[:16]}', flush=True)
    print('\nENTREGUE. Backups em:', bk)


if __name__ == '__main__':
    main(sys.argv[1:] or ['Hematologia', 'Bioquimica'])
