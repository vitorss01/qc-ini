# -*- coding: utf-8 -*-
"""instalar_producao.py -- entrega: instala Etapa 1 + ADR-050/051/052 nos arquivos de PRODUCAO.

Uso:  python instalar_producao.py

1. copia QC_Bioquimica.xlsm e QC_Hematologia.xlsm para _backup_pre_5anos_<data>/
   (e confere que a copia e identica byte a byte);
2. instala na Bioquimica (aplicar_etapa1.py) e porta a Hematologia
   (portar_hema.py) -- os MESMOS scripts validados por testes/validar_tudo.sh;
3. se qualquer passo falhar, devolve o arquivo do backup.

A Imunologia nao e tocada.
"""
import hashlib
import os
import shutil
import subprocess
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..', '..'))
REF_H1 = os.path.join(AQUI, 'ref', 'QC_Hematologia_build_h1.xlsm')


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def rodar(*args):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    r = subprocess.run([sys.executable, *args], cwd=AQUI, env=env, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    print(r.stdout[-3000:])
    if r.returncode != 0:
        print(r.stderr[-3000:])
    return r.returncode == 0


def main():
    bio = os.path.join(RAIZ, 'QC_Bioquimica.xlsm')
    hema = os.path.join(RAIZ, 'QC_Hematologia.xlsm')
    bk = os.path.join(RAIZ, '_backup_pre_5anos_' + time.strftime('%Y-%m-%d_%H%M'))
    os.makedirs(bk, exist_ok=True)
    copias = {}
    for p in (bio, hema):
        d = os.path.join(bk, os.path.basename(p))
        shutil.copy2(p, d)
        if sha(p) != sha(d):
            raise SystemExit(f'backup divergente: {p}')
        copias[p] = d
        print('backup:', d)
    # Por padrao a instalacao roda SOBRE o arquivo de producao atual, preservando o
    # que o usuario lancou desde a ultima instalacao (os instaladores sao idempotentes:
    # cada passo confere se ja esta feito antes de refazer). QC_ORIGINAL existe para
    # reinstalar do zero a partir de uma copia intocada -- e ai o que foi lancado
    # depois daquela copia NAO vem junto.
    orig = os.environ.get('QC_ORIGINAL')
    if orig:
        for p in (bio, hema):
            shutil.copy2(os.path.join(orig, os.path.basename(p)), p)
        print('reinstalando a partir de', orig)
    ok = rodar('aplicar_etapa1.py', bio)
    if ok:
        ok = rodar('portar_hema.py', hema, REF_H1, bio)
    if not ok:
        for p, d in copias.items():
            shutil.copy2(d, p)
        raise SystemExit('*** FALHOU -- arquivos de producao devolvidos do backup ***')
    print('\nENTREGUE. Backups em:', bk)


if __name__ == '__main__':
    main()
