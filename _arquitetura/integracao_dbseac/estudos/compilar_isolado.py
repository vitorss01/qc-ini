# -*- coding: utf-8 -*-
"""Compila modulos VBA um a um numa COPIA, para isolar qual nao compila (sem salvar).
Uso: python compilar_isolado.py <copia.xlsm> <modulo.bas> [...]
Para cada modulo importado: Depurar > Compilar e espera ate 30 s o comando ficar desabilitado (= compilou).
Um erro de compilacao abre um modal no VBE: o cao de guarda de 20 s mata o Excel desta copia (so o PID criado aqui).
"""
import os
import sys
import threading
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..', 'testes'))
sys.path.insert(0, AQUI)
from qa_final import QA  # noqa: E402
from equivalencia_cq import importar  # noqa: E402


def compilar(q, teto=20, espera=30):
    ctl = q.wb.VBProject.VBE.CommandBars.FindControl(1, 578)
    feito, morto = threading.Event(), []

    def cao():
        if not feito.wait(teto):
            morto.append(True)
            q.ex.matar()
    threading.Thread(target=cao, daemon=True).start()
    try:
        if ctl.Enabled:
            ctl.Execute()
    except Exception as e:
        return 'MODAL/ERRO (' + str(e)[:80] + ')' if morto else 'excecao ' + str(e)[:80]
    finally:
        feito.set()
    if morto:
        return 'MODAL DE ERRO DE COMPILACAO'
    t0 = time.time()
    while time.time() - t0 < espera:
        if not ctl.Enabled:
            return f'OK ({time.time() - t0:.1f}s)'
        time.sleep(0.5)
    return 'comando ainda habilitado apos %ds' % espera


def main(arq, mods):
    q = QA('Bioquimica', arq)
    try:
        print('base:', compilar(q), flush=True)
        for m in mods:
            importar(q.wb, m)
            r = compilar(q)
            print(os.path.basename(m), '->', r, flush=True)
            if not r.startswith('OK'):
                break
    finally:
        try:
            q.fechar(salvar=False)
        except Exception:
            pass


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), [os.path.abspath(x) for x in sys.argv[2:]])
