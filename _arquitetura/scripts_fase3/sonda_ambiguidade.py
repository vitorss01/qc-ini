# -*- coding: utf-8 -*-
"""sonda_ambiguidade.py - a duplicata de AtualizarEstatistica trava a cadeia?

HIPOTESE

mOperacao.AtualizarOperacao -- a cadeia chamada apos qualquer gravacao ou
exclusao -- chama AtualizarEstatistica SEM QUALIFICAR. Na Bioquimica de
producao ha DUAS Public Sub com esse nome (mUI, um stub de 4 linhas; e
mEstatistica, o motor). Nome publico duplicado entre modulos torna a chamada
AMBIGUA, e o VBA responde com ERRO DE COMPILACAO -- que On Error nao captura,
porque nao e erro de execucao.

Isso explica por que xl.Run("AtualizarOperacao") trava por dez minutos sem
devolver nada, enquanto EstatPeriodo e AlvoDoLote rodam normalmente: o VBA
compila por modulo, sob demanda. So quem toca o mOperacao esbarra na
ambiguidade.

O TESTE

Com timeout, para nao repetir a travada:

  1. chama AtualizarOperacao com o projeto como esta        -> espera-se travar
  2. REMOVE o stub do mUI em memoria
  3. chama de novo                                           -> espera-se rodar

Se o passo 3 completa e o 1 nao, a ambiguidade e a causa.

Nada e salvo: o arquivo e fechado com Close(False).
"""
import io
import os
import re
import sys
import time
import threading
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w
import pythoncom


def novo():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (1, 3, 5):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Minimized'],
                                stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def rodar_com_limite(xl_id, macro, limite):
    """Chama a macro numa thread propria. Devolve (estado, segundos).

    A thread nao pode ser morta: se o VBA abriu um dialogo modal, ela fica la.
    Mas o resultado ('travou') ja e a informacao que interessa, e o processo
    do Excel e encerrado no fim de qualquer jeito.
    """
    caixa = {}

    def alvo():
        pythoncom.CoInitialize()
        try:
            xl2 = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                xl_id, pythoncom.IID_IDispatch))
            t0 = time.time()
            try:
                xl2.Run(macro)
                caixa['r'] = ('completou', time.time() - t0)
            except Exception as e:
                caixa['r'] = ('erro: %s' % str(e)[:120], time.time() - t0)
        except Exception as e:
            caixa['r'] = ('falha na thread: %s' % str(e)[:80], 0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    t0 = time.time()
    th.start()
    th.join(limite)
    if th.is_alive():
        return ('TRAVOU (sem resposta em %ds)' % limite, time.time() - t0)
    return caixa.get('r', ('sem resultado', 0))


def main(caminho, limite=90):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    try:
        onde = []
        for c in wb.VBProject.VBComponents:
            try:
                cm = c.CodeModule
                txt = cm.Lines(1, cm.CountOfLines)
            except Exception:
                continue
            if re.search(r'^\s*(Public\s+)?Sub\s+AtualizarEstatistica\s*\(',
                         txt, re.M | re.I):
                onde.append(c.Name)
        print('AtualizarEstatistica definida em: %s' % onde)
        if len(onde) < 2:
            print('sem duplicata neste produto: nada a medir')
            return
        print()

        idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, xl)
        est, seg = rodar_com_limite(idst, 'AtualizarOperacao', limite)
        print('1) com a duplicata      -> %s  (%.1fs)' % (est, seg))
        travou1 = est.startswith('TRAVOU')
        if not travou1:
            print()
            print('nao travou: a hipotese da ambiguidade NAO se confirma aqui.')
            return
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force'],
                        stderr=subprocess.DEVNULL)
        time.sleep(4)

    # ---- segunda rodada, com o stub removido -------------------------------
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    try:
        cm = wb.VBProject.VBComponents('mUI').CodeModule
        ini = cm.ProcStartLine('AtualizarEstatistica', 0)
        n = cm.ProcCountLines('AtualizarEstatistica', 0)
        print('   removendo mUI.AtualizarEstatistica: linhas %d..%d' % (ini, ini + n - 1))
        cm.DeleteLines(ini, n)

        idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, xl)
        est2, seg2 = rodar_com_limite(idst, 'AtualizarOperacao', limite * 4)
        print('2) sem o stub do mUI    -> %s  (%.1fs)' % (est2, seg2))
        print()
        if est2.startswith('completou'):
            print('VEREDITO: a duplicata era a causa.')
            print('  Com as duas copias, a chamada sem qualificar do mOperacao e')
            print('  ambigua e o VBA para com erro de COMPILACAO -- que nenhum')
            print('  On Error captura. Removido o stub, a cadeia roda.')
        else:
            print('VEREDITO: removendo o stub a cadeia ainda nao completa (%s).' % est2)
            print('  A ambiguidade e defeito de qualquer forma, mas nao e a unica causa.')
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 90)
