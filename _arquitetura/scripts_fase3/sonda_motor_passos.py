# -*- coding: utf-8 -*-
"""sonda_motor_passos.py - removido o stub, qual passo do motor trava

mEstatistica.AtualizarEstatistica encadeia seis passos:

    InvalidarCache
    AtualizarCalc
    RegistrarEventosWestgard
    AtualizarPainelEng
    AtualizarEstatisticaAba
    AtualizarEixos

Com o stub do mUI removido, AtualizarOperacao continuou sem responder por 360s.
Aqui cada passo roda sozinho, com limite proprio e Excel novo entre um e outro.

Suspeito de saida: RegistrarEventosWestgard escreve na aba Eventos_Westgard, que
EXISTE na Hematologia e NAO existe na Bioquimica.

Nada e salvo.
"""
import io
import os
import sys
import time
import threading
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w
import pythoncom

PASSOS = ['InvalidarCache', 'AtualizarCalc', 'RegistrarEventosWestgard',
          'AtualizarPainelEng', 'AtualizarEstatisticaAba', 'AtualizarEixos']


def matar():
    subprocess.call(['powershell', '-NoProfile', '-Command',
                     'Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force'],
                    stderr=subprocess.DEVNULL)
    time.sleep(4)


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


def com_limite(xl, macro, limite):
    caixa = {}
    idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl)

    def alvo():
        pythoncom.CoInitialize()
        try:
            x = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                idst, pythoncom.IID_IDispatch))
            t0 = time.time()
            try:
                x.Run(macro)
                caixa['r'] = ('completou', time.time() - t0)
            except Exception as e:
                s = str(e)
                if 'não esteja disponível' in s or 'nao esteja disponivel' in s:
                    s = 'macro AMBIGUA ou ausente'
                caixa['r'] = ('erro: %s' % s[:90], time.time() - t0)
        except Exception as e:
            caixa['r'] = ('thread: %s' % str(e)[:60], 0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    th.join(limite)
    if th.is_alive():
        return ('TRAVOU (>%ds)' % limite, float(limite))
    return caixa.get('r', ('sem resultado', 0))


def main(caminho, limite=120):
    caminho = os.path.abspath(caminho)
    print('=== passos do motor, cada um sozinho, limite %ds ===' % limite)
    for macro in PASSOS:
        matar()
        xl = novo()
        wb = xl.Workbooks.Open(caminho)
        travou = False
        try:
            est, seg = com_limite(xl, macro, limite)
            travou = est.startswith('TRAVOU')
            print('  %-26s %-46s %6.1fs' % (macro, est, seg))
        finally:
            if not travou:
                try:
                    wb.Close(False)
                except Exception:
                    pass
                try:
                    xl.Quit()
                except Exception:
                    pass

    # e agora AtualizarEstatistica com o stub fora do caminho
    matar()
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    travou = False
    try:
        cm = wb.VBProject.VBComponents('mUI').CodeModule
        ini = cm.ProcStartLine('AtualizarEstatistica', 0)
        n = cm.ProcCountLines('AtualizarEstatistica', 0)
        cm.DeleteLines(ini, n)
        print()
        print('  (stub do mUI removido em memoria: linhas %d..%d)' % (ini, ini + n - 1))
        est, seg = com_limite(xl, 'AtualizarEstatistica', limite * 3)
        travou = est.startswith('TRAVOU')
        print('  %-26s %-46s %6.1fs' % ('AtualizarEstatistica', est, seg))
    finally:
        if not travou:
            try:
                wb.Close(False)
            except Exception:
                pass
            try:
                xl.Quit()
            except Exception:
                pass
    matar()
    print()
    print('(nenhuma rodada salvou o arquivo)')


main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 120)
