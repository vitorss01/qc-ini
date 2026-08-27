# -*- coding: utf-8 -*-
"""validar_cadeia_bioquimica.py - a cadeia pos-gravacao roda de ponta a ponta?

Prova no arquivo que sera ENTREGUE, nao numa copia de build.

  AtualizarEstatistica   o motor inteiro
  AtualizarOperacao      a cadeia chamada apos qualquer gravacao/exclusao
                         (AtualizarBanco -> View -> Estatistica -> Painel)

Confere tambem que ScreenUpdating volta a True, porque AtualizarEstatistica o
desliga no inicio e nao tem bloco de finalizacao garantida: qualquer erro no
meio deixaria o Excel parecendo congelado para o usuario.

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
import win32gui
import pythoncom


def proc_excel():
    try:
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process EXCEL -EA SilentlyContinue | Measure-Object).Count'],
            stderr=subprocess.DEVNULL, timeout=40)
        return int(out.decode('ascii', 'ignore').strip() or '0') > 0
    except Exception:
        return False


def novo():
    if not proc_excel():
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Start-Process excel.exe -WindowStyle Minimized'],
                        stderr=subprocess.DEVNULL)
        time.sleep(14)
    for t in range(1, 10):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            subprocess.call(['powershell', '-NoProfile', '-Command',
                             'Start-Process excel.exe -WindowStyle Minimized'],
                            stderr=subprocess.DEVNULL)
            time.sleep(6 + 2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def dialogo():
    achado = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == '#32770':
                partes = []

                def filho(hc, _2):
                    try:
                        if win32gui.GetClassName(hc) == 'Static':
                            t = win32gui.GetWindowText(hc)
                            if t:
                                partes.append(t.replace('\r', ' ').replace('\n', ' '))
                    except Exception:
                        pass
                    return True
                win32gui.EnumChildWindows(h, filho, None)
                if partes:
                    achado.append(' | '.join(partes))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return achado


def rodar(xl, macro, limite):
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
                caixa['r'] = ('erro: %s' % str(e)[:140], time.time() - t0)
        except Exception as e:
            caixa['r'] = ('thread: %s' % str(e)[:70], 0.0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    th.join(limite)
    if th.is_alive():
        return ('TRAVOU (>%ds) dialogo=%s' % (limite, dialogo() or 'nenhum'),
                float(limite))
    return caixa.get('r', ('sem resultado', 0.0))


def main(caminho, limite=300):
    caminho = os.path.abspath(caminho)
    ok = fail = 0
    for macro in ('AtualizarEstatistica', 'AtualizarOperacao'):
        xl = novo()
        wb = xl.Workbooks.Open(caminho)
        travou = False
        try:
            est, seg = rodar(xl, macro, limite)
            travou = est.startswith('TRAVOU')
            try:
                su = bool(xl.ScreenUpdating)
            except Exception:
                su = None
            bom = (est == 'completou') and (su is True)
            ok, fail = (ok + 1, fail) if bom else (ok, fail + 1)
            print('  %-5s %-22s %-46s %6.1fs  ScreenUpdating=%s'
                  % ('OK' if bom else 'FALHA', macro, est[:46], seg, su))
        finally:
            if not travou:
                try:
                    xl.ScreenUpdating = True
                except Exception:
                    pass
                try:
                    wb.Close(False)
                except Exception:
                    pass
                try:
                    xl.Quit()
                except Exception:
                    pass
            time.sleep(3)

    print()
    print('%d OK, %d FALHA' % (ok, fail))
    sys.exit(1 if fail else 0)


main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 300)
