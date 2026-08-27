# -*- coding: utf-8 -*-
"""ler_dialogo_vba.py - le o texto do dialogo modal invisivel do VBA

O PROBLEMA QUE ISTO RESOLVE

Com Application.Visible = False, um erro de COMPILACAO do VBA abre uma caixa de
dialogo que ninguem ve e que trava a chamada COM ate alguem clicar. O codigo que
chega ao Python e so 0x800A9C68 -- "erro de compilacao" --, sem dizer QUAL.

Aqui a macro roda numa thread e, enquanto ela esta bloqueada no dialogo, a
thread principal varre as janelas do sistema (classe #32770) e le os controles
de texto. E a unica forma de obter a mensagem real sem sessao interativa.

Uso: python ler_dialogo_vba.py <arquivo.xlsm> <Macro> [segundos]
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


def textos_da_janela(hwnd):
    saida = []

    def filho(h, _):
        try:
            cls = win32gui.GetClassName(h)
            txt = win32gui.GetWindowText(h)
            if txt and cls in ('Static', 'Button', 'Edit'):
                saida.append('[%s] %s' % (cls, txt))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(hwnd, filho, None)
    except Exception:
        pass
    return saida


def varrer():
    achadas = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return True
            cls = win32gui.GetClassName(h)
            tit = win32gui.GetWindowText(h)
            if cls == '#32770' or 'Visual Basic' in tit or 'Microsoft Excel' == tit:
                achadas.append((h, cls, tit))
        except Exception:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    return achadas


def main(caminho, macro, espera=12):
    caminho = os.path.abspath(caminho)
    matar()
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl)

    def alvo():
        pythoncom.CoInitialize()
        try:
            x = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                idst, pythoncom.IID_IDispatch))
            x.Run(macro)
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()

    print('chamando %s e esperando %ds pelo dialogo...' % (macro, espera))
    threading.Thread(target=alvo, daemon=True).start()
    time.sleep(espera)

    js = varrer()
    print()
    print('janelas candidatas: %d' % len(js))
    for h, cls, tit in js:
        print()
        print('  --- hwnd=%s classe=%s titulo=%r' % (h, cls, tit))
        for t in textos_da_janela(h):
            print('      ' + t[:150])
    matar()
    print()
    print('(nada foi salvo)')


main(sys.argv[1], sys.argv[2],
     int(sys.argv[3]) if len(sys.argv) > 3 else 12)
