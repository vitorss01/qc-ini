# -*- coding: utf-8 -*-
"""apontar_erro_compilacao.py - o VBE diz QUAL identificador nao esta definido

O PROBLEMA

"Erro de compilacao: 'Sub' ou 'Function' nao definida" nao diz o nome. Com o
Excel invisivel a caixa e modal e trava a automacao, entao nem o erro chega ao
Python. Procurar o nome por varredura de texto ja consumiu tempo demais e
produziu ruido: identificador de VBA pode vir de modulo, de biblioteca, de
variavel local ou de propriedade, e nenhum parser meu reproduz a resolucao de
escopo do VBA.

A SAIDA

O proprio VBE SELECIONA o identificador ofensor quando levanta esse erro.
VBE.ActiveCodePane.GetSelection() devolve linha e coluna da selecao. Entao:

  1. dispara a compilacao numa thread (uma chamada barata a qualquer macro)
  2. espera a caixa aparecer
  3. LE a selecao do VBE -- modulo, linha, e o texto exato
  4. dispensa a caixa com um clique programatico no OK, para a thread destravar

Uso: python apontar_erro_compilacao.py <arquivo.xlsm> [macro] [espera_s]
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
import win32con
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
    ult = None
    for t in range(1, 8):
        for fn in (lambda: w.DispatchEx('Excel.Application'),
                   lambda: w.Dispatch('Excel.Application')):
            try:
                xl = fn()
                xl.Visible = False
                xl.DisplayAlerts = False
                xl.EnableEvents = False
                xl.AutomationSecurity = 1
                return xl
            except Exception as e:
                ult = e
        subprocess.call(['powershell', '-NoProfile', '-Command',
                         'Start-Process excel.exe -WindowStyle Minimized'],
                        stderr=subprocess.DEVNULL)
        time.sleep(6 + 2.0 * t)
    raise RuntimeError('Excel COM nao subiu: %s' % ult)


def caixa_erro():
    """(hwnd, texto) da caixa modal do VBA, se houver."""
    achado = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return True
            if win32gui.GetClassName(h) != '#32770':
                return True
            tit = win32gui.GetWindowText(h)
            if 'Visual Basic' not in tit and 'Microsoft Excel' not in tit:
                return True
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
                achado.append((h, ' | '.join(partes)))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return achado[0] if achado else (None, None)


def dispensar(hwnd):
    """Clica OK / Fim sem sessao interativa."""
    alvo = []

    def filho(h, _):
        try:
            if win32gui.GetClassName(h) == 'Button':
                t = win32gui.GetWindowText(h).replace('&', '').strip().lower()
                if t in ('ok', 'fim', 'end'):
                    alvo.append(h)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(hwnd, filho, None)
    if alvo:
        win32gui.PostMessage(alvo[0], win32con.BM_CLICK, 0, 0)
        return True
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    return False


def main(caminho, macro='ReconciliarComCalc', espera=20):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    try:
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

        print('disparando a compilacao com %s ...' % macro)
        threading.Thread(target=alvo, daemon=True).start()

        h = txt = None
        for _ in range(espera * 2):
            time.sleep(0.5)
            h, txt = caixa_erro()
            if h:
                break
        if not h:
            print('nenhuma caixa apareceu: o projeto compilou')
            return
        print('caixa: %s' % txt)

        # o VBE seleciona o identificador ofensor
        try:
            vbe = xl.VBE
            pane = vbe.ActiveCodePane
            comp = pane.CodeModule.Parent.Name
            sel = pane.GetSelection()          # (linIni, colIni, linFim, colFim)
            l1, c1, l2, c2 = int(sel[0]), int(sel[1]), int(sel[2]), int(sel[3])
            linha = pane.CodeModule.Lines(l1, 1)
            token = linha[c1 - 1:c2 - 1] if l1 == l2 else '(varias linhas)'
            print()
            print('MODULO      : %s' % comp)
            print('LINHA %-5d : %s' % (l1, linha.strip()[:110]))
            print('SELECIONADO : %r' % token)
        except Exception as e:
            print('nao consegui ler a selecao do VBE: %s' % str(e)[:110])

        if dispensar(h):
            print()
            print('caixa dispensada')
        time.sleep(1)
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


main(sys.argv[1],
     sys.argv[2] if len(sys.argv) > 2 else 'ReconciliarComCalc',
     int(sys.argv[3]) if len(sys.argv) > 3 else 20)
