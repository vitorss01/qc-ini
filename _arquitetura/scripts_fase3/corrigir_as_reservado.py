# -*- coding: utf-8 -*-
"""corrigir_as_reservado.py - 'aS' e a palavra reservada 'As'

O DEFEITO

mEstatistica da Hematologia declara

    Dim aM As Double, aS As Double, etp As Double

VBA nao distingue caixa: 'aS' E a palavra reservada 'As'. O modulo inteiro para
com "Erro de compilacao: Erro de sintaxe" -- ou seja, o motor estatistico da
Hematologia nunca compilou. A Bioquimica escreve alvoM/alvoS e por isso passa.

A auditoria estatica nao pega: ela nao analisa gramatica de VBA, e o proprio
cabecalho dela avisa que nao prova compilacao.

O QUE FAZ

Renomeia o identificador aS para alvoS dentro do modulo, no arquivo, com
verificacao: dispara a compilacao, e se a caixa modal aparecer, LE o texto e
dispensa por clique programatico em vez de travar a automacao.

So salva se compilar.
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
import win32gui
import win32con
import pythoncom

RX = re.compile(r'(?<![A-Za-z0-9_])aS(?![A-Za-z0-9_])')


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


def caixa():
    achado = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return True
            if win32gui.GetClassName(h) != '#32770':
                return True
            if 'Visual Basic' not in win32gui.GetWindowText(h):
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


def dispensar(h):
    alvo = []

    def filho(hc, _):
        try:
            if win32gui.GetClassName(hc) == 'Button':
                t = win32gui.GetWindowText(hc).replace('&', '').strip().lower()
                if t in ('ok', 'fim', 'end'):
                    alvo.append(hc)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(h, filho, None)
    if alvo:
        win32gui.PostMessage(alvo[0], win32con.BM_CLICK, 0, 0)
    else:
        win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)


def compila(xl, macro, espera=25):
    """(compilou, mensagem). Dispensa a caixa para nao travar a automacao."""
    idst = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl)
    fim = {}

    def alvo():
        pythoncom.CoInitialize()
        try:
            x = w.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                idst, pythoncom.IID_IDispatch))
            try:
                x.Run(macro)
                fim['r'] = 'ok'
            except Exception as e:
                fim['r'] = str(e)[:120]
        except Exception as e:
            fim['r'] = 'thread: %s' % str(e)[:60]
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    for _ in range(espera * 2):
        if not th.is_alive():
            break
        time.sleep(0.5)
        h, txt = caixa()
        if h:
            dispensar(h)
            th.join(6)
            return (False, txt)
    th.join(3)
    return (fim.get('r') == 'ok', fim.get('r', 'sem resposta'))


def main(caminho, macro='AtualizarCalc'):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    salvar = False
    try:
        cm = wb.VBProject.VBComponents('mEstatistica').CodeModule
        n = cm.CountOfLines
        txt = cm.Lines(1, n)
        achou = len(RX.findall(txt))
        print('mEstatistica: %d linhas, %d ocorrencia(s) de aS' % (n, achou))
        if achou:
            novo_txt = RX.sub('alvoS', txt)
            cm.DeleteLines(1, n)
            cm.AddFromString(novo_txt)
            print('  renomeado aS -> alvoS (%d)' % achou)
        else:
            print('  nada a renomear')

        ok, msg = compila(xl, macro)
        print('  compilacao via %s: %s' % (macro, 'OK' if ok else msg[:110]))
        if not ok:
            raise SystemExit('nao compilou -- nada salvo')

        wb.Save()
        salvar = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            xl.ScreenUpdating = True
        except Exception:
            pass
        try:
            wb.Close(salvar)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'AtualizarCalc')
