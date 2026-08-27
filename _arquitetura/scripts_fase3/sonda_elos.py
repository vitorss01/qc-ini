# -*- coding: utf-8 -*-
"""sonda_elos.py - qual elo da cadeia pos-gravacao trava

A primeira hipotese (a duplicata de AtualizarEstatistica tornando a chamada
ambigua) foi REFUTADA pela medicao: removido o stub do mUI, AtualizarOperacao
continuou sem responder por 360s. A ambiguidade e defeito, mas nao e a causa.

Aqui cada elo roda SOZINHO, com limite de tempo proprio e o Excel encerrado
entre um e outro -- porque um elo que trava deixa o processo inutilizavel para
o proximo. Cada rodada abre o arquivo do zero e fecha sem salvar.

Tambem procura MsgBox/InputBox nos modulos da cadeia: dialogo modal com o Excel
invisivel e a explicacao mais provavel de "trava sem consumir CPU".
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

ELOS = ['AtualizarBanco', 'AtualizarViewResultados',
        'AtualizarEstatistica', 'AtualizarPainel']


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
                caixa['r'] = ('erro: %s' % str(e)[:110], time.time() - t0)
        except Exception as e:
            caixa['r'] = ('thread: %s' % str(e)[:70], 0)
        finally:
            pythoncom.CoUninitialize()

    th = threading.Thread(target=alvo, daemon=True)
    th.start()
    th.join(limite)
    if th.is_alive():
        return ('TRAVOU (>%ds)' % limite, float(limite))
    return caixa.get('r', ('sem resultado', 0))


def titulo_vbe():
    try:
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             "Get-Process EXCEL -EA SilentlyContinue | "
             "Where-Object {$_.MainWindowTitle} | "
             "ForEach-Object { $_.MainWindowTitle }"],
            stderr=subprocess.DEVNULL, timeout=30)
        return out.decode('cp1252', 'replace').strip().replace('\r\n', ' / ')
    except Exception:
        return '(nao consegui ler)'


def main(caminho, limite=120):
    caminho = os.path.abspath(caminho)

    print('=== dialogos modais nos modulos da cadeia (estatico) ===')
    tmp = (r"C:\Users\vitor\AppData\Local\Temp\claude"
           r"\C--Users-vitor-OneDrive---MSFT-Desktop-QC-INI"
           r"\919a6d1a-4dbb-471e-9759-5f896c9a1c9b\scratchpad\vba_Bioquimica")
    if os.path.isdir(tmp):
        for f in sorted(os.listdir(tmp)):
            try:
                cod = io.open(os.path.join(tmp, f), encoding='cp1252',
                              errors='replace').read()
            except IOError:
                continue
            hits = [(i + 1, l.strip()[:74])
                    for i, l in enumerate(cod.split('\n'))
                    if re.search(r'\b(MsgBox|InputBox|\.Show\b)', l)
                    and not l.strip().startswith("'")]
            if hits and f.replace('.bas', '').replace('.cls', '') in (
                    'mOperacao', 'mDados', 'mUI', 'mEstatistica', 'mBanco',
                    'mLotes', 'mSeguranca'):
                print('  %s:' % f)
                for n, l in hits[:6]:
                    print('     l.%-5d %s' % (n, l))
    print()

    print('=== cada elo, sozinho, com limite de %ds ===' % limite)
    for macro in ELOS:
        matar()
        xl = novo()
        wb = xl.Workbooks.Open(caminho)
        try:
            est, seg = com_limite(xl, macro, limite)
            print('  %-26s %-22s %6.1fs' % (macro, est, seg))
            if est.startswith('TRAVOU'):
                print('        janela: %s' % titulo_vbe()[:110])
        finally:
            if not est.startswith('TRAVOU'):
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
