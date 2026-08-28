# -*- coding: utf-8 -*-
"""estrutura_pasta.py - destrava/retrava a estrutura da pasta de trabalho

Sheets.Add falha com "nao e possivel obter a propriedade Add da classe Sheets"
quando a ESTRUTURA da pasta esta protegida -- e a mensagem nao diz isso. E a
mesma armadilha do ADR-034, quando o montar_modulo_eqa nao destravava a
estrutura e o erro nao explicava por que.

Uso: python estrutura_pasta.py <arquivo.xlsm> destravar|travar
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'


def novo():
    try:
        n = int(subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process EXCEL -EA SilentlyContinue | Measure-Object).Count'],
            stderr=subprocess.DEVNULL, timeout=40).decode('ascii', 'ignore').strip() or '0')
    except Exception:
        n = 0
    if not n:
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


def main(caminho, acao):
    caminho = os.path.abspath(caminho)
    xl = novo()
    wb = xl.Workbooks.Open(caminho)
    try:
        antes = bool(wb.ProtectStructure)
        print('estrutura protegida antes: %s' % antes)
        if acao == 'destravar':
            if antes:
                for s in (SENHA, None):
                    try:
                        wb.Unprotect(s) if s else wb.Unprotect()
                        break
                    except Exception:
                        pass
        else:
            if not antes:
                wb.Protect(SENHA, True, False)
        depois = bool(wb.ProtectStructure)
        print('estrutura protegida depois: %s' % depois)
        if depois != antes:
            wb.Save()
            print('SALVO')
        else:
            print('sem mudanca -- nada salvo')
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main(sys.argv[1], sys.argv[2])
