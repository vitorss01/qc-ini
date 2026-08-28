# -*- coding: utf-8 -*-
"""padronizar_cab_eventos.py - o cabecalho de Eventos_Westgard nos dois produtos

DOIS PROBLEMAS, UM ARQUIVO

1. HEMATOLOGIA COM CABECALHO VELHO
   RegistrarEventosWestgard so reescreve os DADOS, da linha 4 para baixo. Depois
   da migracao para o motor de evento, a Hematologia passou a gravar 14 colunas
   sob um cabecalho de 8 -- dado novo com rotulo antigo, que e pior do que
   qualquer um dos dois sozinho.

2. ACENTO CORROMPIDO NA BIOQUIMICA
   O cabecalho saiu como "NÍveis" e "Evidéncia". Nao e so feio: o Power Query
   casa coluna por NOME, e "Níveis" nao encontra "NÍveis". A carga cairia no
   ramo errado sem dizer por que.

A cura do segundo caso nao e so corrigir o texto -- e parar de depender dele.
EV_NCOL = 14 e um contrato do motor: as posicoes sao fixas. O modelo semantico
passou a nomear por POSICAO, e este script deixa o rotulo visivel correto para
quem abre a planilha.
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"

# a ordem E o contrato: EV_NCOL = 14, posicoes fixas (ADR-045)
CAB = ['Data', 'RUN', 'Analito', 'Níveis', 'Regra', 'Detector', 'Escopo',
       'Classe', 'N', 'R', 'RUN_Inicial', 'Evidência', 'Classificação', 'Z_Max']


def proc_excel():
    try:
        return int(subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process EXCEL -EA SilentlyContinue | Measure-Object).Count'],
            stderr=subprocess.DEVNULL, timeout=40).decode('ascii', 'ignore').strip() or '0') > 0
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


def main():
    for arq in ('QC_Bioquimica.xlsm', 'QC_Hematologia.xlsm'):
        caminho = os.path.join(BASE, arq)
        xl = novo()
        wb = xl.Workbooks.Open(caminho)
        salvar = False
        try:
            ws = wb.Worksheets('Eventos_Westgard')
            antes = [str(ws.Cells(3, c).Value or '') for c in range(1, 15)]
            prot = bool(ws.ProtectContents)
            if prot:
                for s in ('qcini2025', None):
                    try:
                        ws.Unprotect(s) if s else ws.Unprotect()
                        break
                    except Exception:
                        pass
            for i, nome in enumerate(CAB, start=1):
                ws.Cells(3, i).Value = nome
            depois = [str(ws.Cells(3, c).Value or '') for c in range(1, 15)]
            if prot:
                try:
                    ws.Protect('qcini2025', True, True, False, True)
                except Exception:
                    pass
            ult = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
            larg = 0
            for c in range(1, 15):
                if ws.Cells(4, c).Value not in (None, ''):
                    larg = c
            print('%s' % arq)
            print('   antes : %s' % ' | '.join(x for x in antes if x))
            print('   depois: %s' % ' | '.join(depois))
            print('   eventos: %d linha(s) | largura do dado: %d coluna(s)'
                  % (max(0, ult - 3), larg))
            if antes != depois:
                wb.Save()
                salvar = True
                print('   SALVO')
            else:
                print('   ja estava correto')
        finally:
            try:
                wb.Close(salvar)
            except Exception:
                pass
            try:
                xl.Quit()
            except Exception:
                pass
        time.sleep(2)


main()
