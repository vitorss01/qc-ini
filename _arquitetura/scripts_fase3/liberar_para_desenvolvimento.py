# -*- coding: utf-8 -*-
"""liberar_para_desenvolvimento.py - tira senha, protecao e travas das celulas

FASE DE DESENVOLVIMENTO: TUDO EDITAVEL

Enquanto o produto esta sendo construido, protecao atrapalha mais do que
protege -- e ja atrapalhou: varios scripts desta sessao gastaram tentativas
com "aba protegida", "AllowFormattingCells=False" e RPC_E_CALL_REJECTED em
cima de celula travada.

O QUE FAZ

  1. desprotege a ESTRUTURA da pasta (senha qcini2025)
  2. desprotege TODAS as abas, inclusive as ocultas e veryHidden
  3. destrava (Locked = False) TODAS as celulas de todas as abas
  4. deixa toda aba visivel? NAO -- ver abaixo

ABA OCULTA NAO E ABA TRANCADA

Continuam ocultas as abas que sao camada de dados e nao tela: LotesStore,
LiberStore, RegistrosStore, Calc, Login, BI_Data. Torna-las visiveis nao
libera nada -- so polui a barra de abas com estrutura interna. Se voce quiser
ver alguma, o Modo Desenvolvedor ja faz isso, e o "veryHidden" vira "hidden"
aqui para que apareçam na lista de reexibir do Excel.

ISTO E REVERSIVEL

travar_estrutura.ps1 e blindar_artefato.ps1 continuam no build e reaplicam a
protecao no ARTEFATO entregue. Este script mexe no arquivo de trabalho.

Uso: python liberar_para_desenvolvimento.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (2, 4):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(8)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def tenta(fn, vezes=8):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            s = str(e).lower()
            if 'rejeitada' not in s and 'rejected' not in s:
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura (aberto no Excel?): %s' % caminho)
    try:
        xl.Calculation = -4135

        # ---- 1. estrutura ---------------------------------------------------
        if wb.ProtectStructure:
            for p in (SENHA, None):
                try:
                    wb.Unprotect(p) if p else wb.Unprotect()
                    break
                except Exception:
                    pass
        print('estrutura da pasta: %s' % ('AINDA PROTEGIDA' if wb.ProtectStructure else 'liberada'))

        # ---- 2 e 3. abas e celulas -----------------------------------------
        nAbas = 0
        nProt = 0
        nVery = 0
        for ws in wb.Worksheets:
            nAbas += 1
            vis = ws.Visible
            if vis == 2:                      # xlSheetVeryHidden
                ws.Visible = 0                # xlSheetHidden: aparece em "Reexibir"
                nVery += 1
            if ws.ProtectContents:
                nProt += 1
                for p in (SENHA, None):
                    try:
                        ws.Unprotect(p) if p else ws.Unprotect()
                        break
                    except Exception:
                        pass
            # destrava TODAS as celulas -- e o Locked que impede editar quando
            # alguem reproteger a aba mais tarde
            try:
                tenta(lambda s=ws: s.Cells.__setattr__('Locked', False))
            except Exception as e:
                print('   %s: nao destravou (%s)' % (ws.Name, str(e)[:50]))

        print('abas: %d ; estavam protegidas: %d ; veryHidden -> hidden: %d'
              % (nAbas, nProt, nVery))

        # ---- 4. conferencia --------------------------------------------------
        resta = []
        travadas = []
        for ws in wb.Worksheets:
            if ws.ProtectContents:
                resta.append(ws.Name)
            try:
                if ws.Range('A1').Locked:
                    travadas.append(ws.Name)
            except Exception:
                pass
        print('abas ainda protegidas: %s' % (resta if resta else 'nenhuma'))
        print('abas com A1 travada  : %s' % (travadas if travadas else 'nenhuma'))
        if resta or travadas or wb.ProtectStructure:
            raise SystemExit('sobrou protecao -- nada salvo')

        xl.Calculation = -4105
        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    main(sys.argv[1])
