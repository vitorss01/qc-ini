# -*- coding: utf-8 -*-
"""testar_protecao_adr025.py - o cenario que produziu o erro 1004

O defeito so aparecia no ARTEFATO SALVO E REABERTO, nunca na producao aberta:
UserInterfaceOnly nao e persistido, entao ao reabrir a aba esta protegida por
inteiro e as duas escritas de AtualizarFlagsBanco falham com 1004.

Este teste reproduz exatamente isso -- abre o artefato SEM destravar nada e
chama a rotina. E confere as duas metades do contrato:
  1. a rotina consegue escrever
  2. a protecao VOLTA depois (nao pode ficar destravada por efeito colateral)

Uso: python testar_protecao_adr025.py <artefato.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

falhas = []

WRAP = r'''Attribute VB_Name = "mProtTemp"
Option Explicit
' Wrapper TEMPORARIO: Err.Raise cru sob Application.Run abre dialogo modal e
' trava o Excel. Devolvendo string, o teste le o erro em vez de congelar.
Public Function RodarFlags() As String
    On Error GoTo falhou
    AtualizarFlagsBanco
    RodarFlags = "OK"
    Exit Function
falhou:
    RodarFlags = "ERRO " & Err.Number & ": " & Err.Description
End Function

Public Function EstadoAba() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DB_Resultados")
    EstadoAba = "ProtectContents=" & ws.ProtectContents & _
                "|ProtectionMode=" & ws.ProtectionMode & _
                "|BA4_Locked=" & ws.Range("BA4").Locked & _
                "|BA4=" & CStr(ws.Range("BA4").Value) & _
                "|BB4=" & CStr(ws.Range("BB4").Value)
End Function
'''


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def novo_excel():
    for t in range(1, 8):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t == 2:
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(6)
            time.sleep(2.0 * t)
    raise RuntimeError('Excel COM nao subiu')


def main(caminho):
    aqui = os.path.dirname(os.path.abspath(__file__))
    copia = os.path.join(os.environ.get('TEMP', aqui), 'prot_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    bas = os.path.join(os.environ.get('TEMP', aqui), 'mProtTemp.bas')
    open(bas, 'w', encoding='cp1252').write(WRAP)

    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        db = wb.Worksheets('DB_Resultados')
        print('=== ARTEFATO RECEM-ABERTO, NADA DESTRAVADO ===')
        prot0 = db.ProtectContents
        modo0 = db.ProtectionMode
        print('   ProtectContents=%s  ProtectionMode(UserInterfaceOnly)=%s  BA4.Locked=%s'
              % (prot0, modo0, db.Range('BA4').Locked))
        ck('a aba REALMENTE esta protegida (senao o teste nao prova nada)',
           prot0 is True, 'ProtectContents=%s' % prot0)
        ck('UserInterfaceOnly NAO sobreviveu ao salvar (premissa do defeito)',
           modo0 is False, 'ProtectionMode=%s' % modo0)

        # o wrapper precisa entrar; a estrutura pode estar protegida
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect('qcini2025')
        for c in list(wb.VBProject.VBComponents):
            if c.Name == 'mProtTemp':
                wb.VBProject.VBComponents.Remove(c)
        wb.VBProject.VBComponents.Import(bas)
        if estrutura:
            wb.Protect('qcini2025', True, False)

        print('\n=== CHAMAR AtualizarFlagsBanco COM A ABA PROTEGIDA ===')
        r = str(xl.Run('RodarFlags'))
        print('   retorno: %s' % r)
        ck('AtualizarFlagsBanco executa em aba protegida', r == 'OK', r)

        print('\n=== A PROTECAO VOLTOU? ===')
        est = str(xl.Run('EstadoAba'))
        for p in est.split('|'):
            print('   %s' % p)
        ck('a aba continua protegida depois da rotina',
           'ProtectContents=Verdadeiro' in est or 'ProtectContents=True' in est, est[:60])
        ck('BA4 tem nucleo de lote gravado',
           'BA4=|' not in est + '|' and 'BA4=' in est and est.split('BA4=')[1].split('|')[0].strip() != '',
           est.split('BA4=')[1].split('|')[0] if 'BA4=' in est else '?')

        print('\n=== SEGUNDA CHAMADA (idempotencia) ===')
        r2 = str(xl.Run('RodarFlags'))
        print('   retorno: %s' % r2)
        ck('roda de novo sem erro', r2 == 'OK', r2)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print('\n' + '=' * 60)
    if falhas:
        print('FALHAS (%d): %s' % (len(falhas), '; '.join(falhas)))
        sys.exit(1)
    print('TESTE DE PROTECAO: PASSOU')


if __name__ == '__main__':
    main(sys.argv[1])
