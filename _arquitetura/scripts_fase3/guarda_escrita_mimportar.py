# -*- coding: utf-8 -*-
"""guarda_escrita_mimportar.py - ADR-046 na linhagem de PRODUCAO do mImportar

O DEFEITO (achado G da auditoria estatica)

mImportar.MostrarErros e mImportar.LimparAreaImport desprotegem a aba, escrevem
e reprotegem -- sem On Error. Um erro entre o Unprotect e o Protect deixa a aba
DESTRANCADA e o usuario nunca fica sabendo. Troca um defeito visivel por um furo
silencioso.

POR QUE NAO IMPORTAR LiberarEscrita/RestaurarProtecao

O trecho pronto em src_producao/mSeguranca_GUARDA.txt usa a constante
SENHA_PROT, que NAO EXISTE no mSeguranca de producao (299 linhas, senha em
literal). Importa-lo como esta nao compilaria -- e um modulo que nao compila
derruba o projeto inteiro, nao so a importacao.

A linhagem de producao ja tem o idioma dela: captura prot = ws.ProtectContents,
desprotege com IMP_SENHA e reprotege com as mesmas flags. Falta so o caminho de
erro. E o que se acrescenta aqui, preservando o idioma local.

Escreve em src_producao/mImportar.bas (cp1252, sem acento no codigo) e
reimporta no QC_Bioquimica.xlsm -- o unico produto que tem este modulo.
"""
import io
import os
import re
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
FONTE = os.path.join(BASE, '_arquitetura', 'src_producao', 'mImportar.bas')
ALVO = os.path.join(BASE, 'QC_Bioquimica.xlsm')

VELHO_MOSTRAR = """    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:=IMP_SENHA
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    ws.Cells(IMP_CAB, cErr).Value = "Inconsistencias"
    ws.Cells(IMP_CAB, cErr).Font.Bold = True
    If erros.Count = 0 Then
        ws.Cells(IMP_R0, cErr).Value = "(nenhuma)"
    Else
        For i = 1 To erros.Count
            ws.Cells(IMP_R0 + i - 1, cErr).Value = erros(i)
        Next i
    End If
    ws.Columns(cErr).ColumnWidth = 58
    If prot Then ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub"""

NOVO_MOSTRAR = """    ' ADR-046: a janela destrancada tem de fechar tambem no erro.
    ' Sem o On Error, uma excecao entre o Unprotect e o Protect deixa a aba
    ' aberta para edicao e ninguem fica sabendo. O erro original e relancado
    ' depois de reproteger, para nao trocar um defeito por um silencio.
    Dim nErrP As Long, sErrP As String
    prot = ws.ProtectContents
    On Error GoTo restaura
    If prot Then ws.Unprotect Password:=IMP_SENHA
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    ws.Cells(IMP_CAB, cErr).Value = "Inconsistencias"
    ws.Cells(IMP_CAB, cErr).Font.Bold = True
    If erros.Count = 0 Then
        ws.Cells(IMP_R0, cErr).Value = "(nenhuma)"
    Else
        For i = 1 To erros.Count
            ws.Cells(IMP_R0 + i - 1, cErr).Value = erros(i)
        Next i
    End If
    ws.Columns(cErr).ColumnWidth = 58

restaura:
    nErrP = Err.Number: sErrP = Err.Description
    If prot Then
        On Error Resume Next
        ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mImportar.MostrarErros", sErrP
End Sub"""

VELHO_LIMPAR = """    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:=IMP_SENHA
    ws.Range(ws.Cells(IMP_R0, IMP_C_DATA), ws.Cells(IMP_RN, IMP_C_AN0 + nC - 1)).ClearContents
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents
    If prot Then ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
                            DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub"""

NOVO_LIMPAR = """    ' ADR-046: mesma guarda do MostrarErros. Aqui pesa mais, porque a rotina
    ' APAGA a area de colagem: um erro no meio deixaria a aba destrancada logo
    ' depois de uma limpeza.
    Dim nErrP As Long, sErrP As String
    prot = ws.ProtectContents
    On Error GoTo restaura
    If prot Then ws.Unprotect Password:=IMP_SENHA
    ws.Range(ws.Cells(IMP_R0, IMP_C_DATA), ws.Cells(IMP_RN, IMP_C_AN0 + nC - 1)).ClearContents
    ws.Range(ws.Cells(IMP_CAB, cErr), ws.Cells(IMP_RN, cErr)).ClearContents

restaura:
    nErrP = Err.Number: sErrP = Err.Description
    If prot Then
        On Error Resume Next
        ws.Protect Password:=IMP_SENHA, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mImportar.LimparAreaImport", sErrP
End Sub"""


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


def main():
    cod = io.open(FONTE, encoding='cp1252').read()
    for velho, novo_txt, nome in ((VELHO_MOSTRAR, NOVO_MOSTRAR, 'MostrarErros'),
                                  (VELHO_LIMPAR, NOVO_LIMPAR, 'LimparAreaImport')):
        if novo_txt.split('\n')[0].strip() in cod:
            print('  %s: guarda ja aplicada' % nome)
            continue
        if velho not in cod:
            raise SystemExit('nao encontrei o corpo de %s como esperado' % nome)
        cod = cod.replace(velho, novo_txt)
        print('  %s: guarda de erro acrescentada' % nome)

    # o .bas e lido pelo VBA como cp1252; acento aqui vira corrupcao silenciosa
    try:
        cod.encode('cp1252')
    except UnicodeEncodeError as e:
        raise SystemExit('caractere fora do cp1252 no modulo: %s' % e)
    io.open(FONTE, 'w', encoding='cp1252', newline='\r\n').write(cod)
    print('  fonte gravada: %s' % os.path.relpath(FONTE, BASE))

    xl = novo()
    wb = xl.Workbooks.Open(ALVO)
    salvar = False
    try:
        proj = wb.VBProject
        for c in list(proj.VBComponents):
            if c.Name == 'mImportar':
                proj.VBComponents.Remove(c)
                print('  mImportar anterior removido do arquivo')
                break
        proj.VBComponents.Import(FONTE)
        alvo = [c for c in proj.VBComponents if c.Name == 'mImportar']
        if not alvo:
            raise SystemExit('Import nao criou mImportar')
        cm = alvo[0].CodeModule
        print('  mImportar reimportado (%d linhas)' % cm.CountOfLines)
        txt = cm.Lines(1, cm.CountOfLines).lower()
        for nome in ('mostrarerros', 'limpareaimport', 'limpaareaimport',
                     'limparareaimport'):
            pass
        if txt.count('on error goto restaura') < 2:
            raise SystemExit('as duas guardas nao chegaram ao modulo')
        print('  duas guardas presentes no modulo dentro do arquivo')

        # o projeto tem de continuar rodando
        xl.Run('IrParaImportar')
        print('  IrParaImportar executou: o projeto compila com o modulo novo')

        wb.Save()
        salvar = True
        print('SALVO: %s' % ALVO)
    finally:
        try:
            wb.Close(salvar)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


main()
