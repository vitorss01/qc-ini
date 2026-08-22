# -*- coding: utf-8 -*-
"""diagnosticar_bi.py - por que AtualizarBIData parou de responder

CONTEXTO

Depois do ADR-033 a rotina passou a nao voltar mais: Excel ocioso (CPU parada),
processo respondendo, nenhuma mensagem. Esse quadro nao e "lento" -- e caixa de
dialogo modal invisivel, porque a instancia sobe com Visible = False.

Este script NAO corrige nada. Ele instala um invólucro temporario que captura
Err.Number e Err.Description e devolve como texto, roda a rotina, imprime o que
achou e fecha SEM salvar. O modulo de diagnostico nunca chega ao arquivo.

On Error GoTo aqui e INSTRUMENTACAO, nao remendo: existe para o erro aparecer,
nao para ele sumir.

Uso: python diagnosticar_bi.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SHIM = '''Attribute VB_Name = "mDiagBI"
Option Explicit

Public Function DiagAtualizarBI() As String
    On Error GoTo falhou
    mBI.AtualizarBIData
    DiagAtualizarBI = "OK"
    Exit Function
falhou:
    DiagAtualizarBI = "ERRO " & Err.Number & " | " & Err.Description & _
                      " | origem: " & Err.Source
End Function

' Repete so a escrita do bloco, com um array que contem Null, para saber se e
' isso que derruba a rotina.
Public Function DiagEscreveNull() As String
    Dim ws As Worksheet, a(1 To 1, 1 To 3) As Variant
    On Error GoTo falhou
    Set ws = ThisWorkbook.Worksheets("BI_Data")
    a(1, 1) = "sonda"
    a(1, 2) = Null
    a(1, 3) = 1
    ws.Range(ws.Cells(100000, 1), ws.Cells(100000, 3)).Value = a
    ws.Range(ws.Cells(100000, 1), ws.Cells(100000, 3)).ClearContents
    DiagEscreveNull = "escrita com Null: OK"
    Exit Function
falhou:
    DiagEscreveNull = "escrita com Null: ERRO " & Err.Number & " | " & Err.Description
End Function

' Quantas linhas do fato receberiam Null nas colunas novas?
Public Function DiagQuantosNull() As String
    Dim v As Variant, n As Long
    On Error GoTo falhou
    v = mQualidade.MargemETp(Empty, 5#)
    DiagQuantosNull = "MargemETp(Empty,5) devolve: " & TypeName(v) & _
                      " / IsNull=" & CStr(IsNull(v))
    Exit Function
falhou:
    DiagQuantosNull = "ERRO " & Err.Number & " | " & Err.Description
End Function
'''


def novo_excel(visivel):
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = visivel
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


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'),
                         'diag_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    shim = os.path.join(os.environ.get('TEMP', '.'), 'mDiagBI.bas')
    io.open(shim, 'w', encoding='latin-1', newline='\r\n').write(SHIM)

    xl = novo_excel(visivel=True)     # visivel: se houver dialogo, ele aparece
    wb = xl.Workbooks.Open(copia)
    try:
        vbp = wb.VBProject
        # O shim chama mQualidade e mBI. Se qualquer um faltar, o projeto nao
        # COMPILA e a VBA abre caixa modal -- que numa instancia automatizada
        # trava tudo em silencio. Foi o que aconteceu na primeira tentativa:
        # o instrumento acusou "AtualizarBIData travou" quando o que faltava
        # era o modulo que o proprio instrumento importava.
        raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for nome in ('mQualidade', 'mBI'):
            fonte = os.path.join(raiz, 'src_producao', nome + '.bas')
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            vbp.VBComponents.Import(fonte)
            print('importado: %s' % nome)
        for c in list(vbp.VBComponents):
            if c.Name == 'mDiagBI':
                vbp.VBComponents.Remove(c)
        vbp.VBComponents.Import(shim)
        print('shim de diagnostico importado')

        # prova de compilacao ANTES de medir: se isto nao voltar, o problema
        # e de compilacao, e nao da rotina que vamos cronometrar.
        print('compila? mQualidade responde: %r'
              % xl.Run('ClassificarSigma', 4.5))

        print('\n--- 1. o que MargemETp devolve quando falta insumo ---')
        print('   %s' % xl.Run('DiagQuantosNull'))

        print('\n--- 2. escrever um array com Null numa Range ---')
        print('   %s' % xl.Run('DiagEscreveNull'))

        print('\n--- 3. AtualizarBIData sob captura de erro ---')
        t0 = time.time()
        r = xl.Run('DiagAtualizarBI')
        print('   %s   (%.1fs)' % (r, time.time() - t0))

        bi = wb.Worksheets('BI_Data')
        ult = bi.Cells(bi.Rows.Count, 1).End(-4162).Row
        print('   BI_Data: %d linhas, cabecalho 61..65 = %s'
              % (max(0, ult - 1),
                 [str(bi.Cells(1, c).Value) for c in range(61, 66)]))
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    main(sys.argv[1])
