# -*- coding: utf-8 -*-
"""testar_watch_status.py - as flags acompanham o Status por todos os caminhos

Tres caminhos, e o que mudou foi o primeiro:
  1. editar Status DIRETO na celula   -> antes deixava BA:BC obsoletas
  2. ExcluirLogico                    -> ja recalculava
  3. restaurar Status direto na celula-> antes NAO recalculava (causa da 4.2)

Uso: python testar_watch_status.py <artefato.xlsm>
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
WRAP = r'''Attribute VB_Name = "mWatchTemp"
Option Explicit
Public Function T_Excluir(ByVal r As Long, ByVal nv As Long, ByVal an As String) As Long
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(UCase$(Trim$(an))) = 1
    On Error GoTo falhou
    T_Excluir = ExcluirLogico(r, nv, d, _
        "Exclusao do teste do vigia de Status", "Excluido", "Teste", "Resultados")
    Exit Function
falhou:
    T_Excluir = -1
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
            xl.AutomationSecurity = 1
            xl.EnableEvents = True          # o ponto do teste
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
    copia = os.path.join(os.environ.get('TEMP', '.'), 'watch_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    bas = os.path.join(os.environ.get('TEMP', '.'), 'mWatchTemp.bas')
    open(bas, 'w', encoding='cp1252').write(WRAP)

    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        db = wb.Worksheets('DB_Resultados')
        lib = wb.Worksheets('Liberação')

        def flags(lin):
            return (db.Cells(lin, 54).Value, db.Cells(lin, 55).Value)

        def libA():
            return [lib.Cells(r, 1).Value for r in range(4, 30)]

        # linha 4 = primeira do RUN 1: e a que carrega BB=1 e BC=1
        lin = 4
        run = int(db.Cells(lin, 1).Value)
        an = str(db.Cells(lin, 5).Value)
        niv = int(db.Cells(lin, 3).Value)
        st0 = str(db.Cells(lin, 7).Value)
        lib0 = libA()
        print('linha %d: RUN=%s nivel=%s analito=%s status=%s  flags=%s'
              % (lin, run, niv, an, st0, flags(lin)))
        print('Liberacao inicial: %s' % lib0[:8])

        print('\n=== 1. EDITAR STATUS DIRETO NA CELULA ===')
        db.Cells(lin, 7).Value = 'Excluído'
        time.sleep(0.5)
        f1 = flags(lin)
        print('   apos marcar Excluido -> BB=%r BC=%r' % f1)
        ck('BA:BC atualizados pela edicao direta (vigia disparou)',
           f1[0] in (None, '') and f1[1] in (None, ''), str(f1))

        print('\n=== 2. RESTAURAR STATUS DIRETO NA CELULA ===')
        db.Cells(lin, 7).Value = st0
        time.sleep(0.5)
        f2 = flags(lin)
        print('   apos restaurar %s -> BB=%r BC=%r' % (st0, f2[0], f2[1]))
        ck('BA:BC voltam ao estado original', f2 == (1, 1), str(f2))
        ck('Liberacao volta ao estado original', libA() == lib0,
           '%s vs %s' % (libA()[:6], lib0[:6]))

        print('\n=== 3. ExcluirLogico E RESTAURACAO (o ciclo da prova 4.2) ===')
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect('qcini2025')
        for c in list(wb.VBProject.VBComponents):
            if c.Name == 'mWatchTemp':
                wb.VBProject.VBComponents.Remove(c)
        wb.VBProject.VBComponents.Import(bas)
        if estrutura:
            wb.Protect('qcini2025', True, False)

        n = xl.Run('T_Excluir', run, niv, an)
        time.sleep(0.5)
        print('   ExcluirLogico excluiu %s registro(s) -> flags=%s' % (n, flags(lin)))
        ck('ExcluirLogico recalculou as flags',
           flags(lin)[0] in (None, ''), str(flags(lin)))

        db.Cells(lin, 7).Value = st0        # exatamente como a suite restaura
        time.sleep(0.5)
        f3 = flags(lin)
        libF = libA()
        print('   apos restaurar -> flags=%s' % (f3,))
        print('   Liberacao final: %s' % libF[:8])
        ck('flags restauradas apos o ciclo completo', f3 == (1, 1), str(f3))
        ck('Liberacao identica ao estado inicial (causa da 4.2)', libF == lib0,
           '%s vs %s' % (libF[:6], lib0[:6]))
        ck('eventos nao ficaram presos em False', xl.EnableEvents is True,
           str(xl.EnableEvents))
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
    print('VIGIA DE STATUS: TODOS OS CAMINHOS RECALCULAM')


if __name__ == '__main__':
    main(sys.argv[1])
