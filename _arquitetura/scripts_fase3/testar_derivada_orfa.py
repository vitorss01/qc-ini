# -*- coding: utf-8 -*-
"""testar_derivada_orfa.py - a causa raiz da prova 4.2, isolada

O MECANISMO

Apagar A:G de uma linha deixava BA:BC intactas. Uma linha sem RUN com BC=1,
ainda dentro dos intervalos nomeados, faz

    rRUN / ((rRunUnico = 1) * ((""&rLote) = (""&loteAtivo)))

devolver VAZIO/1 = ZERO. O AGREGAR da Liberacao passa a ver uma "corrida 0"
antes de todas e a lista inteira desce uma linha: A4 vira 0, A5 vira 1, A7 vira
3 no lugar de 4. Sao as 52 celulas da prova 4.2.

Enquanto BA:BC eram formula isso se corrigia sozinho. Como VALOR, sobreviviam.

O teste cria a orfa de proposito, confirma que a Liberacao se desloca, chama
AtualizarFlagsBanco e confirma que volta ao lugar.

Uso: python testar_derivada_orfa.py <artefato.xlsm>
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


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AutomationSecurity = 1
            xl.EnableEvents = False      # como a suite: o vigia NAO pode salvar o dia
            return xl
        except Exception:
            if t in (2, 4):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Hidden'],
                                stderr=subprocess.DEVNULL)
                time.sleep(8)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'), 'orfa_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        db = wb.Worksheets('DB_Resultados')
        lib = wb.Worksheets('Liberação')
        try:
            db.Unprotect('qcini2025')
        except Exception:
            pass
        xl.Calculation = -4105

        def libA():
            xl.CalculateFull()
            return [lib.Cells(r, 1).Value for r in range(4, 10)]

        ult = db.Cells(db.Rows.Count, 1).End(-4162).Row
        base = libA()
        print('ultima linha do banco: %d' % ult)
        print('Liberacao A4:A9 inicial: %s' % base)
        ck('estado inicial correto (comeca em 1)', base[0] == 1, str(base))

        print('\n=== 1. CRIAR A ORFA (o que a suite faz ao limpar A:G) ===')
        lin = ult + 1
        # simula a linha importada: entra com dado, ganha as flags, depois A:G e
        # apagado sem que nada recalcule -- exatamente o cenario da linha 671.
        db.Cells(lin, 1).Value = 999
        db.Cells(lin, 5).Value = 'OrfaTeste'
        db.Cells(lin, 7).Value = 'Ativo'
        db.Cells(lin, 4).Value = str(db.Cells(4, 4).Value)
        xl.Run('AtualizarFlagsBanco')
        print('   linha %d criada e marcada: BA=%r BB=%r BC=%r'
              % (lin, db.Cells(lin, 53).Value, db.Cells(lin, 54).Value,
                 db.Cells(lin, 55).Value))
        db.Range(db.Cells(lin, 1), db.Cells(lin, 7)).ClearContents()
        print('   A:G apagado; derivadas restantes: BA=%r BB=%r BC=%r'
              % (db.Cells(lin, 53).Value, db.Cells(lin, 54).Value,
                 db.Cells(lin, 55).Value))
        com_orfa = libA()
        print('   Liberacao A4:A9 com a orfa: %s' % com_orfa)
        ck('a orfa REALMENTE desloca a Liberacao (mecanismo reproduzido)',
           com_orfa != base, '%s vs %s' % (com_orfa, base))

        print('\n=== 2. AtualizarFlagsBanco LIMPA ALEM DO DADO ===')
        xl.Run('AtualizarFlagsBanco')
        print('   derivadas na linha %d: BA=%r BB=%r BC=%r'
              % (lin, db.Cells(lin, 53).Value, db.Cells(lin, 54).Value,
                 db.Cells(lin, 55).Value))
        ck('BA:BC da linha orfa foram apagadas',
           db.Cells(lin, 55).Value in (None, ''), repr(db.Cells(lin, 55).Value))
        depois = libA()
        print('   Liberacao A4:A9 apos: %s' % depois)
        ck('Liberacao volta ao estado inicial', depois == base,
           '%s vs %s' % (depois, base))

        print('\n=== 3. DETERMINISMO: repetir 3 vezes da o mesmo ===')
        vals = []
        for i in range(3):
            xl.Run('AtualizarFlagsBanco')
            vals.append(libA())
        print('   %s' % vals)
        ck('tres execucoes seguidas produzem o mesmo estado',
           all(v == base for v in vals), str(vals))
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
    print('DERIVADA ORFA: MECANISMO REPRODUZIDO E CORRIGIDO')


if __name__ == '__main__':
    main(sys.argv[1])
