# -*- coding: utf-8 -*-
"""corrigir_estat_cv_ref.py - ADR-028: fim dos 76 #REF! da Estatistica

O DEFEITO, E ELE E ESTRUTURAL

Estatistica!H e !I traziam CVi e CVTp do fabricante assim:

    H14 = IF(Analitos!N4="","",Analitos!N4)
    H15 = IF(Analitos!N5="","",Analitos!N5)

Mapeamento POSICIONAL: linha 14 da Estatistica <-> linha 4 da Analitos. So que a
Estatistica tem 40 analitos x 2 NIVEIS -- linhas 14..53 sao o nivel 1 e 54..93 o
nivel 2. Do 54 em diante a conta apontava para Analitos!N44..N83, fora do bloco
de 40 analitos. Quando as colunas da Analitos foram reorganizadas, essas
referencias viraram #REF!: 76 celulas.

O numero 76 e sintoma, nao causa. A causa e amarrar duas tabelas por POSICAO
quando elas tem granularidades diferentes. Repor a coluna certa mantendo o
deslocamento so adiaria o problema ate a proxima insercao de linha.

A CORRECAO

Busca por CHAVE -- o nome do analito -- como todo o resto do sistema ja faz:

    H = INDEX(Analitos!$O$4:$O$43, MATCH($A{linha}, Analitos!$A$4:$A$43, 0))   CVi %
    I = INDEX(Analitos!$N$4:$N$43, MATCH($A{linha}, Analitos!$A$4:$A$43, 0))   CVTp FAB %

Passa a valer para os dois niveis, e nao depende mais de a Estatistica e a
Analitos terem o mesmo numero de linhas na mesma ordem.

O IFERROR aqui cobre "este analito nao esta na Analitos", que e estado de dado
legitimo -- nao esconde referencia quebrada: referencia quebrada e justamente o
que este script elimina.

Uso: python corrigir_estat_cv_ref.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
R0, RN = 14, 93
C_H, C_I = 8, 9

# coluna da Analitos que cada uma representa, no layout atual
ORIGEM = {C_H: ('O', 'CVi %'), C_I: ('N', 'CVTp FAB %')}


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
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        xl.Calculation = -4135
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        try:
            wb.Worksheets('Estatística').Unprotect(SENHA)
        except Exception:
            pass
        es = wb.Worksheets('Estatística')

        antes = 0
        for lin in range(R0, RN + 1):
            for c in (C_H, C_I):
                f = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Formula))
                if '#REF!' in f:
                    antes += 1
        print('celulas com #REF! antes: %d' % antes)

        n = 0
        for lin in range(R0, RN + 1):
            for c, (colA, rot) in ORIGEM.items():
                nova = ('=IF($A{0}="","",IFERROR(INDEX(Analitos!${1}$4:${1}$43,'
                        'MATCH($A{0},Analitos!$A$4:$A$43,0)),""))').format(lin, colA)
                tenta(lambda l=lin, cc=c, ff=nova: es.Cells(l, cc).__setattr__('Formula', ff))
                n += 1
        print('%d celulas reescritas (H <- Analitos!O CVi, I <- Analitos!N CVTp FAB)' % n)

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        depois = 0
        erros = 0
        for lin in range(R0, RN + 1):
            for c in (C_H, C_I):
                f = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Formula))
                if '#REF!' in f:
                    depois += 1
                t = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Text))
                if t.startswith('#'):
                    erros += 1
        print('celulas com #REF! depois: %d ; celulas exibindo erro: %d' % (depois, erros))
        if depois or erros:
            raise SystemExit('ainda ha erro em H/I -- nada salvo')

        if estrutura and not wb.ProtectStructure:
            wb.Protect(SENHA, True, False)
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
