# -*- coding: utf-8 -*-
"""aposentar_limespec.py - ADR-028: as ultimas 240 chamadas ao motor antigo saem

O QUE FOI MEDIDO

Estatistica!L, !M e !P chamavam LimEspec(), que le DB_Especificacoes:

    L = LimEspec(analito, ano, "Variacao Biologica", "BIAS")
    M = LimEspec(analito, ano, "Fabricante",        "BIAS")
    P = LimEspec(analito, ano, "Variacao Biologica", "ETP")

Conferido celula a celula: as tres devolvem VAZIO para os 40 analitos. O
DB_Especificacoes nunca foi populado -- o cadastro do ADR-022 ficou como
pendencia e nunca saiu do papel. Sao 240 formulas mortas, e a coluna P alimenta
StatusETP, que portanto compara contra nada.

Enquanto isso, Analitos!R ("ETp VB") tem valor de verdade -- 15,21 para o
Lactato, por exemplo -- calculado a partir de CVi, CVg e do desempenho
escolhido. A informacao existia; estava era no lugar certo e ninguem ia buscar.

O QUE MUDA

    P  <- Analitos!R          ETp da Variacao Biologica (passa a ter valor)
    L  <- bias VB calculado   fb * RAIZ(CVi^2 + CVg^2), o mesmo termo que a
                              propria formula de Analitos!R ja usa
    M  <- ""                  o fabricante informa ETp e CVTp, NAO bias. Nao ha
                              fonte para essa coluna em lugar nenhum do arquivo.

Sobre M: preencher com zero ou com o CVTp seria inventar um dado que o
fabricante nao forneceu. Vazio e a verdade, e agora e um vazio explicito em vez
de um vazio por engrenagem quebrada.

Sobre L: os coeficientes 0,125 / 0,25 / 0,375 sao os mesmos que Analitos!R usa
para OTI / DES / MIN. Escrever o mesmo criterio aqui mantem a conta identica a
que o gestor ja validou -- nao e uma segunda regra.

Depois deste script, nenhuma formula do arquivo chama LimEspec nem le
Cfg_/DB_/Eng_Especificacoes.

Uso: python aposentar_limespec.py <arquivo.xlsm>
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
C_L, C_M, C_P = 12, 13, 16

BUSCA = 'MATCH($A{0},Analitos!$A$4:$A$43,0)'


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


def form_P(lin):
    """ETp da Variacao Biologica, direto de Analitos!R."""
    return ('=IF($A{0}="","",IFERROR(INDEX(Analitos!$R$4:$R$43,{1}),""))'
            .format(lin, BUSCA.format(lin)))


def form_L(lin):
    """Bias permitido pela VB: fb * RAIZ(CVi^2 + CVg^2).

    fb = 0,125 (OTI) / 0,25 (DES) / 0,375 (MIN) -- os mesmos coeficientes que a
    formula de Analitos!R ja aplica. CVi = Analitos!O, CVg = Analitos!P,
    desempenho = Analitos!Q.
    """
    b = BUSCA.format(lin)
    cvi = 'INDEX(Analitos!$O$4:$O$43,%s)' % b
    cvg = 'INDEX(Analitos!$P$4:$P$43,%s)' % b
    des = 'INDEX(Analitos!$Q$4:$Q$43,%s)' % b
    fb = 'IF({0}="ÓTI",0.125,IF({0}="DES",0.25,IF({0}="MIN",0.375,"")))'.format(des)
    return ('=IF($A{0}="","",IFERROR(IF(OR({1}="",{2}=""),"",'
            '{3}*SQRT({1}^2+{2}^2)),""))').format(lin, cvi, cvg, fb)


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
            for c in (C_L, C_M, C_P):
                if 'LimEspec' in str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Formula)):
                    antes += 1
        print('formulas chamando LimEspec antes: %d' % antes)

        n = 0
        for lin in range(R0, RN + 1):
            tenta(lambda l=lin: es.Cells(l, C_P).__setattr__('Formula', form_P(l)))
            tenta(lambda l=lin: es.Cells(l, C_L).__setattr__('Formula', form_L(l)))
            tenta(lambda l=lin: es.Cells(l, C_M).__setattr__('Formula', '=""'))
            n += 3
        print('%d celulas reescritas (P<-Analitos!R, L<-bias VB, M<-vazio)' % n)

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        rest = 0
        erros = 0
        comP = 0
        for lin in range(R0, RN + 1):
            for c in (C_L, C_M, C_P):
                if 'LimEspec' in str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Formula)):
                    rest += 1
                t = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Text))
                if t.startswith('#'):
                    erros += 1
            v = tenta(lambda l=lin: es.Cells(l, C_P).Value)
            if isinstance(v, (int, float)) and v:
                comP += 1
        print('LimEspec restantes: %d ; celulas com erro: %d' % (rest, erros))
        print('linhas com ETp VB agora preenchido: %d (antes: 0)' % comP)
        if rest or erros:
            raise SystemExit('sobrou LimEspec ou erro -- nada salvo')

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
