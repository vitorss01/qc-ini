# -*- coding: utf-8 -*-
"""ligar_ssot_analitos.py - ADR-027: Analitos!R/S/T como fonte unica de verdade

O QUE ESTAVA ERRADO

Havia TRES fontes concorrentes de ETp no mesmo arquivo:

  Estatistica!O  LimEspec(analito, ano, "CLIA", "ETP") -- CLIA CRAVADO no
                 argumento, ignorando a fonte escolhida em Analitos!R.
                 Alimenta o Sigma da Estatistica (col R).
  Painel!F7/F8   INDEX(engETp, ...) -- motor Eng_Especificacoes (ADR-022).
                 Alimenta o Sigma do Painel (col I).
  Analitos!S     "ETp em uso final %", que escolhe conforme R -- e que NAO era
                 lido por ninguem.

Consequencia concreta: Painel e Estatistica podiam exibir Sigmas DIFERENTES para
o mesmo analito, e trocar a fonte em Analitos!R nao mudava nada em lugar nenhum.

O mesmo valia para o limite de CV: Estatistica!G chamava LimEspec(...,"CLIA","CV")
enquanto Analitos!T ja resolvia o CVTp conforme a fonte.

O QUE ESTE SCRIPT FAZ

  1. cria os nomes espFonte / etpOficial / cvtpOficial sobre Analitos R/S/T
  2. Painel!F7:F8   passa a ler etpOficial
  3. Estatistica!O  passa a ler etpOficial
  4. Estatistica!G  passa a ler cvtpOficial
  5. o Sigma so calcula com ETp NUMERICO

As formulas da aba Analitos NAO sao tocadas. Foram criadas e testadas pelo
gestor; este script apenas liga quem as consome.

SOBRE O GUARDA DO SIGMA

Analitos!S devolve o texto "DEFINIR FONTE QUE CONTENHAM DADOS" quando a fonte
escolhida nao tem dado. Sem guarda o Sigma viraria #VALOR!. ISNUMBER aqui nao
mascara erro: diz "nao ha meta utilizavel", que e a verdade, e e o mesmo criterio
de quatro estados do ADR-023. Zero seria lido como "desempenho pessimo"; vazio e
lido como "sem meta".

Uso: python ligar_ssot_analitos.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
EST_R0, EST_RN = 14, 93
C_G, C_O, C_R = 7, 15, 18          # Estatistica: CVTp, ETp, Sigma
P_F, P_I = 6, 9                    # Painel: ETp, Sigma


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
    """COM com retentativa.

    Escrever formula numa pasta pesada dispara recalculo, e o Excel devolve
    RPC_E_CALL_REJECTED enquanto esta ocupado. Nao e defeito -- e o servidor
    dizendo "agora nao". Abortar deixaria a planilha meio convertida, que e o
    pior estado possivel.
    """
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


def ler(ws, lin, col):
    return str(tenta(lambda: ws.Cells(lin, col).Formula))


def escrever(ws, lin, col, formula):
    def por():
        ws.Cells(lin, col).Formula = formula
    tenta(por)


def idx(nome, chave):
    return 'INDEX(%s,MATCH(%s,Analitos!$A$4:$A$43,0))' % (nome, chave)


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
        xl.Calculation = -4135       # manual: escrever nao recalcula a cada celula
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        for nm in ('Analitos', 'Painel', 'Estatística'):
            try:
                wb.Worksheets(nm).Unprotect(SENHA)
            except Exception:
                pass
        pa = wb.Worksheets('Painel')
        es = wb.Worksheets('Estatística')

        # ---- 1. nomes ------------------------------------------------------
        for n, ref in (('espFonte', '=Analitos!$R$4:$R$43'),
                       ('etpOficial', '=Analitos!$S$4:$S$43'),
                       ('cvtpOficial', '=Analitos!$T$4:$T$43')):
            try:
                wb.Names(n).Delete()
            except Exception:
                pass
            wb.Names.Add(n, ref)
        print('nomes: espFonte, etpOficial, cvtpOficial -> Analitos R/S/T')

        # ---- 2. Painel: ETp deixa de vir do motor --------------------------
        tocadas = []
        for lin in (7, 8):
            if 'engETp' not in ler(pa, lin, P_F):
                print('   Painel!F%d ja nao usava engETp' % lin)
                continue
            escrever(pa, lin, P_F, '=IFERROR(%s,"")' % idx('etpOficial', 'selAnalito'))
            tocadas.append('F%d' % lin)
        print('Painel: %s -> etpOficial' % (', '.join(tocadas) if tocadas else '(nada)'))

        # ---- 3/4. Estatistica ---------------------------------------------
        nO = nG = 0
        for lin in range(EST_R0, EST_RN + 1):
            if 'LimEspec' in ler(es, lin, C_O):
                escrever(es, lin, C_O, '=IF($A%d="","",IFERROR(%s,""))'
                         % (lin, idx('etpOficial', '$A%d' % lin)))
                nO += 1
            if 'LimEspec' in ler(es, lin, C_G):
                escrever(es, lin, C_G, '=IF($A%d="","",IFERROR(%s,""))'
                         % (lin, idx('cvtpOficial', '$A%d' % lin)))
                nG += 1
        print('Estatistica: O -> etpOficial em %d celulas ; G -> cvtpOficial em %d' % (nO, nG))

        # ---- 5. Sigma so com ETp numerico ----------------------------------
        nS = 0
        for lin in range(EST_R0, EST_RN + 1):
            f = ler(es, lin, C_R)
            if f.startswith('=') and 'ISNUMBER($O%d)' % lin not in f:
                escrever(es, lin, C_R,
                         '=IF(OR($F{0}="",$F{0}=0,NOT(ISNUMBER($O{0})),$K{0}=""),"",'
                         '($O{0}-ABS($K{0}))/$F{0})'.format(lin))
                nS += 1
        print('Estatistica: %d celulas de Sigma com guarda ISNUMBER' % nS)

        nP = 0
        for lin in (7, 8):
            f = ler(pa, lin, P_I)
            if f.startswith('=') and 'ISNUMBER($F%d)' % lin not in f:
                escrever(pa, lin, P_I,
                         '=IF(OR($E{0}="",$E{0}=0,NOT(ISNUMBER($F{0}))),"",'
                         '($F{0}-IF(ISNUMBER($G{0}),ABS($G{0}),0))/$E{0})'.format(lin))
                nP += 1
        print('Painel: %d celulas de Sigma com guarda ISNUMBER e ABS no bias' % nP)

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
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
