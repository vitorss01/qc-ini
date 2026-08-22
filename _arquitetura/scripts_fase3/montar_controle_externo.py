# -*- coding: utf-8 -*-
"""montar_controle_externo.py - ADR-032: a aba de EP vira usavel

O QUE MUDA NA EQC_Dados

  C  Rodada       vira LETRA (A, B, C, D) com validacao.
                  Estava 1/2/3 numerico. O laboratorio fala "rodada A"; e o
                  numero se confundia com contagem de amostra.
  E  Provedor     validacao Controllab / CAP
  P  Status SDI   NOVO: |SDI| <= 2 aprova, acima reprova

  Digitados pelo usuario: G resultado do lab, H media do grupo, I SD do grupo,
  K limite inferior, L limite superior. O resto e calculado.

O QUE MUDA NA Estatistica

  Bloco de filtros do EP em K3:P5 -- provedor, ano e rodada. Os tres viram
  nomes (eqProvedor, eqAnoEP, eqRodada) e alimentam TODAS as colunas de EP,
  para que nao exista um card lendo um filtro e outro lendo outro.

  T  Status SDI        consolidado do analito
  U  Status limites    consolidado do analito

  K, M, N, R passam a receber os tres filtros.

SEPARADOR DA VALIDACAO

No XML e no COM a lista vai com VIRGULA ("Controllab,CAP"); o Excel exibe e
aceita ponto-e-virgula na caixa de dialogo, conforme a configuracao regional.
Escrever ";" aqui gravaria um item unico chamado "Controllab;CAP". As
validacoes que ja existiam nesta aba usam virgula -- e o padrao comprovado.

RODADAS: 4 PARA A CONTROLLAB, 3 PARA O CAP

A validacao oferece A..D. O CAP simplesmente nao usa a D. Restringir a lista por
provedor exigiria validacao dependente, que quebra ao copiar linha -- e o custo
nao paga: rodada D de CAP nao existe nos dados e nao entra em media nenhuma.

Uso: python montar_controle_externo.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
MODULO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      'src_producao', 'mCEQ.bas')

EQ_R0, EQ_RN = 4, 1003
C_RODADA, C_PROV, C_SDI, C_STATUS_SDI = 3, 5, 10, 16

EST_R0, EST_RN = 14, 93
C_K, C_M, C_N, C_R = 11, 13, 14, 18
C_T, C_U = 20, 21                      # status SDI e status limites

LETRAS = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F'}


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


def por(ws, lin, col, valor, formula=False):
    def f():
        if formula:
            ws.Cells(lin, col).Formula = valor
        else:
            ws.Cells(lin, col).Value = valor
    tenta(f)


def valida(ws, ref, lista):
    r = ws.Range(ref)
    try:
        r.Validation.Delete()
    except Exception:
        pass
    r.Validation.Add(3, 1, 1, lista)      # xlValidateList, stop
    r.Validation.InCellDropdown = True
    r.Validation.IgnoreBlank = True


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
        for nm in ('EQC_Dados', 'Estatística', 'Painel'):
            try:
                wb.Worksheets(nm).Unprotect(SENHA)
            except Exception:
                pass
        eq = wb.Worksheets('EQC_Dados')
        es = wb.Worksheets('Estatística')

        # ---- 0. modulo -----------------------------------------------------
        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name == 'mCEQ':
                vbp.VBComponents.Remove(c)
        vbp.VBComponents.Import(MODULO)
        print('mCEQ reimportado')

        # ---- 1. rodada numerica vira letra --------------------------------
        ult = eq.Cells(eq.Rows.Count, 1).End(-4162).Row
        conv = 0
        for lin in range(EQ_R0, max(ult, EQ_R0) + 1):
            v = tenta(lambda l=lin: eq.Cells(l, C_RODADA).Value)
            if isinstance(v, (int, float)) and int(v) in LETRAS:
                por(eq, lin, C_RODADA, LETRAS[int(v)])
                conv += 1
        print('rodadas convertidas para letra: %d (ate a linha %d)' % (conv, ult))

        # ---- 2. validacoes na EQC_Dados -----------------------------------
        valida(eq, eq.Range(eq.Cells(EQ_R0, C_RODADA), eq.Cells(EQ_RN, C_RODADA)).Address,
               'A,B,C,D')
        valida(eq, eq.Range(eq.Cells(EQ_R0, C_PROV), eq.Cells(EQ_RN, C_PROV)).Address,
               'Controllab,CAP')
        print('validacoes: Rodada A..D ; Provedor Controllab/CAP')

        # ---- 3. Status SDI na EQC_Dados -----------------------------------
        por(eq, 3, C_STATUS_SDI, 'Status SDI')
        tenta(lambda: eq.Cells(3, C_STATUS_SDI).Font.__setattr__('Bold', True))
        for lin in range(EQ_R0, EQ_RN + 1):
            por(eq, lin, C_STATUS_SDI,
                '=IF($J{0}="","",IF(ABS($J{0})<=2,"OK","FORA (|SDI|>2)"))'.format(lin),
                formula=True)
        print('EQC_Dados!P: Status SDI em %d linhas' % (EQ_RN - EQ_R0 + 1))

        # ---- 4. bloco de filtros na Estatistica ---------------------------
        anos = set()
        for lin in range(EQ_R0, max(ult, EQ_R0) + 1):
            a = tenta(lambda l=lin: eq.Cells(l, 2).Value)
            if isinstance(a, (int, float)):
                anos.add(int(a))
        listaAnos = ','.join(str(a) for a in sorted(anos)) or '2025'

        por(es, 3, 11, 'CONTROLE EXTERNO — origem do bias')
        tenta(lambda: es.Cells(3, 11).Font.__setattr__('Bold', True))
        por(es, 4, 11, 'Provedor')
        por(es, 4, 13, 'Ano EP')
        por(es, 4, 15, 'Rodada')
        por(es, 4, 12, 'Controllab')
        por(es, 4, 14, sorted(anos)[-1] if anos else 2025)
        por(es, 4, 16, 'TODAS')
        valida(es, 'L4', 'Controllab,CAP')
        valida(es, 'N4', listaAnos)
        valida(es, 'P4', 'TODAS,A,B,C,D')
        por(es, 5, 11, '=ResumoFiltroEQ(eqProvedor,eqAnoEP,eqRodada)', formula=True)
        tenta(lambda: es.Cells(5, 11).Font.__setattr__('Italic', True))

        for n, ref in (('eqProvedor', '=Estatística!$L$4'),
                       ('eqAnoEP', '=Estatística!$N$4'),
                       ('eqRodada', '=Estatística!$P$4')):
            try:
                wb.Names(n).Delete()
            except Exception:
                pass
            wb.Names.Add(n, ref)
        print('filtros do EP: L4 provedor, N4 ano (%s), P4 rodada' % listaAnos)

        # ---- 5. colunas da Estatistica passam a usar os filtros -----------
        por(es, 13, C_T, 'Status SDI (EP)')
        por(es, 13, C_U, 'Status limites (EP)')
        for c in (C_T, C_U):
            tenta(lambda cc=c: es.Cells(13, cc).Font.__setattr__('Bold', True))

        n = 0
        for lin in range(EST_R0, EST_RN + 1):
            por(es, lin, C_K,
                '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"ABS",eqProvedor,eqRodada))'
                .format(lin), formula=True)
            por(es, lin, C_M,
                '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"SIGNED",eqProvedor,eqRodada))'
                .format(lin), formula=True)
            por(es, lin, C_N,
                '=IF(OR($A{0}="",NOT(ISNUMBER($F{0})),NOT(ISNUMBER($K{0}))),"",'
                '1.65*$F{0}+ABS($K{0}))'.format(lin), formula=True)
            por(es, lin, C_R,
                '=IF(OR($F{0}="",$F{0}=0,NOT(ISNUMBER($O{0})),NOT(ISNUMBER($K{0}))),"",'
                '($O{0}-ABS($K{0}))/$F{0})'.format(lin), formula=True)
            por(es, lin, C_T,
                '=IF($A{0}="","",mCEQ.StatusSDIeq($A{0},eqAnoEP,eqProvedor,eqRodada))'
                .format(lin), formula=True)
            por(es, lin, C_U,
                '=IF($A{0}="","",mCEQ.StatusLimitesEQ($A{0},eqAnoEP,eqProvedor,eqRodada))'
                .format(lin), formula=True)
            n += 6
        print('%d celulas na Estatistica (K, M, N, R, T, U)' % n)

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        # ---- 6. conferencia ------------------------------------------------
        erros = 0
        comBias = 0
        for lin in range(EST_R0, EST_RN + 1):
            for c in (C_K, C_M, C_N, C_R, C_T, C_U):
                t = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Text))
                if t.startswith('#'):
                    erros += 1
                    if erros <= 5:
                        print('   ERRO linha %d col %d = %s' % (lin, c, t))
            if isinstance(tenta(lambda l=lin: es.Cells(l, C_K).Value), (int, float)):
                comBias += 1
        print('erros: %d ; linhas com |Bias| EQ: %d' % (erros, comBias))
        print('resumo do filtro: %r' % tenta(lambda: es.Cells(5, 11).Value))
        if erros:
            raise SystemExit('erro de formula -- nada salvo')

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
