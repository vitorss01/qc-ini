# -*- coding: utf-8 -*-
"""ligar_bias_ceq.py - ADR-030: ET e Sigma passam a usar o bias do EP

O DEFEITO

Estatistica!K chamava EstatPeriodo(...,"BIAS"), que resolve em
mEstatistica.CalcularBias(mediaObs, alvoDoLote): media do CONTROLE INTERNO
contra o alvo atribuido AO LOTE. O proprio cabecalho dizia "Bias % (alvo do
lote)". Isso mede deriva do CQI, nao erro sistematico do metodo -- e era esse
numero que entrava em ET e em Sigma.

A aritmetica sempre esteve certa dos dois lados: (X - ref)/ref*100, com
|bias| em ET e em Sigma. O que estava errado era o "ref".

O QUE MUDA

  Estatistica!K   |Bias| EQ %          <- mCEQ.BiasEQ(...,"ABS")
  Estatistica!M   Bias EQ % assinado   <- mCEQ.BiasEQ(...,"SIGNED")
  Estatistica!N   ET %                 <- 1,65*CV + ABS(K), so com os dois
  Estatistica!R   Sigma                <- guarda por ISNUMBER, nao por ""
  Painel!H        ET                   <- passa a usar ABS do bias
  Painel!I        Sigma                <- exige bias numerico, nao assume zero

O ano de referencia e ANO(Estat_Fim_Efetiva), e mCEQ resolve dai a rodada
vigente: o maior ano de EP que nao o ultrapasse -- a mesma regra de vigencia que
o ADR-022 usa para especificacao. Exigir coincidencia exata de ano faria o bias
sumir sempre que o CQI passasse na frente do ultimo ciclo publicado.

A COLUNA M ESTAVA MORTA

Era "Lim Bias FAB %" e virou ="" no ADR-028, porque o fabricante informa ETp e
CVTp, nao bias. Conferido: nenhuma formula do arquivo a referencia. Usa-la para
o bias assinado nao desloca nada e nao quebra nada -- e o lugar que ja existia
para guardar o sinal sem destruir a magnitude em K.

DOIS DEFEITOS MENORES CORRIGIDOS DE PASSAGEM, NO PAINEL

  H7 = E7*1,65 + IF(ISNUMBER(G7);G7;0)

Somava o bias COM SINAL: um bias negativo REDUZIA o erro total. E, faltando
bias, somava zero -- afirmando exatidao que ninguem mediu. I7 (Sigma) ja usava
ABS, mas tambem assumia zero na ausencia. Agora os dois exigem bias numerico e
devolvem vazio quando nao ha.

Uso: python ligar_bias_ceq.py <arquivo.xlsm>
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
EST_R0, EST_RN = 14, 93
C_K, C_M, C_N, C_R = 11, 13, 14, 18

# Ano de referencia = fim do periodo em analise. mCEQ.BiasEQ resolve dai a
# rodada vigente: o maior ano de EP que nao o ultrapasse (mesma regra do
# ADR-022 para especificacao).
ANO_REF = 'YEAR(Estat_Fim_Efetiva)'


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


def escrever(ws, lin, col, formula):
    def por():
        ws.Cells(lin, col).Formula = formula
    tenta(por)


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
        for nm in ('Estatística', 'Painel', 'EQC_Dados'):
            try:
                wb.Worksheets(nm).Unprotect(SENHA)
            except Exception:
                pass
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')

        # ---- 1. modulo mCEQ ------------------------------------------------
        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name == 'mCEQ':
                vbp.VBComponents.Remove(c)
        vbp.VBComponents.Import(MODULO)
        print('modulo mCEQ importado')

        # ---- 2. cabecalhos -------------------------------------------------
        escrever(es, 13, C_K, '|Bias| EQ %')
        escrever(es, 13, C_M, 'Bias EQ % (assinado)')
        escrever(es, 13, C_N, 'ET % (1,65·CV + |Bias|)')
        print('cabecalhos: K, M e N renomeados')

        # ---- 3. as tres colunas -------------------------------------------
        n = 0
        for lin in range(EST_R0, EST_RN + 1):
            escrever(es, lin, C_K,
                     '=IF($A%d="","",mCEQ.BiasEQ($A%d,%s,"ABS"))'
                     % (lin, lin, ANO_REF))
            escrever(es, lin, C_M,
                     '=IF($A%d="","",mCEQ.BiasEQ($A%d,%s,"SIGNED"))'
                     % (lin, lin, ANO_REF))
            # ET so existe com CV e bias. Sem um dos dois, nao e zero: e vazio.
            escrever(es, lin, C_N,
                     '=IF(OR($A{0}="",NOT(ISNUMBER($F{0})),NOT(ISNUMBER($K{0}))),"",'
                     '1.65*$F{0}+ABS($K{0}))'.format(lin))
            # Sigma: o guarda tinha de mudar junto.
            #
            # Ele testava $K="" -- valido quando a ausencia era vazia. Agora a
            # ausencia e o TEXTO "SEM EP", que passa por esse teste e chega em
            # ABS("SEM EP") = #VALOR!. ISNUMBER cobre os dois casos e nao
            # depende de qual marcador o modulo devolve.
            escrever(es, lin, C_R,
                     '=IF(OR($F{0}="",$F{0}=0,NOT(ISNUMBER($O{0})),NOT(ISNUMBER($K{0}))),"",'
                     '($O{0}-ABS($K{0}))/$F{0})'.format(lin))
            n += 4
        print('%d celulas reescritas na Estatistica (K, M, N, R)' % n)

        # ---- 4. Painel: ET com modulo, Sigma sem zero artificial ----------
        for lin in (7, 8):
            escrever(pa, lin, 8,
                     '=IF(OR($E{0}="",NOT(ISNUMBER($E{0})),NOT(ISNUMBER($G{0}))),"",'
                     '1.65*$E{0}+ABS($G{0}))'.format(lin))
            escrever(pa, lin, 9,
                     '=IF(OR($E{0}="",$E{0}=0,NOT(ISNUMBER($F{0})),NOT(ISNUMBER($G{0}))),"",'
                     '($F{0}-ABS($G{0}))/$E{0})'.format(lin))
        print('Painel: H7/H8 (ET) e I7/I8 (Sigma) corrigidos')

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        # ---- 5. conferencia ------------------------------------------------
        erros = 0
        comBias = 0
        for lin in range(EST_R0, EST_RN + 1):
            for c in (C_K, C_M, C_N, C_R):
                t = str(tenta(lambda l=lin, cc=c: es.Cells(l, cc).Text))
                if t.startswith('#'):
                    erros += 1
                    if erros <= 5:
                        print('   ERRO %s%d = %s' % ('KMNR'[(C_K, C_M, C_N, C_R).index(c)], lin, t))
            v = tenta(lambda l=lin: es.Cells(l, C_K).Value)
            if isinstance(v, (int, float)):
                comBias += 1
        print('celulas com erro: %d ; linhas com |Bias| EQ calculado: %d' % (erros, comBias))
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
