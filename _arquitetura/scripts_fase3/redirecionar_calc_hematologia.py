# -*- coding: utf-8 -*-
"""redirecionar_calc_hematologia.py - ADR-053: Calc para de calcular sozinho

O DEFEITO

A aba Calc da Hematologia tinha, nas colunas 1-10 e 17-22 de cada bloco de
nivel, VALORES ESTATICOS -- numeros colados numa execucao passada, sem
formula nenhuma. So as colunas 11-16 (flags de Westgard e veredicto) ja eram
formula viva, lendo Eng_Saida via INDEX/MATCH (engDados/engChave).

Prova: RUN 1 N1, Z gravado = 0,689498. BI_Data recalculado (fonte atual) =
0,468809. Motor e Calc divergiam porque Calc simplesmente parou de
acompanhar o dado -- um numero congelado nunca vai bater com um recalculo.

A CORRECAO: O MESMO PADRAO DA BIOQUIMICA, PARA 3 NIVEIS

A Bioquimica reconcilia a ZERO divergencias porque seu Calc usa FORMULA VIVA,
ancorada na mesma fonte canonica que o motor usa para o lote ativo
(Analitos!E:J, que reflete o LotesStore do lote em tela): Z = (valor - media) /
DP, calculado com AS MESMAS medias e DP que AlvoAnalito devolve no motor. Nao
e "ler o Z do motor" -- e "calcular com a MESMA equacao, a partir da MESMA
fonte", que e algebricamente identico e nao inventa uma segunda definicao.

Esta rotina reproduz ESSE padrao para a Hematologia, estendido ao 3o nivel
(Analitos!I:J), sem tocar nas colunas 11-16 (Westgard), que ja estavam certas.

O QUE NAO MUDA

Layout, cores, posicoes, titulos: nada. Cada celula recebe a MESMA formula
que ja existe na Bioquimica na posicao equivalente -- so a letra da coluna
muda, pelo deslocamento de bloco (22 colunas por nivel) e pela ancora do
nivel (N1: Analitos E/F, N2: G/H, N3: I/J).

Uso: python redirecionar_calc_hematologia.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
LIN_INI, LIN_FIM = 3, 182          # mesma capacidade da Bioquimica
NLV = 3
BLOCO = 22

# colunas 0-indexadas DENTRO do bloco (base = 1a coluna do bloco do nivel)
OFF_VAL, OFF_Z, OFF_LINE, OFF_OK, OFF_REJ = 5, 6, 7, 8, 9
OFF_STATUS = 15                      # col16 do bloco (P no nivel 1) -- NAO tocada
OFF_M3, OFF_M2, OFF_M1, OFF_MED, OFF_P1, OFF_P2 = 16, 17, 18, 19, 20, 21

# ancora (media, DP) por nivel: colunas de Analitos e celula-ancora em Calc!linha1
#
# O bloco por nivel tem 22 colunas (val..calib: val,z,line,ok,rej,r1-r5,status,
# m3,m2,m1,med,p1,p2,p3,rep1,rep2,rep3,calib) -- confirmado nos dois produtos,
# 3 niveis ocupam ate a coluna 71 (50 + 22 - 1). Verificado por varredura ao
# vivo: colunas 86..150 estao livres em toda a altura da aba (linhas 1..185).
# Longe do bloco de dados, sem risco de colidir com p3/rep/calib ou com os
# nomes de eixo de grafico (axmin/axmax vivem em BZ..CG = col78..85).
ANCORA_ANALITOS = {1: ('E', 'F'), 2: ('G', 'H'), 3: ('I', 'J')}
ANCORA_CALC_COL = {1: (90, 91), 2: (92, 93), 3: (94, 95)}


def letra(col):
    s = ''
    while col > 0:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    try:
        if wb.ReadOnly:
            raise SystemExit('somente leitura: %s' % caminho)

        estrutura = bool(wb.ProtectStructure)
        if estrutura:
            wb.Unprotect(SENHA)
        ws = wb.Worksheets('Calc')
        reprot = bool(ws.ProtectContents)
        if reprot:
            try:
                ws.Unprotect(SENHA)
            except Exception:
                ws.Unprotect()

        # --- ancoras (media, DP) por nivel, iguais ao padrao da Bioquimica ---
        for nv in range(1, NLV + 1):
            cMed, cDp = ANCORA_ANALITOS[nv]
            colM, colD = ANCORA_CALC_COL[nv]
            ws.Cells(1, colM).Formula = (
                '=IFERROR(INDEX(Analitos!$%s$4:$%s$43,'
                'MATCH(selAnalito,Analitos!$A$4:$A$43,0)),"")' % (cMed, cMed))
            ws.Cells(1, colD).Formula = (
                '=IFERROR(INDEX(Analitos!$%s$4:$%s$43,'
                'MATCH(selAnalito,Analitos!$A$4:$A$43,0)),"")' % (cDp, cDp))
        print('ancoras de media/DP escritas para os 3 niveis')

        # --- colunas genericas (1-5), identicas em qualquer nivel/produto ---
        F_IDX = None   # coluna 1: literal, nunca foi formula -- nao tocar
        F_SEQ = ('=IF(selAnalito="","",IFERROR(AGGREGATE(15,6,'
                 'rRUN/((rAnalito=selAnalito)*(rFirst=1)*'
                 '((""&rLote)=(""&loteAtivo))),$A{r}),""))')
        F_DATA = ('=IF($B{r}="","",MAXIFS(rData,rAnalito,selAnalito,rRUN,$B{r},'
                  'rLote,loteAtivo,rStatus,"Ativo"))')
        F_FILTRO = ('=IF($C{r}="",0,IF(AND(OR(filtroDe="",$C{r}>=filtroDe),'
                    'OR(filtroAte="",$C{r}<=filtroAte),'
                    'OR(NOT(OR(qsel1,qsel2,qsel3,qsel4)),'
                    'AND(qsel1,ROUNDUP(MONTH($C{r})/3,0)=1),'
                    'AND(qsel2,ROUNDUP(MONTH($C{r})/3,0)=2),'
                    'AND(qsel3,ROUNDUP(MONTH($C{r})/3,0)=3),'
                    'AND(qsel4,ROUNDUP(MONTH($C{r})/3,0)=4))),1,0))')
        F_XDATA = '=IF(OR($B{r}="",$D{r}=0),NA(),$A{r})'

        n = 0
        for r in range(LIN_INI, LIN_FIM + 1):
            ws.Cells(r, 2).Formula = F_SEQ.format(r=r)
            ws.Cells(r, 3).Formula = F_DATA.format(r=r)
            ws.Cells(r, 4).Formula = F_FILTRO.format(r=r)
            ws.Cells(r, 5).Formula = F_XDATA.format(r=r)
            n += 1
        print('colunas B-E (genericas) reescritas em %d linhas' % n)

        # --- blocos por nivel ---
        for nv in range(1, NLV + 1):
            base = (nv - 1) * BLOCO + 1
            cVal = letra(base + OFF_VAL)
            cZ = letra(base + OFF_Z)
            cStatus = letra(base + OFF_STATUS)
            colM, colD = ANCORA_CALC_COL[nv]
            lM, lD = letra(colM), letra(colD)

            f_val = ('=IF($B{r}="","",IF(COUNTIFS(rAnalito,selAnalito,'
                     'rNivel,%d,rRUN,$B{r},rLote,loteAtivo,rStatus,"Ativo")=0,'
                     '"",SUMIFS(rValor,rAnalito,selAnalito,rNivel,%d,rRUN,$B{r},'
                     'rLote,loteAtivo,rStatus,"Ativo")))') % (nv, nv)
            f_z = '=IF(OR($%s{r}="",$%s$1=""),"",($%s{r}-$%s$1)/$%s$1)' % (
                cVal, lD, cVal, lM, lD)
            f_line = '=IF(AND($%s{r}<>"",$D{r}=1),$%s{r},NA())' % (cVal, cVal)
            f_ok = '=IF(AND($%s{r}<>"",$D{r}=1,$%s{r}="OK"),$%s{r},NA())' % (
                cVal, cStatus, cVal)
            f_rej = ('=IF(AND($%s{r}<>"",$D{r}=1,$%s{r}<>"OK",$%s{r}<>""),'
                     '$%s{r},NA())') % (cVal, cStatus, cStatus, cVal)
            f_m3 = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1+(-3)*$%s$1,NA())' % (lM, lM, lD)
            f_m2 = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1+(-2)*$%s$1,NA())' % (lM, lM, lD)
            f_m1 = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1+(-1)*$%s$1,NA())' % (lM, lM, lD)
            f_med = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1,NA())' % (lM, lM)
            f_p1 = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1+(1)*$%s$1,NA())' % (lM, lM, lD)
            f_p2 = '=IF(AND($D{r}=1,$%s$1<>""),$%s$1+(2)*$%s$1,NA())' % (lM, lM, lD)

            colVal, colZ = base + OFF_VAL, base + OFF_Z
            colLine, colOk, colRej = base + OFF_LINE, base + OFF_OK, base + OFF_REJ
            colM3, colM2, colM1 = base + OFF_M3, base + OFF_M2, base + OFF_M1
            colMed, colP1, colP2 = base + OFF_MED, base + OFF_P1, base + OFF_P2

            for r in range(LIN_INI, LIN_FIM + 1):
                ws.Cells(r, colVal).Formula = f_val.format(r=r)
                ws.Cells(r, colZ).Formula = f_z.format(r=r)
                ws.Cells(r, colLine).Formula = f_line.format(r=r)
                ws.Cells(r, colOk).Formula = f_ok.format(r=r)
                ws.Cells(r, colRej).Formula = f_rej.format(r=r)
                ws.Cells(r, colM3).Formula = f_m3.format(r=r)
                ws.Cells(r, colM2).Formula = f_m2.format(r=r)
                ws.Cells(r, colM1).Formula = f_m1.format(r=r)
                ws.Cells(r, colMed).Formula = f_med.format(r=r)
                ws.Cells(r, colP1).Formula = f_p1.format(r=r)
                ws.Cells(r, colP2).Formula = f_p2.format(r=r)
            print('nivel %d: bloco (colunas %s..%s) reescrito, %d linhas'
                  % (nv, letra(base), letra(base + BLOCO - 1),
                     LIN_FIM - LIN_INI + 1))

        xl.CalculateFullRebuild()

        # --- QA imediato: nenhuma celula em erro no bloco reescrito ---
        erros = []
        for r in range(LIN_INI, min(LIN_FIM, LIN_INI + 200) + 1):
            for nv in range(1, NLV + 1):
                base = (nv - 1) * BLOCO + 1
                for off in (OFF_VAL, OFF_Z):
                    t = str(ws.Cells(r, base + off).Text or '')
                    if t.startswith('#') and t not in ('#N/A',):
                        erros.append('%s%d = %s' % (letra(base + off), r, t))
        if erros:
            raise SystemExit('celulas em erro apos redirecionar -- nada salvo:\n  %s'
                             % '\n  '.join(erros[:10]))
        print('QA: nenhuma celula em erro nos blocos reescritos')

        if reprot:
            try:
                ws.Protect(SENHA)
            except Exception:
                pass
        if estrutura:
            try:
                wb.Protect(SENHA, True, False)
            except Exception:
                pass
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
    return 0


if __name__ == '__main__':
    sys.exit(main())
