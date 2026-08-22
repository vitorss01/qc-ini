# -*- coding: utf-8 -*-
"""reestruturar_qualidade.py - ADR-033: bias, especificacao, sigma e orcamento

O QUE ESTE SCRIPT MONTA

  Estatistica  G..T   as quatro familias de indicador deixam de estar
                      embaralhadas: bias do EP, especificacao vigente, Six
                      Sigma e orcamento de erro (ET contra ETp) passam a
                      ocupar blocos vizinhos e nomeados.
  Estatistica  96+    resumo de conformidade e a lista dos analitos que
                      estao em margem critica ou ja estouraram o ETp.
  Painel       J10+   dois blocos novos -- SIX SIGMA e ERRO TOTAL vs ETp --
                      separados do bloco descritivo e do bloco Westgard.

POR QUE ABS() NO BIAS

Westgard S, Bayat H, Westgard JO. Analytical Sigma metrics: A review of Six
Sigma implementation tools for medical laboratories. Biochem Med (Zagreb)
2018;28(2):020502, pagina 3: "SM = (TEa% - bias%) / CV. This form of the
equation assumes all variables will be expressed in percentages, and the bias
will be an absolute percentage (the presence of any bias always shrinks the
allowable error, never enlarges it)."

As formulas que estavam na pasta -- H = (F*1,65)+G e L = (K-G)/F -- usavam o
bias COM SINAL. Um bias de -8% aumentava o Sigma em vez de reduzi-lo, e
encolhia o ET em vez de aumenta-lo. Os dois sentidos ficam invertidos em
relacao ao modelo. Ver PROVA 2 em testar_qualidade.py.

SIGMA BAIXO NAO E REPROVACAO

Mesmo artigo, paginas 8 e 9: metodo de Sigma baixo exige mais regras, limites
mais estreitos, mais controles e CQ mais frequente -- "For Three Sigma methods
and lower, however, QC frequency must be greatly increased". A coluna M
classifica o desempenho analitico do metodo; ela nao reprova corrida. Quem
reprova corrida e Westgard, e continua sendo so Westgard.

Uso: python reestruturar_qualidade.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

EST_R0, EST_RN = 14, 93
RES_R0 = 96                       # bloco de resumo, abaixo da tabela
LST_R0 = RES_R0 + 10              # lista de analitos criticos
LST_N = 16                        # quantas linhas de lista
C_HELPER = 21                     # U: indice sequencial dos analitos sinalizados
MAP = 'Analitos!$A$4:$A$43'

CAB_EST = [
    (7,  'Bias EQC (abs) %'),
    (8,  'ET %'),
    (9,  'ESPQ FONTE'),
    (10, 'CVTp %'),
    (11, 'ETp %'),
    (12, 'SIX SIGMA'),
    (13, 'Status sigma'),
    (14, 'Margem ETp (p.p.)'),
    (15, 'Margem ETp %'),
    (16, 'Status margem'),
    (17, 'Status CV'),
    (18, 'Status SDI (EP)'),
    (19, 'Status limites (EP)'),
    (20, 'Bias EQC (sinal) %'),
]

CLASSES = ['Classe mundial', 'Excelente', 'Bom', 'Marginal', 'Inadequado']
MARGENS = ['Dentro do orcamento', 'Margem critica', 'ETp excedido']


# A escada de classificacao aparece em tres lugares (Estatistica M, Painel O12,
# BI). Ela e escrita UMA vez aqui e reutilizada, para nao existir versao B.
def escada_sigma(ref):
    return ('IF({0}>=6,"Classe mundial",IF({0}>=5,"Excelente",'
            'IF({0}>=4,"Bom",IF({0}>=3,"Marginal","Inadequado"))))').format(ref)


def escada_margem(ref):
    return ('IF({0}<0,"ETp excedido",IF({0}<=10,"Margem critica",'
            '"Dentro do orcamento"))').format(ref)


def formulas_est(r):
    ix = 'MATCH($A{0},{1},0)'.format(r, MAP)
    return {
        7:  '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"ABS",eqProvedor,eqRodada))'.format(r),
        8:  '=IF(OR(NOT(ISNUMBER($F{0})),NOT(ISNUMBER($G{0}))),"",'
            '1.65*$F{0}+ABS($G{0}))'.format(r),
        9:  '=IF($A{0}="","",IFERROR(INDEX(espFonte,{1}),""))'.format(r, ix),
        10: '=IF($A{0}="","",IFERROR(INDEX(cvtpOficial,{1}),""))'.format(r, ix),
        11: '=IF($A{0}="","",IFERROR(INDEX(etpOficial,{1}),""))'.format(r, ix),
        12: '=IF(OR(NOT(ISNUMBER($F{0})),$F{0}=0,NOT(ISNUMBER($G{0})),'
            'NOT(ISNUMBER($K{0}))),"",($K{0}-ABS($G{0}))/$F{0})'.format(r),
        13: '=IF(NOT(ISNUMBER($L{0})),"",{1})'.format(r, escada_sigma('$L%d' % r)),
        14: '=IF(OR(NOT(ISNUMBER($K{0})),NOT(ISNUMBER($H{0}))),"",$K{0}-$H{0})'.format(r),
        15: '=IF(OR(NOT(ISNUMBER($N{0})),NOT(ISNUMBER($K{0})),$K{0}=0),"",'
            '$N{0}/$K{0}*100)'.format(r),
        16: '=IF(NOT(ISNUMBER($O{0})),"",{1})'.format(r, escada_margem('$O%d' % r)),
        17: '=IF(OR(NOT(ISNUMBER($F{0})),NOT(ISNUMBER($J{0}))),"",'
            'IF($F{0}<=$J{0},"OK","CV acima da meta"))'.format(r),
        18: '=IF($A{0}="","",mCEQ.StatusSDIeq($A{0},eqAnoEP,eqProvedor,eqRodada))'.format(r),
        19: '=IF($A{0}="","",mCEQ.StatusLimitesEQ($A{0},eqAnoEP,eqProvedor,eqRodada))'.format(r),
        20: '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"SIGNED",eqProvedor,eqRodada))'.format(r),
        # helper: numera 1,2,3... apenas nas linhas sinalizadas, para a lista
        # de criticos poder usar MATCH em vez de formula matricial (CSE).
        C_HELPER: '=IF(OR($P{0}="ETp excedido",$P{0}="Margem critica"),'
                  'COUNTIF($P$14:$P{0},"ETp excedido")+'
                  'COUNTIF($P$14:$P{0},"Margem critica"),"")'.format(r),
    }


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


# .Text devolve o que esta RENDERIZADO. Coluna estreita com formato 0,00 vira
# "####", que comeca com "#" e passaria por erro de formula. O erro de verdade
# tem letra depois do "#" (#N/D, #REF!, #VALOR!, #NOME?, #DIV/0!).
def eh_erro(txt):
    t = str(txt)
    return t.startswith('#') and t.strip('#') != ''


def por(ws, lin, col, valor, formula=False, negrito=False, fmt=None):
    def f():
        c = ws.Cells(lin, col)
        if valor is not None:
            if formula:
                c.Formula = valor
            else:
                c.Value = valor
        if negrito:
            c.Font.Bold = True
        if fmt:
            c.NumberFormat = fmt
    tenta(f)


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura (aberto no Excel?): %s' % caminho)
    try:
        xl.Calculation = -4135
        es = wb.Worksheets('Estatística')
        pa = wb.Worksheets('Painel')

        # ---- 1. Estatistica: cabecalhos ---------------------------------
        for c, txt in CAB_EST:
            por(es, 13, c, txt, negrito=True)
        por(es, 13, C_HELPER, 'ordem crítico', negrito=True)
        print('Estatistica L13: %d cabecalhos (G..T) + helper em U' % len(CAB_EST))

        # ---- 2. Estatistica: formulas -----------------------------------
        n = 0
        for r in range(EST_R0, EST_RN + 1):
            for c, f in formulas_est(r).items():
                por(es, r, c, f, formula=True)
                n += 1
        for c in (7, 8, 10, 11, 12, 14, 15, 20):
            tenta(lambda cc=c: es.Range(es.Cells(EST_R0, cc),
                                        es.Cells(EST_RN, cc))
                  .__setattr__('NumberFormat', '0.00'))
        tenta(lambda: es.Range(es.Cells(EST_R0, C_HELPER),
                               es.Cells(EST_RN, C_HELPER))
              .EntireColumn.__setattr__('Hidden', True))
        print('Estatistica G14:U93: %d formulas (coluna U oculta, e apoio)' % n)

        # ---- 3. Estatistica: resumo de conformidade ---------------------
        por(es, RES_R0, 1, 'RESUMO — SIX SIGMA E ORÇAMENTO DE ERRO',
            negrito=True)
        por(es, RES_R0, 5, '(conta analito × nível avaliado; usa os '
                           'mesmos filtros de período e de EP da tabela acima)')
        por(es, RES_R0 + 1, 1, 'Classificação Sigma', negrito=True)
        por(es, RES_R0 + 1, 2, 'n', negrito=True)
        por(es, RES_R0 + 1, 4, 'Orçamento de erro (ET vs ETp)', negrito=True)
        por(es, RES_R0 + 1, 6, 'n', negrito=True)
        for i, cls in enumerate(CLASSES):
            por(es, RES_R0 + 2 + i, 1, cls)
            por(es, RES_R0 + 2 + i, 2,
                '=COUNTIF($M$14:$M$93,$A{0})'.format(RES_R0 + 2 + i), formula=True)
        for i, mg in enumerate(MARGENS):
            por(es, RES_R0 + 2 + i, 4, mg)
            por(es, RES_R0 + 2 + i, 6,
                '=COUNTIF($P$14:$P$93,$D{0})'.format(RES_R0 + 2 + i), formula=True)
        por(es, RES_R0 + 7, 1, 'Sem EP (sem bias → sem Sigma)')
        por(es, RES_R0 + 7, 2, '=COUNTIF($G$14:$G$93,"SEM EP")', formula=True)
        print('Estatistica L%d: resumo (5 classes de sigma, 3 de margem)' % RES_R0)

        # ---- 4. Estatistica: lista dos criticos -------------------------
        por(es, LST_R0, 1, 'ANALITOS EM MARGEM CRÍTICA OU COM ETp EXCEDIDO',
            negrito=True)
        for c, t in ((1, 'Analito'), (2, 'Nível'), (3, 'ET %'), (4, 'ETp %'),
                     (5, 'Margem %'), (6, 'Situação'), (7, 'Sigma'),
                     (8, 'Status sigma')):
            por(es, LST_R0 + 1, c, t, negrito=True)
        for i in range(LST_N):
            lin = LST_R0 + 2 + i
            base = 'MATCH({0},$U$14:$U$93,0)'.format(i + 1)
            for c, orig in ((1, '$A'), (2, '$B'), (3, '$H'), (4, '$K'),
                            (5, '$O'), (6, '$P'), (7, '$L'), (8, '$M')):
                por(es, lin, c,
                    '=IFERROR(INDEX({0}$14:{0}$93,{1}),"")'.format(orig, base),
                    formula=True)
            for c in (3, 4, 5, 7):
                por(es, lin, c, None, fmt='0.00')
        print('Estatistica L%d: lista com %d vagas (INDEX/MATCH, sem CSE)'
              % (LST_R0, LST_N))

        # ---- 5. Painel: bloco descritivo perde o que nao e descritivo ---
        tenta(lambda: pa.Range('F6:J8').ClearContents())
        por(pa, 5, 1, 'INDICADORES POR NÍVEL — descritivos', negrito=True)

        # ---- 6. Painel: Westgard ganha o veredito da corrida ------------
        por(pa, 5, 12, 'WESTGARD — violações no período '
                       '(o que reprova corrida)', negrito=True)
        por(pa, 6, 19, 'Status', negrito=True)
        for lin in (7, 8):
            por(pa, lin, 19,
                '=IF($B{0}="","",IF($R{0}>0,"REPROVA — "&$R{0}&'
                '" violação(ões)","Sem violação"))'.format(lin),
                formula=True)
        por(pa, 9, 12, 'Só 1-3s, 2-2s e R-4s estão implementadas; '
                       '4-1s e 8x são reservas e permanecem em zero.')

        # ---- 7. Painel: bloco SIX SIGMA ---------------------------------
        bias = 'mCEQ.BiasEQ(selAnalito,eqAnoEP,"ABS",eqProvedor,eqRodada)'
        etp = 'IFERROR(INDEX(etpOficial,MATCH(selAnalito,{0},0)),"")'.format(MAP)
        por(pa, 10, 10, 'SIX SIGMA — desempenho analítico do método',
            negrito=True)
        for i, t in enumerate(['Nível', 'CV % obs', 'Bias EQC (abs) %',
                               'ETp %', 'Sigma', 'Classificação']):
            por(pa, 11, 10 + i, t, negrito=True)
        for lin, orig in ((12, 7), (13, 8)):
            por(pa, lin, 10, '=$A{0}'.format(orig), formula=True)
            por(pa, lin, 11, '=$E{0}'.format(orig), formula=True, fmt='0.00')
            por(pa, lin, 12, '=IF(selAnalito="","",{0})'.format(bias),
                formula=True, fmt='0.00')
            por(pa, lin, 13, '=IF(selAnalito="","",{0})'.format(etp),
                formula=True, fmt='0.00')
            por(pa, lin, 14,
                '=IF(OR(NOT(ISNUMBER($K{0})),$K{0}=0,NOT(ISNUMBER($L{0})),'
                'NOT(ISNUMBER($M{0}))),"",($M{0}-ABS($L{0}))/$K{0})'.format(lin),
                formula=True, fmt='0.00')
            por(pa, lin, 15, '=IF(NOT(ISNUMBER($N{0})),"",{1})'
                .format(lin, escada_sigma('$N%d' % lin)), formula=True)
        por(pa, 14, 10, 'Sigma baixo não reprova corrida: exige mais '
                        'regras, mais controles e CQ mais frequente.')

        # ---- 8. Painel: bloco ERRO TOTAL vs ETp -------------------------
        por(pa, 16, 10, 'ERRO TOTAL vs ETp — orçamento de erro',
            negrito=True)
        for i, t in enumerate(['Nível', 'ET %', 'ETp %', 'Margem (p.p.)',
                               'Margem %', 'Situação']):
            por(pa, 17, 10 + i, t, negrito=True)
        for lin, orig in ((18, 12), (19, 13)):
            por(pa, lin, 10, '=$J{0}'.format(orig), formula=True)
            por(pa, lin, 11,
                '=IF(OR(NOT(ISNUMBER($K{0})),NOT(ISNUMBER($L{0}))),"",'
                '1.65*$K{0}+ABS($L{0}))'.format(orig), formula=True, fmt='0.00')
            por(pa, lin, 12, '=$M{0}'.format(orig), formula=True, fmt='0.00')
            por(pa, lin, 13,
                '=IF(OR(NOT(ISNUMBER($K{0})),NOT(ISNUMBER($L{0}))),"",'
                '$L{0}-$K{0})'.format(lin), formula=True, fmt='0.00')
            por(pa, lin, 14,
                '=IF(OR(NOT(ISNUMBER($M{0})),NOT(ISNUMBER($L{0})),$L{0}=0),"",'
                '$M{0}/$L{0}*100)'.format(lin), formula=True, fmt='0.00')
            por(pa, lin, 15, '=IF(NOT(ISNUMBER($N{0})),"",{1})'
                .format(lin, escada_margem('$N%d' % lin)), formula=True)

        # ---- 9. Painel: margem critica no conjunto ----------------------
        por(pa, 21, 10, 'MARGEM CRÍTICA — todos os analitos', negrito=True)
        for i, (rot, cnt) in enumerate((
                ('ETp excedido', 'ETp excedido'),
                ('Margem crítica (≤10%)', 'Margem critica'),
                ('Dentro do orçamento', 'Dentro do orcamento'))):
            por(pa, 22 + i, 10, rot)
            por(pa, 22 + i, 12,
                '=COUNTIF(Estatística!$P$14:$P$93,"{0}")'.format(cnt),
                formula=True)
        por(pa, 25, 10, 'Os blocos acima usam o período do Painel; a '
                        'Estatística usa o período dela. Filtros '
                        'diferentes, números diferentes — por desenho.')
        # colunas estreitas renderizam "####" e escondem o numero do gestor
        tenta(lambda: pa.Range('J:O').__setattr__('ColumnWidth', 15.5))
        tenta(lambda: es.Range('G:T').__setattr__('ColumnWidth', 15.5))
        print('Painel: descritivo A5, Westgard L5 (+Status S), '
              'Six Sigma J10, ET vs ETp J16, margem critica J21')

        # ---- 10. conferencia -------------------------------------------
        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())
        erros = []
        for r in range(EST_R0, EST_RN + 1):
            for c in range(7, 21):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        for r in list(range(11, 20)) + [22, 23, 24]:
            for c in range(10, 16):
                t = str(tenta(lambda rr=r, cc=c: pa.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Painel!%d,%d=%s' % (r, c, t))
        for r in range(LST_R0 + 2, LST_R0 + 2 + LST_N):
            for c in range(1, 9):
                t = str(tenta(lambda rr=r, cc=c: es.Cells(rr, cc).Text))
                if eh_erro(t):
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        print('celulas em erro: %d' % len(erros))
        for e in erros[:8]:
            print('   %s' % e)
        if erros:
            raise SystemExit('erro de formula -- nada salvo')

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
