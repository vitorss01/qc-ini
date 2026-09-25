# -*- coding: utf-8 -*-
"""portar_hema.py -- leva a Hematologia ao mesmo motor da Bioquimica (ADR-049/050/052).

Uso:  python portar_hema.py <QC_Hematologia.xlsm> <referencia_build_h1.xlsm> <QC_Bioquimica_instalada.xlsm>

POR QUE
  O arquivo de producao da Hematologia rodou o mEstatistica ANTIGO, que gravava
  direto nas abas de interface. Resultado, medido em 19/09/2026:
    Calc ........ 12.614 formulas viraram valor (grafico congelado em agosto)
    Painel ...... 45 formulas viraram valor (indicadores congelados)
    Estatistica . 1.320 formulas viraram valor, e um script posterior aplicou
                  por cima o layout da Bioquimica pela metade (cabecalho no
                  meio dos dados, coluna K com bias onde o titulo dizia ET)
  E ela nao tinha nada do multi-lote nem da capacidade de 60 meses: banco com
  teto fixo de 15.000 linhas e flags BA:BC em formula O(n^2).

O QUE FAZ
  VBA     mesmo codigo-fonte da Bioquimica (src/), com NLV=3 e as celulas de
          filtro da Estatistica da Hematologia; mBanco com teto de 200.000
          linhas; Painel com seletor de lote em H3.
  Banco   BA:BC viram valor (AtualizarFlagsBanco); nomes r* acompanham o dado.
  Lotes   cadastro ate 100 lotes; LotesStore por chave LOTE|ANALITO; lote em
          analise no Painel (H3); Analitos!E:J edita o lote em tela.
  Calc    formulas reconstruidas a partir da referencia (build_h1, a ultima
          versao integra) e ja no desenho da Etapa 1 + ADR-050 (le o motor).
  Painel  indicadores de volta (engPainel), faixa de status do lote, secao
          Sigma apontando para a Estatistica nova (3 niveis).
  Estat.  tabela no desenho da Bioquimica, 3 niveis (linhas 14..133).

Nao salva se qualquer conferencia falhar.
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlh  # noqa: E402
import v2   # noqa: E402
import ux   # noqa: E402  camada de experiencia (ADR-051)
import painel  # noqa: E402  desenho do Painel (ADR-054)

AQUI = os.path.dirname(os.path.abspath(__file__))
SENHA = 'qcini2025'
CAP_HEMA = 200000        # 28 analitos x 3 niveis x 1 corrida/dia x 365 x 5 anos = 153.300 (+30%)

# ---- VBA: (componente, arquivo em src/, substituicoes de produto) ----
SUBST_ESTAT = [
    ('Public Const NLV As Long = 2          \' niveis deste setor',
     'Public Const NLV As Long = 3          \' niveis deste setor (Hematologia)'),
    ('''    Dim visao As String
    visao = UCase$(Trim$(CStr(ws.Range("B4").Value)))
    loteF = Trim$(CStr(ws.Range("H4").Value))
    If visao = "PERSONALIZADO" Then
        Dim dIniPer As Variant, dFimPer As Variant
        dIniPer = ws.Range("B5").Value
        dFimPer = ws.Range("D5").Value
        If IsDate(dIniPer) Then anoDe = Year(CDate(dIniPer)) Else anoDe = 0
        If IsDate(dFimPer) Then anoAte = Year(CDate(dFimPer)) Else anoAte = anoDe
    Else
        anoDe = CLng(Val(ws.Range("D4").Value))
        anoAte = anoDe
    End If''',
     '''    ' Hematologia: Ano De (B3), Ano Ate (D3), Lote (B4 = lote em analise)
    anoDe = CLng(Val(ws.Range("B3").Value))
    anoAte = CLng(Val(ws.Range("D3").Value))
    loteF = Trim$(CStr(ws.Range("B4").Value))'''),
]
MODULOS = [
    ('mApp', 'mApp.bas', None),
    ('mPeriodo', 'mPeriodo.bas', None),   # ADR-054: a Hematologia nao tem mEstatPeriodo
    ('mBanco', 'mBanco.bas', [('Public Const CAP_LINHAS As Long = 150000',
                               f'Public Const CAP_LINHAS As Long = {CAP_HEMA}   \' Hematologia: 28 x 3 niveis x 365 x 5 anos = 153.300')]),
    ('mDados', 'mDados.bas', None),
    ('mLotes', 'mLotes.bas', None),
    ('mEstatistica', 'mEstatistica.bas', SUBST_ESTAT),
    ('mUI', 'hema/mUI.bas', None),
    ('mOperacao', 'hema/mOperacao.bas', None),
    ('mEntrada', 'hema/mEntrada.bas', None),
    ('mCEQ', 'mCEQ.bas', None),
    ('mBI', 'mBI.bas', None),
    ('mSeguranca', 'mSeguranca.bas', None),
    ('Planilha7', 'hema/Planilha7.cls', None),
    ('Planilha3', 'Planilha3.cls', None),
    ('EstaPastaDeTrabalho', 'EstaPastaDeTrabalho.cls', None),
]

NLV = 3
EST_ULT = 133            # ultima linha da tabela da aba Estatistica
LINHA_BLOCO = 42         # 1a linha do bloco de desempenho (abaixo dos 3 graficos)
VAL = {1: 'F', 2: 'AB', 3: 'AX'}            # Calc: valor do nivel
PAR = {1: ('BT', 'BU'), 2: ('BV', 'BW'), 3: ('BX', 'BY')}   # Calc linha 1: media/DP
EIXO = {'BZ': ('AA', 1, '-'), 'CA': ('AB', 1, '+'), 'CB': ('AE', 2, '-'),
        'CC': ('AF', 2, '+'), 'CD': ('AI', 3, '-'), 'CE': ('AJ', 3, '+')}
REP = {1: ('X', 'Y', 'Z'), 2: ('AT', 'AU', 'AV'), 3: ('BP', 'BQ', 'BR')}
REJ = {1: ('J', '$P3'), 2: ('AF', '$AL3'), 3: ('BB', '$BH3')}


def desproteger(ws):
    estava = bool(ws.ProtectContents)
    if estava:
        ws.Unprotect(SENHA)
    return estava


def main(caminho, ref_h1, ref_bio):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    wb = ex.abrir(caminho)
    ref = ex.xl.Workbooks.Open(os.path.abspath(ref_h1), 0, True)       # somente leitura
    bio = ex.xl.Workbooks.Open(os.path.abspath(ref_bio), 0, True)
    ok = False
    try:
        ex.xl.Calculation = -4135
        sh = {s.Name: s for s in wb.Worksheets}
        prot = {n: desproteger(s) for n, s in sh.items()}
        cfg, pai, calc = sh['Configuração'], sh['Painel'], sh['Calc']
        ls, ana, est, eng = sh['LotesStore'], sh['Analitos'], sh['Estatística'], sh['Eng_Saida']
        db = sh['DB_Resultados']

        # ================= VBA =================
        vbp = wb.VBProject
        n = v2.conferir_duplicados([(arq, subst) for _, arq, subst in MODULOS if arq.endswith('.bas')])
        print(f'VBA: {n} nomes publicos, nenhum duplicado')
        for comp, arq, subst in MODULOS:
            v2.substituir_codigo(vbp, comp, v2.codigo(arq, subst))
        print('VBA: modulos instalados (motor unico, NLV=3)')

        # ================= Banco: BA:BC viram valor =================
        ult_db = db.Cells(db.Rows.Count, 1).End(-4162).Row
        f_ba = db.Range('BA4').Formula
        if isinstance(f_ba, str) and f_ba.startswith('='):
            db.Range('BA4:BC15003').ClearContents()
        print(f'Banco: {ult_db - 3} linhas; formulas BA:BC removidas')

        # ================= LotesStore O/P =================
        ls.Range('N1').Copy(ls.Range('O1:P1'))
        ls.Range('O1').Value = 'Analito'
        ls.Range('P1').Value = 'Chave'
        ult = ls.Cells(ls.Rows.Count, 1).End(-4162).Row
        nomes_an = [ana.Cells(r, 1).Value for r in range(4, 44)]
        n_ch = 0
        for r in range(2, ult + 1):
            lote, idx = ls.Cells(r, 1).Value, ls.Cells(r, 2).Value
            if lote in (None, '') or idx in (None, ''):
                continue
            nm = nomes_an[int(idx) - 1] if 1 <= int(idx) <= 40 else None
            if nm in (None, ''):
                continue
            lt = str(int(lote)) if isinstance(lote, float) and lote == int(lote) else str(lote).strip()
            ls.Cells(r, 15).Value = str(nm).strip()
            ls.Cells(r, 16).Value = f'{lt.upper()}|{str(nm).strip().upper()}'
            n_ch += 1
        print(f'LotesStore: {n_ch} linhas com chave')

        # ================= Configuracao: 100 lotes, loteParam, lista =================
        cfg.Range('F1').Copy(cfg.Range('F2'))
        cfg.Range('G1').Copy(cfg.Range('G2'))
        cfg.Range('F2').Value = 'loteParam→'
        cfg.Range('G2').Value = cfg.Range('G1').Value
        if cfg.Range('B76').Value in (None, ''):
            cfg.Range('B75:D75').Copy(cfg.Range('B76:D125'))
            for i in range(51, 101):
                cfg.Cells(25 + i, 2).Value = f'Lote {i:02d}'
            cfg.Range('C76:D125').ClearContents()
        cfg.Range('B23').Value = str(cfg.Range('B23').Value or '').replace('até 50', 'até 100')
        cfg.Range('C21').Formula = '=IFERROR(INDEX(Configuração!$D$26:$D$125,MATCH($C$20,Configuração!$C$26:$C$125,0)),"")'
        cfg.Range('AB1').Value = 'lstLotes (auto - nao editar)'
        cfg.Range('AB2:AB101').Formula = (
            '=IFERROR(INDEX($C$26:$C$125,AGGREGATE(15,6,(ROW($C$26:$C$125)-ROW($C$26)+1)'
            '/($C$26:$C$125<>""),ROWS($AB$2:AB2))),"")')
        cfg.Range('AB1:AB101').Font.Color = 8421504
        for rg in (cfg.Range('G1:G2'), cfg.Range('C20'), sh['Eventos_Westgard'].Range('L2')):
            rg.NumberFormat = '@'

        # ================= Nomes =================
        refs = {
            'loteSel': '=Painel!$H$3',
            'loteAnalise': '=SE(Painel!$H$3="";Configuração!$C$20;Painel!$H$3)',
            'loteParam': '=Configuração!$G$2',
            'filtroDe': '=Painel!$G$3', 'filtroAte': '=Painel!$G$4',
            'filtroAno': '=Painel!$K$3', 'filtroPeriodo': '=Painel!$K$4',
            'lstAnosCIQ': '=Configuração!$Z$2:ÍNDICE(Configuração!$Z$2:$Z$50;MÁXIMO(1;CONT.VALORES(Configuração!$Z$2:$Z$50)))',
            'lsChave': '=LotesStore!$P$2:$P$4001',
            'lstLotes': '=Configuração!$AB$2:ÍNDICE(Configuração!$AB$2:$AB$101;MÁXIMO(1;CONT.SE(Configuração!$AB$2:$AB$101;"?*")'
                        '+CONT.NÚM(Configuração!$AB$2:$AB$101)))',
            'lstAnalitos': '=Analitos!$A$4:ÍNDICE(Analitos!$A$4:$A$43;MÁXIMO(1;CONT.VALORES(Analitos!$A$4:$A$43)))',
            'regLoteCol': '=Configuração!$C$26:$C$125',
            'aInput': '=Analitos!$E$4:$J$43',
            'protegerEstrutura': '=VERDADEIRO',   # a Hematologia chegou com a estrutura protegida (ADR-052)
            'engNTotal': '=Eng_Saida!$K$1', 'engCortadas': '=Eng_Saida!$M$1',
            'engLotesPer': '=Eng_Saida!$O$1', 'engParam': '=Eng_Saida!$Q$1',
        }
        for t, (cm, cs) in enumerate((('C', 'D'), ('E', 'F'), ('G', 'H')), start=1):
            refs[f'lsMedN{t}'] = f'=LotesStore!${cm}$2:${cm}$4001'
            refs[f'lsDPN{t}'] = f'=LotesStore!${cs}$2:${cs}$4001'
        for n, r in refs.items():
            v2.nome(wb, n, r)
        v = cfg.Range('C20').Validation
        v.Delete()
        v.Add(3, 1, 1, '=lstLotes')
        for c, t in (('J1', 'nTotal:'), ('L1', 'cortadas:'), ('N1', 'lotes no período:'), ('P1', 'param:')):
            eng.Range(c).Value = t
        print('Nomes, cadastro e validacao: ok')

        # ================= Calc: reconstrucao a partir da referencia =================
        rc = ref.Sheets('Calc')
        for c in range(2, 72):                          # B..BS, linha 3 (relativa) -> 3:182
            f = rc.Cells(3, c).Formula
            if isinstance(f, str) and f.startswith('='):
                calc.Range(calc.Cells(3, c), calc.Cells(182, c)).Formula = f
        calc.Range('A3:A182').Value = [[i] for i in range(1, 181)]
        for c in range(72, 86):                         # BT1..CG1
            f = rc.Cells(1, c).Formula
            if isinstance(f, str) and f.startswith('='):
                calc.Cells(1, c).Formula = f
        # chave do lote em analise + casas decimais
        calc.Range('CH1').Formula = ('=IFERROR(MATCH(UPPER(TRIM(loteAnalise&""))&"|"&UPPER(TRIM(selAnalito)),'
                                     'lsChave,0),0)')
        calc.Range('CI1').Formula = ('=IFERROR(INDEX(Analitos!$D$4:$D$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))'
                                     ',1)+1')
        for t, (cm, cs) in PAR.items():
            m, s = f'lsMedN{t}', f'lsDPN{t}'
            base = (f'=IF($CH$1=0,"",IF(AND(ISNUMBER(INDEX({m},$CH$1)),ISNUMBER(INDEX({s},$CH$1))),'
                    f'IF(INDEX({s},$CH$1)>0,INDEX(%s,$CH$1),""),""))')
            calc.Range(f'{cm}1').Formula = base % m
            calc.Range(f'{cs}1').Formula = base % s
        for cel, (colA, t, sinal) in EIXO.items():
            f = calc.Range(f'{cel}1').Formula
            velho = f'INDEX(Analitos!${colA}$4:${colA}$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))'
            if velho not in f:
                raise SystemExit(f'Calc!{cel}1: ancora do eixo nao encontrada: {f[:120]}')
            cm, cs = PAR[t]
            calc.Range(f'{cel}1').Formula = f.replace(velho, f'(${cm}$1{sinal}3.3*${cs}$1)')
        # corridas = as do motor
        calc.Range('B3:B182').Formula = (
            '=IF(OR(selAnalito="",(""&engAnalito)<>(""&selAnalito),(""&engLote)<>(""&loteAnalise),'
            '$A3>N(engNRun)),"",INDEX(engRUN,$A3))')
        # lote em analise no lugar do lote ativo (C e valores: v2 troca pelo motor logo abaixo)
        for c in ['C'] + [VAL[t] for t in range(1, NLV + 1)]:
            f = calc.Range(f'{c}3').Formula
            if 'loteAtivo' not in f:
                raise SystemExit(f'Calc!{c}3 sem loteAtivo: {f[:100]}')
            calc.Range(f'{c}3:{c}182').Formula = f.replace('loteAtivo', 'loteAnalise')
        for t, cols in REP.items():
            for c in cols:
                f = calc.Range(f'{c}3').Formula
                if not f.startswith('=IF(AND($D3=1,'):
                    raise SystemExit(f'Calc!{c}3 formato inesperado: {f[:80]}')
                calc.Range(f'{c}3:{c}182').Formula = f.replace(
                    '=IF(AND($D3=1,', '=IF(AND($D3=1,(""&loteAnalise)=(""&loteCarregado),', 1)
        for t, (c, st) in REJ.items():
            f = calc.Range(f'{c}3').Formula
            velho = f'{st}<>"OK",{st}<>""'
            if velho not in f:
                raise SystemExit(f'Calc!{c}3 formato inesperado: {f[:100]}')
            calc.Range(f'{c}3:{c}182').Formula = f.replace(velho, f'{st}="REJEITADO"')
        f = calc.Range('D3').Formula
        if f.count('$C3>=filtroDe') != 1 or f.count('$C3<=filtroAte') != 1:
            raise SystemExit(f'Calc!D3 formato inesperado: {f[:120]}')
        calc.Range('D3:D182').Formula = painel.sem_trimestre(
            f.replace('$C3>=filtroDe', 'INT($C3)>=INT(filtroDe)').replace(
                '$C3<=filtroAte', 'INT($C3)<=INT(filtroAte)'))     # ADR-054
        print('Calc: formulas reconstruidas (Etapa 1)')

        # ================= Painel (ADR-054: mesmo desenho da Bioquimica) =================
        faltam = ('IF(AND(Calc!$BT$1="",Calc!$BV$1="",Calc!$BX$1=""),"N1, N2 e N3",'
                  'TRIM(IF(Calc!$BT$1="","N1 ","")&IF(Calc!$BV$1="","N2 ","")&IF(Calc!$BX$1="","N3","")))')
        faixa_f = (
            '=IF(OR(Calc!$BT$1="",Calc!$BV$1="",Calc!$BX$1=""),'
            '"\u26d4 LOTE "&loteAnalise&" \u00b7 "&selAnalito&": SEM M\u00c9DIA/DP em "&' + faltam +
            '&" \u2014 cadastre na aba Analitos. ",'
            '"Lote "&loteAnalise&" \u00b7 alvo N1 "&FIXED(Calc!$BT$1,Calc!$CI$1)&" \u00b1 "&FIXED(Calc!$BU$1,Calc!$CI$1+1)'
            '&" \u00b7 N2 "&FIXED(Calc!$BV$1,Calc!$CI$1)&" \u00b1 "&FIXED(Calc!$BW$1,Calc!$CI$1+1)'
            '&" \u00b7 N3 "&FIXED(Calc!$BX$1,Calc!$CI$1)&" \u00b1 "&FIXED(Calc!$BY$1,Calc!$CI$1+1)&". ")'
            '&IF(AND(COUNTIF(Calc!$D$3:$D$182,1)=0,selAnalito<>""),"Nenhuma corrida deste lote no per\u00edodo"'
            '&IF(engLotesPer<>""," (com dados: "&engLotesPer&")","")&". ","")'
            '&IF(N(engCortadas)>0,"\u26a0 "&engCortadas&" corrida(s) antigas do per\u00edodo fora do gr\u00e1fico (m\u00e1x. 180). ","")'
            '&IF(OR($G$3<>"",$G$4<>""),"Filtro "&IF($G$3<>"",TEXT($G$3,"dd/mm/aa"),"in\u00edcio")&" a "'
            '&IF($G$4<>"",TEXT($G$4,"dd/mm/aa"),"hoje"),"Sem filtro de data")')
        # ADR-055: o realce "esta regra e a recomendada pelo Sigma" existia so na
        # Bioquimica -- e la estava MORTO desde o ADR-054 (lia Painel!V10:V11,
        # celulas que o Painel novo esvaziou). Agora le a coluna Sigma do Painel,
        # e existe nos dois produtos. Vem ANTES do padronizar: e ele quem usa os
        # nomes para a formatacao condicional do cabecalho de Westgard.
        painel.realce_regras(wb, NLV, [('1-3S', '1_3s'), ('3-1S', '3_1s'), ('6X', '6x'),
                                       ('2of3-2S', '2of3_2s'), ('R4S', 'R_4s')])
        print('ADR-055: matriz de regras do Sigma em Cfg_PlanoQC (novo na Hematologia)')
        painel.padronizar(wb, 'Hematologia', NLV, EST_ULT,
                          ['BT', 'BV', 'BX'],
                          [('1-3S', 13), ('3-1S', 16), ('6X', 17), ('2of3-2S', 14), ('R4S', 15)],
                          'O3:U4', faixa_f, LINHA_BLOCO)
        print('Painel: desenho padrao (topo, filtro Ano/Periodo, bloco de desempenho na coluna A)')

        # ================= Estatistica: tabela no desenho da Bioquimica, 3 niveis =================
        estatistica(est, bio.Sheets('Estatística'))
        print('Estatistica: tabela reconstruida (linhas 14..133)')
        # As referencias saem ANTES de qualquer macro: Application.Run sem o nome
        # da pasta resolve na pasta ativa -- e a Bioquimica tem os mesmos modulos.
        ex.xl.CutCopyMode = False
        ref.Close(False)
        bio.Close(False)
        ref = bio = None
        wb.Activate()

        # ================= Analitos =================
        ana.Range('A2').Formula = (
            '="PARÂMETROS DO LOTE "&loteParam&"  ·  Média/DP valem só para este lote — para ver ou editar '
            'outro lote, selecione-o no Painel (H3).  ·  Lote em uso p/ lançamentos: "&loteAtivo&'
            '"  ·  Especificações (CLIA/VB/ETp) valem para todos os lotes."')

        # ================= ADR-050: Calc le o motor; Liberacao por valor =================
        v2.aplicar_planilhas(ex, wb, NLV)
        print('ADR-050: Calc lendo o motor, Liberacao A:B por valor')
        col_nav = painel.coluna_da_barra(pai)
        print(f'ADR-055: barra de navegacao comeca em {col_nav}1 (depois do titulo + 1 coluna)')
        ux.aplicar(wb, 'Hematologia', 'H3', 'O3:U4', CAP_HEMA, col_nav + '1')
        print('ADR-051: navegacao, Inicio de aplicativo, status do lote colorido, Novo lote')

        # ================= roda =================
        r = ex.run('mLotes.ConferirLotesStore', teto=60)
        print('ConferirLotesStore:', r)
        if not str(r).startswith('ok'):
            raise SystemExit('LotesStore inconsistente')
        t0 = time.time()
        ex.run('mBanco.AtualizarFlagsBanco', teto=600)
        print(f'AtualizarFlagsBanco: {time.time() - t0:.1f}s; rRUN = {wb.Names("rRUN").RefersTo}')
        ex.xl.Calculation = -4105
        t0 = time.time(); ex.xl.CalculateFull(); ex.esperar()
        print(f'CalculateFull: {time.time() - t0:.1f}s')
        ex.run('mLotes.SincronizarLotesAoAbrir', teto=60)
        ex.run('mEstatistica.InvalidarCache', teto=60)
        t0 = time.time()
        ex.run('mEstatistica.AtualizarEstatistica', teto=600)
        print(f'AtualizarEstatistica: {time.time() - t0:.1f}s')
        ex.xl.CalculateFull(); ex.esperar()
        ex.run('AtualizarEixos', teto=60)

        chk = {
            'loteSel': pai.Range('H3').Value, 'loteParam': cfg.Range('G2').Value,
            'BT1': calc.Range('BT1').Value, 'BU1': calc.Range('BU1').Value, 'BX1': calc.Range('BX1').Value,
            'engLote': eng.Range('E1').Value, 'engNRun': eng.Range('I1').Value,
            'CalcB3': calc.Range('B3').Value, 'CalcF3': calc.Range('F3').Value, 'CalcAX3': calc.Range('AX3').Value,
            'PainelB7': pai.Range('B7').Value, 'PainelC9': pai.Range('C9').Value, 'O3': pai.Range('O3').Value,
            'EstC14': est.Range('C14').Value, 'EstD16': est.Range('D16').Value, 'EstL14': est.Range('L14').Value,
            'erros_formula': contar_erros(wb),
            'filtroCalc': ex.xl.WorksheetFunction.CountIf(calc.Range('D3:D182'), 1),
        }
        for k, val in chk.items():
            print(f'  {k:14s} = {val}')
        print('  ano do filtro =', painel.garantir_ano(wb))
        # Teste de fumaca: o VBA compila PROCEDIMENTO a procedimento, sob demanda.
        # Um erro de compilacao numa rotina que ninguem chamou durante a instalacao
        # so aparece no primeiro clique do usuario -- e ai a caixa de erro trava o
        # arquivo (foi o que aconteceu na Hematologia com mEstatPeriodo).
        for macro, args in (('mApp.Ocupado', ()), ('mUI.AjustarGraficos', ()),
                            ('mUI.AjustarLegendas', ()),
                            ('mPeriodo.PeriodoInicio', ('TRIMESTRAL', 2026, 2)),
                            ('mApp.AplicarFiltroAnoMes', ())):
            r = ex.run(macro, *args, teto=60)
            print(f'  fumaca {macro} -> {r}')
            if macro == 'mApp.AplicarFiltroAnoMes' and not str(r).startswith('OK'):
                raise SystemExit(f'filtro de periodo nao respondeu: {r}')
        painel.conferir_sem_qsel(wb)
        if (chk['BT1'] in (None, '') or chk['CalcB3'] in (None, '') or chk['PainelB7'] in (None, '')
                or chk['erros_formula'] or not chk['EstC14'] or not chk['filtroCalc']):
            raise SystemExit('Calc/Painel sem parametro ou sem corrida apos portar -- nao salvo')

        for n, s in sh.items():
            if prot[n]:
                s.Protect(SENHA, False, True, True, True)
        wb.Save()
        ok = True
        print('SALVO:', caminho)
    finally:
        for x in (ref, bio):
            try:
                if x is not None:
                    x.Close(False)
            except Exception:
                pass
        ex.fechar(salvar=False)
        if not ok:
            import traceback
            traceback.print_exc()
            print('*** NAO SALVO ***')
            sys.exit(1)


def estatistica(est, bio_est):
    """Tabela da Estatistica no desenho da Bioquimica (cabecalho na linha 13), 3 niveis."""
    # limpa a tabela antiga (linhas 6..140), mantendo o bloco de controle 1..5
    est.Range('A6:AF140').UnMerge()
    est.Range('A6:AF140').Clear()
    # formato: copia cabecalho e uma linha de dado da Bioquimica
    bio_est.Range('A13:AD14').Copy()
    est.Range('A13').PasteSpecial(-4122)          # xlPasteFormats
    bio_est.Range('A14:AD14').Copy()
    est.Range('A15:AD133').PasteSpecial(-4122)
    est.Application.CutCopyMode = False
    for c in range(1, 31):
        est.Columns(c).ColumnWidth = bio_est.Columns(c).ColumnWidth
    rot = ['Analito', 'Nível', 'n', 'Média', 'DP', 'CV %', 'Bias EQC (abs) %', 'ET %', 'ESPQ FONTE', 'CVTp %',
           'ETp %', 'SIX SIGMA', 'Status sigma', 'Margem ETp (p.p.)', 'Margem ETp %', 'Status margem',
           'Status CV', 'Status SDI (EP)', 'Status limites (EP)', 'Bias EQC (sinal) %', 'ordem crítico',
           'DPM teórico', 'Rendimento teórico %', 'Regras Westgard recomendadas', 'N (medições de controle)',
           'Run Size máx (pacientes)', 'Cobertura do motor Westgard', 'chave analito|nível',
           'N EQA Resultados', 'N EQA Rodadas']
    est.Range('A13:AD13').Value = [rot]
    m = 'MATCH($A14,Analitos!$A$4:$A$43,0)'
    f = {
        'C': '=IF($A14="","",IF(INDEX(engEstat,ROW()-13,3)="","",INDEX(engEstat,ROW()-13,3)))',
        'D': '=IF($A14="","",IF(INDEX(engEstat,ROW()-13,4)="","",INDEX(engEstat,ROW()-13,4)))',
        'E': '=IF($A14="","",IF(INDEX(engEstat,ROW()-13,5)="","",INDEX(engEstat,ROW()-13,5)))',
        'F': '=IF($A14="","",IF(INDEX(engEstat,ROW()-13,6)="","",INDEX(engEstat,ROW()-13,6)))',
        'G': '=IF($A14="","",mCEQ.BiasEQ($A14,eqAnoEP,"ABS",eqProvedor,eqRodada))',
        'H': '=IF(OR(NOT(ISNUMBER($F14)),NOT(ISNUMBER($G14))),"",1.65*$F14+ABS($G14))',
        'I': f'=IF($A14="","",IFERROR(INDEX(Analitos!$Q$4:$Q$43,{m}),""))',
        'J': (f'=IF($A14="","",IFERROR(IF($I14="VB",INDEX(Analitos!$T$4:$T$43,{m}),'
              f'IF(INDEX(Analitos!$S$4:$S$43,{m})<>"",INDEX(Analitos!$S$4:$S$43,{m}),INDEX(Analitos!$T$4:$T$43,{m}))),""))'),
        'K': f'=IF($A14="","",IFERROR(INDEX(Analitos!$R$4:$R$43,{m}),""))',
    }
    # L..AD: iguais aos da Bioquimica (so referenciam a propria linha e as UDFs)
    for c in range(12, 31):
        col = est.Cells(1, c).Address.split('$')[1]
        fb = bio_est.Cells(14, c).Formula
        f[col] = fb
    for col, ff in f.items():
        est.Range(f'{col}14:{col}133').Formula = ff
    # A/B: analito e nivel (3 linhas por analito)
    a = [[f'=Analitos!$A${4 + k // 3}', k % 3 + 1] for k in range(120)]
    est.Range('A14:B133').Formula = a
    # filtros: lote = lote em analise do Painel
    est.Range('A4').Value = 'Lote'
    est.Range('B4').MergeArea.NumberFormatLocal = 'Geral'   # estava como TEXTO: a formula entrava como texto
    est.Range('B4').Formula = '=loteAnalise'
    est.Range('C4').Value = '(refletido do Painel)'
    est.Range('F4').Value = ('Ano De/Até filtram um intervalo (ex.: 3 ou 5 anos). O lote é o escolhido no Painel '
                             '(H3): média, DP e CV nunca misturam lotes.')
    est.Range('A12').Value = 'Bias, ET e Sigma pelo controle EXTERNO (EP), como na Bioquímica; n/Média/DP/CV pelo motor.'
    est.Range('A12').Font.Italic = True
    est.Range('A12').Font.Size = 9


def contar_erros(wb):
    """Formulas com erro no Painel e na Estatistica, e formulas gravadas como TEXTO
    (celula com formato @). O Calc fica de fora: #N/D ali e proposital -- e o
    que faz o grafico nao desenhar o ponto."""
    n = 0
    for nm in ('Painel', 'Estatística', 'Configuração', 'Analitos'):
        ws = wb.Sheets(nm)
        try:
            n += ws.UsedRange.SpecialCells(-4123, 16).Count
        except Exception:
            pass
        for linha in ws.UsedRange.Value or ():
            for v in linha:
                if isinstance(v, str) and v.startswith('=') and len(v) > 2:
                    n += 1
    return n


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
