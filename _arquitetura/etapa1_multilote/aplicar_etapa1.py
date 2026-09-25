# -*- coding: utf-8 -*-
"""aplicar_etapa1.py -- instala a Etapa 1 (multi-lote, ADR-049) num QC_Bioquimica.xlsm.

Uso:  python aplicar_etapa1.py <arquivo.xlsm>

O que faz (tudo por nome/rotulo, e conferido no fim):
  VBA     mLotes (novo), mEstatistica, mEstatPeriodo, mBI, Painel (Planilha7),
          Analitos (Planilha3), EstaPastaDeTrabalho -- de etapa1_multilote/src.
  Store   LotesStore ganha O=Analito e P=Chave (LOTE|ANALITO), preenchidas para
          o que ja existe.
  Nomes   loteSel, loteAnalise, loteParam, ls*, lstLotes, eng* novos;
          regLoteCol e aInput redefinidos.
  Calc    lote em analise no lugar do lote ativo; media/DP lidos da store pela
          chave; limites de eixo derivados da media/DP; lista de corridas vem
          do motor (e so dele); Registros so no lote carregado; ponto sem
          parametro nao vira "violacao".
  Painel  E3 = seletor de lote; faixa O3 fala do lote; M7/M8 nao dizem "Sem
          violacao" onde nao houve avaliacao; spinner refaz o motor.
  Config. cadastro de lotes vai a 100; lista compacta lstLotes.
  Estat.  H4 reflete o lote do Painel.

Nao salva se qualquer conferencia falhar.
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlh  # noqa: E402
import v2   # noqa: E402  camada de desempenho (ADR-050)
import painel  # noqa: E402  desenho do Painel (ADR-054)
import ux   # noqa: E402  camada de experiencia (ADR-051)

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, 'src')
SENHA = 'qcini2025'

MODULOS = ['mLotes', 'mEstatistica', 'mEstatPeriodo', 'mPeriodo', 'mBI',
           # ADR-050 -- desempenho com cinco anos de banco
           'mApp', 'mUI', 'mBanco', 'mOperacao', 'mImportar', 'mCEQ', 'mDados', 'mSeguranca']
DOC_MODULOS = {'Planilha7': 'Planilha7.cls', 'Planilha3': 'Planilha3.cls',
               'EstaPastaDeTrabalho': 'EstaPastaDeTrabalho.cls'}


def codigo(nome_arq):
    with open(os.path.join(SRC, nome_arq), encoding='utf-8', newline='') as f:
        linhas = f.read().replace('\r\n', '\n').split('\n')
    return '\r\n'.join(l for l in linhas if not l.startswith('Attribute ')).strip('\r\n') + '\r\n'


def substituir_codigo(vbp, comp, texto):
    cm = vbp.VBComponents(comp).CodeModule
    n = cm.CountOfLines
    if n:
        cm.DeleteLines(1, n)
    cm.AddFromString(texto)


def nome(wb, n, ref):
    try:
        wb.Names(n).Delete()
    except Exception:
        pass
    wb.Names.Add(n, ref)
    # prova de que o nome resolve (ler .RefersTo de volta nao denuncia corrupcao)
    return wb.Names(n).RefersTo


def desproteger(ws):
    estava = bool(ws.ProtectContents)
    if estava:
        ws.Unprotect(SENHA)
    return estava


def trocar_formula_coluna(ws, col, r0, r1, f):
    """Escreve a formula da linha r0 na coluna inteira (Excel ajusta o relativo)."""
    ws.Range(f'{col}{r0}:{col}{r1}').Formula = f


def main(caminho):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    wb = ex.abrir(caminho)
    ok = False
    try:
        ex.xl.Calculation = -4135      # manual durante a instalacao; um recalculo no fim
        vbp = wb.VBProject
        # ---------------- VBA ----------------
        existentes = [c.Name for c in vbp.VBComponents]
        if 'mLotes' not in existentes:
            raise SystemExit('mLotes nao existe no projeto -- arquivo inesperado')
        n = v2.conferir_duplicados([(m + '.bas', None) for m in MODULOS])
        print(f'VBA: {n} nomes publicos, nenhum duplicado')
        for m in MODULOS:
            v2.substituir_codigo(vbp, m, codigo(m + '.bas'))   # cria o modulo se faltar (mApp)
        for comp, arq in DOC_MODULOS.items():
            substituir_codigo(vbp, comp, codigo(arq))
        print('VBA: modulos substituidos')

        sh = {s.Name: s for s in wb.Worksheets}
        cfg, pai, calc = sh['Configuração'], sh['Painel'], sh['Calc']
        ls, ana, est, eng = sh['LotesStore'], sh['Analitos'], sh['Estatística'], sh['Eng_Saida']
        prot = {n: desproteger(s) for n, s in sh.items()}

        # ---------------- LotesStore: O/P ----------------
        ls.Range('N1').Copy(ls.Range('O1:P1'))
        ls.Range('O1').Value = 'Analito'
        ls.Range('P1').Value = 'Chave'
        ult = ls.Cells(ls.Rows.Count, 1).End(-4162).Row
        nomes_an = [ana.Cells(r, 1).Value for r in range(4, 44)]
        preench = 0
        for r in range(2, ult + 1):
            lote = ls.Cells(r, 1).Value
            idx = ls.Cells(r, 2).Value
            if lote in (None, '') or idx in (None, ''):
                continue
            nm = nomes_an[int(idx) - 1] if 1 <= int(idx) <= 40 else None
            if nm in (None, ''):
                continue
            lt = str(int(lote)) if isinstance(lote, float) and lote == int(lote) else str(lote).strip()
            ls.Cells(r, 15).Value = str(nm).strip()
            ls.Cells(r, 16).Value = f'{lt.upper()}|{str(nm).strip().upper()}'
            preench += 1
        print(f'LotesStore: {preench} linhas com chave')

        # ---------------- Configuracao ----------------
        cfg.Range('F2').Value = 'loteParam→'
        cfg.Range('G2').Value = cfg.Range('G1').Value
        cfg.Range('F1').Copy(cfg.Range('F2'))
        cfg.Range('G1').Copy(cfg.Range('G2'))
        cfg.Range('G2').Value = cfg.Range('G1').Value
        if cfg.Range('B51').Value in (None, ''):
            cfg.Range('B50:D50').Copy(cfg.Range('B51:D125'))
            for i in range(26, 101):
                cfg.Cells(25 + i, 2).Value = f'Lote {i:02d}'
            cfg.Range('C51:D125').ClearContents()
        b23 = str(cfg.Range('B23').Value or '')
        cfg.Range('B23').Value = b23.replace('até 25', 'até 100')
        cfg.Range('AB1').Value = 'lstLotes (auto - nao editar)'
        cfg.Range('Z1').Copy(cfg.Range('AB1'))
        cfg.Range('AB1').Value = 'lstLotes (auto - nao editar)'
        cfg.Range('AB2:AB101').Formula = (
            '=IFERROR(INDEX($C$26:$C$125,AGGREGATE(15,6,(ROW($C$26:$C$125)-ROW($C$26)+1)'
            '/($C$26:$C$125<>""),ROWS($AB$2:AB2))),"")')
        cfg.Range('AB2:AB101').Font.Color = cfg.Range('Z2').Font.Color
        cfg.Range('AB2:AB101').Font.Size = cfg.Range('Z2').Font.Size

        # ---------------- Nomes ----------------
        refs = {
            'loteSel': '=Painel!$H$3',
            # Names.Add le a formula no idioma LOCAL (pt-BR: SE, ÍNDICE, ';') --
            # em ingles e recusada. O arquivo guarda em ingles do mesmo jeito.
            'loteAnalise': '=SE(Painel!$H$3="";Configuração!$C$20;Painel!$H$3)',
            'loteParam': '=Configuração!$G$2',
            'filtroDe': '=Painel!$G$3', 'filtroAte': '=Painel!$G$4',
            'filtroAno': '=Painel!$K$3', 'filtroPeriodo': '=Painel!$K$4',
            'lsChave': '=LotesStore!$P$2:$P$4001',
            'lsMedN1': '=LotesStore!$C$2:$C$4001', 'lsDPN1': '=LotesStore!$D$2:$D$4001',
            'lsMedN2': '=LotesStore!$E$2:$E$4001', 'lsDPN2': '=LotesStore!$F$2:$F$4001',
            'lsMedN3': '=LotesStore!$G$2:$G$4001', 'lsDPN3': '=LotesStore!$H$2:$H$4001',
            # INDEX, e nao OFFSET: OFFSET e VOLATIL e arrasta para todo recalculo
            # tudo que dependa dele (ver lstAnosCIQ/lstAnosCEQ abaixo)
            'lstLotes': '=Configuração!$AB$2:ÍNDICE(Configuração!$AB$2:$AB$101;MÁXIMO(1;CONT.SE(Configuração!$AB$2:$AB$101;"?*")'
                        '+CONT.NÚM(Configuração!$AB$2:$AB$101)))',
            'lstAnosCIQ': '=Configuração!$Z$2:ÍNDICE(Configuração!$Z$2:$Z$50;MÁXIMO(1;CONT.VALORES(Configuração!$Z$2:$Z$50)))',
            'lstAnosCEQ': '=Configuração!$AA$2:ÍNDICE(Configuração!$AA$2:$AA$50;MÁXIMO(1;CONT.VALORES(Configuração!$AA$2:$AA$50)))',
            'regLoteCol': '=Configuração!$C$26:$C$125',
            'lstAnalitos': '=Analitos!$A$4:ÍNDICE(Analitos!$A$4:$A$43;MÁXIMO(1;CONT.VALORES(Analitos!$A$4:$A$43)))',
            'aInput': '=Analitos!$E$4:$J$43',
            'engNTotal': '=Eng_Saida!$K$1', 'engCortadas': '=Eng_Saida!$M$1',
            'engLotesPer': '=Eng_Saida!$O$1', 'engParam': '=Eng_Saida!$Q$1',
        }
        for n, r in refs.items():
            nome(wb, n, r)
        v = cfg.Range('C20').Validation
        v.Delete()
        v.Add(3, 1, 1, '=lstLotes')
        print('Nomes e validacao: ok')

        # ---------------- Eng_Saida: rotulos do cabecalho ----------------
        for c, t in (('J1', 'nTotal:'), ('L1', 'cortadas:'), ('N1', 'lotes no período:'), ('P1', 'param:')):
            eng.Range(c).Value = t

        # ---------------- Calc ----------------
        calc.Range('BH1').Formula = ('=IFERROR(MATCH(UPPER(TRIM(loteAnalise&""))&"|"&UPPER(TRIM(selAnalito)),'
                                     'lsChave,0),0)')
        calc.Range('BI1').Formula = ('=IFERROR(INDEX(Analitos!$D$4:$D$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))'
                                     ',1)+1')
        par = {'AX1': ('lsMedN1', 'lsDPN1', 'M'), 'AY1': ('lsMedN1', 'lsDPN1', 'S'),
               'AZ1': ('lsMedN2', 'lsDPN2', 'M'), 'BA1': ('lsMedN2', 'lsDPN2', 'S')}
        for cel, (m, s, qual) in par.items():
            devolve = f'INDEX({m},$BH$1)' if qual == 'M' else f'INDEX({s},$BH$1)'
            calc.Range(cel).Formula = (
                f'=IF($BH$1=0,"",IF(AND(ISNUMBER(INDEX({m},$BH$1)),ISNUMBER(INDEX({s},$BH$1))),'
                f'IF(INDEX({s},$BH$1)>0,{devolve},""),""))')
        eixos = {
            'BB1': ('INDEX(Analitos!$AC$4:$AC$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))', '($AX$1-3.3*$AY$1)'),
            'BC1': ('INDEX(Analitos!$AD$4:$AD$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))', '($AX$1+3.3*$AY$1)'),
            'BD1': ('INDEX(Analitos!$AG$4:$AG$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))', '($AZ$1-3.3*$BA$1)'),
            'BE1': ('INDEX(Analitos!$AH$4:$AH$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0))', '($AZ$1+3.3*$BA$1)'),
        }
        for cel, (velho, novo) in eixos.items():
            f = calc.Range(cel).Formula
            if novo in f:
                continue                      # ja instalado: reinstalar tem de ser possivel
            if velho not in f:
                raise SystemExit(f'Calc!{cel}: ancora do eixo nao encontrada: {f[:120]}')
            calc.Range(cel).Formula = f.replace(velho, novo)

        # corridas: a lista e a do MOTOR, e so vale se o motor for deste analito e deste lote
        trocar_formula_coluna(calc, 'B', 3, 182,
            '=IF(OR(selAnalito="",(""&engAnalito)<>(""&selAnalito),(""&engLote)<>(""&loteAnalise),'
            '$A3>N(engNRun)),"",INDEX(engRUN,$A3))')
        for col in ('C', 'F', 'AB'):
            f = calc.Range(f'{col}3').Formula
            if 'engData' in f or 'engValN' in f:
                continue          # ja le o motor (ADR-050): nao ha lote na formula
            if 'loteAtivo' not in f:
                raise SystemExit(f'Calc!{col}3 sem loteAtivo: {f[:120]}')
            trocar_formula_coluna(calc, col, 3, 182, f.replace('loteAtivo', 'loteAnalise'))
        for col in ('X', 'Y', 'Z', 'AT', 'AU', 'AV'):
            f = calc.Range(f'{col}3').Formula
            if not f.startswith('=IF(AND($D3=1,'):
                raise SystemExit(f'Calc!{col}3 formato inesperado: {f[:80]}')
            if 'loteCarregado' not in f:
                f = f.replace('=IF(AND($D3=1,', '=IF(AND($D3=1,(""&loteAnalise)=(""&loteCarregado),', 1)
            trocar_formula_coluna(calc, col, 3, 182, f)
        for col, st in (('J', '$P3'), ('AF', '$AL3')):
            f = calc.Range(f'{col}3').Formula
            velho = f'{st}<>"OK",{st}<>""'
            if velho in f:
                trocar_formula_coluna(calc, col, 3, 182, f.replace(velho, f'{st}="REJEITADO"'))
            elif f'{st}="REJEITADO"' not in f:
                raise SystemExit(f'Calc!{col}3 formato inesperado: {f[:120]}')
        # filtro de periodo por DIA (ADR-049): resultado com hora nao sai do ultimo dia
        f = calc.Range('D3').Formula
        f2 = f
        if 'INT($C3)' not in f2:
            if f2.count('$C3>=filtroDe') != 1 or f2.count('$C3<=filtroAte') != 1:
                raise SystemExit(f'Calc!D3 formato inesperado: {f2[:120]}')
            f2 = f2.replace('$C3>=filtroDe', 'INT($C3)>=INT(filtroDe)').replace('$C3<=filtroAte', 'INT($C3)<=INT(filtroAte)')
        f2 = painel.sem_trimestre(f2)      # ADR-054: o trimestre saiu daqui tambem
        if f2 != f:
            trocar_formula_coluna(calc, 'D', 3, 182, f2)
        print('Calc: formulas trocadas')

        # ---------------- Painel ----------------
        for rg in (cfg.Range('G1:G2'), cfg.Range('C20'), eng.Range('E1'), sh['Eventos_Westgard'].Range('L2')):
            rg.NumberFormat = '@'
        # A faixa de status do lote: o texto e daqui, o LUGAR e do painel.py
        # (ADR-054 tirou a faixa de O3:X4 e pos abaixo do grafico do nivel 2).
        faixa_f = (
            '=IF(OR(Calc!$AX$1="",Calc!$AZ$1=""),'
            '"⛔ LOTE "&loteAnalise&" · "&selAnalito&": SEM MÉDIA/DP em "'
            '&IF(AND(Calc!$AX$1="",Calc!$AZ$1=""),"N1 e N2",IF(Calc!$AX$1="","N1","N2"))'
            '&" — cadastre na aba Analitos. Westgard não avaliado. ",'
            '"Lote "&loteAnalise&" · alvo N1 "&FIXED(Calc!$AX$1,Calc!$BI$1)&" ± "&FIXED(Calc!$AY$1,Calc!$BI$1+1)'
            '&" · N2 "&FIXED(Calc!$AZ$1,Calc!$BI$1)&" ± "&FIXED(Calc!$BA$1,Calc!$BI$1+1)&". ")'
            '&IF(AND(COUNTIF(Calc!$D$3:$D$182,1)=0,selAnalito<>""),"Nenhuma corrida deste lote no período"'
            '&IF(engLotesPer<>""," (com dados: "&engLotesPer&")","")&". ","")'
            '&IF(N(engCortadas)>0,"⚠ "&engCortadas&" corrida(s) antigas do período fora do gráfico (máx. 180). ","")'
            '&IF(OR($G$3<>"",$G$4<>""),"Filtro "&IF($G$3<>"",TEXT($G$3,"dd/mm/aa"),"início")&" a "'
            '&IF($G$4<>"",TEXT($G$4,"dd/mm/aa"),"hoje"),"Sem filtro de data")')

        # ---------------- Estatistica / Analitos ----------------
        # Spinner de analito: ~11 s -> ~1,5 s. R/S/T/AC/AD usavam eqProvedor (o
        # provedor do analito EM TELA) para as 80 linhas: cada clique recalculava
        # ~400 UDFs de EQA, e cada linha era avaliada com o provedor de outro
        # analito. Agora cada linha usa o PROPRIO provedor, como a coluna G.
        prov_linha = 'IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A14,Analitos!$A$4:$A$43,0)),"CAP")'
        for col in ('R', 'S', 'T', 'AC', 'AD'):
            f = est.Range(f'{col}14').Formula
            if 'eqProvedor' in f:
                est.Range(f'{col}14:{col}93').Formula = f.replace('eqProvedor', prov_linha)
            elif 'MATCH($A14' not in f:
                raise SystemExit(f'Estatistica!{col}14 formato inesperado: {f[:120]}')
        est.Range('H4').Formula = '=loteAnalise'
        est.Range('I4').Value = '(refletido do Painel)'
        ana.Range('A2').Formula = (
            '="PARÂMETROS DO LOTE "&loteParam&"  ·  Média/DP valem só para este lote — para ver ou editar '
            'outro lote, selecione-o no Painel (H3).  ·  Lote em uso p/ lançamentos: "&loteAtivo')

        # ---------------- ADR-050: Calc le o motor; Liberacao por valor ----------------
        v2.aplicar_planilhas(ex, wb, 2)
        print('ADR-050: Calc lendo o motor, Liberacao A:B por valor')
        # ---------------- ADR-054: Painel no desenho padrao ----------------
        # A faixa de status muda de lugar (O3:X4 -> A39:U40, abaixo do grafico do
        # nivel 2), mas o TEXTO dela e o mesmo: quem sabe o que dizer ali e o
        # arquivo, nao este script. padronizar() preserva a formula que encontrar.
        extras = [
            '=IF(NOT(ISNUMBER(sigmaDoPlano)),"",IF(sigmaDoPlano<3,'
            '"DESEMPENHO INADEQUADO — CQ intensivo pode não ser suficiente. Investigar Bias, CV e '
            'desempenho do método.","Plano de CQ aplicável para o Sigma deste analito."))',
            '=IF(NOT(ISNUMBER(sigmaDoPlano)),"","Motor executa: "&Cfg_PlanoQC!$D$1&"   |   Cobertura: "'
            '&mPlanoQC.CoberturaWestgard(sigmaDoPlano))',
            'No quadro de Westgard, no alto: regra ILUMINADA = recomendada pelo Sigma deste analito · '
            'VERMELHO = regra violada no período · cinza = fora da estratégia atual.',
        ]
        # ADR-055: o realce "esta regra e a recomendada pelo Sigma" existia so na
        # Bioquimica -- e la estava MORTO desde o ADR-054 (lia Painel!V10:V11,
        # celulas que o Painel novo esvaziou). Agora le a coluna Sigma do Painel,
        # e existe nos dois produtos. Vem ANTES do padronizar: e ele quem usa os
        # nomes para a formatacao condicional do cabecalho de Westgard.
        painel.realce_regras(wb, 2, [('1-3S', '1_3s'), ('2-2S', '2_2s'), ('R4S', 'R_4s'),
                                     ('4-1S', '4_1s'), ('8X', '8x')])
        print('ADR-055: matriz de regras do Sigma em Cfg_PlanoQC (sigmaDoPlano vivo)')
        painel.padronizar(wb, 'Bioquímica', 2, 93, ['AX', 'AZ'],
                          [('1-3S', 13), ('2-2S', 14), ('R4S', 15), ('4-1S', 16), ('8X', 17)],
                          'A39:U40', faixa_f, 42, extras=extras, eqa=True,
                          fonte_etp='="Fonte do ETp / CVTp deste analito: "&IFERROR(INDEX('
                                    'Analitos!$S$4:$S$43,MATCH(selAnalito,Analitos!$A$4:$A$43,0)),"—")')
        print('ADR-054: Painel padronizado (topo, filtro Ano/Período, desempenho na coluna A)')
        col_nav = painel.coluna_da_barra(pai)
        print(f'ADR-055: barra de navegacao comeca em {col_nav}1 (depois do titulo + 1 coluna)')
        ux.aplicar(wb, 'Bioquímica', 'H3', 'A39:U40', 150000, col_nav + '1')
        print('ADR-051: navegacao, Inicio de aplicativo, status do lote colorido, Novo lote')

        # ---------------- compila e roda ----------------
        r = ex.run('mLotes.ConferirLotesStore', teto=60)
        print('ConferirLotesStore:', r)
        if not str(r).startswith('ok'):
            raise SystemExit('LotesStore inconsistente')
        ex.xl.Calculation = -4105
        t0 = time.time(); ex.xl.CalculateFull(); ex.esperar()
        print(f'CalculateFull: {time.time() - t0:.1f}s')
        ex.run('mLotes.SincronizarLotesAoAbrir', teto=60)
        ex.run('mEstatistica.InvalidarCache', teto=60)
        t0 = time.time()
        ex.run('mEstatistica.AtualizarEstatistica', teto=600)
        print(f'AtualizarEstatistica: {time.time() - t0:.1f}s')
        ex.xl.CalculateFull()
        ex.run('AtualizarEixos', teto=60)

        # conferencias minimas
        chk = {
            'loteSel': pai.Range('H3').Value, 'loteParam': cfg.Range('G2').Value,
            'AX1': calc.Range('AX1').Value, 'AY1': calc.Range('AY1').Value,
            'engLote': eng.Range('E1').Value, 'engNRun': eng.Range('I1').Value,
            'CalcB3': calc.Range('B3').Value, 'faixa': pai.Range('A39').Value,
            'PainelB7': pai.Range('B7').Value, 'PainelS7': pai.Range('S7').Value,
            'lstLotes': wb.Names('lstLotes').RefersToRange.Rows.Count,
        }
        for k, val in chk.items():
            print(f'  {k:10s} = {val}')
        if chk['AX1'] in (None, '') or chk['CalcB3'] in (None, ''):
            raise SystemExit('Calc sem parametro ou sem corrida apos aplicar -- nao salvo')
        # O filtro do Calc e o que o grafico enxerga: se ele zerar (ou virar
        # #NOME?), o Painel fica vazio e NADA acusa. Confere antes de salvar.
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
        nf = ex.xl.WorksheetFunction.CountIf(calc.Range('D3:D182'), 1)
        print(f'  corridas no filtro do Calc = {nf}')
        if not nf:
            raise SystemExit('filtro do Calc nao deixou passar nenhuma corrida -- nao salvo')

        for n, s in sh.items():
            if prot[n]:
                s.Protect(SENHA, False, True, True, True)  # DrawingObjects=False, Contents, Scenarios, UserInterfaceOnly
        wb.Save()
        ok = True
        print('SALVO:', caminho)
    finally:
        ex.fechar(salvar=False)
        if not ok:
            import traceback
            traceback.print_exc()
            print('*** NAO SALVO ***')
            sys.exit(1)


if __name__ == '__main__':
    main(sys.argv[1])
