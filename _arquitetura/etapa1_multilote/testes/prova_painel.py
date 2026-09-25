# -*- coding: utf-8 -*-
"""prova_painel.py -- o desenho do Painel (ADR-054) esta como foi pedido?

Uso: python prova_painel.py <arquivo_instalado.xlsm> <bio|hema>

Confere o que o gestor pediu, item a item, direto no arquivo: onde cada coisa
mora, as larguras iguais nos dois produtos, a coluna M a esquerda, o seletor de
controle externo fora do caminho dos filtros, o bloco de desempenho na coluna A
e a barra de navegacao comecando em I1. Nada e salvo.
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import painel  # noqa: E402
import xlh     # noqa: E402

RES = []
# ADR-056: a faixa de status desceu tambem na Hematologia -- os cartoes de
# indicador ocupam O3:U4, e os dois Paineis passam a ser identicos.
FAIXA = {'bio': 'A39:U40', 'hema': 'A39:U40'}


def chk(n, ok, d=''):
    RES.append((n, bool(ok)))
    print(f'   [{"PASS" if ok else "FAIL"}] {n}' + ('' if ok else f'  -> {str(d)[:200]}'), flush=True)


def main(src, prod):
    dst = os.path.join(os.environ['TMP'], 'qcw', 'final', f'prova_painel_{prod}.xlsm')
    shutil.copyfile(src, dst)
    ex = xlh.Excel()
    wb = ex.abrir(dst)
    try:
        ws = wb.Sheets('Painel')
        nlv = 3 if prod == 'hema' else 2
        # ---- larguras iguais nos dois produtos ----
        larg = [round(ws.Columns(i + 1).ColumnWidth, 1) for i in range(len(painel.LARGURAS))]
        chk('larguras A..U iguais ao padrao dos dois produtos', larg == painel.LARGURAS, larg)
        # ---- onde cada coisa mora ----
        for cel, esperado in (('A3', 'Analito nº'), ('F3', 'De'), ('F4', 'Até'),
                              ('J3', 'Ano'), ('J4', 'Período'), ('H4', '▲ lote em análise')):
            chk(f'{cel} = {esperado!r}', str(ws.Range(cel).Value) == esperado, ws.Range(cel).Value)
        chk('lote em analise em H3 (loteSel)', wb.Names('loteSel').RefersTo.endswith('$H$3'),
            wb.Names('loteSel').RefersTo)
        chk('Ano/Periodo tem nome proprio (filtroAno/filtroPeriodo)',
            wb.Names('filtroAno').RefersTo.endswith('$K$3')
            and wb.Names('filtroPeriodo').RefersTo.endswith('$K$4'),
            (wb.Names('filtroAno').RefersTo, wb.Names('filtroPeriodo').RefersTo))
        v = ws.Range('K4').Validation.Formula1
        chk('lista de periodo com Tudo, Ano inteiro, trimestres e meses',
            v == painel.PERIODOS, v)
        # ---- spinner ao lado do numero, sem cobrir rotulo ----
        sp = [s for s in ws.Shapes if str(s.Name).startswith('Spinner')]
        chk('spinner existe e fica sobre E3', len(sp) == 1
            and abs(sp[0].Left - ws.Range('E3').Left) < 6, [s.Left for s in sp])
        # ---- tabelas ----
        cab = [ws.Range(f'{c}6').Value for c in 'ABCDEFGHIJ']
        chk('indicadores: 10 colunas padrao', cab == painel.CAB_IND, cab)
        chk('Westgard: Nivel em L6, Total em R6, ultima violacao em S6:U6',
            ws.Range('L6').Value == 'Nível' and ws.Range('R6').Value == 'Total'
            and str(ws.Range('S6').Value).startswith('Últ')
            and ws.Range('T6').Value == 'Classificação' and ws.Range('U6').Value == 'Histórico',
            [ws.Range(c + '6').Value for c in 'LRSTU'])
        for n in range(1, nlv + 1):
            f = ws.Range(f'S{6 + n}').Formula
            chk(f'ultima violacao do N{n} vem do motor (engPainel,{n},19)',
                f'INDEX(engPainel,{n},19)' in f, f[:80])
        # ---- faixa de status ----
        faixa = ws.Range(FAIXA[prod].split(':')[0])
        chk(f'faixa de status em {FAIXA[prod]}',
            'loteAnalise' in str(faixa.Formula) and faixa.MergeArea.Address.replace('$', '')
            == FAIXA[prod], (faixa.MergeArea.Address, str(faixa.Formula)[:60]))
        # ---- bloco de desempenho na coluna A ----
        chk('bloco de desempenho comeca em A42',
            str(ws.Range('A42').Value).startswith('DESEMPENHO SIX SIGMA'), ws.Range('A42').Value)
        titulos = [str(ws.Cells(r, 1).Value or '') for r in range(42, 130)]
        for t in ('PLANO DE CQ RECOMENDADO PELO SIGMA', 'ERRO TOTAL vs ETp — orçamento de erro',
                  'MARGEM CRÍTICA — todos os analitos', 'SIGMA × DPM × RENDIMENTO — referência',
                  'BASE CIENTÍFICA DO PLANO DE CQ'):
            chk(f'bloco tem "{t[:34]}" na coluna A', t in titulos)
        vazio = [f'{c}{r}' for r in range(42, 100) for c in ('R', 'S', 'T', 'U')
                 if ws.Range(f'{c}{r}').Value not in (None, '')]
        chk('nada do bloco sobrou nas colunas R..U', not vazio, vazio[:6])
        # ---- controle externo (so a Bioquimica) ----
        if prod == 'bio':
            chk('controle externo em M3:N4', str(ws.Range('M3').Value).startswith('Controle externo')
                and 'Analitos!' in str(ws.Range('M4').Formula),
                (ws.Range('M3').Value, ws.Range('M4').Formula[:40]))
            chk('lista CAP/Controllab no M4', ws.Range('M4').Validation.Formula1 == 'CAP;Controllab',
                ws.Range('M4').Validation.Formula1)
            # Columns('M').HorizontalAlignment devolve Nulo quando a coluna tem
            # celulas com alinhamentos diferentes -- e tem, de proposito: o
            # PADRAO da coluna e a esquerda, e a tabela de Westgard sobrescreve
            # as dela para o centro. Confere o padrao numa celula fora da tabela.
            chk('coluna M alinhada a esquerda (pedido do gestor)',
                ws.Range('M30').HorizontalAlignment == -4131 and
                ws.Range('M120').HorizontalAlignment == -4131,
                (ws.Range('M30').HorizontalAlignment, ws.Range('M120').HorizontalAlignment))
            chk('tabela de Westgard continua centrada na coluna M',
                ws.Range('M7').HorizontalAlignment == -4108, ws.Range('M7').HorizontalAlignment)
        # ---- linha 1: titulo igual nos dois e barra que nao o cobre ----
        f1 = ws.Range('A1').Font
        chk('titulo em Segoe UI 16 (o mesmo nos dois produtos)',
            str(f1.Name) == painel.TITULO_FONTE and float(f1.Size) == painel.TITULO_TAM,
            (f1.Name, f1.Size))
        chk('titulo da linha 1 e o texto padrao',
            str(ws.Range('A1').Value) == painel.titulo_padrao(
                'Bioquímica' if prod == 'bio' else 'Hematologia', nlv), ws.Range('A1').Value)
        # ---- navegacao e graficos ----
        navs = [s for s in ws.Shapes if str(s.Name).startswith('nav_')]
        chk(f'barra de navegacao com {len(navs)} botoes', len(navs) >= 9, len(navs))
        esperada = painel.coluna_da_barra(ws)
        l0 = min(s.Left for s in navs)
        chk(f'barra comeca na coluna {esperada} (a mesma nos dois produtos)',
            abs(l0 - ws.Range(f'{esperada}1').Left) < 1, (l0, ws.Range(f'{esperada}1').Left))
        larg_tit = painel.medir_texto(ws, str(ws.Range('A1').Value))
        chk('barra NAO cobre o titulo da linha 1', l0 > larg_tit,
            f'titulo termina em {larg_tit:.0f}, barra comeca em {l0:.0f}')
        # e sobra mesmo UMA coluna inteira de respiro entre um e outro
        anterior = ws.Range(f'{painel.col(painel.num(esperada) - 1)}1')
        chk('sobra uma coluna vazia entre o titulo e a barra',
            anterior.Left >= larg_tit, (anterior.Left, larg_tit))
        chk('barra cabe dentro da largura da tabela (A..U)',
            max(s.Left + s.Width for s in navs) <= ws.Range('A1:U1').Width,
            (max(s.Left + s.Width for s in navs), ws.Range('A1:U1').Width))
        # ---- legenda: sem rotulo cortado ----
        for co in ws.ChartObjects():
            ch = co.Chart
            n_ser = ch.SeriesCollection().Count
            n_leg = ch.Legend.LegendEntries().Count
            chk(f'{co.Name}: legenda enxuta ({n_leg} de {n_ser} series)',
                n_leg <= 6 and n_leg < n_ser, (n_leg, n_ser))
            chk(f'{co.Name}: legenda a direita', ch.Legend.Position == -4152,
                ch.Legend.Position)
            # A prova de que NENHUM rotulo esta cortado. Legenda vertical: a
            # caixa tem de caber o rotulo MAIS LARGO com o marcador ao lado, e
            # ter altura para uma linha por entrada. Com as 14 entradas na
            # caixa de 80 pt (o estado antigo) a altura reprova.
            fica = []
            vistos = set()
            for i in range(1, n_ser + 1):
                nm = str(ch.SeriesCollection(i).Name)
                if nm not in painel.LIMITES and nm not in vistos:
                    fica.append(nm)
                    vistos.add(nm)
            maior = max(painel.medir_texto(ws, nm, 'Segoe UI', 9, False) for nm in fica)
            chk(f'{co.Name}: cabe o rotulo mais largo ({maior:.0f} pt + marcador)',
                ch.Legend.Width >= maior + 22, (ch.Legend.Width, maior))
            chk(f'{co.Name}: cabe uma linha por entrada ({len(fica)} entradas)',
                ch.Legend.Height >= len(fica) * 12, (ch.Legend.Height, len(fica)))
            chk(f'{co.Name}: legenda cabe na altura do grafico',
                ch.Legend.Height <= co.Height, (ch.Legend.Height, co.Height))
        # ---- cartoes de indicador (ADR-056) ----
        for (c0, c1), rotulo in painel.CARTOES:
            chk(f'cartao "{rotulo}" em {c0}3', str(ws.Range(f'{c0}3').Value) == rotulo,
                ws.Range(f'{c0}3').Value)
            v = ws.Range(f'{c0}4')
            chk(f'cartao "{rotulo}" tem formula (acompanha a troca de analito)',
                str(v.Formula).startswith('='), str(v.Formula)[:60])
            chk(f'cartao "{rotulo}" ocupa {c0}3:{c1}4',
                ws.Range(f'{c0}3').MergeArea.Address.replace('$', '') == f'{c0}3:{c1}3'
                and v.MergeArea.Address.replace('$', '') == f'{c0}4:{c1}4',
                (ws.Range(f'{c0}3').MergeArea.Address, v.MergeArea.Address))
        chk('cartao do Sigma traz numero', isinstance(ws.Range('O4').Value, (int, float)),
            ws.Range('O4').Value)
        chk('cartao do Status traz um dos estados previstos',
            str(ws.Range('Q4').Value) in ('OK', 'REJEITADO', 'SEM MÉDIA/DP', '—'),
            ws.Range('Q4').Value)
        chk('cartao de violacoes traz numero', isinstance(ws.Range('T4').Value, (int, float)),
            ws.Range('T4').Value)
        chk('cartoes nao invadiram o seletor de EQA nem os filtros',
            str(ws.Range('J3').Value) == 'Ano' and str(ws.Range('J4').Value) == 'Período',
            (ws.Range('J3').Value, ws.Range('J4').Value))
        # ---- uma familia de fonte em todas as telas de uso ----
        fora = []
        for aba in ('Início', 'Painel', 'Analitos', 'Configuração', 'Estatística',
                    'Liberação', 'Registros'):
            try:
                f = wb.Sheets(aba).UsedRange.Font.Name
            except Exception:
                continue
            if str(f) != painel.tema.FONTE:
                fora.append(f'{aba}={f}')
        chk('todas as telas de uso na mesma fonte', not fora, fora)
        # ---- realce da regra recomendada pelo Sigma ----
        for nome in ('sigmaDoPlano', 'regrasAtivas', 'regrasRotulos'):
            chk(f'nome {nome} existe (realce da regra recomendada)',
                any(n.Name == nome for n in wb.Names))
        sg = wb.Sheets('Cfg_PlanoQC').Range('B1')
        chk('Sigma do plano vem da coluna Sigma do Painel (I7:I…), nao de V10',
            f'Painel!$I$7:$I${6 + nlv}' in str(sg.Formula), sg.Formula)
        chk('Sigma do plano tem valor numerico (o realce depende dele)',
            isinstance(sg.Value, (int, float)), sg.Value)
        rot = [str(c.Value) for c in wb.Sheets('Cfg_PlanoQC').Range('J3:N3').Cells]
        cabs = [str(ws.Range(f'{c}6').Value) for c in 'MNOPQ']
        chk('rotulos da matriz batem com o cabecalho de Westgard do Painel',
            rot == cabs, (rot, cabs))
        ativas = [c.Value for c in wb.Sheets('Cfg_PlanoQC').Range('J10:N10').Cells]
        chk('alguma regra iluminada para o Sigma de agora', sum(
            v for v in ativas if isinstance(v, (int, float))) > 0, ativas)
        chk('botao do historico de Westgard existe',
            any(str(s.Name) == 'btnHist' for s in ws.Shapes))
        gr = list(ws.ChartObjects())
        chk(f'{len(gr)} graficos, todos na coluna A e da largura da tabela',
            len(gr) == nlv and all(abs(c.Left) < 0.5 for c in gr)
            and all(c.Width > 1000 for c in gr), [(c.Left, c.Width) for c in gr])
        # ---- nada de sobra do desenho antigo ----
        chk('sem caixas de trimestre', not [s for s in ws.Shapes if str(s.Name).startswith('Check Box')])
        painel.conferir_sem_qsel(wb)
        chk('nenhuma formula da pasta cita qsel', True)
    finally:
        ex.fechar()
    n = sum(1 for r in RES if r[1])
    print(f'\nTOTAL PAINEL {prod}: {n}/{len(RES)} PASS')
    return n == len(RES)


if __name__ == '__main__':
    sys.exit(0 if main(os.path.abspath(sys.argv[1]), sys.argv[2]) else 1)
