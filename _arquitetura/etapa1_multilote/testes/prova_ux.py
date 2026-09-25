# -*- coding: utf-8 -*-
"""prova_ux.py -- a camada de experiencia (ADR-051) funciona? (copia; nada e salvo)

Uso: python prova_ux.py <arquivo_instalado.xlsm> <bio|hema>

Confere: cada botao da barra leva a tela certa; o analito escolhido pelo NOME
vira a posicao do spinner e o motor acompanha; "Media/DP deste lote" abre a
linha do analito; "Trocar lote" leva ao seletor; Novo lote cadastra, recusa
duplicado e data invalida, passa a lote em uso e a lote em analise; o filtro
de periodo refaz o motor; e nada disso deixa o Excel em calculo manual.
"""
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import painel  # noqa: E402  desenho do Painel (ADR-054/055)
import xlh     # noqa: E402

RES = []

# ADR-054: a faixa de status desceu para baixo do grafico na Bioquimica
FAIXA = {'bio': 'A39', 'hema': 'O3'}


def chk(n, ok, d=''):
    RES.append((n, bool(ok)))
    print(f'   [{"PASS" if ok else "FAIL"}] {n}' + ('' if ok else f'  -> {str(d)[:200]}'), flush=True)


def main(src, prod):
    dst = os.path.join(os.environ['TMP'], 'qcw', 'final', f'prova_ux_{prod}.xlsm')
    shutil.copyfile(src, dst)
    ex = xlh.Excel(visivel=True)
    wb = ex.abrir(dst)
    try:
        estrutura0 = bool(wb.ProtectStructure)
        # login/logout de verdade: LockApp esconde tudo menos o Login; UnlockApp mostra
        ex.run('LockApp')
        vis = [s.Name for s in wb.Worksheets if s.Visible == -1]
        chk(f'LockApp: so o Login visivel ({vis})', vis == ['Login'], vis)
        chk('LockApp devolve a protecao de estrutura original', bool(wb.ProtectStructure) == estrutura0,
            (wb.ProtectStructure, estrutura0))
        wb.Names('currentPapel').RefersToRange.Value = 'ANALISTA'
        ex.run('UnlockApp')
        vis = [s.Name for s in wb.Worksheets if s.Visible == -1]
        chk(f'UnlockApp: telas de uso visiveis ({len(vis)}), Calc/LotesStore ocultas',
            'Painel' in vis and 'Início' in vis and 'Calc' not in vis and 'Login' not in vis, vis)
        chk('UnlockApp abre no Inicio', ex.xl.ActiveSheet.Name == 'Início', ex.xl.ActiveSheet.Name)
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for s in wb.Worksheets:
            s.Visible = -1
        wb.Sheets('Início').Activate()
        destinos = {'IrPainel': 'Painel', 'IrAnalitos': 'Analitos', 'IrLotes': 'Configuração',
                    'IrEstatistica': 'Estatística', 'IrLiberacao': 'Liberação', 'IrRegistros': 'Registros',
                    'IrEventos': 'Eventos_Westgard', 'IrInicio': 'Início'}
        for m, aba in destinos.items():
            ex.run('mApp.' + m)
            chk(f'botao {m} -> {aba}', ex.xl.ActiveSheet.Name == aba, ex.xl.ActiveSheet.Name)
        for aba in ('Painel', 'Analitos', 'Configuração', 'Estatística', 'Liberação', 'Registros'):
            nomes = [s.Name for s in wb.Sheets(aba).Shapes]
            chk(f'barra de navegacao em {aba}', all(f'nav_{k}' in nomes for k in
                                                   ('inicio', 'painel', 'lancar', 'analitos', 'lotes', 'estat')),
                nomes[:5])
        chk('Inicio tem os 8 botoes de tarefa',
            sum(1 for s in wb.Sheets('Início').Shapes if s.Name.startswith('tile_')) >= 8)
        # analito pelo nome
        pai = wb.Sheets('Painel')
        ana = wb.Sheets('Analitos')
        alvo = ana.Range('A9').Value
        pai.Range('C3').Value = alvo
        ex.run('mApp.AnalitoEscolhido', pai.Range('C3'))
        ex.esperar()
        chk(f'analito pela lista: B3 = 6 ({alvo})', pai.Range('B3').Value == 6, pai.Range('B3').Value)
        chk('C3 volta a ser formula', str(pai.Range('C3').Formula).startswith('='), pai.Range('C3').Formula)
        chk('motor acompanhou o analito escolhido', wb.Sheets('Eng_Saida').Range('C1').Value == alvo,
            wb.Sheets('Eng_Saida').Range('C1').Value)
        ex.run('mApp.IrParametrosDoAnalito')
        chk('Media/DP deste lote abre a linha do analito na aba Analitos',
            ex.xl.ActiveSheet.Name == 'Analitos' and ex.xl.ActiveCell.Row == 9,
            (ex.xl.ActiveSheet.Name, ex.xl.ActiveCell.Row))
        ex.run('mApp.IrLotePainel')
        cel = wb.Names('loteSel').RefersToRange.Address
        chk(f'Trocar lote leva ao seletor do Painel ({cel})',
            ex.xl.ActiveSheet.Name == 'Painel' and ex.xl.ActiveCell.Address == cel, ex.xl.ActiveCell.Address)
        # Novo lote
        r = ex.run('mApp.CadastrarLote', '7777', '31/12/2027')
        chk(f'Novo lote cadastrado: {r}', str(r).startswith('OK'), r)
        r2 = ex.run('mApp.CadastrarLote', '7777', '')
        chk('Novo lote recusa duplicado', str(r2).startswith('ERRO'), r2)
        r3 = ex.run('mApp.CadastrarLote', '7778', '32/13/2027')
        chk('Novo lote recusa data invalida (e nao cadastra)', str(r3).startswith('ERRO'), r3)
        ex.xl.Calculate()
        lista = [c.Value for c in wb.Names('lstLotes').RefersToRange]
        chk('lote novo aparece na lista de lotes', '7777' in [str(x) for x in lista], lista)
        t0 = time.time()
        ex.run('mApp.UsarLote', '7777')
        ex.esperar()
        chk(f'Usar lote: lote em uso = 7777 ({time.time() - t0:.1f}s)',
            str(wb.Names('loteAtivo').RefersToRange.Value) == '7777', wb.Names('loteAtivo').RefersToRange.Value)
        t0 = time.time()
        ex.run('mApp.AnalisarLote', '7777')
        ex.esperar()
        o3 = str(pai.Range(FAIXA[prod]).Value)
        chk(f'lote novo em analise: faixa avisa SEM MEDIA/DP ({time.time() - t0:.1f}s)',
            o3.startswith('⛔ LOTE 7777'), o3[:80])
        chk('Analitos mostra o lote novo VAZIO (nao herda media/DP)',
            all(v in (None, '') for linha in ana.Range('E4:J10').Value for v in linha), ana.Range('E4:F5').Value)
        lotes = [str(x) for x in lista if x not in (None, '', '7777')]
        ex.run('mApp.AnalisarLote', lotes[0])
        pai.Range('G3').Value = ''
        pai.Range('G4').Value = ''
        ex.run('mEstatistica.FiltroMudou')
        n_all = wb.Sheets('Eng_Saida').Range('I1').Value
        chk(f'FiltroMudou refaz o motor (nRun = {n_all})', n_all is not None)
        chk('calculo automatico ao fim de tudo', ex.xl.Calculation == -4105, ex.xl.Calculation)
        chk('nenhuma operacao ficou presa (mApp)', not ex.run('mApp.Ocupado'))
        chk('faixa de status tem formatacao condicional',
            pai.Range(FAIXA[prod]).FormatConditions.Count >= 3, pai.Range(FAIXA[prod]).FormatConditions.Count)
        # ---- ADR-054: um filtro de periodo so (Ano + Periodo) -------------------
        pai.Activate()
        for n in ('qsel1', 'qsel2', 'qsel3', 'qsel4'):
            try:
                wb.Names(n)
                chk(f'nome {n} apagado (filtro duplicado)', False, 'ainda existe')
            except Exception:
                chk(f'nome {n} apagado (filtro duplicado)', True)
        cxs = [s.Name for s in pai.Shapes if str(s.Name).startswith('Check Box')]
        chk('caixas de trimestre removidas do Painel', not cxs, cxs)
        f = wb.Sheets('Calc').Range('D3').Formula
        chk('filtro do Calc sem qsel', 'qsel' not in f and 'filtroDe' in f, f[:120])
        nf = ex.xl.WorksheetFunction.CountIf(wb.Sheets('Calc').Range('D3:D182'), 1)
        chk(f'filtro do Calc deixa passar corridas (n = {nf})', nf and nf > 0, nf)

        pai.Range('K3').Value = ''
        pai.Range('K4').Value = 'Ano inteiro'
        r = ex.run('mApp.AplicarFiltroAnoMes')
        ano = pai.Range('K3').Value
        chk(f'Ano vazio e preenchido sozinho com o ano dos dados ({ano}) [{r}]',
            isinstance(ano, (int, float)) and 1900 < ano < 2200, r)
        ex.run('mEstatistica.FiltroMudou')
        ex.esperar()
        ex.xl.Calculate()
        n_total = pai.Range('B7').Value
        achou = None
        for p in ('T1', 'T2', 'T3', 'T4'):
            pai.Range('K4').Value = p
            ex.run('mApp.AplicarFiltroAnoMes')
            ex.run('mEstatistica.FiltroMudou')
            ex.esperar()
            ex.xl.Calculate()
            n_p = pai.Range('B7').Value
            if isinstance(n_p, (int, float)) and isinstance(n_total, (int, float)) and n_p < n_total:
                achou = (p, n_p)
                break
        chk(f'trimestre filtra o Painel ({achou[0] if achou else "?"}: n {n_total} -> '
            f'{achou[1] if achou else "?"})', achou is not None, (n_total, achou))
        aba = ex.xl.ActiveSheet.Name
        chk(f'trimestre NAO tira o usuario do Painel (ficou em {aba})', aba == 'Painel', aba)
        chk('trimestre escreve as datas De/Ate',
            pai.Range('G3').Value not in (None, '') and pai.Range('G4').Value not in (None, ''),
            (pai.Range('G3').Value, pai.Range('G4').Value))
        pai.Range('K4').Value = 'Ano inteiro'
        ex.run('mApp.AplicarFiltroAnoMes')
        ex.run('mEstatistica.FiltroMudou')
        ex.esperar()
        ex.xl.Calculate()
        chk(f'voltar para o ano inteiro devolve o periodo todo (n = {pai.Range("B7").Value})',
            pai.Range('B7').Value == n_total, (pai.Range('B7').Value, n_total))
        pai.Range('K4').Value = 'Tudo (sem filtro)'
        r = ex.run('mApp.AplicarFiltroAnoMes')
        chk(f'"Tudo (sem filtro)" limpa as datas De/Ate [{r}]',
            pai.Range('G3').Value in (None, '') and pai.Range('G4').Value in (None, ''),
            (r, pai.Range('G3').Value, pai.Range('G4').Value))
        ex.run('mEstatistica.FiltroMudou')
        ex.esperar()
        ex.xl.Calculate()
        n_tudo = pai.Range('B7').Value
        chk(f'sem filtro o Painel mostra tudo (n = {n_tudo} >= {n_total})',
            isinstance(n_tudo, (int, float)) and n_tudo >= n_total, (n_tudo, n_total))
        chk('filtro nao deixou o Excel em calculo manual', ex.xl.Calculation == -4105, ex.xl.Calculation)

        # ---- ADR-054: o desenho padrao dos dois produtos ------------------------
        chk('titulo do bloco de desempenho na COLUNA A (A42)',
            str(pai.Range('A42').Value).startswith('DESEMPENHO'), pai.Range('A42').Value)
        chk('Westgard mostra a ultima violacao (S6:U6)',
            str(pai.Range('S6').Value).startswith('Últ') and pai.Range('T6').Value == 'Classificação'
            and pai.Range('U6').Value == 'Histórico',
            [pai.Range(c + '6').Value for c in 'STU'])
        chk('indicadores por nivel com Sigma e Status (I6/J6)',
            pai.Range('I6').Value == 'Sigma' and pai.Range('J6').Value == 'Status',
            (pai.Range('I6').Value, pai.Range('J6').Value))
        # ADR-055: a barra comeca DEPOIS do titulo da linha 1, mais uma coluna
        # vazia -- e na MESMA coluna nos dois produtos
        esq = min(s.Left for s in pai.Shapes if str(s.Name).startswith('nav_'))
        alvo = painel.coluna_da_barra(pai)
        chk(f'barra de navegacao comeca na coluna {alvo} ({esq:.0f} pt)',
            abs(esq - pai.Range(f'{alvo}1').Left) < 1, (esq, pai.Range(f'{alvo}1').Left))
        tit = painel.medir_texto(pai, str(pai.Range('A1').Value))
        chk('a barra nao cobre o titulo da linha 1', esq > tit,
            f'titulo ate {tit:.0f} pt, barra a partir de {esq:.0f} pt')
        larg = {}
        janela = wb.Windows(1)
        for z in (70, 130):
            janela.Zoom = z
            ex.run('mUI.AjustarGraficos')
            larg[z] = pai.ChartObjects()(1).Width
        chk(f'grafico acompanha o zoom ({larg[70]:.0f} pt em 70%, {larg[130]:.0f} pt em 130%)',
            larg[130] < larg[70] - 10, larg)
        janela.Zoom = 70
        ex.run('mUI.AjustarGraficos')
        alvo_larg = janela.UsableWidth * 100.0 / janela.Zoom - 6
        chk(f'grafico ocupa a largura visivel (alvo {alvo_larg:.0f} pt)',
            abs(pai.ChartObjects()(1).Width - alvo_larg) < 3, (pai.ChartObjects()(1).Width, alvo_larg))
        chk('todos os graficos comecam na coluna A',
            all(abs(co.Left) < 0.5 for co in pai.ChartObjects()), [co.Left for co in pai.ChartObjects()])
        # redimensionar a janela tambem muda a largura visivel (Workbook_WindowResize)
        ex.xl.EnableEvents = True
        estado0, larg0 = janela.WindowState, None
        try:
            janela.WindowState = -4143            # xlNormal, para poder mudar a largura
            larg0 = janela.Width
            janela.Width = larg0 * 0.6
            ex.esperar()
            depois = pai.ChartObjects()(1).Width
            chk(f'grafico acompanha o redimensionamento da janela ({depois:.0f} pt)',
                depois < larg[70] - 10, (depois, larg[70]))
        finally:
            if larg0:
                janela.Width = larg0
            janela.WindowState = estado0
            ex.xl.EnableEvents = False

        # ---- trocar de analito depois de tudo isso nao pode dar erro ------------
        erro = ''
        try:
            for b in (3, 4, 5):
                pai.Range('B3').Value = b
                ex.run('mEstatistica.PainelMudou')
                ex.esperar()
        except Exception as e:
            erro = str(e)[:160]
        chk('trocar de analito depois do filtro: sem erro', not erro, erro)
        ex.xl.Calculate()
        chk(f'Painel responde ao analito novo (n = {pai.Range("B7").Value})',
            pai.Range('B7').Value is not None and str(pai.Range('C3').Value) != '',
            (pai.Range('B7').Value, pai.Range('C3').Value))
        chk('nada ficou preso depois da troca', not ex.run('mApp.Ocupado'))
    finally:
        ex.fechar()
    n = sum(1 for r in RES if r[1])
    print(f'\nTOTAL UX {prod}: {n}/{len(RES)} PASS')


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), sys.argv[2])
