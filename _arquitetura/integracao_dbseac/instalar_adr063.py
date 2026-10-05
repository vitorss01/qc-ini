# -*- coding: utf-8 -*-
"""instalar_adr063.py -- ADR-063: TODOS os graficos de Levey-Jennings a vista no Painel
(2 na Bioquimica, 3 na Hematologia). Substitui o instalador do ADR-061: o foco por nivel e o
Ctrl+Shift+G sairam; o vigia de zoom do ADR-061 continua.

Uso:  python instalar_adr063.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py); EXIGE compilacao;
  2. tira o botao btnGrafFoco ("Ampliar grafico") se existir, deixa TODOS os graficos visiveis e
     limpa o Texto Alt "qcini-geo:" que o foco gravava;
  3. botao btnGrafTodos ("Ver todos os graficos") na linha 5 do Painel, depois do titulo
     INDICADORES POR NIVEL, no estilo do btnHist; para se nao couber antes do bloco WESTGARD;
  4. FUMACA que CHAMA as rotinas (o VBA compila sob demanda): vigia nao arma em automacao,
     encaixe sai calado com o Excel oculto (o encaixe de verdade e provado em testes/qa_graficos.py,
     com Excel visivel);
  5. devolve a protecao de cada aba e a da estrutura como estavam.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import ux  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

BOTAO = 'btnGrafTodos'
ROTULO = 'Ver todos os gráficos'
ANTIGO = 'btnGrafFoco'
GEO = 'qcini-geo:'


def log(*a):
    print(*a, flush=True)


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:          # outro Excel segura o arquivo: o Save 'funcionaria' sem gravar nada
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        p = wb.Worksheets('Painel')

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        log('2. sem foco por nivel: todos os graficos visiveis')
        try:
            p.Shapes(ANTIGO).Delete()
            log(f'   {ANTIGO} removido')
        except Exception:
            pass
        n = p.ChartObjects().Count
        for co in p.ChartObjects():
            if not co.Visible:
                co.Visible = True
            sh = p.Shapes(co.Name)
            if str(sh.AlternativeText or '').startswith(GEO):
                sh.AlternativeText = ''
        if not all(co.Visible for co in p.ChartObjects()):
            raise SystemExit('ha grafico oculto no Painel')
        log(f'   {n} graficos, todos visiveis')

        log('3. botao ' + BOTAO)
        hist = p.Shapes('btnHist')
        a5 = p.Range('A5')
        x0 = a5.Left + ux.larg_texto(p, str(a5.Value), a5.Font.Name, a5.Font.Size, bool(a5.Font.Bold)) + 14
        larg = 150
        lim = p.Range('L5').Left - 6
        if x0 + larg > lim:
            raise SystemExit(f'{BOTAO} nao cabe na linha 5 ({x0:.0f}+{larg} > {lim:.0f})')
        b = ux._botao(p, BOTAO, 'PainelVerTodos', ROTULO, x0, hist.Top, larg, hist.Height,
                      hist.Fill.ForeColor.RGB, hist.TextFrame2.TextRange.Font.Fill.ForeColor.RGB,
                      tam=hist.TextFrame2.TextRange.Font.Size, negrito=True)
        b.AlternativeText = ('ADR-063: encaixa o Painel na janela -- cabecalho inteiro e TODOS os graficos '
                             'de Levey-Jennings inteiros na tela (volta depois de um zoom manual)')
        log(f'   em L={x0:.0f} T={hist.Top:.0f} W={larg} (limite {lim:.0f})')

        log('4. fumaca (chama cada rotina)')
        run = lambda m, *a: ex.run("'" + wb.Name + "'!" + m, *a, teto=120)
        run('mUI.AjustarGraficos')
        run('mUI.GrafVigiaLigar')                  # automacao: nao arma (Interactive=False) -- e o certo
        est = str(run('mUI.GrafVigiaEstado'))
        if not est.startswith('0|'):
            raise SystemExit(f'vigia armou em automacao: {est}')
        run('mUI.GrafVigiaDesligar')
        run('mUI.GrafVigiaEncerrar')               # o do fechamento: sem tique agendado, nao espera
        r = str(run('mUI.PainelEncaixar'))
        if r.startswith('ERRO') or r.startswith('OK|'):
            # OK nao pode: com o Excel OCULTO o encaixe nao tem janela para medir
            raise SystemExit(f'PainelEncaixar em automacao devolveu {r!r}')
        run('mUI.PainelVerTodos')
        txt = p.Shapes(BOTAO).TextFrame2.TextRange.Text
        if txt != ROTULO:
            raise SystemExit(f'rotulo do botao: {txt!r}')
        log(f'   vigia nao arma em automacao; encaixe sai calado sem janela ({r}); botao "{txt}"')

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
