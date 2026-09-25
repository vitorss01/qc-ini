# -*- coding: utf-8 -*-
"""foto_painel.py -- exporta o Painel como imagem, para ver o desenho sem abrir o arquivo.

Uso: python foto_painel.py <arquivo.xlsm> <saida.png> [linhas]

Copia a faixa visivel do Painel como figura de TELA (inclui gráficos e botões,
que são objetos flutuantes) e exporta por um gráfico temporário. Nada é salvo no
arquivo de origem: o trabalho é feito numa cópia.
"""
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402


def main(src, saida, linhas=60):
    dst = os.path.join(os.environ['TMP'], 'qcw', 'final', 'foto_' + os.path.basename(src))
    shutil.copyfile(src, dst)
    ex = xlh.Excel(visivel=True)
    wb = ex.abrir(dst)
    try:
        try:
            wb.Unprotect('qcini2025')
        except Exception:
            pass
        for s in wb.Worksheets:
            try:
                s.Visible = -1
            except Exception:
                pass
        ws = wb.Sheets('Painel')
        try:
            ws.Unprotect('qcini2025')     # para poder criar o quadro temporario
        except Exception:
            pass
        ws.Activate()
        w = wb.Windows(1)
        w.WindowState = -4137                 # xlMaximized: a tela inteira
        w.Zoom = 70
        w.ScrollRow = 1
        w.ScrollColumn = 1
        ex.run('mUI.AjustarGraficos', teto=60)
        time.sleep(1.5)
        # O grafico e ELASTICO: depois de AjustarGraficos ele pode ser mais
        # largo que A..U. Capturar so ate U cortava a legenda NA FOTO -- e eu
        # cheguei a diagnosticar como defeito do arquivo o que era defeito da
        # minha camera. A faixa vai ate a coluna que cobre o grafico.
        larg_max = max([c.Width for c in ws.ChartObjects()] + [ws.Range('A1:U1').Width])
        ult = 21
        while ult < 40 and ws.Cells(1, ult).Left + ws.Cells(1, ult).Width < larg_max:
            ult += 1
        faixa = ws.Range(ws.Cells(1, 1), ws.Cells(linhas, ult))
        # xlScreen+xlBitmap sai igualzinho a tela, mas EXIGE que a janela esteja
        # sendo desenhada de verdade -- com a sessao bloqueada, ou com outro
        # Excel segurando a area de transferencia, ele falha com "o metodo
        # CopyPicture falhou". xlPicture (metarquivo) nao depende da tela.
        erro = None
        for apar, fmt, nome in ((1, -4147, 'tela/bitmap'), (1, -4142, 'tela/metarquivo'),
                                (2, -4142, 'impressao/metarquivo')):
            try:
                faixa.CopyPicture(apar, fmt)
                print('copia:', nome)
                erro = None
                break
            except Exception as e:                # pode ser transitorio: espera e tenta o proximo
                erro = e
                time.sleep(1.0)
        if erro is not None:
            raise erro
        co = ws.ChartObjects().Add(0, 0, faixa.Width, faixa.Height)
        ch = co.Chart
        # o quadro vazio desenha eixos proprios; eles apareciam por baixo da
        # figura colada e sujavam a imagem
        try:
            ch.ChartArea.ClearContents()   # tira eixos e serie do quadro vazio
        except Exception:
            pass
        try:
            ch.ChartArea.Format.Line.Visible = 0
        except Exception:
            pass
        ch.Paste()
        time.sleep(0.5)
        # o quadro passa a ter o tamanho da FIGURA (e nao o contrario): assim
        # nao sobra fundo de grafico vazio por baixo dela
        fig = ch.Shapes(1)
        fig.Left, fig.Top = 0, 0
        co.Width, co.Height = fig.Width, fig.Height
        print('figura:', round(fig.Width), 'x', round(fig.Height),
              '| quadro:', round(co.Width), 'x', round(co.Height))
        ch.Export(os.path.abspath(saida))
        co.Delete()
        print('imagem:', os.path.abspath(saida), os.path.getsize(saida), 'bytes')
    finally:
        ex.fechar()


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]), sys.argv[2],
         int(sys.argv[3]) if len(sys.argv) > 3 else 60)
