# -*- coding: utf-8 -*-
"""corrigir_bias_painel.py - o Painel lia o Bias da coluna errada

O DEFEITO

Painel!G7 e G8 (rotulados "Bias %") liam Estatistica!$J$14:$J$137. Na
Estatistica v2 a coluna J e "Status CV" -- TEXTO. A coluna do bias e a K,
"Bias % (alvo do lote)".

O guarda IF(ISNUMBER(G7), G7, 0) fazia o texto virar ZERO em silencio:

    ET %  = CV*1,65 + bias   ->  CV*1,65 + 0   ->  SUBESTIMADO
    Sigma = (ETp - bias)/CV  ->  (ETp - 0)/CV  ->  SUPERESTIMADO

O metodo aparentava desempenho melhor do que tem. Num sistema de controle de
qualidade esse e o pior modo de falha: nao o erro visivel, e o numero plausivel
e errado, que ninguem confere porque parece certo.

A causa estrutural: a Estatistica foi reconstruida (v2) com outro conjunto de
colunas e quem a consumia continuou apontando para as coordenadas antigas.
Mesma familia do redirecionar_estatistica.ps1 que ja corrigimos hoje -- o
consumidor nao sabe que a fonte mudou.

Localiza a coluna PELO ROTULO, e nao por letra: se a Estatistica mudar de novo,
este script continua achando o bias.

Uso: python corrigir_bias_painel.py <caminho.xlsm>
"""
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'


def col_letra(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura')
    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        es = None
        for s in wb.Worksheets:
            if s.Name.upper().startswith('ESTAT'):
                es = s
        pa = wb.Worksheets('Painel')
        for sh in (es, pa):
            for t in (lambda: sh.Unprotect(SENHA), lambda: sh.Unprotect()):
                try:
                    t()
                    break
                except Exception:
                    pass

        # ---- acha o cabecalho e a coluna do bias PELO ROTULO -------------
        cabLin = 0
        for r in range(1, 40):
            if str(es.Cells(r, 1).Value or '').strip().upper() == 'ANALITO':
                cabLin = r
                break
        if cabLin == 0:
            raise SystemExit('cabecalho da Estatistica nao localizado')

        colBias = 0
        rotulos = {}
        for c in range(1, 30):
            v = str(es.Cells(cabLin, c).Value or '').strip()
            if v:
                rotulos[c] = v
            if v.upper().startswith('BIAS'):
                colBias = c
        if colBias == 0:
            raise SystemExit('coluna de Bias nao localizada pelo rotulo')
        print('cabecalho na linha %d; Bias na coluna %d (%s) = %r'
              % (cabLin, colBias, col_letra(colBias), rotulos[colBias]))

        prim = cabLin + 1
        ult = prim
        r = prim
        while r < 400:
            f = es.Cells(r, 1).Formula
            if isinstance(f, str) and f.startswith('=Analitos!'):
                ult = r
                r += 1
            else:
                break
        print('bloco de dados: %d..%d' % (prim, ult))

        alvo = "Estatística!$%s$%d:$%s$%d" % (col_letra(colBias), prim,
                                              col_letra(colBias), ult)

        # ---- reescreve as celulas de bias do Painel ----------------------
        trocadas = []
        for lin in range(1, 40):
            for c in range(1, 30):
                f = pa.Cells(lin, c).Formula
                if not isinstance(f, str) or 'Estatística!$' not in f:
                    continue
                if 'MATCH(selAnalito' not in f and 'CORRESP(selAnalito' not in f:
                    continue
                novo = f
                import re
                # troca SO a faixa de valor (INDEX), preservando a de busca
                novo = re.sub(r'Estatística!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+',
                              lambda m, box=[0]: (box.__setitem__(0, box[0] + 1) or
                                                  (alvo if box[0] == 1 else m.group(0))),
                              novo)
                if novo != f:
                    pa.Cells(lin, c).Formula = novo
                    trocadas.append('%s%d' % (col_letra(c), lin))
        if not trocadas:
            raise SystemExit('nenhuma formula do Painel consumia a Estatistica')
        print('Painel: %s -> faixa do Bias (%s)' % (', '.join(trocadas), alvo))

        xl.CalculateFullRebuild()
        for lin in (7, 8):
            print('   G%d=%r  H%d(ET)=%r  I%d(Sigma)=%r'
                  % (lin, pa.Cells(lin, 7).Value, lin, pa.Cells(lin, 8).Value,
                     lin, pa.Cells(lin, 9).Value))
        v = pa.Cells(7, 7).Value
        if isinstance(v, str) and v.strip() != '':
            raise SystemExit('Bias continua vindo como TEXTO (%r) -- nao salvo' % v)

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
