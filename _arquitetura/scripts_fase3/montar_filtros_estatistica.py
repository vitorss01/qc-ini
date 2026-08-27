# -*- coding: utf-8 -*-
"""montar_filtros_estatistica.py - cabecalho de filtros da aba Estatistica

POR QUE PYTHON, E NAO POWERSHELL

Os rotulos tem acento. O Windows PowerShell 5.1 le .ps1 como ANSI, e foi
exatamente assim que a aba ficou com "MEs/Tri/Sem", "atE", "PerIodo" e "NIvel".
Python trata UTF-8 nativamente; o texto que se escreve aqui e o que chega na
celula.

O QUE ESTE SCRIPT FAZ

1. reescreve o cabecalho (linhas 1..12) com rotulos corretos e acentuados
2. cria as VALIDACOES DE DADOS que faltavam -- a causa real de "os filtros nao
   funcionam": o motor sempre respondeu, mas B3 tinha lista com origem VAZIA e
   Ano/Mes nao tinham validacao nenhuma. Sem menu, nao ha como interagir.
3. monta o bloco de ate 5 EXCLUSOES em 2 colunas x 3 linhas
4. cria o espelho contiguo das exclusoes e o nome Estat_Exclusoes
5. reescreve as formulas da tabela para a nova assinatura de EstatPeriodo

O ESPELHO CONTIGUO

A UDF recebe as exclusoes como UM intervalo Nx2. O layout escolhido (2 colunas
x 3 linhas) e bom para os olhos e NAO forma um bloco Nx2 continuo. Em vez de
espremer a tela para agradar a formula, um bloco oculto a direita espelha os 5
pares em 5x2 e e ele que a formula recebe. A tela serve o usuario; o espelho
serve o motor.

Uso:
    python montar_filtros_estatistica.py <caminho.xlsm>
"""
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'

# ---- geometria do novo cabecalho -------------------------------------------
L_TITULO = 1
L_SUB = 2
L_PER = 4          # Visao / Ano / Mes / Lote
L_CUSTOM = 5       # personalizado + janela calculada
L_EXTIT = 7        # titulo do bloco de exclusoes
L_EX0 = 8          # primeira linha de exclusao (8, 9, 10)
L_EFET = 11        # periodo efetivo
L_CAB = 13         # cabecalho da tabela
L_DADOS = 14       # primeira linha de dado

C_ESP0 = 22        # V: espelho contiguo das exclusoes (oculto)

MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
         'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']


def novo_excel():
    return w.DispatchEx('Excel.Application')


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = novo_excel()
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    try:
        wb.EnableAutoRecover = False
    except Exception:
        pass
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    xl.Calculation = -4135          # manual durante a montagem

    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        ws = None
        for s in wb.Worksheets:
            if s.Name.upper().startswith('ESTAT'):
                ws = s
        if ws is None:
            raise SystemExit('aba Estatistica nao encontrada')
        if ws.ProtectContents:
            try:
                ws.Unprotect(SENHA)
            except Exception:
                pass

        # ---- quantas linhas de dado existem hoje -------------------------
        # A tabela comeca na linha 7 no layout antigo; achamos o fim pela
        # coluna A, que traz =Analitos!$A$n.
        prim_antiga = 7
        ult_antiga = prim_antiga
        r = prim_antiga
        while r < 400:
            f = ws.Cells(r, 1).Formula
            if isinstance(f, str) and f.startswith('=Analitos!'):
                ult_antiga = r
                r += 1
            else:
                break
        n_dados = ult_antiga - prim_antiga + 1
        print('linhas de dado no layout antigo: %d (linhas %d..%d)'
              % (n_dados, prim_antiga, ult_antiga))
        if n_dados < 2:
            raise SystemExit('tabela nao localizada -- abortando sem alterar')

        # ---- insere as linhas que faltam para o novo cabecalho -----------
        desloc = L_CAB - 6          # cabecalho antigo estava na linha 6
        if desloc > 0:
            ws.Rows('3:%d' % (2 + desloc)).Insert(-4121)   # xlDown
            print('linhas inseridas no topo: %d' % desloc)

        prim = prim_antiga + desloc
        ult = ult_antiga + desloc

        # ---- limpa a area de cabecalho ------------------------------------
        ws.Range(ws.Cells(1, 1), ws.Cells(L_CAB - 1, 40)).ClearContents()
        for rr in range(1, L_CAB):
            try:
                ws.Rows(rr).UnMerge()
            except Exception:
                pass

        # ---- titulo -------------------------------------------------------
        ws.Cells(L_TITULO, 1).Value = 'ESTATÍSTICA DE DESEMPENHO ANALÍTICO'
        ws.Cells(L_TITULO, 1).Font.Bold = True
        ws.Cells(L_TITULO, 1).Font.Size = 14
        ws.Cells(L_SUB, 1).Value = (
            'Imprecisão (CV), inexatidão (Bias) e erro total (ET) confrontados '
            'com as especificações de qualidade vigentes no período.')
        ws.Cells(L_SUB, 1).Font.Italic = True

        def rotulo(r, c, txt, largura=None):
            cel = ws.Cells(r, c)
            cel.Value = txt
            cel.Font.Bold = True
            cel.HorizontalAlignment = -4152      # direita
            if largura:
                ws.Columns(c).ColumnWidth = largura

        def campo(r, c):
            cel = ws.Cells(r, c)
            cel.Interior.Color = 15921906        # amarelo claro: entrada
            cel.Borders.LineStyle = 1
            cel.HorizontalAlignment = -4108
            cel.Locked = False
            return cel

        # ---- linha do periodo ---------------------------------------------
        ws.Cells(L_PER - 1, 1).Value = 'PERÍODO ANALISADO'
        ws.Cells(L_PER - 1, 1).Font.Bold = True
        ws.Cells(L_PER - 1, 1).Font.Color = 10498160

        rotulo(L_PER, 1, 'Visão')
        campo(L_PER, 2)
        rotulo(L_PER, 3, 'Ano')
        campo(L_PER, 4)
        rotulo(L_PER, 5, 'Mês / Tri / Sem')
        campo(L_PER, 6)
        rotulo(L_PER, 7, 'Lote')
        campo(L_PER, 8)
        ws.Cells(L_PER, 9).Value = '(vazio = todos os lotes)'
        ws.Cells(L_PER, 9).Font.Italic = True
        ws.Cells(L_PER, 9).Font.Color = 8421504

        rotulo(L_CUSTOM, 1, 'Personalizado')
        campo(L_CUSTOM, 2)
        rotulo(L_CUSTOM, 3, 'até')
        campo(L_CUSTOM, 4)
        ws.Cells(L_CUSTOM, 5).Value = '(só vale com Visão = PERSONALIZADO)'
        ws.Cells(L_CUSTOM, 5).Font.Italic = True
        ws.Cells(L_CUSTOM, 5).Font.Color = 8421504

        # janela calculada -- saida do motor, nao entrada do usuario
        rotulo(L_CUSTOM, 7, 'Janela')
        ws.Cells(L_CUSTOM, 8).Formula = '=JanelaInicio($B$%d,$D$%d,$F$%d,$B$%d,$D$%d)' % (
            L_PER, L_PER, L_PER, L_CUSTOM, L_CUSTOM)
        ws.Cells(L_CUSTOM, 9).Value = 'a'
        ws.Cells(L_CUSTOM, 9).HorizontalAlignment = -4108
        ws.Cells(L_CUSTOM, 10).Formula = '=JanelaFim($B$%d,$D$%d,$F$%d,$B$%d,$D$%d)' % (
            L_PER, L_PER, L_PER, L_CUSTOM, L_CUSTOM)
        for c in (8, 10):
            ws.Cells(L_CUSTOM, c).NumberFormat = 'dd/mm/yyyy'
            ws.Cells(L_CUSTOM, c).Font.Bold = True

        # ---- bloco de exclusoes: 2 colunas x 3 linhas ---------------------
        ws.Cells(L_EXTIT, 1).Value = 'EXCLUIR PERÍODOS DO CÁLCULO'
        ws.Cells(L_EXTIT, 1).Font.Bold = True
        ws.Cells(L_EXTIT, 1).Font.Color = 10498160
        ws.Cells(L_EXTIT, 5).Value = ('Resultados dentro de qualquer intervalo preenchido '
                                      'ficam fora de n, Média, DP, CV, Bias e ET.')
        ws.Cells(L_EXTIT, 5).Font.Italic = True
        ws.Cells(L_EXTIT, 5).Font.Color = 8421504

        # 1,2,3 na esquerda (B:C) e 4,5 na direita (E:F)
        pos = [(L_EX0 + 0, 2), (L_EX0 + 1, 2), (L_EX0 + 2, 2),
               (L_EX0 + 0, 5), (L_EX0 + 1, 5)]
        for i, (lin, col) in enumerate(pos, start=1):
            rotulo(lin, col - 1, '%d.' % i)
            campo(lin, col)
            campo(lin, col + 1)
            ws.Cells(lin, col).NumberFormat = 'dd/mm/yyyy'
            ws.Cells(lin, col + 1).NumberFormat = 'dd/mm/yyyy'

        # ---- espelho contiguo (oculto) ------------------------------------
        for i, (lin, col) in enumerate(pos):
            a = ws.Cells(L_EX0 + i, C_ESP0)
            b = ws.Cells(L_EX0 + i, C_ESP0 + 1)
            # .Address e PROPRIEDADE no win32com (ja devolve absoluto);
            # chamar como metodo da "'str' object is not callable".
            ref_a = ws.Cells(lin, col).Address
            ref_b = ws.Cells(lin, col + 1).Address
            a.Formula = '=IF(%s="","",%s)' % (ref_a, ref_a)
            b.Formula = '=IF(%s="","",%s)' % (ref_b, ref_b)
        ws.Cells(L_EX0 - 1, C_ESP0).Value = 'espelho contíguo das exclusões (não editar)'
        ws.Columns(C_ESP0).Hidden = True
        ws.Columns(C_ESP0 + 1).Hidden = True

        for nome, ref in (
            ('Estat_Exclusoes', "='%s'!$%s$%d:$%s$%d" % (
                ws.Name, 'V', L_EX0, 'W', L_EX0 + 4)),
            ('Estat_Ini_Efetiva', "='%s'!$H$%d" % (ws.Name, L_CUSTOM)),
            ('Estat_Fim_Efetiva', "='%s'!$J$%d" % (ws.Name, L_CUSTOM)),
        ):
            try:
                wb.Names(nome).Delete()
            except Exception:
                pass
            wb.Names.Add(nome, ref)
        print('nomes: Estat_Exclusoes (V%d:W%d), Estat_Ini_Efetiva, Estat_Fim_Efetiva'
              % (L_EX0, L_EX0 + 4))

        # ---- periodo efetivo ----------------------------------------------
        rotulo(L_EFET, 1, 'Período efetivo')
        cel = ws.Cells(L_EFET, 2)
        # Formato TEXTO faz a formula ser guardada como string -- a celula
        # herdou "@" do layout antigo e exibia o proprio texto da formula.
        #
        # NumberFormat e LOCALIZADO no COM: num Excel pt-BR "General" e
        # recusado e o nome aceito e "Geral". Tentar os dois evita amarrar o
        # script ao idioma da maquina que o rodou.
        geral_ok = False
        for nome_fmt in ('General', 'Geral'):
            try:
                cel.NumberFormat = nome_fmt
                geral_ok = True
                break
            except Exception:
                pass
        if not geral_ok:
            raise SystemExit('nao foi possivel tirar a celula do formato Texto')
        cel.Formula = '=PeriodoEfetivo(Estat_Ini_Efetiva,Estat_Fim_Efetiva,Estat_Exclusoes)'
        cel.Font.Bold = True
        ws.Range(ws.Cells(L_EFET, 2), ws.Cells(L_EFET, 12)).Merge()

        # ---- validacoes: a causa real do "nao funciona" -------------------
        def validar(rng, lista):
            rng.Validation.Delete()
            rng.Validation.Add(3, 1, 1, ','.join(lista))   # xlValidateList
            rng.Validation.IgnoreBlank = True
            rng.Validation.InCellDropdown = True

        validar(ws.Cells(L_PER, 2), ['ANUAL', 'SEMESTRAL', 'TRIMESTRAL',
                                     'MENSAL', 'PERSONALIZADO'])
        validar(ws.Cells(L_PER, 6), [str(i) for i in range(1, 13)])

        # anos: do banco, para o dropdown so oferecer o que existe
        db = wb.Worksheets('DB_Resultados')
        ultdb = db.Cells(db.Rows.Count, 1).End(-4162).Row
        anos = set()
        if ultdb >= 4:
            vals = db.Range(db.Cells(4, 2), db.Cells(ultdb, 2)).Value
            for linha in vals:
                v = linha[0]
                if hasattr(v, 'year'):
                    anos.add(v.year)
        if not anos:
            import datetime
            anos = {datetime.date.today().year}
        anos = sorted(anos)
        validar(ws.Cells(L_PER, 4), [str(a) for a in anos])
        print('validações: Visão(5), Mês/Tri/Sem(1-12), Ano(%s)'
              % ','.join(str(a) for a in anos))

        # valores iniciais coerentes
        if not str(ws.Cells(L_PER, 2).Value or '').strip():
            ws.Cells(L_PER, 2).Value = 'ANUAL'
        ws.Cells(L_PER, 4).Value = anos[-1]
        if not str(ws.Cells(L_PER, 6).Value or '').strip():
            ws.Cells(L_PER, 6).Value = 1

        # ---- formulas da tabela: nova assinatura --------------------------
        # O ARGUMENTO DE LOTE TEM DE SER REESCRITO, NAO PROCURADO.
        #
        # Ao inserir linhas no topo, o Excel DESLOCA sozinho as referencias das
        # formulas: o antigo $B$4 (lote) virou $B$11 -- que no layout novo e a
        # celula "Periodo efetivo". Procurar por "$B$4" nao acha nada, a formula
        # fica apontando para texto, nenhum lote casa e a tabela inteira zera.
        # Por isso o ultimo argumento e reescrito por posicao, e nao por valor.
        # Busca de texto, nao regex: o alvo tem "$" e parenteses, que em regex
        # exigem escape duplo e ja produziram erro de sintaxe aqui. Recortar
        # entre "Estat_Exclusoes," e o ")" seguinte e mais simples e mais facil
        # de conferir de olho.
        trocadas = 0
        alvo_lote = '$H$%d' % L_PER
        marca = 'Estat_Exclusoes,'
        for r in range(prim, ult + 1):
            for c in range(1, 40):
                f = ws.Cells(r, c).Formula
                if not isinstance(f, str) or 'EstatPeriodo(' not in f:
                    continue
                novo = f.replace(
                    'Data_Inicio_Exclusao,Data_Fim_Exclusao', 'Estat_Exclusoes')
                i = novo.find(marca)
                if i >= 0:
                    k = i + len(marca)
                    j = novo.find(')', k)
                    if j > k:
                        # DB_Carimbo entra como 8o argumento (ADR-049). Ele nao
                        # participa da conta: serve para a celula DECLARAR que
                        # depende do banco, senao so um rebuild completo
                        # atualizaria a Estatistica -- e o fluxo real do
                        # sistema chama Application.Calculate, que nao faz
                        # rebuild.
                        novo = novo[:k] + alvo_lote + ',DB_Carimbo' + novo[j:]
                if novo != f:
                    ws.Cells(r, c).Formula = novo
                    trocadas += 1
        print('fórmulas migradas (exclusões + lote -> %s): %d' % (alvo_lote, trocadas))

        for nome in ('Data_Inicio_Exclusao', 'Data_Fim_Exclusao'):
            try:
                wb.Names(nome).Delete()
            except Exception:
                pass

        ws.Columns(1).ColumnWidth = 16
        ws.Rows(3).RowHeight = 6
        ws.Rows(6).RowHeight = 6
        ws.Rows(12).RowHeight = 6

        xl.Calculation = -4105
        xl.CalculateFullRebuild()
        if estrutura and not wb.ProtectStructure:
            wb.Protect(SENHA, True, False)
        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            xl.Calculation = -4105
        except Exception:
            pass
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    if len(sys.argv) < 2:
        raise SystemExit('uso: python montar_filtros_estatistica.py <arquivo.xlsm>')
    main(sys.argv[1])
