# -*- coding: utf-8 -*-
"""qa_painel_westgard.py - o que o Painel EXIBE depois do ADR-045

POR QUE LER CELULA E NAO TIRAR PRINT

O que interessa aqui e o TEXTO que chega ao analista, e isso a celula entrega
com precisao. Captura de tela nesta maquina ja trouxe janela de sistema de
paciente junto; o risco nao se paga por uma informacao que a celula da melhor.

O QUE VERIFICA

  1. o resumo de Westgard do Painel usa o formato "N ev / M res", separando
     EVENTO de RESULTADO MARCADO (era um "Nx" ambiguo)
  2. os numeros exibidos batem com MetricasWestgard -- ou seja, a tela nao tem
     uma terceira contagem propria
  3. a regra exibida pertence a matriz do produto (nada de 10x na tela)
  4. onde nao ha violacao, o Painel mostra o travessao e nao "0 ev / 0 res"
  5. a Estatistica nao ficou com celula em erro (#VALOR!, #NOME?) depois da
     mudanca

Uso: python qa_painel_westgard.py <arquivo.xlsm>
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

ABA_ENG = 'Eng_Saida'
ABA_EST = 'Estat' + chr(0x00ED) + 'stica'
TRAVESSAO = chr(0x2014)          # em dash, o que AtualizarPainelEng escreve

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-58s %s' % (nome[:58], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for linha in str(detalhe).split('\n')[:8]:
            print('        %s' % linha)


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, True)
    try:
        area = str(xl.Run('AreaDoProduto'))
        nlv = 3 if area.upper() == 'HEMATOLOGIA' else 2
        matriz = [r.strip() for r in str(xl.Run('MatrizWestgard')).split(';')]
        print('=' * 78)
        print('%s -- o que o Painel exibe (ADR-045)' % area)
        print('=' * 78)
        print('   matriz do produto: %s' % ' / '.join(matriz))
        print()

        # O Painel mostra UM analito de cada vez, escolhido pelo Spinner e
        # guardado no nome selAnalito. E dele que o resumo por nivel fala.
        alvo = str(wb.Names('selAnalito').RefersToRange.Value or '').strip()
        print('   analito no Painel: %s' % (alvo or '(nao definido)'))

        # roda o motor inteiro: e assim que a tela e alimentada
        xl.Run('AtualizarEstatistica')
        xl.Run('AtualizarPainelEng')

        eng = wb.Worksheets(ABA_ENG)

        # ANCORAR NO CABECALHO, e nao no conteudo da coluna U.
        #
        # Procurar pelo conteudo parecia mais robusto e nao era: quando o
        # analito NAO tem violacao, AtualizarPainelEng poe o travessao na
        # coluna S e deixa a U VAZIA. O bloco existe, mas fica invisivel para
        # quem procura na U -- e o teste morria com "nao achei o bloco" num
        # produto perfeitamente saudavel (Glicose, sem violacao).
        #
        # O cabecalho ("ultViolacao" em S, "historico" em U) existe sempre,
        # independente de haver violacao.
        cab = 0
        for r in range(1, 400):
            s19 = str(eng.Cells(r, 19).Value or '').strip().lower()
            s21 = str(eng.Cells(r, 21).Value or '').strip().lower()
            if s19.startswith('ultviolacao') or s21.startswith('historico'):
                cab = r
                break
        if not cab:
            raise SystemExit('nao achei o cabecalho do bloco de Westgard em %s'
                             % ABA_ENG)
        achou = []
        for k in range(1, nlv + 1):
            r = cab + k
            v = eng.Cells(r, 21).Value
            achou.append((r, '' if v is None else str(v).strip(),
                          eng.Cells(r, 19).Value))
        print('   resumo por nivel (Eng_Saida!S e !U):')
        for r, s, regra in achou:
            print('     linha %-4d %-26s %s' % (r, str(regra), s))
        print()

        # Nivel COM violacao = coluna U preenchida. Sem violacao, U fica vazia e
        # o travessao aparece na S. Filtrar por "nao tem travessao na U" daria
        # verdadeiro para a celula vazia e trataria nivel limpo como se tivesse
        # dado -- justamente o caso que este teste precisa distinguir.
        comDado = [(r, s, g) for r, s, g in achou
                   if s and TRAVESSAO not in s]
        semDado = [(r, s, g) for r, s, g in achou if not s]

        # --- 1. formato novo --------------------------------------------------
        checar('resumo usa o formato "N ev / M res" (evento x marca)',
               all(re.search(r'\d+ ev / \d+ res', s) for _, s, _ in comDado)
               or not comDado,
               'linhas fora do formato: %s'
               % [s for _, s, _ in comDado if not re.search(r'\d+ ev / \d+ res', s)])

        checar('nenhum resumo no formato antigo e ambiguo ("Nx  |")',
               not any(re.search(r'^\d+x\s+\|', s) for _, s, _ in achou))

        # --- 2. a tela nao tem contagem propria -------------------------------
        # O resumo do Painel e uma LEITURA do motor. Se os numeros divergirem,
        # existe uma terceira contagem em algum lugar -- que e o defeito que
        # esta serie de ADRs vem eliminando.
        divergem = []
        if alvo:
            for idx, (r, s, _) in enumerate(achou, start=1):
                m = re.search(r'(\d+) ev / (\d+) res', s)
                if not m:
                    continue
                ev_tela, res_tela = int(m.group(1)), int(m.group(2))
                metr = str(xl.Run('MetricasWestgard', alvo, idx))
                ev_mot, res_mot = [int(x or 0) for x in metr.split('|')[:2]]
                if (ev_tela, res_tela) != (ev_mot, res_mot):
                    divergem.append('N%d tela=%d/%d motor=%d/%d'
                                    % (idx, ev_tela, res_tela, ev_mot, res_mot))
            checar('numeros da tela = MetricasWestgard (sem terceira contagem)',
                   not divergem, '\n'.join(divergem))
        else:
            print('   %-58s %s'
                  % ('analito do Painel nao identificado; item 2 nao rodou',
                     'SKIP'))

        # --- 3. regra exibida pertence a matriz -------------------------------
        # O Painel escreve "3_1s <sep> RUN 25", com Chr$(183) de separador.
        # Ele entra na quebra por espaco e vira um "token" -- que nao e regra
        # nenhuma. Tratar separador como conteudo fazia o teste reprovar uma
        # tela correta.
        SEPARADORES = {chr(0x00B7), chr(0x2014), chr(0x2013), '-', ':'}
        exibidas = set()
        for _, _, regra in comDado:
            texto = str(regra or '')
            for tok in re.split(r'[+\s|]+', texto):
                tok = tok.strip()
                if not tok or tok in SEPARADORES:
                    continue
                if not tok.isdigit() and tok.upper() != 'RUN':
                    exibidas.add(tok)
        forasteiras = sorted(x for x in exibidas if x not in matriz)
        checar('regra exibida pertence a matriz do produto', not forasteiras,
               'fora da matriz: %s' % ', '.join(forasteiras))
        checar('nenhum 10x na tela', '10x' not in exibidas)

        # --- 4. sem violacao mostra travessao ---------------------------------
        # O travessao vive na coluna S (ultima violacao), nao na U.
        checar('nivel sem violacao mostra travessao, nao "0 ev / 0 res"',
               not any('0 ev / 0 res' in s for _, s, _ in achou)
               and all(TRAVESSAO in str(g or '') for _, _, g in semDado),
               'niveis sem violacao com coluna S inesperada: %s'
               % [str(g) for _, _, g in semDado])

        # --- 5. Estatistica sem celula em erro --------------------------------
        est = wb.Worksheets(ABA_EST)
        erros = []
        for r in range(1, 120):
            for c in range(1, 26):
                t = str(est.Cells(r, c).Text or '')
                if t.startswith('#') and len(t) > 1:
                    erros.append('%s!%s%d = %s'
                                 % (ABA_EST, chr(64 + c) if c <= 26 else '?', r, t))
        checar('Estatistica sem celula em erro', not erros, '\n'.join(erros[:8]))
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
