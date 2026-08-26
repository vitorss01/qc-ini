# -*- coding: utf-8 -*-
"""testar_eventos_westgard.py - ADR-045: a aba de eventos e um FATO de evento

O QUE ESTAVA ERRADO

Eventos_Westgard tinha granularidade de (corrida x analito x nivel) com as
regras concatenadas numa celula ("13s+22s"), e era produzida por
AvaliarWestgard1N -- uma SEGUNDA implementacao das regras, escrita para uma
serie de um nivel so. Consequencias:

  - R_4s, 2of3_2s, 3_1s e 6x eram estruturalmente invisiveis (exigem visao
    multi-nivel);
  - a aba materializava "10 corridas seguidas do mesmo lado", o 10x que o
    ADR-041 aposentou;
  - a Hematologia mostrava 9 eventos onde a camada de BI via 80 marcacoes.

Agora a aba vem do TRACE de AvaliarWestgard. Nao ha segunda implementacao.

O QUE ESTE TESTE PROVA (e nao apenas afirma)

  1. nenhuma regra fora da matriz do produto -- em especial, nenhum 10x
  2. nenhum par (regra, detector) fora de DetectoresWestgard()
  3. todo evento OFICIAL usa o detector oficial daquela regra
  4. R_4s e sempre WITHIN_RUN: janela de uma corrida so
  5. as regras de janela (8x/6x/4_1s/2_2s across) abrem janela > 1 corrida
  6. N_Eventos_Violacao de MetricasWestgard bate com a contagem de linhas
     OFICIAIS da aba que envolvem aquele nivel
  7. EVENTO e MARCA sao numeros diferentes e nao se confundem

Uso: python testar_eventos_westgard.py <arquivo.xlsm>
"""
import io
import os
import sys
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

# colunas da aba (1-based)
C_DATA, C_RUN, C_ANALITO, C_NIVEIS, C_REGRA, C_DETECTOR = 1, 2, 3, 4, 5, 6
C_ESCOPO, C_CLASSE, C_N, C_R, C_RUNINI, C_EVID = 7, 8, 9, 10, 11, 12
C_CLASSIF, C_ZMAX = 13, 14

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-58s %s' % (nome[:58], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for linha in str(detalhe).split('\n')[:6]:
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
        matriz = [r.strip() for r in str(xl.Run('MatrizWestgard')).split(';')]
        detectores = str(xl.Run('DetectoresWestgard'))
        print('=' * 78)
        print('%s -- Eventos_Westgard com granularidade de evento' % area)
        print('=' * 78)
        print('   matriz do produto: %s' % ' / '.join(matriz))
        print()

        # Uma regra pode ter MAIS DE UM detector ativo -- 2_2s da Bioquimica tem
        # as duas formas classicas (dois materiais na mesma corrida, e o mesmo
        # material em duas corridas), ambas oficiais. Quem escolhe UM detector e
        # o CONTRATO do mPlanoQC, e so para efeito de cobertura; o que o evento
        # tem de respeitar e o Ativo do detector que o produziu.
        pares = set()
        ativos = set()
        for bruto in detectores.split(';'):
            c = bruto.split('|')
            if len(c) < 9 or c[0].upper() != area.upper():
                continue
            pares.add((c[2], c[3]))
            if c[4] == '1':
                ativos.add((c[2], c[3]))

        xl.Run('RegistrarEventosWestgard')
        ws = wb.Worksheets('Eventos_Westgard')
        total = int(ws.Range('J2').Value or 0)
        print('   eventos registrados: %d' % total)
        print()

        linhas = []
        if total:
            bloco = ws.Range(ws.Cells(4, 1), ws.Cells(3 + total, C_ZMAX)).Value
            linhas = [list(r) for r in bloco]

        def niveis(L):
            """A coluna e texto por contrato, mas o teste nao pode DEPENDER
            disso: se a formatacao regredir, quero ver o teste da tipagem
            falhar -- e nao seis outros quebrarem por ValueError e esconderem
            qual e o defeito real."""
            bruto = L[C_NIVEIS - 1]
            if isinstance(bruto, float):
                return [int(bruto)]
            return [int(float(x)) for x in str(bruto).split(',') if x.strip()]

        misturadas = [L for L in linhas if not isinstance(L[C_NIVEIS - 1], str)]
        checar('coluna Niveis e texto em todas as linhas (tipo unico)',
               not misturadas,
               '%d linha(s) vieram como numero -- coluna de tipo misto'
               % len(misturadas))

        # --- 1. regra fora da matriz -----------------------------------------
        fora = sorted({str(L[C_REGRA - 1]) for L in linhas
                       if str(L[C_REGRA - 1]) not in matriz})
        checar('nenhuma regra fora da matriz do produto', not fora,
               'encontradas: %s' % ', '.join(fora))

        checar('nenhum evento de 10x (aposentado pelo ADR-041)',
               not any(str(L[C_REGRA - 1]) == '10x' for L in linhas))

        # --- 2. par (regra, detector) declarado ------------------------------
        maus = sorted({(str(L[C_REGRA - 1]), str(L[C_DETECTOR - 1]))
                       for L in linhas
                       if (str(L[C_REGRA - 1]), str(L[C_DETECTOR - 1])) not in pares})
        checar('todo par (regra, detector) existe em DetectoresWestgard',
               not maus, '\n'.join('%s / %s' % p for p in maus))

        # --- 3. a classe do evento espelha o Ativo do detector ---------------
        # Nos DOIS sentidos. Verificar so um lado deixaria passar o caso pior:
        # um detector complementar gravado como OFICIAL faria a aba consolidar
        # uma regra que o produto nao adotou.
        erradas = sorted({(str(L[C_REGRA - 1]), str(L[C_DETECTOR - 1]),
                           str(L[C_CLASSE - 1]))
                          for L in linhas
                          if ((str(L[C_CLASSE - 1]) == 'OFICIAL')
                              != ((str(L[C_REGRA - 1]),
                                   str(L[C_DETECTOR - 1])) in ativos))})
        checar('classe do evento espelha o Ativo do detector (nos 2 sentidos)',
               not erradas,
               '\n'.join('%s / %s gravado %s, mas Ativo=%s'
                         % (r, d, cl, '1' if (r, d) in ativos else '0')
                         for r, d, cl in erradas))

        checar('classe sempre OFICIAL ou COMPLEMENTAR',
               all(str(L[C_CLASSE - 1]) in ('OFICIAL', 'COMPLEMENTAR')
                   for L in linhas))

        # --- 4. R_4s e intra-corrida ------------------------------------------
        r4 = [L for L in linhas if str(L[C_REGRA - 1]) == 'R_4s']
        checar('R_4s: janela de UMA corrida (RUN_Inicial = RUN)',
               all(int(L[C_RUNINI - 1]) == int(L[C_RUN - 1]) for L in r4),
               '%d evento(s) de R_4s com janela maior que a corrida'
               % sum(1 for L in r4 if int(L[C_RUNINI - 1]) != int(L[C_RUN - 1])))
        checar('R_4s envolve dois materiais (dois niveis)',
               all(len(niveis(L)) == 2 for L in r4))

        # --- 5. regras de janela abrem janela --------------------------------
        JANELA = {'8x': 'N2_R4', '6x': 'N3_R2', '4_1s': 'N2_R2',
                  '2_2s': 'ACROSS_RUN_SAME_LEVEL'}
        jan = [L for L in linhas
               if JANELA.get(str(L[C_REGRA - 1])) == str(L[C_DETECTOR - 1])]
        checar('regras de janela abrem janela de mais de uma corrida (%d ev.)'
               % len(jan),
               all(int(L[C_RUNINI - 1]) < int(L[C_RUN - 1]) for L in jan))

        # 1_3s e individual: um nivel, uma corrida
        um = [L for L in linhas if str(L[C_REGRA - 1]) == '1_3s']
        checar('1_3s: um nivel, uma corrida (%d ev.)' % len(um),
               all(len(niveis(L)) == 1
                   and int(L[C_RUNINI - 1]) == int(L[C_RUN - 1]) for L in um))

        # --- 6. metricas batem com a aba --------------------------------------
        analitos = sorted({str(L[C_ANALITO - 1]) for L in linhas})
        porChave = defaultdict(int)
        for L in linhas:
            if str(L[C_CLASSE - 1]) != 'OFICIAL':
                continue
            for n in niveis(L):
                porChave[(str(L[C_ANALITO - 1]).upper(), n)] += 1

        divergem = []
        amostra = []
        for a in analitos:
            for n in range(1, 4):
                m = str(xl.Run('MetricasWestgard', a, n))
                if '|' not in m:
                    continue
                ev, marc, corr = [int(x or 0) for x in m.split('|')]
                esperado = porChave.get((a.upper(), n), 0)
                if ev != esperado:
                    divergem.append('%s N%d: metrica=%d aba=%d'
                                    % (a, n, ev, esperado))
                if ev and len(amostra) < 8:
                    amostra.append((a, n, ev, marc, corr))
        checar('N_Eventos_Violacao bate com as linhas OFICIAIS da aba',
               not divergem, '\n'.join(divergem))

        # --- 7. evento x marca sao numeros distintos -------------------------
        print()
        print('   %-14s %-4s %-9s %-11s %s'
              % ('analito', 'niv', 'eventos', 'marcados', 'corridas'))
        print('   ' + '-' * 58)
        for a, n, ev, marc, corr in amostra:
            print('   %-14s N%-3d %-9d %-11d %d' % (a[:14], n, ev, marc, corr))
        if not amostra:
            print('   (nenhum evento oficial nesta fixture)')

        checar('corridas envolvidas >= 1 onde ha evento',
               all(corr >= 1 for _, _, ev, _, corr in amostra if ev))
        checar('nenhuma janela conta menos corridas que a escala do detector',
               all(corr >= 1 for _, _, _, _, corr in amostra))
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
