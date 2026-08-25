# -*- coding: utf-8 -*-
"""testar_cobertura_westgard.py - ADR-044: a cobertura reprova o que deve

O QUE ESTE TESTE PROVA

CoberturaWestgard so pode responder TOTAL quando o detector OFICIAL de cada
regra do plano bate em nome, N, R e escopo com o contrato. Provar isso exige
mostrar que ela REPROVA um motor errado -- e um motor errado nao existe na
pasta.

Por isso a comparacao foi extraida para ConferirContrato(regras, tabela): o
teste passa uma tabela MUTADA e verifica a recusa. Sem o parametro, estes
casos seriam encenacao -- nao ha como alterar uma Const em execucao.

CASOS

  A. 8x com detector longitudinal em vez de N2_R4 ......... reprova
  B. 6x com detector longitudinal em vez de N3_R2 ......... reprova
  C. R_4s com escopo ACROSS_RUN em vez de WITHIN_RUN ...... reprova
  D. 3_1s com detector SAME_LEVEL_R3 em vez de N3_R1 ...... reprova
  E. R trocado (8x com R=8 em vez de 4) ................... reprova
  F. detector oficial desativado .......................... reprova
  G. tabela canonica intacta .............................. aprova

Uso: python testar_cobertura_westgard.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

resultados = []


def anota(nome, esperado_reprova, resposta):
    reprovou = bool(str(resposta).strip())
    ok = (reprovou == esperado_reprova)
    resultados.append((nome, esperado_reprova, resposta, ok))
    print('   %-52s %-8s %s'
          % (nome[:52], 'reprova' if esperado_reprova else 'aprova',
             'PASS' if ok else 'FAIL'))
    if not ok:
        print('        resposta do motor: %r' % (str(resposta)[:90],))
    elif reprovou:
        print('        -> %s' % str(resposta)[:90])


def mutar(tabela, area, regra, det_de, **campos):
    """Troca campos de UMA linha da tabela.
    Colunas: 0 Area | 1 Niveis | 2 Regra | 3 Detector | 4 Ativo | 5 N | 6 R |
             7 Escopo | 8 Limiar"""
    idx = {'detector': 3, 'ativo': 4, 'n': 5, 'r': 6, 'escopo': 7, 'limiar': 8}
    fora = []
    for linha in tabela.split(';'):
        c = linha.split('|')
        if (len(c) > 8 and c[0].upper() == area.upper()
                and c[2].upper() == regra.upper()
                and c[3].upper() == det_de.upper()):
            for k, v in campos.items():
                c[idx[k]] = str(v)
        fora.append('|'.join(c))
    return ';'.join(fora)


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
        canonica = str(xl.Run('DetectoresWestgard'))
        matriz = str(xl.Run('MatrizWestgard')).replace(';', ' / ')
        print('=' * 78)
        print('%s -- cobertura estrutural' % area)
        print('=' * 78)
        print('   plano de referencia: %s' % matriz)
        print()

        # G. canonica intacta aprova
        anota('tabela canonica intacta', False,
              xl.Run('ConferirContrato', matriz, canonica))

        if area == 'BIOQUIMICA':
            # A. 8x oficial vira o longitudinal
            t = mutar(canonica, area, '8x', 'N2_R4', ativo=0)
            t = mutar(t, area, '8x', 'SAME_LEVEL_R8', ativo=1)
            anota('8x oficial = SAME_LEVEL_R8 (era N2_R4)', True,
                  xl.Run('ConferirContrato', matriz, t))
            # E. R errado
            t = mutar(canonica, area, '8x', 'N2_R4', r=8)
            anota('8x com R=8 (esperado R=4)', True,
                  xl.Run('ConferirContrato', matriz, t))
            # F. oficial desativado, sem substituto
            t = mutar(canonica, area, '4_1s', 'N2_R2', ativo=0)
            anota('4_1s sem detector oficial ativo', True,
                  xl.Run('ConferirContrato', matriz, t))
        else:
            # B. 6x oficial vira o longitudinal
            t = mutar(canonica, area, '6x', 'N3_R2', ativo=0)
            t = mutar(t, area, '6x', 'SAME_LEVEL_R6', ativo=1)
            anota('6x oficial = SAME_LEVEL_R6 (era N3_R2)', True,
                  xl.Run('ConferirContrato', matriz, t))
            # D. 3_1s oficial vira o longitudinal
            t = mutar(canonica, area, '3_1s', 'N3_R1', ativo=0)
            t = mutar(t, area, '3_1s', 'SAME_LEVEL_R3', ativo=1)
            anota('3_1s oficial = SAME_LEVEL_R3 (era N3_R1)', True,
                  xl.Run('ConferirContrato', matriz, t))
            # E. N errado
            t = mutar(canonica, area, '6x', 'N3_R2', n=2)
            anota('6x com N=2 (esperado N=3)', True,
                  xl.Run('ConferirContrato', matriz, t))

        # C. escopo errado -- vale nos dois
        t = mutar(canonica, area, 'R_4s', 'WITHIN_RUN', escopo='ACROSS_RUN')
        anota('R_4s com escopo ACROSS_RUN', True,
              xl.Run('ConferirContrato', matriz, t))

        # a cobertura completa, com a tabela real
        print()
        print('   CoberturaWestgard(3,5) = %s' % xl.Run('CoberturaWestgard', 3.5))
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    falhas = [r for r in resultados if not r[3]]
    print('TOTAL: %d PASS, %d FAIL'
          % (len(resultados) - len(falhas), len(falhas)))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
