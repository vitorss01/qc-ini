# -*- coding: utf-8 -*-
"""testar_listobjects_dados.py - as duas tabelas que o ETL vai ler pelo nome

ESCOPO: so as interfaces criadas pelo ADR-051. Nao repete a bateria da Fase 1.

  A. tblEventos_Westgard existe
  B. uma linha por evento (linhas da tabela == total em J2)
  C. reconstruir o historico nao destroi a tabela, e ela acompanha o tamanho
  D. sem linha fantasma: nenhuma linha da tabela sem chave
  E. tblEQA_Base existe
  F. cabecalhos integros -- inclusive os dois acentos que estavam errados
  G. as duas sao alcancaveis PELO NOME, como o Power Query as le

O item C e o que importa de verdade: a rotina escreve num Range, nao na
tabela. Se o ListObject nao acompanhar, o ETL passa a ler sobra (linha vazia)
ou falta (evento real fora da tabela) -- e a aba mostraria tudo, entao o erro
nao apareceria olhando a planilha.

Uso: python testar_listobjects_dados.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

EV_TAB = 'tblEventos_Westgard'
EQ_TAB = 'tblEQA_Base'

CAB_EV = ['Data', 'RUN', 'Analito', 'N' + chr(0x00ED) + 'veis', 'Regra',
          'Detector', 'Escopo', 'Classe', 'N', 'R', 'RUN_Inicial',
          'Evid' + chr(0x00EA) + 'ncia', 'Classifica' + chr(0x00E7) +
          chr(0x00E3) + 'o', 'Z_Max']

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-54s %s' % (nome[:54], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for l in str(detalhe).split('\n')[:5]:
            print('        %s' % l)


def tabela(wb, nome):
    """Procura em todas as abas -- e assim que o Power Query acha."""
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            if lo.Name == nome:
                return ws, lo
    return None, None


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, False)
    try:
        area = str(xl.Run('AreaDoProduto'))
        print('=' * 74)
        print('LISTOBJECTS DE DADOS -- %s' % area)
        print('=' * 74)
        print()

        # --- A / G. existe e e alcancavel pelo nome -----------------------
        wsEv, loEv = tabela(wb, EV_TAB)
        checar('A. %s existe' % EV_TAB, loEv is not None)
        wsEq, loEq = tabela(wb, EQ_TAB)
        checar('E. %s existe' % EQ_TAB, loEq is not None)
        if loEv is None or loEq is None:
            print('\nTOTAL: %d FAIL' % len(falhas))
            return 1

        for nome, lo, ws in ((EV_TAB, loEv, wsEv), (EQ_TAB, loEq, wsEq)):
            n = lo.ListRows.Count
            print('   %-22s %-14s %d linha(s) x %d coluna(s)  aba %s'
                  % (nome, lo.Range.Address, n, lo.ListColumns.Count, ws.Name))

        # G: o Power Query resolve o nome sem saber a aba. Range("nome")
        # emula exatamente isso.
        alcanca = []
        for nome in (EV_TAB, EQ_TAB):
            try:
                wb.Application.Range(nome).Rows.Count
            except Exception as e:
                alcanca.append('%s: %s' % (nome, e))
        checar('G. as duas resolvem PELO NOME, sem citar a aba',
               not alcanca, '\n'.join(alcanca))

        # --- B. uma linha por evento --------------------------------------
        total = int(wsEv.Range('J2').Value or 0)
        checar('B. linhas da tabela == total de eventos (J2 = %d)' % total,
               loEv.ListRows.Count == total or (total == 0 and loEv.ListRows.Count == 1),
               'tabela tem %d' % loEv.ListRows.Count)

        # --- D. sem linha fantasma ----------------------------------------
        # Chave do evento: Analito e Regra. Linha sem os dois e sobra de um
        # historico anterior que a tabela nao encolheu.
        vazias = 0
        if loEv.ListRows.Count and total:
            dados = loEv.DataBodyRange.Value
            for r in dados:
                if not str(r[2] or '').strip() or not str(r[4] or '').strip():
                    vazias += 1
        checar('D. nenhuma linha da tabela sem Analito/Regra', vazias == 0,
               '%d linha(s) fantasma' % vazias)

        # --- F. cabecalhos integros ---------------------------------------
        cab = [str(c.Name) for c in loEv.ListColumns]
        checar('F. cabecalho de eventos igual ao contrato (14 colunas)',
               cab == CAB_EV,
               'esperado %s\nveio     %s' % (CAB_EV, cab))
        cabEq = [str(c.Name) for c in loEq.ListColumns]
        checar('F. cabecalho de EQA com 21 colunas, sem vazio',
               len(cabEq) == 21 and all(x.strip() for x in cabEq),
               '%d colunas: %s' % (len(cabEq), cabEq))

        # --- C. reconstruir o historico nao destroi a tabela ---------------
        antes = loEv.Range.Address
        xl.Run('RegistrarEventosWestgard')
        wsEv2, loEv2 = tabela(wb, EV_TAB)
        checar('C. a tabela sobrevive a RegistrarEventosWestgard',
               loEv2 is not None)
        if loEv2 is not None:
            total2 = int(wsEv2.Range('J2').Value or 0)
            checar('C. e acompanha o tamanho do historico reconstruido',
                   loEv2.ListRows.Count == total2
                   or (total2 == 0 and loEv2.ListRows.Count == 1),
                   'J2=%d, tabela=%d (antes %s, agora %s)'
                   % (total2, loEv2.ListRows.Count, antes, loEv2.Range.Address))
            vaz2 = 0
            if total2:
                for r in loEv2.DataBodyRange.Value:
                    if not str(r[2] or '').strip() or not str(r[4] or '').strip():
                        vaz2 += 1
            checar('D. sem linha fantasma apos a reconstrucao', vaz2 == 0,
                   '%d linha(s)' % vaz2)
    finally:
        try:
            wb.Close(False)       # nada e salvo
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
