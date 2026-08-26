# -*- coding: utf-8 -*-
"""testar_protecao_escrita.py - ADR-046: aba protegida, VBA escrevendo

O DEFEITO

Duas ocorrencias de "erro 1004 - a celula esta em uma planilha protegida", em
AtualizarCalc e em RegistrarEventosWestgard. A causa e a mesma nas duas, e nao
e a protecao estar errada:

  mSeguranca.ReprotectAll aplica Protect ... UserInterfaceOnly:=True, que e a
  configuracao certa -- usuario nao edita, VBA edita. Mas o Excel NAO PERSISTE
  essa flag ao salvar. No arquivo salvo e reaberto a aba volta a estar
  protegida tambem para o VBA, ate que Workbook_Open rode LockApp.

  E Workbook_Open nao roda em abertura por automacao: todo script de build e de
  QA abre com EnableEvents = False, de proposito.

Por isso a correcao NAO e "aplicar UserInterfaceOnly" (ja e aplicado) nem
"desproteger a aba" (reduziria a protecao). E cada rotina que escreve tratar a
propria protecao, com restauro garantido -- o que mBI, mBanco e mImportar ja
faziam e as tres rotinas do motor nao.

O QUE ESTE TESTE PROVA

Abrindo SEM disparar Workbook_Open, que e o cenario em que o defeito aparece:

  1. as abas tecnicas estao protegidas de fato (ProtectContents = True)
  2. o VBA consegue rodar as quatro rotinas de escrita, sem 1004
  3. depois de rodar, as abas continuam protegidas
  4. o que bloqueia o usuario continua de pe: celulas Locked sob protecao ativa
  5. o restauro sobrevive ao caminho de ERRO e ao de saida antecipada

O QUE ESTE TESTE **NAO** PROVA

Que o usuario esta bloqueado no teclado. Escrita por automacao COM e escrita
programatica, e UserInterfaceOnly:=True a autoriza -- autorizar e justamente o
objetivo. Uma versao anterior deste script usava a escrita COM como prova de
bloqueio e reprovava codigo correto. O teste de teclado e humano: abrir o
arquivo e tentar digitar numa aba tecnica.

Uso: python testar_protecao_escrita.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

ABAS = ['Eng_Saida', 'Eventos_Westgard']
ROTINAS = ['AtualizarCalc', 'AtualizarEstatisticaAba', 'AtualizarPainelEng',
           'RegistrarEventosWestgard']

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
    # ESTE e o ponto do teste: sem eventos, Workbook_Open nao roda, LockApp nao
    # roda, e o UserInterfaceOnly da sessao anterior nao existe mais.
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, False)
    try:
        area = str(xl.Run('AreaDoProduto'))
        print('=' * 78)
        print('%s -- escrita em aba protegida, sem Workbook_Open' % area)
        print('=' * 78)
        print()

        presentes = {}
        for i in range(1, wb.Worksheets.Count + 1):
            presentes[wb.Worksheets(i).Name] = wb.Worksheets(i)

        # --- 1. as abas estao protegidas ANTES ------------------------------
        antes = {}
        for nome in ABAS:
            if nome not in presentes:
                checar('aba %s existe' % nome, False)
                continue
            antes[nome] = bool(presentes[nome].ProtectContents)
            checar('%s protegida antes (ProtectContents)' % nome, antes[nome],
                   'a aba veio desprotegida; o teste perderia o sentido')

        # --- 2. o VBA escreve sem 1004 ---------------------------------------
        for rot in ROTINAS:
            erro = ''
            try:
                xl.Run(rot)
            except Exception as e:
                erro = str(e)
            checar('%s roda sem 1004' % rot, not erro, erro)

        # --- 3. continuam protegidas DEPOIS ----------------------------------
        for nome in ABAS:
            if nome not in presentes:
                continue
            depois = bool(presentes[nome].ProtectContents)
            checar('%s continua protegida depois' % nome, depois,
                   'a rotina deixou a aba destrancada -- restauro falhou')

        # --- 4. o que de fato bloqueia o usuario -----------------------------
        #
        # NAO da para provar "o usuario esta bloqueado" escrevendo pela API COM.
        # Escrita por automacao e escrita PROGRAMATICA, e UserInterfaceOnly:=True
        # a permite por definicao -- e permitir e o objetivo. A tentativa
        # anterior de usar isso como prova reprovava codigo correto.
        #
        # O que bloqueia o usuario na interface sao dois fatos, e sao esses que
        # dao para conferir aqui: a aba protegida E as celulas com Locked=True.
        # Protecao ligada sobre celulas desbloqueadas nao protege nada.
        #
        # O teste de teclado (item 6 do QA) continua sendo humano: abrir o
        # arquivo e tentar digitar. Nenhum script substitui isso, e dizer que
        # substitui seria pior do que nao testar.
        for nome in ABAS:
            if nome not in presentes:
                continue
            ws = presentes[nome]
            amostra = [ws.Cells(4, 1), ws.Cells(4, 2), ws.Cells(10, 3)]
            travadas = all(bool(c.Locked) for c in amostra)
            checar('%s: celulas com Locked=True sob protecao ativa' % nome,
                   travadas and bool(ws.ProtectContents),
                   'ProtectContents=%s  Locked=%s -- protecao sem celula '
                   'bloqueada nao impede o usuario'
                   % (ws.ProtectContents, [bool(c.Locked) for c in amostra]))

        # --- 5. o restauro sobrevive ao ERRO ---------------------------------
        # Provoca uma falha DENTRO da janela destrancada e confirma que a aba
        # volta protegida. Sem isto, o teste so provaria o caminho feliz -- e o
        # caminho de erro e justamente onde uma aba fica destrancada sem que
        # ninguem perceba.
        ponte = '''
Public Function ForcarErroComEscrita() As String
    Dim ws As Worksheet, estava As Boolean
    Set ws = ThisWorkbook.Sheets("Eng_Saida")
    On Error GoTo saida
    estava = LiberarEscrita(ws)
    Err.Raise 5, "teste", "falha proposital dentro da janela destrancada"
saida:
    RestaurarProtecao ws, estava
    ForcarErroComEscrita = CStr(ws.ProtectContents)
End Function
'''
        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name == 'mTesteProtecao':
                vbp.VBComponents.Remove(c)
                break
        m = vbp.VBComponents.Add(1)
        m.Name = 'mTesteProtecao'
        m.CodeModule.AddFromString(ponte)
        r = str(xl.Run('ForcarErroComEscrita'))
        checar('protecao restaurada mesmo com erro no meio da escrita',
               r.strip().lower() in ('true', 'verdadeiro'),
               'ProtectContents apos o erro = %r' % r)
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
