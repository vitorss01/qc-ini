# -*- coding: utf-8 -*-
"""testar_compilacao.py - o equivalente ao Debug > Compile VBAProject

POR QUE NAO E SO "ABRIR E VER SE DA ERRO"

O Excel nao expoe um metodo Compile pela automacao. O que existe e uma
propriedade util: quando QUALQUER modulo do projeto nao compila, o VBA recusa
executar QUALQUER macro -- e o erro que chega ao cliente COM e um generico
"Classe nao registrada" ou "macro nao disponivel", que nao parece erro de
compilacao nenhum. Foi exatamente esse o sintoma quando LiberarEscrita sumiu
no splice: quatro rotinas "falharam com 1004" quando na verdade o projeto
inteiro nao compilava.

Entao a prova e comportamental e vale: se um Application.Run trivial responde,
o projeto compilou. Sondar VARIOS pontos de entrada, de modulos diferentes,
acrescenta a garantia de que cada modulo chave esta presente e publico -- que
e a outra metade do que o Debug > Compile pega.

O QUE PROVA

  1. o projeto compila (uma sonda qualquer responde)
  2. cada API publica sondada existe e e alcancavel por Application.Run
  3. as duas rotinas da guarda de protecao (ADR-046) estao no projeto

O QUE NAO PROVA

Que o codigo esta certo. Compilar e o piso, nao o teto.

Uso: python testar_compilacao.py <arquivo.xlsm>
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

# (nome, argumentos, modulo de origem) -- so funcoes SEM efeito colateral
SONDAS = [
    ('AreaDoProduto', (), 'mEstatistica'),
    ('MatrizWestgard', (), 'mEstatistica'),
    ('DetectoresWestgard', (), 'mEstatistica'),
    ('ValidarMatrizWestgard', (), 'mEstatistica'),
    ('LimiarSequencialWestgard', (), 'mEstatistica'),
    ('ClassificarSigma', (4.5,), 'mQualidade'),
    ('CoberturaWestgard', (3.5,), 'mPlanoQC'),
    ('ContratoWestgard', (), 'mPlanoQC'),
    ('VersaoCfg', (), 'mConfig'),
    ('PapelSistema', (), 'mAuditoria'),
    # mBI entra pela reconciliacao, que so LE e devolve texto. E a sonda que
    # cobre GarantirAba/Cab -- onde apareceu um relato de "variavel nao
    # declarada". Se o mBI nao compilasse, nenhuma sonda acima responderia
    # (o VBA compila o projeto inteiro antes de executar), mas deixar isso por
    # inferencia e pior do que sondar o modulo diretamente.
    ('ReconciliarBancoBI', (), 'mBI'),
]

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-52s %s' % (nome[:52], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for l in str(detalhe).split('\n')[:4]:
            print('        %s' % l)


def main():
    caminho = os.path.abspath(sys.argv[1])
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, True)
    try:
        print('=' * 78)
        print('COMPILACAO -- %s' % os.path.basename(caminho))
        print('=' * 78)
        print()

        for nome, args, mod in SONDAS:
            erro = ''
            try:
                xl.Run(nome, *args)
            except Exception as e:
                erro = str(e)
            checar('%-24s (%s)' % (nome, mod), not erro, erro)

        # A guarda do ADR-046 tem de existir e ser publica. Sondar por
        # Application.Run exige um Worksheet como argumento, entao a
        # conferencia e no proprio projeto.
        print()
        nomes = set()
        for c in wb.VBProject.VBComponents:
            cm = c.CodeModule
            if cm.CountOfLines:
                txt = cm.Lines(1, cm.CountOfLines)
                for alvo in ('LiberarEscrita', 'RestaurarProtecao'):
                    if ('Public Function %s' % alvo) in txt or \
                       ('Public Sub %s' % alvo) in txt:
                        nomes.add((alvo, c.Name))
        for alvo in ('LiberarEscrita', 'RestaurarProtecao'):
            onde = [m for a, m in nomes if a == alvo]
            checar('%s publica, em UM modulo so' % alvo,
                   len(onde) == 1,
                   'encontrada em: %s' % (onde or 'nenhum'))
            if len(onde) == 1:
                print('        -> %s' % onde[0])
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
