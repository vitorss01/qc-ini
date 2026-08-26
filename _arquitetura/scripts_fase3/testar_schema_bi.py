# -*- coding: utf-8 -*-
"""testar_schema_bi.py - schema gate da camada BI (§24 da Fase 2)

POR QUE NAO E CONTAGEM DE COLUNAS

Trocar "60 colunas" por "84 colunas" nao e schema gate: passa com o mesmo
numero de colunas erradas. O que quebra o BI em silencio nao e a contagem --
e a coluna que muda de nome, de tipo ou de significado sem ninguem perceber.
Foi exatamente o que aconteceu: W_10x carregou 8x e 6x por meses, e as flags
Usar_* da Hematologia ficaram falsas sem que nenhum teste acusasse.

O gate compara DUAS FONTES INDEPENDENTES:

  CONTRATO_BI.json   especificacao escrita a mao
  BI_Data            o que o mBI realmente gravou

Se o contrato fosse gerado do mBI, espelharia o erro e nao provaria nada.

O QUE VALIDA (§24: nomes, tipos, obrigatorios, semantica, chaves, duplicidade)

  1. todo campo do contrato para este produto existe na aba
  2. toda coluna da aba esta no contrato (nada entra sem ser declarado)
  3. nenhum nome proibido reapareceu -- em especial W_10x
  4. os campos do OUTRO produto nao existem aqui
  5. tipo de cada coluna compativel com o declarado
  6. campo obrigatorio nunca vem vazio
  7. chave do grao unica e sem vazio
  8. as cinco regras de Westgard, na ordem, sao as da matriz do produto
  9. flags Usar_* coerentes com o texto do plano (a divergencia do ADR-047)
 10. Classificacao_Sigma segue a escada unica do ADR-043

Uso: python testar_schema_bi.py <arquivo.xlsm> [--contrato <json>]
"""
import io
import json
import os
import sys
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)
import win32com.client as w

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PADRAO = os.path.join(RAIZ, 'CONTRATO_BI.json')

ESCADA = [(3.0, 'Desempenho inadequado'), (4.0, 'Marginal'), (5.0, 'Bom'),
          (6.0, 'Excelente'), (float('inf'), 'Classe mundial')]

falhas = []


def checar(nome, ok, detalhe=''):
    if not ok:
        falhas.append(nome)
    print('   %-56s %s' % (nome[:56], 'PASS' if ok else 'FAIL'))
    if not ok and detalhe:
        for l in str(detalhe).split('\n')[:8]:
            print('        %s' % l)


def tipo_ok(valor, tipo):
    """Vazio nunca reprova tipo -- obrigatoriedade e outra verificacao."""
    if valor is None or valor == '':
        return True
    if tipo == 'texto':
        return True
    if tipo == 'inteiro':
        return isinstance(valor, (int, float)) and float(valor) == int(valor)
    if tipo == 'decimal':
        return isinstance(valor, (int, float))
    if tipo in ('data', 'datahora'):
        return hasattr(valor, 'year') or isinstance(valor, datetime)
    if tipo == 'logico':
        return isinstance(valor, bool)
    return False


def classificar(sg):
    for lim, rot in ESCADA:
        if sg < lim:
            return rot
    return 'Classe mundial'


def main():
    caminho = os.path.abspath(sys.argv[1])
    cjson = PADRAO
    if '--contrato' in sys.argv:
        cjson = sys.argv[sys.argv.index('--contrato') + 1]
    contrato = json.load(io.open(cjson, encoding='utf-8'))

    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho, 0, True)
    try:
        area = str(xl.Run('AreaDoProduto')).lower()
        matriz = [r.strip() for r in str(xl.Run('MatrizWestgard')).split(';')]
        print('=' * 78)
        print('SCHEMA GATE -- %s (%s)' % (os.path.basename(caminho), area))
        print('=' * 78)
        print('   contrato %s  |  matriz %s' % (contrato['versao'], ' / '.join(matriz)))
        print()

        ws = wb.Worksheets(contrato['aba'])
        ur = ws.UsedRange
        ncolAba = ur.Column + ur.Columns.Count - 1
        nlin = ur.Row + ur.Rows.Count - 1
        cab = [str(c or '').strip() for c in
               ws.Range(ws.Cells(1, 1), ws.Cells(1, ncolAba)).Value[0]]
        cab = [c for c in cab if c]
        idx = {n: i for i, n in enumerate(cab)}

        meus = [c for c in contrato['campos']
                if c['produto'] in ('ambos', area)]
        alheios = [c for c in contrato['campos']
                   if c['produto'] not in ('ambos', area)]

        # --- 1. tudo que o contrato exige existe --------------------------
        faltando = [c['nome'] for c in meus if c['nome'] not in idx]
        checar('todo campo do contrato existe na aba (%d)' % len(meus),
               not faltando, 'ausentes: %s' % ', '.join(faltando))

        # --- 2. nada entra sem ser declarado ------------------------------
        declarados = {c['nome'] for c in contrato['campos']}
        intrusos = [c for c in cab if c not in declarados]
        checar('nenhuma coluna fora do contrato', not intrusos,
               'nao declaradas: %s' % ', '.join(intrusos))

        # --- 3. nomes proibidos -------------------------------------------
        proib = [c for c in cab if c in contrato['proibidos']['nomes']]
        checar('nenhum nome proibido (W_10x e afins)', not proib,
               'encontrados: %s' % ', '.join(proib))

        # --- 4. o produto nao carrega campo do outro ----------------------
        vazados = [c['nome'] for c in alheios if c['nome'] in idx]
        checar('nenhum campo do outro produto nesta aba', not vazados,
               'vazados: %s' % ', '.join(vazados))

        # --- 5. contagem declarada ----------------------------------------
        checar('%d colunas, como o contrato declara'
               % contrato['n_colunas_por_produto'],
               len(cab) == contrato['n_colunas_por_produto'],
               'a aba tem %d' % len(cab))

        # --- amostra para tipo / obrigatorio / chave ----------------------
        ate = min(nlin, 2000)
        dados = ws.Range(ws.Cells(2, 1), ws.Cells(ate, len(cab))).Value
        dados = [r for r in dados if any(x not in (None, '') for x in r)]
        print('   %d linhas amostradas de %d' % (len(dados), nlin - 1))
        print()

        # --- 5. tipos ------------------------------------------------------
        maus = []
        for c in meus:
            if c['nome'] not in idx:
                continue
            i = idx[c['nome']]
            for r in dados:
                if not tipo_ok(r[i], c['tipo']):
                    maus.append('%s: esperado %s, veio %r'
                                % (c['nome'], c['tipo'], r[i]))
                    break
        checar('tipo de cada coluna compativel com o contrato', not maus,
               '\n'.join(maus))

        # --- 6. obrigatorios ------------------------------------------------
        vazios = []
        for c in meus:
            if not c['obrigatorio'] or c['nome'] not in idx:
                continue
            i = idx[c['nome']]
            n = sum(1 for r in dados if r[i] in (None, ''))
            if n:
                vazios.append('%s: %d linha(s) vazias' % (c['nome'], n))
        checar('campo obrigatorio nunca vem vazio', not vazios,
               '\n'.join(vazios))

        # --- 7. chave do grao ----------------------------------------------
        chave = contrato['chave']
        if all(k in idx for k in chave):
            vistos, dup = set(), 0
            for r in dados:
                k = tuple(r[idx[x]] for x in chave)
                if k in vistos:
                    dup += 1
                vistos.add(k)
            checar('chave %s unica' % '+'.join(chave), dup == 0,
                   '%d duplicata(s)' % dup)
        else:
            checar('chave %s presente' % '+'.join(chave), False)

        # --- 8. as regras, na ordem, sao as da matriz -----------------------
        esperado = contrato['ordem_das_regras'][area]
        checar('ordem das regras do contrato = MatrizWestgard do motor',
               esperado == matriz,
               'contrato %s\nmotor    %s' % (esperado, matriz))
        wcols = ['W_' + r for r in matriz]
        ucols = ['Usar_' + r for r in matriz]
        checar('colunas W_* sao as da matriz deste produto',
               all(c in idx for c in wcols),
               'faltam: %s' % [c for c in wcols if c not in idx])
        checar('colunas Usar_* sao as da matriz deste produto',
               all(c in idx for c in ucols),
               'faltam: %s' % [c for c in ucols if c not in idx])

        # --- 9. flag x texto do plano (a divergencia do ADR-047) ------------
        if 'Regra_Westgard_Recomendada' in idx and all(c in idx for c in ucols):
            iTexto = idx['Regra_Westgard_Recomendada']
            div = []
            for r in dados:
                txt = str(r[iTexto] or '')
                if not txt:
                    continue
                tokens = {t.strip() for t in txt.replace('/', ' ').split()}
                for regra, col in zip(matriz, ucols):
                    noTexto = regra in tokens
                    naFlag = (r[idx[col]] is True)
                    if noTexto != naFlag:
                        div.append('%s: texto diz %s, %s = %s'
                                   % (regra, 'SIM' if noTexto else 'NAO',
                                      col, naFlag))
                        break
                if div:
                    break
            checar('flags Usar_* concordam com o texto do plano', not div,
                   '\n'.join(div))

        # --- 10. escada unica de Sigma --------------------------------------
        if 'Sigma' in idx and 'Classificacao_Sigma' in idx:
            erradas = []
            for r in dados:
                sg, cl = r[idx['Sigma']], r[idx['Classificacao_Sigma']]
                if not isinstance(sg, (int, float)) or not cl:
                    continue
                if str(cl).strip() != classificar(float(sg)):
                    erradas.append('Sigma %.3f classificado como %r, esperado %r'
                                   % (sg, cl, classificar(float(sg))))
                    if len(erradas) >= 3:
                        break
            checar('Classificacao_Sigma segue a escada unica (ADR-043)',
                   not erradas, '\n'.join(erradas))
    finally:
        wb.Close(False)
        xl.Quit()

    print()
    print('TOTAL: %d FAIL' % len(falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
