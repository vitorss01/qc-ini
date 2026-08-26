# -*- coding: utf-8 -*-
"""gerar_powerquery.py - a tipagem do Power Query sai do CONTRATO

POR QUE GERAR, E NAO MANTER A MAO

A lista de Table.TransformColumnTypes tinha 81 das 84 colunas do contrato --
AtualizadoEmUTC, Vigencia_Inicio e Vigencia_Fim entravam sem tipo declarado.
Ninguem percebeu porque nada comparava as duas listas. Gerando a partir do
CONTRATO_BI.json, divergir deixa de ser possivel: a fonte e uma so.

Isto NAO enfraquece o schema gate. O gate compara CONTRATO x EXCEL -- duas
fontes independentes. Aqui o Power Query e CONSUMIDOR do contrato, nao uma
terceira opiniao sobre ele.

TOLERANCIA A COLUNA AUSENTE

A tipagem passa a ser filtrada por Table.ColumnNames. Sem isso, uma coluna que
so existe num produto (W_6x na Hematologia, W_8x na Bioquimica) derrubaria o
refresh inteiro com "column not found" -- e e exatamente essa a forma do
contrato depois do ADR-047.

Uso:
    python gerar_powerquery.py --mostrar
    python gerar_powerquery.py --aplicar
"""
import io
import json
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONTRATO = os.path.join(RAIZ, 'CONTRATO_BI.json')
TMDL = os.path.join(RAIZ, '..', 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
                    'definition', 'tables', 'Fato_QC.tmdl')
TMDL = os.path.normpath(TMDL)

TIPO_M = {
    'texto': 'type text',
    'inteiro': 'Int64.Type',
    'decimal': 'type number',
    'data': 'type date',
    'datahora': 'type datetime',
    'logico': 'type logical',
}

IND = '\t\t\t\t'          # indentacao do bloco M dentro do TMDL


def montar():
    c = json.load(io.open(CONTRATO, encoding='utf-8'))
    pares = []
    for campo in c['campos']:
        m = TIPO_M.get(campo['tipo'])
        if not m:
            raise SystemExit('tipo sem mapeamento M: %r' % campo['tipo'])
        pares.append('{"%s", %s}' % (campo['nome'], m))

    linhas = []
    linhas.append(IND + '    // TIPAGEM GERADA DE CONTRATO_BI.json (nao editar a mao).')
    linhas.append(IND + '    // gerar_powerquery.py --aplicar reescreve este bloco.')
    linhas.append(IND + '    //')
    linhas.append(IND + '    // A lista traz a UNIAO dos dois produtos: depois do ADR-047 cada um')
    linhas.append(IND + '    // grava o nome da SUA regra (W_2_2s na Bioquimica, W_2of3_2s na')
    linhas.append(IND + '    // Hematologia), entao metade das colunas de Westgard so existe de um')
    linhas.append(IND + '    // lado. Por isso a tipagem e filtrada por Table.ColumnNames: sem o')
    linhas.append(IND + '    // filtro, uma coluna ausente derruba o refresh inteiro.')
    linhas.append(IND + '    TiposContrato = {')
    for i in range(0, len(pares), 3):
        linhas.append(IND + '        ' + ', '.join(pares[i:i + 3]) +
                      (',' if i + 3 < len(pares) else ''))
    linhas.append(IND + '    },')
    linhas.append(IND + '    ColunasPresentes = Table.ColumnNames(Tabela),')
    linhas.append(IND + '    TiposAplicaveis = List.Select(TiposContrato,')
    linhas.append(IND + '        each List.Contains(ColunasPresentes, _{0})),')
    linhas.append(IND + '    Tipada = Table.TransformColumnTypes(Tabela, TiposAplicaveis),')
    return '\n'.join(linhas), len(pares)


def aplicar(bloco):
    txt = io.open(TMDL, encoding='utf-8', newline='').read()
    nl = '\r\n' if '\r\n' in txt else '\n'
    t = txt.replace('\r\n', '\n')

    ini = t.index('Tipada = Table.TransformColumnTypes(Tabela, {')
    # recua ate o inicio da linha
    ini = t.rindex('\n', 0, ini) + 1
    # o bloco antigo termina na primeira linha que fecha com "}),"
    fim = t.index('\n', t.index('}),', ini)) + 1

    novo = t[:ini] + bloco + '\n' + t[fim:]
    io.open(TMDL, 'w', encoding='utf-8', newline=nl).write(novo)


def main():
    bloco, n = montar()
    if '--mostrar' in sys.argv:
        print(bloco)
        print()
        print('%d colunas tipadas' % n)
        return 0
    if '--aplicar' in sys.argv:
        aplicar(bloco)
        print('Fato_QC.tmdl: tipagem regerada com %d colunas do contrato' % n)
        return 0
    print(__doc__)
    return 2


if __name__ == '__main__':
    sys.exit(main())
