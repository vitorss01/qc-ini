# -*- coding: utf-8 -*-
"""montar_relatorio_4paginas.py - ADR-029: o relatorio ganha os dois produtos

O QUE JA EXISTIA

O PBIP tinha tres paginas -- Painel, Estatistica e Visao Gerencial -- todas
completas e todas SO da Bioquimica: Levey-Jennings, 16 cards, 8 slicers, tabela
de Westgard, distribuicao de Sigma. O modelo lia UM arquivo.

O QUE ESTE SCRIPT FAZ

  1. o modelo passa a ler os DOIS artefatos de build e a empilha-los
  2. o parametro de caminho deixa de apontar para o perfil da outra maquina
  3. as paginas viram quatro, filtradas por Produto:

        Bioquimica - Painel        Bioquimica - Estatistica
        Hematologia - Painel       Hematologia - Estatistica

     A Visao Gerencial e preservada como quinta pagina: ela cruza os produtos e
     nao competiria com nenhuma das quatro.

POR QUE EMPILHAR EM VEZ DE DOIS MODELOS

A tabela fato ja nasce com as colunas Produto e WorkbookID -- o contrato do
ADR-026 foi desenhado para isto. Dois modelos separados obrigariam a duplicar
as 60 medidas DAX, e duas copias da formula do Sigma e exatamente a divergencia
que o ADR-027 passou a sessao inteira eliminando.

A COMBINACAO E FEITA ANTES DA TIPAGEM

Os dois arquivos tem o mesmo schema de 60 colunas, entao Table.Combine no dado
cru e depois um unico Table.TransformColumnTypes. Tipar duas vezes abriria
espaco para os dois lados divergirem no tipo de uma coluna.

O FILTRO DE PAGINA E DE PAGINA, NAO DE VISUAL

Cada pagina recebe um filterConfig no proprio page.json. Filtrar visual a visual
deixaria um card sem filtro em algum canto mostrando o produto errado -- e o
pedido pede explicitamente que isso nao aconteca.

Uso: python montar_relatorio_4paginas.py
"""
import io
import os
import re
import sys
import json
import shutil
import hashlib

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PBI = os.path.join(RAIZ, 'PowerBI')
MODELO = os.path.join(PBI, 'QC_INI_Bioquimica.SemanticModel', 'definition')
REPORT = os.path.join(PBI, 'QC_INI_Bioquimica.Report', 'definition')
PAGES = os.path.join(REPORT, 'pages')

PERFIL = os.path.expanduser('~')
CAM_BIO = os.path.join(PERFIL, 'QCINI_build_hardening1_Bioquimica', 'QC_Bioquimica.xlsm')
CAM_HEM = os.path.join(PERFIL, 'QCINI_build_hardening1_Hematologia', 'QC_Hematologia.xlsm')


def ler(p):
    return io.open(p, encoding='utf-8').read()


def escrever(p, s):
    io.open(p, 'w', encoding='utf-8', newline='\n').write(s)


def novo_id(semente):
    return hashlib.sha1(semente.encode('utf-8')).hexdigest()[:20]


# ---------------------------------------------------------------- 1. modelo
def ajustar_expressoes():
    p = os.path.join(MODELO, 'expressions.tmdl')
    s = ler(p)
    s = re.sub(r'expression pCaminhoQC = "[^"]*"',
               'expression pCaminhoQC = "%s"' % CAM_BIO.replace('\\', '\\\\'), s, count=1)
    if 'pCaminhoHema' not in s:
        bloco = (
            '\nexpression pCaminhoHema = "%s" meta [IsParameterQuery=true, Type="Text", '
            'IsParameterQueryRequired=true]\n'
            '\tlineageTag: %s\n\n'
            '\tannotation PBI_NavigationStepName = Navegação\n\n'
            '\tannotation PBI_ResultType = Text\n'
            % (CAM_HEM.replace('\\', '\\\\'), '8c4f21a0-1111-4c2b-9f30-aa5566778899'))
        s = s.rstrip() + '\n' + bloco
    escrever(p, s)
    print('parametros: pCaminhoQC -> %s' % CAM_BIO)
    print('            pCaminhoHema -> %s' % CAM_HEM)


def ajustar_fato():
    p = os.path.join(MODELO, 'tables', 'Fato_QC.tmdl')
    s = ler(p)
    if 'TabelaHema' in s:
        print('Fato_QC ja empilha os dois produtos')
        return
    velho_fonte = 'Fonte = Excel.Workbook(File.Contents(pCaminhoQC), null, true),'
    if velho_fonte not in s:
        raise SystemExit('linha Fonte nao encontrada em Fato_QC.tmdl')
    # a indentacao do TMDL e por tabulacao; preserva a da linha original
    ind = s[:s.index(velho_fonte)].split('\n')[-1]
    novo = (
        'Fonte = Excel.Workbook(File.Contents(pCaminhoQC), null, true),\n'
        '%s// Os dois produtos entram na MESMA fato. A coluna Produto ja vem do\n'
        '%s// motor (ADR-026) e e ela que separa as paginas. Dois modelos\n'
        '%s// separados obrigariam a duplicar as 60 medidas -- e duas copias da\n'
        '%s// formula do Sigma sao a divergencia que o ADR-027 eliminou.\n'
        '%sFonteHema = Excel.Workbook(File.Contents(pCaminhoHema), null, true),'
        % (ind, ind, ind, ind, ind))
    s = s.replace(velho_fonte, novo, 1)

    velho_tab = 'Tabela = Fonte{[Item="tblBI_Fato", Kind="Table"]}[Data],'
    if velho_tab not in s:
        raise SystemExit('linha Tabela nao encontrada')
    novo_tab = (
        'TabelaBio = Fonte{[Item="tblBI_Fato", Kind="Table"]}[Data],\n'
        '%sTabelaHema = FonteHema{[Item="tblBI_Fato", Kind="Table"]}[Data],\n'
        '%s// Combina no dado CRU e tipa uma vez so: tipar os dois lados em\n'
        '%s// separado abriria espaco para divergirem no tipo de uma coluna.\n'
        '%sTabela = Table.Combine({TabelaBio, TabelaHema}),'
        % (ind, ind, ind, ind))
    s = s.replace(velho_tab, novo_tab, 1)
    escrever(p, s)
    print('Fato_QC: passa a empilhar Bioquimica + Hematologia')


# ---------------------------------------------------------------- 2. paginas
def nome_pagina(d):
    return json.load(io.open(os.path.join(PAGES, d, 'page.json'), encoding='utf-8'))


def filtro_produto(valor):
    """filterConfig de pagina em Fato_QC[Produto]."""
    return {
        "filters": [{
            "name": "filtroProduto",
            "field": {
                "Column": {
                    "Expression": {"SourceRef": {"Entity": "Fato_QC"}},
                    "Property": "Produto"
                }
            },
            "type": "Categorical",
            "filter": {
                "Version": 2,
                "From": [{"Name": "f", "Entity": "Fato_QC", "Type": 0}],
                "Where": [{
                    "Condition": {
                        "In": {
                            "Expressions": [{
                                "Column": {
                                    "Expression": {"SourceRef": {"Source": "f"}},
                                    "Property": "Produto"
                                }
                            }],
                            "Values": [[{"Literal": {"Value": "'%s'" % valor}}]]
                        }
                    }
                }]
            },
            "howCreated": "Auto",
            "isLockedInViewMode": True,
            "isHiddenInViewMode": False
        }]
    }


def clonar_pagina(origem, novo_nome, produto):
    destino_id = novo_id(novo_nome)
    dorig = os.path.join(PAGES, origem)
    ddest = os.path.join(PAGES, destino_id)
    if os.path.exists(ddest):
        shutil.rmtree(ddest)
    shutil.copytree(dorig, ddest)
    # os visuais precisam de nomes proprios, senao dois visuais compartilham id
    vdir = os.path.join(ddest, 'visuals')
    for v in os.listdir(vdir):
        novo_v = novo_id(destino_id + v)
        os.rename(os.path.join(vdir, v), os.path.join(vdir, novo_v))
        pj = os.path.join(vdir, novo_v, 'visual.json')
        j = json.load(io.open(pj, encoding='utf-8'))
        j['name'] = novo_v
        json.dump(j, io.open(pj, 'w', encoding='utf-8', newline='\n'),
                  ensure_ascii=False, indent=2)
    pj = os.path.join(ddest, 'page.json')
    j = json.load(io.open(pj, encoding='utf-8'))
    j['name'] = destino_id
    j['displayName'] = novo_nome
    j['filterConfig'] = filtro_produto(produto)
    json.dump(j, io.open(pj, 'w', encoding='utf-8', newline='\n'),
              ensure_ascii=False, indent=2)
    print('   pagina criada: %-30s [%s]' % (novo_nome, destino_id))
    return destino_id


def renomear(origem, novo_nome, produto):
    pj = os.path.join(PAGES, origem, 'page.json')
    j = json.load(io.open(pj, encoding='utf-8'))
    j['displayName'] = novo_nome
    j['filterConfig'] = filtro_produto(produto)
    json.dump(j, io.open(pj, 'w', encoding='utf-8', newline='\n'),
              ensure_ascii=False, indent=2)
    print('   pagina renomeada: %-28s [%s]' % (novo_nome, origem))


def main():
    print('=== 1. MODELO ===')
    ajustar_expressoes()
    ajustar_fato()

    print()
    print('=== 2. PAGINAS ===')
    atual = {}
    for d in os.listdir(PAGES):
        if os.path.isdir(os.path.join(PAGES, d)):
            atual[nome_pagina(d).get('displayName', '')] = d
    print('   existentes: %s' % list(atual))

    painel = atual.get('Painel')
    estat = atual.get('Estatística') or atual.get('Estatistica')
    geral = atual.get('Visão Gerencial') or atual.get('Visao Gerencial')
    if not painel or not estat:
        raise SystemExit('paginas base nao encontradas: %s' % list(atual))

    hp = clonar_pagina(painel, 'Hematologia — Painel', 'Hematologia')
    he = clonar_pagina(estat, 'Hematologia — Estatística', 'Hematologia')
    renomear(painel, 'Bioquímica — Painel', 'Bioquimica')
    renomear(estat, 'Bioquímica — Estatística', 'Bioquimica')

    ordem = [painel, estat, hp, he] + ([geral] if geral else [])
    pj = os.path.join(PAGES, 'pages.json')
    j = json.load(io.open(pj, encoding='utf-8'))
    j['pageOrder'] = ordem
    j['activePageName'] = painel
    json.dump(j, io.open(pj, 'w', encoding='utf-8', newline='\n'),
              ensure_ascii=False, indent=2)
    print('   ordem: %d paginas, ativa = Bioquimica - Painel' % len(ordem))
    print()
    print('PRONTO')


if __name__ == '__main__':
    main()
