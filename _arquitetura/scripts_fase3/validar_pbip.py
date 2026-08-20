# -*- coding: utf-8 -*-
"""validar_pbip.py - o relatorio abre? antes de pedir para o Power BI descobrir

Um PBIP quebrado nao avisa em texto: o Power BI Desktop abre uma caixa de erro e
para. Como a validacao visual custa caro, estas conferencias baratas rodam antes
e pegam a maioria dos defeitos que a clonagem de paginas pode introduzir.

Confere:
  1. todo JSON e TMDL abre
  2. nomes de pagina e de visual sao unicos (id repetido = visual sumido)
  3. o name de cada visual bate com a pasta que o contem
  4. pageOrder cita exatamente as paginas existentes
  5. os filtros de pagina apontam para coluna que existe no modelo
  6. as medidas referenciadas pelos visuais existem no modelo

Uso: python validar_pbip.py
"""
import io
import os
import re
import sys
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PBI = os.path.join(RAIZ, 'PowerBI')
MODELO = os.path.join(PBI, 'QC_INI_Bioquimica.SemanticModel', 'definition')
PAGES = os.path.join(PBI, 'QC_INI_Bioquimica.Report', 'definition', 'pages')

falhas = []


def ck(nome, cond, det=''):
    print(('  OK   ' if cond else '  FALHA') + '  ' + nome + (('  -> ' + det) if det else ''))
    if not cond:
        falhas.append(nome)


def main():
    # ---- 1. tudo abre ----------------------------------------------------
    ruins = []
    njson = 0
    for base, _, arqs in os.walk(PBI):
        for a in arqs:
            if a.endswith('.json'):
                njson += 1
                try:
                    json.load(io.open(os.path.join(base, a), encoding='utf-8'))
                except Exception as e:
                    ruins.append('%s: %s' % (a, str(e)[:60]))
    ck('todos os %d JSON abrem' % njson, not ruins, '; '.join(ruins[:3]))

    # ---- 2/3. paginas e visuais -----------------------------------------
    pastas = sorted(d for d in os.listdir(PAGES)
                    if os.path.isdir(os.path.join(PAGES, d)))
    nomes = []
    ids_visuais = []
    desalinhados = []
    for d in pastas:
        j = json.load(io.open(os.path.join(PAGES, d, 'page.json'), encoding='utf-8'))
        nomes.append(j.get('displayName'))
        ck('page.json de %s tem name = pasta' % j.get('displayName'),
           j.get('name') == d, '%s vs %s' % (j.get('name'), d))
        vd = os.path.join(PAGES, d, 'visuals')
        for v in sorted(os.listdir(vd)):
            ids_visuais.append(v)
            vj = json.load(io.open(os.path.join(vd, v, 'visual.json'), encoding='utf-8'))
            if vj.get('name') != v:
                desalinhados.append('%s/%s (name=%s)' % (d, v, vj.get('name')))
    ck('nomes de pagina unicos', len(set(nomes)) == len(nomes), str(nomes))
    ck('ids de visual unicos em todo o relatorio',
       len(set(ids_visuais)) == len(ids_visuais),
       '%d ids, %d unicos' % (len(ids_visuais), len(set(ids_visuais))))
    ck('name de cada visual bate com a pasta', not desalinhados,
       '; '.join(desalinhados[:3]))

    # ---- 4. pageOrder ----------------------------------------------------
    pj = json.load(io.open(os.path.join(PAGES, 'pages.json'), encoding='utf-8'))
    ordem = pj.get('pageOrder', [])
    ck('pageOrder cobre todas as paginas',
       set(ordem) == set(pastas), '%d na ordem, %d pastas' % (len(ordem), len(pastas)))
    ck('activePageName existe', pj.get('activePageName') in pastas,
       str(pj.get('activePageName')))

    # ---- 5. colunas e medidas do modelo ---------------------------------
    tmdl = ''
    for base, _, arqs in os.walk(MODELO):
        for a in arqs:
            if a.endswith('.tmdl'):
                tmdl += io.open(os.path.join(base, a), encoding='utf-8').read() + '\n'
    colunas = set(re.findall(r'(?m)^\s*column\s+([\'"]?)([^\r\n\'"=]+)\1\s*$', tmdl))
    colunas = {c[1].strip() for c in colunas}
    colunas |= {m.strip().strip("'") for m in
                re.findall(r'(?m)^\s*column\s+(.+?)\s*=', tmdl)}
    medidas = {m.strip().strip("'") for m in
               re.findall(r'(?m)^\s*measure\s+(.+?)\s*=', tmdl)}
    print('  modelo: %d colunas, %d medidas' % (len(colunas), len(medidas)))

    # filtros de pagina
    faltando = []
    for d in pastas:
        j = json.load(io.open(os.path.join(PAGES, d, 'page.json'), encoding='utf-8'))
        for f in (j.get('filterConfig') or {}).get('filters', []):
            prop = (f.get('field') or {}).get('Column', {}).get('Property')
            if prop and prop not in colunas:
                faltando.append('%s -> %s' % (j.get('displayName'), prop))
    ck('filtros de pagina usam coluna existente', not faltando, '; '.join(faltando))

    # medidas citadas pelos visuais
    citadas = set()
    for d in pastas:
        vd = os.path.join(PAGES, d, 'visuals')
        for v in os.listdir(vd):
            txt = io.open(os.path.join(vd, v, 'visual.json'), encoding='utf-8').read()
            for m in re.findall(r'"Measure":\s*\{[^}]*"Property":\s*"([^"]+)"', txt):
                citadas.add(m)
    orfas = sorted(m for m in citadas if m not in medidas)
    ck('toda medida citada pelos visuais existe no modelo (%d citadas)' % len(citadas),
       not orfas, '; '.join(orfas[:6]))

    print()
    print('=' * 60)
    if falhas:
        print('FALHAS (%d):' % len(falhas))
        for f in falhas:
            print('   - %s' % f)
        sys.exit(1)
    print('PBIP ESTRUTURALMENTE VALIDO')


if __name__ == '__main__':
    main()
