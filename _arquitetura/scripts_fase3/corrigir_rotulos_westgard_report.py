# -*- coding: utf-8 -*-
"""corrigir_rotulos_westgard_report.py - o nome fisico nunca chega ao usuario

O PROBLEMA

As medidas por posicao passaram a se chamar 'Viol R2', 'Viol R4' e 'Viol R5',
porque as posicoes 2, 4 e 5 sao regras diferentes em cada area. Trocar so a
referencia deixaria "Viol R5" no cabecalho da tabela -- nome fisico na cara do
analista, que nao diz nada.

A SAIDA

No PBIR, a referencia da medida e o ROTULO exibido sao campos diferentes:

    Property        / queryRef   -> a medida (fisico)
    nativeQueryRef              -> o que aparece na tela

As paginas ja sao por area ("Bioquimica - Painel", "Hematologia - Painel"),
entao o rotulo correto e conhecido sem nenhuma logica nova:

    posicao   Bioquimica   Hematologia
    R2        2_2s         2of3_2s
    R4        4_1s         3_1s
    R5        8x           6x

Assim o usuario ve a regra da area dele, e o modelo continua com uma coluna so
por posicao -- que e o que permite as duas areas viverem na mesma fato.

Tambem corrige 'Viol 10x': 10x nao existe em lugar nenhum operacional desde o
ADR-041, e continuava como rotulo de uma coluna de tabela.
"""
import io
import os
import re
import sys
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
PAGES = os.path.join(BASE, 'PowerBI', 'QC_INI_Bioquimica.Report',
                     'definition', 'pages')

# medida antiga -> (medida nova, rotulo Bioquimica, rotulo Hematologia)
MAPA = {
    'Viol 2_2s': ('Viol R2', '2_2s', '2of3_2s'),
    'Viol 4_1s': ('Viol R4', '4_1s', '3_1s'),
    'Viol 10x': ('Viol R5', '8x', '6x'),
    'Viol 8x': ('Viol R5', '8x', '6x'),
}


def area_da_pagina(dirpag):
    p = os.path.join(dirpag, 'page.json')
    try:
        j = json.load(io.open(p, encoding='utf-8'))
    except Exception:
        return None
    nome = (j.get('displayName') or '')
    # o displayName vem com acento; basta o prefixo
    baixo = nome.lower()
    if baixo.startswith('hemato') or 'hemato' in baixo:
        return 'Hematologia'
    if baixo.startswith('bio') or 'bio' in baixo:
        return 'Bioquimica'
    return None


def main():
    tot = mud = 0
    for pag in sorted(os.listdir(PAGES)):
        dirpag = os.path.join(PAGES, pag)
        if not os.path.isdir(dirpag):
            continue
        area = area_da_pagina(dirpag)
        vis = os.path.join(dirpag, 'visuals')
        if not os.path.isdir(vis):
            continue
        for v in sorted(os.listdir(vis)):
            f = os.path.join(vis, v, 'visual.json')
            if not os.path.exists(f):
                continue
            tot += 1
            t = io.open(f, encoding='utf-8').read()
            orig = t
            for antigo, (nova, rotBio, rotHem) in MAPA.items():
                if antigo not in t:
                    continue
                rot = rotHem if area == 'Hematologia' else rotBio
                # o rotulo exibido: so quando a area e conhecida
                if area:
                    t = t.replace('"nativeQueryRef": "%s"' % antigo,
                                  '"nativeQueryRef": "%s"' % rot)
                # a referencia da medida
                t = t.replace('"Property": "%s"' % antigo,
                              '"Property": "%s"' % nova)
                t = t.replace('"queryRef": "Fato_QC.%s"' % antigo,
                              '"queryRef": "Fato_QC.%s"' % nova)
                # sobra de rotulo em qualquer outro lugar
                t = t.replace('"%s"' % antigo, '"%s"' % nova)
            if t != orig:
                io.open(f, 'w', encoding='utf-8', newline='\n').write(t)
                mud += 1
                print('  %-12s %s/%s' % (area or '(sem area)', pag[:8], v[:8]))
    print()
    print('%d visual(is) examinados, %d corrigido(s)' % (tot, mud))

    # confere que nao sobrou nome fisico visivel nem regra aposentada
    resto = {}
    for raiz, _d, arqs in os.walk(PAGES):
        for a in arqs:
            if a != 'visual.json':
                continue
            t = io.open(os.path.join(raiz, a), encoding='utf-8').read()
            for m in re.finditer(r'"nativeQueryRef":\s*"([^"]*)"', t):
                r = m.group(1)
                if re.search(r'\bR[245]\b', r) or '10x' in r:
                    resto.setdefault(r, 0)
                    resto[r] += 1
    print()
    if resto:
        print('AINDA VISIVEL AO USUARIO: %s' % resto)
        sys.exit(1)
    print('nenhum rotulo posicional (R2/R4/R5) nem 10x visivel ao usuario')


main()
