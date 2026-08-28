# -*- coding: utf-8 -*-
"""conferir_relatorio.py - o PBIR referencia algo que o modelo nao tem?

Um visualContainer malformado, ou que aponta para uma medida inexistente, nao
degrada: impede o RELATORIO INTEIRO de abrir. E so se descobre abrindo o
Desktop, um arquivo por vez.

Aqui cada visual.json e lido, cada referencia (Entity, Property) e cruzada com
o TMDL, e o que sobra e divergencia de verdade.
"""
import io
import os
import re
import sys
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
MOD = os.path.join(BASE, 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
                   'definition')
REL = os.path.join(BASE, 'PowerBI', 'QC_INI_Bioquimica.Report', 'definition')
PAGES = os.path.join(REL, 'pages')

ok = fail = 0


def reg(nome, cond, det=''):
    global ok, fail
    if cond:
        ok += 1
    else:
        fail += 1
    print('  %-5s %-54s %s' % ('OK' if cond else 'FALHA', nome[:54], det[:56]))


def ler(p):
    raw = io.open(p, 'rb').read()
    enc = 'utf-8-sig' if raw[:3] == b'\xef\xbb\xbf' else 'utf-8'
    return raw.decode(enc)


def modelo():
    """{entidade: set(colunas e medidas)}"""
    d = {}
    tdir = os.path.join(MOD, 'tables')
    for f in os.listdir(tdir):
        if not f.endswith('.tmdl'):
            continue
        s = ler(os.path.join(tdir, f))
        m = re.search(r'^table\s+(\S+)', s, re.M)
        if not m:
            continue
        ent = m.group(1).strip("'")
        nomes = set()
        # NOME DE MEDIDA PODE TER ESPACO SEM ASPAS EM TMDL.
        #
        # "measure % Grupos ET Conformes =" e valido, e cortar no primeiro
        # espaco fazia o comparador acusar 172 referencias orfas que existem.
        # Le-se ate o "=" (ou ate o fim da linha, no caso de coluna).
        # E o separador e " = " COM ESPACOS: existe medida chamada
        # "Grupos Sigma >= 6", e cortar no primeiro "=" a truncava em
        # "Grupos Sigma >" -- mais uma orfa inventada pelo parser.
        for r in re.finditer(r'^\s*(?:column|measure)\s+(.+?)(?:\s+=\s|\s*$)', s, re.M):
            nome = r.group(1).strip().strip("'")
            if nome:
                nomes.add(nome)
        d[ent] = nomes
    return d


def main():
    mods = modelo()
    print('modelo: %d tabelas' % len(mods))
    for e in sorted(mods):
        print('   %-26s %d campo(s)' % (e, len(mods[e])))
    print()

    # ---- paginas ------------------------------------------------------
    pj = json.load(io.open(os.path.join(PAGES, 'pages.json'), encoding='utf-8'))
    ordem = pj.get('pageOrder', [])
    dirs = sorted(d for d in os.listdir(PAGES)
                  if os.path.isdir(os.path.join(PAGES, d)))
    reg('toda pasta de pagina esta na ordem',
        set(dirs) == set(ordem),
        'so em disco: %s' % sorted(set(dirs) - set(ordem)))
    reg('activePageName existe', pj.get('activePageName') in dirs)

    nomes_pag = {}
    for d in dirs:
        p = os.path.join(PAGES, d, 'page.json')
        try:
            j = json.load(io.open(p, encoding='utf-8'))
            nomes_pag[d] = j.get('displayName', '(sem nome)')
            bom = j.get('name') == d and bool(j.get('displayName'))
        except Exception as e:
            nomes_pag[d] = '(ilegivel)'
            bom = False
        reg('page.json valido: %s' % nomes_pag[d], bom)

    # ---- visuais -------------------------------------------------------
    print()
    ruins, orfas, nvis = [], [], 0
    rotulos_fisicos = []
    for d in dirs:
        vdir = os.path.join(PAGES, d, 'visuals')
        if not os.path.isdir(vdir):
            continue
        for v in sorted(os.listdir(vdir)):
            f = os.path.join(vdir, v, 'visual.json')
            if not os.path.exists(f):
                continue
            nvis += 1
            try:
                j = json.load(io.open(f, encoding='utf-8'))
            except Exception as e:
                ruins.append('%s/%s: json invalido (%s)' % (d[:8], v[:8], e))
                continue
            if not j.get('name') or not j.get('position') or not j.get('visual'):
                ruins.append('%s/%s: falta name/position/visual' % (d[:8], v[:8]))
                continue
            if j.get('name') != v:
                ruins.append('%s/%s: name diverge da pasta' % (d[:8], v[:8]))
            txt = io.open(f, encoding='utf-8').read()
            for m in re.finditer(
                    r'"SourceRef":\s*\{\s*"Entity":\s*"([^"]+)"\s*\}\s*\},\s*"Property":\s*"([^"]+)"',
                    txt):
                ent, prop = m.group(1), m.group(2)
                if ent not in mods:
                    orfas.append('%s: entidade %s' % (nomes_pag.get(d, d), ent))
                elif prop not in mods[ent]:
                    orfas.append('%s: %s[%s]' % (nomes_pag.get(d, d), ent, prop))
            for m in re.finditer(r'"nativeQueryRef":\s*"([^"]*)"', txt):
                r = m.group(1)
                if re.search(r'\bR[1-5]\b', r) or '10x' in r:
                    rotulos_fisicos.append('%s: %s' % (nomes_pag.get(d, d), r))

    reg('visualContainer bem formado', not ruins,
        '%d problema(s): %s' % (len(ruins), ruins[:2]))
    reg('nenhuma referencia orfa ao modelo', not orfas,
        '%d: %s' % (len(orfas), sorted(set(orfas))[:3]))
    reg('nenhum nome fisico visivel ao usuario', not rotulos_fisicos,
        '%s' % sorted(set(rotulos_fisicos))[:3])
    print('  (%d visuais em %d paginas)' % (nvis, len(dirs)))

    # ---- as paginas exigidas existem -----------------------------------
    print()
    tem = ' | '.join(nomes_pag.values()).lower()
    for chave, rot in (('painel', 'Painel'), ('estat', 'Estatística'),
                       ('eventos westgard', 'Eventos Westgard'),
                       ('eqa', 'Controle Externo (EQA)'),
                       ('gerencial', 'Visão Gerencial')):
        reg('pagina presente: %s' % rot, chave in tem)

    print()
    print('%d OK, %d FALHA' % (ok, fail))
    sys.exit(1 if fail else 0)


main()
