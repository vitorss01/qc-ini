# -*- coding: utf-8 -*-
"""gerar_paginas_bi.py - as duas paginas que faltavam no relatorio

Acrescenta EVENTOS WESTGARD e CONTROLE EXTERNO (EQA) ao PBIR existente, sem
tocar nas cinco paginas que ja funcionam.

POR QUE GERADO POR SCRIPT, E NAO A MAO

Sao ~40 arquivos visual.json com o mesmo esqueleto. Escritos um a um, a chance
de um campo divergir e alta -- e um visualContainer malformado impede o
relatorio INTEIRO de abrir, nao so aquele visual. Aqui o esqueleto e um so.

O QUE NAO SE FAZ AQUI

Nenhum rotulo posicional (R2/R4/R5) aparece: na pagina de eventos a regra vem
de Fato_Eventos_Westgard[Regra], que ja e o nome real por area; onde se quer o
nome pela dimensao, usa-se Dim_Regra_Westgard[Regra].
"""
import io
import os
import json
import hashlib
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
DEF = os.path.join(BASE, 'PowerBI', 'QC_INI_Bioquimica.Report', 'definition')
PAGES = os.path.join(DEF, 'pages')

SCH_VIS = ('https://developer.microsoft.com/json-schemas/fabric/item/report/'
           'definition/visualContainer/2.11.0/schema.json')
SCH_PAG = ('https://developer.microsoft.com/json-schemas/fabric/item/report/'
           'definition/page/1.0.0/schema.json')

AZUL = "'#1F3864'"


def ident(txt):
    return hashlib.sha1(txt.encode('utf-8')).hexdigest()[:20]


def campo(entidade, prop, medida=False, ativo=False):
    tipo = 'Measure' if medida else 'Column'
    d = {
        'field': {tipo: {'Expression': {'SourceRef': {'Entity': entidade}},
                         'Property': prop}},
        'queryRef': '%s.%s' % (entidade, prop),
        'nativeQueryRef': prop,
    }
    if ativo:
        d['active'] = True
    return d


def rotular(proj, rotulo):
    proj['nativeQueryRef'] = rotulo
    return proj


def visual(nome, tipo, x, y, w, h, z, queryState, objects=None):
    v = {
        '$schema': SCH_VIS,
        'name': nome,
        'position': {'x': x, 'y': y, 'z': z, 'height': h, 'width': w,
                     'tabOrder': z},
        'visual': {'visualType': tipo, 'query': {'queryState': queryState}},
    }
    if objects:
        v['visual']['objects'] = objects
    return v


def card(nome, x, y, w, h, z, ent, medida, rotulo):
    obj = {
        'labels': [{'properties': {
            'color': {'solid': {'color': {'expr': {'Literal': {'Value': AZUL}}}}},
            'fontSize': {'expr': {'Literal': {'Value': '21D'}}},
            'bold': {'expr': {'Literal': {'Value': 'true'}}},
            'fontFamily': {'expr': {'Literal': {'Value': "'Segoe UI'"}}}}}],
    }
    return visual(nome, 'card', x, y, w, h, z,
                  {'Values': {'projections': [
                      rotular(campo(ent, medida, medida=True), rotulo)]}}, obj)


def slicer(nome, x, y, w, h, z, ent, col, rotulo):
    return visual(nome, 'slicer', x, y, w, h, z,
                  {'Values': {'projections': [
                      rotular(campo(ent, col, ativo=True), rotulo)]}})


def texto(nome, x, y, w, h, z, txt, tam='12pt'):
    obj = {'general': [{'properties': {'paragraphs': [{'textRuns': [
        {'value': txt, 'textStyle': {'fontSize': tam, 'fontWeight': 'bold',
                                     'color': '#1F3864',
                                     'fontFamily': 'Segoe UI'}}]}]}}]}
    return visual(nome, 'textbox', x, y, w, h, z, {}, obj)


def barras(nome, x, y, w, h, z, entC, colC, rotC, entY, medY, rotY,
           tipo='clusteredBarChart'):
    return visual(nome, tipo, x, y, w, h, z, {
        'Category': {'projections': [rotular(campo(entC, colC, ativo=True), rotC)]},
        'Y': {'projections': [rotular(campo(entY, medY, medida=True), rotY)]}})


def tabela(nome, x, y, w, h, z, colunas):
    proj = []
    for ent, prop, med, rot in colunas:
        proj.append(rotular(campo(ent, prop, medida=med), rot))
    return visual(nome, 'tableEx', x, y, w, h, z,
                  {'Values': {'projections': proj}})


def gravar_pagina(nome_exib, visuais):
    pid = ident(nome_exib)
    d = os.path.join(PAGES, pid)
    os.makedirs(os.path.join(d, 'visuals'), exist_ok=True)
    io.open(os.path.join(d, 'page.json'), 'w', encoding='utf-8',
            newline='\n').write(json.dumps({
                '$schema': SCH_PAG, 'name': pid, 'displayName': nome_exib,
                'displayOption': 'FitToPage', 'height': 720, 'width': 1280},
                ensure_ascii=False, indent=2) + '\n')
    for v in visuais:
        vd = os.path.join(d, 'visuals', v['name'])
        os.makedirs(vd, exist_ok=True)
        io.open(os.path.join(vd, 'visual.json'), 'w', encoding='utf-8',
                newline='\n').write(json.dumps(v, ensure_ascii=False,
                                               indent=2) + '\n')
    print('  %-28s %s  (%d visuais)' % (nome_exib, pid, len(visuais)))
    return pid


# ===================== EVENTOS WESTGARD =====================
EV = 'Fato_Eventos_Westgard'
v_ev = [
    texto(ident('ev_tit'), 16, 12, 700, 34,
          'EVENTOS DE WESTGARD  ·  uma linha = um evento', '16pt'),
    texto(ident('ev_nota'), 16, 46, 1248, 30,
          'Um evento pode marcar varios resultados: 6x N3/R2 e UM evento e marca '
          'SEIS. Por isso eventos, resultados marcados e corridas sao numeros '
          'diferentes.', '9pt'),

    slicer(ident('ev_s1'), 16, 86, 200, 150, 1000, EV, 'Area', 'Área'),
    slicer(ident('ev_s2'), 224, 86, 200, 150, 1100, EV, 'Analito', 'Analito'),
    slicer(ident('ev_s3'), 432, 86, 200, 150, 1200, EV, 'Regra', 'Regra'),
    slicer(ident('ev_s4'), 640, 86, 200, 150, 1300, EV, 'Detector', 'Detector'),
    slicer(ident('ev_s5'), 848, 86, 200, 150, 1400, EV, 'Classe', 'Classe'),
    slicer(ident('ev_s6'), 1056, 86, 208, 150, 1500, 'Dim_Data', 'Data', 'Período'),

    card(ident('ev_c1'), 16, 248, 220, 84, 2000, EV,
         'N_Eventos_Violacao', 'Eventos'),
    card(ident('ev_c2'), 244, 248, 240, 84, 2100, EV,
         'N_Resultados_Marcados', 'Resultados marcados'),
    card(ident('ev_c3'), 492, 248, 240, 84, 2200, EV,
         'N_Corridas_Com_Violacao', 'Corridas com violação'),
    card(ident('ev_c4'), 740, 248, 220, 84, 2300, EV,
         'N_Eventos_Oficiais', 'Eventos oficiais'),

    barras(ident('ev_b1'), 16, 344, 480, 200, 3000,
           EV, 'Regra', 'Regra', EV, 'N_Eventos_Violacao', 'Eventos'),
    barras(ident('ev_b2'), 504, 344, 480, 200, 3100,
           EV, 'Analito', 'Analito', EV, 'N_Eventos_Violacao', 'Eventos'),
    barras(ident('ev_b3'), 992, 344, 272, 200, 3200,
           EV, 'Escopo', 'Escopo', EV, 'N_Eventos_Violacao', 'Eventos'),

    tabela(ident('ev_t1'), 16, 552, 1248, 156, 4000, [
        (EV, 'Data', False, 'Data'),
        (EV, 'RUN', False, 'RUN'),
        (EV, 'Analito', False, 'Analito'),
        (EV, 'Regra', False, 'Regra'),
        (EV, 'Detector', False, 'Detector'),
        (EV, 'Escopo', False, 'Escopo'),
        (EV, 'Niveis', False, 'Níveis'),
        (EV, 'N_Niveis', False, 'N'),
        (EV, 'R_Corridas', False, 'R'),
        (EV, 'RUN_Inicial', False, 'RUN inicial'),
        (EV, 'RUN_Final', False, 'RUN final'),
        (EV, 'Evidencia', False, 'Evidência'),
        (EV, 'Z_Max', False, 'Z máximo'),
    ]),
]

# ===================== CONTROLE EXTERNO (EQA) =====================
EQ = 'Fato_EQA'
v_eqa = [
    texto(ident('eq_tit'), 16, 12, 800, 34,
          'CONTROLE EXTERNO DA QUALIDADE  ·  CAP e Controllab', '16pt'),
    texto(ident('eq_nota'), 16, 46, 1248, 30,
          'O |Bias| consolidado e a media DENTRO da rodada e depois entre '
          'rodadas. A avaliacao original do provedor e preservada ao lado do '
          'status padronizado.', '9pt'),

    slicer(ident('eq_s1'), 16, 86, 200, 150, 1000, EQ, 'Area', 'Área'),
    slicer(ident('eq_s2'), 224, 86, 200, 150, 1100, EQ, 'Provedor', 'Provedor'),
    slicer(ident('eq_s3'), 432, 86, 200, 150, 1200, EQ, 'Ano', 'Ano'),
    slicer(ident('eq_s4'), 640, 86, 200, 150, 1300, EQ, 'Rodada', 'Rodada'),
    slicer(ident('eq_s5'), 848, 86, 200, 150, 1400, EQ, 'Analito', 'Analito'),
    slicer(ident('eq_s6'), 1056, 86, 208, 150, 1500, EQ, 'Status', 'Avaliação'),

    card(ident('eq_c1'), 16, 248, 196, 84, 2000, EQ,
         'N_EQA_Resultados', 'Resultados'),
    card(ident('eq_c2'), 220, 248, 196, 84, 2100, EQ,
         'N_EQA_Rodadas', 'Rodadas'),
    card(ident('eq_c3'), 424, 248, 196, 84, 2200, EQ,
         'N_Nao_Conformidades', 'Não aceitáveis'),
    card(ident('eq_c4'), 628, 248, 196, 84, 2300, EQ,
         'Aceitabilidade_pct', '% aceitabilidade'),
    card(ident('eq_c5'), 832, 248, 196, 84, 2400, EQ,
         'Bias_Abs_EQA_pct', '|Bias| %'),
    card(ident('eq_c6'), 1036, 248, 228, 84, 2500, EQ,
         'SDI_Medio', 'SDI médio'),

    barras(ident('eq_b1'), 16, 344, 410, 200, 3000,
           EQ, 'Analito', 'Analito', EQ, 'Bias_Abs_EQA_pct', '|Bias| %'),
    barras(ident('eq_b2'), 434, 344, 410, 200, 3100,
           EQ, 'Analito', 'Analito', EQ, 'N_Nao_Conformidades',
           'Não aceitáveis'),
    barras(ident('eq_b3'), 852, 344, 412, 200, 3200,
           EQ, 'Rodada', 'Rodada', EQ, 'Bias_Abs_EQA_pct', '|Bias| %',
           tipo='lineChart'),

    tabela(ident('eq_t1'), 16, 552, 1248, 156, 4000, [
        (EQ, 'Provedor', False, 'Provedor'),
        (EQ, 'Ano', False, 'Ano'),
        (EQ, 'Rodada', False, 'Rodada'),
        (EQ, 'Analito', False, 'Analito'),
        (EQ, 'Amostra', False, 'Amostra'),
        (EQ, 'Resultado', False, 'Resultado'),
        (EQ, 'Valor_Alvo', False, 'Valor alvo'),
        (EQ, 'Bias_pct', False, 'Bias %'),
        (EQ, 'Bias_Abs_pct', False, '|Bias| %'),
        (EQ, 'SDI', False, 'SDI'),
        (EQ, 'Limite_Inferior', False, 'Limite inf.'),
        (EQ, 'Limite_Superior', False, 'Limite sup.'),
        (EQ, 'Avaliacao_Original', False, 'Avaliação do provedor'),
        (EQ, 'Status', False, 'Status'),
    ]),
]


def main():
    print('paginas geradas:')
    p1 = gravar_pagina('Eventos Westgard', v_ev)
    p2 = gravar_pagina('Controle Externo — EQA', v_eqa)

    # entra na ordem, sem mexer nas cinco que ja existem
    pj = os.path.join(PAGES, 'pages.json')
    j = json.load(io.open(pj, encoding='utf-8'))
    ordem = j.get('pageOrder', [])
    for p in (p1, p2):
        if p not in ordem:
            ordem.append(p)
    j['pageOrder'] = ordem
    io.open(pj, 'w', encoding='utf-8', newline='\n').write(
        json.dumps(j, ensure_ascii=False, indent=2) + '\n')
    print()
    print('pages.json: %d paginas na ordem' % len(ordem))


main()
