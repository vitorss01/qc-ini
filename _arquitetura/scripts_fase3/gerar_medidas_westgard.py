# -*- coding: utf-8 -*-
"""gerar_medidas_westgard.py - colunas e medidas de Westgard, a partir do contrato

POR QUE GERAR

Depois do ADR-047 o contrato tem OITO colunas de Westgard -- a uniao dos dois
produtos, cada um preenchendo as cinco da sua area. Manter a mao 8 colunas
W_*, 8 colunas Usar_* e 8 familias de 6 medidas e 64 blocos TMDL que precisam
concordar entre si e com o contrato. Foi manutencao manual que deixou
'Violada 8x' lendo Fato_QC[W_10x] por meses.

O MODELO NAO GANHA UMA MEDIDA GENERICA

Cada regra tem as suas medidas, com o nome dela. O §3 proibe coluna generica
de significado variavel, e uma medida 'Viol Regra 2' que muda de sentido
conforme a area seria a mesma armadilha com outro nome.

Isso funciona porque as paginas ja sao por produto: a pagina da Bioquimica
liga 1_3s/2_2s/R_4s/4_1s/8x e a da Hematologia liga
1_3s/2of3_2s/R_4s/3_1s/6x. O §20 sai de graca -- nao ha condicional visual,
ha a medida certa em cada pagina.

IDEMPOTENTE

lineageTag deriva do nome por UUID5, entao rodar duas vezes produz o mesmo
arquivo. Sem isso cada execucao trocaria os identificadores e o Power BI
trataria as colunas como novas, perdendo a ligacao com os visuais.

Uso:
    python gerar_medidas_westgard.py --mostrar
    python gerar_medidas_westgard.py --aplicar
"""
import io
import json
import os
import re
import sys
import uuid

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONTRATO = os.path.join(RAIZ, 'CONTRATO_BI.json')
TMDL = os.path.normpath(os.path.join(
    RAIZ, '..', 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
    'definition', 'tables', 'Fato_QC.tmdl'))

NS = uuid.UUID('7f3a1c22-9d84-4e51-8b60-qc1ni0000000'.replace('q', 'a')
               .replace('c1ni', 'c1a4'))

FAMILIAS = ['Viol', 'Recomendada', 'Violada', 'Violacoes', 'Estado', 'Cor',
            'Rotulo']


def tag(nome):
    return str(uuid.uuid5(NS, nome))


def rotulo(regra):
    """1_3s -> 1-3s, mantendo a convencao ja usada nas paginas."""
    return regra.replace('_', '-')


def blocos_coluna(regras):
    fora = []
    for r in regras:
        for pref, tipo in (('W_', 'int64'), ('Usar_', 'boolean')):
            nome = pref + r
            fora.append(
                '\tcolumn %s\n'
                '\t\tdataType: %s\n'
                '\t\tlineageTag: %s\n'
                '\t\tsummarizeBy: none\n'
                '\t\tsourceColumn: %s\n'
                '\n'
                '\t\tannotation SummarizationSetBy = Automatic\n'
                % (nome, tipo, tag('col:' + nome), nome))
    return '\n'.join(fora)


def blocos_medida(regras):
    fora = []
    for r in regras:
        W = 'Fato_QC[W_%s]' % r
        U = 'Fato_QC[Usar_%s]' % r
        m = [
            ("Viol %s" % r,
             "SUM ( %s )" % W,
             "\t\tformatString: #,0\n\t\tdisplayFolder: 05 Westgard\n"),
            ("Recomendada %s" % r,
             "IF ( SELECTEDVALUE ( %s ) = TRUE (), 1, 0 )" % U,
             "\t\tformatString: 0\n\t\tdisplayFolder: 11 Regras recomendadas\n"),
            ("Violada %s" % r,
             "VAR ultRun = MAX ( Fato_QC[RUN] ) RETURN IF ( CALCULATE ( SUM ( %s ), "
             "Fato_QC[RUN] = ultRun ) > 0, 1, 0 )" % W,
             "\t\tformatString: 0\n\t\tdisplayFolder: 11 Regras recomendadas\n"),
            ("Violacoes %s no periodo" % r,
             "SUM ( %s )" % W,
             "\t\tformatString: #,0\n\t\tdisplayFolder: 11 Regras recomendadas\n"),
            ("Estado %s" % r,
             "VAR rec = [Recomendada %s] VAR vio = [Violada %s] RETURN SWITCH ( TRUE (), "
             "rec = 1 && vio = 1, \"RECOMENDADA - VIOLADA\", rec = 1, \"RECOMENDADA\", "
             "vio = 1, \"FORA DO PLANO - VIOLADA\", \"fora do plano\" )" % (r, r),
             "\t\tdisplayFolder: 11 Regras recomendadas\n"),
            # VIOLACAO > RECOMENDACAO > NEUTRO (§11): o vermelho vence o verde,
            # e violada fora do plano ainda alerta em laranja.
            ("Cor %s" % r,
             "VAR e = [Estado %s] RETURN SWITCH ( e, \"RECOMENDADA - VIOLADA\", \"#C0392B\", "
             "\"FORA DO PLANO - VIOLADA\", \"#D68910\", \"RECOMENDADA\", \"#1E8449\", "
             "\"#9AA0A6\" )" % r,
             "\t\tdisplayFolder: 11 Regras recomendadas\n"),
            ("Rotulo %s" % r, '"%s"' % rotulo(r),
             "\t\tdisplayFolder: 11 Regras recomendadas\n"),
        ]
        for nome, expr, extra in m:
            fora.append("\tmeasure '%s' = %s\n%s\t\tlineageTag: %s\n"
                        % (nome, expr, extra, tag('med:' + nome)))
    return '\n'.join(fora)


def limpar(t, regras_todas):
    """Remove colunas e medidas das familias, inclusive as de nome legado."""
    legado = ['1_3s', '2_2s', 'R_4s', '4_1s', '10x', '8x', '2of3_2s', '3_1s', '6x']
    alvos = sorted(set(regras_todas + legado))
    n = 0

    for r in alvos:
        for pref in ('W_', 'Usar_'):
            pad = re.compile(
                r'\n\tcolumn %s%s\n(?:\t\t.*\n|\n(?=\t\t))*' % (pref, re.escape(r)))
            t, k = pad.subn('\n', t)
            n += k
        for fam in FAMILIAS:
            nome = ('%s %s no periodo' % (fam, r)) if fam == 'Violacoes' \
                else ('%s %s' % (fam, r))
            pad = re.compile(
                r"\n\tmeasure '%s' = .*\n(?:\t\t.*\n|\n(?=\t\t))*"
                % re.escape(nome))
            t, k = pad.subn('\n', t)
            n += k
    return t, n


def main():
    c = json.load(io.open(CONTRATO, encoding='utf-8'))
    ordem = c['ordem_das_regras']
    regras = []
    for a in ('bioquimica', 'hematologia'):
        for r in ordem[a]:
            if r not in regras:
                regras.append(r)

    cols = blocos_coluna(regras)
    meds = blocos_medida(regras)
    total = ' + '.join('[Viol %s]' % r for r in regras)

    if '--mostrar' in sys.argv:
        print('regras (%d): %s' % (len(regras), ', '.join(regras)))
        print()
        print(cols[:600] + '\n...')
        print()
        print(meds[:900] + '\n...')
        print()
        print("Viol Total = %s" % total)
        return 0

    if '--aplicar' not in sys.argv:
        print(__doc__)
        return 2

    txt = io.open(TMDL, encoding='utf-8', newline='').read()
    nl = '\r\n' if '\r\n' in txt else '\n'
    t = txt.replace('\r\n', '\n')

    t, removidos = limpar(t, regras)
    print('removidos %d blocos das familias antigas' % removidos)

    # Viol Total passa a somar as OITO: as do outro produto vem em branco, e
    # branco + numero em DAX e o numero. Fica correto nos dois sem condicional.
    t = re.sub(r"\n\tmeasure 'Viol Total' = .*\n",
               "\n\tmeasure 'Viol Total' = %s\n" % total, t)

    # reinsere logo antes da primeira medida existente, mantendo o arquivo
    # organizado por blocos
    marca = "\n\tmeasure '"
    pos = t.index(marca)
    t = t[:pos] + '\n' + cols + '\n' + meds + t[pos:]

    io.open(TMDL, 'w', encoding='utf-8', newline=nl).write(t)
    print('Fato_QC.tmdl: %d regras x (2 colunas + 7 medidas) regeradas'
          % len(regras))
    return 0


if __name__ == '__main__':
    sys.exit(main())
