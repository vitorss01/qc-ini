# -*- coding: utf-8 -*-
"""corrigir_espec_efetiva.py - o rotulo de especificacao volta a dizer a verdade

O DEFEITO

A coluna Especificacao_Efetiva decidia assim:

    if [Situacao_Especificacao] = "CADASTRADA" then [Fonte_Especificacao]
    else "Sem meta"

O ADR-028 passou a gravar "ANALITOS_VIGENTE" nessa coluna. A comparacao
literal deixou de casar e TODOS os analitos passaram a ser rotulados
"Sem meta" -- inclusive os 62 grupos que TEM meta. No painel do gestor isso
aparece como "Especificacao em uso: Sem meta", ou seja, o dashboard afirmando
que o laboratorio nao tem criterio de qualidade nenhum.

O erro de fundo foi acoplar o rotulo a um VOCABULARIO (o texto da situacao)
em vez de acopla-lo ao FATO (existe ETp?). Vocabulario muda -- e mudou.

A CORRECAO

A meta existe quando ha ETp e ha fonte declarada. E o mesmo criterio que as
medidas [Grupos com Meta] e [Sigma] ja usam, entao o cartao passa a concordar
com os numeros ao lado dele em vez de contradize-los.

Ausencia de meta continua sendo um terceiro estado explicito, nunca uma
aprovacao silenciosa (ADR-023).

Uso: python corrigir_espec_efetiva.py
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

NOVO = '''// A meta existe quando ha ETp e fonte, nao quando a Situacao tem um
        // texto especifico: o ADR-028 trocou "CADASTRADA" por
        // "ANALITOS_VIGENTE" e a comparacao literal rotulou TODOS os
        // analitos como "Sem meta". Mesmo criterio de [Grupos com Meta].
        if [ETp_pct] <> null
           and [Fonte_Especificacao] <> null
           and Text.Trim(Text.From([Fonte_Especificacao])) <> ""
        then [Fonte_Especificacao] else "Sem meta", type text)'''

ALVOS = [
    os.path.join(RAIZ, '_arquitetura', 'scripts_fase3', 'gerar_pbip.py'),
    os.path.join(RAIZ, 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
                 'definition', 'tables', 'Fato_QC.tmdl'),
]

# O recuo difere entre os arquivos: o gerador usa espacos e o .tmdl usa TAB
# dentro do bloco da particao. O padrao aceita os dois e o recuo capturado e
# reaplicado a cada linha do trecho novo -- TMDL e sensivel a indentacao.
PADRAO = re.compile(
    r'([ \t]*)if \[Situacao_Especificacao\] = "CADASTRADA"\r?\n'
    r'[ \t]*then \[Fonte_Especificacao\] else "Sem meta", type text\)'
)


def substituir(texto):
    m = PADRAO.search(texto)
    if not m:
        return None
    recuo = m.group(1)
    novo = NOVO.replace('\n        ', '\n' + recuo)
    return texto[:m.start()] + recuo + novo + texto[m.end():]


def main():
    mudou = 0
    for caminho in ALVOS:
        nome = os.path.basename(caminho)
        if not os.path.isfile(caminho):
            print('  ausente: %s' % nome)
            continue
        texto = io.open(caminho, encoding='utf-8').read()
        saida = substituir(texto)
        if saida is None:
            if 'ETp_pct] <> null' in texto:
                print('  ja corrigido: %s' % nome)
            else:
                print('  NAO CASOU o padrao em %s' % nome)
            continue
        io.open(caminho, 'w', encoding='utf-8', newline='').write(saida)
        print('  corrigido: %s' % nome)
        mudou += 1
    print()
    print('arquivos alterados: %d' % mudou)
    return 0


if __name__ == '__main__':
    sys.exit(main())
