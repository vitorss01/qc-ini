# -*- coding: utf-8 -*-
"""apontar_fontes_pbip.py - os parametros de caminho do PBIP nesta maquina

POR QUE ESTE SCRIPT EXISTE

O PBIP e versionado e viaja entre maquinas. Os dois parametros de caminho
(pCaminhoQC e pCaminhoHema) sao a unica coisa dentro dele que NAO pode ser
igual nos dois lados: o perfil do usuario muda. Corrigir isso a mao no
expressions.tmdl seria edicao manual no produto final, que e exatamente o
que o ADR-021 proibe -- entao a correcao entra por script.

DOIS DEFEITOS QUE ESTE SCRIPT CONSERTA

1. Perfil errado. Os caminhos apontavam para C:\\Users\\vitor\\..., o perfil
   da outra maquina. Aqui o usuario e vitor.santos e o refresh falhava.

2. Escape duplicado. Em M a barra invertida NAO e caractere de escape:
   "C:\\\\Users\\\\vitor" e um caminho com barras DUPLAS, que nao existe.
   pCaminhoQC estava correto (barra simples) e pCaminhoHema estava com barra
   dupla -- os dois lados do mesmo commit, escritos de formas diferentes.

ESCOLHA DO ARTEFATO

Ha mais de um build de Bioquimica no disco e eles NAO sao equivalentes:
o de 13/08 carrega 34 colunas e o lote de fixture QC-99999901; o de 19/08
carrega as 60 colunas do contrato ADR-026 e o lote real. O script escolhe
pelo conteudo -- numero de colunas e ausencia de fixture -- e nao pela data
ou pelo nome da pasta, que sao pistas e nao provas.

Uso:
    python apontar_fontes_pbip.py                 # confere e corrige
    python apontar_fontes_pbip.py --conferir      # so confere, nao grava
"""
import argparse
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
EXPR = os.path.join(RAIZ, 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
                    'definition', 'expressions.tmdl')

PERFIL = os.environ.get('USERPROFILE', os.path.expanduser('~'))

# Candidatos por produto, em ordem de preferencia. A ordem e uma dica; quem
# decide e a conferencia de conteudo.
CANDIDATOS = {
    'pCaminhoQC': [
        os.path.join(PERFIL, 'QCINI_build_hardening1_BI_Bioquimica', 'QC_Bioquimica.xlsm'),
        os.path.join(PERFIL, 'QCINI_build_hardening1_Bioquimica', 'QC_Bioquimica.xlsm'),
        os.path.join(PERFIL, 'QCINI_build_hardening1', 'QC_Bioquimica.xlsm'),
    ],
    'pCaminhoHema': [
        os.path.join(PERFIL, 'QCINI_build_hardening1_Hematologia', 'QC_Hematologia.xlsm'),
        os.path.join(PERFIL, 'QCINI_build_hardening1', 'QC_Hematologia.xlsm'),
    ],
}


def inspecionar(caminho):
    """(existe, n_colunas, tem_fixture) lendo o .xlsm como zip -- sem abrir o
    Excel, que custa dezenas de segundos por arquivo."""
    if not os.path.isfile(caminho):
        return False, 0, False
    import zipfile
    try:
        with zipfile.ZipFile(caminho) as z:
            nomes = z.namelist()
            tabelas = [n for n in nomes if n.startswith('xl/tables/')]
            for t in tabelas:
                xml = z.read(t).decode('utf-8', 'replace')
                if 'tblBI_Fato' not in xml:
                    continue
                cols = len(re.findall(r'<tableColumn\b', xml))
                # o lote de fixture aparece na cadeia de texto compartilhada
                fixture = False
                if 'xl/sharedStrings.xml' in nomes:
                    ss = z.read('xl/sharedStrings.xml').decode('utf-8', 'replace')
                    fixture = 'QC-99999' in ss
                return True, cols, fixture
    except Exception as e:
        print('   aviso: nao consegui ler %s (%s)' % (caminho, e))
    return False, 0, False


def escolher(param):
    print('\n%s' % param)
    melhor = None
    for c in CANDIDATOS[param]:
        tem, cols, fix = inspecionar(c)
        if not os.path.isfile(c):
            print('   ausente   %s' % c)
            continue
        marca = 'tblBI_Fato %d col%s' % (cols, ', COM FIXTURE' if fix else '') \
            if tem else 'SEM tblBI_Fato'
        print('   %-28s %s' % (marca, c))
        # Criterio: mais colunas vence (contrato mais novo); empate vai para o
        # build mais recente. Sem o desempate, dois builds de 60 colunas
        # fariam a ordem da lista decidir -- e a ordem da lista nao sabe qual
        # deles nasceu depois do ADR-028.
        if tem and not fix:
            chave = (cols, os.path.getmtime(c))
            if melhor is None or chave > melhor[1]:
                melhor = (c, chave)
    if melhor:
        import datetime
        quando = datetime.datetime.fromtimestamp(melhor[1][1])
        print('   -> escolhido: %s (%d colunas, build de %s)'
              % (melhor[0], melhor[1][0], quando.strftime('%d/%m %H:%M')))
        return melhor[0]
    print('   -> NENHUM candidato serve')
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--conferir', action='store_true',
                    help='so relata, nao grava')
    a = ap.parse_args()

    if not os.path.isfile(EXPR):
        raise SystemExit('nao encontrei %s' % EXPR)
    texto = io.open(EXPR, encoding='utf-8').read()

    print('expressions.tmdl: %s' % EXPR)
    print('perfil desta maquina: %s' % PERFIL)

    novos, faltando = {}, []
    for param in ('pCaminhoQC', 'pCaminhoHema'):
        c = escolher(param)
        if c:
            novos[param] = c
        else:
            faltando.append(param)

    print()
    mudou = False
    for param, caminho in novos.items():
        # barra simples: em M a barra invertida e literal, nao escape
        literal = caminho.replace('"', '""')
        padrao = re.compile(r'(expression\s+%s\s*=\s*)"[^"]*"' % re.escape(param))
        if not padrao.search(texto):
            print('   nao achei a expressao %s no arquivo' % param)
            continue
        atual = padrao.search(texto).group(0)
        novo = '%s"%s"' % (padrao.search(texto).group(1), literal)
        if atual == novo:
            print('   %s ja esta correto' % param)
            continue
        texto = padrao.sub(lambda m: '%s"%s"' % (m.group(1), literal), texto, count=1)
        print('   %s -> %s' % (param, caminho))
        mudou = True

    if faltando:
        print()
        print('PENDENTE: sem artefato utilizavel para: %s' % ', '.join(faltando))
        print('   o Table.Combine da fato e rigido: se um lado falta, o refresh')
        print('   inteiro falha -- nao so as paginas daquele produto.')

    if a.conferir:
        print('\n(--conferir: nada gravado)')
    elif mudou:
        io.open(EXPR, 'w', encoding='utf-8', newline='\r\n').write(texto)
        print('\nexpressions.tmdl atualizado')
    else:
        print('\nnada a mudar')

    return 1 if faltando else 0


if __name__ == '__main__':
    sys.exit(main())
