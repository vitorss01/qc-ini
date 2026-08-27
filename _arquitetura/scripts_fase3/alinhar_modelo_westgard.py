# -*- coding: utf-8 -*-
"""alinhar_modelo_westgard.py - a fato unica volta a aceitar os dois produtos

O PROBLEMA

Fato_QC combina Bioquimica e Hematologia numa fato so (Table.Combine), separadas
pela coluna Produto -- desenho do ADR-026, para nao duplicar as 114 medidas.

Depois que o contrato passou a publicar o nome REAL da regra por area, os dois
lados deixaram de ter o mesmo esquema:

    Bioquimica   W_1_3s  W_2_2s     W_R_4s  W_4_1s  W_8x
    Hematologia  W_1_3s  W_2of3_2s  W_R_4s  W_3_1s  W_6x

Table.Combine com nomes diferentes nao falha: ele UNE, e as colunas que so
existem de um lado ficam meio nulas. Silenciosamente, metade das violacoes
sumiria das medidas.

A DECISAO

As posicoes 1 e 3 tem o MESMO nome nas duas areas (1_3s e R_4s): ficam como
estao. So as posicoes 2, 4 e 5 divergem, e essas passam a ter nome POSICIONAL
na fato -- W_R2, W_R4, W_R5.

Nome posicional nao e perda de informacao, e o contrario: W_R2 nao afirma nada
sobre qual regra e, enquanto um W_2_2s carregando 2of3_2s afirmaria algo FALSO.
O nome real por area vive em Dim_Regra_Westgard, que e onde ele pode variar sem
mentir.

Tambem corrige W_10x, que o modelo ainda declarava: essa coluna nao existe mais
na fonte desde que o nome passou a vir do motor, e uma coluna declarada que a
fonte nao tem QUEBRA a atualizacao do Power BI.
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
MOD = os.path.join(BASE, 'PowerBI', 'QC_INI_Bioquimica.SemanticModel', 'definition')
FATO = os.path.join(MOD, 'tables', 'Fato_QC.tmdl')

# posicao -> (nome na Bioquimica, nome na Hematologia, nome posicional)
POS = {
    2: ('2_2s', '2of3_2s', 'R2'),
    4: ('4_1s', '3_1s', 'R4'),
    5: ('8x', '6x', 'R5'),
}


def ler(p):
    raw = io.open(p, 'rb').read()
    enc = 'utf-8-sig' if raw[:3] == b'\xef\xbb\xbf' else 'utf-8'
    return raw.decode(enc), enc


def gravar(p, s, enc):
    io.open(p, 'w', encoding=enc, newline='\n').write(s)


def main():
    s, enc = ler(FATO)
    orig = s

    # ---- 1. renomeia cada lado ANTES do Table.Combine ---------------------
    # O TMDL usa TAB para indentar e CRLF para terminar linha. Casar por
    # literal com espacos falha; o regex preserva a indentacao real do arquivo.
    mComb = re.search(r'^([ \t]*)Tabela = Table\.Combine\(\{TabelaBio, TabelaHema\}\),',
                      s, re.M)
    assert mComb, 'Table.Combine nao encontrado'
    ind = mComb.group(1)
    velho = mComb.group(0)
    ren_bio = ', '.join('{"W_%s","W_%s"}' % (b, p) for _, (b, _h, p) in sorted(POS.items()))
    ren_bio += ', ' + ', '.join('{"Usar_%s","Usar_%s"}' % (b, p)
                                for _, (b, _h, p) in sorted(POS.items()))
    ren_hem = ', '.join('{"W_%s","W_%s"}' % (h, p) for _, (_b, h, p) in sorted(POS.items()))
    ren_hem += ', ' + ', '.join('{"Usar_%s","Usar_%s"}' % (h, p)
                                for _, (_b, h, p) in sorted(POS.items()))
    novo_bruto = '''// NOME POSICIONAL NAS TRES POSICOES QUE DIVERGEM POR AREA.
//
// As posicoes 1 e 3 chamam-se 1_3s e R_4s nas duas areas e
// ficam como estao. As outras tres tem nome diferente em
// cada produto, e Table.Combine com nomes diferentes NAO
// falha: ele une e deixa meia coluna nula, fazendo metade
// das violacoes sumir das medidas sem nenhum aviso.
//
// W_R2 nao afirma qual regra e; um W_2_2s carregando
// 2of3_2s afirmaria algo falso. O nome real por area vive
// em Dim_Regra_Westgard.
//
// MissingField.Ignore: um arquivo mais antigo, ainda sem o
// rename no motor, entra sem derrubar a carga.
RenBio = Table.RenameColumns(TabelaBio, {%s}, MissingField.Ignore),
RenHema = Table.RenameColumns(TabelaHema, {%s}, MissingField.Ignore),
Tabela = Table.Combine({RenBio, RenHema}),''' % (ren_bio, ren_hem)
    novo = '\n'.join(ind + l if l.strip() else l
                     for l in novo_bruto.split('\n'))
    s = s.replace(velho, novo)

    # ---- 2. lista de tipos e declaracoes de coluna ------------------------
    trocas = []
    for _, (b, _h, p) in sorted(POS.items()):
        trocas.append(('{"W_%s",' % b, '{"W_%s",' % p))
        trocas.append(('{"Usar_%s",' % b, '{"Usar_%s",' % p))
        trocas.append(('column W_%s' % b, 'column W_%s' % p))
        trocas.append(('column Usar_%s' % b, 'column Usar_%s' % p))
        trocas.append(('Fato_QC[W_%s]' % b, 'Fato_QC[W_%s]' % p))
        trocas.append(('Fato_QC[Usar_%s]' % b, 'Fato_QC[Usar_%s]' % p))
    # W_10x era o nome antigo da posicao 5 no modelo
    trocas += [('{"W_10x",', '{"W_R5",'), ('column W_10x', 'column W_R5'),
               ('Fato_QC[W_10x]', 'Fato_QC[W_R5]')]
    feitas = 0
    for a, b in trocas:
        n = s.count(a)
        if n:
            s = s.replace(a, b)
            feitas += n
    print('substituicoes de nome de coluna: %d' % feitas)

    # ---- 3. nomes de medida que carregavam a regra ------------------------
    med = [("'Viol 2_2s'", "'Viol R2'"), ("'Viol 4_1s'", "'Viol R4'"),
           ("'Viol 10x'", "'Viol R5'"), ("'Viol 8x'", "'Viol R5'")]
    for a, b in med:
        if a in s:
            s = s.replace(a, b)
            print('medida renomeada: %s -> %s' % (a, b))

    if s == orig:
        print('nada mudou')
    else:
        gravar(FATO, s, enc)
        print('Fato_QC.tmdl gravado')

    # ---- 4. dimensao com o nome real por area ----------------------------
    linhas = []
    nomes = {1: ('1_3s', '1_3s', 'R1'), 3: ('R_4s', 'R_4s', 'R3')}
    nomes.update(POS)
    for pos in sorted(nomes):
        b, h, p = nomes[pos]
        linhas.append('{"Bioquimica", %d, "R%d", "%s"}' % (pos, pos, b))
        linhas.append('{"Hematologia", %d, "R%d", "%s"}' % (pos, pos, h))
    corpo = ',\n                            '.join(linhas)

    dim = '''table Dim_Regra_Westgard

	/// Nome da regra de Westgard em cada area, por posicao.
	///
	/// A fato guarda as posicoes que divergem com nome neutro (W_R2, W_R4,
	/// W_R5) porque as duas areas rodam matrizes diferentes e uma coluna
	/// chamada W_2_2s carregando 2of3_2s seria uma afirmacao falsa. O nome
	/// verdadeiro vive aqui, onde pode variar por produto sem mentir.
	///
	/// Bioquimica  (2 niveis)  1_3s  2_2s      R_4s  4_1s  8x
	/// Hematologia (3 niveis)  1_3s  2of3_2s   R_4s  3_1s  6x
	///
	/// 10x nao aparece: nao existe mais em lugar nenhum operacional desde o
	/// ADR-041.

	column Produto
		dataType: string
		summarizeBy: none
		sourceColumn: Produto

	column Posicao
		dataType: int64
		formatString: 0
		summarizeBy: none
		sourceColumn: Posicao

	column Slot
		dataType: string
		summarizeBy: none
		sourceColumn: Slot

	column Regra
		dataType: string
		summarizeBy: none
		sourceColumn: Regra

	partition Dim_Regra_Westgard = m
		mode: import
		source =
				let
				    Linhas = #table(
				        type table [Produto = text, Posicao = Int64.Type,
				                    Slot = text, Regra = text],
				        {
                            %s
				        })
				in
				    Linhas

	annotation PBI_ResultType = Table
''' % corpo
    p = os.path.join(MOD, 'tables', 'Dim_Regra_Westgard.tmdl')
    io.open(p, 'w', encoding='utf-8', newline='\n').write(dim)
    print('Dim_Regra_Westgard.tmdl criada (%d linhas de dado)' % len(linhas))


main()
