# -*- coding: utf-8 -*-
"""tema.py -- a paleta e a escala tipografica, num lugar so (ADR-056).

Antes as cores viviam em hexadecimal espalhado por ux.py, painel.py e os
scripts da Fase 3 -- trocar o visual exigia cacar hex em seis arquivos. Aqui
elas tem nome, e o nome diz PARA QUE serve, nao qual e a cor: quando a paleta
mudar, `FUNDO_CARTAO` continua sendo o fundo do cartao.

ATENCAO AO FORMATO. O Excel guarda cor como inteiro **BGR**, nao RGB: o
vermelho puro e 0x0000FF e o azul puro e 0xFF0000. Use `cor('#13282C')`, que
faz a inversao, em vez de escrever o inteiro na mao -- dois defeitos de cor
deste projeto nasceram de um hex copiado de um site de design direto para o
codigo.
"""


def cor(hexa):
    """'#13282C' -> inteiro BGR que o Excel entende."""
    h = hexa.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


# ---------------------------------------------------------------- marca
TINTA = cor('#13282C')          # texto de titulo e de destaque
MARCA = cor('#1F6F6B')          # verde da casa: cabecalho de tabela, botao
MARCA_CLARA = cor('#3CD6B0')    # realce sobre fundo escuro (aba ativa)
ACAO = cor('#1F6F3C')           # verde de acao (lancar, novo lote)

# ---------------------------------------------------------------- superficie
PAPEL = cor('#FFFFFF')          # fundo das telas
FAIXA = cor('#EDF6F4')          # faixa de secao, fundo de bloco
FUNDO_CARTAO = cor('#F7FAFA')   # cartao de indicador
ZEBRA = cor('#FAFBFB')          # linha alternada de tabela longa
CAMPO = cor('#F2F2F2')          # celula que o usuario preenche

# ---------------------------------------------------------------- traco
REGUA = cor('#D8DEDD')          # regra horizontal entre linhas de tabela
CONTORNO = cor('#DCE6E4')       # contorno de cartao
CAMPO_BORDA = cor('#A6A6A6')    # contorno de celula editavel

# ---------------------------------------------------------------- texto
TEXTO = cor('#262626')
TEXTO_FRACO = cor('#595959')    # nota, legenda, rodape
TEXTO_APAGADO = cor('#BFBFBF')  # zero que nao precisa gritar
SOBRE_MARCA = cor('#FFFFFF')    # texto sobre o verde da casa

# ---------------------------------------------------------------- semantica
OK_TXT = cor('#1F6F1F')
OK_FUNDO = cor('#E3F4EA')
ALERTA_TXT = cor('#7A5000')
ALERTA_FUNDO = cor('#FFF4CE')
ERRO_TXT = cor('#C00000')
ERRO_FUNDO = cor('#FDE2E1')
INFO_TXT = cor('#9C7A00')
FALTA_TXT = cor('#007A9C')      # "sem media/DP": falta dado, nao e erro
FALTA_FUNDO = cor('#CEF4FF')

# ---------------------------------------------------------------- tipografia
# Uma familia so, cinco tamanhos. Mais do que cinco vira ruido; menos nao da
# hierarquia. Segoe UI e a fonte do proprio Windows -- existe em toda maquina
# onde este arquivo vai abrir, e nao tem o ar de planilha que o Calibri tem.
FONTE = 'Segoe UI'
TITULO = 16        # titulo da tela (linha 1)
SECAO = 11         # titulo de bloco
CORPO = 10         # dado de tabela
MIUDO = 9          # cabecalho de tabela, legenda de grafico
NOTA = 8           # rodape, dica, rotulo de cartao
CARTAO = 14        # o numero grande do cartao de indicador
