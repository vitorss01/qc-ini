# -*- coding: utf-8 -*-
"""estender_modelo_plano_qc.py - o plano de CQ chega ao modelo semantico

O QUE ESTE SCRIPT FAZ

Leva as oito colunas do ADR-040 (Sigma_Plano, Nivel_Governante,
Classificacao_Sigma_Plano e as cinco Usar_*) do contrato tblBI_Fato para
dentro do PBIP: tipagem no Power Query, colunas no TMDL e as medidas que os
visuais consomem.

O QUE ELE DELIBERADAMENTE NAO FAZ

Nao recalcula faixa de Sigma, nao decide qual nivel governa e nao monta a
lista de regras. Tudo isso ja vem decidido do motor (ADR-019/ADR-040). As
medidas aqui SELECIONAM e ROTULAM -- um SWITCH de faixa escrito em DAX seria
a segunda copia da regra, e duas copias divergem no primeiro ajuste.

Por isso as medidas usam SELECTEDVALUE e nao AVERAGE/MIN sobre as colunas do
plano: o plano e um ATRIBUTO do grupo (analito, lote), nao uma grandeza para
agregar. Com varios analitos no filtro, SELECTEDVALUE devolve vazio e a tela
diz "selecione um analito" -- que e a verdade -- em vez de exibir a media de
run sizes de analitos diferentes, que nao significa nada.

RECOMENDADA x VIOLADA

Sao duas dimensoes distintas e o roteiro de validacao exige que nao se
confundam:

  Recomendada_X  vem de Usar_X    -- o plano que o Sigma exige
  Violada_X      vem de W_X       -- o que de fato aconteceu na corrida

Uma regra pode estar recomendada e nao violada, violada e nao recomendada
(violacao historica de uma regra fora do plano atual), ou as duas coisas.

Uso: python estender_modelo_plano_qc.py
"""
import io
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TMDL = os.path.join(RAIZ, 'PowerBI', 'QC_INI_Bioquimica.SemanticModel',
                    'definition', 'tables', 'Fato_QC.tmdl')

TEXTO, INT, REAL, BOOL = 'string', 'int64', 'double', 'boolean'

# (nome, tipo no modelo, tipo no M, oculta)
NOVAS = [
    ('Sigma_Plano', REAL, 'type number', False),
    ('Nivel_Governante', TEXTO, 'type text', False),
    ('Classificacao_Sigma_Plano', TEXTO, 'type text', False),
    ('Usar_1_3s', BOOL, 'type logical', False),
    ('Usar_2_2s', BOOL, 'type logical', False),
    ('Usar_R_4s', BOOL, 'type logical', False),
    ('Usar_4_1s', BOOL, 'type logical', False),
    ('Usar_8x', BOOL, 'type logical', False),
]

# Colunas do ADR-035 que ja existem no contrato mas nunca chegaram ao modelo.
# Sem elas os cartoes de plano ficariam sem origem.
ADR035 = [
    ('DPM_Teorico', REAL, 'type number', False),
    ('Yield_Teorico', REAL, 'type number', False),
    ('Regra_Westgard_Recomendada', TEXTO, 'type text', False),
    ('N_Controle_Recomendado', TEXTO, 'type text', False),
    ('RunSize_Max_Recomendado', TEXTO, 'type text', False),
    ('Frequencia_QC_Descricao', TEXTO, 'type text', False),
    ('Cobertura_Motor_Westgard', TEXTO, 'type text', True),
    ('Referencia_Plano_QC', TEXTO, 'type text', True),
    ('Classificacao_Sigma', TEXTO, 'type text', True),
    ('Bias_Observado_abs_pct', REAL, 'type number', False),
    ('Margem_ETp_pp', REAL, 'type number', False),
    ('Margem_ETp_pct', REAL, 'type number', False),
    ('Status_Margem_ETp', TEXTO, 'type text', False),
    ('Provedor_EQA', TEXTO, 'type text', False),
    ('Ano_EQA', INT, 'Int64.Type', False),
    ('Rodada_EQA', TEXTO, 'type text', False),
]

REGRAS = [('1_3s', 'W_1_3s'), ('2_2s', 'W_2_2s'), ('R_4s', 'W_R_4s'),
          ('4_1s', 'W_4_1s'), ('8x', 'W_10x')]

VERDE, AMBAR, VERMELHO, CINZA, AZUL = ('#1E8449', '#D68910', '#C0392B',
                                       '#9AA0A6', '#1F3864')


def guid(semente):
    import uuid
    return str(uuid.uuid5(uuid.NAMESPACE_URL, 'qcini/pbip/' + semente))


def bloco_coluna(nome, tipo, oculta):
    ln = ['\tcolumn %s' % nome,
          '\t\tdataType: %s' % tipo]
    if oculta:
        ln.append('\t\tisHidden')
    ln.append('\t\tlineageTag: %s' % guid('col/Fato_QC/' + nome))
    ln.append('\t\tsummarizeBy: none')
    ln.append('\t\tsourceColumn: %s' % nome)
    ln.append('')
    ln.append('\t\tannotation SummarizationSetBy = Automatic')
    ln.append('')
    return '\n'.join(ln)


def medidas():
    """(nome, dax, formato, pasta). DAX de uma linha sai inline no TMDL; o
    bloco recuado faria o analisador engolir dataType e lineageTag para dentro
    da formula -- foi assim que Faixa_Sigma nasceu em erro."""
    m = []
    m.append(('Sigma Plano', 'SELECTEDVALUE ( Fato_QC[Sigma_Plano] )',
              '#,0.00', '10 Plano de CQ'))
    m.append(('Nivel Governante',
              'VAR v = SELECTEDVALUE ( Fato_QC[Nivel_Governante] ) '
              'RETURN IF ( ISBLANK ( v ), "SEM DADOS", v )',
              None, '10 Plano de CQ'))
    m.append(('Classificacao do Plano',
              'VAR v = SELECTEDVALUE ( Fato_QC[Classificacao_Sigma_Plano] ) '
              'RETURN IF ( ISBLANK ( v ), "SEM DADOS", v )',
              None, '10 Plano de CQ'))
    m.append(('Regras Recomendadas',
              'VAR v = SELECTEDVALUE ( Fato_QC[Regra_Westgard_Recomendada] ) '
              'RETURN IF ( ISBLANK ( v ) || v = "", "--", v )',
              None, '10 Plano de CQ'))
    # N e run size chegam como TEXTO porque a faixa abaixo de 3 Sigma os deixa
    # vazios de proposito. Converter para numero transformaria o vazio em 0 --
    # "rode zero controles", que e o oposto de "nao ha plano automatico".
    m.append(('N Controle',
              'VAR v = SELECTEDVALUE ( Fato_QC[N_Controle_Recomendado] ) '
              'RETURN IF ( ISBLANK ( v ) || v = "", "--", v )',
              None, '10 Plano de CQ'))
    m.append(('Run Size',
              'VAR v = SELECTEDVALUE ( Fato_QC[RunSize_Max_Recomendado] ) '
              'RETURN IF ( ISBLANK ( v ) || v = "", "--", v )',
              None, '10 Plano de CQ'))
    m.append(('Frequencia QC',
              'VAR v = SELECTEDVALUE ( Fato_QC[Frequencia_QC_Descricao] ) '
              'RETURN IF ( ISBLANK ( v ) || v = "", "--", v )',
              None, '10 Plano de CQ'))
    # O status do PLANO nao e o status da corrida. Sigma baixo pede metodo
    # melhor; nao reprova a corrida de hoje, que e governada por Westgard.
    m.append(('Status Plano QC',
              'VAR c = SELECTEDVALUE ( Fato_QC[Classificacao_Sigma_Plano] ) '
              'RETURN SWITCH ( TRUE (), ISBLANK ( c ), "SEM DADOS", '
              'c = "SEM DADOS", "SEM DADOS", '
              'c = "Desempenho inadequado", '
              '"DESEMPENHO INADEQUADO - REAVALIAR METODO", UPPER ( c ) )',
              None, '10 Plano de CQ'))
    m.append(('DPM Teorico', 'SELECTEDVALUE ( Fato_QC[DPM_Teorico] )',
              '#,0', '10 Plano de CQ'))
    m.append(('Rendimento Teorico', 'SELECTEDVALUE ( Fato_QC[Yield_Teorico] )',
              '0.00%', '10 Plano de CQ'))
    m.append(('Sigma N1',
              'CALCULATE ( SELECTEDVALUE ( Fato_QC[Sigma_Obs] ), Fato_QC[Nivel] = 1 )',
              '#,0.00', '10 Plano de CQ'))
    m.append(('Sigma N2',
              'CALCULATE ( SELECTEDVALUE ( Fato_QC[Sigma_Obs] ), Fato_QC[Nivel] = 2 )',
              '#,0.00', '10 Plano de CQ'))

    for rg, wcol in REGRAS:
        m.append(('Recomendada %s' % rg,
                  'IF ( SELECTEDVALUE ( Fato_QC[Usar_%s] ) = TRUE (), 1, 0 )' % rg,
                  '0', '11 Regras recomendadas'))
        # "VIOLADA NA CORRIDA", e nao "violada algum dia". Somar W_* sobre o
        # historico inteiro marca TODAS as regras como violadas -- com 110
        # resultados por analito em seis meses, qualquer regra dispara em
        # algum momento, e a tela inteira fica vermelha sem informar nada.
        # O escopo e a ULTIMA corrida do contexto, coerente com [Veredito
        # Atual] e [Z Score Atual].
        m.append(('Violada %s' % rg,
                  'VAR ultRun = MAX ( Fato_QC[RUN] ) '
                  'RETURN IF ( CALCULATE ( SUM ( Fato_QC[%s] ), '
                  'Fato_QC[RUN] = ultRun ) > 0, 1, 0 )' % wcol,
                  '0', '12 Regras violadas'))
        # A contagem historica continua disponivel, com OUTRO nome: serve a
        # tabela de Westgard do periodo, nao ao semaforo do plano.
        m.append(('Violacoes %s no periodo' % rg,
                  'SUM ( Fato_QC[%s] )' % wcol, '#,0', '12 Regras violadas'))
        # QUATRO estados, nao dois. Recomendacao e violacao sao dimensoes
        # independentes (secoes 8 e 10 do roteiro): uma regra violada que NAO
        # faz parte do plano nao pode ser lida como recomendada, e uma regra
        # recomendada e violada precisa mostrar as duas coisas ao mesmo tempo.
        m.append(('Estado %s' % rg,
                  'VAR rec = [Recomendada %s] VAR vio = [Violada %s] '
                  'RETURN SWITCH ( TRUE (), '
                  'rec = 1 && vio = 1, "RECOMENDADA - VIOLADA", '
                  'rec = 1, "RECOMENDADA", '
                  'vio = 1, "FORA DO PLANO - VIOLADA", '
                  '"fora do plano" )' % (rg, rg),
                  None, '13 Estado das regras'))
        # Violacao dentro do plano e vermelha: e o alerta que o gestor tem de
        # ver primeiro. Violacao FORA do plano e ambar -- aconteceu, merece
        # atencao, mas nao entra no plano so por ter acontecido. Regra fora do
        # plano e sem violacao fica cinza, nunca oculta.
        m.append(('Cor %s' % rg,
                  'VAR e = [Estado %s] RETURN SWITCH ( e, '
                  '"RECOMENDADA - VIOLADA", "%s", '
                  '"FORA DO PLANO - VIOLADA", "%s", '
                  '"RECOMENDADA", "%s", "%s" )'
                  % (rg, VERMELHO, AMBAR, VERDE, CINZA),
                  None, '14 Cores das regras'))
        m.append(('Rotulo %s' % rg,
                  '"%s"' % rg.replace('_', '-'), None, '15 Rotulos'))
    return m


def bloco_medida(nome, dax, formato, pasta):
    seguro = "'%s'" % nome
    ln = ['\tmeasure %s = %s' % (seguro, dax)]
    if formato:
        ln.append('\t\tformatString: %s' % formato)
    ln.append('\t\tlineageTag: %s' % guid('med/' + nome))
    ln.append('\t\tdisplayFolder: %s' % pasta)
    ln.append('')
    ln.append('\t\tannotation PBI_FormatHint = {"isGeneralNumber":true}')
    ln.append('')
    return '\n'.join(ln)


def main():
    if not os.path.isfile(TMDL):
        raise SystemExit('nao encontrei %s' % TMDL)
    s = io.open(TMDL, encoding='utf-8').read()

    todas = ADR035 + NOVAS
    faltam_col = [t for t in todas if ('\tcolumn %s\n' % t[0]) not in s]
    print('colunas a acrescentar: %d de %d' % (len(faltam_col), len(todas)))

    # ---- 1. tipagem no Power Query -------------------------------------
    # Entra ANTES do fecho da lista de Table.TransformColumnTypes.
    alvo_m = [t for t in faltam_col if ('{"%s",' % t[0]) not in s]
    if alvo_m:
        mm = re.search(r'(\{"Sigma", type number\})', s)
        if not mm:
            raise SystemExit('nao achei a ancora de tipagem {"Sigma", type number}')
        adic = ''.join(', {"%s", %s}' % (t[0], t[2]) for t in alvo_m)
        s = s[:mm.end()] + adic + s[mm.end():]
        print('tipagem M: %d colunas acrescentadas' % len(alvo_m))

    # ---- 2. colunas no TMDL --------------------------------------------
    if faltam_col:
        anc = s.index('\tcolumn Faixa_Sigma =')
        novo = ''.join(bloco_coluna(n, t, oc) for n, t, _, oc in faltam_col)
        s = s[:anc] + novo + s[anc:]
        print('TMDL: %d colunas acrescentadas' % len(faltam_col))

    # ---- 3. medidas -----------------------------------------------------
    # REESCREVE as que ja existem em vez de so acrescentar as que faltam.
    # Sem isso o script vira mao unica: corrigir a definicao de uma medida
    # exigiria editar o .tmdl a mao, que e exatamente o que ele existe para
    # evitar. Remove o bloco antigo pelo nome e reinsere o atual.
    meds = medidas()
    nomes = set(m[0] for m in meds)
    removidas = 0
    for nome in nomes:
        marca = "\tmeasure '%s' =" % nome
        while marca in s:
            ini = s.index(marca)
            # o bloco termina na proxima declaracao de mesmo nivel
            resto = s[ini + len(marca):]
            fim = len(s)
            for prox in ('\n\tmeasure ', '\n\tcolumn ', '\n\tpartition ',
                         '\n\tannotation '):
                p = resto.find(prox)
                if p != -1:
                    fim = min(fim, ini + len(marca) + p + 1)
            s = s[:ini] + s[fim:]
            removidas += 1
    if removidas:
        print('TMDL: %d medidas antigas removidas para reescrita' % removidas)

    anc = s.index('\tpartition Fato_QC = m')
    novo = ''.join(bloco_medida(*m) for m in meds)
    s = s[:anc] + novo + s[anc:]
    print('TMDL: %d medidas gravadas' % len(meds))

    io.open(TMDL, 'w', encoding='utf-8', newline='').write(s)
    print('\nFato_QC.tmdl atualizado')
    print('  medidas do plano  : %d' % len([m for m in meds if 'Plano' in m[3]]))
    print('  medidas das regras: %d' % len([m for m in meds if 'egras' in m[3]
                                            or 'Estado' in m[3] or 'Cores' in m[3]
                                            or 'Rotulos' in m[3]]))
    return 0


if __name__ == '__main__':
    sys.exit(main())
