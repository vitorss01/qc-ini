# -*- coding: utf-8 -*-
"""painel.py -- ADR-054: o Painel dos dois produtos no MESMO desenho.

O que o gestor pediu (20/09/2026) e o que cada peca resolve:

  topo igual        A Hematologia tinha o topo limpo; a Bioquimica tinha o
                    spinner por cima do rotulo "Ano" e as caixas de trimestre
                    por cima do seletor de Controle Externo. Agora os dois
                    produtos tem a MESMA linha 3/4: analito, lote, De/Ate,
                    Ano/Periodo -- e a mesma largura de coluna (A..U).
  filtro unico      Havia dois caminhos para dizer "que periodo eu quero":
                    4 caixas de trimestre (qsel1..4) e a lista Ano+Periodo.
                    Dois caminhos = duas verdades: marcar 2o Tri e escolher T1
                    esvaziava o Painel sem dizer por que. Ficou UM: Ano +
                    Periodo (Tudo / Ano inteiro / T1..T4 / S1..S2 / Jan..Dez),
                    que e o que a Hematologia nao tinha.
  desempenho em A   O bloco de Sigma ficava na coluna R (Bio) / J (Hema) --
                    fora da tela sem rolar para o lado. Vai para a coluna A,
                    abaixo dos graficos, em grupos de colunas mescladas (as
                    colunas de cima sao estreitas; o bloco precisa de largura).
  ultima violacao   A Bioquimica nao mostrava. O motor ja calculava para os
                    dois (engPainel colunas 19..21); so a interface nao lia.
  grafico elastico  mUI.AjustarGraficos: a largura do grafico acompanha a
                    largura VISIVEL da janela, em qualquer zoom.

Tudo aqui e apresentacao. Nenhuma conta muda: os numeros continuam vindo de
Eng_Saida (engPainel) e da aba Estatistica.
"""

import ux

# ---------------------------------------------------------------- larguras
# As da Hematologia (o topo que o gestor aprovou). A Bioquimica passa a usar
# as mesmas: mesmo desenho, mesma posicao de cada coluna nos dois produtos.
LARGURAS = [8.2, 5.2, 12.2, 7.2, 8.2, 6.2, 13.2, 8.2, 8.2, 8.9, 14.9, 15.9,
            14.9, 12.9, 21.9, 23.9, 17.9, 6.2, 18.0, 18.0, 22.0]      # A..U

AZUL = 0x2C2813        # #13282C  texto de titulo
VERDE = 0x6B6F1F       # #1F6F6B  fundo de cabecalho
CLARO = 0xF4F6ED       # #EDF6F4  fundo de faixa
BRANCO = 0xFFFFFF
CINZA = 0x595959
BORDA = 0xD8C8C8       # #C8C8D8

# grupos de colunas mescladas do bloco de desempenho (coluna inicial = A)
G_SIGMA = [('A', 'B'), ('C', 'D'), ('E', 'G'), ('H', 'I'), ('J', 'J'), ('K', 'L'), ('M', 'N'), ('O', 'O')]
G_PLANO = [('A', 'B'), ('C', 'D'), ('E', 'H'), ('I', 'J'), ('K', 'K'), ('L', 'O'), ('P', 'P')]
G_ERRO = [('A', 'B'), ('C', 'D'), ('E', 'G'), ('H', 'I'), ('J', 'J'), ('K', 'L')]
G_REF = [('A', 'B'), ('C', 'D'), ('E', 'G')]
G_MARGEM = [('A', 'D'), ('E', 'E')]
G_NOTA = ('A', 'O')

# "(todos)" era ambiguo ao lado de um seletor de Ano: todos os anos, ou o ano
# inteiro? A lista agora diz qual e qual (ADR-054).
TUDO = 'Tudo (sem filtro)'
ANO_INTEIRO = 'Ano inteiro'
PERIODOS = (TUDO + ';' + ANO_INTEIRO + ';T1;T2;T3;T4;S1;S2;'
            'Jan;Fev;Mar;Abr;Mai;Jun;Jul;Ago;Set;Out;Nov;Dez')

LINKS = [
    ('Westgard et al., 2018 — Sigma metrics e tabela Sigma × DPM ↗',
     'https://doi.org/10.11613/BM.2018.020502'),
    ('Westgard & Westgard, 2019 — Sigma Rules with Run Sizes ↗',
     'https://doi.org/10.1093/ajcp/aqy158'),
    ('Peng et al., 2021 — aplicação prática de regras, N e run size ↗',
     'https://doi.org/10.1002/jcla.23665'),
    ('CLSI C24-Ed4 — SQC baseado em risco (base conceitual) ↗',
     'https://clsi.org/standards/products/method-evaluation/documents/c24/'),
]

CAB_IND = ['Nível', 'n', 'Média', 'DP', 'CV %', 'ETp %', 'Bias %', 'ET %', 'Sigma', 'Status']
FMT_IND = ['Geral', 'Geral', '0,0000', '0,0000', '0,00"%"', '0,00"%"', '0,00"%"', '0,00"%"', '0,00', 'Geral']

# ---------------------------------------------------------------- linha 1
# O titulo estava em Calibri 22 na Bioquimica e Calibri 18 na Hematologia: o
# mesmo titulo, dois tamanhos. Como a barra de navegacao comeca DEPOIS do
# titulo, isso tambem fazia a barra cair em coluna diferente nos dois -- e na
# Bioquimica ela entrava 61 pt POR CIMA do titulo (ADR-055).
TITULO_FONTE = 'Segoe UI'
TITULO_TAM = 16
SUB_TAM = 9


def titulo_padrao(produto, nlv):
    return f'PAINEL · LEVEY-JENNINGS ({produto}) — {nlv} NÍVEIS'


# Os dois titulos possiveis. A barra usa o MAIOR deles: assim a coluna onde ela
# comeca e a mesma nos dois produtos, que foi o que o gestor pediu.
TITULOS_PADRAO = (titulo_padrao('Bioquímica', 2), titulo_padrao('Hematologia', 3))

# ---------------------------------------------------------------- legenda
# A legenda tinha 14 entradas numa caixa de 80 pt: os rotulos saiam cortados
# ("Viol...", "Rep..."). Sete sao as linhas de limite, que se leem pela POSICAO
# no grafico, e "Repeticao" aparecia 3x (uma por replica).
LIMITES = ('-3s', '-2s', '-1s', '+1s', '+2s', '+3s')


# ---------------------------------------------------------------- utilidades
def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def num(n):
    v = 0
    for ch in n:
        v = v * 26 + (ord(ch) - 64)
    return v


def medir_texto(ws, texto, fonte=TITULO_FONTE, tam=TITULO_TAM, negrito=True):
    """Largura REAL do texto, em pontos (uma medida so, em ux.larg_texto)."""
    return ux.larg_texto(ws, texto, fonte, tam, negrito)


def coluna_da_barra(ws):
    """Primeira coluna livre para a barra de navegacao da linha 1.

    O gestor: "verificar ate onde as informacoes escritas da linha 1 estao indo
    e colocar o menu numa posicao que nao sobreponha elas -- portanto pular uma
    coluna e inserir apos esse espaco". Entao: acha a coluna onde o titulo
    TERMINA, deixa a seguinte vazia, e devolve a proxima. Mede o maior dos dois
    titulos para a coluna ser a MESMA nos dois produtos.
    """
    larg = max(medir_texto(ws, t) for t in TITULOS_PADRAO)
    c = 1
    while c < len(LARGURAS) and ws.Cells(1, c).Left + ws.Cells(1, c).Width < larg:
        c += 1
    c += 2                     # c+1 fica VAZIA (o respiro pedido); a barra vem em c+2
    fim = ws.Range(f'A1:{col(len(LARGURAS))}1').Width
    if ws.Cells(1, c).Left + ux.LARG_BARRA > fim:
        raise SystemExit(f'barra de navegacao nao cabe a partir de {col(c)}: '
                         f'{ws.Cells(1, c).Left:.0f} + {ux.LARG_BARRA} > {fim:.0f}')
    return col(c)


def legenda(ws):
    """Deixa na legenda so o que nao da para ler pela posicao no grafico.

    Saem as seis linhas de limite (-3s..+3s) e as repeticoes de nome duplicado;
    ficam Media, Resultado, OK, Violacao, Repeticao e Calibracao.

    A legenda FICA a direita. Tentei embaixo, para devolver largura ao grafico,
    e ela caiu em cima do rotulo "Corrida (RUN)" e do numero do eixo -- a area
    de plotagem destes graficos tem layout manual, entao o Excel nao reflui
    sozinho. E a conta nao fechava: a largura e elastica (acompanha a janela),
    a ALTURA nao -- na Hematologia sao 150 pt para tres graficos, e a faixa de
    baixo custaria 22% deles. A direita, com seis entradas, a caixa se
    redimensiona sozinha para 72 pt e o rotulo mais largo ocupa 42.

    Reposicionar tambem TIRA a largura fixa de 80 pt que estava gravada no
    arquivo: era ela que apertava as quatorze entradas.
    """
    n_limpas = 0
    for co in ws.ChartObjects():
        ch = co.Chart
        try:
            ch.HasLegend = True
            lg = ch.Legend
            lg.Position = -4152                 # xlLegendPositionRight
            lg.Font.Name = 'Segoe UI'
            lg.Font.Size = 9
            n = ch.SeriesCollection().Count
            if ch.Legend.LegendEntries().Count != n:
                continue                        # ja limpa: mexer de novo apagaria o resto
            vistos, apagar = set(), []
            for i in range(1, n + 1):
                nome = str(ch.SeriesCollection(i).Name)
                if nome in LIMITES or nome in vistos:
                    apagar.append(i)
                else:
                    vistos.add(nome)
            for i in reversed(apagar):          # de tras para frente: o indice dos
                ch.Legend.LegendEntries(i).Delete()   # que ficam nao se desloca
            n_limpas += 1
        except Exception:
            continue
    return n_limpas


# ------------------------------------------------- realce da regra recomendada
# Uma coluna por regra, 1/0, DERIVADA do texto da coluna Regras da propria
# tabela de plano -- nao ha segunda fonte: mudar a faixa muda o realce junto.
MAT_C0 = 10                    # coluna J: a tabela vai ate H nos dois produtos


def realce_regras(wb, nlv, regras):
    """Cfg_PlanoQC ganha a matriz de regras e os nomes que a formatacao le.

    regras: [(rotulo como esta no Painel, token como esta na coluna Regras)],
            na mesma ordem das colunas M..Q do Painel.

    Existia so na Bioquimica -- e la estava MORTO: o Sigma do plano era lido de
    Painel!V10:V11, celulas que o Painel novo esvaziou (ADR-054). Resultado:
    "SEM DADOS", nenhuma regra iluminada e as duas notas do bloco em branco.
    Agora le a coluna Sigma do Painel (I7:I{6+nlv}), nos dois produtos.
    """
    cfg = wb.Sheets('Cfg_PlanoQC')
    ult = 6 + nlv
    # a descricao que morava em A1:A2 da Hematologia desce: as duas primeiras
    # linhas passam a ser os escalares, iguais nos dois arquivos
    for r in (1, 2):
        v = cfg.Range(f'A{r}').Value
        if isinstance(v, str) and v.strip().startswith(('tblPlanoQC', 'Fonte única')):
            cfg.Range(f'A{11 + r}').Value = v
    cfg.Range('A1').Value = 'Sigma efetivo do plano'
    cfg.Range('B1').Formula = (f'=IF(COUNT(Painel!$I$7:$I${ult})=0,"SEM DADOS",'
                               f'MIN(Painel!$I$7:$I${ult}))')
    cfg.Range('C1').Value = 'Regras que o motor executa nesta faixa'
    cfg.Range('D1').Formula = '=IF($B$2<1,"",INDEX($D$4:$D$8,$B$2))'
    cfg.Range('A2').Value = 'Linha da faixa vigente (índice na matriz):'
    cfg.Range('B2').Formula = ('=IF(NOT(ISNUMBER($B$1)),0,SUMPRODUCT(($A$4:$A$8<=$B$1)*'
                               '($B$4:$B$8>$B$1)*ROW($A$4:$A$8))-3)')
    cfg.Range('C2').Value = 'Nível que governa o plano: o de MENOR Sigma'
    for c in ('A1', 'A2', 'C1', 'C2'):
        cfg.Range(c).Font.Italic = True
    for j, (rotulo, token) in enumerate(regras):
        c = MAT_C0 + j
        cel = cfg.Cells(3, c)
        cel.Value = rotulo
        cel.Font.Bold = True
        cfg.Columns(c).ColumnWidth = 9
        for r in range(4, 9):
            cfg.Cells(r, c).Formula = (f'=IF(ISNUMBER(SEARCH("{token}",'
                                       f'IF($D{r}="",$D$1,$D{r}))),1,0)')
        cfg.Cells(10, c).Formula = (f'=IF(OR(NOT(ISNUMBER($B$1)),$B$2<1),0,'
                                    f'INDEX({col(c)}$4:{col(c)}$8,$B$2))')
        cfg.Cells(10, c).Font.Bold = True
    cfg.Cells(2, MAT_C0).Value = 'matriz derivada da coluna Regras — não editar à mão'
    cfg.Cells(9, MAT_C0).Value = 'regras ativas para o Sigma selecionado:'
    cfg.Cells(2, MAT_C0).Font.Italic = True
    cfg.Cells(9, MAT_C0).Font.Italic = True
    c1, c2 = col(MAT_C0), col(MAT_C0 + len(regras) - 1)
    for nome, ref in (('regrasAtivas', f'=Cfg_PlanoQC!${c1}$10:${c2}$10'),
                      ('regrasRotulos', f'=Cfg_PlanoQC!${c1}$3:${c2}$3'),
                      ('sigmaDoPlano', '=Cfg_PlanoQC!$B$1')):
        try:
            wb.Names(nome).Delete()
        except Exception:
            pass
        wb.Names.Add(nome, ref)
    for lo in list(cfg.ListObjects):
        if lo.Name == 'tblPlanoQC_Sigma':
            try:
                lo.Resize(cfg.Range(cfg.Cells(3, 1), cfg.Cells(8, MAT_C0 + len(regras) - 1)))
            except Exception:
                pass
    return True


def _est(c, ult, nivel):
    """Celula da aba Estatistica para (analito em tela, nivel)."""
    return (f'=IFERROR(INDEX(Estatística!${c}$14:${c}${ult},'
            f'MATCH(selAnalito&"|"&{nivel},Estatística!$AB$14:$AB${ult},0)),"")')


def _borda(rg, cor=BORDA):
    for i in (7, 8, 9, 10):        # esquerda, topo, base, direita
        b = rg.Borders(i)
        b.LineStyle = 1
        b.Weight = 2
        b.Color = cor


def titulo(ws, r, texto):
    rg = ws.Range(f'A{r}:U{r}')
    rg.UnMerge()
    rg.Merge()
    c = ws.Range(f'A{r}')
    c.Value = texto
    c.Font.Bold = True
    c.Font.Size = 11
    c.Font.Name = 'Segoe UI'
    c.Font.Color = AZUL
    rg.Interior.Color = CLARO
    c.HorizontalAlignment = -4131
    c.VerticalAlignment = -4108
    ws.Rows(r).RowHeight = 20
    b = rg.Borders(9)
    b.LineStyle = 1
    b.Weight = 2
    b.Color = VERDE
    return r + 1


def nota(ws, r, texto, altura=None):
    a, b = G_NOTA
    rg = ws.Range(f'{a}{r}:{b}{r}')
    rg.UnMerge()
    rg.Merge()
    c = ws.Range(f'{a}{r}')
    if isinstance(texto, str) and texto.startswith('='):
        c.Formula = texto
    else:
        c.Value = texto
    c.Font.Size = 8
    c.Font.Italic = True
    c.Font.Name = 'Segoe UI'
    c.Font.Color = CINZA
    c.WrapText = True
    c.HorizontalAlignment = -4131
    c.VerticalAlignment = -4108
    ws.Rows(r).RowHeight = altura or 15
    return r + 1


def tabela(ws, r, grupos, cabecalhos, linhas, formatos=None):
    """Cabecalho + linhas, cada coluna logica num grupo de colunas mescladas."""
    for (c0, c1) in grupos:
        rg = ws.Range(f'{c0}{r}:{c1}{r}')
        if c0 != c1:
            rg.UnMerge()
            rg.Merge()
    for (c0, c1), t in zip(grupos, cabecalhos):
        c = ws.Range(f'{c0}{r}')
        c.Value = t
        c.Font.Bold = True
        c.Font.Size = 9
        c.Font.Name = 'Segoe UI'
        c.Font.Color = BRANCO
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        c.WrapText = True
    fim = num(grupos[-1][1])
    cab = ws.Range(f'A{r}:{col(fim)}{r}')
    cab.Interior.Color = VERDE
    _borda(cab, VERDE)
    ws.Rows(r).RowHeight = 26
    r += 1
    for lin in linhas:
        for (c0, c1) in grupos:
            rg = ws.Range(f'{c0}{r}:{c1}{r}')
            if c0 != c1:
                rg.UnMerge()
                rg.Merge()
        for i, ((c0, c1), v) in enumerate(zip(grupos, lin)):
            c = ws.Range(f'{c0}{r}')
            if formatos:
                c.NumberFormatLocal = formatos[i]
            if v is None:
                continue
            if isinstance(v, str) and v.startswith('='):
                c.Formula = v
            else:
                c.Value = v
            c.Font.Size = 10
            c.Font.Name = 'Segoe UI'
            c.HorizontalAlignment = -4108
        faixa = ws.Range(f'A{r}:{col(fim)}{r}')
        _borda(faixa)
        ws.Rows(r).RowHeight = 16
        r += 1
    return r


# ------------------------------------------------------- bloco de desempenho
def bloco_desempenho(ws, r0, nlv, ult, extras=(), fonte_etp=None):
    """DESEMPENHO SIX SIGMA + PLANO DE CQ + ERRO TOTAL + referencias, na coluna A.

    extras: notas proprias do produto (formulas de texto ja prontas).
    """
    r = titulo(ws, r0, 'DESEMPENHO SIX SIGMA — analito selecionado')
    r = nota(ws, r, 'Valores vindos da aba Estatística. O período e o filtro de EQA (provedor / ano / rodada) '
                    'são os definidos lá; o filtro de datas deste Painel manda no gráfico e nos descritivos, '
                    'não aqui.', 22)
    if fonte_etp:
        r = nota(ws, r, fonte_etp)
    r = tabela(ws, r, G_SIGMA,
               ['Nível', 'CV % obs', 'Bias EQC (abs) %', 'ETp %', 'Sigma', 'Classificação',
                'DPM teórico', 'Rendimento teórico %'],
               [[f'N{n}'] + [_est(c, ult, n) for c in ('F', 'G', 'K', 'L', 'M', 'V', 'W')]
                for n in range(1, nlv + 1)],
               ['Geral', '0,00', '0,00', '0,00', '0,00', 'Geral', '#.##0', '0,0000'])
    r = nota(ws, r, 'DPM teórico estimado pelo Sigma: usa a convenção de short-term Sigma com deslocamento de '
                    '1,5 SD. É um benchmark teórico de desempenho, e não uma contagem observada de erros em '
                    'resultados de pacientes.', 22)
    r += 1

    r = titulo(ws, r, 'PLANO DE CQ RECOMENDADO PELO SIGMA')
    linhas = []
    for n in range(1, nlv + 1):
        lr = r + n          # linha fisica desta linha da tabela (cabecalho em r)
        linhas.append([f'N{n}', _est('L', ult, n), _est('X', ult, n), _est('Y', ult, n),
                       _est('Z', ult, n),
                       f'=IF(NOT(ISNUMBER($C{lr})),"",mPlanoQC.PlanoQC($C{lr},"FREQUENCIA"))',
                       _est('AA', ult, n)])
    r = tabela(ws, r, G_PLANO,
               ['Nível', 'Sigma', 'Regras Westgard', 'N (medições)', 'Run Size máx',
                'Frequência de CQ', 'Cobertura do motor'],
               linhas, ['Geral', '0,00', 'Geral', 'Geral', 'Geral', 'Geral', 'Geral'])
    r = nota(ws, r, 'O run size é recomendação de planejamento de SQC baseada em desempenho Sigma e risco. '
                    'Não substitui requisitos regulatórios, de acreditação, instruções do fabricante ou '
                    'procedimentos internos mais restritivos.', 22)
    r = nota(ws, r, 'N é o número TOTAL de medições de controle no evento — não o número de níveis. Como '
                    'distribuir entre níveis, materiais e replicatas depende da configuração do laboratório.', 22)
    r = nota(ws, r, 'Run Size é quantos pacientes podem passar entre eventos de CQ. Não confundir com a regra R_4s.')
    for t in extras:
        r = nota(ws, r, t, 22)
    r += 1

    r = titulo(ws, r, 'ERRO TOTAL vs ETp — orçamento de erro')
    r = tabela(ws, r, G_ERRO,
               ['Nível', 'ET %', 'ETp %', 'Margem (p.p.)', 'Margem %', 'Situação'],
               [[f'N{n}'] + [_est(c, ult, n) for c in ('H', 'K', 'N', 'O', 'P')]
                for n in range(1, nlv + 1)],
               ['Geral', '0,00', '0,00', '0,00', '0,00', 'Geral'])
    r += 1

    r = titulo(ws, r, 'MARGEM CRÍTICA — todos os analitos')
    for rot, chave in (('ETp excedido', 'ETp excedido'),
                       ('Margem crítica (≤10%)', 'Margem critica'),
                       ('Dentro do orçamento', 'Dentro do orcamento')):
        for (c0, c1) in G_MARGEM:
            rg = ws.Range(f'{c0}{r}:{c1}{r}')
            if c0 != c1:
                rg.UnMerge()
                rg.Merge()
        ws.Range(f'A{r}').Value = rot
        ws.Range(f'A{r}').Font.Size = 10
        ws.Range(f'A{r}').Font.Name = 'Segoe UI'
        ws.Range(f'A{r}').HorizontalAlignment = -4131
        c = ws.Range('E' + str(r))
        c.Formula = f'=COUNTIF(Estatística!$P$14:$P${ult},"{chave}")'
        c.Font.Bold = True
        c.Font.Size = 10
        c.HorizontalAlignment = -4108
        _borda(ws.Range(f'A{r}:E{r}'))
        ws.Rows(r).RowHeight = 16
        r += 1
    r += 1

    r = titulo(ws, r, 'SIGMA × DPM × RENDIMENTO — referência')
    linhas = []
    for i, s in enumerate([6, 5.5, 5, 4.5, 4, 3.5, 3, 2.5, 2]):
        lr = r + 1 + i
        linhas.append([s, f'=mPlanoQC.DPMdoSigma($A{lr})', f'=mPlanoQC.RendimentoDoSigma($A{lr})'])
    r = tabela(ws, r, G_REF, ['Sigma', 'DPM teórico', 'Rendimento %'], linhas,
               ['0,0', '#.##0,0', '0,00000'])
    r += 1

    r = titulo(ws, r, 'BASE CIENTÍFICA DO PLANO DE CQ')
    r = nota(ws, r, 'Metodologia: Westgard Sigma Rules with Run Sizes. Base conceitual de SQC baseado em '
                    'risco: CLSI C24-Ed4.')
    for texto, url in LINKS:
        a, b = G_NOTA
        rg = ws.Range(f'{a}{r}:{b}{r}')
        rg.UnMerge()
        rg.Merge()
        c = ws.Range(f'{a}{r}')
        c.Hyperlinks.Delete()
        ws.Hyperlinks.Add(c, url, '', texto, texto)
        c.Font.Size = 9
        c.Font.Name = 'Segoe UI'
        c.HorizontalAlignment = -4131
        ws.Rows(r).RowHeight = 15
        r += 1
    return r


# ------------------------------------------------------------------- cabecalho
def _rotulo(ws, cel, texto, tam=9):
    c = ws.Range(cel)
    c.Value = texto
    c.Font.Bold = True
    c.Font.Size = tam
    c.Font.Name = 'Segoe UI'
    c.Font.Color = AZUL
    c.HorizontalAlignment = -4152        # direita: cola no campo que rotula
    c.VerticalAlignment = -4108


def _campo(ws, faixa, lista, titulo_ajuda, msg):
    rg = ws.Range(faixa)
    if ':' in faixa:
        rg.UnMerge()
        rg.Merge()
    c = ws.Range(faixa.split(':')[0])
    c.Font.Bold = True
    c.Font.Size = 10
    c.Font.Name = 'Segoe UI'
    c.HorizontalAlignment = -4108
    c.VerticalAlignment = -4108
    rg.Interior.Color = 0xF2F2F2
    _borda(rg, 0xA6A6A6)
    rg.Locked = False
    try:
        c.Validation.Delete()
    except Exception:
        pass
    if lista:
        c.Validation.Add(3, 1, 1, lista)
        c.Validation.ShowError = False
        c.Validation.InputTitle = titulo_ajuda
        c.Validation.InputMessage = msg
    return c


def cabecalho(ws, produto, nlv, ult, param, rotulos_regras, falta_media, alvo):
    """Linhas 1..(6+nlv): titulo, filtros, indicadores e Westgard -- iguais nos dois."""
    # --------- larguras e alturas ---------
    for i, w in enumerate(LARGURAS):
        ws.Columns(i + 1).ColumnWidth = w
    ws.Range(ws.Columns(len(LARGURAS) + 1), ws.Columns(len(LARGURAS) + 12)).ColumnWidth = 8.1
    ws.Rows(1).RowHeight = 40.2
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 17.4
    ws.Rows(4).RowHeight = 14.4
    ws.Rows(5).RowHeight = 14.4
    ws.Rows(6).RowHeight = 28.2

    # --------- titulo e subtitulo ---------
    for faixa in (f'A1:{col(len(LARGURAS))}1', f'A2:{col(len(LARGURAS))}2'):
        rg = ws.Range(faixa)
        rg.UnMerge()
        rg.Merge()
    ws.Range('A1').Value = titulo_padrao(produto, nlv)
    ws.Range('A2').Value = ('Analito: spinner ▲▼ ou a lista (C3)  ·  Lote em análise: H3  ·  '
                            'Período: Ano + Período (K3/K4) ou as datas De/Até (G3/G4)')
    # o MESMO titulo estava em dois tamanhos (Calibri 22 e 18): mesma fonte,
    # mesmo tamanho nos dois -- e a barra de navegacao cai na mesma coluna
    for cel, tam, negrito in (('A1', TITULO_TAM, True), ('A2', SUB_TAM, False)):
        f = ws.Range(cel).Font
        f.Name = TITULO_FONTE
        f.Size = tam
        f.Bold = negrito

    # --------- linha 3/4: filtros ---------
    _rotulo(ws, 'A3', 'Analito nº')
    _rotulo(ws, 'F3', 'De')
    _rotulo(ws, 'F4', 'Até')
    _rotulo(ws, 'J3', 'Ano')
    _rotulo(ws, 'J4', 'Período')
    for cel in ('G3', 'G4'):
        c = ws.Range(cel)
        c.NumberFormatLocal = 'dd/mm/aaaa'
        c.Font.Size = 10
        c.HorizontalAlignment = -4108
        c.Locked = False
        _borda(ws.Range(cel), 0xA6A6A6)
    _campo(ws, 'H3:I3', '=lstLotes', 'Lote em análise',
           'Escolha o lote. Média, DP, limites, gráfico, Westgard e Estatística passam a usar SÓ os '
           'parâmetros deste lote.')
    ws.Range('H3:I3').NumberFormatLocal = '"Lote "0;;;"Lote "@'
    ws.Range('H3:I3').Font.Color = 0xFF0000
    ws.Range('H3:I3').Interior.Color = 0xF2F2F2
    c = ws.Range('H4')
    c.Value = '▲ lote em análise'
    c.Font.Size = 8
    c.Font.Italic = True
    c.Font.Color = 0x808080
    c.HorizontalAlignment = -4108
    _campo(ws, 'K3', '=lstAnosCIQ', 'Ano',
           'Ano do período. Ano + Período preenchem as datas De/Até (G3/G4).')
    _campo(ws, 'K4', PERIODOS, 'Período',
           'Tudo = todos os anos · Ano inteiro = o ano de K3 · T1..T4 = trimestre · S1/S2 = semestre · Jan..Dez = mês.\n'
           'Escolher aqui reescreve De/Até. Para um intervalo qualquer, digite direto em G3/G4.')

    # --------- linha 5: titulos das duas tabelas ---------
    for faixa, texto in (('A5:J5', 'INDICADORES POR NÍVEL'),
                         ('L5:U5', 'WESTGARD — violações no período')):
        rg = ws.Range(faixa)
        rg.UnMerge()
        rg.Merge()
        c = ws.Range(faixa.split(':')[0])
        c.Value = texto
        c.Font.Bold = True
        c.Font.Size = 10
        c.Font.Name = 'Segoe UI'
        c.Font.Color = AZUL
        c.HorizontalAlignment = -4131
        rg.Interior.Color = CLARO

    # --------- linha 6: cabecalhos ---------
    cab_w = ['Nível'] + [r for r, _ in rotulos_regras] + ['Total', 'Últ. violação', 'Classificação', 'Histórico']
    for letras, rotulos in (('ABCDEFGHIJ', CAB_IND), ('LMNOPQRSTU', cab_w)):
        for c, t in zip(letras, rotulos):
            cel = ws.Range(f'{c}6')
            cel.Value = t
            cel.Font.Bold = True
            cel.Font.Size = 9
            cel.Font.Name = 'Segoe UI'
            cel.Font.Color = BRANCO
            cel.HorizontalAlignment = -4108
            cel.VerticalAlignment = -4108
            cel.WrapText = True
        rg = ws.Range(f'{letras[0]}6:{letras[-1]}6')
        rg.Interior.Color = VERDE
        _borda(rg, VERDE)

    # --------- linhas 7..: os dados, lidos do motor ---------
    for n in range(1, nlv + 1):
        r = 6 + n
        ws.Rows(r).RowHeight = 17
        f = {
            'A': f'N{n}',
            'B': f'=IF(INDEX(engPainel,{n},2)="","",INDEX(engPainel,{n},2))',
            'C': f'=IF(INDEX(engPainel,{n},3)="","",INDEX(engPainel,{n},3))',
            'D': f'=IF(INDEX(engPainel,{n},4)="","",INDEX(engPainel,{n},4))',
            'E': f'=IF(INDEX(engPainel,{n},5)="","",INDEX(engPainel,{n},5))',
            'F': f'=IF(INDEX(engPainel,{n},6)="","",INDEX(engPainel,{n},6))',
            'G': _est('G', ult, n),
            'H': f'=IF($E{r}="","",$E{r}*1.65+ABS(IF(ISNUMBER($G{r}),$G{r},0)))',
            'I': f'=IF(OR($E{r}="",$E{r}=0,$F{r}=""),"",($F{r}-ABS(IF(ISNUMBER($G{r}),$G{r},0)))/$E{r})',
            'J': (f'=IF(Calc!${param[n - 1]}$1="","SEM MÉDIA/DP",'
                  f'IF(INDEX(engPainel,{n},10)="","",INDEX(engPainel,{n},10)))'),
            'L': f'N{n}',
        }
        for i, (_, slot) in enumerate(rotulos_regras):
            f['MNOPQ'[i]] = f'=IF(INDEX(engPainel,{n},{slot})="","",INDEX(engPainel,{n},{slot}))'
        for i, slot in enumerate((18, 19, 20, 21)):
            f['RSTU'[i]] = f'=IF(INDEX(engPainel,{n},{slot})="","",INDEX(engPainel,{n},{slot}))'
        for c, v in f.items():
            cel = ws.Range(f'{c}{r}')
            if isinstance(v, str) and v.startswith('='):
                cel.Formula = v
            else:
                cel.Value = v
            cel.Font.Size = 10
            cel.Font.Name = 'Segoe UI'
            cel.HorizontalAlignment = -4108
            cel.VerticalAlignment = -4108
        for c, fmt in zip('ABCDEFGHIJ', FMT_IND):
            ws.Range(f'{c}{r}').NumberFormatLocal = fmt
        for c in 'STU':                       # textos: alinhados a esquerda
            ws.Range(f'{c}{r}').HorizontalAlignment = -4131
            ws.Range(f'{c}{r}').Font.Size = 9
        _borda(ws.Range(f'A{r}:J{r}'))
        _borda(ws.Range(f'L{r}:U{r}'))

    # --------- faixa de status do lote ---------
    if ':' in alvo:
        rg = ws.Range(alvo)
        rg.UnMerge()
        rg.Merge()
    c = ws.Range(alvo.split(':')[0])
    c.Font.Size = 9
    c.Font.Name = 'Segoe UI'
    c.WrapText = True
    c.HorizontalAlignment = -4131
    c.VerticalAlignment = -4108
    c.Formula = falta_media


# ---------------------------------------------------------------- driver
def _limpar(ws):
    """Tudo da linha 3 para baixo e refeito. Os graficos sao objetos: nao sao tocados."""
    rg = ws.Range('A3:AG130')
    rg.UnMerge()
    rg.Clear()
    try:
        rg.FormatConditions.Delete()
    except Exception:
        pass


def _fc(rg, tipo, op, f1, cor=None, fundo=None, negrito=None, parar=False):
    fc = rg.FormatConditions.Add(tipo, op, f1) if op is not None else rg.FormatConditions.Add(tipo, None, f1)
    if cor is not None:
        fc.Font.Color = cor
    if fundo is not None:
        fc.Interior.Color = fundo
    if negrito is not None:
        fc.Font.Bold = negrito
    if parar:
        fc.StopIfTrue = True
    return fc


def padronizar(wb, produto, nlv, ult, param, rotulos_regras, faixa, faixa_formula,
               linha_bloco, extras=(), eqa=False, fonte_etp=None):
    """Refaz o Painel no desenho padrao. Devolve a ultima linha escrita.

    param           celulas do Calc que dizem se o lote tem media/DP, por nivel
    rotulos_regras  [(rotulo, coluna do motor)] das 5 regras de Westgard
    faixa           onde vai a faixa de status do lote
    faixa_formula   o texto dessa faixa (difere no numero de niveis)
    linha_bloco     primeira linha do bloco de desempenho (abaixo dos graficos)
    extras          notas proprias do produto, no fim do bloco do plano
    eqa             True na Bioquimica: seletor de Controle Externo em M3:N4
    fonte_etp       nota com a origem do ETp (so onde o produto tem essa coluna)
    """
    ws = wb.Sheets('Painel')
    cfg = wb.Sheets('Configuração')

    def texto(v):
        if v in (None, ''):
            return ''
        return str(int(v)) if isinstance(v, float) and v == int(v) else str(v).strip()

    # ---- o que precisa sobreviver a reconstrucao ----
    # O lote em analise mudou de celula (E3 -> H3) e o rotulo "Ano" morava em
    # H3: ler a celula pelo endereco pegaria o rotulo. So vale como lote o que
    # esta CADASTRADO; senao, o lote em uso (Configuracao!C20).
    cadastrados = {texto(c.Value).upper() for c in cfg.Range('C26:C125').Cells if c.Value not in (None, '')}
    lote = ''
    for cel in ('H3', 'E3'):
        v = texto(ws.Range(cel).Value)
        if v and v.upper() in cadastrados:
            lote = v
            break
    if not lote:
        lote = texto(cfg.Range('C20').Value)
    ano = None
    for cel in ('K3', 'I3'):
        v = ws.Range(cel).Value
        if isinstance(v, (int, float)) and 1900 < v < 2200:
            ano = int(v)
            break
    if ano is None:
        # sem ano, escolher "T1" na lista nao filtraria nada e o usuario acharia
        # que quebrou. O padrao e o ULTIMO ano com dado (Configuracao!Z, escrita
        # por mDados.AtualizarListasAno).
        anos = [int(c.Value) for c in cfg.Range('Z2:Z50').Cells
                if isinstance(c.Value, (int, float)) and 1900 < c.Value < 2200]
        ano = max(anos) if anos else None
    periodo = None
    for cel in ('K4', 'I4'):
        v = str(ws.Range(cel).Value or '').strip()
        if v == '(todos)':        # rotulo antigo: era o ano inteiro
            v = ANO_INTEIRO
        if v in PERIODOS.split(';'):
            periodo = v
            break
    de, ate = ws.Range('G3').Value, ws.Range('G4').Value
    nan = ws.Range('B3').Value
    if faixa_formula is None:
        # quem sabe o que a faixa diz e o arquivo, nao o instalador: ela so
        # muda de lugar (ADR-054), o texto e o que ja estava la.
        for cel in (faixa.split(':')[0], 'O3'):
            f = ws.Range(cel).Formula
            if isinstance(f, str) and f.startswith('=') and 'loteAnalise' in f:
                faixa_formula = f
                break
        if faixa_formula is None:
            raise SystemExit('Painel: nao achei a formula da faixa de status (O3)')
    prov = None
    if eqa:
        for cel in ('M4', 'L3'):
            f = ws.Range(cel).Formula
            if isinstance(f, str) and 'Analitos!' in f:
                prov = f
                break
        if prov is None:
            prov = ('=IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($C$3,Analitos!$A$4:$A$43,0)),"CAP")')

    # ---- caixas de trimestre: o filtro passa a ser um so (Ano + Periodo) ----
    for s in list(ws.Shapes):
        nm = str(s.Name)
        if nm.startswith('Check Box') or nm == 'btnHist':
            s.Delete()
    for n in ('qsel1', 'qsel2', 'qsel3', 'qsel4'):
        try:
            wb.Names(n).Delete()
        except Exception:
            pass

    _limpar(ws)

    # ---- coluna M alinhada a esquerda (pedido do gestor); a tabela de
    #      Westgard sobrescreve celula a celula, entao nada fica torto ----
    ws.Columns('M').HorizontalAlignment = -4131

    cabecalho(ws, produto, nlv, ult, param, rotulos_regras, faixa_formula, faixa)

    # ---- analito: numero (spinner), nome (lista) e o rodape ----
    c = ws.Range('B3')
    c.Value = nan if nan not in (None, '') else 1
    c.Font.Bold = True
    c.Font.Size = 11
    c.HorizontalAlignment = -4108
    c.Locked = False
    try:
        c.Validation.Delete()
    except Exception:
        pass
    c.Validation.Add(1, 1, 7, '1')
    _campo(ws, 'C3:D3', '=lstAnalitos', 'Analito', 'Escolha o analito na lista (ou use o spinner ▲▼).')
    ws.Range('C3').Formula = '=IFERROR(INDEX(Analitos!$A$4:$A$43,$B$3),"")'
    rg = ws.Range('A4:B4')
    rg.Merge()
    ws.Range('A4').Value = '(spinner ▲▼)'
    ws.Range('A4').Font.Size = 8
    ws.Range('A4').Font.Italic = True
    ws.Range('A4').Font.Color = 0x808080
    rg = ws.Range('C4:E4')
    rg.Merge()
    ws.Range('C4').Formula = ('="Analito "&$B$3&" de "&SUMPRODUCT(--(Analitos!$A$4:$A$43<>""))'
                              '&"  ·  Lançamentos no lote: "&loteAtivo')
    ws.Range('C4').Font.Size = 8
    ws.Range('C4').Font.Color = 0x808080
    ws.Range('C4').ShrinkToFit = True

    # ---- valores que sobreviveram ----
    if lote:
        ws.Range('H3').Value = "'" + lote
    if ano:
        ws.Range('K3').Value = ano
    # sem periodo escolhido, nao inventa janela: mostra tudo
    ws.Range('K4').Value = periodo or TUDO
    if de not in (None, ''):
        ws.Range('G3').Value = de
    if ate not in (None, ''):
        ws.Range('G4').Value = ate

    # ---- Controle Externo (so a Bioquimica tem): M3:N4 ----
    if eqa:
        rg = ws.Range('M3:N3')
        rg.Merge()
        ws.Range('M3').Value = 'Controle externo (EQA)'
        ws.Range('M3').Font.Bold = True
        ws.Range('M3').Font.Size = 9
        ws.Range('M3').Font.Name = 'Segoe UI'
        ws.Range('M3').Font.Color = AZUL
        ws.Range('M3').HorizontalAlignment = -4131
        _campo(ws, 'M4:N4', 'CAP;Controllab', 'Controle externo',
               'Provedor do EQA deste analito. Trocar aqui grava na aba Analitos e refaz bias, ET e Sigma.')
        ws.Range('M4').Formula = prov
        ws.Range('M4').HorizontalAlignment = -4131

    # ---- formatacao condicional das duas tabelas ----
    ult_lin = 6 + nlv
    sg = ws.Range(f'I7:I{ult_lin}')
    _fc(sg, 1, 7, '=4', cor=0x1F6F1F, negrito=True)          # Sigma >= 4
    _fc(sg, 1, 6, '=3', cor=0x0000C0, negrito=True)          # Sigma < 3
    st = ws.Range(f'J7:J{ult_lin}')
    _fc(st, 1, 3, '="OK"', cor=0x1F6F1F, negrito=True)
    _fc(st, 1, 3, '="REJEITADO"', cor=0x0000C0, fundo=0xE1E2FD, negrito=True)
    _fc(st, 1, 3, '="SEM MÉDIA/DP"', cor=0x007A9C, fundo=0xCEF4FF, negrito=True)
    zer = ws.Range(f'M7:R{ult_lin}')
    _fc(zer, 1, 3, '=0', cor=0xBFBFBF)                        # zero nao grita
    _fc(ws.Range(f'M7:Q{ult_lin}'), 1, 5, '=0', cor=0x0000C0, negrito=True)
    # cabecalho das regras: iluminada = recomendada pelo Sigma deste analito,
    # vermelha = violada no periodo. So onde o produto tem a tabela de plano
    # por regra (Bioquimica) -- as formulas falam pt-BR (ADR-039).
    try:
        wb.Names('regrasRotulos')
        wb.Names('regrasAtivas')
        for c in 'MNOPQ':
            cab = ws.Range(f'{c}6')
            # vermelho de AVISO, nao de alarme: o cabecalho inteiro em vermelho
            # solido dominava a tela e fazia um painel normal parecer um incendio
            _fc(cab, 2, None, f'=SOMA(${c}$7:${c}${ult_lin})>0',
                cor=0x0000A0, fundo=0xD5D5FF, negrito=True, parar=True)
            _fc(cab, 2, None, f'=SOMARPRODUTO((regrasRotulos={c}6)*regrasAtivas)=1',
                cor=AZUL, fundo=0xB0D63C, negrito=True)
    except Exception:
        pass

    # ---- bloco de desempenho, na coluna A, abaixo dos graficos ----
    fim = bloco_desempenho(ws, linha_bloco, nlv, ult, extras, fonte_etp)

    # ---- spinner do analito ao lado do numero (E3), como na Hematologia ----
    for s in list(ws.Shapes):
        if str(s.Name).startswith('Spinner'):
            s.Left = ws.Range('E3').Left + 2
            s.Top = ws.Range('E3').Top
            s.Width = 20.4
            s.Height = ws.Range('E3:E4').Height
            s.Placement = 3
            s.OnAction = 'PainelMudou'
    # ---- botao do historico de Westgard, ao lado do titulo da tabela ----
    ux._botao(ws, 'btnHist', 'AbrirEventosWestgard', 'Histórico auditável dos eventos ▸',
              ws.Range('S5').Left, ws.Range('S5').Top - 1, ws.Range('S5:U5').Width - 4, 15,
              VERDE, BRANCO, tam=8)
    # ---- botao de desenvolvedor: no canto, fora do caminho ----
    for s in list(ws.Shapes):
        if str(s.Name) == 'btnDev':
            s.Left = ws.Range('A1').MergeArea.Left + ws.Range('A1').MergeArea.Width - 18
            s.Top = 2

    # ---- graficos: largura inicial = largura da tabela (o VBA reajusta) ----
    larg = ws.Range(f'A1:{col(len(LARGURAS))}1').Width
    for co in ws.ChartObjects():
        co.Placement = 2
        co.Left = 0
        co.Width = larg
    legenda(ws)
    return fim

def sem_trimestre(f):
    """Tira a clausula das caixas de trimestre (qsel1..4) da formula do filtro.

    O filtro por trimestre morava em TRES lugares: as caixas do Painel, o motor
    (CarregarFiltro) e esta formula do Calc. Apagar os nomes sem mexer aqui
    deixava Calc!D em #NOME? -- e #NOME? em D e zero corridas no grafico, sem
    erro visivel na tela. Um filtro, um dono (ADR-054).
    """
    i = f.find(',OR(NOT(OR(qsel')
    if i < 0:
        return f
    j, nivel = i + 1, 0
    while j < len(f):
        if f[j] == '(':
            nivel += 1
        elif f[j] == ')':
            nivel -= 1
            if nivel == 0:
                break
        j += 1
    novo = f[:i] + f[j + 1:]
    if 'qsel' in novo:
        raise SystemExit(f'filtro do Calc: sobrou qsel depois do corte: {novo[:160]}')
    return novo


def conferir_sem_qsel(wb):
    """Nenhuma formula da pasta pode citar os nomes que deixaram de existir."""
    achados = []
    for ws in wb.Worksheets:
        try:
            c = ws.Cells.Find('qsel', None, -4123, 2)      # xlFormulas, xlPart
        except Exception:
            c = None
        if c is not None:
            achados.append(f'{ws.Name}!{c.Address}')
    if achados:
        raise SystemExit('formulas ainda citam qsel (nome apagado): ' + ', '.join(achados))
    return True


def garantir_ano(wb):
    """Preenche o Ano do Painel com o ultimo ano que tem dado.

    Sem ano, escolher "T1" na lista nao filtra nada -- e nada na tela explica
    por que. A lista de anos (Configuracao!Z) so fica pronta no fim da
    instalacao, entao isto roda depois do recalculo.
    """
    ws, cfg = wb.Sheets('Painel'), wb.Sheets('Configuração')
    if ws.Range('K3').Value not in (None, ''):
        return None
    anos = [int(c.Value) for c in cfg.Range('Z2:Z50').Cells
            if isinstance(c.Value, (int, float)) and 1900 < c.Value < 2200]
    if not anos:
        return None
    ws.Range('K3').Value = max(anos)
    return max(anos)
