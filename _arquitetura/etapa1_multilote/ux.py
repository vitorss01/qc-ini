# -*- coding: utf-8 -*-
"""ux.py -- camada de EXPERIENCIA (ADR-051): o Excel com cara e uso de aplicativo.

Aplicada pelos instaladores (aplicar_etapa1.py / portar_hema.py) sobre uma pasta
aberta. Nada aqui calcula -- so apresenta e navega:

  Barra de navegacao  em toda tela de uso, no alto a direita: Inicio, Painel,
                      Lancar, Media/DP, Lotes, Estatistica, Liberacao, Registros,
                      Westgard. A tela atual fica destacada.
  Inicio              vira a tela inicial de um aplicativo: botoes grandes das
                      tarefas do dia, "Situacao agora" (lote em uso e validade,
                      ultima corrida, banco, alertas) e "Como usar" em 3 passos.
  Painel              analito tambem pela LISTA (alem do spinner); faixa de
                      status do lote colorida (verde / ambar / vermelho);
                      atalho "Media/DP deste lote".
  Analitos            atalho "Trocar lote" (vai ao seletor do Painel).
  Configuracao        botao "Novo lote" (cadastro guiado).
"""

NAV = [  # chave, macro, rotulo
    ('inicio', 'IrInicio', 'Início'),
    ('painel', 'IrPainel', 'Painel'),
    ('lancar', 'IrLancar', '+ Lançar'),
    ('analitos', 'IrAnalitos', 'Média/DP'),
    ('lotes', 'IrLotes', 'Lotes'),
    ('estat', 'IrEstatistica', 'Estatística'),
    ('liber', 'IrLiberacao', 'Liberação'),
    ('reg', 'IrRegistros', 'Registros'),
    ('eventos', 'IrEventos', 'Westgard'),
]
ATIVO = {'Início': 'inicio', 'Painel': 'painel', 'Importar': 'lancar', 'Resultados': 'lancar',
         'Analitos': 'analitos', 'Configuração': 'lotes', 'Estatística': 'estat', 'Liberação': 'liber',
         'Registros': 'reg', 'Eventos_Westgard': 'eventos'}
TECNICAS = {'Calc', 'Eng_Saida', 'LotesStore', 'LiberStore', 'RegistrosStore', 'DB_Resultados',
            'Cfg_Status', 'Cfg_Westgard_Escopo', 'Cfg_PlanoQC', 'EQA_Base', 'BI_Data', 'Login',
            'T_Entrada', 'Audit_Legenda', 'Audit_Log', 'EQC_Dados', 'Usuarios', 'Corridas'}

# medidas da barra: 9 botoes de 66 pt + 3 de respiro, mais o atalho do Painel
NAV_W, NAV_H, NAV_GAP = 66, 19, 3
EXTRA_PAINEL = 128            # "✎ Média/DP deste lote"
LARG_BARRA = len(NAV) * (NAV_W + NAV_GAP) + EXTRA_PAINEL + 8

AZUL_ESCURO = 0x2C2813        # BGR de #13282C (faixa de titulo)
VERDE_TXT = 0xB0D63C          # #3CD6B0 (titulo)
BTN = 0x6B6F1F                # #1F6F6B
BTN_ATIVO = 0xB0D63C
BRANCO = 0xFFFFFF


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def _botao(ws, nome, macro, texto, left, top, w, h, fundo, cor_txt, tam=9, negrito=True):
    try:
        ws.Shapes(nome).Delete()
    except Exception:
        pass
    s = ws.Shapes.AddShape(5, left, top, w, h)       # retangulo arredondado
    s.Name = nome
    s.OnAction = macro
    s.Placement = 3                                   # nao se move/redimensiona com celulas
    s.Line.Visible = 0
    s.Fill.ForeColor.RGB = fundo
    try:
        s.Adjustments.SetItem(1, 0.25)
    except Exception:
        pass
    tf = s.TextFrame2
    tf.MarginLeft = tf.MarginRight = 2
    tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3                             # meio
    tf.WordWrap = True
    tr = tf.TextRange
    tr.Text = texto
    tr.ParagraphFormat.Alignment = 2                  # centro
    tr.Font.Size = tam
    tr.Font.Bold = -1 if negrito else 0
    tr.Font.Name = 'Segoe UI'
    tr.Font.Fill.ForeColor.RGB = cor_txt
    return s


def larg_texto(ws, texto, fonte, tam, negrito):
    """Largura REAL do texto em pontos (caixa temporaria com AutoSize).

    A estimativa antiga -- len(texto) * tamanho * 0,62 -- errava por mais de
    15% em fonte proporcional, e era ela que deixava a barra de navegacao por
    cima do titulo da linha 1 (ADR-055).
    """
    try:
        s = ws.Shapes.AddTextbox(1, 0, 800, 10, 20)
    except Exception:
        return len(texto) * float(tam or 11) * 0.62 + 24     # planilha protegida
    try:
        tf = s.TextFrame2
        tf.WordWrap = False
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tr = tf.TextRange
        tr.Text = texto or ' '
        tr.Font.Name = fonte
        tr.Font.Size = tam
        tr.Font.Bold = -1 if negrito else 0
        tf.AutoSize = 1
        return float(s.Width)
    finally:
        try:
            s.Delete()
        except Exception:
            pass


def barra_navegacao(ws, extras=(), inicio=None):
    """Botoes de navegacao no alto da tela, alinhados a direita da faixa de titulo.
    extras: [(nome, macro, texto, largura)] -- atalhos proprios da tela, a esquerda.
    inicio: celula onde a barra comeca. No Painel vem de painel.coluna_da_barra:
            a primeira coluna depois do titulo MAIS uma vazia (ADR-055)."""
    for s in list(ws.Shapes):
        if s.Name.startswith('nav_'):
            s.Delete()
    a1 = ws.Range('A1').MergeArea
    topo = ws.Range('A1').Top
    alt = a1.Height
    if alt < 24:
        alt = ws.Range('A1:A2').Height
    w, h, gap = NAV_W, NAV_H, NAV_GAP
    larg_extras = sum(e[3] + 8 for e in extras)
    total = len(NAV) * (w + gap) + larg_extras
    direita = a1.Left + a1.Width
    # nunca por cima do titulo -- e "nunca" so vale com a largura MEDIDA
    f = ws.Range('A1').Font
    larg_titulo = larg_texto(ws, str(ws.Range('A1').Value or ''),
                             str(f.Name), float(f.Size or 11), bool(f.Bold)) + 18
    if inicio:
        left0 = max(ws.Range(inicio).Left, a1.Left + larg_titulo)
    else:
        left0 = max(a1.Left + larg_titulo, direita - total - 6)
    top = topo + max(1, (alt - h) / 2)
    ativo = ATIVO.get(ws.Name, '')
    x = left0
    for nome, macro, texto, larg in extras:
        _botao(ws, nome, macro, texto, x, top, larg, h, rgb(31, 111, 60), BRANCO, tam=8)
        x += larg + 8
    for i, (k, macro, rot) in enumerate(NAV):
        eh = (k == ativo)
        _botao(ws, f'nav_{k}', macro, rot, x + i * (w + gap), top, w, h,
               BTN_ATIVO if eh else BTN, AZUL_ESCURO if eh else BRANCO, tam=8)
    # a faixa escura do titulo acompanha a barra (em telas de titulo estreito
    # os botoes passariam do fim da faixa, sobre fundo branco)
    fim_barra = left0 + total + 8
    if fim_barra > direita:
        c0 = a1.Column + a1.Columns.Count
        c = c0
        while ws.Cells(1, c).Left < fim_barra and c < c0 + 60:
            c += 1
        linhas = a1.Rows.Count if alt == a1.Height else 2
        cor = ws.Range('A1').Interior.Color
        ws.Range(ws.Cells(1, c0), ws.Cells(linhas, c)).Interior.Color = cor


def telas_de_uso(wb):
    return [ws for ws in wb.Worksheets if ws.Name not in TECNICAS]


def inicio(ws, produto):
    """Tela inicial: tarefas, situacao agora e 'como usar'."""
    for s in list(ws.Shapes):
        if s.Name.startswith('tile_') or s.Name.startswith('nav_'):
            s.Delete()
    # a lista de links vira a grade de botoes
    ws.Range('B11:H40').UnMerge()
    for c in ws.Range('B12:G20').Cells:
        try:
            if c.Hyperlinks.Count:
                c.Hyperlinks.Delete()
        except Exception:
            pass
    ws.Range('B11:H40').Clear()
    ws.Range('B11').Value = 'O QUE VOCÊ QUER FAZER?'
    ws.Range('B11').Font.Bold = True
    ws.Range('B11').Font.Size = 12
    ws.Range('B11').Font.Color = rgb(19, 40, 44)
    tiles = [
        ('painel', 'IrPainel', 'Painel', 'gráficos de Levey-Jennings e Westgard por lote'),
        ('lancar', 'IrLancar', '+ Lançar resultados', 'a corrida do dia (ou várias de uma vez)'),
        ('novolote', 'NovoLote', 'Novo lote de controle', 'cadastrar, usar e digitar média/DP'),
        ('analitos', 'IrAnalitos', 'Média e DP por lote', 'parâmetros e especificações'),
        ('estat', 'IrEstatistica', 'Estatística e Sigma', 'CV, bias, ET, Sigma e plano de CQ'),
        ('liber', 'IrLiberacao', 'Liberação', 'iniciar e finalizar a corrida (assinatura)'),
        ('reg', 'IrRegistros', 'Repetições e calibração', 'aparecem marcadas no gráfico'),
        ('eqa', 'IrEQA', 'Controle externo (EQA)', 'rodadas CAP / Controllab, SDI e bias'),
    ]
    x0 = ws.Range('B12').Left
    y0 = ws.Range('B12').Top + 2
    w, h, gx, gy = 205, 46, 10, 10
    for i, (k, macro, tit, sub) in enumerate(tiles):
        col, lin = i % 4, i // 4
        s = _botao(ws, f'tile_{k}', macro, tit + '\n' + sub, x0 + col * (w + gx), y0 + lin * (h + gy), w, h,
                   BTN if k != 'lancar' else rgb(31, 111, 60), BRANCO, tam=11)
        n1 = len(tit)
        try:                                          # subtitulo menor e sem negrito
            ch = s.raw.TextFrame.Characters(n1 + 2, len(sub))
            ch.Font.Size = 8
            ch.Font.Bold = False
        except Exception:
            pass
    # ---- situacao agora ----
    r0 = 21
    ws.Range(f'B{r0}').Value = 'SITUAÇÃO AGORA'
    ws.Range(f'B{r0}').Font.Bold = True
    ws.Range(f'B{r0}').Font.Size = 12
    ws.Range(f'B{r0}').Font.Color = rgb(19, 40, 44)
    linhas = [
        ('Lote em uso (lançamentos)',
         '=loteAtivo&IF(Configuração!$C$21="",""," · validade "&TEXT(Configuração!$C$21,"dd/mm/aaaa")'
         '&IF(Configuração!$C$21-TODAY()<0,"  ⛔ VENCIDO",IF(Configuração!$C$21-TODAY()<=30,"  ⚠ vence em "'
         '&(Configuração!$C$21-TODAY())&" dia(s)","")))'),
        ('Lote em análise (Painel)', '=loteAnalise&"  ·  escolha outro na lista do Painel"'),
        ('Última corrida registrada', '=IF(COUNT(rData)=0,"nenhuma",TEXT(MAX(rData),"dd/mm/aaaa"))'),
        ('Resultados no banco',
         '=TEXT(COUNT(rValor),"#.##0")&" de "&TEXT(capBanco,"#.##0")&" ("&TEXT(COUNT(rValor)/capBanco,"0%")&" da capacidade)"'),
        ('Lotes cadastrados', '=COUNTA(Configuração!$C$26:$C$125)&" de 100"'),
        ('Analitos sem média/DP no lote em uso',
         '=SUMPRODUCT((Analitos!$A$4:$A$43<>"")*(COUNTIFS(lsChave,UPPER(TRIM(loteAtivo&""))&"|"'
         '&UPPER(TRIM(Analitos!$A$4:$A$43)),lsMedN1,"<>")=0))'),
    ]
    for i, (rot, f) in enumerate(linhas):
        r = r0 + 1 + i
        ws.Range(f'B{r}').Value = rot
        ws.Range(f'C{r}').Formula = f
        ws.Range(f'C{r}:G{r}').Merge()
        ws.Range(f'B{r}:G{r}').Interior.Color = rgb(237, 246, 244)
        ws.Range(f'B{r}').Font.Bold = True
        ws.Range(f'C{r}').HorizontalAlignment = -4131
        ws.Range(f'B{r}:G{r}').Borders(9).Color = rgb(200, 220, 216)
    ultima = r0 + len(linhas)
    # alerta: analitos sem parametro em vermelho
    c = ws.Range(f'C{ultima}')
    c.FormatConditions.Delete()
    fc = c.FormatConditions.Add(1, 5, '=0')           # xlCellValue, xlGreater
    fc.Font.Color = rgb(192, 0, 0)
    fc.Font.Bold = True
    # ---- como usar ----
    r1 = ultima + 2
    ws.Range(f'B{r1}').Value = 'COMO USAR — 3 PASSOS'
    ws.Range(f'B{r1}').Font.Bold = True
    ws.Range(f'B{r1}').Font.Size = 12
    ws.Range(f'B{r1}').Font.Color = rgb(19, 40, 44)
    passos = [
        '1.  Chegou lote novo de controle?  Clique em "Novo lote de controle": cadastra, pergunta se passa a '
        'ser o lote dos lançamentos e abre a digitação de média e DP.',
        '2.  Todo dia:  "+ Lançar resultados".  Os gráficos, o Westgard e a Estatística se atualizam sozinhos.',
        '3.  Para analisar:  "Painel".  Escolha o LOTE e o ANALITO nas listas (ou ▲▼).  Gráfico, limites, '
        'Westgard e Estatística usam SÓ a média/DP do lote escolhido — lote antigo continua consultável.',
        'Dica:  a barra no alto de cada tela leva a qualquer lugar.  Cada troca leva menos de 1 segundo.',
    ]
    for i, t in enumerate(passos):
        r = r1 + 1 + i
        ws.Range(f'B{r}').Value = t
        ws.Range(f'B{r}').Font.Size = 10
        ws.Range(f'B{r}:H{r}').Merge()
        ws.Range(f'B{r}').WrapText = True
        ws.Rows(r).RowHeight = 28
    _botao(ws, 'tile_guias', 'AlternarGuias', 'mostrar/ocultar guias das planilhas',
           ws.Range(f'B{r1 + len(passos) + 1}').Left, ws.Range(f'B{r1 + len(passos) + 1}').Top + 4, 190, 16,
           rgb(237, 246, 244), rgb(31, 111, 107), tam=7, negrito=False)
    ws.Range(f'B{r1 + len(passos)}').Font.Italic = True
    ws.Range(f'B{r1 + len(passos)}').Font.Color = rgb(89, 89, 89)
    for r in range(12, 22):
        ws.Rows(r).RowHeight = 15


def _ao_lado_da_barra(ws, nome, macro, texto, largura, fundo, cor):
    """Botao extra da tela, a esquerda da barra de navegacao (mesma altura)."""
    ref = ws.Shapes('nav_inicio')
    return _botao(ws, nome, macro, texto, ref.Left - largura - 10, ref.Top, largura, ref.Height,
                  fundo, cor, tam=8)


def painel(ws, cel_lote, faixa, inicio_nav='I1'):
    """Status do lote colorido; atalho Media/DP deste lote."""
    o3 = ws.Range(faixa)
    o3.FormatConditions.Delete()
    # a faixa mudou de lugar (ADR-054): a regra aponta para a PROPRIA faixa
    ref = ws.Range(faixa.split(':')[0]).Address      # ja vem absoluto ($A$39)
    # FORMATACAO CONDICIONAL FALA PORTUGUES (ADR-039): a formula vai em pt-BR
    c1 = o3.FormatConditions.Add(2, None, f'=ESQUERDA({ref};1)="⛔"')
    c1.Interior.Color = rgb(253, 226, 225)
    c1.Font.Color = rgb(156, 0, 6)
    c1.StopIfTrue = True
    c2 = o3.FormatConditions.Add(2, None, f'=ÉNÚM(LOCALIZAR("Nenhuma corrida";{ref}))')
    c2.Interior.Color = rgb(255, 244, 206)
    c2.Font.Color = rgb(122, 80, 0)
    c2.StopIfTrue = True
    c3 = o3.FormatConditions.Add(2, None, f'={ref}<>""')
    c3.Interior.Color = rgb(227, 244, 234)
    c3.Font.Color = rgb(19, 40, 44)
    o3.Font.Bold = True
    o3.WrapText = True
    o3.VerticalAlignment = -4108
    barra_navegacao(ws, [('nav_mediadp', 'IrParametrosDoAnalito', '✎ Média/DP deste lote',
                          EXTRA_PAINEL)], inicio=inicio_nav)


def analitos(ws):
    barra_navegacao(ws, [('nav_trocarlote', 'IrLotePainel', 'Trocar lote ▸', 96)])


def configuracao(ws):
    barra_navegacao(ws, [('nav_novolote', 'NovoLote', '+ Novo lote', 96)])
    alvo = ws.Range('D23')
    _botao(ws, 'nav_novolote2', 'NovoLote', '+ Novo lote', alvo.Left + alvo.Width + 6, alvo.Top - 2, 96, 20,
           rgb(31, 111, 60), BRANCO, tam=9)
    ws.Range('B24').Value = ('Clique em "+ Novo lote" (cadastro guiado) ou digite na próxima linha livre. '
                             'Lote = núcleo do código, sem "QC-" e sem os 2 dígitos finais de nível.')


def aplicar(wb, produto, cel_lote, faixa, cap, inicio_nav='I1'):
    """Instala tudo. produto: 'Bioquímica' | 'Hematologia'.

    inicio_nav: coluna onde a barra do Painel comeca (painel.coluna_da_barra).
    """
    nm = 'capBanco'
    try:
        wb.Names(nm).Delete()
    except Exception:
        pass
    wb.Names.Add(nm, f'={cap}')
    sh = {s.Name: s for s in wb.Worksheets}
    for ws in telas_de_uso(wb):
        if ws.Name != 'Início':               # no Inicio os botoes grandes ja sao a navegacao
            barra_navegacao(ws)
    inicio(sh['Início'], produto)
    painel(sh['Painel'], cel_lote, faixa, inicio_nav)
    analitos(sh['Analitos'])
    configuracao(sh['Configuração'])
    # C3 do Painel (analito) tambem pela lista -- a Hematologia ja recebe no portar
    c3 = sh['Painel'].Range('C3')
    try:
        c3.Validation.Delete()
    except Exception:
        pass
    c3.Validation.Add(3, 1, 1, '=lstAnalitos')
    c3.Validation.ShowError = False
    c3.Validation.InputTitle = 'Analito'
    c3.Validation.InputMessage = 'Escolha o analito na lista (ou use o spinner ▲▼).'
    c3.MergeArea.Locked = False
