# -*- coding: utf-8 -*-
"""instalar_adr064.py -- ADR-064/ADR-067: INCERTEZA DE MEDICAO (Nordtest TR 537: CIQ + CEQ) na aba Estatistica e no
Painel, e correcoes do CEQ achadas no caminho.

Uso:  python instalar_adr064.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Idempotente. So:
  1. instala todo o VBA atual das fontes (codigo_atual.py: mIncerteza novo; mCEQ.ViesEQ; mEQA.NLabsPorChave;
     texto acentuado corrigido em mEQA/mPlanoQC); EXIGE compilacao;
  2. Hematologia: Estatistica R, S, T, AC, AD passam a usar o provedor do filtro (eqProvedor), como a G ja
     usava -- liam Analitos!AR, coluna que so existe na Bioquimica, e davam "SEM EP" sempre;
  3. Analitos: REMOVE o campo "u(cal) % fabricante" de instalacoes antigas (ADR-067: sem certificado de
     calibrador; a incerteza e so CIQ + CEQ);
  4. Estatistica: janela de 12 meses (MU_Ini/MU_Fim) e as colunas AF:AU da incerteza na mesma tabela
     analito|nivel (nome MU_Faixa, usado por mIncerteza.RecalcularIncerteza);
  5. Painel: bloco "INCERTEZA DE MEDICAO -- analito selecionado" inserido ANTES de "MARGEM CRITICA"
     (localizado pelo rotulo, nunca por numero de linha), lendo da Estatistica pela chave analito|nivel;
  6. FUMACA: compila, recalcula e confere numeros de verdade (u(Rw) > 0, U = 2 uc, status valido);
  7. devolve a protecao de cada aba e a da estrutura como estavam.
"""
import os
import sys

import pythoncom

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

XL_PASTE_FORMATS = -4122
C0 = 32          # AF: primeira coluna da incerteza na Estatistica (AE fica de respiro)
LIN_TAB = 14     # primeira linha da tabela analito|nivel
MARCA_PAINEL = 'INCERTEZA DE MEDIÇÃO'
# textos de mIncerteza.ClassificarIncerteza (com MU_MESES_MIN = 6 e MU_N_MIN = 60)
STATUS_CORES = (('Não atende', 13551615), ('Mínimo', 10284031), ('Desejável', 13561798), ('Ótimo', 13561798),
                ('Sem meta', 15921906), ('Insuficiente', 15921906))
VIES_CORES = (('RELEVANTE: investigar/corrigir', 13551615),
              ('detectável (relevância não avaliada)', 10284031), ('detectável, dentro do permitido', 10284031),
              ('não detectável', 13561798))


def log(*a):
    print(*a, flush=True)


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def achar_cabecalho(ws, linha, texto, ate=80):
    for c in range(1, ate):
        if str(ws.Cells(linha, c).Value or '').strip().lower() == texto.lower():
            return c
    return 0


def ultima_linha_tabela(e):
    ult = LIN_TAB
    for r in range(LIN_TAB, 400):
        if str(e.Cells(r, 28).Formula or '').startswith('='):      # AB = chave analito|nivel
            ult = r
        elif r > ult + 3:
            break
    return ult


def formato_data(xl):
    """dd/mm/aaaa nos codigos do Excel INSTALADO (International: 21 dia, 20 mes, 19 ano). NumberFormat
    'dd/mm/yyyy' por automacao num Excel em portugues mostra 'yyyy' literal -- medido na Hematologia (AF3/AF4)."""
    t = xl.International                   # sem indice a automacao devolve a lista toda (base 1 no Excel)
    d, m, y = t[21 - 1], t[20 - 1], t[19 - 1]
    if not all(isinstance(x, str) and len(x) == 1 for x in (d, m, y)):
        raise SystemExit(f'codigos de data inesperados: {d!r} {m!r} {y!r}')
    return f'{d}{d}/{m}{m}/{y}{y}{y}{y}'


def nf_local(xl, nf):
    """Formato numerico nos codigos do Excel INSTALADO. Por esta automacao o NumberFormat e lido na lingua
    local: '0.00' num Excel em portugues vira separador de MILHAR e mostra '001' (medido). '@' e '0' valem
    em qualquer lingua."""
    if nf in ('@', '0'):
        return nf
    casas = len(nf.split('.')[1]) if '.' in nf else 0
    dec = xl.International[3 - 1]                     # xlDecimalSeparator
    return '0' + (dec + '0' * casas if casas else '')


def selecionar_a1(ws):
    """Devolve a selecao para A1 (o PasteSpecial deixa a faixa colada selecionada -- e isso e salvo no
    arquivo). Na entrega blindada a aba esta OCULTA: mostra so durante o ajuste e devolve. Cosmetico: nunca
    derruba a instalacao."""
    try:
        ws.Application.CutCopyMode = False
        vis = ws.Visible
        if vis != -1:
            ws.Visible = -1
        ws.Activate()
        ws.Range('A1').Select()
        if vis != -1:
            ws.Visible = vis
    except Exception as e:
        log(f'   (selecao em A1 nao ajustada em {ws.Name}: {e})')


def cores_status(rng, regras):
    """Formatacao condicional por VALOR IGUAL a um texto (xlCellValue, xlEqual, ="texto"): constante de
    texto, sem funcao nem separador -- nao depende do idioma do Excel (formula de FormatConditions e lida na
    lingua da instalacao). O 'texto contem' (xlTextString) nao aceita os argumentos pela automacao."""
    rng.FormatConditions.Delete()
    for txt, cor in regras:
        fc = rng.FormatConditions.Add(1, 3, '="' + txt.replace('"', '""') + '"')      # xlCellValue, xlEqual
        fc.Interior.Color = cor


def provedor_padrao_hema(e):
    """Hematologia: eqProvedor apontava para L4, celula do MEIO da nota mesclada F4:M4 -- nunca guardou valor
    ("todos os provedores"). Com dados do Controllab, bias, SDI, limites e a triagem do vies misturariam CAP e
    Controllab; o contador de K5 mostrava 0 linhas (auditoria 04/10/2026). O filtro ganha celula propria
    (Q4 rotulo, R4 valor), padrao CAP, editavel, com lista; o nome passa a apontar para ela."""
    wb = e.Parent
    try:
        alvo = wb.Names('eqProvedor').RefersToRange
    except Exception:                                # referencia invalida: refaz em R4
        alvo = e.Range('L4')
    if alvo.MergeArea.Count == 1 and str(alvo.Value or '').strip():
        return f'{alvo.Address} = {alvo.Value} (ja definido)'
    if alvo.MergeArea.Count > 1:                     # celula presa numa mescla: muda de lugar
        for a in ('Q4', 'R4'):
            if e.Range(a).MergeCells or str(e.Range(a).Formula or '').strip() not in ('', 'Provedor', 'CAP', 'Controllab'):
                raise SystemExit(f'Estatistica!{a} ocupada ({e.Range(a).Formula!r}): nao da para mover o provedor')
        e.Range('Q4').Value = 'Provedor'
        e.Range('O4').Copy()                          # mesmo visual do rotulo "Rodada"
        e.Range('Q4').PasteSpecial(XL_PASTE_FORMATS)
        e.Range('P4').Copy()                          # e da celula do filtro de rodada
        e.Range('R4').PasteSpecial(XL_PASTE_FORMATS)
        alvo = e.Range('R4')
        # repontar SEM apagar (apagar quebraria as formulas). Names.Add sobre o nome existente: o .RefersTo
        # por esta automacao foi lido em L1C1 local e gravou "=Estatística!L4C18" (medido)
        wb.Names.Add('eqProvedor', "='Estatística'!$R$4")
        if wb.Names('eqProvedor').RefersToRange.Address != '$R$4':
            raise SystemExit(f"eqProvedor nao foi repontado: {wb.Names('eqProvedor').RefersTo}")
    alvo.Value = 'CAP'
    alvo.Locked = False
    sep = e.Application.International[5 - 1]             # xlListSeparator
    alvo.Validation.Delete()
    alvo.Validation.Add(3, 1, 1, 'CAP' + sep + 'Controllab')   # xlValidateList
    return f'{alvo.Address} = CAP (padrao, lista CAP/Controllab)'


def corrigir_provedor_hema(e, ult):
    """Hematologia: R, S, T, AC, AD procuravam o provedor em Analitos!AR (so a Bio tem)."""
    velho = 'IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A{r},Analitos!$A$4:$A$43,0)),"CAP")'
    n = 0
    for r in range(LIN_TAB, ult + 1):
        for c in ('R', 'S', 'T', 'AC', 'AD'):
            f = str(e.Range(f'{c}{r}').Formula)
            v = velho.format(r=r)
            if v in f:
                e.Range(f'{c}{r}').Formula = f.replace(v, 'eqProvedor')
                n += 1
    return n


def remover_ucal(wb):
    """ADR-067: o laboratorio nao usa certificado de calibrador. Limpa o campo e o nome de instalacoes antigas
    (a coluna fica vazia, nao e excluida: excluir deslocaria referencias de outras colunas da Analitos)."""
    a = wb.Worksheets('Analitos')
    c = achar_cabecalho(a, 3, 'u(cal) % fabricante (k=1)', 80)
    if c:
        rng = a.Range(a.Cells(2, c), a.Cells(43, c))
        rng.Validation.Delete()
        rng.ClearContents()
        rng.ClearFormats()
        rng.Locked = True
    for n in list(wb.Names):
        if n.Name == 'ucalFabricante':
            n.Delete()
    return col(c) if c else None


def coluna_cvi(wb):
    a = wb.Worksheets('Analitos')
    c = achar_cabecalho(a, 3, 'CVi %', 60)
    if not c:
        raise SystemExit('Analitos: coluna "CVi %" nao encontrada na linha 3')
    return f"Analitos!${col(c)}$4:${col(c)}$43"


def bloco_estatistica(wb, ult, exc, prov, cvi):
    e = wb.Worksheets('Estatística')
    m = f'MATCH($A{{r}},Analitos!$A$4:$A$43,0)'
    ciq = 'mIncerteza.IncertezaCIQ($A{r},$B{r},"%s",MU_Ini,MU_Fim' + exc + ')'
    val = ciq % 'VALIDADE'
    cvim = 'mIncerteza.CVIMeta(IFERROR(INDEX(' + cvi + ',' + m + '),""),$I{r},$J{r})'
    vies = 'mCEQ.ViesEQ($A{r},eqAnoEP,"%s",' + prov + ',eqRodada)'
    bperm = 'IF(AND($I{r}="VB",ISNUMBER($K{r}),ISNUMBER($J{r})),$K{r}-1.65*$J{r},"")'   # ADR-068: ETp pode ser "-"
    cab = [
        ('u(Rw) %\nCIQ 12 m, lotes n≥20', '=IF(OR($A{r}="",$A{r}=0),"",IF(' + val + '="INSUFICIENTE","",' + (ciq % 'CV') + '))', '0.00'),
        ('gl\nΣ(n−1)', '=IF(OR($A{r}="",$A{r}=0),"",' + (ciq % 'GL') + ')', '0'),
        ('Lotes\n(n ≥ 20)', '=IF(OR($A{r}="",$A{r}=0),"",' + (ciq % 'LOTES') + ')', '0'),
        ('Dias\ncobertos', '=IF(OR($A{r}="",$A{r}=0),"",' + (ciq % 'DIAS') + ')', '0'),
        ('u(bias) %\nCEQ (Nordtest)', '=IF(OR($A{r}="",$A{r}=0),"",IFERROR(' + (vies % 'UBIAS') + ',""))', '0.00'),
        ('uc %\n(k=1)', '=IF(AND(ISNUMBER($AF{r}),ISNUMBER($AJ{r})),SQRT($AF{r}^2+$AJ{r}^2),"")', '0.00'),
        ('U %\n(k=2, ~95%)', '=IF(ISNUMBER($AK{r}),2*$AK{r},"")', '0.00'),
        ('U na unidade\n(média dos lotes)', '=IF(ISNUMBER($AL{r}),$AL{r}*' + (ciq % 'MEDIA') + '/100,"")', '0.000'),
        ('Meta u %\n(0,50 × CVI)', '=IF(OR($A{r}="",$A{r}=0),"",IFERROR(0.5*' + cvim + ',""))', '0.00'),
        ('Classe da\nincerteza', '=IF(OR($A{r}="",$A{r}=0),"",mIncerteza.ClasseIncerteza($AK{r},' + cvim + ',' + val + '))', None),
        ('Situação', '=IF(OR($A{r}="",$A{r}=0),"",mIncerteza.SituacaoIncerteza(' + val + ',$AJ{r},' + (vies % 'SITUACAO') + ',' + (vies % 'N') + '))', None),
        ('Alertas', '=IF(OR($A{r}="",$A{r}=0),"",mIncerteza.AlertasIncerteza($A{r},$B{r},MU_Ini,MU_Fim,$AF{r},' + cvim + ',$AJ{r}' + exc + '))', None),
        ('Viés CEQ médio %\n(com sinal)', '=IF(OR($A{r}="",$A{r}=0),"",' + (vies % 'MEDIA') + ')', '0.00'),
        ('Viés do CEQ\n(triagem)', '=IF(OR($A{r}="",$A{r}=0),"",mCEQ.ViesEQ($A{r},eqAnoEP,"TRIAGEM",' + prov + ',eqRodada,' + bperm + '))', None),
        ('RMS viés % / u(Cref) %\n(CEQ)', '=IF(OR($A{r}="",$A{r}=0),"",IFERROR(TEXT(' + (vies % 'RMS') + ',"0.00")&" / "&IFERROR(TEXT(' + (vies % 'UCREF') + ',"0.00"),"—"),""))', None),
        ('CEQ amostras\n/ rodadas', '=IF(OR($A{r}="",$A{r}=0),"",' + (vies % 'N') + '&" / "&' + (vies % 'NRODADAS') + ')', None),
    ]
    assert col(C0) == 'AF' and col(C0 + 4) == 'AJ' and col(C0 + 14) == 'AT', 'colunas fora do lugar'
    cl = col(C0 + len(cab) - 1)                                          # AU
    # cabecalho do bloco (linhas 9 a 12) e a janela de 12 meses
    e.Range(f'AF9:{cl}12').UnMerge()
    e.Range('AF9').Value = 'INCERTEZA DE MEDIÇÃO — Nordtest TR 537 (top-down: CIQ + CEQ)'
    e.Range('AF9').Font.Bold = True
    e.Range('AF9').Font.Size = 12
    e.Range('AF10').Value = ('u(Rw) = CIQ de longo prazo (12 meses): CV de cada lote de controle (n ≥ 20) agrupado pelos graus '
                             'de liberdade · u(bias) = CEQ: raiz(RMS dos vieses² + u(Cref)²), u(Cref) = DP do grupo / raiz(nº de '
                             'laboratórios), ≥ 6 amostras de ≥ 2 rodadas (preferível 10) · uc = raiz(u(Rw)² + u(bias)²) · U = 2 × uc · '
                             'meta: u ≤ 0,50 × CVI (EFLM; ótimo 0,25, mínimo 0,75) · CIQ válido com ≥ 180 dias e gl ≥ 100; provisório '
                             'com 90–179 dias ou gl 30–99 · sem u(bias) o U não é estimado (u(Rw) sozinha não é a incerteza) · '
                             'ALERTAS: lotes curtos; CV heterogêneo = CVmáx/CVmín > 2; inativados > 5%; componente dominante')
    e.Range('AF10').Font.Size = 9
    e.Range('AF11').Value = 'Janela (12 m)'
    e.Range('AF11').Font.Bold = True
    # janela MOVEL de 12 meses terminando no ultimo resultado (revisao metodologica do ADR-064): a incerteza
    # e sempre a atual. Amarrar ao fim do periodo da aba deixava a Hematologia (Ano De/Ate = 2025) com 2 meses
    # de dados e nenhuma estimativa -- o periodo da aba continua mandando no CV do lote, nao na incerteza
    # so o que PARTICIPA da estatistica: uma data futura digitada num manual vira MANUAL_INCOMPLETO mas
    # continua na tabela, e o MAX da coluna inteira jogava a janela no futuro (auditoria 04/10/2026)
    e.Range('AI11').Formula = ('=IFERROR(IF(MAXIFS(tblCQ_Final[DATA],tblCQ_Final[PARTICIPA_ESTATISTICA],"SIM")>0,'
                               'INT(MAXIFS(tblCQ_Final[DATA],tblCQ_Final[PARTICIPA_ESTATISTICA],"SIM")),""),"")')
    e.Range('AG11').Formula = '=IF(ISNUMBER(AI11),EDATE(AI11,-12)+1,"")'
    e.Range('AH11').Value = 'a'
    fd = formato_data(e.Application)
    for a in ('AG11', 'AI11'):
        e.Range(a).NumberFormatLocal = fd
    if str(e.Range('AE3').Value or '').startswith('Início efetivo'):     # Hematologia: mesmo defeito, antigo
        e.Range('AF3:AF4').NumberFormatLocal = fd
    e.Range('AJ11').Value = ('o viés do CEQ ENTRA em U pelo RMS (Nordtest) e também é TRIADO (média com sinal; detectável se '
                             '|média| > 2 EP): viés relevante se investiga/corrige')
    e.Range('AJ11').Font.Size = 9
    e.Range('AJ11').Font.Italic = True
    ii.nome(wb, 'MU_Ini', "='Estatística'!$AG$11")
    ii.nome(wb, 'MU_Fim', "='Estatística'!$AI$11")
    # cabecalho da tabela e corpo
    e.Range('AD13').Copy()
    e.Range(f'AF13:{cl}13').PasteSpecial(XL_PASTE_FORMATS)
    e.Range(f'AF13:{cl}13').WrapText = True
    # REINSTALACAO: celula ja formatada como texto ('@', instalacao antiga) guarda a formula como TEXTO
    # (medido: todas as linhas viravam a formula literal da linha 14). Limpa o formato antes de escrever;
    # o formato e reaplicado logo abaixo.
    e.Range(f'AF{LIN_TAB}:{cl}{ult}').ClearFormats()
    for j, (titulo, f, nf) in enumerate(cab):
        c = C0 + j
        e.Cells(13, c).Value = titulo
        e.Range(e.Cells(LIN_TAB, c), e.Cells(ult, c)).Formula = f.format(r=LIN_TAB)   # relativo: o Excel propaga
        e.Columns(c).ColumnWidth = 12
    e.Range(f'AC{LIN_TAB}:AC{ult}').Copy()
    e.Range(f'AF{LIN_TAB}:{cl}{ult}').PasteSpecial(XL_PASTE_FORMATS)
    for j, (titulo, f, nf) in enumerate(cab):
        c = C0 + j
        if nf:                                        # texto: sem '@' (formula em celula '@' vira texto)
            e.Range(e.Cells(LIN_TAB, c), e.Cells(ult, c)).NumberFormatLocal = nf_local(e.Application, nf)
    for c in (C0 + 10, C0 + 11, C0 + 13):            # situacao, alertas, triagem: texto a esquerda
        e.Range(e.Cells(LIN_TAB, c), e.Cells(ult, c)).HorizontalAlignment = -4131
    e.Columns(C0 + 9).ColumnWidth = 14        # classe
    e.Columns(C0 + 10).ColumnWidth = 40       # situacao
    e.Columns(C0 + 11).ColumnWidth = 60       # alertas
    e.Columns(C0 + 13).ColumnWidth = 34       # triagem do vies
    # cores do status
    st = e.Range(f'{col(C0 + 9)}{LIN_TAB}:{col(C0 + 9)}{ult}')
    cores_status(st, STATUS_CORES)
    cores_status(e.Range(f'{col(C0 + 13)}{LIN_TAB}:{col(C0 + 13)}{ult}'), VIES_CORES)
    ii.nome(wb, 'MU_Faixa', f"='Estatística'!$AF${LIN_TAB}:${cl}${ult}")
    selecionar_a1(e)                                  # nao deixa a faixa colada selecionada no arquivo
    return cl


def bloco_painel(wb, nlv):
    p = wb.Worksheets('Painel')
    # idempotente: bloco ja instalado e so reescrito
    t = None
    alvo = None
    for r in range(40, 200):
        v = str(p.Cells(r, 1).Value or '')
        if v.startswith(MARCA_PAINEL):
            alvo = r
        if v.startswith('ERRO TOTAL') and t is None:
            t = r
        if v.startswith('MARGEM CR'):
            marg = r
            break
    else:
        raise SystemExit('Painel: rotulo "MARGEM CRÍTICA" nao encontrado')
    if t is None:
        raise SystemExit('Painel: bloco "ERRO TOTAL vs ETp" nao encontrado (modelo de formato)')
    altura = 4 + nlv                       # titulo, nota, cabecalho, niveis, rodape
    if alvo is None:
        p.Rows(f'{marg}:{marg + altura}').Insert()        # + 1 linha de respiro
        alvo = marg
        t = t if t < alvo else t + altura + 1
    r0 = alvo
    faixa = p.Range(p.Cells(r0, 1), p.Cells(r0 + altura, 21))
    faixa.UnMerge()
    faixa.ClearContents()
    faixa.FormatConditions.Delete()
    # formatos copiados do bloco ERRO TOTAL vs ETp: titulo, cabecalho e linha de nivel
    p.Rows(t).Copy()
    p.Rows(r0).PasteSpecial(XL_PASTE_FORMATS)
    p.Rows(t + 1).Copy()
    p.Rows(r0 + 2).PasteSpecial(XL_PASTE_FORMATS)
    for k in range(nlv):
        p.Rows(t + 2).Copy()
        p.Rows(r0 + 3 + k).PasteSpecial(XL_PASTE_FORMATS)
    p.Rows(r0 + 1).RowHeight = 15
    p.Rows(r0 + 3 + nlv).RowHeight = 15
    # cabecalho das colunas M, O, P (o modelo so ia ate K:L)
    for a in ('M', 'N', 'O', 'P'):                    # modelo J: unica coluna NAO mesclada do bloco
        p.Range(f'J{t + 1}').Copy()
        p.Range(f'{a}{r0 + 2}').PasteSpecial(XL_PASTE_FORMATS)
        for k in range(nlv):
            p.Range(f'J{t + 2}').Copy()
            p.Range(f'{a}{r0 + 3 + k}').PasteSpecial(XL_PASTE_FORMATS)
    p.Range(f'M{r0 + 2}:N{r0 + 2}').Merge()
    for k in range(nlv):
        p.Range(f'M{r0 + 3 + k}:N{r0 + 3 + k}').Merge()
    p.Cells(r0, 1).Value = 'INCERTEZA DE MEDIÇÃO — analito selecionado (Nordtest: CIQ + CEQ)'
    p.Cells(r0 + 1, 1).Formula = '=mIncerteza.NotaIncerteza(MU_Ini,MU_Fim)'
    p.Cells(r0 + 1, 1).Font.Size = 8
    p.Cells(r0 + 1, 1).Font.Italic = True
    cab = [('A', 'Nível', None), ('C', 'u(Rw) %', 'AF'), ('E', 'U % (k=2)', 'AL'), ('H', 'U (unid.)', 'AM'),
           ('J', 'Meta u %', 'AN'), ('K', 'Classe', 'AO'), ('M', 'Situação', 'AP'),
           ('O', 'Viés CEQ médio %', 'AR'), ('P', 'Viés do CEQ (triagem)', 'AS')]
    nf = {'AF': '0.00', 'AL': '0.00', 'AM': '0.000', 'AN': '0.00', 'AR': '0.00'}
    for a, titulo, src in cab:
        p.Range(f'{a}{r0 + 2}').Value = titulo
    for k in range(1, nlv + 1):
        r = r0 + 2 + k
        p.Range(f'A{r}').Value = f'N{k}'
        for a, titulo, src in cab[1:]:
            if p.Range(f'{a}{r}').NumberFormat == '@':      # formula em celula de texto viraria texto
                p.Range(f'{a}{r}').NumberFormatLocal = nf_local(p.Application, '0.00')
            p.Range(f'{a}{r}').Formula = (f"=IFERROR(INDEX('Estatística'!${src}$14:${src}$133,"
                                          f"MATCH(selAnalito&\"|\"&{k},'Estatística'!$AB$14:$AB$133,0)),\"\")")
            if src in nf:
                p.Range(f'{a}{r}').NumberFormatLocal = nf_local(p.Application, nf[src])
    p.Cells(r0 + 3 + nlv, 1).Value = ('u(bias) do CEQ, RMS / u(Cref) e alertas na aba Estatística (AJ, AT, AQ). Base: Nordtest '
                                      'TR 537; Magnusson 2012; Cui 2017; Hann 2017; Martinello 2020; Michaelis 2026; meta EFLM '
                                      '(Braga & Panteghini 2020).')
    p.Cells(r0 + 3 + nlv, 1).Font.Size = 8
    p.Cells(r0 + 3 + nlv, 1).Font.Italic = True
    selecionar_a1(p)
    cores_status(p.Range(f'K{r0 + 3}:L{r0 + 2 + nlv}'), STATUS_CORES)
    cores_status(p.Range(f'P{r0 + 3}:P{r0 + 2 + nlv}'), VIES_CORES)
    for k in range(nlv):                               # textos longos: quebra de linha
        p.Rows(r0 + 3 + k).RowHeight = 30
        for a in ('M', 'P'):
            p.Range(f'{a}{r0 + 3 + k}').WrapText = True
            p.Range(f'{a}{r0 + 3 + k}').Font.Size = 8
    return r0


def acabamento_painel(wb, nlv):
    """Agente 07 (interface), achados na revisao visual de 04/10/2026:
    - titulo do grafico e do eixo X com IncludeInLayout = False: o titulo do nivel flutuava sobre as linhas do
      grafico e "Corrida (RUN)" ficava em cima do rotulo do eixo. Incluidos no layout, a area de plotagem cede;
    - Plano de CQ: o texto da Frequencia de CQ e longo e estava centralizado -- cortado nas duas bordas."""
    p = wb.Worksheets('Painel')
    n = 0
    for co in p.ChartObjects():
        ch = co.Chart
        if ch.HasTitle:
            ch.ChartTitle.IncludeInLayout = True
            n += 1
        ax = ch.Axes(1)
        if ax.HasTitle:
            ax.AxisTitle.IncludeInLayout = True
    plano = [r for r in range(40, 200) if str(p.Cells(r, 1).Value or '').startswith('PLANO DE CQ')]
    feitos = 0
    if plano:
        h = plano[0] + 1
        cf = [c for c in range(1, 22) if str(p.Cells(h, c).Value or '').startswith('Frequência')]
        if cf:
            for k in range(1, nlv + 1):
                cel = p.Cells(h + k, cf[0]).MergeArea
                cel.HorizontalAlignment = -4131
                cel.WrapText = True
                cel.Font.Size = 8
                p.Rows(h + k).RowHeight = 24
                feitos += 1
    return n, feitos


def main(produto, caminho, salvar=True):
    caminho = os.path.abspath(caminho)
    bio = produto.startswith('Bio')
    nlv = 2 if bio else 3
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.Calculation = -4135                      # manual durante a montagem

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        e = wb.Worksheets('Estatística')
        ult = ultima_linha_tabela(e)
        log(f'   tabela da Estatistica: linhas {LIN_TAB}..{ult}')
        if not bio:
            n = corrigir_provedor_hema(e, ult)
            pv = provedor_padrao_hema(e)
            log(f'2. Hematologia: {n} formulas do CEQ passam a usar eqProvedor (R, S, T, AC, AD); provedor {pv}')
        a = wb.Worksheets('Analitos')
        tem_ar = str(a.Range('AR3').Value or '').strip().lower() == 'provedor eqa'
        prov = ('IFERROR(INDEX(Analitos!$AR$4:$AR$43,MATCH($A{r},Analitos!$A$4:$A$43,0)),"CAP")'
                if tem_ar else 'eqProvedor')
        exc = ',Estat_Exclusoes' if any(n.Name == 'Estat_Exclusoes' for n in wb.Names) else ''
        cvi = coluna_cvi(wb)
        ucal = remover_ucal(wb)
        log(f'3. Analitos: campo u(cal) {"removido da coluna " + ucal if ucal else "ausente"}; CVI em {cvi}; provedor {"por analito (AR)" if tem_ar else "eqProvedor"}; '
            f'exclusoes {"sim" if exc else "nao"}')
        cl = bloco_estatistica(wb, ult, exc, prov.replace('{r}', '{r}'), cvi)
        log(f'4. Estatistica: incerteza em AF:{cl}, janela MU_Ini/MU_Fim, MU_Faixa')
        r0 = bloco_painel(wb, nlv)
        log(f'5. Painel: bloco de incerteza na linha {r0} (antes de MARGEM CRITICA)')
        nt, nf = acabamento_painel(wb, nlv)
        log(f'   acabamento: {nt} titulos de grafico no layout; Frequencia de CQ em {nf} linhas sem corte')

        log('6. fumaca')
        ex.xl.Calculation = -4105
        run = lambda m, *x: ex.run("'" + wb.Name + "'!" + m, *x, teto=600)
        run('mIncerteza.RecalcularIncerteza')
        ex.esperar()
        ok = 0
        ruins = []
        for r in range(LIN_TAB, ult + 1):
            an = e.Range(f'A{r}').Value
            if not an:
                continue
            urw, uc, U, st = (e.Range(f'AF{r}').Value, e.Range(f'AK{r}').Value, e.Range(f'AL{r}').Value,
                              str(e.Range(f'AO{r}').Value or ''))
            if st not in ('Ótimo', 'Desejável', 'Mínimo', 'Não atende', 'Sem meta', 'Insuficiente'):
                ruins.append((r, an, 'classe', st))
            if isinstance(urw, float) and urw > 0:
                if isinstance(uc, float):            # sem u(bias) do CEQ: uc e U ficam vazios (ADR-067)
                    if not (isinstance(U, float) and abs(U - 2 * uc) < 1e-9 and uc >= urw - 1e-12):
                        ruins.append((r, an, urw, uc, U))
                elif uc not in (None, ''):
                    ruins.append((r, an, 'uc', uc))
                ok += 1
            if not st:
                ruins.append((r, an, 'sem status'))
            for c in range(C0, C0 + 16):
                v = e.Cells(r, c).Value
                if isinstance(v, int) and v < 0 and v > -2147483648 + 100:
                    pass
            if any(isinstance(e.Cells(r, c).Value, int) and e.Cells(r, c).Text.startswith('#') for c in range(C0, C0 + 16)):
                ruins.append((r, an, 'erro de formula'))
        if ruins or ok == 0:
            raise SystemExit(f'fumaca da incerteza falhou: {ok} linhas com u(Rw); problemas {ruins[:6]}')
        log(f'   {ok} linhas analito|nivel com u(Rw); U = 2 uc conferido; classe valida; janela '
            f'{e.Range("AG11").Text} a {e.Range("AI11").Text}')
        txt = str(wb.Worksheets('Painel').Cells(r0 + 1, 1).Value)
        if not txt.startswith(('Janela', 'Sem dados')):
            raise SystemExit(f'nota do Painel: {txt!r}')

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
