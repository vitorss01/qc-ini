# -*- coding: utf-8 -*-
"""instalar_adr102.py -- ADR-102 (25/11/2026): alvos do laboratorio por lote e periodo (manual) e vigilancia contra a
referencia da bula/Unity, no rodape do Painel (linhas 97 a 183). Adendo 25/11: bloco da EXCECAO TEMPORARIA (valores
manuais, linhas 105-111, trancado ate o botao de liberar); referencia, observacoes e parecer desceram 8 linhas.

Uso:  python instalar_adr102.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]
      python instalar_adr102.py --autoteste      (estatico, sem Excel: layout daqui = constantes do mAlvosRef e do
                                                 Painel.cls, modulo no codigo_atual, etapa no entregar)

Idempotente (a 2a execucao registra 0 mudancas: o layout e refeito do zero, preservando o lote, as datas e o parecer
digitados; a grade e as observacoes vem das abas tecnicas). Reinstalar sobre o layout ANTERIOR ao adendo (A97:Z175, sem o
bloco da excecao) tambem funciona: le o parecer da posicao antiga (A168), confere que A176:Z183 esta livre, limpa a area
inteira (valores, formatos, mesclas, validacoes) e remonta; a referencia digitada esta na RefBulaStore (nao no Painel) e
volta pela fumaca (AtualizarTudo). O bloco da excecao sai sempre TRANCADO e vazio.
Passos: 1. todo o VBA atual das fontes (codigo_atual.py: mAlvosRef novo, Painel.cls com o Worksheet_Change da secao,
mIntegracao com a verificacao no fim do ATUALIZAR); EXIGE compilacao;
2. abas ocultas RefBulaStore, RefBulaObs, RefBulaParecer (cabecalho, formatos, protegidas);
3. LotesStore AA1:AO1 = auditoria por nivel (N_USADO, DATA_INI, DATA_FIM, CONDICAO, ATUALIZADO_EM);
4. Cfg_Integracao ALVO_N_RECOMENDADO = 20 e ALVO_N_IDEAL = 90 (so se faltarem: o valor ajustado fica);
5. Painel linhas 97-183 (aborta se a area tiver conteudo que nao seja desta secao, ou grafico/forma nela): titulos,
   cabecalhos, entradas destravadas (B98:C98 com a lista dos lotes, E98, G98, grade, parecer), bloco da excecao
   TRANCADO e cinza (B107:F1xx, B110:U110, lista Lote anterior/Bula em D107:D1xx), botoes "Aplicar Alvos ao Lote",
   "Liberar digitacao manual", "Gravar valores manuais" e "Registrar parecer"
   (mover com as celulas: o encaixe muda a altura das linhas de cima);
6. fumaca: mAlvosRef.AtualizarTudo, LimitesNTexto, travas, botoes;
7. protecao devolvida e Save com os eventos desligados (mtime conferido). Qualquer falha para ANTES de salvar.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))

MUDANCAS = []
TITULO = 'ALVOS DO LABORATÓRIO — apuração por lote e período'
R_TIT, R_ENT, R_CAB, R_NV1, R_REGRA = 97, 98, 99, 100, 103
# adendo 25/11: excecao temporaria (valores manuais) logo abaixo da apuracao
R_MAN_TIT, R_MAN_CAB, R_MAN1, R_MAN_JUST, R_MAN_BOT = 105, 106, 107, 110, 111
C_MAN_MED, C_MAN_DP, C_MAN_ORIG, C_MAN_LOTE, C_MAN_INFO, C_MAN_STATUS, C_MAN_JUST, C_MARCA = 2, 3, 4, 5, 7, 11, 2, 26
R_REF_TIT, R_REF_CAB, R_REF1, N_REF = 113, 114, 115, 40
R_OBS_TIT, R_OBS_CAB, R_OBS1, N_OBS, R_OBS_TOT = 156, 157, 158, 15, 173
R_PAR_TIT, R_PAR_TXT, R_PAR_TXT_FIM, R_PAR_BOT, R_PAR_ULT = 175, 176, 179, 180, 181
R_FIM, C_FIM_AREA = 183, 26                    # area conferida: A97:Z183
R_FIM_ANTIGO, R_PAR_TXT_ANTIGO = 175, 168      # layout antes do adendo (A97:Z175, parecer em A168)
TITULO_MAN = 'VALORES MANUAIS TEMPORÁRIOS (exceção)'
C_APURADO = 26                                 # Z98
BOTOES = {'adr102_aplicar': ('Aplicar Alvos ao Lote', 'mAlvosRef.AplicarAlvosAoLote', f'K{R_ENT}'),
          'adr102_liberar': ('Liberar digitação manual', 'mAlvosRef.LiberarDigitacaoManual',
                             f'A{R_MAN_BOT}'),
          'adr102_manual_gravar': ('Gravar valores manuais', 'mAlvosRef.GravarValoresManuais', f'K{R_MAN_BOT}'),
          'adr102_parecer': ('Registrar parecer', 'mAlvosRef.RegistrarParecer', f'A{R_PAR_BOT}')}
CFG = [('ALVO_N_RECOMENDADO', 20, 'ADR-102: N mínimo recomendado para o alvo do laboratório (abaixo: só TEMPORÁRIO, '
                                  'com ciência; < 2 bloqueia)'),
       ('ALVO_N_IDEAL', 90, 'ADR-102: N ideal para o alvo do laboratório (informativo)')]
AUD_NOMES = ['N_USADO_N', 'DATA_INI_N', 'DATA_FIM_N', 'CONDICAO_N', 'ATUALIZADO_EM_N']
LS_C_AUD1 = 27                                 # AA
ABAS = {'RefBulaStore': ['LOTE', 'ANALITO', 'CHAVE', 'N1_MEDIA', 'N1_LI', 'N1_LS', 'N2_MEDIA', 'N2_LI', 'N2_LS',
                         'N3_MEDIA', 'N3_LI', 'N3_LS', 'FONTE', 'REGISTRADO_EM', 'REGISTRADO_POR'],
        'RefBulaObs': ['CHAVE', 'ID_REGISTRO', 'LOTE', 'ANALITO', 'NIVEL', 'DATA_HORA', 'RESULTADO', 'LI', 'LS',
                       'SITUACAO', 'STATUS', 'CRIADA_EM', 'STATUS_EM', 'STATUS_POR', 'ID_PARECER', 'EQUIPAMENTO', 'RUN'],
        'RefBulaParecer': ['ID_PARECER', 'DATA_HORA', 'USUARIO', 'LOTE', 'N_OBSERVACOES', 'PARECER']}
# formatos por aba: colunas de texto ('@') e de data/hora
TEXTO = {'RefBulaStore': ['A', 'M'], 'RefBulaObs': ['B', 'C', 'Q'], 'RefBulaParecer': ['A', 'D', 'F']}  # M/F: texto livre ('-', '=' nao viram formula)
DATAHORA = {'RefBulaStore': ['N'], 'RefBulaObs': ['F', 'L', 'M'], 'RefBulaParecer': ['B']}
AZUL_ESCURO = 0x2C2813
BRANCO = 0xFFFFFF
AZUL = 0x7F3F1F
CINZA = 0x595959
AMARELO_CLARO = 0xCCF2FF                       # BGR de #FFF2CC (entrada)
CINZA_CLARO = 0xF2F2F2
CINZA_TRANCADO = 0xD9D9D9                      # #D9D9D9 = RGB(217, 217, 217) do mAlvosRef.EstadoManual
LISTA_ORIGEM_DP = ('Lote anterior', 'Bula')     # = ORIG_LOTE_ANT / ORIG_BULA do mAlvosRef
REGRA = ('Regra do N: < 2 bloqueia · 2 a 19 NÃO RECOMENDADO (grava só como TEMPORÁRIO, com ciência) · 20 a 89 '
         'aceitável · ≥ 90 ideal. Só grava no botão; o valor fica estático na LotesStore.')


def log(m):
    print(m, flush=True)


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def c_fonte(nlv):
    return 2 + 3 * nlv


def grade(nlv):
    return f'B{R_REF1}:{col(c_fonte(nlv))}{R_REF1 + N_REF - 1}'


def man_lote(nlv):
    """Lote anterior da excecao (E:F mescladas, um nivel por linha): o Intersect do Painel.cls."""
    return f'{col(C_MAN_LOTE)}{R_MAN1}:{col(C_MAN_LOTE + 1)}{R_MAN1 + nlv - 1}'


# =============================================================== autoteste estatico (sem Excel)
def autoteste():
    falhou = []

    def ok(cod, cond, ev=''):
        print(('PASS ' if cond else 'FAIL ') + cod + (f' -- {ev}' if ev else ''), flush=True)
        if not cond:
            falhou.append(cod)

    bas = open(os.path.join(AQUI, 'src', 'comum', 'mAlvosRef.bas'), encoding='utf-8').read()
    const = {m.group(1): int(m.group(2)) for m in re.finditer(r'Private Const (\w+) As Long = (\d+)', bas)}
    esperado = {'R_TIT': R_TIT, 'R_ENT': R_ENT, 'R_NV1': R_NV1, 'R_REGRA': R_REGRA, 'R_REF1': R_REF1, 'N_REF': N_REF,
                'R_OBS_TIT': R_OBS_TIT, 'R_OBS1': R_OBS1, 'N_OBS_TELA': N_OBS, 'R_OBS_TOT': R_OBS_TOT,
                'R_PAR_TXT': R_PAR_TXT, 'R_PAR_ULT': R_PAR_ULT, 'C_LOTE': 2, 'C_DE': 5, 'C_ATE': 7,
                'C_APURADO': C_APURADO, 'LS_C_AUD1': LS_C_AUD1, 'LS_AUD_N': len(AUD_NOMES),
                'R_MAN_TIT': R_MAN_TIT, 'R_MAN1': R_MAN1, 'R_MAN_JUST': R_MAN_JUST, 'C_MAN_MED': C_MAN_MED,
                'C_MAN_DP': C_MAN_DP, 'C_MAN_ORIG': C_MAN_ORIG, 'C_MAN_LOTE': C_MAN_LOTE, 'C_MAN_INFO': C_MAN_INFO,
                'C_MAN_STATUS': C_MAN_STATUS, 'C_MAN_JUST': C_MAN_JUST, 'C_MARCA': C_MARCA}
    dif = {k: (const.get(k), v) for k, v in esperado.items() if const.get(k) != v}
    ok('S01 layout do instalador = constantes do mAlvosRef.bas', not dif, dif)
    for nome, cab in ABAS.items():
        m = re.search(r'Case ABA_' + {'RefBulaStore': 'REF', 'RefBulaObs': 'OBS', 'RefBulaParecer': 'PAR'}[nome] +
                      r'\s+cab = Array\((.*?)\)\n', bas.replace('\r\n', '\n').replace(' _\n', ''), re.S)
        lidos = re.findall(r'"([^"]+)"', m.group(1)) if m else []
        ok(f'S02 cabecalho {nome} igual no VBA e no instalador', lidos == cab, lidos)
    for prod, nlv in (('bio', 2), ('hema', 3)):
        cls = open(os.path.join(AQUI, 'src', prod, 'Painel.cls'), encoding='utf-8').read()
        ok(f'S03 Painel.cls {prod}: Intersect da grade {grade(nlv)}, da entrada B98:C98,E98,G98, do lote anterior '
           f'{man_lote(nlv)} e do analito (B3:C3 tranca a excecao); nenhum endereco do layout antigo',
           f'Me.Range("{grade(nlv)}")' in cls and 'Me.Range("B98:C98,E98,G98")' in cls
           and 'mAlvosRef.SecaoMudou' in cls and 'mAlvosRef.GradeEditada Target' in cls
           and f'Me.Range("{man_lote(nlv)}")' in cls and 'mAlvosRef.ManualEditado Target' in cls
           and 'mAlvosRef.ManualContextoMudou' in cls and 'B107:H146' not in cls and 'B107:K146' not in cls)
        integ = open(os.path.join(AQUI, 'src', prod, 'mIntegracao.bas'), encoding='utf-8').read()
        chama = 'resRef = mAlvosRef.VerificarAoAtualizar()'
        fim_motor = 'mDados.EncerrarSnapshot\n    tempos'
        resumo = 'IIf(Len(resRef) > 0, vbCrLf & vbCrLf & resRef, "") & _\n           vbCrLf & vbCrLf & "Tempo: "'
        ok(f'S04 mIntegracao {prod}: verificacao no fim do motor (On Error local), texto antes do "Tempo:"',
           chama in integ and fim_motor in integ and ('On Error Resume Next\n    ' + chama) in integ
           and integ.index(chama) < integ.index(fim_motor) and resumo in integ)
    import codigo_atual
    ok('S05 codigo_atual: mAlvosRef entre os modulos NOVOS', 'mAlvosRef' in codigo_atual.NOVOS)
    import entregar
    src = open(os.path.join(AQUI, 'entregar.py'), encoding='utf-8').read()
    ok('S06 entregar: etapa adr102 logo depois da adr101; qa_alvos_manual nas suites do lote',
       "('adr102', ['instalar_adr102.py', produto, alvo])" in src
       and src.index("('adr101'") < src.index("('adr102'") < src.index("('blindagem'")
       and 'qa_alvos_manual.py' in entregar.SUITES_DO_LOTE and 'qa_alvos_manual.py' in entregar.SUITES)
    ok('S07 grade: Bio B115:H154, Hema B115:K154; lote anterior Bio E107:F108, Hema E107:F109',
       grade(2) == 'B115:H154' and grade(3) == 'B115:K154' and man_lote(2) == 'E107:F108' and man_lote(3) == 'E107:F109')
    # S08 layout sem sobreposicao: apuracao < excecao < referencia < observacoes < parecer <= fim
    seq = [R_TIT, R_ENT, R_CAB, R_NV1, R_NV1 + 2, R_REGRA, R_MAN_TIT, R_MAN_CAB, R_MAN1, R_MAN1 + 2, R_MAN_JUST,
           R_MAN_BOT, R_REF_TIT, R_REF_CAB, R_REF1, R_REF1 + N_REF - 1, R_OBS_TIT, R_OBS_CAB, R_OBS1,
           R_OBS1 + N_OBS - 1, R_OBS_TOT, R_PAR_TIT, R_PAR_TXT, R_PAR_TXT_FIM, R_PAR_BOT, R_PAR_ULT, R_FIM]
    ok('S08 layout em ordem e sem sobreposicao (97..183); o antigo (A97:Z175) cabe na area nova',
       all(a < b for a, b in zip(seq, seq[1:])) and R_FIM_ANTIGO < R_FIM and R_PAR_TXT_ANTIGO < R_PAR_TXT, seq)
    # S09 bula/Unity e complemento: a parte da referencia (vista, grade, verificacao, observacoes, parecer) NUNCA
    # grava na LotesStore. So AplicarAlvos e GravarManualCore chamam GravarParametroStore.
    i_ref = bas.index("' ============================ REFERENCIA (RefBulaStore)")
    i_ent = bas.index("' ============================ ENTRADA DA SECAO")
    parte_ref = bas[i_ref:i_ent]
    proibido = [x for x in ('ABA_LS', 'LS_C_', 'GravarParametroStore', '"LotesStore"', 'AplicarAlvos', 'GravarManualCore')
                if x in parte_ref]
    escritores = sorted(m.group(1) for m in re.finditer(
        r'^(?:Public |Private )?(?:Function|Sub) (\w+)\((?:(?!^End (?:Function|Sub)).)*?GravarParametroStore', bas,
        re.S | re.M))
    ok('S09 referencia bula/Unity nao grava na LotesStore (nenhum ABA_LS/LS_C_/GravarParametroStore na parte da '
       'referencia); GravarParametroStore so em AplicarAlvos e GravarManualCore',
       not proibido and escritores == ['AplicarAlvos', 'GravarManualCore'], {'proibido': proibido,
                                                                           'escritores': escritores})
    # S10 adendo: pergunta 1/2/3, origem com o modo, excecao com a origem e a condicao do ADR, aviso amarelo, Workbook
    pasta = open(os.path.join(AQUI, 'src', 'comum', 'EstaPastaDeTrabalho.cls'), encoding='utf-8').read()
    faltam = [x for x in ('modo = PerguntarModo(msg)', '" [" & ModoTxt(modo) & "]"', '"EXCECAO TEMPORARIA "',
                          '"EXCECAO_TEMPORARIA"', '"ALVO_MANUAL_LIBERADO"', '"PARAMETRO_LOTE_EXCECAO"',
                          'temporário: substituir pelos valores do laboratório',
                          f'ORIG_LOTE_ANT As String = "{LISTA_ORIGEM_DP[0]}"',
                          f'ORIG_BULA As String = "{LISTA_ORIGEM_DP[1]}"') if x not in bas]
    ok('S10 adendo no VBA: PerguntarModo (1/2/3), origem MANUAL ... [MODO], EXCECAO TEMPORARIA, EXCECAO_TEMPORARIA, '
       'ALVO_MANUAL_LIBERADO, aviso amarelo, Lote anterior/Bula = lista do instalador; Workbook_Open/BeforeClose trancam',
       not faltam and pasta.count('mAlvosRef.BloquearManual') == 2, {'faltam': faltam})
    # S11 marca temporaria POR CAMPO (revisao 25/11): CONDICAO_Nk gravado so via CondicaoDasMarcas (nunca a condicao
    # do N ou EXCECAO_TEMPORARIA direto), o amarelo tambem le AA:AO, a excecao sobre valor PROPRIO pede confirmacao e
    # o proprio lote nao vale como "Lote anterior"
    def corpo(nome):
        m = re.search(r'^(?:Public |Private )?(?:Function|Sub) ' + nome + r'\(.*?^End (?:Function|Sub)', bas, re.S | re.M)
        return m.group(0) if m else ''
    ap, gm, va = corpo('AplicarAlvos'), corpo('GravarManualCore'), corpo('AtualizarVistaAlvos')
    falta11 = [x for x, ok_ in (
        ('AplicarAlvos grava condNova', 'ls.Cells(rLs, k + 3).Value = condNova(nv)' in ap
         and 'k + 3).Value = cond(nv)' not in ap and 'CondicaoDasMarcas(' in ap),
        ('GravarManualCore grava CondicaoDasMarcas', 'c + 3).Value = CondicaoDasMarcas(' in gm
         and 'c + 3).Value = COND_EXCECAO' not in gm),
        ('GravarManualCore: 2a confirmacao do PROPRIO', '"CANCELADO|proprio"' in gm),
        ('GravarManualCore: proprio lote recusado', 'StrComp(la, lote, vbTextCompare) = 0' in gm),
        ('AtualizarVistaAlvos le a marca do AA:AO', 'MarcasDaCondicao' in va and 'LS_C_AUD1' in va)) if not ok_]
    ok('S11 marca temporaria por campo no CONDICAO_Nk (Aplicar parcial e excecao parcial nao apagam a marca do campo '
       'mantido); amarelo tambem pelo AA:AO; excecao sobre valor PROPRIO pede 2a confirmacao; proprio lote recusado',
       not falta11, {'faltam': falta11})
    total = 11 + len(ABAS) + 2
    print(f'AUTOTESTE: {total - len(falhou)}/{total} PASS' + (f' -- FAIL {falhou}' if falhou else ''), flush=True)
    return 1 if falhou else 0


# =============================================================== instalacao (Excel)
def abas_tecnicas(wb, cd, ii):
    nomes = [w.Name for w in wb.Worksheets]
    for nome, cab in ABAS.items():
        if nome not in nomes:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = nome
            MUDANCAS.append(f'aba {nome}')
        ws = wb.Worksheets(nome)
        ws.Unprotect(ii.SENHA)
        atual = [str(c or '') for c in ws.Range(ws.Cells(1, 1), ws.Cells(1, len(cab))).Value[0]]
        if atual != cab:
            if any(atual) and atual != cab:
                raise SystemExit(f'{nome} com cabecalho diferente ({atual}) -- nada foi salvo')
            ws.Range(ws.Cells(1, 1), ws.Cells(1, len(cab))).Value = [cab]
            ws.Rows(1).Font.Bold = True
            MUDANCAS.append(f'{nome} cabecalho')
        for c in TEXTO[nome]:
            if str(ws.Range(f'{c}2').NumberFormat) != '@':
                ws.Range(f'{c}2:{c}20000').NumberFormat = '@'   # so o corpo: o cabecalho fica General
        for c in DATAHORA[nome]:
            if 'hh' not in str(ws.Range(f'{c}2').NumberFormat):
                cd.fmt_data(ws.Range(f'{c}2:{c}20000'), 'dd/mm/yyyy hh:mm:ss')
        if ws.Visible != 0:
            ws.Visible = 0                                     # xlSheetHidden
        ws.Protect(ii.SENHA, False, True, True, True)          # o mesmo de mSeguranca.RestaurarProtecao


def lotesstore(wb, cd):
    ls = wb.Worksheets('LotesStore')
    esperado = [f'{n}{k}' for k in (1, 2, 3) for n in AUD_NOMES]
    rng = ls.Range(ls.Cells(1, LS_C_AUD1), ls.Cells(1, LS_C_AUD1 + len(esperado) - 1))
    atual = [str(c or '') for c in rng.Value[0]]
    if atual == esperado:
        return
    if any(atual):
        raise SystemExit(f'LotesStore AA1:AO1 com outro conteudo ({atual[:4]}) -- nada foi salvo')
    corpo = ls.Range(ls.Cells(2, LS_C_AUD1), ls.Cells(4001, LS_C_AUD1 + len(esperado) - 1))
    if int(wb.Application.WorksheetFunction.CountA(corpo)):
        raise SystemExit('LotesStore AA2:AO4001 tem dados sem cabecalho -- nada foi salvo')
    rng.Value = [esperado]
    rng.Font.Bold = True
    for k in range(3):
        base = LS_C_AUD1 + 5 * k
        cd.fmt_data(ls.Range(ls.Cells(2, base + 1), ls.Cells(4001, base + 2)), 'dd/mm/yyyy')
        cd.fmt_data(ls.Range(ls.Cells(2, base + 4), ls.Cells(4001, base + 4)), 'dd/mm/yyyy hh:mm:ss')
    MUDANCAS.append('LotesStore AA1:AO1 auditoria por nivel')


def cfg(wb):
    lo = wb.Worksheets('Cfg_Integracao').ListObjects('tblConfigIntegracao')
    ks = [str(r[0] or '').strip().upper() for r in lo.DataBodyRange.Value]
    for chave, valor, desc in CFG:
        if chave in ks:
            continue
        lin = lo.ListRows.Add()
        lin.Range.Cells(1, 1).Value = chave
        lin.Range.Cells(1, 2).Value = valor
        if lo.ListColumns.Count >= 3:
            lin.Range.Cells(1, 3).Value = desc
        MUDANCAS.append(f'Cfg {chave}={valor}')


def mesclar(ws, ende):
    r = ws.Range(ende)
    if not (r.MergeCells is True and str(r.MergeArea.Address) == str(r.Address)):
        r.UnMerge()
        r.Merge()
    return r


def titulo(ws, ende, texto, tam=11):
    r = mesclar(ws, ende)
    r.Cells(1, 1).Value = texto
    r.Font.Bold = True
    r.Font.Size = tam
    r.Font.Color = BRANCO
    r.Interior.Color = AZUL_ESCURO
    r.HorizontalAlignment = -4131                      # esquerda
    r.VerticalAlignment = -4108


def cabecalho(ws, linha, textos, c0=1):
    for i, t in enumerate(textos):
        c = ws.Cells(linha, c0 + i)
        c.Value = t
        c.Font.Bold = True
        c.Font.Size = 9
        c.Font.Color = AZUL
        c.Interior.Color = CINZA_CLARO
        c.WrapText = True
        c.HorizontalAlignment = -4108
        c.Borders(9).LineStyle = 1                     # xlEdgeBottom


def entrada(r):
    r.Locked = False
    r.Interior.Color = AMARELO_CLARO
    r.Borders.LineStyle = 1
    r.Borders.Color = 0xBFBFBF


def botao(ws, nome_b, texto, celula, macro):
    # Padrao do Painel (pedido do usuario 25/11): mesmo tamanho, cor e fonte do botao existente "btnGrafTodos"
    # (ADR-063, que por sua vez copia o btnHist); retangulo arredondado, texto centralizado nos dois eixos.
    nomes = [str(s.Name) for s in ws.Shapes]
    modelo = None
    for nm in ('btnGrafTodos', 'btnHist'):
        if nm in nomes:
            modelo = ws.Shapes(nm)
            break
    if modelo is not None:
        larg, alt = float(modelo.Width), float(modelo.Height)
        fundo = modelo.Fill.ForeColor.RGB
        mtr = modelo.TextFrame2.TextRange.Font
        cor_txt, tam, fonte = mtr.Fill.ForeColor.RGB, float(mtr.Size), str(mtr.Name)
    else:
        larg, alt, fundo, cor_txt, tam, fonte = 150.0, 24.0, 0x6B6F1F, BRANCO, 9.0, 'Segoe UI'
    c = ws.Range(celula)
    if float(c.RowHeight) < alt + 4:                   # a linha do botao comporta a altura padrao
        c.EntireRow.RowHeight = alt + 4
    if nome_b in nomes:
        b = ws.Shapes(nome_b)
    else:
        b = ws.Shapes.AddShape(5, c.Left + 2, c.Top + 1, larg, alt)
        b.Name = nome_b
        MUDANCAS.append(f'Painel botao {nome_b}')
    tf = b.TextFrame2
    tf.AutoSize = 0                                    # tamanho fixo, igual ao modelo
    b.Width, b.Height = larg, alt
    b.Left = c.Left + 2
    b.Top = c.Top + max(0.0, (float(c.Height) - alt) / 2)
    if str(tf.TextRange.Text) != texto:
        tf.TextRange.Text = texto
    if not str(b.OnAction).endswith(macro):
        b.OnAction = macro
    try:
        b.Adjustments.SetItem(1, 0.25)
    except Exception:
        pass
    tf.WordWrap = -1
    tf.MarginLeft = tf.MarginRight = 2
    tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3                              # meio
    tf.HorizontalAnchor = 2                            # centro
    tr = tf.TextRange
    tr.ParagraphFormat.Alignment = 2                   # centro
    tr.Font.Size = tam
    tr.Font.Bold = -1
    tr.Font.Name = fonte
    tr.Font.Fill.ForeColor.RGB = cor_txt
    b.Fill.ForeColor.RGB = fundo
    b.Line.Visible = 0
    b.Placement = 2                                    # xlMove: o encaixe muda a altura das linhas 9..38


def painel(wb, nlv, cd):
    p = wb.Worksheets('Painel')
    area = p.Range(p.Cells(R_TIT, 1), p.Cells(R_FIM, C_FIM_AREA))
    reinstalar = str(p.Cells(R_TIT, 1).Value or '') == TITULO
    guardado = {}
    if reinstalar:
        # layout atual (bloco da excecao em A105) ou o anterior ao adendo (parecer em A168, area A97:Z175)
        novo = str(p.Cells(R_MAN_TIT, 1).Value or '').startswith(TITULO_MAN)
        for e in (f'B{R_ENT}', f'E{R_ENT}', f'G{R_ENT}'):
            guardado[e] = p.Range(e).Value
        guardado[f'A{R_PAR_TXT}'] = p.Range(f'A{R_PAR_TXT if novo else R_PAR_TXT_ANTIGO}').Value
        if not novo:
            extra = p.Range(p.Cells(R_FIM_ANTIGO + 1, 1), p.Cells(R_FIM, C_FIM_AREA))
            ocup = [c.Address for c in extra.Cells if c.Formula not in (None, '')]
            if ocup:
                raise SystemExit(f'Painel A{R_FIM_ANTIGO + 1}:Z{R_FIM} ocupada ({ocup[:5]}): o layout do adendo nao '
                                 'cabe -- nada foi salvo')
            MUDANCAS.append(f'Painel: layout anterior ao adendo (A{R_TIT}:Z{R_FIM_ANTIGO}) refeito com a excecao')
    else:
        ocup = [c.Address for c in area.Cells if c.Formula not in (None, '')]
        if ocup:
            raise SystemExit(f'Painel A{R_TIT}:Z{R_FIM} ocupada ({ocup[:5]}) -- nada foi salvo')
    for s in p.Shapes:
        if str(s.Name) in BOTOES:
            continue
        try:
            de, ate = s.TopLeftCell.Row, s.BottomRightCell.Row
        except Exception:                               # noqa: BLE001 -- comentario de celula etc.
            continue
        if ate >= R_TIT and de <= R_FIM:
            raise SystemExit(f'Painel: forma/grafico {s.Name} nas linhas {R_TIT}-{R_FIM} -- nada foi salvo')
    if reinstalar:
        area.UnMerge()
        area.Clear()
        area.Validation.Delete()                       # listas antigas (B98, D107..): o Clear nao garante tirar
        p.Range(f'A{R_TIT}:A{R_FIM}').EntireRow.RowHeight = p.StandardHeight   # as linhas mudaram de papel
    area.Locked = True
    area.Font.Size = 9
    area.Font.Name = 'Segoe UI'

    # ---- 1. alvos do laboratorio ----
    titulo(p, f'A{R_TIT}:U{R_TIT}', TITULO, 12)
    p.Rows(R_TIT).RowHeight = 22
    for c, t in (('A', 'Lote'), ('D', 'De'), ('F', 'Até'), ('H', 'Analito')):
        p.Range(f'{c}{R_ENT}').Value = t
        p.Range(f'{c}{R_ENT}').Font.Bold = True
        p.Range(f'{c}{R_ENT}').HorizontalAlignment = -4152      # direita
    lote = mesclar(p, f'B{R_ENT}:C{R_ENT}')
    lote.NumberFormat = '@'
    entrada(lote)
    nome_lista = '=lstLotes'
    try:
        wb.Names('lstLotes').RefersToRange
    except Exception:                                   # noqa: BLE001
        nome_lista = '=regLoteCol'
    lote.Validation.Delete()
    lote.Validation.Add(3, 1, 1, nome_lista)             # lista, Parar
    lote.Validation.IgnoreBlank = True
    lote.Validation.InCellDropdown = True
    lote.Validation.InputTitle = 'Lote da apuração'
    lote.Validation.InputMessage = 'Obrigatório. Os alvos são gravados SÓ neste lote; a referência abaixo é deste lote.'
    lote.Validation.ErrorTitle = 'Lote não cadastrado'
    lote.Validation.ErrorMessage = 'Escolha um lote da lista (cadastro da aba Configuração).'
    for e in (f'E{R_ENT}', f'G{R_ENT}'):
        r = p.Range(e)
        cd.fmt_data(r, 'dd/mm/yyyy')
        entrada(r)
    an = mesclar(p, f'I{R_ENT}:J{R_ENT}')
    an.Formula = '=selAnalito'
    an.Font.Bold = True
    an.Interior.Color = CINZA_CLARO
    p.Cells(R_ENT, C_APURADO).Font.Color = 0xA6A6A6
    p.Cells(R_ENT, C_APURADO).Font.Size = 7
    p.Rows(R_ENT).RowHeight = 22
    cab = ['Nível', 'N', 'Média calc.', 'DP calc.', 'CV%', 'Condição do N', 'Média atual', 'DP atual', 'Origem atual']
    cabecalho(p, R_CAB, cab)
    mesclar(p, f'I{R_CAB}:U{R_CAB}')
    for r in range(R_NV1, R_NV1 + nlv):
        mesclar(p, f'I{r}:U{r}').HorizontalAlignment = -4131
        p.Range(f'A{r}').Value = f'N{r - R_NV1 + 1}'
        p.Range(f'A{r}').Font.Bold = True
        p.Range(f'B{r}:I{r}').Borders(9).LineStyle = 1
        p.Range(f'B{r}:I{r}').Borders(9).Color = 0xD9D9D9
    rg = mesclar(p, f'A{R_REGRA}:U{R_REGRA}')
    rg.Value = REGRA
    rg.WrapText = True
    rg.Font.Color = CINZA
    rg.Font.Italic = True
    rg.VerticalAlignment = -4160
    p.Rows(R_REGRA).RowHeight = 36

    # ---- 1b. excecao temporaria: valores manuais (adendo 25/11) -- TRANCADO e cinza ate o botao de liberar ----
    titulo(p, f'A{R_MAN_TIT}:J{R_MAN_TIT}', f'{TITULO_MAN} — Média da bula e DP do lote anterior; só no início do lote')
    stt = mesclar(p, f'{col(C_MAN_STATUS)}{R_MAN_TIT}:U{R_MAN_TIT}')
    stt.Cells(1, 1).Value = ('TRANCADA (padrão). Exceção: só no início do lote, sem valores próprios -- botão '
                             '"Liberar digitação manual".')
    stt.Font.Bold = True
    stt.Font.Color = 0x404040
    stt.Interior.Color = CINZA_TRANCADO
    stt.ShrinkToFit = True
    p.Cells(R_MAN_TIT, C_MARCA).Font.Color = 0xA6A6A6
    p.Cells(R_MAN_TIT, C_MARCA).Font.Size = 7
    p.Rows(R_MAN_TIT).RowHeight = 20
    cabecalho(p, R_MAN_CAB, ['Nível', 'Média (bula)', 'DP', 'Origem do DP', 'Lote anterior'])
    mesclar(p, f'{col(C_MAN_LOTE)}{R_MAN_CAB}:{col(C_MAN_LOTE + 1)}{R_MAN_CAB}')
    cabecalho(p, R_MAN_CAB, ['Informação (o DP do lote anterior vem da LotesStore; pode corrigir)'], c0=C_MAN_INFO)
    mesclar(p, f'{col(C_MAN_INFO)}{R_MAN_CAB}:U{R_MAN_CAB}')
    p.Rows(R_MAN_CAB).RowHeight = 26
    sep = cd.sep_lista(wb.Application)                 # ADR-095: lista com o separador do Excel
    for k in range(1, nlv + 1):
        r = R_MAN1 + k - 1
        p.Range(f'A{r}').Value = f'N{k}'
        p.Range(f'A{r}').Font.Bold = True
        mesclar(p, f'{col(C_MAN_LOTE)}{r}:{col(C_MAN_LOTE + 1)}{r}').NumberFormat = '@'     # "010" fica "010"
        inf = mesclar(p, f'{col(C_MAN_INFO)}{r}:U{r}')
        inf.Font.Color = CINZA
        inf.ShrinkToFit = True
        for c in (C_MAN_MED, C_MAN_DP):
            p.Cells(r, c).NumberFormat = '0.0###'
        o = p.Cells(r, C_MAN_ORIG)
        o.Validation.Delete()
        o.Validation.Add(3, 1, 1, sep.join(LISTA_ORIGEM_DP))          # lista, Parar
        o.Validation.IgnoreBlank = True
        o.Validation.InCellDropdown = True
        o.Validation.ErrorTitle = 'Origem do DP'
        o.Validation.ErrorMessage = 'Escolha Lote anterior ou Bula.'
    p.Range(f'A{R_MAN_JUST}').Value = 'Justificativa'
    p.Range(f'A{R_MAN_JUST}').Font.Bold = True
    js = mesclar(p, f'{col(C_MAN_JUST)}{R_MAN_JUST}:U{R_MAN_JUST}')
    js.NumberFormat = '@'                              # justificativa iniciada por '-' ou '=' nao vira formula
    js.WrapText = True
    js.VerticalAlignment = -4160
    p.Rows(R_MAN_JUST).RowHeight = 30
    p.Rows(R_MAN_BOT).RowHeight = 24
    for end_ in (f'{col(C_MAN_MED)}{R_MAN1}:{col(C_MAN_LOTE + 1)}{R_MAN1 + nlv - 1}',
                 f'{col(C_MAN_JUST)}{R_MAN_JUST}:U{R_MAN_JUST}'):   # uma area por vez (uniao por texto falha via COM)
        man = p.Range(end_)
        man.Locked = True                              # TRANCADO por padrao (so o botao de liberar destrava)
        man.Interior.Color = CINZA_TRANCADO
        man.Borders.LineStyle = 1
        man.Borders.Color = 0xBFBFBF

    # ---- 2. referencia da bula / Unity ----
    titulo(p, f'A{R_REF_TIT}:U{R_REF_TIT}', 'VALORES DE REFERÊNCIA (bula / Unity) — lote da seção (B98). Digite Média, '
           'LI e LS por nível; resultado ativo fora de LI–LS gera observação (informativo)')
    cab = ['Analito']
    for k in range(1, nlv + 1):
        cab += [f'N{k} Média', f'N{k} LI', f'N{k} LS']
    cab += ['Fonte (bula / Unity)', 'Fora da faixa (lote)']
    cabecalho(p, R_REF_CAB, cab)
    p.Rows(R_REF_CAB).RowHeight = 26
    gr = p.Range(grade(nlv))
    entrada(gr)
    p.Range(p.Cells(R_REF1, c_fonte(nlv)), p.Cells(R_REF1 + N_REF - 1, c_fonte(nlv))).NumberFormat = '@'  # texto livre
    p.Range(p.Cells(R_REF1, 1), p.Cells(R_REF1 + N_REF - 1, 1)).Font.Bold = True
    fora = p.Range(p.Cells(R_REF1, c_fonte(nlv) + 1), p.Cells(R_REF1 + N_REF - 1, c_fonte(nlv) + 1))
    fora.Font.Color = 0x06009C
    fora.ShrinkToFit = True

    # ---- 3. observacoes automaticas ----
    titulo(p, f'A{R_OBS_TIT}:F{R_OBS_TIT}', 'OBSERVAÇÕES AUTOMÁTICAS (referência da bula / Unity)')
    st = mesclar(p, f'G{R_OBS_TIT}:U{R_OBS_TIT}')
    st.Font.Bold = True
    p.Rows(R_OBS_TIT).RowHeight = 20
    cabecalho(p, R_OBS_CAB, ['Data/hora', 'Analito', 'Nível', 'Resultado', 'LI', 'LS', 'Situação', 'Equipamento',
                             'ID do resultado', 'Observação desde'])
    tot = mesclar(p, f'A{R_OBS_TOT}:U{R_OBS_TOT}')
    tot.WrapText = True
    tot.Font.Italic = True
    tot.Font.Color = CINZA
    p.Rows(R_OBS_TOT).RowHeight = 26

    # ---- 4. parecer tecnico ----
    titulo(p, f'A{R_PAR_TIT}:U{R_PAR_TIT}', 'PARECER TÉCNICO DO ANALISTA — razão da ciência e conduta (registrado com '
           'usuário e data/hora; dá ciência às observações pendentes do lote)')
    tx = mesclar(p, f'A{R_PAR_TXT}:U{R_PAR_TXT_FIM}')
    entrada(tx)
    tx.NumberFormat = '@'                      # parecer iniciado por '-' ou '=' nao vira formula
    tx.WrapText = True
    tx.VerticalAlignment = -4160
    tx.HorizontalAlignment = -4131
    for r in range(R_PAR_TXT, R_PAR_TXT_FIM + 1):
        p.Rows(r).RowHeight = 18
    p.Rows(R_PAR_BOT).RowHeight = 24
    ul = mesclar(p, f'A{R_PAR_ULT}:U{R_PAR_ULT}')
    ul.WrapText = True
    ul.Font.Color = CINZA
    p.Rows(R_PAR_ULT).RowHeight = 30
    for nome_b, (texto, macro, cel) in BOTOES.items():
        botao(p, nome_b, texto, cel, macro)
    for e, v in guardado.items():
        if v not in (None, ''):
            p.Range(e).Value = v
    if not reinstalar:                                 # reinstalar refaz o layout do zero (sem mudanca de conteudo)
        MUDANCAS.append(f'Painel A{R_TIT}:U{R_FIM}')


def fumaca(ex, wb, nlv):
    nome = wb.Name
    p = wb.Worksheets('Painel')
    if not ex.run(f"'{nome}'!mAlvosRef.SecaoInstalada", teto=60):
        raise SystemExit('fumaca: mAlvosRef.SecaoInstalada = False -- nada foi salvo')
    lim = str(ex.run(f"'{nome}'!mAlvosRef.LimitesNTexto", teto=60))
    if not re.fullmatch(r'\d+\|\d+', lim):
        raise SystemExit(f'fumaca: LimitesNTexto = {lim!r} -- nada foi salvo')
    r = str(ex.run(f"'{nome}'!mAlvosRef.AtualizarTudo", teto=600))
    if not r.startswith('OK|OK|'):
        raise SystemExit(f'fumaca: AtualizarTudo = {r[:300]!r} -- nada foi salvo')
    travas = {f'B{R_ENT}': False, f'E{R_ENT}': False, f'G{R_ENT}': False, f'B{R_REF1}': False,
              f'{col(c_fonte(nlv))}{R_REF1 + N_REF - 1}': False, f'A{R_PAR_TXT}': False,
              f'A{R_TIT}': True, f'I{R_ENT}': True, f'A{R_REF1}': True, f'{col(c_fonte(nlv) + 1)}{R_REF1}': True,
              f'A{R_OBS1}': True, f'A{R_PAR_ULT}': True,
              # adendo: a excecao sai TRANCADA (so o botao de liberar destrava)
              f'B{R_MAN1}': True, f'C{R_MAN1}': True, f'D{R_MAN1}': True, f'E{R_MAN1}': True,
              f'B{R_MAN1 + nlv - 1}': True, f'B{R_MAN_JUST}': True}
    errados = [e for e, v in travas.items() if bool(p.Range(e).Locked) != v]
    if errados:
        raise SystemExit(f'fumaca: travas erradas em {errados} -- nada foi salvo')
    for nome_b, (texto, macro, _) in BOTOES.items():
        b = p.Shapes(nome_b)
        if not str(b.OnAction).endswith(macro) or int(b.Placement) != 2:
            raise SystemExit(f'fumaca: botao {nome_b} sem a macro {macro} ou sem mover com as celulas -- nada foi salvo')
    nomes_ab = [w.Name for w in wb.Worksheets]
    if any(a not in nomes_ab for a in ABAS):
        raise SystemExit('fumaca: abas tecnicas do ADR-102 ausentes -- nada foi salvo')
    if str(ex.run(f"'{nome}'!mAlvosRef.EstadoManualTexto", teto=60)) != 'TRANCADA' or \
            str(p.Cells(R_MAN_TIT, C_MARCA).Value or ''):
        raise SystemExit('fumaca: a digitacao manual (excecao) nao saiu TRANCADA -- nada foi salvo')
    log(f'   ok: limites {lim}; {r[:80]}; I98 = {p.Range(f"I{R_ENT}").Text!r}; excecao TRANCADA')


def main(produto, caminho, salvar=True):
    import xlh  # noqa: E402
    import instalar_integracao as ii  # noqa: E402
    import codigo_atual  # noqa: E402
    import camada_dados as cd  # noqa: E402
    from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                   # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    nlv = 2 if produto.startswith('Bio') else 3
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.EnableEvents = False
        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)
        ex.xl.Calculation = -4135
        log('2. abas tecnicas RefBulaStore / RefBulaObs / RefBulaParecer')
        abas_tecnicas(wb, cd, ii)
        log('3. LotesStore AA:AO (auditoria por nivel)')
        lotesstore(wb, cd)
        log('4. Cfg_Integracao ALVO_N_RECOMENDADO / ALVO_N_IDEAL')
        cfg(wb)
        log(f'5. Painel linhas {R_TIT}-{R_FIM}')
        painel(wb, nlv, cd)
        ex.xl.Calculation = -4105
        log('6. fumaca')
        ex.xl.EnableEvents = False
        fumaca(ex, wb, nlv)
        ex.xl.EnableEvents = False
        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))
        if salvar:
            antes = os.path.getmtime(wb.FullName)
            wb.Save()
            if os.path.getmtime(wb.FullName) == antes:
                raise SystemExit('o Save nao chegou ao arquivo (data de gravacao igual) -- conferir')
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    if a[:1] == ['--autoteste']:
        sys.exit(autoteste())
    if len(a) < 2:
        raise SystemExit(__doc__)
    main(a[0], a[1], salvar='--sem-salvar' not in a)
