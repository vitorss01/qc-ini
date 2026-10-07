# -*- coding: utf-8 -*-
"""camada_dados.py -- monta, numa pasta QC_INI aberta, a camada de dados nova:

    DB_SEAC.xlsm --(Power Query)--> DB_RECEBIMENTO -> DB_ORGANIZADO
                                         |
           DB_MANUAL + INATIVACAO + COMENTARIOS_TECNICOS
                                         v
                                    DB_CQ_FINAL --> QA_INTEGRACAO

Tudo idempotente: aba que existe e reaproveitada, tabela de ENTRADA que existe
NUNCA e recriada (preserva o que o usuario digitou), consulta e regravada com o
M de pq/*.m, tabela de saida e recarregada.

Usado pelo instalador (instalar_integracao.py) e pelos testes.
"""
import os
import re
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import pqlib  # noqa: E402
import tema  # noqa: E402
import config_produtos as cp  # noqa: E402

SENHA = 'qcini2025'
LINHA_CAB = 6           # cabecalho de toda tabela das abas novas
FOLGA_ENTRADA = {'tblResultados_Manuais': 300, 'tblInativacao_NaoConformes': 500, 'tblComentariosTecnicos': 500}

# ---------------------------------------------------------------- abas (ordem do fluxo)
ABAS = [
    ('DB',
     'DB — banco de dados recebido do DB_SEAC',
     'Camada 1 · ingestão · somente leitura · acumula, nunca apaga · uma linha por resultado',
     ['Colunas e nomes iguais aos da aba DADOS_ do DB_SEAC, mais ID_REGISTRO (chave estável = prefixo + id do SIPEC) '
      'e RECEBIDO_EM (primeira chegada).',
      'Atualizada pelo botão ⟳ ATUALIZAR DADOS. Resultado que sai da janela de 12 meses do DB_SEAC continua aqui.']),
    ('DB_ORGANIZADO',
     'DB_ORGANIZADO — grade horizontal de conferência',
     'Visão como as tabelas dinâmicas do DB_SEAC: uma linha por instante/nível, um analito por coluna',
     ['Só conferência: o Levey-Jennings e a Estatística NÃO leem daqui (leem a aba Principal - Resultados).',
      'RESULTADOS_AGREGADOS > 0: mais de um resultado do mesmo analito no mesmo instante (a célula mostra a média).']),
    ('Digitar Resultados',
     'DIGITAR RESULTADOS — continuidade quando o interfaceamento falha',
     'Entrada paralela (staging) · entra na DB_CQ_FINAL pelo mesmo caminho dos resultados automáticos',
     ['1) Preencha DATA, HORA, LOTE, NÍVEL, ANALITO e RESULTADO na próxima linha livre — o ID (MAN_0001, MAN_0002…) '
      'é gerado sozinho.  2) Clique ⟳ ATUALIZAR DADOS.',
      '3) Confira STATUS: ATIVO = entrou na estatística · CONFLITO_MANUAL = o mesmo resultado já veio pelo '
      'interfaceamento (o manual fica guardado, fora do cálculo) · MANUAL_INCOMPLETO = falta campo.']),
    ('Inativar',
     'INATIVAR — retirar resultados da estatística',
     'O resultado continua na base (auditoria) e sai do cálculo · REGISTRAR - LJ marcado = X vermelho no gráfico · '
     'desmarcado = não aparece',
     ['Digite o ID (o número que aparece ao passar o mouse no ponto do Levey-Jennings), escolha o analito, escreva o '
      'motivo e clique ⟳ ATUALIZAR DADOS.',
      'Para reativar, apague a linha.  Se o ID não for do analito escolhido, nada é inativado e a aba QA_INTEGRACAO avisa.']),
    ('COMENTARIOS_TECNICOS',
     'COMENTÁRIOS TÉCNICOS — observações por resultado',
     'Relacionados pelo ID_REGISTRO · a justificativa da inativação pode ficar aqui ou no MOTIVO da aba Inativar',
     ['Uma linha por comentário. Mais de um comentário para o mesmo ID aparece junto na DB_CQ_FINAL (separados por " | ").',
      'Comentário de resultado reativado fica como histórico (o QA informa, não acusa erro).']),
    ('Principal - Resultados',
     'PRINCIPAL - RESULTADOS — fonte oficial dos resultados do CQ (DB_CQ_FINAL)',
     'Single source of truth do Levey-Jennings, Westgard, Estatística e Power BI · gerada só pelo Power Query',
     ['Uma linha = um resultado. Inativados CONTINUAM aqui (PARTICIPA_ESTATISTICA = NÃO). '
      'TIPO_PLOTAGEM_LJ: NORMAL = ponto · X_VERMELHO = X no gráfico · NAO_PLOTAR = só auditoria.',
      'Para inativar: digite o número do ID e o analito na aba Inativar.  Não edite esta tabela — '
      'ela é refeita a cada ⟳ ATUALIZAR DADOS.']),
    ('QA_INTEGRACAO',
     'QA DA INTEGRAÇÃO — governança dos dados',
     'ERRO = processo inadequado (agir) · ALERTA = o sistema protegeu o dado, alguém deve olhar · INFO = registro',
     ['Refeito a cada ⟳ ATUALIZAR DADOS. Lê o que a DB_CQ_FINAL decidiu; não refaz regra.',
      '']),
    ('Cfg_Integracao',
     'CONFIGURAÇÃO DA INTEGRAÇÃO DB_SEAC',
     'Parâmetros (CHAVE | VALOR) e de/para de analitos · uso do administrador',
     ['Mudar o caminho do DB_SEAC, o prefixo do ID ou as tolerâncias é editar uma célula: nenhuma consulta muda.',
      'De/para: ANALITO_ORIGEM (nome do equipamento) → ANALITO_QCINI (nome do cadastro). Vazio = recebido, '
      'guardado, fora do Painel.']),
]
SAIDAS = {   # consulta -> (aba, tabela)
    'DB_RECEBIMENTO': ('DB', 'tblDB_Recebimento'),
    'DB_ORGANIZADO': ('DB_ORGANIZADO', 'tblDB_Organizado'),
    'DB_CQ_FINAL': ('Principal - Resultados', 'tblCQ_Final'),
    'QA_INTEGRACAO': ('QA_INTEGRACAO', 'tblQA_Integracao'),
}
ORDEM_CARGA = ['DB_RECEBIMENTO', 'DB_ORGANIZADO', 'DB_CQ_FINAL', 'QA_INTEGRACAO']
CONSULTAS = ['CFG_INTEGRACAO', 'SEAC_ORIGEM'] + ORDEM_CARGA

# ---------------------------------------------------------------- entradas
# (cabecalho, largura em caracteres, formato, tipo) -- tipo: 'in' digitado, 'auto' carimbado pelo VBA,
# 'f' formula de conferencia (calculada, bloqueada).
# Coluna de FORMULA nunca tem formato '@' (texto): com '@' o Excel guarda a formula como TEXTO literal e a
# celula mostra "=IF(..." em vez do valor (defeito achado em 06/10/2026 nas tres abas de entrada; ADR-070).
LOOKUP = '=IF([@ID_REGISTRO]="","",IFERROR(XLOOKUP([@ID_REGISTRO],tblCQ_Final[ID_REGISTRO],tblCQ_Final[{c}]),"—"))'
ENTRADAS = {
    'tblResultados_Manuais': ('Digitar Resultados', [
        ('ID_REGISTRO', 13, '@', 'auto'),
        ('DATA', 11, 'dd/mm/yyyy', 'in'),
        ('HORA', 8, 'hh:mm', 'in'),
        ('EQUIPAMENTO', 14, '@', 'in'),
        ('MATRIZ', 13, '@', 'in'),
        ('LOTE', 10, '@', 'in'),
        ('NIVEL', 7, '0', 'in'),
        ('ANALITO', 26, '@', 'in'),
        ('RESULTADO', 11, 'General', 'in'),
        ('UNIDADE', 9, '@', 'in'),
        ('MOTIVO', 28, '@', 'in'),
        ('USUARIO', 14, '@', 'auto'),
        ('REGISTRADO_EM', 16, 'dd/mm/yyyy hh:mm', 'auto'),
        ('RUN', 11, '0', LOOKUP.format(c='RUN')),
        ('STATUS', 20, 'General', LOOKUP.format(c='STATUS_ANALITICO')),
        ('DETALHE', 50, 'General', '=IF([@ID_REGISTRO]="","",IFERROR(XLOOKUP([@ID_REGISTRO],tblCQ_Final[ID_REGISTRO],'
                                   'tblCQ_Final[MOTIVO_EXCLUSAO_AUTOMATICA]),"atualize os dados"))'),
    ]),
    # ADR-070: so o que o usuario preenche (ID, ANALITO de conferencia, caixa, MOTIVO) + a trilha de auditoria
    # (data e usuario, carimbados pelo VBA, travados). Sem formula: o Power Query confere o ID contra o analito.
    'tblInativacao_NaoConformes': ('Inativar', [
        ('ID_REGISTRO', 14, '@', 'in'),
        ('ANALITO', 24, '@', 'in'),
        ('REGISTRAR - LJ', 11, 'General', 'in'),
        ('MOTIVO', 50, '@', 'in'),
        ('DATA_INATIVACAO', 16, 'dd/mm/yyyy hh:mm', 'auto'),
        ('USUARIO', 14, '@', 'auto'),
    ]),
    'tblComentariosTecnicos': ('COMENTARIOS_TECNICOS', [
        ('ID_REGISTRO', 14, '@', 'in'),
        ('COMENTARIO_TECNICO', 70, '@', 'in'),
        ('DATA', 16, 'dd/mm/yyyy hh:mm', 'auto'),
        ('USUARIO', 14, '@', 'auto'),
        ('ANALITO', 24, 'General', LOOKUP.format(c='ANALITO')),
        ('NIVEL', 7, '0', LOOKUP.format(c='NIVEL')),
        ('INATIVADO', 10, 'General', LOOKUP.format(c='INATIVACAO_REGISTRADA')),
    ]),
}
# ADR-070: layout da tblInativacao_NaoConformes ate 06/10/2026 -- colunas de FORMULA que a migracao descarta
# (o ANALITO antigo era XLOOKUP; o novo e digitado e, nas linhas antigas, vem da tblCQ_Final)
INAT_FORMULAS_LEGADO = ['ANALITO', 'NIVEL', 'LOTE', 'DATA_HORA', 'RUN', 'RESULTADO', 'PLOTAGEM_LJ', 'JUSTIFICATIVA']
MOTIVOS_MANUAL = ['Falha do interfaceamento', 'Resultado não transmitido', 'Equipamento em contingência', 'Outro']

LARG_SAIDA = {'ID_REGISTRO': 13, 'DATA': 11, 'DATA_HORA': 16, 'HORA': 8, 'ANALITO': 24, 'ANALITO_ORIGEM': 12,
              'STATUS_ANALITICO': 19, 'TIPO_PLOTAGEM_LJ': 15, 'COMENTARIO_TECNICO': 40, 'GOVERNANCA': 30,
              'MOTIVO_INATIVACAO': 30,
              'MOTIVO_EXCLUSAO_AUTOMATICA': 40, 'DETALHE': 90, 'TESTE': 48, 'RECEBIDO_EM': 16, 'EQUIPAMENTO': 13}
FMT_SAIDA = {'DATA': 'dd/mm/yyyy', 'DATA_HORA': 'dd/mm/yyyy hh:mm:ss', 'HORA': 'hh:mm:ss', 'RECEBIDO_EM': 'dd/mm/yyyy hh:mm',
             'DATA_INATIVACAO': 'dd/mm/yyyy hh:mm', 'RUN': '0', 'ID_ORIGEM': '0', 'ITEM_ID': '0'}


# ================================================================ utilitarios
def log(*a):
    print(*a, flush=True)


def aba(wb, nome, depois=None):
    for ws in wb.Worksheets:
        if ws.Name == nome:
            return ws, False
    ref = depois if depois is not None else wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(None, ref)
    ws.Name = nome
    return ws, True


def desproteger(ws):
    try:
        if ws.ProtectContents:
            ws.Unprotect(SENHA)
    except Exception:
        pass


def cabecalho(ws, titulo, subtitulo, instrucoes, ncols=12):
    ws.Cells.Font.Name = tema.FONTE
    ws.Cells.Font.Size = tema.CORPO
    ult = ws.Cells(1, ncols).Address.replace('$1', '').replace('$', '')
    for lin, txt, tam, cor_txt, alt in ((1, titulo, tema.TITULO, tema.MARCA_CLARA, 40.2),
                                        (2, subtitulo, tema.MIUDO, tema.SOBRE_MARCA, 20)):
        r = ws.Range(f'A{lin}:{ult}{lin}')
        try:
            r.UnMerge()
        except Exception:
            pass
        r.Merge()
        r.Interior.Color = tema.TINTA
        r.VerticalAlignment = -4108
        r.HorizontalAlignment = -4131
        r.IndentLevel = 1
        ws.Range(f'A{lin}').Value = txt
        f = ws.Range(f'A{lin}').Font
        f.Size, f.Color, f.Bold = tam, cor_txt, lin == 1
        ws.Rows(lin).RowHeight = alt
    for i, txt in enumerate(instrucoes[:2]):
        c = ws.Range(f'A{3 + i}')
        c.Value = txt
        c.Font.Size, c.Font.Color = tema.MIUDO, tema.TEXTO_FRACO
        c.WrapText = False
    ws.Rows(5).RowHeight = 22


def textos_topo(wb, nome):
    """So os TEXTOS do topo de uma aba (titulo, subtitulo e as 2 linhas de instrucao), sem refazer a formatacao
    da aba (cabecalho() reformata a aba inteira). Usado por instaladores incrementais (ADR-070)."""
    for n, tit, sub, instr in ABAS:
        if n == nome:
            ws = [w for w in wb.Worksheets if w.Name == nome][0]
            ws.Range('A1').Value = tit
            ws.Range('A2').Value = sub
            for i, txt in enumerate(instr[:2]):
                ws.Range(f'A{3 + i}').Value = txt
            return ws
    raise KeyError(nome)


def formatar_cabecalho_tabela(lo, tipos=None):
    hr = lo.HeaderRowRange
    hr.Font.Bold = True
    hr.Font.Size = tema.MIUDO
    hr.Font.Color = tema.SOBRE_MARCA
    hr.Interior.Color = tema.MARCA
    hr.WrapText = True
    hr.VerticalAlignment = -4108
    lo.Parent.Rows(hr.Row).RowHeight = 30
    if tipos:
        for j, t in enumerate(tipos, start=1):
            c = hr.Cells(1, j)
            if t == 'in':
                c.Interior.Color = tema.ACAO
            elif t == 'auto':
                c.Interior.Color = tema.MARCA
            else:                                   # formula de conferencia
                c.Interior.Color = tema.cor('#5F7472')


def larguras(ws, lo, mapa, padrao=11):
    for j, col in enumerate(lo.ListColumns, start=1):
        w = mapa.get(col.Name, padrao)
        ws.Columns(lo.Range.Column + j - 1).ColumnWidth = w


# ================================================================ configuracao
def montar_cfg(wb, produto, ws):
    """Cria/atualiza tblConfigIntegracao e tblDeParaAnalitos. Valor ja editado pelo usuario e preservado."""
    for nome, cab, linhas, dest in (
            ('tblConfigIntegracao', ['CHAVE', 'VALOR', 'DESCRICAO'], cp.linhas_cfg(produto), f'A{LINHA_CAB}'),
            ('tblDeParaAnalitos', ['MATRIZ', 'ANALITO_ORIGEM', 'ANALITO_QCINI', 'BLOCO', 'ORDEM', 'OBS'],
             cp.linhas_depara(produto), f'E{LINHA_CAB}')):
        lo = pqlib.tabela(wb, nome)
        if lo is None:
            lo = pqlib.criar_tabela_dados(ws, nome, cab, linhas, dest)
            formatar_cabecalho_tabela(lo)
            log(f'  {nome}: criada ({len(linhas)} linhas)')
        elif nome == 'tblConfigIntegracao':
            # acrescenta chave nova; nao sobrescreve valor existente
            atuais = {str(r['CHAVE']).strip().upper(): r for r in pqlib.ler_tabela(lo)}
            novas = [l for l in linhas if l[0].upper() not in atuais]
            for l in novas:
                lr = lo.ListRows.Add()
                for j, v in enumerate(l, start=1):
                    lr.Range.Cells(1, j).Value = v
            log(f'  {nome}: existia; {len(novas)} chave(s) nova(s)')
        else:
            log(f'  {nome}: existia (preservada)')
    for col, w in zip('ABCDEFGHIJ', (24, 70, 60, 4, 9, 18, 30, 10, 7, 46)):
        ws.Columns(col).ColumnWidth = w


# ================================================================ entradas
def formato_geral(rng):
    """Formato Geral numa faixa. 'General' e recusado via COM no Excel pt-BR; 'Geral' e o nome local.
    Ultimo recurso: limpar os formatos (o Geral e o padrao)."""
    for tentativa in (lambda: setattr(rng, 'NumberFormat', 'General'),
                      lambda: setattr(rng, 'NumberFormatLocal', 'Geral'),
                      lambda: rng.ClearFormats()):
        try:
            tentativa()
            return
        except Exception:
            continue


def _norm_id(v, prefixo):
    """Espelho do NormId do Power Query (DB_CQ_FINAL/QA_INTEGRACAO) e do mIntegracao.NormalizarId."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = str(v).strip().replace(' ', '').replace('\xa0', '').upper()
    if not t:
        return None
    dig = lambda s: s != '' and s.isdigit()      # noqa: E731
    sem0 = lambda d: d.lstrip('0') or '0'        # noqa: E731
    pre = prefixo + '-'
    if dig(t):
        return pre + sem0(t)
    if t.startswith(pre) and dig(t[len(pre):]):
        return pre + sem0(t[len(pre):])
    if t.startswith('MAN_') and dig(t[4:]):
        return 'MAN_' + t[4:].rjust(4, '0')
    return t


def _coluna(lo, nome, n):
    """Valores (Value2: data = numero de serie, sem fuso do pywin32) de uma coluna da tabela, como lista."""
    v = lo.ListColumns(nome).DataBodyRange.Value2
    return [r[0] for r in v] if n > 1 else [v]


def migrar_inativacao(wb, lo, cols):
    """ADR-070: tblInativacao_NaoConformes do layout antigo (ID, caixa, carimbo + 8 colunas de FORMULA) para o
    novo (ID, ANALITO, caixa, MOTIVO, carimbo). Toda linha digitada e preservada NA MESMA POSICAO (ID,
    REGISTRAR - LJ, DATA_INATIVACAO, USUARIO); o ANALITO das linhas antigas vem da tblCQ_Final pelo ID (o mesmo
    que a formula antiga buscava), o MOTIVO fica vazio (a justificativa antiga continua valendo pelos
    COMENTARIOS_TECNICOS). Idempotente: no layout novo nao faz nada. Devolve um resumo (dict)."""
    cab = [c[0] for c in cols]
    atuais = [c.Name for c in lo.ListColumns]
    if atuais == cab:
        return {'migrada': False}
    ws = lo.Parent
    nome = lo.Name
    largura_cab = ws.Range('A1').MergeArea.Width          # a barra de navegacao foi montada sobre ela
    n = lo.ListRows.Count
    # o que a tabela antiga tinha de DIGITADO/CARIMBADO; colunas de formula sao descartadas
    legado = set(INAT_FORMULAS_LEGADO) if 'MOTIVO' not in atuais else set()
    dados = {h: _coluna(lo, h, n) for h in atuais if h in cab and h not in legado} if n else {}
    prefixo = ''
    cfg = pqlib.tabela(wb, 'tblConfigIntegracao')
    if cfg is not None:
        for r in pqlib.ler_tabela(cfg):
            if str(r.get('CHAVE') or '').strip().upper() == 'PREFIXO_ID':
                prefixo = str(r.get('VALOR') or '').strip().upper()
    fin = pqlib.tabela(wb, 'tblCQ_Final')
    analito_de = {}
    if fin is not None and fin.ListRows.Count:
        nf = fin.ListRows.Count
        for i, a in zip(_coluna(fin, 'ID_REGISTRO', nf), _coluna(fin, 'ANALITO', nf)):
            if i not in (None, ''):
                analito_de[str(i).strip().upper()] = a
    ids = dados.get('ID_REGISTRO', [None] * n)
    an = dados.get('ANALITO', [None] * n)
    preenchidos = 0
    for k, i in enumerate(ids):
        idn = _norm_id(i, prefixo)
        if idn and an[k] in (None, '') and idn in analito_de:
            an[k] = analito_de[idn]
            preenchidos += 1
    dados['ANALITO'] = an
    # ultima linha com algum conteudo: as linhas (inclusive as vazias no meio) ficam onde estavam
    ult = 0
    for k in range(n):
        if any(dados[h][k] not in (None, '') for h in dados if h != 'REGISTRAR - LJ') or \
                dados.get('REGISTRAR - LJ', [None] * n)[k] is True:
            ult = k + 1
    antigas = [{h: dados[h][k] for h in dados} for k in range(ult)]
    r0, c0 = lo.Range.Row, lo.Range.Column
    area = lo.Range.Address
    lo.Delete()
    ws.Range(area).Clear()                                # conteudo, formatos, caixa de selecao e validacoes
    try:
        ws.Range(area).Validation.Delete()
    except Exception:
        pass
    folga = max(FOLGA_ENTRADA[nome], ult + 100)
    for j, h in enumerate(cab):
        ws.Cells(r0, c0 + j).Value = h
    novo = ws.ListObjects.Add(1, ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + folga, c0 + len(cab) - 1)), None, 1)
    novo.Name = nome
    novo.TableStyle = 'TableStyleLight1'
    if ult:
        for j, h in enumerate(cab):
            if h in dados:
                col = novo.ListColumns(h).DataBodyRange
                faixa = ws.Range(col.Cells(1, 1), col.Cells(ult, 1))
                if h in ('ID_REGISTRO', 'ANALITO', 'MOTIVO', 'USUARIO'):
                    faixa.NumberFormat = '@'
                faixa.Value2 = tuple((dados[h][k],) for k in range(ult))
    return {'migrada': True, 'linhas': ult, 'ids': sum(1 for i in ids[:ult] if i not in (None, '')),
            'analito_preenchido': preenchidos, 'largura_cab': largura_cab, 'antigas': antigas}


def igualar_largura_cabecalho(ws, ncols_tabela, alvo, ncols=12):
    """Depois da migracao a tabela tem menos colunas: as colunas vazias que sobram ate a ultima do cabecalho
    (L) ficam com a largura que deixa a faixa do titulo (A1:L1, onde a barra de navegacao foi montada) com a
    MESMA largura de antes -- os botoes, de posicao fixa, continuam alinhados a direita da faixa."""
    resto = list(range(ncols_tabela + 1, ncols + 1))
    if not resto or not alvo:
        return
    tabela = sum(ws.Columns(c).Width for c in range(1, ncols_tabela + 1))
    alvo_col = max(3.0, (alvo - tabela) / len(resto))           # pontos por coluna, iguais
    for c in resto:
        col = ws.Columns(c)
        for _ in range(3):                                      # largura em caracteres -> pontos nao e linear
            pt_por_car = (col.Width / col.ColumnWidth) if col.ColumnWidth else 5.25
            w = col.ColumnWidth + (alvo_col - col.Width) / pt_por_car
            col.ColumnWidth = max(0.5, min(60, w))
            if abs(col.Width - alvo_col) < 1:
                break


def montar_entrada(wb, produto, nome):
    nome_aba, cols = ENTRADAS[nome]
    ws = [w for w in wb.Worksheets if w.Name == nome_aba][0]
    lo = pqlib.tabela(wb, nome)
    cab = [c[0] for c in cols]
    criada = False
    migracao = None
    if lo is None:
        r0 = LINHA_CAB
        folga = FOLGA_ENTRADA[nome]
        for j, h in enumerate(cab, start=1):
            ws.Cells(r0, j).Value = h
        rng = ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + folga, len(cab)))
        lo = ws.ListObjects.Add(1, rng, None, 1)
        lo.Name = nome
        lo.TableStyle = 'TableStyleLight1'
        criada = True
    elif nome == 'tblInativacao_NaoConformes' and [c.Name for c in lo.ListColumns] != cab:
        # ADR-070: layout novo, sem formula; preserva as linhas
        migracao = migrar_inativacao(wb, lo, cols)
        lo = pqlib.tabela(wb, nome)
    else:
        atuais = [c.Name for c in lo.ListColumns]
        if atuais[:len(cab)] != cab:
            # tabela antiga com outro layout: acrescenta so as colunas que faltam (nunca apaga digitacao)
            for h in cab:
                if h not in atuais:
                    lo.ListColumns.Add().Name = h
    formatar_cabecalho_tabela(lo, [c[3] if c[3] in ('in', 'auto') else 'f' for c in cols])
    larguras(ws, lo, {c[0]: c[1] for c in cols})
    if migracao and migracao.get('migrada'):
        igualar_largura_cabecalho(ws, len(cab), migracao['largura_cab'])
    ws.Cells.Locked = True
    for h, w, fmt, tipo in cols:
        colr = lo.ListColumns(h).DataBodyRange
        formula = tipo not in ('in', 'auto')
        if formula or fmt == 'General':
            # coluna de formula NUNCA em '@' (a formula viraria texto literal); Geral pelo nome local
            if formula or str(colr.Cells(1, 1).NumberFormat) == '@':
                formato_geral(colr)
        else:
            colr.NumberFormat = fmt
        if tipo == 'in':
            colr.Locked = False          # o que o usuario digita fica destravado com a aba protegida
        elif tipo == 'auto':
            # carimbado pelo VBA (data/usuario). O ID manual e gerado pelo VBA, mas o
            # usuario pode digitar um MAN_ proprio -- por isso fica destravado
            colr.Locked = not (h == 'ID_REGISTRO' and nome == 'tblResultados_Manuais')
            colr.Interior.Color = tema.FUNDO_CARTAO
        else:
            if fmt != 'General':
                colr.NumberFormat = fmt
            # '[@COL]' e recusado via COM no Excel pt-BR; a forma longa e aceita. Reescrita sempre: a formula
            # que ficou guardada como texto (formato '@') volta a calcular
            colr.Formula2 = re.sub(r'\[@([^\]]+)\]', lambda m: f'{nome}[[#This Row],[{m.group(1)}]]', tipo)
            colr.Locked = True
            colr.Font.Color = tema.TEXTO_FRACO
    validacoes(wb, produto, nome, lo)
    if nome == 'tblInativacao_NaoConformes':
        caixa = lo.ListColumns('REGISTRAR - LJ').DataBodyRange
        try:
            caixa.CellControl.SetCheckbox()
        except Exception as e:     # Excel sem checkbox nativo: lista SIM/NAO como alternativa
            log('  checkbox nativo indisponivel, usando lista SIM/NÃO:', e)
            caixa.Validation.Delete()
            caixa.Validation.Add(3, 1, 1, 'SIM,NÃO')
        caixa.HorizontalAlignment = -4108
    log(f'  {nome}: {"criada" if criada else "existia"} · {lo.ListRows.Count} linhas · {len(cab)} colunas')
    return lo


def validacoes(wb, produto, nome, lo):
    def lista(col, itens_ou_formula, msg):
        r = lo.ListColumns(col).DataBodyRange
        r.Validation.Delete()
        src = itens_ou_formula if itens_ou_formula.startswith('=') else itens_ou_formula
        r.Validation.Add(3, 1, 1, src)
        r.Validation.IgnoreBlank = True
        r.Validation.ErrorMessage = msg
        r.Validation.ShowError = True

    if nome == 'tblResultados_Manuais':
        p = cp.PRODUTOS[produto]
        lista('EQUIPAMENTO', ','.join(p['equipamentos']), 'Escolha o equipamento da lista.')
        matrizes = sorted({m for m, *_ in p['depara'] if m}) or [p['cfg']['MATRIZ_PADRAO'][0]]
        lista('MATRIZ', ','.join(matrizes), 'Escolha a matriz da lista.')
        lista('LOTE', '=lstLotes', 'Use um lote cadastrado (aba Configuração).')
        lista('ANALITO', '=lstAnalitos', 'Use o nome do analito como está no cadastro (aba Analitos).')
        lista('MOTIVO', ','.join(MOTIVOS_MANUAL), 'Escolha o motivo da lista.')
        n = lo.ListColumns('NIVEL').DataBodyRange
        n.Validation.Delete()
        n.Validation.Add(1, 1, 1, '1', str(p['cfg']['NIVEIS'][0]))          # inteiro entre 1 e NIVEIS
        n.Validation.ErrorMessage = f'Nível de 1 a {p["cfg"]["NIVEIS"][0]}.'
        d = lo.ListColumns('DATA').DataBodyRange
        d.Validation.Delete()
        d.Validation.Add(4, 1, 5, '1/1/2020')                             # data >= 2020
        d.Validation.ErrorMessage = 'Data inválida.'
        r = lo.ListColumns('RESULTADO').DataBodyRange
        r.Validation.Delete()
        r.Validation.Add(2, 1, 7, '-1E+307')                              # decimal
        r.Validation.ErrorMessage = 'Resultado numérico.'
    elif nome == 'tblInativacao_NaoConformes':
        # ADR-070: o analito e a CONFERENCIA do ID -- escolhido da lista do cadastro (a mesma do Painel)
        lista('ANALITO', '=lstAnalitos', 'Escolha o analito da lista (o mesmo nome da aba Analitos).')

        def dica(col, titulo, msg):
            r = lo.ListColumns(col).DataBodyRange
            r.Validation.Delete()
            r.Validation.Add(0, 1, 1)                                     # xlValidateInputOnly: so a mensagem
            r.Validation.InputTitle = titulo
            r.Validation.InputMessage = msg
            r.Validation.ShowInput = True
        dica('ID_REGISTRO', 'ID do resultado', 'Digite só o número do ID (ex.: 314216): é o número que aparece ao '
                                              'passar o mouse no ponto do Levey-Jennings, no Painel.')
        dica('MOTIVO', 'Motivo (obrigatório)', 'Por que este resultado sai da estatística? '
                                               'Ex.: repetição, erro pré-analítico, controle trocado.')


# ================================================================ consultas e saidas
def gravar_consultas(wb):
    wb.Queries.FastCombine = True
    for q in CONSULTAS:
        pqlib.gravar_query(wb, q, pqlib.m_de(q))


def carregar_saida(wb, consulta, m_alternativo=None):
    nome_aba, nome_tab = SAIDAS[consulta]
    ws = [w for w in wb.Worksheets if w.Name == nome_aba][0]
    desproteger(ws)
    if m_alternativo is not None:
        pqlib.gravar_query(wb, consulta, m_alternativo)
    lo = pqlib.tabela(wb, nome_tab)
    if lo is None:
        lo, t = pqlib.carregar_em_tabela(ws, consulta, nome_tab, f'A{LINHA_CAB}')
        lo.TableStyle = 'TableStyleLight1'
    else:
        t = pqlib.atualizar(lo)
    formatar_cabecalho_tabela(lo)
    larguras(ws, lo, LARG_SAIDA)
    for col in lo.ListColumns:
        if col.Name in FMT_SAIDA and col.DataBodyRange is not None:
            col.DataBodyRange.NumberFormat = FMT_SAIDA[col.Name]
    return lo, t


def resumo_saidas(wb):
    """Linha 5 de cada aba de saida: contagem viva por formula (sem VBA)."""
    f = {
        'DB': '="Recebidos: "&TEXT(ROWS(tblDB_Recebimento),"#,##0")&" resultados  ·  último resultado: "&'
                          'TEXT(MAX(tblDB_Recebimento[DATA_HORA]),"dd/mm/yyyy hh:mm")&"  ·  última chegada: "&'
                          'TEXT(MAX(tblDB_Recebimento[RECEBIDO_EM]),"dd/mm/yyyy hh:mm")',
        'DB_ORGANIZADO': '="Linhas (instante × nível): "&TEXT(ROWS(tblDB_Organizado),"#,##0")',
        'Principal - Resultados': '="Total: "&TEXT(ROWS(tblCQ_Final),"#,##0")&"  ·  participam: "&'
                       'TEXT(COUNTIFS(tblCQ_Final[PARTICIPA_ESTATISTICA],"SIM"),"#,##0")&"  ·  X vermelho: "&'
                       'COUNTIFS(tblCQ_Final[TIPO_PLOTAGEM_LJ],"X_VERMELHO")&"  ·  inativados: "&'
                       'COUNTIFS(tblCQ_Final[STATUS_ANALITICO],"INATIVADO")&"  ·  manuais: "&'
                       'COUNTIFS(tblCQ_Final[ORIGEM_RESULTADO],"MANUAL")',
        'QA_INTEGRACAO': '="ERROS: "&COUNTIFS(tblQA_Integracao[SEVERIDADE],"ERRO")&"   ·   ALERTAS: "&'
                         'COUNTIFS(tblQA_Integracao[SEVERIDADE],"ALERTA")&"   ·   INFO: "&'
                         'COUNTIFS(tblQA_Integracao[SEVERIDADE],"INFO")',
    }
    for nome_aba, formula in f.items():
        ws = [w for w in wb.Worksheets if w.Name == nome_aba][0]
        c = ws.Range('A5')
        c.Formula2 = formula
        c.Font.Bold = True
        c.Font.Size = tema.CORPO
        c.Font.Color = tema.TINTA


def montar(wb, produto, carregar=True, m_final=None, limite=None):
    """Monta tudo. m_final/limite: so para testes (M alternativo / primeiras N linhas recebidas)."""
    try:
        wb.Unprotect(SENHA)
    except Exception:
        pass
    ref = None
    for n in ('Registros', 'Eventos_Westgard', 'Painel'):
        if any(w.Name == n for w in wb.Worksheets):
            ref = wb.Worksheets(n)
            break
    for nome, tit, sub, instr in ABAS:
        ws, nova = aba(wb, nome, ref)
        ref = ws
        desproteger(ws)
        cabecalho(ws, tit, sub, instr)
        ws.Tab.Color = tema.MARCA if nome not in ('Cfg_Integracao',) else tema.TEXTO_FRACO
        log(f'aba {nome}: {"criada" if nova else "existia"}')
    cfg = wb.Worksheets('Cfg_Integracao')
    montar_cfg(wb, produto, cfg)
    gravar_consultas(wb)
    tempos = {}
    if carregar:
        # 1a carga: as tabelas de entrada referenciam tblCQ_Final nas formulas de conferencia,
        # entao a saida nasce antes delas (a ordem das CONSULTAS nao depende disso)
        for q in ORDEM_CARGA:
            m_alt = None
            if q == 'DB_CQ_FINAL' and (m_final or limite):
                m_alt = m_final or pqlib.m_de('DB_CQ_FINAL')
                if limite:
                    alvo = 'R0 = Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}[Content],'
                    assert alvo in m_alt
                    m_alt = m_alt.replace(alvo, 'R0 = Table.FirstN(Excel.CurrentWorkbook(){[Name = "tblDB_Recebimento"]}'
                                                f'[Content], {int(limite)}),')
            lo, t = carregar_saida(wb, q, m_alt)
            tempos[q] = (round(t, 1), lo.ListRows.Count, lo.ListColumns.Count)
            log(f'  {q}: {t:.1f}s · {lo.ListRows.Count} linhas · {lo.ListColumns.Count} colunas')
        for nome in ENTRADAS:
            montar_entrada(wb, produto, nome)
        resumo_saidas(wb)
    return tempos
