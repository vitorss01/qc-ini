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
     'INATIVAR — retirar resultados da população principal (não conformes / repetições)',
     'Soft-delete por ID_REGISTRO: o resultado continua na base, sai da população estatística',
     ['Digite (ou cole) o ID_REGISTRO copiado da aba Principal - Resultados.  REGISTRAR - LJ marcado (padrão) = aparece no '
      'Levey-Jennings como X VERMELHO · desmarcado = não aparece no gráfico.',
      'Para REATIVAR, apague a linha.  Toda inativação exige justificativa na aba COMENTARIOS_TECNICOS '
      '(sem ela o QA acusa ERRO DE GOVERNANÇA).  Depois clique ⟳ ATUALIZAR DADOS.']),
    ('COMENTARIOS_TECNICOS',
     'COMENTÁRIOS TÉCNICOS — justificativa de cada inativação',
     'Relacionados pelo ID_REGISTRO · obrigatórios para todo resultado inativado',
     ['Uma linha por comentário. Mais de um comentário para o mesmo ID aparece junto na DB_CQ_FINAL (separados por " | ").',
      'Comentário de resultado reativado fica como histórico (o QA informa, não acusa erro).']),
    ('Principal - Resultados',
     'PRINCIPAL - RESULTADOS — fonte oficial dos resultados do CQ (DB_CQ_FINAL)',
     'Single source of truth do Levey-Jennings, Westgard, Estatística e Power BI · gerada só pelo Power Query',
     ['Uma linha = um resultado. Inativados CONTINUAM aqui (PARTICIPA_ESTATISTICA = NÃO). '
      'TIPO_PLOTAGEM_LJ: NORMAL = ponto · X_VERMELHO = X no gráfico · NAO_PLOTAR = só auditoria.',
      'Para inativar: copie o ID_REGISTRO e cole na aba Inativar.  Não edite esta tabela — '
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
# 'f' formula de conferencia (calculada, bloqueada)
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
        ('STATUS', 20, '@', LOOKUP.format(c='STATUS_ANALITICO')),
        ('DETALHE', 50, '@', '=IF([@ID_REGISTRO]="","",IFERROR(XLOOKUP([@ID_REGISTRO],tblCQ_Final[ID_REGISTRO],'
                             'tblCQ_Final[MOTIVO_EXCLUSAO_AUTOMATICA]),"atualize os dados"))'),
    ]),
    'tblInativacao_NaoConformes': ('Inativar', [
        ('ID_REGISTRO', 14, '@', 'in'),
        ('REGISTRAR - LJ', 11, 'General', 'in'),
        ('DATA_INATIVACAO', 16, 'dd/mm/yyyy hh:mm', 'auto'),
        ('USUARIO', 14, '@', 'auto'),
        ('ANALITO', 24, '@', LOOKUP.format(c='ANALITO')),
        ('NIVEL', 7, '0', LOOKUP.format(c='NIVEL')),
        ('LOTE', 9, '@', LOOKUP.format(c='LOTE')),
        ('DATA_HORA', 16, 'dd/mm/yyyy hh:mm', LOOKUP.format(c='DATA_HORA')),
        ('RUN', 11, '0', LOOKUP.format(c='RUN')),
        ('RESULTADO', 11, 'General', LOOKUP.format(c='RESULTADO')),
        ('PLOTAGEM_LJ', 13, '@', LOOKUP.format(c='TIPO_PLOTAGEM_LJ')),
        ('JUSTIFICATIVA', 26, '@', '=IF([@ID_REGISTRO]="","",IF(COUNTIF(tblComentariosTecnicos[ID_REGISTRO],'
                                   '[@ID_REGISTRO])>0,"OK","FALTA — ver COMENTARIOS_TECNICOS"))'),
    ]),
    'tblComentariosTecnicos': ('COMENTARIOS_TECNICOS', [
        ('ID_REGISTRO', 14, '@', 'in'),
        ('COMENTARIO_TECNICO', 70, '@', 'in'),
        ('DATA', 16, 'dd/mm/yyyy hh:mm', 'auto'),
        ('USUARIO', 14, '@', 'auto'),
        ('ANALITO', 24, '@', LOOKUP.format(c='ANALITO')),
        ('NIVEL', 7, '0', LOOKUP.format(c='NIVEL')),
        ('INATIVADO', 10, '@', LOOKUP.format(c='INATIVACAO_REGISTRADA')),
    ]),
}
MOTIVOS_MANUAL = ['Falha do interfaceamento', 'Resultado não transmitido', 'Equipamento em contingência', 'Outro']

LARG_SAIDA = {'ID_REGISTRO': 13, 'DATA': 11, 'DATA_HORA': 16, 'HORA': 8, 'ANALITO': 24, 'ANALITO_ORIGEM': 12,
              'STATUS_ANALITICO': 19, 'TIPO_PLOTAGEM_LJ': 15, 'COMENTARIO_TECNICO': 40, 'GOVERNANCA': 30,
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
def montar_entrada(wb, produto, nome):
    nome_aba, cols = ENTRADAS[nome]
    ws = [w for w in wb.Worksheets if w.Name == nome_aba][0]
    lo = pqlib.tabela(wb, nome)
    cab = [c[0] for c in cols]
    criada = False
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
    else:
        atuais = [c.Name for c in lo.ListColumns]
        if atuais[:len(cab)] != cab:
            # tabela antiga com outro layout: acrescenta so as colunas que faltam (nunca apaga digitacao)
            for h in cab:
                if h not in atuais:
                    lo.ListColumns.Add().Name = h
    formatar_cabecalho_tabela(lo, [c[3] if c[3] in ('in', 'auto') else 'f' for c in cols])
    larguras(ws, lo, {c[0]: c[1] for c in cols})
    body = lo.DataBodyRange
    ws.Cells.Locked = True
    for h, w, fmt, tipo in cols:
        colr = lo.ListColumns(h).DataBodyRange
        if fmt != 'General':                # 'General' e recusado pelo Excel pt-BR via COM; e o padrao mesmo
            colr.NumberFormat = fmt
        if tipo == 'in':
            colr.Locked = False          # o que o usuario digita fica destravado com a aba protegida
        elif tipo == 'auto':
            # carimbado pelo VBA (data/usuario). O ID manual e gerado pelo VBA, mas o
            # usuario pode digitar um MAN_ proprio -- por isso fica destravado
            colr.Locked = not (h == 'ID_REGISTRO' and nome == 'tblResultados_Manuais')
            colr.Interior.Color = tema.FUNDO_CARTAO
        else:
            # '[@COL]' e recusado via COM no Excel pt-BR; a forma longa e aceita
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
