# -*- coding: utf-8 -*-
"""qa_manual_substitui.py -- ADR-082: o resultado MANUAL substitui um resultado do interfaceamento por ESCOLHA do
usuario (coluna SUBSTITUI_ID da aba Digitar Resultados).

Uso: python qa_manual_substitui.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Pelo caminho real: escreve nas abas com os eventos ligados (o Worksheet_Change normaliza, preenche e carimba) e roda o
ATUALIZAR DADOS (mIntegracao.AtualizarDadosAutomatico). MODO_FONTE = HISTORICO durante o teste (nada novo chega do
DB_SEAC no meio: as comparacoes de RUN e PQ x VBA ficam exatas). Nada salvo.

  M01 estrutura: SUBSTITUI_ID no fim da tblResultados_Manuais, texto, destravada, cabecalho de digitacao; o topo da aba
      explica a substituicao
  M02 digitar o SUBSTITUI_ID preenche EQUIPAMENTO, MATRIZ, LOTE, NIVEL, ANALITO, DATA e HORA do resultado indicado (so o
      vazio), normaliza o ID e gera o MAN_; ID inexistente nao preenche nada (e nao abre caixa em automacao)
  M03 substituicao valida: manual ATIVO/SIM/NORMAL com ID_RELACIONADO = automatico; automatico SUBSTITUIDO_MANUAL, NÃO,
      NAO_PLOTAR, sem RUN e sem corrida, ID_RELACIONADO = manual, MOTIVO "substituido pelo manual ... (motivo: ...)";
      o manual fica com o RUN e a corrida do dia que eram do automatico; nenhum outro RUN muda
  M04 grafico: corrida do lote = oraculo independente (qa_corrida_lote), sem lacuna; o manual no lugar do automatico
  M05 QA: I04 do automatico (o manual, os dois valores, a diferenca e o motivo); nenhum E12; nenhum A01 do manual; o A06
      do manual (se houver) tem o texto de quem substitui e NAO lista o automatico substituido
  M06 Audit_Log: RESULTADO_SUBSTITUIDO_POR_MANUAL (automatico; resAnt = valor dele, resNovo = o do manual, RUN = o do
      manual) e RESULTADO_MANUAL_INCLUIDO (manual)
  M07 dica do ponto do manual: "ID MAN_x (substitui <ID>)" DENTRO da parte do ID; aviso da janela de inativar (o efeito
      de sempre + inativar o manual nao traz o automatico de volta; cabe no rotulo: <= 150 caracteres); a dica de um ponto
      do interfaceamento nao muda
  M08 editar o RESULTADO do manual que substitui: RESULTADO_MANUAL_ALTERADO no Audit_Log (antes -> depois)
  M09 desfazer: apagar o SUBSTITUI_ID -> o automatico volta ATIVO com o RUN e a corrida de antes; o manual volta as
      regras sem SUBSTITUI_ID (CONFLITO_MANUAL pelo horario, A01); SUBSTITUICAO_POR_MANUAL_DESFEITA no Audit_Log
  M10 SUBSTITUI_ID invalido -> E12 com o motivo exato e NADA substituido: ID inexistente, chave diferente (outro nivel),
      ID de um manual, retransmissao (se a base tiver), segundo manual para o mesmo ID (CONFLITO_MANUAL "outro manual
      (...) ja substitui ...", A01; o 1o vale -- e continua valendo INATIVADO)
  M11 interfaceamento INATIVADO indicado: continua INATIVADO (X vermelho, GOVERNANCA OK, sem E01/A04) e o manual vale;
      vinculo nos dois sentidos; SEM corrida propria: o RUN e a corrida do manual, nenhum outro RUN do grupo muda; o I04
      diz que a inativacao prevalece e como o automatico volta (reativar); nenhum A06 do manual lista o inativado; a
      janela de reativar o X avisa que ele continua fora
  M12 SUBSTITUI_ID sem MOTIVO: MANUAL_INCOMPLETO "MOTIVO (obrigatorio para substituir um resultado)" + E04; nada muda
  M13 SEM_VALOR do interfaceamento indicado: SUBSTITUIDO_MANUAL e o manual ATIVO (se a base tiver SEM_VALOR)
  M14 idempotencia: ATUALIZAR de novo -> tblCQ_Final identica e nenhuma linha nova de RESULTADO_*/SUBSTITUICAO_* na trilha
  M15 PQ x VBA: a tblCQ_Final do motor VBA = a do Power Query (MOTOR_DADOS = PQ), com todas as substituicoes acima, e
      os mesmos achados E04/E09/E12/A01/A06/A09/I04 na tblQA_Integracao (o QA inteiro e do estudos/equivalencia_qa.py)
  M16 o 1o manual (o que vale) INATIVADO continua substituindo: o E12 do 2o manual diz isso e que a correcao e apagar o
      SUBSTITUI_ID do 1o; nenhum A09 do 2o (o automatico nao esta inativado)
  M17 seguindo o E12 (apagar o SUBSTITUI_ID do 1o): o 2o passa a substituir -- SUBSTITUICAO_POR_MANUAL_ALTERADA na trilha
      (antes = valor do 1o, depois = do 2o, com quem lancou o manual); o 1o vira CONFLITO_MANUAL com o automatico e o
      A01 dele diz que vale o 2o (nao "informe o ID em SUBSTITUI_ID"), sem A09
  M18 inativar DEPOIS um automatico ja SUBSTITUIDO_MANUAL: INATIVADO/X_VERMELHO com o RUN e a corrida do manual, nenhum
      outro RUN do grupo muda; RESULTADO_INATIVADO na trilha (nenhum REATIVADO/PLOTAGEM)
  M19 SEM_DATA_HORA do interfaceamento indicado (se a base tiver): SUBSTITUIDO_MANUAL e sem E09 (o I04 registra)
  M00 seguranca do teste: arquivo de entrada intacto
Sai com 1 se algum falhar.
"""
import collections
import datetime as dt
import json
import os
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA, reg, RES, NAO, sha256_arquivo  # noqa: E402
from qa_inativar_ux import Leitor, preparar_painel, series, curto, consecutivos, ativos_em_ordem, num_pt  # noqa: E402
import qa_corrida_lote as cl  # noqa: E402

SIM = 'SIM'
TB_MAN = 'tblResultados_Manuais'
TB_INAT = 'tblInativacao_NaoConformes'
TB_FIN = 'tblCQ_Final'
TB_QA = 'tblQA_Integracao'
COL = 'SUBSTITUI_ID'
SEP = ' ' + chr(183) + ' '
MOT = 'Falha do interfaceamento'
EPOCA = dt.datetime(1899, 12, 30)


def txt(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return '' if v is None else str(v).strip()


def iint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def dia(v):
    return v.date() if hasattr(v, 'date') else None


def achados(qa, cod, id_=None):
    return [x for x in qa if x['CODIGO'] == cod and (id_ is None or x['ID_REGISTRO'] == id_)]


def resumo(r):
    if not r:
        return None
    return {k: r.get(k) for k in ('ID_REGISTRO', 'ORIGEM_RESULTADO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                  'TIPO_PLOTAGEM_LJ', 'RUN', 'CORRIDA_NO_DIA', 'RESULTADO', 'ID_RELACIONADO',
                                  'MOTIVO_EXCLUSAO_AUTOMATICA', 'GOVERNANCA')}


def runs(fin):
    return {i: (iint(r['RUN']), iint(r['CORRIDA_NO_DIA'])) for i, r in fin.items()}


def grupo(r):
    """(equipamento, matriz, lote, analito, dia): o grupo da corrida do dia."""
    return (txt(r['EQUIPAMENTO']), txt(r['MATRIZ']), txt(r['LOTE']), txt(r['ANALITO']), dia(r['DATA_HORA']))


def tabela_v2(q, nome, ignorar=()):
    """{ID: tupla de Value2 (vazio = None)} sem as colunas ignoradas -- como o S6 do teste_incremental."""
    lo = q.lo(nome)
    cab = [c.Name for c in lo.ListColumns]
    lin = [list(r) for r in lo.DataBodyRange.Value2] if lo.ListRows.Count else []
    i = cab.index('ID_REGISTRO')
    keep = [j for j, c in enumerate(cab) if c not in ignorar]
    return {r[i]: tuple(None if r[j] in ('', None) else r[j] for j in keep) for r in lin}


def qa_multiconjunto(q):
    lo = q.lo(TB_QA)
    if not lo.ListRows.Count:
        return collections.Counter()
    out = []
    for row in lo.DataBodyRange.Value:
        k = []
        for j, v in enumerate(row):
            if v in ('', None):
                v = None
            elif j == 5:
                v = int(float(v))
            elif j == 6 and hasattr(v, 'year'):
                v = v.replace(tzinfo=None, microsecond=0).isoformat()
            k.append(v)
        out.append(tuple(k))
    return collections.Counter(out)


def audit_ult(q):
    al = q.ws('Audit_Log')
    return al.Cells(al.Rows.Count, 1).End(-4162).Row


def audit_desde(q, r0):
    """Linhas do Audit_Log depois da linha r0: acao, RUN, resAnt/resNovo, stAnt/stNovo, motivo ("ID ...") e parecer."""
    al = q.ws('Audit_Log')
    ult = audit_ult(q)
    if ult <= r0:
        return []
    v = al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 24)).Value
    if not isinstance(v[0], tuple):
        v = (v,)
    return [{'acao': str(x[6] or ''), 'run': x[9], 'res_ant': x[16], 'res_novo': x[17], 'st_ant': x[20],
             'st_novo': x[21], 'motivo': str(x[22] or ''), 'parecer': str(x[23] or '')} for x in v]


def grupo_dia(r):
    """grupo() com o dia da DATA quando nao ha DATA_HORA (SEM_DATA_HORA): o grupo em que o manual entra."""
    d = r['DATA_HORA'] if isinstance(r.get('DATA_HORA'), dt.datetime) else r.get('DATA')
    return (txt(r['EQUIPAMENTO']), txt(r['MATRIZ']), txt(r['LOTE']), txt(r['ANALITO']), dia(d))


def listados_a06(det):
    """Os IDs que o A06 lista (a parte depois de 'equipamento/lote/nivel/analito...: ' ate '. Os listados')."""
    s = str(det or '')
    p = s.find('analito no dia: ')
    p = s.find('analito: ') if p < 0 else p
    if p < 0:
        return []
    s = s[s.find(': ', p) + 2:]
    return [x.strip() for x in s.split('. Os listados')[0].split(',') if x.strip()]


def eventos(aud, acao, id_):
    return [x for x in aud if acao in x['acao'].split('+') and x['motivo'] == f'ID {id_}']


def numero(v):
    """Numero de uma celula do Audit_Log: o mAuditoria grava resAnt/resNovo como TEXTO (CStr, virgula decimal no
    Windows pt-BR); a celula pode voltar texto ou numero."""
    if isinstance(v, (int, float)):
        return float(v)
    return num_pt(v)


def perto(a, b):
    x, y = numero(a), numero(b)
    if x is None or y is None:
        return False
    return abs(x - y) <= 1e-9 * max(1.0, abs(y))


# ---------------------------------------------------------------- escrita na aba Digitar Resultados (eventos ligados)
def escrever_manual(q, valores):
    """Escreve na 1a linha livre (coluna ID_REGISTRO vazia), celula a celula NA ORDEM do dict (dispara o evento a cada
    celula). Devolve a linha da tabela."""
    return q.escrever_linha(TB_MAN, valores)


def linha_manual(q, r):
    lo = q.lo(TB_MAN)
    cab = [c.Name for c in lo.ListColumns]
    v = lo.DataBodyRange.Rows(r).Value[0]
    return {c: (x.replace(tzinfo=None) if hasattr(x, 'tzinfo') and getattr(x, 'tzinfo', None) else x)
            for c, x in zip(cab, v)}


def gravar_celula(q, r, coluna, valor):
    cel = q.lo(TB_MAN).ListColumns(coluna).DataBodyRange.Cells(r, 1)
    if valor is None:
        cel.ClearContents()
    else:
        cel.Value = valor


def serial_hora(d):
    """Fracao do dia ao segundo (o que o preenchimento automatico grava na HORA)."""
    return (d.hour * 3600 + d.minute * 60 + d.second) / 86400.0


# ---------------------------------------------------------------- a suite
def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    L = Leitor(q)
    tempos = {}

    def run(macro, *a, teto=300):
        return str(q.run(macro, *a, teto=teto))

    def atualizar(rotulo):
        t0 = time.time()
        r, _ = q.atualizar()
        tempos[rotulo] = round(time.time() - t0, 1)
        return r

    try:
        # ------------------------------------------------------------ preparo: historico, motor VBA, tela
        modo0 = q.cfg('MODO_FONTE', 'HISTORICO')
        motor0 = q.cfg('MOTOR_DADOS', 'VBA')
        preparar_painel(q)
        atualizar('M00_base')
        fin0 = q.final_por_id()
        base = L.eng_por_run()
        e0 = cl.eng(q, q.nlv)
        nlv = q.nlv
        # X: ponto N1 do interfaceamento no meio do grafico (vizinhas dos dois lados)
        cand = []
        for rn, d in base.items():
            i = d['id'][0]
            r = fin0.get(i)
            if r and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO' \
                    and d['cor'][0] is not None:
                cand.append((rn, d['linha'], i))
        if len(cand) < 10:
            raise RuntimeError(f'pré-condição ausente: menos de 10 pontos N1 do interfaceamento no LJ ({e0["analito"]}): '
                               f'{len(cand)}')
        rn_x, _, X = cand[len(cand) // 2]
        x0 = fin0[X]
        v_x = float(x0['RESULTADO'])
        v_m = round(v_x * 1.03 + 0.01, 4)
        g_x = grupo(x0)

        # ------------------------------------------------------------ M01 estrutura
        lo = q.lo(TB_MAN)
        cab = [c.Name for c in lo.ListColumns]
        col_s = lo.ListColumns(COL).DataBodyRange if COL in cab else None
        hr = lo.HeaderRowRange
        cor_cab = (hr.Cells(1, cab.index(COL) + 1).Interior.Color if COL in cab else None,
                   hr.Cells(1, cab.index('DATA') + 1).Interior.Color)
        topo = ' '.join(str(q.ws('Digitar Resultados').Range(f'A{i}').Value or '') for i in (3, 4))
        reg('M01 Estrutura: SUBSTITUI_ID é a última coluna da tblResultados_Manuais, formato texto, destravada, cabeçalho '
            'na cor das colunas de digitação; o topo da aba explica como substituir',
            cab[-1:] == [COL] and col_s is not None and col_s.NumberFormat == '@' and col_s.Locked is False
            and cor_cab[0] == cor_cab[1] and 'SUBSTITUI_ID' in topo,
            {'colunas': cab, 'formato': col_s.NumberFormat if col_s is not None else None,
             'travada': col_s.Locked if col_s is not None else None, 'cor_cabecalho': cor_cab, 'topo': topo[:300]})

        # ------------------------------------------------------------ M02 preenchimento ao digitar o SUBSTITUI_ID
        r_m = escrever_manual(q, {COL: curto(X)})
        lin = linha_manual(q, r_m)
        dhx = x0['DATA_HORA']
        M = str(lin.get('ID_REGISTRO') or '')
        ok_hora = isinstance(lin.get('HORA'), (float, dt.datetime))
        hora_lida = lin.get('HORA')
        if isinstance(hora_lida, dt.datetime):
            hora_lida = (hora_lida - hora_lida.replace(hour=0, minute=0, second=0, microsecond=0)).total_seconds() / 86400.0
        ok2 = (M.startswith('MAN_') and txt(lin.get(COL)) == X and txt(lin.get('EQUIPAMENTO')) == txt(x0['EQUIPAMENTO'])
               and txt(lin.get('MATRIZ')) == txt(x0['MATRIZ']) and txt(lin.get('LOTE')) == txt(x0['LOTE'])
               and iint(lin.get('NIVEL')) == iint(x0['NIVEL']) and txt(lin.get('ANALITO')) == txt(x0['ANALITO'])
               and dia(lin.get('DATA')) == dhx.date() and ok_hora
               and abs(float(hora_lida) - serial_hora(dhx)) < 1.5 / 86400 and lin.get('USUARIO'))
        gravar_celula(q, r_m, 'RESULTADO', v_m)
        gravar_celula(q, r_m, 'MOTIVO', MOT)
        # ID inexistente: nada preenchido (e, com Interactive = False, nenhuma caixa trava a automacao)
        r_n = escrever_manual(q, {COL: '987654321'})
        lin_n = linha_manual(q, r_n)
        vazio_n = all(lin_n.get(c) in (None, '') for c in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'NIVEL', 'ANALITO', 'DATA',
                                                             'HORA'))
        gravar_celula(q, r_n, COL, None)
        reg('M02 Digitar o SUBSTITUI_ID (só o número) preenche EQUIPAMENTO, MATRIZ, LOTE, NÍVEL, ANALITO, DATA e HORA do '
            'resultado indicado, normaliza o ID e gera o MAN_; ID inexistente não preenche nada',
            ok2 and vazio_n,
            {'linha': {k: lin.get(k) for k in ('ID_REGISTRO', COL, 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO',
                                               'DATA', 'HORA', 'USUARIO')},
             'automatico': resumo(x0), 'inexistente_vazio': vazio_n})

        # ------------------------------------------------------------ ATUALIZAR 1: a substituicao
        r_aud1 = audit_ult(q)
        atualizar('M03_atualizar')
        fin1 = q.final_por_id()
        qa1 = q.ler(TB_QA)
        aud1 = audit_desde(q, r_aud1)
        m1, x1 = fin1.get(M), fin1.get(X)
        ok3 = (m1 is not None and m1['STATUS_ANALITICO'] == 'ATIVO' and m1['PARTICIPA_ESTATISTICA'] == SIM
               and m1['TIPO_PLOTAGEM_LJ'] == 'NORMAL' and m1['ID_RELACIONADO'] == X
               and m1['MOTIVO_EXCLUSAO_AUTOMATICA'] == f'substitui o resultado do interfaceamento {X}'
               and x1 is not None and x1['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and x1['PARTICIPA_ESTATISTICA'] == NAO
               and x1['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and iint(x1['RUN']) is None and iint(x1['CORRIDA_NO_DIA']) is None
               and x1['ID_RELACIONADO'] == M
               and x1['MOTIVO_EXCLUSAO_AUTOMATICA'] == f'substituido pelo manual {M} (motivo: {MOT})')
        run_igual = m1 is not None and (iint(m1['RUN']), iint(m1['CORRIDA_NO_DIA'])) == \
            (iint(x0['RUN']), iint(x0['CORRIDA_NO_DIA']))
        r0_, r1_ = runs(fin0), runs(fin1)
        outros = [i for i in r0_ if i not in (X, M) and r1_.get(i) != r0_[i]]
        reg('M03 Substituição válida: manual ATIVO/SIM/NORMAL apontando o automático; automático SUBSTITUIDO_MANUAL, NÃO, '
            'NAO_PLOTAR, sem RUN e sem corrida, apontando o manual, motivo com o manual e o MOTIVO; o manual fica com o RUN '
            'e a corrida do dia do automático; nenhum outro RUN muda',
            ok3 and run_igual and not outros,
            {'manual': resumo(m1), 'automatico_antes': resumo(x0), 'automatico_depois': resumo(x1),
             'run_igual': run_igual, 'outros_runs_mudaram': outros[:10]})

        # ------------------------------------------------------------ M04 grafico
        e1 = cl.eng(q, nlv)
        orc = cl.oraculo(list(fin1.values()), e1['analito'], e1['lote'], e1['equip'], nlv)
        dif = cl.comparar(e1, orc, nlv)
        pr1 = L.eng_por_run()
        ids_n1 = [d['id'][0] for d in pr1.values()]
        seq = ativos_em_ordem(pr1, 0)
        reg('M04 Gráfico: corrida do lote = oráculo independente (só os ativos, sem lacuna) em todos os níveis; o manual '
            'aparece no N1 e o automático substituído não',
            not dif and consecutivos(seq) and M in ids_n1 and X not in ids_n1 and len(seq) == len(ativos_em_ordem(base, 0)),
            {'divergencias': dif[:10], 'n1_corridas': [len(ativos_em_ordem(base, 0)), len(seq)],
             'manual_no_grafico': M in ids_n1, 'automatico_no_grafico': X in ids_n1})

        # ------------------------------------------------------------ M05 QA
        i04 = achados(qa1, 'I04', X)
        det = str(i04[0]['DETALHE']) if i04 else ''
        a06m = achados(qa1, 'A06', M)
        ok_a06m = all(str(x['DETALHE']).startswith(f'este manual substitui {X};') and X not in listados_a06(x['DETALHE'])
                      and 'SUBSTITUI_ID' not in str(x['DETALHE']) for x in a06m)
        reg('M05 QA: I04 do automático substituído (o manual, os dois valores, a diferença e o motivo); nenhum E12; '
            'nenhum A01 do manual; o A06 do manual (se houver) é o de quem substitui e não lista o automático',
            len(i04) == 1 and f'substituido pelo manual {M} (motivo: {MOT})' in det and 'vale o manual (' in det
            and 'este resultado (' in det and 'Diferenca manual - automatico: ' in det
            and 'fora do calculo e do grafico' in det and det.endswith('apague o SUBSTITUI_ID (ou a linha) do manual.')
            and not achados(qa1, 'E12') and not achados(qa1, 'A01', M) and ok_a06m,
            {'I04': i04, 'E12': achados(qa1, 'E12')[:3], 'A01_manual': achados(qa1, 'A01', M), 'A06_manual': a06m})

        # ------------------------------------------------------------ M06 Audit_Log
        ev_x = eventos(aud1, 'RESULTADO_SUBSTITUIDO_POR_MANUAL', X)
        ev_m = eventos(aud1, 'RESULTADO_MANUAL_INCLUIDO', M)
        reg('M06 Audit_Log: RESULTADO_SUBSTITUIDO_POR_MANUAL do automático (antes = valor dele, depois = valor do manual, '
            'RUN = o do manual) e RESULTADO_MANUAL_INCLUIDO do manual',
            len(ev_x) == 1 and perto(ev_x[0]['res_ant'], v_x) and perto(ev_x[0]['res_novo'], v_m)
            and iint(ev_x[0]['run']) == iint(m1['RUN']) and len(ev_m) == 1,
            {'automatico': ev_x, 'manual': ev_m, 'acoes': [x['acao'] for x in aud1][-12:]})

        # ------------------------------------------------------------ M07 dica e janela de inativar
        s_res, _ = series(q, 1)
        lm = next((d['linha'] for d in pr1.values() if d['id'][0] == M), None)
        dica_m = run('mUI.DicaPonto', 1, s_res, lm + 1, teto=60) if lm is not None and s_res else ''
        partes_m = dica_m.split(SEP)
        esperado_id = f'ID {M} (substitui {curto(X)})'
        outro = next(((d['linha'], d['id'][0]) for d in pr1.values() if d['id'][0] and d['id'][0] != M
                      and not str(d['id'][0]).startswith('MAN_')), (None, None))
        dica_o = run('mUI.DicaPonto', 1, s_res, outro[0] + 1, teto=60) if outro[0] is not None else ''
        efeito = run('mUI.EfeitoInativar', M, teto=60)
        efeito_o = run('mUI.EfeitoInativar', str(outro[1] or ''), teto=60)
        reg('M07 Dica do ponto do manual: "ID MAN_x (substitui <ID>)" dentro da parte do ID (mesmo número de partes); a '
            'janela de inativar do manual mantém o efeito de sempre e avisa que inativá-lo NÃO traz o automático de volta, '
            'num texto que cabe no rótulo (<= 150 caracteres; o padrão tem 127); ponto do interfaceamento: dica e janela '
            'como antes',
            esperado_id in partes_m and len(partes_m) == 5 and 'substitui' not in dica_o and len(dica_o.split(SEP)) == 5
            and f'NÃO traz o automático {curto(X)} de volta' in efeito and 'sai da média' in efeito
            and len(efeito) <= 150 and efeito_o == '',
            {'dica_manual': dica_m, 'esperado': esperado_id, 'dica_outro': dica_o, 'efeito': efeito,
             'efeito_outro': efeito_o})

        # ------------------------------------------------------------ ATUALIZAR 2: edicao do manual e os casos invalidos
        alvos = [r for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO'
                 and isinstance(r['DATA_HORA'], dt.datetime) and grupo(r) != g_x and iint(r['NIVEL']) == 1
                 and r['ANALITO_CADASTRADO'] == SIM]
        alvos.sort(key=lambda r: (r['DATA_HORA'], str(r['ID_REGISTRO'])))
        if len(alvos) < 60:
            raise RuntimeError(f'pré-condição ausente: poucos resultados N1 do interfaceamento ({len(alvos)})')
        usados, escolhidos = set(), []
        for k in range(1, 10):                         # grupos distintos entre si (um caso nao mexe no outro)
            for r in alvos[len(alvos) * k // 10:]:
                if grupo(r) not in usados:
                    usados.add(grupo(r))
                    escolhidos.append(r)
                    break
        if len(escolhidos) < 6:
            raise RuntimeError(f'pré-condição ausente: menos de 6 grupos distintos com N1 do interfaceamento '
                               f'({len(escolhidos)})')
        Y2, Y3, Z, W, V, U = escolhidos[:6]
        v_m2 = round(v_m + 0.5, 4)
        gravar_celula(q, r_m, 'RESULTADO', v_m2)                                     # M08
        # M10a ID inexistente (chave e hora de Y3, escritas antes)
        dh3 = Y3['DATA_HORA']
        campos3 = {'EQUIPAMENTO': Y3['EQUIPAMENTO'], 'MATRIZ': Y3['MATRIZ'], 'LOTE': txt(Y3['LOTE']), 'NIVEL': 1,
                   'ANALITO': Y3['ANALITO'], 'DATA': dt.datetime(dh3.year, dh3.month, dh3.day),
                   'HORA': serial_hora(dh3 + dt.timedelta(minutes=37)), 'RESULTADO': 1.5, 'MOTIVO': MOT}
        r_a = escrever_manual(q, dict(campos3, **{COL: '987654321'}))
        # M10b chave diferente: EQUIPAMENTO/MATRIZ de Y2 e NIVEL 2 digitados antes do SUBSTITUI_ID (o resto vem de Y2;
        # sem o equipamento digitado, o padrao do CFG entraria primeiro e a chave poderia cair no grupo do X)
        r_b = escrever_manual(q, {'EQUIPAMENTO': Y2['EQUIPAMENTO'], 'MATRIZ': Y2['MATRIZ'], 'NIVEL': 2,
                                  COL: curto(Y2['ID_REGISTRO']), 'RESULTADO': 2.5, 'MOTIVO': MOT})
        # M10c ID de um manual (o M)
        r_c = escrever_manual(q, dict(campos3, HORA=serial_hora(dh3 + dt.timedelta(minutes=77)), **{COL: M}))
        # M10d retransmissao (se houver)
        dup = next((r for r in fin0.values() if r['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM'
                    and isinstance(r['DATA_HORA'], dt.datetime)), None)
        r_d = escrever_manual(q, {COL: dup['ID_REGISTRO'], 'RESULTADO': 3.5, 'MOTIVO': MOT}) if dup else None
        # M10e dois manuais para o mesmo Z: o 1o (linha menor) vale; o 2o, 2 min depois, e CONFLITO_MANUAL. O SUBSTITUI_ID
        # vem antes (preenche a linha, inclusive EQUIPAMENTO/MATRIZ); a HORA corrigida depois e edicao de quem digita
        r_e1 = escrever_manual(q, {COL: curto(Z['ID_REGISTRO']), 'RESULTADO': 4.5, 'MOTIVO': MOT})
        r_e2 = escrever_manual(q, {COL: curto(Z['ID_REGISTRO']),
                                   'HORA': serial_hora(Z['DATA_HORA'] + dt.timedelta(minutes=2)),
                                   'RESULTADO': 4.6, 'MOTIVO': MOT})
        # M16 o 1o manual (o que vale) inativado -- o MAN_ ainda nao esta na Principal: o ANALITO vai digitado
        id_e1 = str(linha_manual(q, r_e1)['ID_REGISTRO'] or '')
        q.escrever_linha(TB_INAT, {'ID_REGISTRO': id_e1, 'ANALITO': Z['ANALITO'],
                                   'MOTIVO': 'QA ADR-082: manual substituto inativado'})
        # M11 W inativado (aba Inativar, eventos ligados: analito e caixa sozinhos) e substituido
        q.escrever_linha(TB_INAT, {'ID_REGISTRO': curto(W['ID_REGISTRO']), 'MOTIVO': 'QA ADR-082: inativado e substituido'})
        r_w = escrever_manual(q, {COL: curto(W['ID_REGISTRO']), 'RESULTADO': 5.5, 'MOTIVO': MOT})
        # M12 sem MOTIVO
        r_v = escrever_manual(q, {COL: curto(V['ID_REGISTRO']), 'RESULTADO': 6.5})
        # M13 SEM_VALOR (se houver)
        sv = next((r for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                   and r['STATUS_ANALITICO'] == 'SEM_VALOR' and isinstance(r['DATA_HORA'], dt.datetime)
                   and grupo(r) != g_x and grupo(r) not in usados), None)
        r_s = escrever_manual(q, {COL: sv['ID_REGISTRO'], 'RESULTADO': 7.5, 'MOTIVO': MOT}) if sv else None
        # M18 U substituido agora; inativado no ATUALIZAR 3
        r_u = escrever_manual(q, {COL: curto(U['ID_REGISTRO']), 'RESULTADO': 8.5, 'MOTIVO': MOT})
        # M19 SEM_DATA_HORA (se houver; com DATA -- o preenchimento traz so a data, a HORA vai digitada)
        sdh = next((r for r in fin0.values() if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and r['STATUS_ANALITICO'] == 'SEM_DATA_HORA' and r['INATIVACAO_REGISTRADA'] != SIM
                    and isinstance(r.get('DATA'), dt.datetime) and grupo_dia(r) != g_x and grupo_dia(r) not in usados
                    and r['ANALITO_CADASTRADO'] == SIM), None)
        r_t = escrever_manual(q, {COL: sdh['ID_REGISTRO'], 'HORA': 0.5, 'RESULTADO': 9.5, 'MOTIVO': MOT}) if sdh else None
        ids = {k: str(linha_manual(q, r)['ID_REGISTRO'] or '') for k, r in
               (('a', r_a), ('b', r_b), ('c', r_c), ('d', r_d), ('e1', r_e1), ('e2', r_e2), ('w', r_w), ('v', r_v),
                ('s', r_s), ('u', r_u), ('t', r_t)) if r}
        r_aud2 = audit_ult(q)
        atualizar('M08_atualizar')
        fin2 = q.final_por_id()
        qa2 = q.ler(TB_QA)
        aud2 = audit_desde(q, r_aud2)

        # ------------------------------------------------------------ M08 edicao do valor do manual
        ev8 = eventos(aud2, 'RESULTADO_MANUAL_ALTERADO', M)
        reg('M08 Editar o RESULTADO do manual que substitui: RESULTADO_MANUAL_ALTERADO na trilha (antes -> depois); a '
            'substituição continua',
            len(ev8) == 1 and perto(ev8[0]['res_ant'], v_m) and perto(ev8[0]['res_novo'], v_m2)
            and fin2[X]['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and perto(fin2[M]['RESULTADO'], v_m2),
            {'eventos': ev8, 'automatico': resumo(fin2.get(X)), 'manual': resumo(fin2.get(M))})

        # ------------------------------------------------------------ M10 invalidos
        def e12(k):
            return achados(qa2, 'E12', ids.get(k))

        def nota(k):
            return str((fin2.get(ids.get(k)) or {}).get('MOTIVO_EXCLUSAO_AUTOMATICA') or '')
        casos = {
            'inexistente': bool(e12('a')) and 'nao existe nos resultados do interfaceamento' in str(e12('a')[0]['DETALHE'])
            and 'SUBSTITUI_ID invalido: ' in nota('a'),
            'chave_diferente': bool(e12('b')) and 'tem outra chave' in str(e12('b')[0]['DETALHE'])
            and fin2[Y2['ID_REGISTRO']]['STATUS_ANALITICO'] == 'ATIVO',
            'id_de_manual': bool(e12('c')) and 'e resultado manual' in str(e12('c')[0]['DETALHE']),
            'retransmissao': dup is None or (bool(e12('d')) and 'e retransmissao de' in str(e12('d')[0]['DETALHE'])
                                             and fin2[dup['ID_REGISTRO']]['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM'),
            'dois_manuais': fin2[Z['ID_REGISTRO']]['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL'
            and fin2[Z['ID_REGISTRO']]['ID_RELACIONADO'] == ids['e1']
            and fin2[ids['e1']]['STATUS_ANALITICO'] == 'INATIVADO'          # M16: inativado, continua substituindo
            and fin2[ids['e2']]['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and fin2[ids['e2']]['ID_RELACIONADO'] == ids['e1']
            and f'outro manual ({ids["e1"]}) ja substitui {Z["ID_REGISTRO"]}' in nota('e2')
            and bool(e12('e2')) and bool(achados(qa2, 'A01', ids['e2'])),
        }
        subst2 = sorted(i for i, r in fin2.items() if r['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL')
        esperado_sub = sorted([X, Z['ID_REGISTRO'], U['ID_REGISTRO']] + ([sv['ID_REGISTRO']] if sv else [])
                              + ([sdh['ID_REGISTRO']] if sdh else []))
        invalidos = [ids[k] for k in ('a', 'b', 'c', 'd', 'e2') if k in ids]
        reg('M10 SUBSTITUI_ID inválido: E12 com o motivo exato e NADA substituído -- ID inexistente, chave diferente '
            '(outro nível), ID de um manual, retransmissão (se houver), 2º manual para o mesmo ID (CONFLITO_MANUAL '
            '"outro manual (...) já substitui ...", A01; o 1º vale)',
            all(casos.values()) and subst2 == esperado_sub
            and all(not str((fin2.get(i) or {}).get('ID_RELACIONADO') or '').startswith(('BIO-', 'HEM-'))
                    or fin2[i]['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' for i in invalidos),
            {'casos': casos, 'substituidos': subst2, 'esperado': esperado_sub,
             'E12': [(x['ID_REGISTRO'], x['DETALHE']) for x in achados(qa2, 'E12')][:8],
             'manuais': {k: resumo(fin2.get(i)) for k, i in ids.items()}})

        # ------------------------------------------------------------ M11 INATIVADO continua INATIVADO
        w2, mw = fin2.get(W['ID_REGISTRO']), fin2.get(ids['w'])
        i04w = achados(qa2, 'I04', W['ID_REGISTRO'])
        r2_ = runs(fin2)
        # sem corrida propria: o RUN/corrida do manual, e nenhum outro RUN do grupo do W muda (antes da revisao o
        # inativado ficava nos reais e o manual, no mesmo instante, abria uma posicao a mais: as seguintes andavam +1)
        mesmo_run_w = (w2 is not None and mw is not None and iint(w2['RUN']) is not None
                       and (iint(w2['RUN']), iint(w2['CORRIDA_NO_DIA'])) == (iint(mw['RUN']), iint(mw['CORRIDA_NO_DIA'])))
        mudou_w = [i for i, r in fin0.items() if i != W['ID_REGISTRO'] and isinstance(r['DATA_HORA'], dt.datetime)
                   and grupo(r) == grupo(W) and r2_.get(i) != r0_[i]]
        a06w = achados(qa2, 'A06', ids['w'])
        ef_r = run('mUI.EfeitoReativar', W['ID_REGISTRO'], teto=60)
        ef_r_o = run('mUI.EfeitoReativar', Y2['ID_REGISTRO'], teto=60)
        reg('M11 Interfaceamento INATIVADO indicado no SUBSTITUI_ID: continua INATIVADO (X vermelho, GOVERNANÇA OK, sem '
            'E01 nem A04) e o manual vale; vínculo nos dois sentidos; sem corrida própria (o RUN e a corrida do manual) e '
            'nenhum outro RUN do grupo muda; o I04 diz que a inativação prevalece e que o automático só volta reativado; '
            'nenhum A06 do manual lista o inativado; a janela de reativar o X avisa que ele continua fora',
            w2 is not None and w2['STATUS_ANALITICO'] == 'INATIVADO' and w2['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
            and str(w2['GOVERNANCA']) == 'OK' and w2['ID_RELACIONADO'] == ids['w'] and mesmo_run_w and not mudou_w
            and mw is not None and mw['STATUS_ANALITICO'] == 'ATIVO' and mw['ID_RELACIONADO'] == W['ID_REGISTRO']
            and not achados(qa2, 'E01', W['ID_REGISTRO']) and not achados(qa2, 'A04', W['ID_REGISTRO'])
            and len(i04w) == 1 and 'continua INATIVADO (a inativacao prevalece' in str(i04w[0]['DETALHE'])
            and 'reative-o (aba Inativar)' in str(i04w[0]['DETALHE'])
            and all(W['ID_REGISTRO'] not in listados_a06(x['DETALHE']) for x in a06w)
            and ids['w'] in ef_r and 'continua fora' in ef_r and len(ef_r) <= 150 and ef_r_o == '',
            {'automatico': resumo(w2), 'manual': resumo(mw), 'mesmo_run': mesmo_run_w, 'grupo_mudou': mudou_w[:10],
             'I04': i04w, 'A06_manual': a06w, 'efeito_reativar': ef_r, 'efeito_reativar_outro': ef_r_o})

        # ------------------------------------------------------------ M16 o 1o manual inativado continua substituindo
        e12_e2 = achados(qa2, 'E12', ids['e2'])
        det16 = str(e12_e2[0]['DETALHE']) if e12_e2 else ''
        reg('M16 O 1º manual (o que vale) INATIVADO continua substituindo: o E12 do 2º diz isso e que a correção é apagar o '
            'SUBSTITUI_ID do 1º; nenhum A09 do 2º (quem está inativado é um manual, não o automático)',
            len(e12_e2) == 1 and f'O manual {ids["e1"]} esta INATIVADO, mas continua substituindo' in det16
            and f'apague o SUBSTITUI_ID do {ids["e1"]}' in det16 and 'Corrija o SUBSTITUI_ID' not in det16
            and not achados(qa2, 'A09', ids['e2']),
            {'E12': e12_e2, 'A09': achados(qa2, 'A09', ids['e2']), 'primeiro': resumo(fin2.get(ids['e1'])),
             'segundo': resumo(fin2.get(ids['e2']))})

        # ------------------------------------------------------------ M19 SEM_DATA_HORA substituido: sem E09
        if sdh:
            t2, mt = fin2.get(sdh['ID_REGISTRO']), fin2.get(ids['t'])
            reg('M19 SEM_DATA_HORA do interfaceamento indicado: SUBSTITUIDO_MANUAL, sem RUN, o manual ATIVO; o ERRO E09 '
                'sai (o usuário já resolveu: o I04 registra)',
                t2 is not None and t2['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and iint(t2['RUN']) is None
                and t2['ID_RELACIONADO'] == ids['t'] and mt is not None and mt['STATUS_ANALITICO'] == 'ATIVO'
                and mt['ID_RELACIONADO'] == sdh['ID_REGISTRO'] and not achados(qa2, 'E09', sdh['ID_REGISTRO'])
                and len(achados(qa2, 'I04', sdh['ID_REGISTRO'])) == 1,
                {'automatico': resumo(t2), 'manual': resumo(mt), 'E09': achados(qa2, 'E09', sdh['ID_REGISTRO']),
                 'I04': achados(qa2, 'I04', sdh['ID_REGISTRO'])})
        else:
            RES.append({'teste': 'M19 SEM_DATA_HORA substituído -- a base não tem resultado SEM_DATA_HORA do '
                                 'interfaceamento com DATA', 'resultado': 'INFO', 'evidencia': {}})

        # ------------------------------------------------------------ M12 sem MOTIVO
        mv = fin2.get(ids['v'])
        reg('M12 SUBSTITUI_ID sem MOTIVO: MANUAL_INCOMPLETO "MOTIVO (obrigatorio para substituir um resultado)" + E04; o '
            'automático continua ATIVO',
            mv is not None and mv['STATUS_ANALITICO'] == 'MANUAL_INCOMPLETO'
            and 'MOTIVO (obrigatorio para substituir um resultado)' in str(mv['MOTIVO_EXCLUSAO_AUTOMATICA'])
            and bool(achados(qa2, 'E04', ids['v'])) and fin2[V['ID_REGISTRO']]['STATUS_ANALITICO'] == 'ATIVO'
            and not achados(qa2, 'E12', ids['v']),
            {'manual': resumo(mv), 'automatico': resumo(fin2.get(V['ID_REGISTRO']))})

        # ------------------------------------------------------------ M13 SEM_VALOR
        if sv:
            s2, ms = fin2.get(sv['ID_REGISTRO']), fin2.get(ids['s'])
            reg('M13 SEM_VALOR do interfaceamento indicado (a falha típica): SUBSTITUIDO_MANUAL, sem RUN; o manual ATIVO',
                s2 is not None and s2['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and iint(s2['RUN']) is None
                and ms is not None and ms['STATUS_ANALITICO'] == 'ATIVO' and ms['ID_RELACIONADO'] == sv['ID_REGISTRO'],
                {'automatico': resumo(s2), 'manual': resumo(ms)})
        else:
            RES.append({'teste': 'M13 SEM_VALOR substituído -- a base não tem resultado SEM_VALOR do interfaceamento',
                        'resultado': 'INFO', 'evidencia': {}})

        # ------------------------------------------------------------ ATUALIZAR 3: desfazer (apagar o SUBSTITUI_ID)
        gravar_celula(q, r_m, COL, None)
        gravar_celula(q, r_e1, COL, None)                                            # M17 (seguir o E12 do M16)
        q.escrever_linha(TB_INAT, {'ID_REGISTRO': curto(U['ID_REGISTRO']),           # M18 inativar o ja substituido
                                   'MOTIVO': 'QA ADR-082: substituido e depois inativado'})
        r_aud3 = audit_ult(q)
        atualizar('M09_atualizar')
        fin3 = q.final_por_id()
        qa3 = q.ler(TB_QA)
        aud3 = audit_desde(q, r_aud3)
        x3, m3 = fin3.get(X), fin3.get(M)
        r3_ = runs(fin3)
        grupo_x = [i for i, r in fin0.items() if i not in (X, M) and isinstance(r['DATA_HORA'], dt.datetime)
                   and grupo(r) == g_x]
        mudou_grupo = [i for i in grupo_x if r3_.get(i) != r0_[i]]
        ev9 = eventos(aud3, 'SUBSTITUICAO_POR_MANUAL_DESFEITA', X)
        ev9m = eventos(aud3, 'RESULTADO_MANUAL_STATUS', M)
        reg('M09 Desfazer (apagar o SUBSTITUI_ID): o automático volta ATIVO com o RUN e a corrida de antes, sem vínculo; o '
            'manual volta às regras sem SUBSTITUI_ID (CONFLITO_MANUAL pelo horário, A01); a corrida do dia é a de antes; '
            'SUBSTITUICAO_POR_MANUAL_DESFEITA (antes = valor do manual, depois = do automático) e RESULTADO_MANUAL_STATUS',
            x3 is not None and x3['STATUS_ANALITICO'] == 'ATIVO' and (iint(x3['RUN']), iint(x3['CORRIDA_NO_DIA'])) == r0_[X]
            and x3['ID_RELACIONADO'] is None and x3['MOTIVO_EXCLUSAO_AUTOMATICA'] is None
            and m3 is not None and m3['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and m3['ID_RELACIONADO'] == X
            and str(m3['MOTIVO_EXCLUSAO_AUTOMATICA']).startswith('coincide com resultado do interfaceamento')
            and bool(achados(qa3, 'A01', M)) and not achados(qa3, 'I04', X) and not mudou_grupo
            and len(ev9) == 1 and perto(ev9[0]['res_ant'], v_m2) and perto(ev9[0]['res_novo'], v_x) and len(ev9m) == 1,
            {'automatico': resumo(x3), 'manual': resumo(m3), 'grupo_do_dia_mudou': mudou_grupo[:10], 'desfeita': ev9,
             'status_manual': ev9m})

        # ------------------------------------------------------------ M17 o 2o manual passa a substituir
        z3, e1_3, e2_3 = fin3.get(Z['ID_REGISTRO']), fin3.get(ids['e1']), fin3.get(ids['e2'])
        ev17 = eventos(aud3, 'SUBSTITUICAO_POR_MANUAL_ALTERADA', Z['ID_REGISTRO'])
        a01_e1 = achados(qa3, 'A01', ids['e1'])
        reg('M17 Apagar o SUBSTITUI_ID do 1º manual: o 2º passa a substituir -- SUBSTITUICAO_POR_MANUAL_ALTERADA na trilha '
            '(antes = valor do 1º, depois = do 2º, com quem lançou o manual); o 1º vira CONFLITO_MANUAL com o automático e '
            'o A01 dele diz que vale o 2º (não "informe o ID em SUBSTITUI_ID"), sem A09',
            z3 is not None and z3['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and z3['ID_RELACIONADO'] == ids['e2']
            and e2_3 is not None and e2_3['STATUS_ANALITICO'] == 'ATIVO' and e2_3['ID_RELACIONADO'] == Z['ID_REGISTRO']
            and e1_3 is not None and e1_3['STATUS_ANALITICO'] == 'CONFLITO_MANUAL'
            and e1_3['ID_RELACIONADO'] == Z['ID_REGISTRO']
            and str(e1_3['MOTIVO_EXCLUSAO_AUTOMATICA']).startswith('coincide com resultado do interfaceamento')
            and len(ev17) == 1 and perto(ev17[0]['res_ant'], 4.5) and perto(ev17[0]['res_novo'], 4.6)
            and 'manual lancado' in ev17[0]['parecer']
            and len(a01_e1) == 1
            and f'vale o manual {ids["e2"]}, que ja substitui o automatico {Z["ID_REGISTRO"]}' in str(a01_e1[0]['DETALHE'])
            and 'SUBSTITUI_ID' not in str(a01_e1[0]['DETALHE']) and not achados(qa3, 'A09', ids['e1']),
            {'automatico': resumo(z3), 'primeiro': resumo(e1_3), 'segundo': resumo(e2_3), 'alterada': ev17,
             'A01_primeiro': a01_e1, 'A09_primeiro': achados(qa3, 'A09', ids['e1'])})

        # ------------------------------------------------------------ M18 inativar depois um ja substituido
        u2, u3, mu3 = fin2.get(U['ID_REGISTRO']), fin3.get(U['ID_REGISTRO']), fin3.get(ids['u'])
        mudou_u = [i for i, r in fin2.items() if i != U['ID_REGISTRO'] and isinstance(r['DATA_HORA'], dt.datetime)
                   and grupo(r) == grupo(U) and r3_.get(i) != r2_.get(i)]
        ev18 = eventos(aud3, 'RESULTADO_INATIVADO', U['ID_REGISTRO'])
        ev18x = [x for x in aud3 if x['motivo'] == f'ID {U["ID_REGISTRO"]}'
                 and any(a in x['acao'].split('+') for a in ('RESULTADO_REATIVADO', 'PLOTAGEM_LJ_ALTERADA'))]
        reg('M18 Inativar DEPOIS um automático já SUBSTITUIDO_MANUAL: INATIVADO/X_VERMELHO com o RUN e a corrida do manual '
            '(sem corrida própria) e nenhum outro RUN do grupo muda; RESULTADO_INATIVADO na trilha, sem REATIVADO/PLOTAGEM',
            u2 is not None and u2['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL' and iint(u2['RUN']) is None
            and u3 is not None and u3['STATUS_ANALITICO'] == 'INATIVADO' and u3['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
            and u3['ID_RELACIONADO'] == ids['u'] and mu3 is not None and iint(u3['RUN']) is not None
            and (iint(u3['RUN']), iint(u3['CORRIDA_NO_DIA'])) == (iint(mu3['RUN']), iint(mu3['CORRIDA_NO_DIA']))
            and not mudou_u and len(ev18) == 1 and not ev18x,
            {'antes': resumo(u2), 'depois': resumo(u3), 'manual': resumo(mu3), 'grupo_mudou': mudou_u[:10],
             'inativado': ev18, 'outros_eventos': ev18x})

        # ------------------------------------------------------------ M14 idempotencia
        cq3 = tabela_v2(q, TB_FIN)
        r_aud4 = audit_ult(q)
        atualizar('M14_atualizar')
        cq4 = tabela_v2(q, TB_FIN)
        aud4 = [x for x in audit_desde(q, r_aud4) if x['acao'].startswith(('RESULTADO_', 'SUBSTITUICAO_', 'PLOTAGEM_'))]
        difs = [i for i in set(cq3) | set(cq4) if cq3.get(i) != cq4.get(i)]
        reg('M14 Idempotência: ATUALIZAR de novo deixa a Principal - Resultados idêntica e não grava nenhuma linha de '
            'resultado na trilha',
            not difs and not aud4, {'ids_diferentes': difs[:10], 'trilha_nova': aud4[:10]})

        # ------------------------------------------------------------ M15 PQ x VBA
        cq_vba = tabela_v2(q, TB_FIN, ('RECEBIDO_EM',))
        qa_vba = qa_multiconjunto(q)
        q.cfg('MOTOR_DADOS', 'PQ')
        try:
            atualizar('M15_atualizar_PQ')
            cq_pq = tabela_v2(q, TB_FIN, ('RECEBIDO_EM',))
            qa_pq = qa_multiconjunto(q)
        finally:
            q.cfg('MOTOR_DADOS', 'VBA')
        dif_cq = [i for i in set(cq_vba) | set(cq_pq) if cq_vba.get(i) != cq_pq.get(i)]
        so_vba, so_pq = qa_vba - qa_pq, qa_pq - qa_vba
        # a equivalencia do QA inteiro e do equivalencia_qa.py (com a leitura da origem); aqui contam os codigos que o
        # ADR-082 mexe -- os outros vao como evidencia (no HISTORICO o PQ e o VBA leem a origem de jeitos diferentes)
        cods = ('E04', 'E09', 'E12', 'A01', 'A06', 'A09', 'I04')
        so_vba82 = [k for k in so_vba if k[1] in cods]
        so_pq82 = [k for k in so_pq if k[1] in cods]
        st_pq = collections.Counter(r['STATUS_ANALITICO'] for r in q.ler(TB_FIN))
        reg('M15 PQ x VBA: com todas as substituições acima, a Principal - Resultados do motor VBA = a do Power Query '
            '(MOTOR_DADOS = PQ), ID a ID, coluna a coluna; na tblQA_Integracao, os mesmos achados E04/E09/E12/A01/A06/A09/I04',
            not dif_cq and not so_vba82 and not so_pq82 and len(cq_pq) == len(cq_vba)
            and st_pq.get('SUBSTITUIDO_MANUAL', 0) >= 1,
            {'linhas': [len(cq_vba), len(cq_pq)], 'ids_diferentes': dif_cq[:10],
             'exemplo': {i: (cq_vba.get(i), cq_pq.get(i)) for i in dif_cq[:2]},
             'qa_so_vba_adr082': [list(map(str, k)) for k in so_vba82[:5]],
             'qa_so_pq_adr082': [list(map(str, k)) for k in so_pq82[:5]],
             'qa_outros_codigos_diferentes': sorted({k[1] for k in list(so_vba) + list(so_pq) if k[1] not in cods}),
             'status_pq': dict(st_pq)})
        q.cfg('MODO_FONTE', modo0)
        q.cfg('MOTOR_DADOS', motor0)
    except Exception as ex_geral:                    # noqa: BLE001
        reg('M99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:],
             'tempos_s': tempos})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('M00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1, 'tempos_s': tempos})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} manual substitui (ADR-082): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===',
          flush=True)
    return falhas


if __name__ == '__main__':
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
    sys.exit(1 if falhas else 0)
