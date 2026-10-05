# -*- coding: utf-8 -*-
"""QA do ETL de RECEBIMENTO com DB_SEAC SINTETICO (E01-E06 do plano) -- roda numa COPIA ja instalada.

Uso: python qa_etl_recebimento.py <Hematologia|Bioquimica> <copia.xlsm> [saida.json]

Prova o elo DB_SEAC -> tblDB_Recebimento -> tblCQ_Final COMPARANDO COM A ORIGEM, o que nenhum
teste anterior fazia (T02 so confere a unicidade). A origem e um DB_SEAC sintetico gerado aqui
(gerar_seac_sintetico.py) com o cabecalho LIDO da propria copia (colunas da tblDB_Recebimento
menos ID_REGISTRO e RECEBIDO_EM) e linhas clonadas de resultados REAIS ATIVO (analito e lote
cadastrados), com IDs 9000001+ (faixa que nao existe na base real), em dias sem dado real.
A copia passa a ler essa origem como o administrador faria: Cfg_Integracao MODO_FONTE = SEAC e
CAMINHO_DB_SEAC = <arquivo sintetico> (mesmo mecanismo do T20 do qa_final.py).

SEGURANCA
- O arquivo recebido por argumento NUNCA e aberto pelo Excel: o teste copia-o (shutil.copy2) para
  uma pasta temporaria ao lado dele e trabalha em duas copias descartaveis (A: E01-E05, B: E06,
  uma de cada vez -- so um Excel por vez). Nada e salvo (fechar(salvar=False)); no fim confere que o
  SHA-256 do arquivo de entrada nao mudou e apaga a pasta temporaria.
- Nenhuma caixa de dialogo: a atualizacao e mIntegracao.AtualizarDadosAutomatico (QA.atualizar);
  nunca mApp.NovoLote nem macro com InputBox/MsgBox. Nao grava NumberFormat; grava valores.

ATUALIZACOES (cada uma leva dezenas de segundos; os tempos vao para a evidencia e para o E00):
  copia A: L1 (S01-S05) | L2, L3, L4 (+S06-S09, S13: mesmo arquivo 3x) | L5 (origem alterada)
           | L6 (+S10, S11) | L7 (+S12 'ABC') | L8 (+S14/S15 DATA_HORA nula -- por ultimo, porque se
           a final quebrar o recebimento ja guardou a linha e toda atualizacao seguinte quebraria)
  copia B: B0 (linha de base HISTORICO) | E06a1..a3 (+1 refeita da linha de base por cenario que
           hoje passa calado) | E06b0 (SEAC sintetico) | E06b1 (falha rapida) | E06b2 (so a consulta)
  Se a carga L2 nao entrar por causa da linha sem ID, o teste tenta de novo sem ela (no maximo 2x)
  e diz isso na evidencia.

CONTRATO-ALVO (testes que HOJE devem falhar trazem "(hoje falha: prova Dnn)"): ver o plano
E01-E06 e os defeitos D01, D02, D03, D04, D13, D16.

Desvio consciente do plano, registrado na evidencia do E03: a colisao fracionaria e S09B
(ID_ORIGEM 9000009.4) contra S09A (9000009) NA MESMA CARGA, e nao contra S1 -- S1 entra na carga
L1 para que o E01 prove o recebimento sem a contaminacao do D01.
"""
import collections
import datetime as dt
import hashlib
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES  # noqa: E402  (qa_final poe pqlib/xlh no sys.path)
import pqlib  # noqa: E402
import gerar_seac_sintetico as gss  # noqa: E402

SIM = 'SIM'
NAO = 'NÃO'
SENHA = 'qcini2025'
BASE_ID = 9000000
CORRIDA_REAL = ('ATIVO', 'INATIVADO', 'SEM_VALOR')          # pq/DB_CQ_FINAL.m CORRIDA_REAL
TEMPOS = []
MOTIVO = {}
REGISTRADOS = set()

TITULOS = collections.OrderedDict([
    ('E01a', 'Recebimento por conjunto (DB_SEAC sintético): os IDs novos são exatamente {PREFIXO-9000001..9000005}'),
    ('E01b', 'Recebimento = origem célula a célula em TODAS as colunas do DB_SEAC (VALOR, DATA_HORA, ANALITO, LOTE, '
             'NIVEL, ITEM_ID, ...)'),
    ('E01c', 'Um único RECEBIDO_EM para a carga, preenchido'),
    ('E01d', 'Histórico pré-existente intacto: mesmo conjunto de linhas e mesmos valores (inclusive RECEBIDO_EM)'),
    ('E01e', 'Final: os 5 IDs uma vez cada, ATIVO/SIM/NORMAL/INTERFACEAMENTO, RESULTADO = VALOR, RUN = oráculo da regra '
             'de corrida'),
    ('E02a', 'P-02: linha da origem sem ID (S6 em branco, S6b nula) não é acumulada no recebimento nem recarimbada a '
             'cada atualização (hoje falha: prova D02)'),
    ('E02b', 'P-02: cada linha sem ID gera ERRO E06 "Linha da origem sem ID (rejeitada)" com DATA_HORA, ANALITO, '
             'ID_AMOSTRA e NIVEL no DETALHE (hoje falha: prova D02)'),
    ('E02c', 'P-02 conservação: |linhas da origem| = |já recebidas| + |IDs novos| + |repetições do mesmo ID| + '
             '|rejeitadas reportadas (E06)| (hoje falha: prova D02)'),
    ('E03a', 'ID repetido na mesma carga (9000007 iguais; 9000008 VALOR 5/500; 9000009 x 9000009.4): exatamente 1 '
             'linha por ID no recebimento, vale a de menor ITEM_ID (hoje falha: prova D01)'),
    ('E03b', 'Final: nunca 2 linhas com o mesmo ID_REGISTRO e nenhum ATIVO em dobro por (equip, lote, nível, analito, '
             'RUN) (hoje falha: prova D01)'),
    ('E03c', 'QA: ERRO E07 "ID repetido na origem" para cada ID repetido, e só para eles (hoje falha: prova D01)'),
    ('E04a', 'ID_ORIGEM em texto (" 9000010 " com espaços, "9000011"): viram PREFIXO-9000010/9000011, únicos e ATIVO'),
    ('E04b', 'ID_ORIGEM não numérico ("ABC"): a atualização falha com mensagem explícita OU a linha é rejeitada com '
             'ERRO no QA; nunca erro de célula silencioso nem perda muda'),
    ('E04c', 'Interface com DATA preenchida e DATA_HORA nula: a atualização conclui, a linha fica fora do RUN com ERRO '
             'E09 e o vizinho do grupo mantém o RUN correto (hoje falha: prova D16)'),
    ('E13c', 'Conflito manual × interface ATRAVESSA a meia-noite: interface S16 às 23:58 e manual com a mesma chave às '
             '00:02 do dia seguinte -> o manual vira CONFLITO_MANUAL com ID_RELACIONADO = S16 e fica fora da estatística '
             '(hoje falha: prova D09c)'),
    ('E05a', 'Retransmissão S13 (cópia de S2 com ID 9000013): DUPLICIDADE_ORIGEM, ID_RELACIONADO = PREFIXO-9000002, '
             'A02; S2 continua ATIVO'),
    ('E05b', 'Idempotência: 3 atualizações com o mesmo arquivo -> |Recebimento|, RECEBIDO_EM de cada ID e final '
             '(ID/RUN/STATUS) idênticos; RUN de S1-S5 igual ao da carga L1'),
    ('E05c', 'Origem altera o VALOR de S1 já recebido: o recebido e a final mantêm o valor original e o QA dá A03'),
    ('E05d', 'Origem altera só o ANALITO de S3 (VALOR igual): o recebido mantém e o QA dá A03 (hoje falha: prova D13)'),
    ('E06p', 'E06 pré-condição: cópia HISTORICO com 1 inativação justificada (INATIVADO) e 1 manual ATIVO'),
    ('E06a1', 'tblInativacao_NaoConformes renomeada: a atualização FALHA com ERRO| citando a tabela; motor (Eng_Saida '
              'G1) e final intactos (hoje falha: prova D03)'),
    ('E06a2', 'tblResultados_Manuais renomeada: a atualização FALHA com ERRO| citando a tabela; motor e final '
              'intactos (hoje falha: prova D03)'),
    ('E06a3', 'tblComentariosTecnicos renomeada: a atualização FALHA com ERRO| citando a tabela; motor e final '
              'intactos (hoje falha: prova D03)'),
    ('E06b1', 'MODO SEAC com tblDB_Recebimento renomeada: ATUALIZAR DADOS falha com ERRO| citando a tabela; motor, '
              'final e recebimento intactos (guarda atual do VBA)'),
    ('E06b2', 'MODO SEAC, camada Power Query: o refresh do recebimento sem a própria tabela (final não vazia) FALHA '
              'explícito em vez de reconstruir só com a janela do DB_SEAC (hoje falha: prova D04)'),
])
SESSAO = {k: ('B' if k.startswith('E06') else 'A') for k in TITULOS}


class PreCondicao(Exception):
    pass


class Interrompida(Exception):
    pass


def R(tid, ok, evid):
    REGISTRADOS.add(tid)
    reg(f'{tid} {TITULOS[tid]}', bool(ok), evid)


# =============================================================================
# utilidades
# =============================================================================
def sha256(caminho):
    h = hashlib.sha256()
    with open(caminho, 'rb') as f:
        for bloco in iter(lambda: f.read(1 << 20), b''):
            h.update(bloco)
    return h.hexdigest()


def semfuso(v):
    if isinstance(v, dt.datetime) and v.tzinfo is not None:
        return v.replace(tzinfo=None)
    return v


def norm(v):
    """Valor comparavel entre a origem (openpyxl) e o que o COM devolve da tabela."""
    if v is None:
        return None
    if isinstance(v, str):
        return v if v != '' else None
    if isinstance(v, bool):
        return v
    if isinstance(v, dt.datetime):
        v = semfuso(v)
        v = dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second, v.microsecond)
        return (v + dt.timedelta(microseconds=500000)).replace(microsecond=0)   # ruido de ms do COM
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        f = float(v)
        return int(f) if f.is_integer() else round(f, 9)
    return v


def dia_de(v):
    v = semfuso(v)
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def id_esperado(prefixo, v):
    """ID_REGISTRO que a regra-alvo atribui ao ID_ORIGEM gravado (None = sem ID valido)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, int):
        return f'{prefixo}-{v}'
    if isinstance(v, float):
        return f'{prefixo}-{int(round(v))}'               # Int64 do Power Query arredonda
    t = str(v).strip()
    return f'{prefixo}-{int(t)}' if t.isdigit() else None


def numero_id(i, prefixo):
    s = str(i or '')
    if s.startswith(prefixo + '-') and s[len(prefixo) + 1:].isdigit():
        return int(s[len(prefixo) + 1:])
    return None


def ids_de(linhas):
    return {r.get('ID_REGISTRO') for r in linhas if r.get('ID_REGISTRO') not in (None, '')}


def agrupar(linhas, campo='ID_REGISTRO'):
    d = collections.defaultdict(list)
    for r in linhas:
        d[r.get(campo)].append(r)
    return d


def ler_cols(q, nome, cols):
    """Le so as colunas pedidas (a final tem ~44 mil linhas x 39 colunas)."""
    lo = q.lo(nome)
    if lo is None:
        raise PreCondicao(f'tabela {nome} não encontrada')
    if lo.ListRows.Count == 0:
        return []
    colunas = []
    for c in cols:
        v = lo.ListColumns(c).DataBodyRange.Value
        if not isinstance(v, tuple):
            v = ((v,),)
        colunas.append([semfuso(x[0]) for x in v])
    return [dict(zip(cols, t)) for t in zip(*colunas)]


COLS_FINAL = ['ID_REGISTRO', 'ORIGEM_RESULTADO', 'EQUIPAMENTO', 'MATRIZ', 'DATA', 'DATA_HORA', 'CORRIDA_NO_DIA', 'RUN',
              'NIVEL', 'LOTE', 'ANALITO', 'ANALITO_ORIGEM', 'ANALITO_CADASTRADO', 'RESULTADO', 'STATUS_ANALITICO',
              'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'COMENTARIO_TECNICO', 'TEM_JUSTIFICATIVA', 'GOVERNANCA',
              'ID_RELACIONADO', 'ID_ORIGEM', 'ID_AMOSTRA', 'ITEM_ID']


def ler_final(q):
    return ler_cols(q, 'tblCQ_Final', COLS_FINAL)


def resumo_final(r):
    return {k: r.get(k) for k in ('ID_REGISTRO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ',
                                  'RUN', 'RESULTADO', 'ID_RELACIONADO', 'DATA_HORA', 'NIVEL')}


def estado_final(fin):
    return collections.Counter((r['ID_REGISTRO'], r['STATUS_ANALITICO'], r['PARTICIPA_ESTATISTICA'],
                                r['TIPO_PLOTAGEM_LJ'], r['RUN'], norm(r['RESULTADO']), r['COMENTARIO_TECNICO'],
                                r['GOVERNANCA']) for r in fin)


def diff_estado(a, b):
    sumiu, surgiu = a - b, b - a
    return {'linhas': [sum(a.values()), sum(b.values())], 'n_mudancas': sum(sumiu.values()) + sum(surgiu.values()),
            'antes': [list(x) for x in list(sumiu.elements())[:4]],
            'depois': [list(x) for x in list(surgiu.elements())[:4]]}


def carimbo_motor(q):
    return semfuso(q.ws('Eng_Saida').Range('G1').Value)


def textos_data(d):
    return {d.strftime('%d/%m/%Y'), d.strftime('%Y-%m-%d'), d.strftime('%d/%m/%y')}


def tentar(q, rotulo):
    """Uma atualizacao (ATUALIZAR DADOS pela entrada de automacao). Devolve (erro, resumo, segundos)."""
    t0 = time.time()
    try:
        resumo, _ = q.atualizar()
        erro = None
    except TimeoutError:
        raise                                             # Excel morto pelo watchdog: a sessao acabou
    except Exception as e:
        resumo, erro = None, str(e)
    t = round(time.time() - t0, 1)
    TEMPOS.append({'atualizacao': rotulo, 's': t, 'ok': erro is None})
    print(f'    atualização {rotulo}: {t}s ' + ('OK' if erro is None else 'ERRO: ' + erro[:300]), flush=True)
    return erro, resumo, t


def camadas_ok(erro):
    """A atualizacao passou por TODAS as consultas (tblDB_Recebimento .. tblQA_Integracao), mesmo que tenha
    falhado depois, no motor/auditoria (ex.: ID em dobro na final derrubando o motor -- D01). Assim um
    defeito nao mascara o teste de outro, e o erro vai sempre para a evidencia."""
    if erro is None:
        return True
    return any(s in erro for s in ('Etapa: conferir', 'Etapa: registrar a auditoria', 'Etapa: atualizar o motor'))


def com_teto(q, fn, teto=900):
    """Chamada COM longa com watchdog (mesma ideia do xlh.Excel.run)."""
    feito, estourou = threading.Event(), []

    def cao():
        if not feito.wait(teto):
            estourou.append(True)
            q.ex.matar()
    threading.Thread(target=cao, daemon=True).start()
    try:
        return fn()
    except Exception:
        if estourou:
            raise TimeoutError(f'chamada passou de {teto}s -- Excel encerrado')
        raise
    finally:
        feito.set()


def desproteger(ws):
    try:
        ws.Unprotect(SENHA)
    except Exception:
        pass
    return not ws.ProtectContents


def renomear(q, antigo, novo):
    lo = q.lo(antigo)
    if lo is None:
        raise PreCondicao(f'tabela {antigo} não encontrada para renomear')
    desproteger(lo.Parent)
    lo.Name = novo
    if q.lo(novo) is None:
        raise Interrompida(f'não renomeou {antigo} -> {novo}')


def oraculo_run(linhas, gap):
    """Oraculo independente da regra de corrida de pq/DB_CQ_FINAL.m (~330-370) para UM grupo
    (equipamento, matriz, lote, analito) num dia: bloco enquanto o intervalo <= GAP; dentro do bloco,
    o 2o resultado do mesmo nivel abre posicao nova; corrida = ordem de (bloco, posicao);
    RUN = yymmdd*100 + corrida. Linhas sem DATA_HORA ficam fora (contrato D16)."""
    def k_nulo(v):
        return (0, 0) if v is None else (1, v)
    rs = [r for r in linhas if r.get('DATA_HORA') is not None]
    ordem = sorted(rs, key=lambda r: (r['DATA_HORA'], int(r['NIVEL']), k_nulo(r.get('ITEM_ID')), str(r['ID_REGISTRO'])))
    bloco, b, ant = {}, -1, None
    for i, r in enumerate(ordem):
        if ant is None or (r['DATA_HORA'] - ant).total_seconds() / 60.0 > gap:
            b += 1
        bloco[i], ant = b, r['DATA_HORA']
    por_bn = collections.defaultdict(list)
    for i, r in enumerate(ordem):
        por_bn[(bloco[i], int(r['NIVEL']))].append(i)
    pos = {}
    for idx in por_bn.values():
        idx.sort(key=lambda i: (ordem[i]['DATA_HORA'], k_nulo(ordem[i].get('ITEM_ID')), str(ordem[i]['ID_REGISTRO'])))
        for p, i in enumerate(idx, start=1):
            pos[i] = p
    segs = {s: n for n, s in enumerate(sorted({(bloco[i], pos[i]) for i in range(len(ordem))}), start=1)}
    out = {}
    for i, r in enumerate(ordem):
        base = int(dia_de(r['DATA']).strftime('%y%m%d'))
        c = segs[(bloco[i], pos[i])]
        out[r['ID_REGISTRO']] = base * 100 + c if c <= 99 else None
    return out


def dias_vazios(listas, n):
    """n dias SEM nenhum dado (recebido ou manual, qualquer analito), do mais recente para tras,
    sempre antes de hoje (manual com data futura vira MANUAL_INCOMPLETO)."""
    ocup = set()
    for linhas in listas:
        for r in linhas:
            d = dia_de(r.get('DATA'))
            if d:
                ocup.add(d)
    if not ocup:
        raise PreCondicao('base sem nenhuma DATA: não há de onde clonar')
    d = min(max(ocup), dt.date.today() - dt.timedelta(days=1))
    limite = d - dt.timedelta(days=900)
    out = []
    while len(out) < n and d > limite:
        if d not in ocup:
            out.append(d)
        d -= dt.timedelta(days=1)
    if len(out) < n:
        raise PreCondicao(f'só {len(out)} dia(s) sem dado real nos últimos 900 dias; preciso de {n}')
    return out


# =============================================================================
# contexto: CFG, cabecalho, molde real e as linhas sinteticas
# =============================================================================
class Contexto:
    OBRIG = ('ID_ORIGEM', 'ITEM_ID', 'DATA', 'DATA_HORA', 'VALOR', 'NIVEL', 'ANALITO', 'ID_AMOSTRA')

    def __init__(self, q, produto, tmp, rotulo, n_dias=5):
        self.q, self.produto, self.tmp, self.rotulo = q, produto, tmp, rotulo
        self.nlv = q.nlv
        self.cfg0 = {}
        for k in ('PREFIXO_ID', 'TABELA_ORIGEM', 'GAP_CORRIDA_MIN', 'MODO_FONTE', 'CAMINHO_DB_SEAC'):
            try:
                self.cfg0[k] = q.cfg(k)
            except IndexError:
                raise PreCondicao(f'tblConfigIntegracao sem a chave {k}')
        self.pref = str(self.cfg0['PREFIXO_ID'] or '').strip()
        self.tabela = str(self.cfg0['TABELA_ORIGEM'] or '').strip()
        self.gap = float(self.cfg0['GAP_CORRIDA_MIN'])
        if not self.pref or not self.tabela:
            raise PreCondicao(f'PREFIXO_ID/TABELA_ORIGEM vazios no CFG: {self.cfg0}')
        if self.gap >= 200:
            raise PreCondicao(f'GAP_CORRIDA_MIN = {self.gap}: o desenho das corridas sintéticas usa intervalos >= 239 min')
        lo = q.lo('tblDB_Recebimento')
        if lo is None:
            raise PreCondicao('tblDB_Recebimento não encontrada na cópia')
        self.cab_rec = [c.Name for c in lo.ListColumns]
        self.cab = [c for c in self.cab_rec if c not in gss.PROIBIDAS]
        faltam = [c for c in self.OBRIG if c not in self.cab]
        if faltam:
            raise PreCondicao(f'tblDB_Recebimento sem as colunas {faltam}: {self.cab_rec}')
        self.rec0 = q.ler('tblDB_Recebimento')
        self.fin0 = ler_final(q)
        if not self.rec0 or not self.fin0:
            raise PreCondicao(f'recebimento ({len(self.rec0)}) ou final ({len(self.fin0)}) vazios')
        ocupados = sorted(i for i in ids_de(self.rec0) | ids_de(self.fin0)
                          if (numero_id(i, self.pref) or 0) >= BASE_ID)
        if ocupados:
            raise PreCondicao(f'a faixa sintética (>= {BASE_ID}) já existe na base: {ocupados[:5]}')
        self.escolher_molde()
        self.dias = dias_vazios([self.rec0, self.fin0], n_dias)
        self.montar_specs()

    # ------------------------------------------------------------------ molde
    def escolher_molde(self):
        q = self.q
        rec_id = {r['ID_REGISTRO']: r for r in self.rec0 if r.get('ID_REGISTRO')}
        lote_an = None
        try:
            v = q.wb.Application.Evaluate('loteAnalise')
            v = v.Value if hasattr(v, 'Value') else v
            lote_an = str(int(v)) if isinstance(v, float) else str(v).strip()
        except Exception:
            pass
        eq_pai = None
        if q.bio:
            try:
                eq_pai = str(q.ws('Painel').Range('L4').Value).strip()
            except Exception:
                pass
        grupos, cont = collections.defaultdict(dict), collections.Counter()
        for r in self.fin0:
            if (r['ORIGEM_RESULTADO'] != 'INTERFACEAMENTO' or r['STATUS_ANALITICO'] != 'ATIVO'
                    or r['ANALITO_CADASTRADO'] != SIM or not r['LOTE'] or r['DATA_HORA'] is None
                    or r['NIVEL'] in (None, '')
                    or r['ID_REGISTRO'] not in rec_id or not isinstance(r['RESULTADO'], (int, float))
                    or not isinstance(rec_id[r['ID_REGISTRO']].get('VALOR'), (int, float))):
                continue
            k = (r['EQUIPAMENTO'], r['MATRIZ'], str(r['LOTE']), r['ANALITO'])
            cont[k] += 1
            nv = int(r['NIVEL'])
            if nv not in grupos[k] or r['DATA_HORA'] > grupos[k][nv]['DATA_HORA']:
                grupos[k][nv] = r
        cands = [k for k in grupos if all(n in grupos[k] for n in range(1, self.nlv + 1))]
        if not cands:
            raise PreCondicao(f'nenhum (equipamento, lote, analito) cadastrado com resultado ATIVO nos {self.nlv} níveis')
        cands.sort(key=lambda k: (k[2] == lote_an, eq_pai is None or k[0] == eq_pai, cont[k]), reverse=True)
        self.chave = cands[0]
        self.fin_molde = grupos[self.chave]
        self.rec_molde = {n: rec_id[r['ID_REGISTRO']] for n, r in self.fin_molde.items()}
        self.valor = {n: float(self.rec_molde[n]['VALOR']) for n in self.rec_molde}
        self.lote_painel, self.equip_painel = lote_an, eq_pai
        eq, an0 = self.rec_molde[1].get('EQUIPAMENTO'), self.rec_molde[1].get('ANALITO')
        outros = collections.Counter(r.get('ANALITO') for r in self.rec0
                                     if r.get('EQUIPAMENTO') == eq and isinstance(r.get('ANALITO'), str)
                                     and r.get('ANALITO') not in ('', an0))
        self.analito_alt = outros.most_common(1)[0][0] if outros else None

    def descricao(self):
        return {'grupo(equip, matriz, lote, analito)': list(self.chave), 'lote_painel': self.lote_painel,
                'equip_painel': self.equip_painel, 'molde_ids': {n: r['ID_REGISTRO'] for n, r in self.fin_molde.items()},
                'valores': self.valor, 'dias_sem_dado': [str(d) for d in self.dias], 'GAP_CORRIDA_MIN': self.gap,
                'PREFIXO_ID': self.pref, 'TABELA_ORIGEM': self.tabela, 'colunas_origem': self.cab}

    # ------------------------------------------------------------------ linhas S
    def spec(self, nome, dia, h, m, nivel, id_origem, item_id, valor=None, amostra=None, sem_hora=False):
        return {'_nome': nome, '_modelo': self.rec_molde[nivel], '_id': id_esperado(self.pref, id_origem),
                'ID_ORIGEM': id_origem, 'ITEM_ID': item_id, 'NIVEL': nivel, 'DATA': dia,
                'DATA_HORA': None if sem_hora else dt.datetime.combine(dia, dt.time(h, m)),
                'VALOR': self.valor[nivel] if valor is None else valor, 'ID_AMOSTRA': amostra or f'QASINT-{nome}'}

    def montar_specs(self):
        d1, d2, d3, d4, d5 = self.dias[:5]
        s = {}
        slots = []
        for c, (h, m) in enumerate(((8, 0), (14, 0), (18, 0)), start=1):
            for k in range(self.nlv):
                slots.append((k + 1, h, m + k, c))
        self.corrida_esperada = {}
        for i in range(5):                       # D1: S01..S05, N1/N2(/N3) em horarios conhecidos
            nv, h, m, c = slots[i]
            nome = f'S{i + 1:02d}'
            s[nome] = self.spec(nome, d1, h, m, nv, BASE_ID + i + 1, BASE_ID + i + 1)
            self.corrida_esperada[nome] = c
        s['S13'] = dict(s['S02'], _nome='S13', ID_ORIGEM=BASE_ID + 13, _id=f'{self.pref}-{BASE_ID + 13}')
        # D2: sem ID (texto em branco / celula nula)
        s['S06A'] = self.spec('S06A', d2, 8, 0, 1, ' ', BASE_ID + 6)
        s['S06B'] = self.spec('S06B', d2, 8, 1, 2, None, BASE_ID + 106)
        # D3: ID repetido na mesma carga
        s['S07A'] = self.spec('S07A', d3, 9, 0, 1, BASE_ID + 7, BASE_ID + 7, amostra='QASINT-S07')
        s['S07B'] = dict(s['S07A'], _nome='S07B')
        s['S08A'] = self.spec('S08A', d3, 9, 1, 2, BASE_ID + 8, BASE_ID + 8, valor=5.0, amostra='QASINT-S08')
        s['S08B'] = dict(s['S08A'], _nome='S08B', VALOR=500.0)
        s['S09A'] = self.spec('S09A', d3, 13, 0, 1, BASE_ID + 9, BASE_ID + 9)
        s['S09B'] = self.spec('S09B', d3, 13, 0, 1, BASE_ID + 9.4, BASE_ID + 109,
                              valor=round(self.valor[1] * 1.2 + 0.1, 4))
        # D4: ID em formato nao canonico
        s['S10'] = self.spec('S10', d4, 10, 0, 1, f' {BASE_ID + 10} ', BASE_ID + 10)
        s['S11'] = self.spec('S11', d4, 10, 1, 2, f'{BASE_ID + 11}', BASE_ID + 11)
        s['S12'] = self.spec('S12', d4, 15, 0, 1, 'ABC', BASE_ID + 12)
        # D5: DATA_HORA nula (D16) + vizinho valido no mesmo grupo/dia
        s['S14'] = self.spec('S14', d5, 0, 0, 1, BASE_ID + 14, BASE_ID + 14, sem_hora=True)
        s['S15'] = self.spec('S15', d5, 9, 0, 1, BASE_ID + 15, BASE_ID + 15)
        # D09c: interface no fim do dia (o manual do dia seguinte as 00:02 e o mesmo resultado)
        s['S16'] = self.spec('S16', d5, 23, 58, 2, BASE_ID + 16, BASE_ID + 16)
        self.specs = s

    def arquivo(self, rotulo, nomes, alter=None):
        linhas = []
        for n in nomes:
            l = dict(self.specs[n])
            l.update((alter or {}).get(n, {}))
            linhas.append(l)
        path = os.path.join(self.tmp, f'DB_SEAC_sintetico_{self.rotulo}_{rotulo}.xlsx')
        esc = gss.gerar(path, self.produto, self.cab, self.rec_molde[1], linhas, tabela=self.tabela)
        for e, l in zip(esc, linhas):
            e['_nome'], e['_id'] = l['_nome'], l['_id']
        return path, esc

    def linhas_grupo(self, fin, dia):
        return [r for r in fin if (r['EQUIPAMENTO'], r['MATRIZ'], str(r['LOTE']), r['ANALITO']) == self.chave
                and dia_de(r['DATA']) == dia and r['STATUS_ANALITICO'] in CORRIDA_REAL]


def carregar(q, caminho_xlsx, rotulo):
    q.cfg('CAMINHO_DB_SEAC', caminho_xlsx)
    return tentar(q, rotulo)


def ev_linha(e):
    return {k: e.get(k) for k in ('_nome', '_id', 'ID_ORIGEM', 'ITEM_ID', 'NIVEL', 'DATA_HORA', 'VALOR', 'ANALITO',
                                  'ID_AMOSTRA')}


# =============================================================================
# copia A: E01-E05
# =============================================================================
def conferir_e01(ctx, esc, rec0, rec1, fin1, erro, t):
    base = {'erro_atualizacao': erro, 't_s': t}
    novos = ids_de(rec1) - ids_de(rec0)
    esperados = {e['_id'] for e in esc}
    R('E01a', camadas_ok(erro) and len(esperados) == 5 and novos == esperados,
      dict(base, esperados=sorted(esperados), novos=sorted(novos)[:10], n_novos=len(novos),
           linhas=[len(rec0), len(rec1)], molde=ctx.descricao()))

    rid = agrupar(rec1)
    difs, conferidas = [], 0
    for e in esc:
        rows = rid.get(e['_id'], [])
        if len(rows) != 1:
            difs.append({'linha': e['_nome'], 'linhas_no_recebimento': len(rows)})
            continue
        for c in ctx.cab:
            conferidas += 1
            if norm(e[c]) != norm(rows[0].get(c)):
                difs.append({'linha': e['_nome'], 'coluna': c, 'origem': e[c], 'recebido': rows[0].get(c)})
    R('E01b', camadas_ok(erro) and len(esc) == 5 and conferidas == 5 * len(ctx.cab) and not difs,
      dict(base, celulas_conferidas=conferidas, colunas=ctx.cab, divergencias=difs[:8],
           origem=[ev_linha(e) for e in esc]))

    carimbos = {norm(r.get('RECEBIDO_EM')) for i in novos for r in rid.get(i, [])}
    R('E01c', camadas_ok(erro) and bool(novos) and len(carimbos) == 1 and None not in carimbos,
      dict(base, carimbos=sorted(str(c) for c in carimbos)) if novos else
      {**base, 'pré-condição ausente': 'nenhum ID novo no recebimento'})

    def tupla(r):
        return tuple(norm(r.get(c)) for c in ctx.cab_rec)
    antes = collections.Counter(tupla(r) for r in rec0)
    depois = collections.Counter(tupla(r) for r in rec1 if r.get('ID_REGISTRO') not in novos)
    # linha SEM ID ja guardada antes (D02 da base real) some na proxima carga: e perda de historico e
    # derruba o E01d -- a evidencia separa esse caso para nao confundir com defeito do anti-join
    sem_id_antes = [r for r in rec0 if r.get('ID_REGISTRO') in (None, '')]
    R('E01d', camadas_ok(erro) and antes == depois and len(rec1) == len(rec0) + len(novos) and len(rec0) > 0,
      dict(base, linhas_antes=len(rec0), linhas_depois=len(rec1), novos=len(novos),
           linhas_sem_ID_na_base_antes=len(sem_id_antes),
           obs_sem_ID=('linhas sem ID já guardadas somem na carga (perda por D02)' if sem_id_antes else None),
           sumiram_ou_mudaram=[list(x) for x in list((antes - depois).elements())[:3]],
           surgiram=[list(x) for x in list((depois - antes).elements())[:3]]))

    fid = agrupar(fin1)
    d1 = ctx.dias[0]
    orac = oraculo_run(ctx.linhas_grupo(fin1, d1), ctx.gap)
    probs, vistos = [], []
    for e in esc:
        rows = fid.get(e['_id'], [])
        if len(rows) != 1:
            probs.append({'linha': e['_nome'], 'linhas_na_final': len(rows)})
            continue
        r = rows[0]
        run_regra = int(d1.strftime('%y%m%d')) * 100 + ctx.corrida_esperada[e['_nome']]
        vistos.append(dict(resumo_final(r), oraculo=orac.get(e['_id']), regra_desenhada=run_regra))
        if not (r['STATUS_ANALITICO'] == 'ATIVO' and r['PARTICIPA_ESTATISTICA'] == SIM
                and r['TIPO_PLOTAGEM_LJ'] == 'NORMAL' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                and norm(r['RESULTADO']) == norm(e['VALOR'])
                and r['RUN'] is not None and int(r['RUN']) == orac.get(e['_id']) == run_regra):
            probs.append({'linha': e['_nome'], 'final': resumo_final(r), 'oraculo': orac.get(e['_id']),
                          'regra_desenhada': run_regra})
    R('E01e', camadas_ok(erro) and len(esc) == 5 and len(vistos) == 5 and not probs,
      dict(base, problemas=probs[:6], final=vistos, linhas_do_grupo_no_dia=len(orac)))


def sessao_a(produto, caminho, tmp):
    copia = os.path.join(tmp, 'copia_A_E01_E05.xlsm')
    shutil.copy2(caminho, copia)
    q = QA(produto, copia)
    try:
        ctx = Contexto(q, produto, tmp, 'A')
        pref = ctx.pref
        print('    molde:', json.dumps(ctx.descricao(), ensure_ascii=False, default=str)[:900], flush=True)
        q.cfg('MODO_FONTE', 'SEAC')
        s1_5 = ['S01', 'S02', 'S03', 'S04', 'S05']

        # ------------------------------------------------------------ L1: E01
        path1, esc1 = ctx.arquivo('L1', s1_5)
        erro1, _, t1 = carregar(q, path1, 'A-L1 (S01-S05)')
        rec1, fin1 = q.ler('tblDB_Recebimento'), ler_final(q)
        conferir_e01(ctx, esc1, ctx.rec0, rec1, fin1, erro1, t1)
        if not camadas_ok(erro1):
            raise Interrompida(f'a carga L1 (S01-S05) falhou: {erro1[:300]}')
        run_l1 = {r['ID_REGISTRO']: r['RUN'] for r in fin1 if r['ID_REGISTRO'] in {e['_id'] for e in esc1}}

        # ------------------------------------------------------------ L2..L4: E02, E03, E05a/b
        nomes2 = s1_5 + ['S06A', 'S06B', 'S07A', 'S07B', 'S08A', 'S08B', 'S09A', 'S09B', 'S13']
        tentativas = []
        while True:
            path2, esc2 = ctx.arquivo(f'L2_{len(tentativas) + 1}', nomes2)
            erro2, _, t2 = carregar(q, path2, f'A-L2 tentativa {len(tentativas) + 1}')
            rec2 = q.ler('tblDB_Recebimento')
            tentativas.append({'linhas': list(nomes2), 'erro': erro2, 't_s': t2, 'recebimento': len(rec2)})
            if erro2 is None or len(rec2) > len(rec1):
                break
            retirar = next((n for n in ('S06A', 'S06B') if n in nomes2), None)
            if retirar is None:
                raise Interrompida(f'a carga L2 não entra nem sem as linhas sem ID: {erro2[:300]}')
            nomes2.remove(retirar)                # a linha sem ID derrubou a leitura da origem: segue sem ela
        fin2, qa2 = ler_final(q), q.ler('tblQA_Integracao')
        erro3, _, t3 = carregar(q, path2, 'A-L3 (mesmo arquivo)')
        rec3, fin3, qa3 = q.ler('tblDB_Recebimento'), ler_final(q), q.ler('tblQA_Integracao')
        erro4, _, t4 = carregar(q, path2, 'A-L4 (mesmo arquivo)')
        rec4, fin4, qa4 = q.ler('tblDB_Recebimento'), ler_final(q), q.ler('tblQA_Integracao')
        base2 = {'tentativas_L2': tentativas, 'erros_L3_L4': [erro3, erro4], 't_L2_L3_L4_s': [t2, t3, t4]}

        # E02 ---------------------------------------------------------
        s06 = [e for e in esc2 if e['_nome'] in ('S06A', 'S06B')]
        marc06 = {e['ID_AMOSTRA'] for e in s06}
        r2_06 = [r for r in rec2 if r.get('ID_AMOSTRA') in marc06]
        r3_06 = [r for r in rec3 if r.get('ID_AMOSTRA') in marc06]

        def carimbos(rec):
            return collections.Counter((r.get('ID_REGISTRO') or ('SEM_ID', r.get('ID_AMOSTRA')), norm(r.get('RECEBIDO_EM')))
                                       for r in rec)
        c2, c3 = carimbos(rec2), carimbos(rec3)
        ev = dict(base2, linhas_sem_id_na_origem=[ev_linha(e) for e in s06],
                  no_recebimento_L2=[{k: r.get(k) for k in ('ID_REGISTRO', 'ID_AMOSTRA', 'RECEBIDO_EM')} for r in r2_06],
                  no_recebimento_L3=[{k: r.get(k) for k in ('ID_REGISTRO', 'ID_AMOSTRA', 'RECEBIDO_EM')} for r in r3_06],
                  recarimbadas_entre_L2_e_L3=[list(x) for x in list((c2 - c3).elements())[:4]],
                  na_final=[resumo_final(r) for r in fin3 if r.get('ID_AMOSTRA') in marc06])
        if not s06:
            ev['pré-condição ausente'] = 'nenhuma linha sem ID chegou a ser carregada (ver tentativas_L2)'
        R('E02a', bool(s06) and camadas_ok(erro2) and not r2_06 and not r3_06 and c2 == c3, ev)

        e06 = [r for r in qa3 if r.get('CODIGO') == 'E06']
        faltam = []
        an_ok = {ctx.rec_molde[1].get('ANALITO'), ctx.fin_molde[1].get('ANALITO')}
        for e in s06:
            # contrato: DATA_HORA (data E hora), ANALITO, ID_AMOSTRA e NIVEL no DETALHE
            hora = e['DATA_HORA'].strftime('%H:%M') if isinstance(e.get('DATA_HORA'), dt.datetime) else None
            achou = [f for f in e06 if f.get('SEVERIDADE') == 'ERRO' and e['ID_AMOSTRA'] in str(f.get('DETALHE'))
                     and any(a and str(a) in str(f.get('DETALHE')) for a in an_ok)
                     and any(tx in str(f.get('DETALHE')) for tx in textos_data(e['DATA']))
                     and hora is not None and hora in str(f.get('DETALHE'))
                     and str(e['NIVEL']) in str(f.get('DETALHE'))]
            if not achou:
                faltam.append(e['_nome'])
        R('E02b', bool(s06) and not faltam,
          {'sem_achado_completo': faltam, 'E06': [{k: f.get(k) for k in ('SEVERIDADE', 'TESTE', 'DATA_HORA', 'DETALHE')}
                                                  for f in e06][:4],
           'codigos_QA_L3': dict(collections.Counter(f.get('CODIGO') for f in qa3)),
           **({'pré-condição ausente': 'nenhuma linha sem ID na origem'} if not s06 else {})})

        ids_antes = ids_de(rec1)
        novos2 = ids_de(rec2) - ids_antes
        # linha rejeitada = a que tem achado E06 com o seu marcador ID_AMOSTRA (assim a conta nao depende de a
        # correcao tratar o fracionario S09B como repeticao (E07) ou como ID invalido (E06))
        e06_l2 = [f for f in qa2 if f.get('CODIGO') == 'E06' and f.get('SEVERIDADE') == 'ERRO']
        rejeitadas = [e for e in esc2 if any(e['ID_AMOSTRA'] in str(f.get('DETALHE')) for f in e06_l2)]
        nomes_rej = {e['_nome'] for e in rejeitadas}
        n_ja = sum(1 for e in esc2 if e['_id'] in ids_antes and e['_nome'] not in nomes_rej)
        cont_novos = collections.Counter(e['_id'] for e in esc2
                                         if e['_id'] and e['_id'] not in ids_antes and e['_nome'] not in nomes_rej)
        n_rep = sum(c - 1 for c in cont_novos.values())
        n_rej = len(rejeitadas)
        R('E02c', bool(s06) and camadas_ok(erro2) and len(esc2) == n_ja + len(novos2) + n_rep + n_rej
          and novos2 == set(cont_novos) and len(e06_l2) == n_rej,
          {'linhas_origem': len(esc2), 'ja_recebidas': n_ja, 'ids_novos': len(novos2), 'repeticoes_mesmo_id': n_rep,
           'rejeitadas_reportadas_E06': n_rej, 'rejeitadas': sorted(nomes_rej), 'achados_E06': len(e06_l2),
           'ids_novos_lista': sorted(novos2), 'ids_novos_esperados': sorted(cont_novos), 'erro_L2': erro2})

        # E03 ---------------------------------------------------------
        alvo = {f'{pref}-{BASE_ID + 7}': ('S07A', 'S07B'), f'{pref}-{BASE_ID + 8}': ('S08A', 'S08B'),
                f'{pref}-{BASE_ID + 9}': ('S09A', 'S09B')}
        r2id, f2id = agrupar(rec2), agrupar(fin2)
        s09a = next(e for e in esc2 if e['_nome'] == 'S09A')
        det, ok_a = {}, camadas_ok(erro2)
        for i in alvo:
            rr = r2id.get(i, [])
            det[i] = {'recebimento': [{k: r.get(k) for k in ('ITEM_ID', 'VALOR', 'ID_ORIGEM', 'RECEBIDO_EM')} for r in rr],
                      'final': [resumo_final(r) for r in f2id.get(i, [])]}
            ok_a = ok_a and len(rr) == 1
        regra = r2id.get(f'{pref}-{BASE_ID + 9}', [])
        ok_regra = (len(regra) == 1 and norm(regra[0].get('ITEM_ID')) == BASE_ID + 9
                    and norm(regra[0].get('VALOR')) == norm(s09a['VALOR']))
        s08 = r2id.get(f'{pref}-{BASE_ID + 8}', [])
        R('E03a', ok_a and ok_regra and len(s08) == 1 and norm(s08[0].get('VALOR')) in (5, 500),
          {'por_id': det, 'regra_menor_ITEM_ID_9000009': ok_regra,
           'valor_mantido_9000008': [r.get('VALOR') for r in s08],
           'desvio_do_plano': 'S9 = 9000009.4 colide com S09A (9000009) na mesma carga; S1 já entrou na carga L1'})

        cont_ids = collections.Counter(r['ID_REGISTRO'] for r in fin2 if r['ID_REGISTRO'])
        dup_ids = {i: n for i, n in cont_ids.items() if n > 1}
        # chave do motor so nos dias sinteticos: defeito antigo da base real nao contamina este teste (o T04 cuida dele)
        runs_sint = {int(d.strftime('%y%m%d')) for d in ctx.dias}
        motor = collections.Counter((r['EQUIPAMENTO'], r['LOTE'], r['NIVEL'], r['ANALITO'], r['RUN'])
                                    for r in fin2 if r['PARTICIPA_ESTATISTICA'] == SIM and r['RUN'] not in (None, '')
                                    and int(r['RUN']) // 100 in runs_sint)
        dup_motor = {str(k): n for k, n in motor.items() if n > 1}
        part_alvo = {i: sum(1 for r in f2id.get(i, []) if r['PARTICIPA_ESTATISTICA'] == SIM) for i in alvo}
        R('E03b', camadas_ok(erro2) and not dup_ids and not dup_motor and all(len(f2id.get(i, [])) == 1 for i in alvo),
          {'ids_repetidos_na_final': dict(list(dup_ids.items())[:8]), 'chave_motor_com_mais_de_1': dict(list(dup_motor.items())[:5]),
           'linhas_na_final': {i: len(f2id.get(i, [])) for i in alvo}, 'participantes_por_id': part_alvo})

        e07 = [f for f in qa2 if f.get('CODIGO') == 'E07' and f.get('SEVERIDADE') == 'ERRO']
        ids07 = {f.get('ID_REGISTRO') for f in e07}
        nossos_unicos = {e['_id'] for e in esc2 if e['_id'] and e['_id'] not in alvo}
        R('E03c', camadas_ok(erro2) and set(alvo) <= ids07 and not (ids07 & nossos_unicos),
          {'E07_ids': sorted(str(i) for i in ids07)[:10], 'esperados': sorted(alvo),
           'E07_indevido_em_ID_unico': sorted(ids07 & nossos_unicos),
           'E07': [{k: f.get(k) for k in ('ID_REGISTRO', 'TESTE', 'DETALHE')} for f in e07][:4]})

        # E05a --------------------------------------------------------
        i13, i02 = f'{pref}-{BASE_ID + 13}', f'{pref}-{BASE_ID + 2}'
        f13, f02 = f2id.get(i13, []), f2id.get(i02, [])
        a02 = [f for f in qa2 if f.get('CODIGO') == 'A02' and f.get('ID_REGISTRO') == i13]
        R('E05a', camadas_ok(erro2) and len(f13) == 1 and f13[0]['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM'
          and f13[0]['ID_RELACIONADO'] == i02 and f13[0]['PARTICIPA_ESTATISTICA'] != SIM and bool(a02)
          and len(f02) == 1 and f02[0]['STATUS_ANALITICO'] == 'ATIVO',
          {'S13': [resumo_final(r) for r in f13], 'S02': [resumo_final(r) for r in f02],
           'A02': [f.get('DETALHE') for f in a02][:2], 'S13_no_recebimento': len(r2id.get(i13, []))})

        # E05b --------------------------------------------------------
        def carimbo_por_id(rec):
            d = collections.defaultdict(list)
            for r in rec:
                if r.get('ID_REGISTRO'):
                    d[r['ID_REGISTRO']].append(norm(r.get('RECEBIDO_EM')))
            return {k: sorted(v, key=str) for k, v in d.items()}

        def estado_ids(fin):
            return collections.Counter((r['ID_REGISTRO'], r['RUN'], r['STATUS_ANALITICO']) for r in fin if r['ID_REGISTRO'])
        cp = [carimbo_por_id(x) for x in (rec2, rec3, rec4)]
        ef = [estado_ids(x) for x in (fin2, fin3, fin4)]
        run_s = {i: [run_l1.get(i)] + [next((r['RUN'] for r in agrupar(f).get(i, [])), None) for f in (fin2, fin3, fin4)]
                 for i in run_l1}
        mudou_carimbo = [k for k in cp[0] if cp[0][k] != cp[1].get(k) or cp[0][k] != cp[2].get(k)]
        R('E05b', all(camadas_ok(x) for x in (erro2, erro3, erro4)) and len(rec2) == len(rec3) == len(rec4)
          and cp[0] == cp[1] == cp[2] and ef[0] == ef[1] == ef[2] and len(run_l1) == 5
          and all(len(set(v)) == 1 for v in run_s.values()),
          dict(base2, recebimento=[len(rec2), len(rec3), len(rec4)], final=[len(fin2), len(fin3), len(fin4)],
               ids_com_RECEBIDO_EM_alterado=mudou_carimbo[:5], final_igual=[ef[0] == ef[1], ef[1] == ef[2]],
               RUN_S1_S5_L1_L2_L3_L4=run_s))

        # ------------------------------------------------------------ L5: E05c/E05d (origem alterada)
        i01, i03 = f'{pref}-{BASE_ID + 1}', f'{pref}-{BASE_ID + 3}'
        orig01 = next(e for e in esc1 if e['_nome'] == 'S01')
        orig03 = next(e for e in esc1 if e['_nome'] == 'S03')
        a03_antes = {f.get('ID_REGISTRO') for f in qa4 if f.get('CODIGO') == 'A03'}
        v_novo = round(orig01['VALOR'] * 1.1 + 1.0, 4)
        alter = {'S01': {'VALOR': v_novo}}
        if ctx.analito_alt:
            alter['S03'] = {'ANALITO': ctx.analito_alt}
        path5, _ = ctx.arquivo('L5', nomes2, alter)
        erro5, _, t5 = carregar(q, path5, 'A-L5 (S01 VALOR e S03 ANALITO alterados na origem)')
        rec5, fin5, qa5 = q.ler('tblDB_Recebimento'), ler_final(q), q.ler('tblQA_Integracao')
        r5, f5 = agrupar(rec5), agrupar(fin5)
        a03 = collections.defaultdict(list)
        for f in qa5:
            if f.get('CODIGO') == 'A03':
                a03[f.get('ID_REGISTRO')].append(f.get('DETALHE'))
        rec01, fin01 = r5.get(i01, []), f5.get(i01, [])
        ev = {'erro_L5': erro5, 't_s': t5, 'valor_origem_antes': orig01['VALOR'], 'valor_origem_agora': v_novo,
              'recebido': [r.get('VALOR') for r in rec01], 'final': [r.get('RESULTADO') for r in fin01],
              'A03': a03.get(i01, [])[:2], 'A03_ja_existia_antes': i01 in a03_antes}
        if i01 in a03_antes:
            ev['pré-condição ausente'] = 'A03 de S1 já existia antes da alteração (asserção vazia)'
        R('E05c', camadas_ok(erro5) and i01 not in a03_antes and len(rec01) == 1 and norm(rec01[0].get('VALOR')) == norm(orig01['VALOR'])
          and len(fin01) == 1 and norm(fin01[0]['RESULTADO']) == norm(orig01['VALOR']) and bool(a03.get(i01)), ev)
        rec03 = r5.get(i03, [])
        ev = {'erro_L5': erro5, 'analito_origem_antes': orig03['ANALITO'], 'analito_origem_agora': ctx.analito_alt,
              'valor': orig03['VALOR'], 'recebido': [r.get('ANALITO') for r in rec03], 'A03': a03.get(i03, [])[:2],
              'A03_ja_existia_antes': i03 in a03_antes}
        if not ctx.analito_alt:
            ev['pré-condição ausente'] = 'não há outro analito do mesmo equipamento para a troca'
        if i03 in a03_antes:
            ev['pré-condição ausente'] = 'A03 de S3 já existia antes da alteração (asserção vazia)'
        R('E05d', camadas_ok(erro5) and bool(ctx.analito_alt) and i03 not in a03_antes and len(rec03) == 1
          and rec03[0].get('ANALITO') == orig03['ANALITO'] and bool(a03.get(i03)), ev)

        # ------------------------------------------------------------ L6: E04a (ID em texto)
        nomes6 = nomes2 + ['S10', 'S11']
        path6, esc6 = ctx.arquivo('L6', nomes6, alter)
        erro6, _, t6 = carregar(q, path6, 'A-L6 (+S10 " 9000010 ", S11 "9000011")')
        rec6, fin6 = q.ler('tblDB_Recebimento'), ler_final(q)
        r6, f6 = agrupar(rec6), agrupar(fin6)
        ids1011 = [f'{pref}-{BASE_ID + 10}', f'{pref}-{BASE_ID + 11}']
        R('E04a', camadas_ok(erro6) and all(len(r6.get(i, [])) == 1 and len(f6.get(i, [])) == 1
                                        and f6[i][0]['STATUS_ANALITICO'] == 'ATIVO' for i in ids1011),
          {'erro_L6': erro6, 't_s': t6, 'origem': [ev_linha(e) for e in esc6 if e['_nome'] in ('S10', 'S11')],
           'recebimento': {i: [{k: r.get(k) for k in ('ID_ORIGEM', 'VALOR')} for r in r6.get(i, [])] for i in ids1011},
           'final': {i: [resumo_final(r) for r in f6.get(i, [])] for i in ids1011},
           'linhas_S10_S11_por_ID_AMOSTRA': [r.get('ID_REGISTRO') for r in rec6
                                             if r.get('ID_AMOSTRA') in ('QASINT-S10', 'QASINT-S11')]})
        if erro6 is not None and len(rec6) == len(rec5):
            nomes6 = list(nomes2)                    # a origem com S10/S11 nao foi lida: segue sem elas

        # ------------------------------------------------------------ L7: E04b (ID 'ABC')
        nomes7 = nomes6 + ['S12']
        n_rec_a, n_fin_a, g1_a = len(rec6), len(fin6), carimbo_motor(q)
        path7, _ = ctx.arquivo('L7', nomes7, alter)
        erro7, _, t7 = carregar(q, path7, 'A-L7 (+S12 "ABC")')
        rec7, fin7, qa7 = q.ler('tblDB_Recebimento'), ler_final(q), q.ler('tblQA_Integracao')
        g1_b = carimbo_motor(q)
        s12_rec = [r for r in rec7 if r.get('ID_AMOSTRA') == 'QASINT-S12']
        achado12 = [f for f in qa7 if f.get('SEVERIDADE') == 'ERRO'
                    and ('QASINT-S12' in str(f.get('DETALHE')) or 'ABC' in str(f.get('ID_REGISTRO') or ''))]
        if not camadas_ok(erro7):
            caminho_ok = (('ABC' in erro7 or 'ID_ORIGEM' in erro7) and len(rec7) == n_rec_a and len(fin7) == n_fin_a
                          and g1_a == g1_b)
            ramo = 'falhou'
        else:
            caminho_ok = bool(achado12) and not s12_rec
            ramo = 'concluiu'
        R('E04b', caminho_ok,
          {'ramo': ramo, 'erro_L7': (erro7 or '')[:400], 't_s': t7, 'recebimento': [n_rec_a, len(rec7)],
           'final': [n_fin_a, len(fin7)], 'carimbo_motor_inalterado': g1_a == g1_b,
           'S12_no_recebimento': [{k: r.get(k) for k in ('ID_REGISTRO', 'ID_ORIGEM', 'RECEBIDO_EM')} for r in s12_rec],
           'achado_ERRO': [{k: f.get(k) for k in ('CODIGO', 'ID_REGISTRO', 'DETALHE')} for f in achado12][:3]})
        if erro7 is not None and len(rec7) == n_rec_a:
            nomes7 = list(nomes6)                    # a origem com 'ABC' nao foi lida: L8 segue sem ela

        # ------------------------------------------------------------ L8: E04c (DATA_HORA nula) -- por ultimo
        nomes8 = nomes7 + ['S14', 'S15', 'S16']
        path8, esc8 = ctx.arquivo('L8', nomes8, alter)
        # E13c: manual com a chave do S16 as 00:02 do dia seguinte (mesma atualizacao)
        f16 = ctx.fin_molde[2]
        d16 = ctx.dias[4] + dt.timedelta(days=1)
        cols_man = {c.Name for c in q.lo('tblResultados_Manuais').ListColumns}
        man16 = {'DATA': float((dt.datetime.combine(d16, dt.time()) - dt.datetime(1899, 12, 30)).days),
                 'HORA': (2.0 / 60) / 24, 'EQUIPAMENTO': f16['EQUIPAMENTO'], 'MATRIZ': f16['MATRIZ'],
                 'LOTE': str(f16['LOTE']), 'NIVEL': 2, 'ANALITO': f16['ANALITO'],
                 'RESULTADO': float(ctx.valor[2]), 'MOTIVO': 'QA E13c: falha do interfaceamento (teste)'}
        lin16 = q.escrever_linha('tblResultados_Manuais', {k: v for k, v in man16.items() if k in cols_man})
        erro8, _, t8 = carregar(q, path8, 'A-L8 (+S14 DATA_HORA nula, S15, S16 23:58 + manual 00:02)')
        rec8, fin8, qa8 = q.ler('tblDB_Recebimento'), ler_final(q), q.ler('tblQA_Integracao')
        i14, i15 = f'{pref}-{BASE_ID + 14}', f'{pref}-{BASE_ID + 15}'
        f8 = agrupar(fin8)
        d5 = ctx.dias[4]
        orac = oraculo_run(ctx.linhas_grupo(fin8, d5), ctx.gap)
        run15 = int(d5.strftime('%y%m%d')) * 100 + 1
        e09 = [f for f in qa8 if f.get('CODIGO') == 'E09' and f.get('SEVERIDADE') == 'ERRO' and f.get('ID_REGISTRO') == i14]
        s14, s15 = f8.get(i14, []), f8.get(i15, [])
        R('E04c', camadas_ok(erro8) and len(s15) == 1 and s15[0]['STATUS_ANALITICO'] == 'ATIVO' and s15[0]['RUN'] is not None
          and int(s15[0]['RUN']) == run15 == orac.get(i15) and len(s14) == 1 and s14[0]['RUN'] in (None, '')
          and bool(e09),
          {'erro_L8': (erro8 or '')[:400], 't_s': t8, 'S14': [resumo_final(r) for r in s14],
           'S15': [resumo_final(r) for r in s15], 'RUN_esperado_S15': run15, 'oraculo_S15': orac.get(i15),
           'S14_no_recebimento': len(agrupar(rec8).get(i14, [])), 'E09': [f.get('DETALHE') for f in e09][:2]})
        i16 = f'{pref}-{BASE_ID + 16}'
        id16 = q.lo('tblResultados_Manuais').ListColumns('ID_REGISTRO').DataBodyRange.Cells(lin16, 1).Value
        m16 = f8.get(str(id16), [])
        R('E13c', camadas_ok(erro8) and len(f8.get(i16, [])) == 1 and len(m16) == 1
          and m16[0]['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and m16[0]['ID_RELACIONADO'] == i16
          and m16[0]['PARTICIPA_ESTATISTICA'] == NAO,
          {'erro_L8': (erro8 or '')[:300], 'interface_S16': [resumo_final(r) for r in f8.get(i16, [])],
           'manual': [resumo_final(r) for r in m16], 'id_manual': id16})
    finally:
        q.fechar(salvar=False)


# =============================================================================
# copia B: E06
# =============================================================================
def linha_de_base(q, xid, man_id):
    fin = ler_final(q)
    g = agrupar(fin)
    fx, fm = g.get(xid, []), g.get(man_id, [])
    ok = (len(fx) == 1 and fx[0]['STATUS_ANALITICO'] == 'INATIVADO' and fx[0]['TEM_JUSTIFICATIVA'] == SIM
          and fx[0]['GOVERNANCA'] == 'OK' and len(fm) == 1 and fm[0]['STATUS_ANALITICO'] == 'ATIVO'
          and fm[0]['PARTICIPA_ESTATISTICA'] == SIM)
    return ok, {'inativado': [dict(resumo_final(r), GOVERNANCA=r['GOVERNANCA']) for r in fx],
                'manual': [resumo_final(r) for r in fm]}, fin


def sessao_b(produto, caminho, tmp):
    copia = os.path.join(tmp, 'copia_B_E06.xlsm')
    shutil.copy2(caminho, copia)
    q = QA(produto, copia)
    try:
        q.cfg('MODO_FONTE', 'HISTORICO')
        fin0, rec0 = ler_final(q), q.ler('tblDB_Recebimento')
        cands = [r for r in fin0 if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO'
                 and r['ANALITO_CADASTRADO'] == SIM and r['DATA_HORA'] is not None and r['LOTE']
                 and isinstance(r['RESULTADO'], (int, float))]
        if not cands:
            raise PreCondicao('nenhum resultado ATIVO do interfaceamento de analito cadastrado para inativar')
        X = max(cands, key=lambda r: r['DATA_HORA'])
        xid = X['ID_REGISTRO']
        dia = dias_vazios([rec0, fin0], 1)[0]
        q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': xid})
        q.escrever_linha('tblComentariosTecnicos', {'ID_REGISTRO': xid,
                                                    'COMENTARIO_TECNICO': 'QA E06: inativação justificada (teste).'})
        serie = float((dt.datetime.combine(dia, dt.time()) - dt.datetime(1899, 12, 30)).days)
        cols_man = {c.Name for c in q.lo('tblResultados_Manuais').ListColumns}
        manual = {'DATA': serie, 'HORA': 10.0 / 24, 'EQUIPAMENTO': X['EQUIPAMENTO'], 'MATRIZ': X['MATRIZ'],
                  'LOTE': str(X['LOTE']), 'NIVEL': int(X['NIVEL']), 'ANALITO': X['ANALITO'],
                  'RESULTADO': float(X['RESULTADO']), 'MOTIVO': 'QA E06: falha do interfaceamento (teste)'}
        lin_man = q.escrever_linha('tblResultados_Manuais', {k: v for k, v in manual.items() if k in cols_man})
        erro0, _, t0 = tentar(q, 'B0 linha de base (HISTORICO)')
        man_id = q.lo('tblResultados_Manuais').ListColumns('ID_REGISTRO').DataBodyRange.Cells(lin_man, 1).Value
        base_ok, sinais, _ = linha_de_base(q, xid, man_id)
        R('E06p', camadas_ok(erro0) and base_ok,
          {'erro': erro0, 't_s': t0, 'id_inativado': xid, 'id_manual': man_id, 'dia_do_manual': str(dia), **sinais})

        cenarios = [('tblInativacao_NaoConformes', 'E06a1'), ('tblResultados_Manuais', 'E06a2'),
                    ('tblComentariosTecnicos', 'E06a3')]
        for tabela, tid in cenarios:
            if not base_ok:
                R(tid, False, {'pré-condição ausente': 'linha de base E06p não estabelecida', **sinais})
                continue
            fin_a = ler_final(q)
            est_a, g1_a = estado_final(fin_a), carimbo_motor(q)
            novo = 'qaRenomeada_' + tid
            renomear(q, tabela, novo)
            try:
                erro, _, t = tentar(q, f'B-{tid} ({tabela} renomeada)')
            finally:
                renomear(q, novo, tabela)
            ok_b, sinais_d, fin_b = linha_de_base(q, xid, man_id)
            est_b, g1_b = estado_final(fin_b), carimbo_motor(q)
            R(tid, erro is not None and tabela in erro and g1_a == g1_b and est_a == est_b,
              {'erro': (erro or '')[:400], 't_s': t, 'carimbo_motor_inalterado': g1_a == g1_b,
               'final_inalterada': est_a == est_b, 'diferencas_final': diff_estado(est_a, est_b) if est_a != est_b else None,
               'estado_depois': sinais_d})
            if est_a != est_b or g1_a != g1_b:          # passou calado e mudou a final: refaz a linha de base
                erro_r, _, _ = tentar(q, f'B-{tid} refaz a linha de base')
                base_ok, sinais, _ = linha_de_base(q, xid, man_id)
                if erro_r is not None or not base_ok:
                    print('    linha de base não voltou:', erro_r, sinais, flush=True)

        # ------------------------------------------------------------ (b) MODO SEAC sem tblDB_Recebimento
        ctx = Contexto(q, produto, tmp, 'B')
        path, esc = ctx.arquivo('E06b', ['S01', 'S02', 'S03', 'S04', 'S05'])
        q.cfg('MODO_FONTE', 'SEAC')
        q.cfg('CAMINHO_DB_SEAC', path)
        erro_s, _, t_s = tentar(q, 'B-E06b0 linha de base SEAC (sintético)')
        rec_s = q.ler('tblDB_Recebimento')
        novos = ids_de(rec_s) - ids_de(ctx.rec0)
        pre = erro_s is None and novos == {e['_id'] for e in esc}
        if not pre:
            ev = {'pré-condição ausente': 'a origem sintética não foi recebida em MODO SEAC', 'erro': erro_s,
                  'novos': sorted(novos)}
            R('E06b1', False, ev)
            R('E06b2', False, ev)
            return
        fin_a = ler_final(q)
        est_a, g1_a, n_rec_a = estado_final(fin_a), carimbo_motor(q), len(rec_s)
        ids_rec_a = ids_de(rec_s)
        novo = 'qaRenomeada_Recebimento'
        renomear(q, 'tblDB_Recebimento', novo)
        try:
            erro1, _, t1 = tentar(q, 'B-E06b1 (tblDB_Recebimento renomeada)')
            lo_r = q.lo(novo)
            n_rec_1 = lo_r.ListRows.Count
            fin_b = ler_final(q)
            est_b, g1_b = estado_final(fin_b), carimbo_motor(q)
            R('E06b1', erro1 is not None and 'tblDB_Recebimento' in erro1 and g1_a == g1_b and est_a == est_b
              and n_rec_1 == n_rec_a,
              {'erro': (erro1 or '')[:400], 't_s': t1, 'carimbo_motor_inalterado': g1_a == g1_b,
               'final_inalterada': est_a == est_b, 'recebimento': [n_rec_a, n_rec_1],
               'linha_de_base_SEAC_s': t_s})

            ws = lo_r.Parent
            if not desproteger(ws):
                R('E06b2', False, {'pré-condição ausente': f'aba {ws.Name} continua protegida: o refresh falharia por '
                                                          'proteção, não pela regra'})
            else:
                t0 = time.time()
                try:
                    com_teto(q, lambda: pqlib.atualizar(q.lo(novo)), 900)
                    erro2 = None
                except TimeoutError:
                    raise
                except Exception as e:
                    erro2 = str(e)
                t2 = round(time.time() - t0, 1)
                TEMPOS.append({'atualizacao': 'B-E06b2 (só a consulta do recebimento)', 's': t2, 'ok': erro2 is None})
                linhas_d = pqlib.ler_tabela(q.lo(novo))
                ids_d = {r.get('ID_REGISTRO') for r in linhas_d if r.get('ID_REGISTRO')}
                R('E06b2', erro2 is not None and ('tblDB_Recebimento' in erro2 or 'CARGA_INICIAL' in erro2)
                  and len(linhas_d) == n_rec_a and ids_d == ids_rec_a,
                  {'erro': (erro2 or '')[:400], 't_s': t2, 'recebimento_linhas': [n_rec_a, len(linhas_d)],
                   'ids_perdidos': len(ids_rec_a - ids_d), 'exemplo_ids_depois': sorted(ids_d)[:6]})
        finally:
            renomear(q, novo, 'tblDB_Recebimento')
    finally:
        q.fechar(salvar=False)


# =============================================================================
def executar(produto, caminho, saida):
    t_ini = time.time()
    h0 = sha256(caminho)
    tmp = tempfile.mkdtemp(prefix='_qa_etl_rec_', dir=os.path.dirname(caminho))
    print('PASTA_TEMPORARIA', tmp, flush=True)
    try:
        for nome, fn in (('A', sessao_a), ('B', sessao_b)):
            try:
                fn(produto, caminho, tmp)
            except PreCondicao as e:
                MOTIVO[nome] = f'pré-condição ausente: {e}'
            except Exception as e:
                MOTIVO[nome] = f'sessão {nome} interrompida: {e!r} | ' + traceback.format_exc()[-900:]
            if nome in MOTIVO:
                print('    ' + MOTIVO[nome][:900], flush=True)
    finally:
        for tid in TITULOS:                          # nenhum teste "pula" calado
            if tid not in REGISTRADOS:
                R(tid, False, {'pré-condição ausente': MOTIVO.get(SESSAO[tid], 'cenário não chegou a executar')})
        h1 = sha256(caminho)
        shutil.rmtree(tmp, ignore_errors=True)
        reg('E00 Segurança: arquivo de entrada intacto (SHA-256 igual antes e depois; trabalho só em cópias '
            'descartáveis, nada salvo, pasta temporária apagada)', h0 == h1 and not os.path.exists(tmp),
            {'sha256_antes': h0, 'sha256_depois': h1, 'pasta_temporaria_apagada': not os.path.exists(tmp),
             'atualizacoes': TEMPOS, 'n_atualizacoes': len(TEMPOS),
             'tempo_total_s': round(time.time() - t_ini, 1)})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ETL recebimento: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
