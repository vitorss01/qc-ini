# -*- coding: utf-8 -*-
"""QA da camada ETL (Power Query -> tblCQ_Final -> motor) -- itens E07 a E16 do plano.

Uso: python qa_etl.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json] [--parte todas|1|2]

  --parte todas ... (padrao) uma sessao: E07 base, E08, E09, E10, E11, E12+E14, E13, E15, E16
  --parte 1 ....... E07 base, E08, E09, E10, E11 (E15/E16 sobre o que esta parte fez)
  --parte 2 ....... E07 base, E12+E14, E13, E15, E16
  (dividir em 2 aberturas, se a sessao unica ficar longa demais; cada parte abre a sua copia)

SEGURANCA
  - o arquivo recebido por argumento NAO e aberto: a suite copia para <nome>_qa_etl_tmp.xlsm, abre a
    copia, nunca salva (q.fechar(salvar=False)) e apaga a copia no fim. O SHA-256 do arquivo de entrada e
    conferido antes e depois (E00).
  - nenhuma macro com dialogo: so mIntegracao.AtualizarDadosAutomatico (QA.atualizar),
    mEstatistica.AtualizarEstatistica, RecalcularEstatPeriodo, mIncerteza.RecalcularIncerteza e as
    funcoes mLotes.MediaDoLote/DPDoLote. As tabelas de entrada sao escritas celula a celula com os
    eventos ligados (o Worksheet_Change carimba e normaliza), como o usuario faz.
  - nenhum NumberFormat e gravado; valores entram como float/datetime/texto.
  - assercao sem alvo (conjunto vazio, corrida fora da janela, cenario sem dia viavel) = FAIL com
    'pre-condicao ausente' e a evidencia. Nenhum teste pula calado.

CONTRATO
  Os testes assertam o comportamento-ALVO das correcoes (D05, D06, D08, D09, D10, D12, D14). Onde o
  codigo atual ainda nao corrigiu, o FAIL e a prova do defeito: o nome do teste diz "(hoje falha: prova Dnn)".
"""
import collections
import datetime as dt
import hashlib
import json
import math
import os
import random
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, NAO  # noqa: E402

SIM = 'SIM'
SENHA = 'qcini2025'
EPOCA = dt.datetime(1899, 12, 30)
TOL_REL = 1e-9
RESULTADOS = os.path.join(AQUI, 'resultados')

TB_FINAL = 'tblCQ_Final'
TB_REC = 'tblDB_Recebimento'
TB_MAN = 'tblResultados_Manuais'
TB_INAT = 'tblInativacao_NaoConformes'
TB_COM = 'tblComentariosTecnicos'
TB_QA = 'tblQA_Integracao'
COL_LJ = 'REGISTRAR - LJ'

STATUS = ('MANUAL_INCOMPLETO', 'DUPLICIDADE_ORIGEM', 'CONFLITO_MANUAL', 'SEM_DATA_HORA', 'SEM_VALOR', 'INATIVADO', 'ATIVO')   # SEM_DATA_HORA: D16
CORRIDA_REAL = ('ATIVO', 'INATIVADO', 'SEM_VALOR')
# colunas da tblCQ_Final que uma inativacao (e o seu comentario) pode mudar -- e nenhuma outra
COLS_INATIVACAO = {'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'REGISTRAR_RESULTADO_NO_LJ', 'TIPO_PLOTAGEM_LJ',
                   'INATIVACAO_REGISTRADA', 'DATA_INATIVACAO', 'USUARIO_INATIVACAO', 'TEM_JUSTIFICATIVA',
                   'GOVERNANCA', 'COMENTARIO_TECNICO', 'MOTIVO_INATIVACAO'}          # MOTIVO_INATIVACAO: ADR-070
COLS_CAIXA = {'REGISTRAR_RESULTADO_NO_LJ', 'TIPO_PLOTAGEM_LJ'}
IDX_AQ = 42            # Estatistica!AQ (alertas da incerteza; carrega o % de inativados -- PEXCL)
IDX_ORDEM = 20         # Estatistica!U 'ordem critico': contagem CORRIDA de linhas criticas desde a 1a linha -- muda
                       # quando OUTRO analito muda de status (ex.: o extremo B no mesmo refresh), nao e indicador do analito


# =============================================================================
#  utilitarios puros (sem COM)
# =============================================================================
def sha256(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def eh_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def iint(v):
    if v is None or isinstance(v, bool) or v == '':
        return None
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


def serial(d):
    return (d.replace(tzinfo=None) - EPOCA).total_seconds() / 86400.0


def dia(d):
    """Int() do serial do Excel -- o 'dia' que o VBA compara."""
    return int(math.floor(serial(d)))


def de_serial(n):
    return EPOCA + dt.timedelta(days=float(n))


def meia_noite(d):
    return dt.datetime(d.year, d.month, d.day)


def cstr(v):
    """CStr do VBA para o que vem de celula (o que interessa: inteiro sem '.0')."""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'Verdadeiro' if v else 'Falso'
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v).replace('.', ',')
    return str(v)


def up(v):
    return cstr(v).strip().upper()


def txt(v):
    """Txt do Power Query: Text.Trim(Text.From(x)); vazio = null."""
    if v is None:
        return None
    if isinstance(v, bool):
        s = 'true' if v else 'false'
    elif isinstance(v, float) and v.is_integer():
        s = str(int(v))
    elif isinstance(v, dt.datetime):
        s = v.isoformat()
    else:
        s = str(v)
    s = s.strip()
    return s or None


def digitos(t):
    return t != '' and all('0' <= c <= '9' for c in t)


def norm_id(x, pref, alvo=True):
    """NormId. alvo=True = contrato D14 (PQ, QA e VBA iguais): tira espaco E NBSP; numero puro com zeros a
    esquerda casa com o ID canonico ('00123' -> PREFIXO-123). alvo=False = regra de hoje."""
    t = txt(x)
    if t is None:
        return None
    u = t.upper().replace(' ', '')
    if alvo:
        u = u.replace(' ', '')
    if not u:
        return None
    if digitos(u):
        return f'{pref}-{int(u)}' if alvo else f'{pref}-{u}'
    if u.startswith('MAN_') and digitos(u[4:]):
        return 'MAN_' + u[4:].rjust(4, '0')
    return u


def como_data(v):
    """ComoData do VBA (mEstatPeriodo): serial do DIA; vazio = 0 (sem limite)."""
    if v is None:
        return 0
    if isinstance(v, dt.datetime):
        return dia(v)
    if eh_num(v):
        return int(math.floor(v))
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return 0
        try:
            return int(math.floor(float(s.replace(',', '.'))))
        except ValueError:
            pass
        for fmt in ('%d/%m/%Y', '%d/%m/%Y %H:%M', '%d/%m/%Y %H:%M:%S'):
            try:
                return dia(dt.datetime.strptime(s, fmt))
            except ValueError:
                pass
    return 0


def perto(a, b, tol=TOL_REL):
    """Igualdade com tolerancia RELATIVA (vazio so casa com vazio)."""
    va, vb = a in (None, ''), b in (None, '')
    if va or vb:
        return va and vb
    if not (eh_num(a) and eh_num(b)):
        return a == b
    if a == b:
        return True
    return abs(a - b) <= tol * max(abs(a), abs(b))


def iguais(a, b, tol=TOL_REL):
    """Igualdade profunda de grades COM (tuplas de tuplas): numero com |delta| relativo < tol."""
    if isinstance(a, (tuple, list)) and isinstance(b, (tuple, list)):
        return len(a) == len(b) and all(iguais(x, y, tol) for x, y in zip(a, b))
    if eh_num(a) and eh_num(b):
        return perto(a, b, tol)
    return a == b


def difs_grade(a, b, rotulo='', limite=8, tol=TOL_REL):
    """Lista as celulas diferentes de duas grades (para a evidencia)."""
    out = []
    if not isinstance(a, (tuple, list)) or not isinstance(b, (tuple, list)):
        return [] if iguais(a, b, tol) else [(rotulo, a, b)]
    if len(a) != len(b):
        out.append((rotulo + ' tamanho', len(a), len(b)))
    for i, (x, y) in enumerate(zip(a, b)):
        if isinstance(x, (tuple, list)):
            out += difs_grade(x, y, f'{rotulo}[{i}]', limite, tol)
        elif not iguais(x, y, tol):
            out.append((f'{rotulo}[{i}]', x, y))
        if len(out) >= limite:
            break
    return out[:limite]


def ev(r, *cols):
    return {c: (r or {}).get(c) for c in cols} if r is not None else None


EV_COLS = ('ID_REGISTRO', 'ORIGEM_RESULTADO', 'ANALITO', 'LOTE', 'NIVEL', 'DATA_HORA', 'RESULTADO', 'RUN',
           'CORRIDA_NO_DIA', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'ID_RELACIONADO',
           'MOTIVO_EXCLUSAO_AUTOMATICA')


def resumo(r):
    return ev(r, *EV_COLS)


def por_id(fin):
    d, dup = {}, []
    for r in fin:
        k = r['ID_REGISTRO']
        if k in d:
            dup.append(k)
        d[k] = r
    return d, dup


def diff_final(a, b):
    """{ID: [colunas diferentes]} entre dois snapshots {ID: linha}."""
    out = {}
    for k in set(a) | set(b):
        ra, rb = a.get(k), b.get(k)
        if ra is None or rb is None:
            out[k] = ['<linha ausente em um dos estados>']
            continue
        cols = [c for c in rb if not iguais(ra.get(c), rb.get(c))]
        if cols:
            out[k] = cols
    return out


def achados(qa, cod, id_=None):
    return [x for x in qa if x['CODIGO'] == cod and (id_ is None or x['ID_REGISTRO'] == id_)]


# =============================================================================
#  ORACULO DO RUN -- reimplementacao independente de pq/DB_CQ_FINAL.m (~linhas 316-368)
# =============================================================================
def _chave_nula(v):
    return (0, 0) if v is None or v == '' else (1, v)


def _gk(r):
    # Text.Combine ignora nulos: {EQUIPAMENTO, MATRIZ, LOTE, ANALITO, yyyyMMdd(DATA)}
    partes = [r.get('EQUIPAMENTO'), r.get('MATRIZ'), r.get('LOTE'), r.get('ANALITO'), r['DATA'].strftime('%Y%m%d')]
    return '|'.join(cstr(p) for p in partes if p is not None and p != '')


def oraculo_run(linhas_finais, gap_min):
    """{ID: (CORRIDA_NO_DIA, RUN)} pela regra escrita no ADR-057 e no PQ:
      Reais = ATIVO/INATIVADO/SEM_VALOR (fora incompleto, retransmissao e conflito);
      (1) ordena por _GK=(equip, matriz, lote, analito, dia), DATA_HORA, NIVEL, ITEM_ID, ID; bloco novo quando
          muda o _GK ou o intervalo para a linha anterior passa de GAP minutos;
      (2) dentro de (bloco, nivel), ordem por DATA_HORA, ITEM_ID, ID: o 2o resultado do mesmo nivel abre a
          posicao seguinte (_POS);
      (3) CORRIDA_NO_DIA = ordem de (bloco, _POS) no _GK; RUN = yymmdd*100 + corrida (nulo acima de 99).
    Contrato D16: linha sem DATA_HORA fica FORA do calculo da corrida (sem RUN)."""
    reais = [r for r in linhas_finais if r.get('STATUS_ANALITICO') in CORRIDA_REAL
             and isinstance(r.get('DATA_HORA'), dt.datetime) and isinstance(r.get('DATA'), dt.datetime)]
    for r in reais:
        r['_GK_ORC'] = _gk(r)
    ordem = sorted(reais, key=lambda r: (r['_GK_ORC'], r['DATA_HORA'], _chave_nula(r.get('NIVEL')),
                                         _chave_nula(r.get('ITEM_ID')), cstr(r['ID_REGISTRO'])))
    bl = {}
    ant = None
    for i, r in enumerate(ordem):
        if ant is None or r['_GK_ORC'] != ant['_GK_ORC'] or \
                (r['DATA_HORA'] - ant['DATA_HORA']).total_seconds() / 60.0 > gap_min:
            ini = i
        bl[id(r)] = ini
        ant = r
    por_bloco = collections.defaultdict(list)
    for r in ordem:
        por_bloco[bl[id(r)]].append(r)
    pos = {}
    for b, rs in por_bloco.items():
        rs2 = sorted(rs, key=lambda r: (_chave_nula(r.get('NIVEL')), r['DATA_HORA'],
                                        _chave_nula(r.get('ITEM_ID')), cstr(r['ID_REGISTRO'])))
        nv_ant, p = object(), 0
        for r in rs2:
            p = p + 1 if r.get('NIVEL') == nv_ant else 1
            nv_ant = r.get('NIVEL')
            pos[id(r)] = p
    segs = collections.defaultdict(set)
    for r in ordem:
        segs[r['_GK_ORC']].add((bl[id(r)], pos[id(r)]))
    corrida = {}
    for g, s in segs.items():
        for k, seg in enumerate(sorted(s), start=1):
            corrida[(g, seg)] = k
    out = {}
    for r in ordem:
        c = corrida[(r['_GK_ORC'], (bl[id(r)], pos[id(r)]))]
        run = None if c > 99 else int(r['DATA'].strftime('%y%m%d')) * 100 + c
        out[r['ID_REGISTRO']] = (c, run)
    for r in reais:
        r.pop('_GK_ORC', None)
    return out


def blocos_do_grupo(rows, gap_min):
    """Blocos (lista de listas) de um _GK, pela mesma regra do oraculo."""
    rs = sorted(rows, key=lambda r: (r['DATA_HORA'], _chave_nula(r.get('NIVEL')), _chave_nula(r.get('ITEM_ID')),
                                     cstr(r['ID_REGISTRO'])))
    out = []
    for r in rs:
        if not out or (r['DATA_HORA'] - out[-1][-1]['DATA_HORA']).total_seconds() / 60.0 > gap_min:
            out.append([r])
        else:
            out[-1].append(r)
    return out


# =============================================================================
#  contexto da sessao (COM)
# =============================================================================
class Ctx:
    def __init__(self, q, produto):
        self.q, self.produto = q, produto
        self.bio, self.nlv = q.bio, q.nlv
        self.pref = str(q.cfg('PREFIXO_ID') or '').strip().upper()
        self.gap = float(q.cfg('GAP_CORRIDA_MIN'))
        self.tol = float(q.cfg('TOLERANCIA_CONFLITO_MIN'))
        self.cadastro = [str(v[0] or '').strip() for v in q.ws('Analitos').Range('A4:A43').Value]
        self.tempos = []
        self.ids = collections.defaultdict(set)      # ID -> cenarios que o envolvem (matriz E16)
        self.usados = set()                          # analitos ja usados por algum cenario (MAIUSCULAS)
        self.tela = None
        self.agora = dt.datetime.now()
        self.hoje = dt.date.today()
        self.e12 = {}
        self.painel_orig = self.painel_ler()
        self._an = None                              # ID -> ANALITO (ADR-070: a Inativar pede o analito)

    # ---------------------------------------------------------------- atualizacao
    def atualizar(self, rotulo):
        resumo_, t = self.q.atualizar()
        self.q.ex.esperar()                           # o Excel termina o recalculo antes da proxima leitura
        self.tempos.append({'etapa': rotulo, 's': round(t, 1)})
        print(f'   ATUALIZAR [{rotulo}] {t:.1f} s', flush=True)
        return resumo_, t

    def motor(self):
        t0 = time.time()
        self.q.run('mEstatistica.AtualizarEstatistica')
        if self.bio:
            self.q.run('RecalcularEstatPeriodo')
        try:
            self.q.run('mIncerteza.RecalcularIncerteza')
        except Exception as e:                       # noqa: BLE001 -- registra e segue
            print('   aviso: mIncerteza.RecalcularIncerteza:', e, flush=True)
        return time.time() - t0

    # ---------------------------------------------------------------- Painel
    def _celulas_painel(self):
        return ('B3', 'G3', 'G4', 'H3') + (('L4',) if self.bio else ())

    def painel_ler(self):
        p = self.q.ws('Painel')
        return {a: p.Range(a).Formula for a in self._celulas_painel()}

    def painel_restaurar(self, f, motor=True):
        q = self.q
        q.ex.xl.EnableEvents = False
        try:
            p = q.ws('Painel')
            p.Unprotect(SENHA)
            for a, v in f.items():
                p.Range(a).Formula = v
        finally:
            q.ex.xl.EnableEvents = True
        if motor:
            self.motor()

    def configurar(self, analito=None, de=None, ate=None, lote=None, equip=None, motor=True):
        """Analito/periodo/lote/equipamento do Painel, com os eventos desligados (como o qa_final T05),
        e o motor rodado depois. Datas como NUMERO DE SERIE (o pywin32 desloca datetime ingenuo)."""
        q = self.q
        q.ex.xl.EnableEvents = False
        try:
            p = q.ws('Painel')
            p.Unprotect(SENHA)
            if analito is not None:
                p.Range('B3').Value = self.cadastro.index(analito) + 1
            if de is not None:
                p.Range('G3').Value = float(dia(de))
            if ate is not None:
                p.Range('G4').Value = float(dia(ate))
            if lote is not None:
                if lote == '':
                    p.Range('H3').ClearContents()
                else:
                    p.Range('H3').Value = str(lote)
            if equip is not None and self.bio:
                p.Range('L4').Value = equip
        finally:
            q.ex.xl.EnableEvents = True
        return self.motor() if motor else 0.0

    def lote_painel(self):
        v = self.q.wb.Application.Evaluate('loteAnalise')
        v = v.Value if hasattr(v, 'Value') else v
        return cstr(v).strip()

    def equip_filtro(self):
        # Bioquimica: o seletor do Painel. Hematologia: o EQUIPAMENTO_PADRAO do CFG (D12 -- antes "", todos)
        try:
            v = self.q.wb.Names('selEquipamento').RefersToRange.Value
        except Exception:                            # noqa: BLE001 -- Hematologia: nome = ""
            v = None
        if not str(v or '').strip():
            v = self.q.cfg('EQUIPAMENTO_PADRAO')
        return up(v)

    # ---------------------------------------------------------------- tabelas de entrada
    def cel(self, tabela, coluna, r):
        return self.q.lo(tabela).ListColumns(coluna).DataBodyRange.Cells(r, 1)

    def valor(self, tabela, coluna, r):
        return self.cel(tabela, coluna, r).Value

    def apagar_linha(self, tabela, r):
        """Apaga a linha como o usuario faz (seleciona e Delete): UM evento, so as celulas digitaveis
        (as colunas de formula de conferencia ficam). Celula a celula, com o evento ligado, a tabela
        manual geraria um MAN_ novo no meio da limpeza."""
        lo = self.q.lo(tabela)
        alvo = None
        for c in lo.ListColumns:
            cel = c.DataBodyRange.Cells(r, 1)
            if not cel.HasFormula:
                # Union, nao Range("A1,B1,..."): o endereco de varias areas passa de 255 caracteres e o Excel
                # recusa (0x800A03EC, medido na 1a execucao)
                alvo = cel if alvo is None else self.q.ex.xl.Union(alvo, cel)
        if alvo is not None:
            ws = lo.Parent                            # colunas carimbadas (DATA/USUARIO) sao travadas: a aba
            prot = bool(ws.ProtectContents)           # e liberada so para a limpeza, como o LiberarEscrita do VBA
            if prot:
                ws.Unprotect('qcini2025')
            try:
                alvo.ClearContents()
            finally:
                if prot:
                    ws.Protect('qcini2025', False, True, True, True)   # senha, desenhos, conteudo, cenarios, UI only

    def manual(self, quando, analito, lote, nivel, resultado, equip, matriz=None, id_=None, omitir=()):
        """Uma linha na aba Digitar Resultados. Devolve (linha da tabela, ID que o evento deixou na celula)."""
        v = {}
        if id_ is not None:
            v['ID_REGISTRO'] = id_
        d0 = meia_noite(quando)
        v['DATA'] = d0
        v['HORA'] = (quando - d0).total_seconds() / 86400.0
        v['EQUIPAMENTO'] = equip
        if matriz:
            v['MATRIZ'] = matriz
        v['LOTE'] = str(lote)
        v['NIVEL'] = int(nivel)
        v['ANALITO'] = analito
        v['RESULTADO'] = float(resultado) if eh_num(resultado) else resultado
        v['MOTIVO'] = 'Falha do interfaceamento'
        for c in omitir:
            v.pop(c, None)
        r = self.q.escrever_linha(TB_MAN, v)
        return r, self.valor(TB_MAN, 'ID_REGISTRO', r)

    def analito_de(self, id_):
        """ANALITO do resultado, como o usuario o le no grafico (tblCQ_Final) ou, para o manual ainda nao
        atualizado, na aba Digitar Resultados. None = ID que nao existe."""
        i = norm_id(id_, self.pref)
        if not i:
            return None
        for _ in range(2):
            if self._an is not None and i in self._an:
                return self._an[i]
            self._an = {}
            lo = self.q.lo(TB_FINAL)
            for a, b in zip(lo.ListColumns('ID_REGISTRO').DataBodyRange.Value, lo.ListColumns('ANALITO').DataBodyRange.Value):
                self._an[cstr(a[0]).strip().upper()] = b[0]
            for r in self.q.ler(TB_MAN):
                k = norm_id(r.get('ID_REGISTRO'), self.pref)
                if k and k not in self._an:
                    self._an[k] = r.get('ANALITO')
        return None

    def inativar(self, id_, comentario=None, analito='auto', motivo=None):
        """Uma linha na aba Inativar. ADR-070: com o ANALITO do resultado (como o usuario faz; 'auto' = o da
        tblCQ_Final/Digitar Resultados) e, se dado, o MOTIVO."""
        v = {'ID_REGISTRO': id_}
        an = self.analito_de(id_) if analito == 'auto' else analito
        if an not in (None, ''):
            v['ANALITO'] = an
        if motivo:
            v['MOTIVO'] = motivo
        r = self.q.escrever_linha(TB_INAT, v)
        rc = None
        if comentario:
            rc = self.q.escrever_linha(TB_COM, {'ID_REGISTRO': id_, 'COMENTARIO_TECNICO': comentario})
        return r, rc

    def equip_do_analito(self, fin, analito):
        if self.bio and self.tela:
            return self.tela['equip']
        c = collections.Counter(r['EQUIPAMENTO'] for r in fin if up(r['ANALITO']) == analito.upper() and r['EQUIPAMENTO'])
        if c:
            return c.most_common(1)[0][0]
        return str(self.q.cfg('EQUIPAMENTO_PADRAO') or '')

    def marcar(self, tag, *ids):
        for i in ids:
            if i:
                self.ids[i].add(tag)


# =============================================================================
#  leituras e snapshots
# =============================================================================
def ler_final(ctx):
    return ctx.q.ler(TB_FINAL)


def snap_eng(q):
    """Eng_Saida inteira que o motor publica (sem o carimbo G1), engPainel e engEstat."""
    eng = q.ws('Eng_Saida')
    return {'cab': tuple(eng.Range(a).Value for a in ('C1', 'E1', 'I1', 'K1', 'M1', 'O1', 'Q1', 'S1', 'U1')),
            'grade': eng.Range(eng.Cells(3, 1), eng.Cells(182, 38)).Value,
            'painel': q.nome('engPainel').Value,
            'estat': q.nome('engEstat').Value}


def snap_estat(q):
    return q.ws('Estatística').Range(f'A14:AU{13 + 40 * q.nlv}').Value


def estat_do_analito(q, analito):
    return [tuple(l) for l in snap_estat(q) if str(l[0] or '').strip().upper() == analito.upper()]


def snap_eventos(q, analito=None):
    ws = q.ws('Eventos_Westgard')
    n = iint(ws.Range('J2').Value) or 0
    if n <= 0:
        return []
    v = ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 14)).Value
    linhas = [tuple(l) for l in (v if isinstance(v[0], tuple) else (v,))]
    if analito:
        linhas = [l for l in linhas if str(l[2] or '').strip().upper() == analito.upper()]
    return linhas


def snap_qa(q):
    qa = q.ler(TB_QA)
    return qa, collections.Counter(tuple(r.values()) for r in qa), \
        collections.Counter((r['CODIGO'], r['ID_REGISTRO']) for r in qa)


def x_multiconjunto(e):
    """{(RUN, nivel, valor)} de todos os X vermelhos da Eng_Saida (3 slots x niveis)."""
    out = collections.Counter()
    for i, run in enumerate(e['runs']):
        linha = e['x'][i]
        for nv in range(len(linha) // 3):
            for s in range(3):
                v = linha[nv * 3 + s]
                if v not in (None, ''):
                    out[(run, nv + 1, v)] += 1
    return out


def painel_n(painel, nivel=1):
    return iint(painel[nivel - 1][1]) or 0


def contadores(fin, analito, lote):
    return collections.Counter((r['STATUS_ANALITICO'], r['TIPO_PLOTAGEM_LJ']) for r in fin
                               if up(r['ANALITO']) == analito.upper() and cstr(r['LOTE']).strip() == lote)


# =============================================================================
#  E07 -- reconciliacao global por CONJUNTO (invariante)
# =============================================================================
def derivar_manuais(man, pref):
    """Grupos {ID base -> [linhas]} pela regra do PQ: linha conta se tiver ID/DATA/HORA/LOTE/NIVEL/ANALITO/
    RESULTADO; ID fora do padrao MAN_nnnn -> MAN_INVALIDO_L<n>; o mesmo ID em varias linhas = UM ID base."""
    grupos = collections.OrderedDict()
    for n, r in enumerate(man, start=1):
        if not any(txt(r.get(c)) is not None for c in ('ID_REGISTRO', 'DATA', 'HORA', 'LOTE', 'NIVEL', 'ANALITO',
                                                         'RESULTADO')):
            continue
        idn = norm_id(r.get('ID_REGISTRO'), pref)
        base = idn if (idn and idn.startswith('MAN_') and digitos(idn[4:])) else f'MAN_INVALIDO_L{n}'
        grupos.setdefault(base, []).append(n)
    return grupos


def vencedor_manual(grupos, ids_final):
    """Para cada ID base, qual linha ficou com o ID sem sufixo: as outras tem de estar como <base>~L<n>.
    Devolve ({base: linha vencedora}, [divergencias])."""
    venc, div = {}, []
    for base, linhas in grupos.items():
        suf = {n for n in linhas if f'{base}~L{n}' in ids_final}
        sem = [n for n in linhas if n not in suf]
        if base not in ids_final or len(sem) != 1:
            div.append({'id_base': base, 'linhas': linhas, 'com_sufixo': sorted(suf), 'sem_sufixo_na_final': base in ids_final})
        else:
            venc[base] = sem[0]
    return venc, div


def filtro_estatistica(ctx):
    """Reproduz o recorte da aba Estatistica (lido do VBA, nao presumido):
      Hematologia: C..F = INDEX(engEstat) <- mEstatistica.AtualizarEstatisticaAba -> EstatBasica:
        PARTICIPA=SIM, lote = Estatistica!B4 (exato, vazio = todos), ANO(DATA_HORA) entre B3 e D3 (B3=0: todos),
        equipamento = selEquipamento (vazio = todos), RESULTADO numerico; DP de duas passadas.
      Bioquimica: C..F = EstatPeriodo(analito, nivel, ..., Estat_Ini_Efetiva, Estat_Fim_Efetiva, Estat_Exclusoes,
        $H$4) -> mEstatPeriodo.Agregar: PARTICIPA=SIM, equipamento do Painel, DIA(DATA_HORA) em [ini, fim]
        (0 = sem limite), fora das janelas de exclusao, lote exato; DP por soma de quadrados.
    Devolve (funcao linha -> chave (ANALITO, nivel) ou None, descricao)."""
    q = ctx.q
    e = q.ws('Estatística')
    eq = ctx.equip_filtro()
    if not q.bio:
        def val(v):
            try:
                return int(float(v))
            except (TypeError, ValueError):
                return 0
        ano_de, ano_ate = val(e.Range('B3').Value), val(e.Range('D3').Value)
        if ano_ate < ano_de:
            ano_ate = ano_de
        lote = cstr(e.Range('B4').Value).strip()

        def pred(r, recorte_so=False):
            an = cstr(r['ANALITO']).strip()
            if not an or (not recorte_so and up(r['PARTICIPA_ESTATISTICA']) != SIM):
                return None
            if eq and up(r['EQUIPAMENTO']) != eq:
                return None
            if lote and cstr(r['LOTE']).strip() != lote:
                return None
            if not eh_num(r['RESULTADO']):
                return None
            if ano_de != 0:
                dh = r['DATA_HORA']
                if not isinstance(dh, dt.datetime) or not (ano_de <= dh.year <= ano_ate):
                    return None
            return (an.upper(), iint(r['NIVEL']))
        return pred, {'produto': 'Hematologia', 'ano_de': ano_de, 'ano_ate': ano_ate, 'lote': lote, 'equip': eq}

    ini = como_data(e.Range('H5').Value)
    fim = como_data(e.Range('J5').Value)
    exc = []
    try:
        for a, b in q.wb.Names('Estat_Exclusoes').RefersToRange.Value:
            a, b = como_data(a), como_data(b)
            if a > 0 and b > 0 and b >= a:
                exc.append((a, b))
    except Exception:                                # noqa: BLE001 -- sem exclusoes
        pass
    lote = cstr(e.Range('H4').Value).strip()

    def pred(r, recorte_so=False):
        an = cstr(r['ANALITO']).strip()
        if not an or (not recorte_so and up(r['PARTICIPA_ESTATISTICA']) != SIM):
            return None
        if eq and up(r['EQUIPAMENTO']) != eq:
            return None
        dh = r['DATA_HORA']
        if not isinstance(dh, dt.datetime) or not eh_num(r['RESULTADO']):
            return None
        d = dia(dh)
        if (ini > 0 and d < ini) or (fim > 0 and d > fim):
            return None
        if any(a <= d <= b for a, b in exc):
            return None
        if lote and cstr(r['LOTE']).strip() != lote:
            return None
        return (an.upper(), cstr(r['NIVEL']).strip())
    return pred, {'produto': 'Bioquimica', 'ini': ini, 'fim': fim, 'exclusoes': exc, 'lote': lote, 'equip': eq}


def estat_independente(ctx, fin):
    pred, desc = filtro_estatistica(ctx)
    pops = collections.defaultdict(list)
    for r in fin:
        k = pred(r)
        if k is not None:
            pops[k].append(float(r['RESULTADO']))
    out = {}
    for k, v in pops.items():
        n = len(v)
        if not ctx.bio:                              # EstatBasica: media e DP de duas passadas
            media = sum(v) / n
            s = sum((x - media) ** 2 for x in v) if n >= 2 else 0.0
            dp = math.sqrt(s / (n - 1)) if n >= 2 and s > 0 else 0.0
        else:                                        # mEstatPeriodo: soma e soma de quadrados
            soma = somaq = 0.0
            for x in v:
                soma += x
                somaq += x * x
            media = soma / n
            num = somaq - n * media * media if n >= 2 else 0.0
            dp = math.sqrt(max(num, 0.0) / (n - 1)) if n >= 2 else 0.0
        cv = dp / media * 100.0 if n >= 2 and media != 0 else ''
        out[k] = (n, media if n else '', dp if n >= 2 else '', cv)
    return out, pred, desc


def conferir_estatistica(ctx, fin, etapa):
    q = ctx.q
    ind, pred, desc = estat_independente(ctx, fin)
    grade = q.ws('Estatística').Range(f'A14:F{13 + 40 * q.nlv}').Value
    div, n_lin, n_pos = [], 0, 0
    for i, l in enumerate(grade):
        a, b = l[0], l[1]
        if a in (None, '') or a == 0:
            continue
        n_lin += 1
        k = (str(a).strip().upper(), iint(b) if not ctx.bio else cstr(b).strip())
        n, m, s, cv = ind.get(k, (0, '', '', ''))
        n_aba = iint(l[2]) if l[2] not in (None, '') else 0
        ok = n_aba == n and perto(l[3], m) and perto(l[4], s) and perto(l[5], cv)
        if n:
            n_pos += 1
        if not ok:
            div.append({'linha': 14 + i, 'analito': a, 'nivel': b, 'aba': l[2:6], 'independente': (n, m, s, cv)})
    ok = not div and n_pos > 0
    reg(f'E07 [{etapa}] Estatística (A14:F) = recálculo independente em TODAS as linhas analito×nível '
        f'(n exato; média, DP e CV com tolerância relativa 1e-9)' + ('' if n_pos else ' -- pré-condição ausente: nenhuma linha com n>0'),
        ok, {'linhas_conferidas': n_lin, 'linhas_com_n>0': n_pos, 'recorte': desc, 'divergencias': div[:10],
             'n_divergencias': len(div)})
    return ok


def reconciliar_global(ctx, etapa, fin=None):
    """E07: identidades do mapa real, por CONJUNTO de IDs, e a Estatistica recalculada. Chamada no estado-base e
    depois de cada etapa. Devolve a final lida (para reaproveitar)."""
    q = ctx.q
    t0 = time.time()
    fin = fin if fin is not None else ler_final(ctx)
    rec = q.ler(TB_REC)
    man = q.ler(TB_MAN)
    inat = q.ler(TB_INAT)
    ids = [r['ID_REGISTRO'] for r in fin]
    F = set(ids)
    dup = [k for k, v in collections.Counter(ids).items() if v > 1]
    vazios = sum(1 for i in ids if i in (None, ''))
    R = {r['ID_REGISTRO'] for r in rec if r['ID_REGISTRO'] not in (None, '')}
    grupos = derivar_manuais(man, ctx.pref)
    venc, div_man = vencedor_manual(grupos, F)
    M = set()
    for base, linhas in grupos.items():
        M.add(base)
        for n in linhas:
            if venc.get(base) != n:
                M.add(f'{base}~L{n}')
    esperado = R | M
    faltam = sorted(esperado - F)
    sobram = sorted(F - esperado)
    st = collections.Counter(r['STATUS_ANALITICO'] for r in fin)
    fora_part = {k: v for k, v in st.items() if k not in STATUS}
    s_sim = {r['ID_REGISTRO'] for r in fin if r['PARTICIPA_ESTATISTICA'] == SIM}
    s_ativo = {r['ID_REGISTRO'] for r in fin if r['STATUS_ANALITICO'] == 'ATIVO'}
    s_normal = {r['ID_REGISTRO'] for r in fin if r['TIPO_PLOTAGEM_LJ'] == 'NORMAL'}
    s_x = {r['ID_REGISTRO'] for r in fin if r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'}
    s_x_esp = {r['ID_REGISTRO'] for r in fin if r['STATUS_ANALITICO'] == 'INATIVADO'
               and r['REGISTRAR_RESULTADO_NO_LJ'] == SIM}
    s_reg = {r['ID_REGISTRO'] for r in fin if r['INATIVACAO_REGISTRADA'] == SIM}
    # ADR-070: vale a 1a linha de cada ID na Inativar, e so com o ANALITO igual ao do resultado
    an_fin = {r['ID_REGISTRO']: up(r['ANALITO']) for r in fin}
    s_reg_esp, vistos_inat = set(), set()
    for r in inat:
        i = norm_id(r['ID_REGISTRO'], ctx.pref)
        if not i or i in vistos_inat:
            continue
        vistos_inat.add(i)
        if i in F and txt(r.get('ANALITO')) is not None and up(r.get('ANALITO')) == an_fin[i]:
            s_reg_esp.add(i)
    plot_inval = sorted({r['TIPO_PLOTAGEM_LJ'] for r in fin} - {'NORMAL', 'X_VERMELHO', 'NAO_PLOTAR'}, key=str)
    checagens = {
        'ids_unicos_e_nao_vazios': not dup and vazios == 0,
        'final_igual_rec_uniao_manuais': not faltam and not sobram,
        'manuais_um_id_sem_sufixo_por_grupo': not div_man,
        'particao_status_exata': not fora_part and sum(st.values()) == len(fin),
        'SIM=ATIVO=NORMAL': s_sim == s_ativo == s_normal,
        'X=INATIVADO_e_LJ': s_x == s_x_esp,
        'plotagem_valida': not plot_inval,
        'INATIVACAO_REGISTRADA=NormId(Inativar)∩Final∩analito': s_reg == s_reg_esp,
    }
    ok = all(checagens.values()) and len(fin) > 0 and len(R) > 0
    reg(f'E07 [{etapa}] Reconciliação por conjunto: IDs(Final) = Recebimento(ID≠nulo) ∪ manuais derivados (NormId, '
        f'~L<n>, MAN_INVALIDO_L<n>), IDs únicos, partição exata por STATUS, {{SIM}}={{ATIVO}}={{NORMAL}}, '
        f'{{X}}={{INATIVADO∧LJ}}, {{INATIVACAO_REGISTRADA}}=NormId(Inativar)∩Final' +
        ('' if len(fin) and len(R) else ' -- pré-condição ausente: final ou recebimento vazio'),
        ok, {'checagens': checagens, 'linhas_final': len(fin), 'ids_recebimento': len(R), 'ids_manuais_derivados': len(M),
             'ids_duplicados': dup[:20], 'ids_vazios': vazios, 'faltam_na_final': faltam[:20], 'sobram_na_final': sobram[:20],
             'manuais_divergentes': div_man[:10], 'status': dict(st), 'status_fora_da_lista': fora_part,
             'SIM-ATIVO': sorted(s_sim ^ s_ativo)[:10], 'ATIVO-NORMAL': sorted(s_ativo ^ s_normal)[:10],
             'X_divergentes': sorted(s_x ^ s_x_esp)[:10], 'INAT_REG_divergentes': sorted(s_reg ^ s_reg_esp)[:10],
             'plotagem_invalida': plot_inval, 'tempo_leitura_s': round(time.time() - t0, 1)})
    conferir_estatistica(ctx, fin, etapa)
    return fin


# =============================================================================
#  tela do Painel (o mesmo criterio do qa_final T05)
# =============================================================================
def escolher_tela(ctx, fin):
    q = ctx.q
    lote = ctx.lote_painel()
    eq = str(q.ws('Painel').Range('L4').Value or '').strip() if ctx.bio else None
    cont = collections.Counter(cstr(r['ANALITO']).strip() for r in fin if cstr(r['LOTE']).strip() == lote
                               and r['PARTICIPA_ESTATISTICA'] == SIM and (eq is None or up(r['EQUIPAMENTO']) == eq.upper()))
    an = next((a for a, _ in cont.most_common() if a in ctx.cadastro), None)
    if an is None:
        return None
    dts = [r['DATA_HORA'] for r in fin if cstr(r['ANALITO']).strip() == an and cstr(r['LOTE']).strip() == lote
           and r['PARTICIPA_ESTATISTICA'] == SIM and isinstance(r['DATA_HORA'], dt.datetime)
           and (eq is None or up(r['EQUIPAMENTO']) == eq.upper())]
    t = ctx.configurar(analito=an, de=min(dts), ate=max(dts))
    ctx.tela = {'analito': an, 'lote': lote, 'equip': eq, 'de': min(dts), 'ate': max(dts)}
    ctx.tela_painel = ctx.painel_ler()
    ctx.usados.add(an.upper())
    print(f'   tela: {ctx.tela} (motor {t:.1f} s)', flush=True)
    return ctx.tela


def voltar_tela(ctx, motor=True):
    ctx.painel_restaurar(ctx.tela_painel, motor=motor)


def corridas_motor(fin, analito, lote, equip):
    """Eixo de corridas do motor (AtualizarCalc): uniao das corridas SIM e X do (analito, lote[, equip]),
    data = 1a linha vista (ordem da tabela), ordem por (data, RUN). Lista de (serial, RUN)."""
    prim = {}
    for passada in ('SIM', 'X'):
        for r in fin:
            if up(r['ANALITO']) != analito.upper() or cstr(r['LOTE']).strip() != lote:
                continue
            if equip and up(r['EQUIPAMENTO']) != equip.upper():
                continue
            sel = r['PARTICIPA_ESTATISTICA'] == SIM if passada == 'SIM' else \
                (r['PARTICIPA_ESTATISTICA'] != SIM and r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO')
            if not sel:
                continue
            k = iint(r['RUN']) or 0
            if k not in prim:
                prim[k] = (serial(r['DATA_HORA']) if isinstance(r['DATA_HORA'], dt.datetime) else 0.0, k)
    return sorted(prim.values())


def linhas_do_dia(fin, analito, lote, d):
    return [r for r in fin if up(r['ANALITO']) == analito.upper() and cstr(r['LOTE']).strip() == lote
            and isinstance(r['DATA_HORA'], dt.datetime) and r['DATA_HORA'].date() == d]


def dia_livre(ctx, fin, analito, evitar=()):
    """Dia mais recente ANTES de hoje sem nenhum resultado (interface ou manual) do analito."""
    ocup = {r['DATA_HORA'].date() for r in fin if up(r['ANALITO']) == analito.upper()
            and isinstance(r['DATA_HORA'], dt.datetime)}
    d = ctx.hoje - dt.timedelta(days=1)
    for _ in range(800):
        if d not in ocup and d not in evitar:
            return d
        d -= dt.timedelta(days=1)
    return None


def analitos_candidatos(ctx, fin, lote, equip):
    c = collections.Counter(cstr(r['ANALITO']).strip() for r in fin if cstr(r['LOTE']).strip() == lote
                            and r['PARTICIPA_ESTATISTICA'] == SIM and (not equip or up(r['EQUIPAMENTO']) == equip.upper()))
    return [a for a, _ in c.most_common() if a in ctx.cadastro and a.upper() not in ctx.usados]


# =============================================================================
#  E08 -- chaves na aba Inativar
# =============================================================================
def e08(ctx, fin0):
    q, pref = ctx.q, ctx.pref
    cand, vistos = [], set()
    for r in sorted((r for r in fin0 if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO'
                     and isinstance(r['DATA_HORA'], dt.datetime) and str(r['ID_REGISTRO']).startswith(pref + '-')
                     and up(r['ANALITO']) not in ctx.usados), key=lambda r: r['DATA_HORA']):
        if up(r['ANALITO']) in vistos:
            continue
        vistos.add(up(r['ANALITO']))
        cand.append(r)
        if len(cand) == 4:
            break
    if len(cand) < 4:
        reg('E08 Chaves na aba Inativar -- pré-condição ausente: menos de 4 resultados ATIVO do interfaceamento',
            False, {'encontrados': [c['ID_REGISTRO'] for c in cand]})
        return
    A, B, C, D = (c['ID_REGISTRO'] for c in cand)
    num = lambda i: i[len(pref) + 1:]              # noqa: E731
    ids_final0 = {r['ID_REGISTRO'] for r in fin0}
    tem_man1 = 'MAN_0001' in ids_final0
    lin = {}
    an_c = {c['ID_REGISTRO']: c['ANALITO'] for c in cand}               # ADR-070: a linha leva o analito
    lin['A'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': f'{pref.lower()}-{num(A)} ', 'ANALITO': an_c[A]})
    lin['B'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': num(B), 'ANALITO': an_c[B]})
    # '00<C>': a coluna ID e texto ('@'); confere SEM evento que a celula guardou o texto com os zeros
    q.ex.xl.EnableEvents = False
    try:
        lin['C'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': '00' + num(C)})
        cru_c = ctx.valor(TB_INAT, 'ID_REGISTRO', lin['C'])
    finally:
        q.ex.xl.EnableEvents = True
    ctx.cel(TB_INAT, 'ID_REGISTRO', lin['C']).Value = '00' + num(C)          # agora com o evento
    ctx.cel(TB_INAT, 'ANALITO', lin['C']).Value = an_c[C]
    lin['D'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': f'{pref}- {num(D)}', 'ANALITO': an_c[D]})
    inexistente = f'{pref}-999999999'
    lin['X'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': inexistente})
    lin['A2'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': f'{pref}-{num(A)}', 'ANALITO': an_c[A]})
    an_m = ctx.analito_de('MAN_0001') if tem_man1 else None
    lin['M'] = q.escrever_linha(TB_INAT, dict({'ID_REGISTRO': 'man_1'}, **({'ANALITO': an_m} if an_m else {})))
    celulas = {k: ctx.valor(TB_INAT, 'ID_REGISTRO', r) for k, r in lin.items()}
    ctx.marcar('E08', A, B, C, D, inexistente, 'MAN_0001')
    ctx.atualizar('E08')
    fin1 = ler_final(ctx)
    f1, _ = por_id(fin1)
    _, _, qa_ids = snap_qa(q)
    qa = q.ler(TB_QA)
    e02 = achados(qa, 'E02')
    e03 = achados(qa, 'E03')

    def inat_x(i):
        r = f1.get(i)
        return r is not None and r['STATUS_ANALITICO'] == 'INATIVADO' and r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
    reg('E08a Inativar "hem-<A> " (minúsculo + espaço) e "<B>" (número puro): INATIVADO/X_VERMELHO e a célula '
        'normalizada pelo evento', inat_x(A) and inat_x(B) and celulas['A'] == A and celulas['B'] == B,
        {'A': resumo(f1.get(A)), 'B': resumo(f1.get(B)), 'celula_A': celulas['A'], 'celula_B': celulas['B']})
    reg('E08b A repetido (2ª linha): E02 no QA e uma única linha de A na final',
        bool([x for x in e02 if x['ID_REGISTRO'] == A]) and sum(1 for r in fin1 if r['ID_REGISTRO'] == A) == 1,
        {'E02': [x['DETALHE'] for x in e02 if x['ID_REGISTRO'] == A], 'linhas_A': sum(1 for r in fin1 if r['ID_REGISTRO'] == A)})
    reg(f'E08c Inexistente ({inexistente}): E03 no QA, sem efeito na final',
        bool([x for x in e03 if x['ID_REGISTRO'] == inexistente]) and inexistente not in f1,
        {'E03': [x['DETALHE'] for x in e03 if x['ID_REGISTRO'] == inexistente][:2]})
    pre_c = isinstance(cru_c, str) and cru_c == '00' + num(C)
    reg('E08d Zeros à esquerda ("00<C>") casam com o ID canônico após D14 (hoje falha: prova D14)'
        + ('' if pre_c else ' -- pré-condição ausente: a célula não guardou o texto com zeros'),
        pre_c and inat_x(C),
        {'C': C, 'celula_sem_evento': cru_c, 'celula_apos_evento': celulas['C'], 'final_C': resumo(f1.get(C)),
         'E03_do_C': [x['ID_REGISTRO'] for x in e03 if num(C) in str(x['ID_REGISTRO'])]})
    reg('E08e NBSP interno ("HEM-\\u00a0<D>") casa com o ID canônico após D14 (hoje falha: prova D14)', inat_x(D),
        {'D': D, 'celula_apos_evento': repr(celulas['D']), 'final_D': resumo(f1.get(D)),
         'E03_do_D': [repr(x['ID_REGISTRO']) for x in e03 if num(D) in str(x['ID_REGISTRO'])]})
    if tem_man1:
        ok_m = inat_x('MAN_0001')
        evm = resumo(f1.get('MAN_0001'))
    else:
        ok_m = bool([x for x in e03 if x['ID_REGISTRO'] == 'MAN_0001']) and 'MAN_0001' not in f1
        evm = {'E03': [x['DETALHE'] for x in e03 if x['ID_REGISTRO'] == 'MAN_0001']}
    reg(f'E08f "man_1" é a chave MAN_0001 ({"existe: INATIVADO" if tem_man1 else "não existe: E03"})', ok_m,
        {'celula': celulas['M'], 'MAN_0001_existia': tem_man1, 'resultado': evm})
    ids1 = {r['ID_REGISTRO'] for r in fin1}
    reg('E08g |Final| e o conjunto de IDs não mudam com a digitação de chaves', len(fin1) == len(fin0) and ids1 == ids_final0,
        {'linhas': [len(fin0), len(fin1)], 'novos': sorted(ids1 - ids_final0)[:5], 'sumidos': sorted(ids_final0 - ids1)[:5]})
    reconciliar_global(ctx, 'E08', fin1)
    # limpeza: C, D e man_1 sairiam do invariante E07 (contrato D14) e o man_1 contaminaria o proximo MAN_
    for k in ('C', 'D', 'M'):
        ctx.apagar_linha(TB_INAT, lin[k])
    ctx.e08 = {'A': A, 'B': B, 'linhas': lin}
    ctx.usados |= {up(c['ANALITO']) for c in cand}


# =============================================================================
#  E09 -- chaves na aba Digitar Resultados (+ variante D06 do E10 no mesmo refresh)
# =============================================================================
def ids_usados(ctx, fin):
    q = ctx.q
    s = {r['ID_REGISTRO'] for r in fin}
    for t in (TB_MAN, TB_INAT, TB_COM):
        for r in q.ler(t):
            i = norm_id(r.get('ID_REGISTRO'), ctx.pref)
            if i:
                s.add(i)
    return s


def proximo_livre(usados, k):
    while f'MAN_{k:04d}' in usados:
        k += 1
    return k


def e09(ctx, fin0):
    q = ctx.q
    lote, eqp = ctx.tela['lote'], ctx.tela['equip']
    a9 = next(iter(analitos_candidatos(ctx, fin0, lote, eqp)), None)
    d9 = dia_livre(ctx, fin0, a9) if a9 else None
    if not a9 or not d9:
        reg('E09 Chaves na aba Digitar Resultados -- pré-condição ausente: analito/dia sem interface', False,
            {'analito': a9, 'dia': d9})
        return
    ctx.usados.add(a9.upper())
    eq9 = ctx.equip_do_analito(fin0, a9)
    vals = sorted(r['RESULTADO'] for r in fin0 if up(r['ANALITO']) == a9.upper() and iint(r['NIVEL']) == 1
                  and r['PARTICIPA_ESTATISTICA'] == SIM and eh_num(r['RESULTADO']))
    v9 = round(vals[len(vals) // 2], 4) if vals else 1.0
    h = lambda hh: dt.datetime.combine(d9, dt.time(hh, 7))   # noqa: E731
    man = lambda hh, **kw: ctx.manual(h(hh), a9, lote, 1, kw.pop('res', v9), eq9, **kw)   # noqa: E731
    usados0 = ids_usados(ctx, fin0)
    real = next(r for r in fin0 if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['STATUS_ANALITICO'] == 'ATIVO'
                and up(r['ANALITO']) != a9.upper())
    x_real = real['ID_REGISTRO']
    # (a) ID vazio -> MAN_n automatico
    ra, id_a = man(1)
    # (b) man_k -> MAN_000k
    k7 = proximo_livre(usados0 | {norm_id(id_a, ctx.pref)}, 7)
    rb, id_b = man(2, id_=f'man_{k7}')
    # (c) ID do interfaceamento digitado num manual: com prefixo e como numero puro
    rc1, id_c1 = man(3, id_=x_real)
    rc2, id_c2 = man(4, id_=x_real[len(ctx.pref) + 1:])
    # (d) duas linhas completas com o mesmo MAN_
    k50 = proximo_livre(usados0, 50)
    id50 = f'MAN_{k50:04d}'
    rd1, _ = man(5, id_=id50)
    rd2, _ = man(6, id_=id50)
    # (e) MAN_001 alem de um MAN_0001 existente
    # alvo = o ID automatico do (a) (com D05 nao e mais necessariamente MAN_0001); digitado com 3 digitos
    alvo_e = id_a if (isinstance(id_a, str) and id_a.startswith('MAN_') and id_a[4:].isdigit()) else 'MAN_0001'
    curto_e = 'MAN_' + str(int(alvo_e[4:])).zfill(3)
    man_tab = q.ler(TB_MAN)
    lin_0001 = [n for n, r in enumerate(man_tab, 1) if norm_id(r.get('ID_REGISTRO'), ctx.pref) == alvo_e]
    re_, id_e = man(7, id_=curto_e)
    # (f) MAN_0060 incompleto na linha k e completo na linha k+1
    k60 = proximo_livre(usados0 | {f'MAN_{k50:04d}'}, 60)
    id60 = f'MAN_{k60:04d}'
    rf1, _ = man(8, id_=id60, omitir=('RESULTADO',))
    rf2, _ = man(9, id_=id60)
    # (g) MAN_k como ultimo, inativado com comentario
    rg, id_g = man(10)
    ctx.inativar(id_g, 'QA ETL E09g: manual inativado antes de a linha ser apagada.')
    ctx.marcar('E09', id_a, id_b, id50, f'{id50}~L{rd2}', alvo_e, f'{alvo_e}~L{re_}', id60,
               f'{id60}~L{rf1}', f'{id60}~L{rf2}', id_g, x_real, f'MAN_INVALIDO_L{rc1}', f'MAN_INVALIDO_L{rc2}')
    ctx.atualizar('E09 #1')
    fin1 = ler_final(ctx)
    f1, _ = por_id(fin1)
    qa = q.ler(TB_QA)
    reg('E09a Manual com ID vazio recebe MAN_n automático e fica ATIVO',
        str(id_a or '').startswith('MAN_') and (f1.get(id_a) or {}).get('STATUS_ANALITICO') == 'ATIVO',
        {'id': id_a, 'final': resumo(f1.get(id_a)), 'dia': d9, 'analito': a9})
    id_b_esp = f'MAN_{k7:04d}'
    reg(f'E09b "man_{k7}" vira {id_b_esp} (evento e PQ) e fica ATIVO',
        id_b == id_b_esp and (f1.get(id_b_esp) or {}).get('STATUS_ANALITICO') == 'ATIVO',
        {'celula': id_b, 'final': resumo(f1.get(id_b_esp))})
    inv1, inv2 = f'MAN_INVALIDO_L{rc1}', f'MAN_INVALIDO_L{rc2}'
    ok_c = all((f1.get(i) or {}).get('STATUS_ANALITICO') == 'MANUAL_INCOMPLETO' and achados(qa, 'E04', i)
               for i in (inv1, inv2)) and all(iguais(tuple(real.values()), tuple(f1.get(x_real, {}).values())) for _ in [0])
    reg('E09c ID do interfaceamento digitado no manual (com prefixo e número puro): MAN_INVALIDO_L<n>, '
        'MANUAL_INCOMPLETO, E04; o resultado real fica intacto', ok_c,
        {'celulas': [id_c1, id_c2], 'finais': [resumo(f1.get(inv1)), resumo(f1.get(inv2))],
         'real_intacto': iguais(tuple(real.values()), tuple(f1.get(x_real, {}).values())), 'real': x_real})
    dup2 = f'{id50}~L{rd2}'
    reg(f'E09d Duas linhas completas {id50}: a 1ª ATIVO, a 2ª "{dup2}" CONFLITO_MANUAL + A01',
        (f1.get(id50) or {}).get('STATUS_ANALITICO') == 'ATIVO'
        and (f1.get(dup2) or {}).get('STATUS_ANALITICO') == 'CONFLITO_MANUAL' and bool(achados(qa, 'A01', dup2)),
        {'primeira': resumo(f1.get(id50)), 'segunda': resumo(f1.get(dup2)), 'A01': len(achados(qa, 'A01', dup2))})
    if lin_0001:
        tarde = max(lin_0001[0], re_)
        ok_e = alvo_e in f1 and f'{alvo_e}~L{tarde}' in f1 and \
            f1[f'{alvo_e}~L{tarde}']['STATUS_ANALITICO'] == 'CONFLITO_MANUAL'
        reg('E09e "MAN_001" além de um MAN_0001 existente: o mesmo ID; a linha posterior vira MAN_0001~L<n> '
            'CONFLITO_MANUAL', ok_e and id_e == alvo_e,
            {'celula': id_e, 'linha_existente': lin_0001[0], 'linha_nova': re_,
             'sufixado': resumo(f1.get(f'{alvo_e}~L{tarde}')), 'alvo': alvo_e, 'digitado': curto_e})
    else:
        reg('E09e "MAN_001" além de um MAN_0001 existente -- pré-condição ausente: não há MAN_0001 na tabela manual',
            False, {'celula': id_e, 'id_a': id_a})
    comp, inc = f1.get(id60), f1.get(f'{id60}~L{rf1}')
    reg(f'E09f {id60} incompleto na linha {rf1} e completo na linha {rf2}: a COMPLETA vale (ATIVO) e a incompleta '
        f'vira {id60}~L{rf1} (hoje falha: prova D08)',
        rf2 == rf1 + 1 and comp is not None and comp['STATUS_ANALITICO'] == 'ATIVO' and eh_num(comp['RESULTADO'])
        and inc is not None and inc['STATUS_ANALITICO'] in ('MANUAL_INCOMPLETO', 'CONFLITO_MANUAL'),
        {'linhas': [rf1, rf2], id60: resumo(comp), f'{id60}~L{rf1}': resumo(inc),
         f'{id60}~L{rf2}': resumo(f1.get(f'{id60}~L{rf2}'))})
    pre_g = (f1.get(id_g) or {}).get('STATUS_ANALITICO') == 'INATIVADO'
    reconciliar_global(ctx, 'E09 #1', fin1)

    # (g) apaga a linha MAN_k e digita um manual novo com ID vazio; + variante D06 do E10 (mesmo refresh)
    usados_g = ids_usados(ctx, fin1)
    ctx.apagar_linha(TB_MAN, rg)
    rg2, id_g2 = man(11)
    A, B = ctx.e08['A'], ctx.e08['B'] if hasattr(ctx, 'e08') else (None, None)
    lB = ctx.e08['linhas']['B'] if hasattr(ctx, 'e08') else None
    lA = ctx.e08['linhas']['A'] if hasattr(ctx, 'e08') else None
    talvez_ok = None
    if lB:
        ctx.cel(TB_INAT, COL_LJ, lB).ClearContents()           # caixa vazia numa inativacao ja carimbada
    if lA:
        try:
            ctx.cel(TB_INAT, COL_LJ, lA).Value = 'TALVEZ'       # valor fora da lista
            talvez_ok = ctx.valor(TB_INAT, COL_LJ, lA)
        except Exception as e:                                  # noqa: BLE001
            talvez_ok = f'ERRO: {e}'
    ctx.marcar('E09g', id_g2)
    ctx.atualizar('E09 #2 (+E10 D06)')
    fin2 = ler_final(ctx)
    f2, _ = por_id(fin2)
    qa = q.ler(TB_QA)
    novo = f2.get(id_g2)
    reg('E09g Apagar o último manual (inativado e comentado) e digitar outro: o novo recebe um MAN_ NUNCA usado e '
        'não herda inativação/comentário (hoje falha: prova D05)'
        + ('' if pre_g else ' -- pré-condição ausente: o MAN_k não ficou INATIVADO antes de ser apagado'),
        pre_g and id_g2 not in usados_g and novo is not None and novo['STATUS_ANALITICO'] == 'ATIVO'
        and novo['INATIVACAO_REGISTRADA'] == NAO and novo['TEM_JUSTIFICATIVA'] == NAO,
        {'MAN_k_apagado': id_g, 'novo_id': id_g2, 'novo_ja_usado': id_g2 in usados_g, 'novo_final': resumo(novo),
         'inativacao_herdada': (novo or {}).get('INATIVACAO_REGISTRADA'), 'comentario_herdado': (novo or {}).get('COMENTARIO_TECNICO')})
    if lB:
        rB = f2.get(B)
        reg('E10v D06 "REGISTRAR - LJ" vazio numa inativação já carimbada = NÃO (NAO_PLOTAR), coerente com a caixa '
            'desmarcada (hoje falha: prova D06)',
            rB is not None and rB['STATUS_ANALITICO'] == 'INATIVADO' and rB['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
            and rB['REGISTRAR_RESULTADO_NO_LJ'] == NAO,
            {'id': B, 'celula_LJ': ctx.valor(TB_INAT, COL_LJ, lB), 'DATA_INATIVACAO': ctx.valor(TB_INAT, 'DATA_INATIVACAO', lB),
             'final': resumo(rB)})
    if lA:
        rA = f2.get(A)
        a08 = achados(qa, 'A08', A)
        reg('E10v D06 "REGISTRAR - LJ" com valor não reconhecido ("TALVEZ"): NAO_PLOTAR e ALERTA A08 '
            '(hoje falha: prova D06)'
            + ('' if talvez_ok == 'TALVEZ' else ' -- pré-condição ausente: a célula não aceitou o texto'),
            talvez_ok == 'TALVEZ' and rA is not None and rA['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and bool(a08),
            {'id': A, 'celula_LJ': talvez_ok, 'final': resumo(rA), 'A08': [x['DETALHE'] for x in a08]})
    ctx.marcar('E10v', A, B)
    reconciliar_global(ctx, 'E09 #2', fin2)
    ctx.e09 = {'analito': a9, 'dia': d9}
    return fin2


# =============================================================================
#  E10 -- inativacao e caixa, linha a linha (SIM -> NAO -> SIM)
# =============================================================================
def e10(ctx, fin0):
    q = ctx.q
    an, lote, eqp = ctx.tela['analito'], ctx.tela['lote'], ctx.tela['equip']
    e0 = q.eng_saida()
    f0, _ = por_id(fin0)
    alvo = None
    for i in range(e0['n'] - 1, -1, -1):
        v = e0['val'][i]
        if e0['fil'][i] == 1 and isinstance(v[0], float) and any(isinstance(x, float) for x in v[1:]):
            cands = [r for r in fin0 if up(r['ANALITO']) == an.upper() and cstr(r['LOTE']).strip() == lote
                     and iint(r['NIVEL']) == 1 and iint(r['RUN']) == e0['runs'][i] and r['PARTICIPA_ESTATISTICA'] == SIM
                     and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and r['RESULTADO'] == v[0]
                     and (not ctx.bio or up(r['EQUIPAMENTO']) == up(eqp))]
            if cands:
                alvo = (cands[0], i)
                break
    if not alvo:
        reg('E10 Inativação SIM→NÃO→SIM -- pré-condição ausente: nenhum N1 do interfaceamento visível no LJ com outro '
            'nível na mesma corrida', False, {'analito': an, 'lote': lote, 'corridas_na_janela': e0['n']})
        return fin0
    X, sl0 = alvo
    xid, run_x, xval = X['ID_REGISTRO'], iint(X['RUN']), X['RESULTADO']
    ctx.marcar('E10', xid)
    rec0 = collections.Counter(tuple(r.values()) for r in q.ler(TB_REC))
    s0, p0 = snap_eng(q), q.nome('engPainel').Value
    x0 = x_multiconjunto(e0)

    # F1: inativar com o padrao (SIM) e comentario
    r_inat, _ = ctx.inativar(xid, 'QA ETL E10: resultado fora do esperado, repetido (teste automatizado).')
    ctx.atualizar('E10 F1 inativar')
    fin1 = ler_final(ctx)
    f1, _ = por_id(fin1)
    e1, p1 = q.eng_saida(), q.nome('engPainel').Value
    d01 = diff_final(f0, f1)
    fora = {k: v for k, v in d01.items() if k != xid or not set(v) <= COLS_INATIVACAO}
    rec1 = collections.Counter(tuple(r.values()) for r in q.ler(TB_REC))
    reg('E10a diff(F0,F1): só a linha X muda e só nas colunas da inativação (STATUS, PARTICIPA, REGISTRAR, TIPO, '
        'INATIVACAO_REGISTRADA, DATA/USUARIO_INATIVACAO, TEM_JUSTIFICATIVA, GOVERNANCA, COMENTARIO); RUN de todos os '
        'IDs igual; recebimento idêntico',
        set(d01) == {xid} and not fora and rec0 == rec1 and f1[xid]['STATUS_ANALITICO'] == 'INATIVADO',
        {'x': xid, 'colunas_X': d01.get(xid), 'outras_linhas_ou_colunas': dict(list(fora.items())[:5]),
         'recebimento_identico': rec0 == rec1, 'final_X': resumo(f1.get(xid))})
    sl1 = e1['runs'].index(run_x) if run_x in e1['runs'] else None
    x1 = x_multiconjunto(e1)
    novos_x = x1 - x0
    reg('E10b F1: X no slot do RUN e no nível certo (3 slots × níveis varridos); nenhum outro X novo',
        sl1 is not None and xval in e1['x'][sl1][0:3] and novos_x == collections.Counter({(run_x, 1, xval): 1})
        and not (x0 - x1),
        {'RUN': run_x, 'slot': sl1, 'X_N1': e1['x'][sl1][0:3] if sl1 is not None else None,
         'X_novos': [list(k) for k in novos_x], 'X_sumidos': [list(k) for k in (x0 - x1)]})
    reg('E10c engPainel F1: n(N1) caiu exatamente 1 em relação a F0; os outros níveis não mudam',
        painel_n(p0, 1) - painel_n(p1, 1) == 1 and all(painel_n(p0, k) == painel_n(p1, k) for k in range(2, ctx.nlv + 1)),
        {'F0': [l[:5] for l in p0], 'F1': [l[:5] for l in p1]})

    # F2: desmarcar
    ctx.cel(TB_INAT, COL_LJ, r_inat).Value = False
    ctx.atualizar('E10 F2 desmarcar')
    fin2 = ler_final(ctx)
    f2, _ = por_id(fin2)
    e2, p2 = q.eng_saida(), q.nome('engPainel').Value
    d12 = diff_final(f1, f2)
    sl2 = e2['runs'].index(run_x) if run_x in e2['runs'] else None
    reg('E10d F2 (desmarcado): corrida na janela (slot não nulo), nenhum X em slot nenhum além dos de F0; só '
        'REGISTRAR/TIPO da linha X mudam; engPainel = F1' + ('' if sl2 is not None else
                                                              ' -- pré-condição ausente: a corrida saiu da janela'),
        sl2 is not None and x_multiconjunto(e2) == x0 and set(d12) == {xid} and set(d12[xid]) <= COLS_CAIXA
        and f2[xid]['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and iguais(p2, p1),
        {'slot': sl2, 'X_N1': e2['x'][sl2][0:3] if sl2 is not None else None, 'colunas_X': d12.get(xid),
         'outras': [k for k in d12 if k != xid][:5], 'painel_F1': [l[:5] for l in p1], 'painel_F2': [l[:5] for l in p2]})

    # F3: marcar de novo
    ctx.cel(TB_INAT, COL_LJ, r_inat).Value = True
    ctx.atualizar('E10 F3 marcar')
    fin3 = ler_final(ctx)
    f3, _ = por_id(fin3)
    e3, p3 = q.eng_saida(), q.nome('engPainel').Value
    s3 = snap_eng(q)
    d23 = diff_final(f2, f3)
    reg('E10e F3 (marcado de novo): o X volta ao mesmo slot; só REGISTRAR/TIPO da linha X mudam; engPainel = F1',
        x_multiconjunto(e3) == x1 and set(d23) == {xid} and set(d23[xid]) <= COLS_CAIXA and iguais(p3, p1)
        and f3[xid]['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO',
        {'colunas_X': d23.get(xid), 'X_novos_vs_F1': [list(k) for k in (x_multiconjunto(e3) - x1)],
         'painel_F3': [l[:5] for l in p3]})

    # F4: refresh extra
    ctx.atualizar('E10 F4 refresh extra')
    fin4 = ler_final(ctx)
    f4, _ = por_id(fin4)
    s4 = snap_eng(q)
    d34 = diff_final(f3, f4)
    reg('E10f F4 = F3: final (todas as colunas, todas as linhas), Eng_Saida e engPainel idênticos; engPainel F1=F2=F3=F4',
        not d34 and iguais(s4['grade'], s3['grade']) and iguais(s4['painel'], s3['painel'])
        and iguais(s4['estat'], s3['estat']) and iguais(s4['painel'], p1),
        {'linhas_diferentes': dict(list(d34.items())[:5]),
         'eng_diferencas': difs_grade(s3['grade'], s4['grade'], 'grade'), 'painel_dif': difs_grade(p1, s4['painel'], 'painel')})
    ctx.e10 = {'x': xid}
    return reconciliar_global(ctx, 'E10', fin4)


# =============================================================================
#  E11 -- contaminacao zero com registro extremo
# =============================================================================
def _ultimo_ok(rows_dia, gap):
    lim = dt.time(23, 50)
    if not rows_dia:
        return True
    ult = max(r['DATA_HORA'] for r in rows_dia)
    alvo = dt.datetime.combine(ult.date(), lim)
    return (alvo - ult).total_seconds() / 60.0 > gap + 1


def _num_lote(ctx, lote, analito, nivel=1):
    m = ctx.q.run('mLotes.MediaDoLote', lote, analito, nivel)
    s = ctx.q.run('mLotes.DPDoLote', lote, analito, nivel)
    return (float(m) if eh_num(m) else None), (float(s) if eh_num(s) else None)


def _escolher_dia(ctx, fin, analito, lote, eqp, cmin, cmax):
    runs = corridas_motor(fin, analito, lote, eqp)
    dias = sorted({int(math.floor(s)) for s, _ in runs})
    hoje_s = dia(meia_noite(ctx.hoje))
    melhor = None
    for d in dias:
        cum = sum(1 for s, _ in runs if math.floor(s) <= d)
        if cum < cmin or cum > cmax or d >= hoje_s:
            continue
        if _ultimo_ok(linhas_do_dia(fin, analito, lote, de_serial(d).date()), ctx.gap):
            melhor = (d, cum)
    return melhor, runs


def snap_e11(ctx, analito, lote):
    q = ctx.q
    fin = ler_final(ctx)
    return {'fin': fin, 'eng': q.eng_saida(), 'painel': q.nome('engPainel').Value,
            'estat': estat_do_analito(q, analito), 'eventos': snap_eventos(q, analito),
            'cont': contadores(fin, analito, lote), 'run': {r['ID_REGISTRO']: r['RUN'] for r in fin},
            'cortadas': iint(q.ws('Eng_Saida').Range('M1').Value) or 0}


def estat_sem_aq(linhas):
    return [tuple(v for j, v in enumerate(l) if j not in (IDX_AQ, IDX_ORDEM)) for l in linhas]


def conferir_contaminacao(rot, s0, s2, run_m, extremo, esperado_cont, prova=''):
    e0, e2 = s0['eng'], s2['eng']
    eng_ok = (e2['runs'][:len(e0['runs'])] == e0['runs'] and iguais(e2['val'][:e0['n']], e0['val'])
              and iguais(e2['flags'][:e0['n']], e0['flags']) and iguais(e2['x'][:e0['n']], e0['x']))
    if extremo is not None:
        sl = e2['runs'].index(run_m) if run_m in e2['runs'] else None
        eng_ok = eng_ok and sl is not None and e2['n'] == e0['n'] + 1 and extremo in e2['x'][sl][0:3]
    else:
        eng_ok = eng_ok and e2['n'] == e0['n'] and iguais(e2['x'], e0['x'])
    dc = dict(s2['cont'] - s0['cont'])
    neg = dict(s0['cont'] - s2['cont'])
    ok_p = iguais(s2['painel'], s0['painel'])
    ok_e = iguais(estat_sem_aq(s2['estat']), estat_sem_aq(s0['estat']))
    ok_w = s2['eventos'] == s0['eventos']
    ok_c = dc == esperado_cont and not neg
    ok_r = all(s2['run'].get(i) == v for i, v in s0['run'].items())
    reg(f'{rot}{prova}', ok_p and ok_e and ok_w and ok_c and eng_ok and ok_r,
        {'painel_igual': ok_p, 'painel_dif': difs_grade(s0['painel'], s2['painel'], 'engPainel'),
         'estatistica_igual(sem AQ)': ok_e, 'estatistica_dif': difs_grade(estat_sem_aq(s0['estat']), estat_sem_aq(s2['estat']), 'Estat'),
         'eventos_westgard_iguais': ok_w, 'eventos': [len(s0['eventos']), len(s2['eventos'])],
         'contadores_delta': {str(k): v for k, v in dc.items()}, 'contadores_negativos': {str(k): v for k, v in neg.items()},
         'eng_saida_ok': eng_ok, 'corridas': [s0['eng']['n'], s2['eng']['n']], 'RUN_inalterados': ok_r,
         'AQ_alertas': [[l[IDX_AQ] for l in s0['estat']], [l[IDX_AQ] for l in s2['estat']]]})


def e11(ctx, fin0):
    q = ctx.q
    lote, eqp = ctx.tela['lote'], ctx.tela['equip']
    hoje_s = dia(meia_noite(ctx.hoje))
    cands = analitos_candidatos(ctx, fin0, lote, eqp)
    # ---- variante A: janela NAO saturada, lote com media/DP (para o 1_3s)
    A = None
    for a in cands:
        m, s = _num_lote(ctx, lote, a)
        if m is None or not s:
            continue
        esc, runs = _escolher_dia(ctx, fin0, a, lote, eqp, 20, 170)
        if esc:
            A = (a, esc[0], esc[1], runs[0][0], m, s)
            break
    if not A:
        reg('E11 Contaminação zero (variante A) -- pré-condição ausente: nenhum analito do lote com média/DP e um dia '
            'final de período com ≤170 corridas e 23:50 livre', False, {'lote': lote, 'candidatos': cands[:10]})
        return fin0
    aA, dA, cumA, iniA, mA, sA = A
    ctx.usados.add(aA.upper())
    # ---- variante B: janela saturada (>180 corridas ate o fim do periodo)
    B, sint = None, []
    for b in [c for c in cands if c != aA]:
        esc, runs = _escolher_dia(ctx, fin0, b, lote, eqp, 181, 10 ** 9)
        if esc:
            B = (b, esc[0], esc[1], runs[0][0])
            break
    if not B:
        # gera com manuais sinteticos nos dias ANTERIORES ao 1o resultado do lote (outro _GK, nada colide)
        melhor = None
        for b in [c for c in cands if c != aA]:
            esc, runs = _escolher_dia(ctx, fin0, b, lote, eqp, 10, 10 ** 9)
            if esc and (melhor is None or esc[1] > melhor[2]):
                melhor = (b, esc[0], esc[1], runs[0][0])
        if melhor:
            b, dB, cumB, ini_b = melhor
            precisa = 181 + 2 - cumB
            passo = max(ctx.gap, ctx.tol) + 15
            por_dia = max(1, int((23 * 60 - 30) // passo))
            mB, sB = _num_lote(ctx, lote, b)
            vals = [r['RESULTADO'] for r in fin0 if up(r['ANALITO']) == b.upper() and iint(r['NIVEL']) == 1
                    and r['PARTICIPA_ESTATISTICA'] == SIM and eh_num(r['RESULTADO'])]
            base = mB if mB is not None else (sorted(vals)[len(vals) // 2] if vals else 1.0)
            sd = sB or (abs(base) * 0.02) or 1.0
            d0 = de_serial(math.floor(ini_b)).date()
            eqb = ctx.equip_do_analito(fin0, b)
            for i in range(precisa):
                dd = d0 - dt.timedelta(days=1 + i // por_dia)
                quando = dt.datetime.combine(dd, dt.time(0, 30)) + dt.timedelta(minutes=passo * (i % por_dia))
                sint.append((quando, round(base + (0.3 if i % 2 else -0.3) * sd, 4)))
            B = (b, dB, cumB + precisa, dia(min(x[0] for x in sint)))
            escrever_sinteticos(ctx, b, lote, eqb, sint)
    if B:
        ctx.usados.add(B[0].upper())
    # ---- S0 (A)
    ctx.configurar(analito=aA, de=de_serial(iniA), ate=de_serial(dA))
    s0 = snap_e11(ctx, aA, lote)
    pre_a = s0['eng']['n'] <= 178 and s0['cortadas'] == 0 and painel_n(s0['painel']) > 1
    ext_a = round(mA + 20 * sA, 6)
    quandoA = dt.datetime.combine(de_serial(dA).date(), dt.time(23, 50))
    eqA = ctx.equip_do_analito(fin0, aA)
    rA, idA = ctx.manual(quandoA, aA, lote, 1, ext_a, eqA)
    ctx.marcar('E11A', idA)
    ctx.atualizar('E11 S1 (manual extremo A' + (f' + {len(sint)} sintéticos B' if sint else '') + ')')
    s1 = snap_e11(ctx, aA, lote)
    f1, _ = por_id(s1['fin'])
    mA_row = f1.get(idA)
    run_m = iint((mA_row or {}).get('RUN'))
    sl = s1['eng']['runs'].index(run_m) if run_m in s1['eng']['runs'] else None
    dia_runs0 = [iint(r['RUN']) for r in linhas_do_dia(s0['fin'], aA, lote, quandoA.date()) if r['RUN'] not in (None, '')]
    renum = [i for i, v in s0['run'].items() if s1['run'].get(i) != v]
    reg('E11a S1 ≠ S0 (sensibilidade): manual extremo (média+20·DP às 23:50) ATIVO em corrida própria no fim, sem '
        'renumerar; n(N1)+1, média deslocada e 1_3s no seu slot'
        + ('' if pre_a else ' -- pré-condição ausente: janela saturada/cortada ou n<2'),
        pre_a and mA_row is not None and mA_row['STATUS_ANALITICO'] == 'ATIVO' and sl is not None
        and run_m == (max(dia_runs0) + 1 if dia_runs0 else None) and not renum
        and painel_n(s1['painel']) == painel_n(s0['painel']) + 1
        and not perto(s1['painel'][0][2], s0['painel'][0][2]) and s1['eng']['flags'][sl][0] == 1,
        {'analito': aA, 'lote': lote, 'periodo': [de_serial(iniA), de_serial(dA)], 'corridas_ate_o_fim': cumA,
         'manual': resumo(mA_row), 'extremo': ext_a, 'alvo_lote': [mA, sA], 'slot': sl,
         'flags_N1': s1['eng']['flags'][sl][:7] if sl is not None else None, 'renumerados': renum[:5],
         'painel_S0': [l[:5] for l in s0['painel']], 'painel_S1': [l[:5] for l in s1['painel']],
         'corridas_janela_S0': s0['eng']['n'], 'cortadas_S0': s0['cortadas']})
    # ---- S0 (B): depois dos sinteticos, antes do extremo B
    s0b = None
    if B:
        bB, dB, cumB, iniB = B
        ctx.configurar(analito=bB, de=de_serial(iniB), ate=de_serial(dB))
        s0b = snap_e11(ctx, bB, lote)
    # ---- S2 (A): inativar com comentario, LJ = SIM
    ctx.configurar(analito=aA, de=de_serial(iniA), ate=de_serial(dA), motor=False)
    rIA, _ = ctx.inativar(idA, 'QA ETL E11: valor extremo de teste (média + 20 DP).')
    ctx.atualizar('E11 S2 inativar A')
    s2 = snap_e11(ctx, aA, lote)
    conferir_contaminacao('E11b S2 = S0 (|Δ| < 1e-9): engPainel completo, linha da Estatística (u(Rw) incluída, '
                          'AQ/PEXCL permitida), Eventos_Westgard, RUNs; só X_VERMELHO/INATIVADO +1 e o X extremo no '
                          'RUN do manual', s0, s2, run_m, ext_a, {('INATIVADO', 'X_VERMELHO'): 1})
    # ---- S3 (A): LJ = NAO; no mesmo refresh entra o extremo B (outro analito, outra serie)
    ctx.cel(TB_INAT, COL_LJ, rIA).Value = False
    idB = rB_ = None
    if B:
        bB, dB, cumB, iniB = B
        mB, sB = _num_lote(ctx, lote, bB)
        if mB is None or not sB:
            vs = [r['RESULTADO'] for r in s2['fin'] if up(r['ANALITO']) == bB.upper() and iint(r['NIVEL']) == 1
                  and r['PARTICIPA_ESTATISTICA'] == SIM and eh_num(r['RESULTADO'])]
            mB = sum(vs) / len(vs) if vs else 1.0
            sB = (max(vs) - min(vs)) / 4 if len(vs) > 1 and max(vs) > min(vs) else abs(mB) * 0.05 or 1.0
        ext_b = round(mB + 20 * sB, 6)
        rB_, idB = ctx.manual(dt.datetime.combine(de_serial(dB).date(), dt.time(23, 50)), bB, lote, 1, ext_b,
                              ctx.equip_do_analito(s2['fin'], bB))
        ctx.marcar('E11B', idB)
    ctx.atualizar('E11 S3 LJ=NÃO A' + (' + extremo B' if B else ''))
    s3 = snap_e11(ctx, aA, lote)
    conferir_contaminacao('E11c S3 = S0 com LJ = NÃO: indicadores e eventos iguais, X ausente (Eng_Saida = S0)',
                          s0, s3, run_m, None, {('INATIVADO', 'NAO_PLOTAR'): 1})
    if not B:
        reg('E11d Variante B (janela saturada, >180 corridas) -- pré-condição ausente: nem o dado real nem a geração '
            'sintética deram um período saturado', False, {'lote': lote})
        return reconciliar_global(ctx, 'E11', s3['fin'])
    # ---- variante B
    ctx.configurar(analito=bB, de=de_serial(iniB), ate=de_serial(dB))
    s1b = snap_e11(ctx, bB, lote)
    fb, _ = por_id(s1b['fin'])
    pre_b = s0b['eng']['n'] == 180 and s0b['eng']['fil'][0] == 1 and (fb.get(idB) or {}).get('STATUS_ANALITICO') == 'ATIVO'
    reg('E11d Variante B, S1 ≠ S0 (sensibilidade): extremo ATIVO numa janela saturada (180 corridas, a mais antiga '
        'dentro do filtro)' + ('' if pre_b else ' -- pré-condição ausente: janela não saturada ou manual não ATIVO'),
        pre_b and not iguais(s1b['painel'], s0b['painel']),
        {'analito': bB, 'periodo': [de_serial(iniB), de_serial(dB)], 'corridas_ate_o_fim': cumB,
         'sinteticos': len(sint), 'janela_S0': s0b['eng']['n'], 'fil_mais_antiga': s0b['eng']['fil'][0] if s0b['eng']['n'] else None,
         'manual': resumo(fb.get(idB)), 'painel_S0': [l[:5] for l in s0b['painel']], 'painel_S1': [l[:5] for l in s1b['painel']]})
    ctx.inativar(idB, 'QA ETL E11 variante B: valor extremo de teste.')
    ctx.atualizar('E11 S2 inativar B')
    s2b = snap_e11(ctx, bB, lote)
    ok_p = iguais(s2b['painel'], s0b['painel'])
    ok_e = iguais(estat_sem_aq(s2b['estat']), estat_sem_aq(s0b['estat']))
    ok_w = s2b['eventos'] == s0b['eventos']
    reg('E11e Variante B, S2 = S0 com a janela NK=180 saturada: o X não consome a janela -- n/média/DP/CV/Sigma do '
        'Painel, Estatística e eventos iguais (hoje falha: prova D10)' +
        ('' if pre_b else ' -- pré-condição ausente (ver E11d)'),
        pre_b and ok_p and ok_e and ok_w,
        {'painel_igual': ok_p, 'painel_dif': difs_grade(s0b['painel'], s2b['painel'], 'engPainel'),
         'estatistica_igual': ok_e, 'eventos_iguais': ok_w, 'janela': [s0b['eng']['n'], s2b['eng']['n']],
         'cortadas': [s0b['cortadas'], s2b['cortadas']], 'primeira_corrida_janela': [s0b['eng']['runs'][:1], s2b['eng']['runs'][:1]]})
    voltar_tela(ctx)
    ctx.e11 = {'A': aA, 'B': bB}
    return reconciliar_global(ctx, 'E11', s2b['fin'])


def escrever_sinteticos(ctx, analito, lote, equip, pontos):
    """Bloco de manuais sinteticos colado de uma vez em DATA..RESULTADO (um evento: o mesmo caminho de uma
    colagem do usuario -- o evento numera os MAN_ e carimba linha a linha). Garante as linhas da tabela antes."""
    q = ctx.q
    lo = q.lo(TB_MAN)
    cab = [c.Name for c in lo.ListColumns]
    seq = ['DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO']
    i0 = cab.index('DATA')
    if cab[i0:i0 + len(seq)] != seq:
        for quando, v in pontos:                     # layout inesperado: linha a linha
            ctx.manual(quando, analito, lote, 1, v, equip)
        return
    ids = [r[0] for r in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
    ult = max([i for i, v in enumerate(ids, 1) if v not in (None, '')] or [0])
    falta = ult + len(pontos) + 15 - lo.ListRows.Count
    ws = lo.Parent
    if falta > 0:
        ws.Unprotect(SENHA)
        rg = lo.Range
        lo.Resize(ws.Range(rg.Cells(1, 1), ws.Cells(rg.Row + rg.Rows.Count - 1 + falta, rg.Column + rg.Columns.Count - 1)))
    linhas = tuple((meia_noite(quando), (quando - meia_noite(quando)).total_seconds() / 86400.0, equip, None, str(lote), 1,
                    analito, float(v)) for quando, v in pontos)
    corpo = lo.DataBodyRange
    alvo = ws.Range(corpo.Cells(ult + 1, i0 + 1), corpo.Cells(ult + len(pontos), i0 + len(seq)))
    alvo.Value = linhas
    ids2 = [r[0] for r in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value][ult:ult + len(pontos)]
    for i in ids2:
        ctx.marcar('E11B-sintetico', i)
    print(f'   {len(pontos)} manuais sintéticos; IDs {ids2[:1]}..{ids2[-1:]}', flush=True)


# =============================================================================
#  E12 (+E14 no mesmo refresh) -- insercao manual e RUN; casos-limite de agrupamento
# =============================================================================
def _minuto_acima(d):
    d2 = d.replace(second=0, microsecond=0)
    return d2 if d2 >= d else d2 + dt.timedelta(minutes=1)


def planejar_e12(ctx, fin, pred_est):
    gap, tol = ctx.gap, ctx.tol
    # M_bloco (dentro do GAP e fora da TOL) so existe com GAP > TOL + 2; na configuracao do laboratorio
    # (GAP 10 <= TOL 30) todo manual dentro do GAP de uma corrida e conflito -- o cenario sai do plano
    com_bloco = gap > tol + 2
    grupos = collections.defaultdict(list)
    for r in fin:
        if r['STATUS_ANALITICO'] in CORRIDA_REAL and isinstance(r['DATA_HORA'], dt.datetime) \
                and isinstance(r['DATA'], dt.datetime):
            grupos[(cstr(r['EQUIPAMENTO']), cstr(r['MATRIZ']), cstr(r['LOTE']), cstr(r['ANALITO']), r['DATA'].date())].append(r)
    t = ctx.tela
    hoje = ctx.hoje

    def prioridade(k):
        rows = grupos[k]
        tela = k[3].upper() == t['analito'].upper() and k[2] == t['lote'] and (not ctx.bio or k[0].upper() == up(t['equip']))
        rec = any(pred_est(r, True) for r in rows)
        return (0 if tela else 1, 0 if rec else 1, -k[4].toordinal())
    G = lambda m: dt.timedelta(minutes=m)          # noqa: E731
    for k in sorted(grupos, key=prioridade):
        if k[4] >= hoje or k[3].upper() in ctx.usados - {t['analito'].upper()}:
            continue
        bl = blocos_do_grupo(grupos[k], gap)
        if len(bl) < 2:
            continue
        b1, b2 = bl[0], bl[1]
        niveis_b1 = collections.Counter(iint(r['NIVEL']) for r in b1)
        if any(v > 1 for v in niveis_b1.values()):
            continue
        n1_b1 = [r for r in b1 if iint(r['NIVEL']) == 1 and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and eh_num(r['RESULTADO'])]
        n1_b2 = [r for r in b2 if iint(r['NIVEL']) == 1 and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and eh_num(r['RESULTADO'])
                 and r['STATUS_ANALITICO'] in ('ATIVO', 'INATIVADO')]
        if not n1_b1 or not n1_b2:
            continue
        d0 = meia_noite(k[4])
        s1, e1 = b1[0]['DATA_HORA'], b1[-1]['DATA_HORA']
        s2 = b2[0]['DATA_HORA']
        ult = max(r['DATA_HORA'] for r in grupos[k])
        m_ini = (s1 - G(max(gap, tol) + 5)).replace(second=0, microsecond=0)   # fora do GAP E da TOL
        if m_ini < d0 + G(1):
            continue
        if com_bloco:
            m_bloco = _minuto_acima(max(e1 + G(1), n1_b1[0]['DATA_HORA'] + G(tol + 2)))
            if (m_bloco - e1).total_seconds() / 60 > gap - 1:
                continue
            m_meio = _minuto_acima(m_bloco + G(gap + 3))
        else:
            m_bloco = None
            m_meio = _minuto_acima(e1 + G(max(gap, tol) + 3))
        if (s2 - m_meio).total_seconds() / 60 <= max(gap, tol) + 2:
            continue
        m_fim = d0 + G(23 * 60 + 59)
        if (m_fim - ult).total_seconds() / 60 <= max(gap, tol) + 1:
            continue
        return {'gk': k, 'b1': b1, 'b2': b2, 'n1_b1': n1_b1[0], 'n1_b2': n1_b2[0], 'M_ini': m_ini, 'M_bloco': m_bloco,
                'M_meio': m_meio, 'M_fim': m_fim, 'M_inst': n1_b2[0]['DATA_HORA'], 'grupo': grupos[k]}, None
    return None, {'motivo': 'nenhum _GK/dia com ≥2 corridas e espaço para M_ini, M_bloco, M_meio e M_fim',
                  'grupos_avaliados': len(grupos), 'GAP': gap, 'TOL': tol}


def conferir_oraculo(ctx, fin, etapa):
    orc = oraculo_run(fin, ctx.gap)
    div = []
    for r in fin:
        esp = orc.get(r['ID_REGISTRO'], (None, None))
        got = (iint(r['CORRIDA_NO_DIA']), iint(r['RUN']))
        if esp != got:
            div.append({'id': r['ID_REGISTRO'], 'status': r['STATUS_ANALITICO'], 'final': got, 'oraculo': esp,
                        'DATA_HORA': r['DATA_HORA']})
    reg(f'E12 [{etapa}] Oráculo Python independente do RUN (_GK, ordenação, corte por GAP, _POS por nível, corrida '
        f'do dia) = CORRIDA_NO_DIA/RUN da final em 100% dos IDs', not div and len(orc) > 0,
        {'ids': len(fin), 'ids_com_corrida': len(orc), 'divergentes': len(div), 'exemplos': div[:8], 'GAP': ctx.gap})
    return not div


def e12_e14(ctx, fin0):
    q = ctx.q
    pred_est, _ = filtro_estatistica(ctx)
    conferir_oraculo(ctx, fin0, 'antes do E12')
    plano, motivo = planejar_e12(ctx, fin0, pred_est)
    run0 = {r['ID_REGISTRO']: iint(r['RUN']) for r in fin0}
    cor0 = {r['ID_REGISTRO']: iint(r['CORRIDA_NO_DIA']) for r in fin0}
    ids_m = {}
    if plano is None:
        reg('E12 Inserção manual e RUN -- pré-condição ausente', False, motivo)
    else:
        eqg, mzg, ltg, ang, dg = plano['gk']
        v1 = float(plano['n1_b1']['RESULTADO'])
        for nome, quando, fator in (('M_ini', plano['M_ini'], 1.003), ('M_bloco', plano['M_bloco'], 1.006),
                                    ('M_meio', plano['M_meio'], 1.004), ('M_fim', plano['M_fim'], 1.005)):
            if quando is None:                        # M_bloco impossivel pela configuracao (GAP <= TOL)
                continue
            _, ids_m[nome] = ctx.manual(quando, ang, ltg, 1, round(v1 * fator, 6), eqg, matriz=mzg)
        _, ids_m['M_inst'] = ctx.manual(plano['M_inst'], ang, ltg, 1, plano['n1_b2']['RESULTADO'], eqg, matriz=mzg)
        ctx.marcar('E12', *ids_m.values())
        ctx.marcar('E12-interface', *(r['ID_REGISTRO'] for r in plano['grupo']))
    # ---------------- E14 no mesmo refresh
    t = ctx.tela
    an, L1 = t['analito'], t['lote']
    eqA = ctx.equip_do_analito(fin0, an)
    equips = sorted({cstr(r['EQUIPAMENTO']) for r in fin0 if r['EQUIPAMENTO']})
    distintos = {a: len({r['EQUIPAMENTO'] for r in fin0 if up(r['ANALITO']) == a.upper()}) for a in ctx.cadastro if a}
    eqB = next((e for e in equips if e.upper() != up(eqA)), 'QA-EQUIP-B')
    lotes_an = collections.Counter(cstr(r['LOTE']).strip() for r in fin0 if up(r['ANALITO']) == an.upper())
    L2 = next((l for l, _ in lotes_an.most_common() if l and l != L1), 'QAETL2')
    evitar = {plano['gk'][4]} if plano else set()
    d14 = dia_livre(ctx, fin0, an, evitar)
    m14 = {}
    if d14:
        base = [r['RESULTADO'] for r in fin0 if up(r['ANALITO']) == an.upper() and iint(r['NIVEL']) == 1
                and r['PARTICIPA_ESTATISTICA'] == SIM and eh_num(r['RESULTADO'])]
        vb = round(sorted(base)[len(base) // 2], 4) if base else 1.0
        dez = dt.datetime.combine(d14, dt.time(10, 0))
        lote_u = f'QAU{d14.strftime("%y%m%d")}'
        for nome, lote, eqx, v in (('L1', L1, eqA, vb), ('L2', L2, eqA, round(vb * 1.07, 6)),
                                   ('B', L1, eqB, round(vb * 1.13, 6)), ('U', lote_u, eqA, round(vb * 0.91, 6))):
            _, idx = ctx.manual(dez, an, lote, 1, v, eqx)
            m14[nome] = {'id': idx, 'lote': lote, 'equip': eqx, 'valor': v}
        ctx.marcar('E14', *(x['id'] for x in m14.values()))
    ctx.atualizar('E12+E14')
    fin1 = ler_final(ctx)
    f1, dup1 = por_id(fin1)
    qa = q.ler(TB_QA)
    conferir_oraculo(ctx, fin1, 'E12 após inserção')
    if plano:
        gk = plano['gk']
        no_grupo = {r['ID_REGISTRO'] for r in plano['grupo']}
        c_b1 = {cor0[r['ID_REGISTRO']] for r in plano['b1']}
        apos_b1 = 3 if 'M_bloco' in ids_m else 2         # M_ini + (M_bloco) + M_meio antes das corridas seguintes
        esperado = {i: cor0[i] + (1 if cor0[i] in c_b1 else apos_b1) for i in no_grupo if cor0[i] is not None}
        div_desl = {i: (cor0[i], iint(f1.get(i, {}).get('CORRIDA_NO_DIA')), e) for i, e in esperado.items()
                    if iint(f1.get(i, {}).get('CORRIDA_NO_DIA')) != e}
        outros = [i for i in run0 if i not in no_grupo and iint(f1.get(i, {}).get('RUN')) != run0[i]]
        novos = set(f1) - set(run0)
        mm = {k: f1.get(v) for k, v in ids_m.items()}
        ok_ativos = all(mm[k] is not None and mm[k]['STATUS_ANALITICO'] == 'ATIVO' for k in ('M_ini', 'M_meio', 'M_fim'))
        reg('E12a M_ini, M_meio e M_fim ATIVO; corridas posteriores do mesmo _GK/dia deslocadas (+1 após M_ini; +3 '
            'depois de M_bloco e M_meio), ordem relativa preservada, nenhum ID alterado, nenhum RUN de outro _GK alterado',
            ok_ativos and not div_desl and not outros and set(run0) <= set(f1) and not dup1,
            {'_GK': [str(x) for x in gk], 'manuais': {k: resumo(v) for k, v in mm.items()},
             'deslocamento_divergente': dict(list(div_desl.items())[:8]), 'outros_GK_alterados': outros[:8],
             'ids_novos': len(novos), 'ids_duplicados': dup1[:5]})
        rb = mm.get('M_bloco')
        run_b1 = max(iint(f1[r['ID_REGISTRO']]['RUN']) for r in plano['b1'])
        if 'M_bloco' not in ids_m:
            RES.append({'teste': 'E12b M_bloco impossível nesta configuração (GAP <= TOL: todo manual dentro do GAP de '
                                 'uma corrida é conflito pela TOL) -- coberto pela regra de CONFLITO (E12c)',
                        'resultado': 'INFO', 'evidencia': {'GAP': ctx.gap, 'TOL': ctx.tol}})
        else:
          reg('E12b M_bloco (dentro do GAP da corrida 1, fora da TOL do N1 dela): posição 2 do bloco, RUN próprio logo '
            'após a corrida 1, ATIVO e A06',
            rb is not None and rb['STATUS_ANALITICO'] == 'ATIVO' and iint(rb['RUN']) == run_b1 + 1
            and iint(rb['CORRIDA_NO_DIA']) == 3 and bool(achados(qa, 'A06', ids_m['M_bloco'])),
            {'M_bloco': resumo(rb), 'RUN_corrida1': run_b1, 'A06': len(achados(qa, 'A06', ids_m['M_bloco']))})
        ri = mm['M_inst']
        reg('E12c M_inst (mesmo DATA_HORA de um N1 do interfaceamento): CONFLITO_MANUAL, ID_RELACIONADO = interface, '
            'A01, fora do RUN',
            ri is not None and ri['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and ri['ID_RELACIONADO'] == plano['n1_b2']['ID_REGISTRO']
            and bool(achados(qa, 'A01', ids_m['M_inst'])) and iint(ri['RUN']) is None,
            {'M_inst': resumo(ri), 'interface': plano['n1_b2']['ID_REGISTRO']})
        ctx.e12 = ids_m
    # ---------------- E14: conferencias
    e14(ctx, fin0, fin1, m14, d14, distintos, eqA, eqB, L1, L2)
    # ---------------- E12: 2 refreshes extras
    base_runs = {r['ID_REGISTRO']: (iint(r['RUN']), iint(r['CORRIDA_NO_DIA'])) for r in fin1}
    difs = []
    for k in range(2):
        ctx.atualizar(f'E12 extra {k + 1}')
        fk = ler_final(ctx)
        atual = {r['ID_REGISTRO']: (iint(r['RUN']), iint(r['CORRIDA_NO_DIA'])) for r in fk}
        difs.append([i for i in set(base_runs) | set(atual) if base_runs.get(i) != atual.get(i)][:8])
    reg('E12d Após 2 refreshes extras todos os RUN (e CORRIDA_NO_DIA) ficam iguais', all(not d for d in difs) and len(fk) == len(fin1),
        {'diferentes_por_refresh': difs, 'linhas': [len(fin1), len(fk)]})
    return reconciliar_global(ctx, 'E12+E14', fk)


def e14(ctx, fin0, fin1, m14, d14, distintos, eqA, eqB, L1, L2):
    q = ctx.q
    f1, _ = por_id(fin1)
    if not d14 or len(m14) < 4:
        reg('E14 Casos-limite de agrupamento -- pré-condição ausente: dia sem interface para o analito em tela', False,
            {'dia': d14})
        return
    run01 = int(d14.strftime('%y%m%d')) * 100 + 1
    rows = {k: f1.get(v['id']) for k, v in m14.items()}
    reg('E14a L1, L2, equipamento B e lote único: cada um em _GK próprio, ATIVO, RUN yyMMdd01 (nada se junta)',
        all(r is not None and r['STATUS_ANALITICO'] == 'ATIVO' and iint(r['RUN']) == run01 for r in rows.values()),
        {'RUN_esperado': run01, 'linhas': {k: resumo(v) for k, v in rows.items()},
         'equipamentos_distintos_por_analito_antes': {a: n for a, n in distintos.items() if n} if not ctx.bio else None,
         'equip_A': eqA, 'equip_B': eqB, 'L1': L1, 'L2': L2})
    an = ctx.tela['analito']
    vA, v2, vB, vU = (m14[k]['valor'] for k in ('L1', 'L2', 'B', 'U'))

    def valores_na_eng(e):
        out = []
        for i in range(e['n']):
            out += [v for v in e['val'][i] if isinstance(v, float)] + [v for v in e['x'][i] if isinstance(v, float)]
        return out
    # lote L1 (equipamento da tela)
    ctx.configurar(analito=an, de=meia_noite(d14), ate=meia_noite(d14), lote=L1)
    e = q.eng_saida()
    sl = e['runs'].index(run01) if run01 in e['runs'] else None
    vals_l1 = valores_na_eng(e)
    reg('E14b Índice ANALITO|LOTE: no LJ do lote L1 o RUN yyMMdd01 mostra o valor de L1; o de L2 não aparece',
        sl is not None and e['val'][sl][0] == (vA if ctx.bio else e['val'][sl][0]) and v2 not in vals_l1
        and (vA in vals_l1),
        {'slot': sl, 'N1_no_slot': e['val'][sl][0] if sl is not None else None, 'valores': {'L1': vA, 'L2': v2, 'B': vB}})
    if ctx.bio:
        reg('E14c Bioquímica: o equipamento B fica FORA do LJ do equipamento A (mesmo lote, mesmo RUN, mesmo instante)',
            sl is not None and vB not in vals_l1 and e['val'][sl][0] == vA,
            {'equip_painel': e['equip'], 'N1_no_slot': e['val'][sl][0] if sl is not None else None, 'vB': vB})
    else:
        sim_run = [r for r in fin1 if up(r['ANALITO']) == an.upper() and cstr(r['LOTE']).strip() == L1
                   and iint(r['NIVEL']) == 1 and iint(r['RUN']) == run01 and r['PARTICIPA_ESTATISTICA'] == SIM]
        no_slot = [e['val'][sl][0]] if sl is not None else []
        pn = painel_n(q.nome('engPainel').Value)
        padrao = ctx.equip_filtro()
        do_padrao = [r for r in sim_run if up(r['EQUIPAMENTO']) == padrao]
        a10 = [x for x in q.ler('tblQA_Integracao') if x.get('CODIGO') == 'A10']
        reg('E14c Hematologia: o mesmo RUN de 2 equipamentos não sobrescreve nem soma em dobro -- o Painel é a série '
            'do equipamento padrão do CFG (o ponto do RUN é o dele, n(N1) = 1) e o outro equipamento gera ALERTA A10 '
            '(hoje falha: prova D12)',
            len(sim_run) == 2 and len(do_padrao) == 1 and no_slot == [do_padrao[0]['RESULTADO']] and pn == 1 and bool(a10),
            {'resultados_SIM_no_RUN': [(r['ID_REGISTRO'], r['EQUIPAMENTO'], r['RESULTADO']) for r in sim_run],
             'pontos_no_slot': no_slot, 'n_N1_painel': pn, 'equipamento_padrao': padrao,
             'A10': [x.get('DETALHE') for x in a10][:2],
             'equipamentos_distintos_por_analito_antes': {a: n for a, n in distintos.items() if n}})
        if not ctx.bio:
            pass
    # lote L2
    ctx.configurar(lote=L2)
    e = q.eng_saida()
    sl2 = e['runs'].index(run01) if run01 in e['runs'] else None
    vals_l2 = valores_na_eng(e)
    reg('E14d No LJ do lote L2 o mesmo RUN yyMMdd01 mostra só o valor de L2 (L1 e B ausentes)',
        sl2 is not None and e['val'][sl2][0] == v2,   # o ponto do RUN e o de L2 (valor igual pode existir em outra corrida)
        {'slot': sl2, 'N1_no_slot': e['val'][sl2][0] if sl2 is not None else None})
    if ctx.bio:
        ctx.configurar(lote=L1, equip=eqB)
        e = q.eng_saida()
        slb = e['runs'].index(run01) if run01 in e['runs'] else None
        reg('E14e Bioquímica: o equipamento B é uma série própria (no LJ de B, RUN yyMMdd01 = valor de B)',
            slb is not None and e['val'][slb][0] == vB,      # o ponto do RUN e o de B (valor igual pode existir em outra corrida)
            {'equip_painel': e['equip'], 'N1_no_slot': e['val'][slb][0] if slb is not None else None, 'vA': vA, 'vB': vB})
        ctx.configurar(equip=ctx.tela['equip'], motor=False)
    # lote unico
    ctx.configurar(lote=m14['U']['lote'])
    p = q.nome('engPainel').Value
    erros = [v for l in p for v in l if isinstance(v, int) and v < -2146820000]
    reg('E14f Lote novo com um único resultado: n=1, média = o valor, DP vazio, sem erro no motor',
        painel_n(p) == 1 and perto(p[0][2], vU) and p[0][3] in (None, '') and not erros,
        {'engPainel_N1': p[0][:5], 'valor': vU, 'erros': erros})
    # primeira e ultima corrida do periodo do Painel (tela)
    voltar_tela(ctx)
    e = q.eng_saida()
    t = ctx.tela
    runs = corridas_motor(fin1, t['analito'], t['lote'], t['equip'] if ctx.bio else '')
    de_s, ate_s = dia(t['de']), dia(t['ate'])
    per = [(s, r) for s, r in runs if de_s <= math.floor(s) <= ate_s]
    corte = iint(q.ws('Eng_Saida').Range('M1').Value) or 0
    if per:
        prim, ult = per[0][1], per[-1][1]
        fora_jan = [r for _, r in per if r not in e['runs']]
        ok = ult in e['runs'] and e['fil'][e['runs'].index(ult)] == 1 and \
            ((prim in e['runs'] and e['fil'][e['runs'].index(prim)] == 1) or corte == len(fora_jan))
        reg('E14g Primeira e última corrida do período do Painel aparecem no Eng_Saida (ou o corte da janela é '
            'informado com a contagem exata em engCortadas)', ok,
            {'primeira': prim, 'ultima': ult, 'corridas_no_periodo': len(per), 'janela': e['n'], 'engCortadas': corte,
             'fora_da_janela': len(fora_jan)})
    else:
        reg('E14g Primeira e última corrida do período -- pré-condição ausente: período sem corridas', False, {'tela': t})


# =============================================================================
#  E13 -- manual e inativacao do mesmo ID; manual corrigindo interface inativado
# =============================================================================
def e13(ctx, fin0):
    q = ctx.q
    pred_est, _ = filtro_estatistica(ctx)
    f0, _ = por_id(fin0)
    tol = ctx.tol
    # (a) MAN_k ATIVO dentro do recorte da Estatistica (de preferencia os do E12)
    pref_ids = [ctx.e12.get(k) for k in ('M_ini', 'M_meio', 'M_fim') if ctx.e12.get(k)]
    outros = [r['ID_REGISTRO'] for r in fin0 if r['ORIGEM_RESULTADO'] == 'MANUAL' and r['STATUS_ANALITICO'] == 'ATIVO'
              and digitos(str(r['ID_REGISTRO'])[4:]) and str(r['ID_REGISTRO']).startswith('MAN_')]
    ma = next((f0[i] for i in pref_ids + outros if i in f0 and f0[i]['STATUS_ANALITICO'] == 'ATIVO' and pred_est(f0[i])), None)
    usados_an = set(ctx.usados)
    # (b) interface ATIVO, +5 min no mesmo dia, sem outro interface da mesma chave a <= 10 min
    chave = lambda r: (r['EQUIPAMENTO'], r['MATRIZ'], r['LOTE'], iint(r['NIVEL']), r['ANALITO'])   # noqa: E731
    por_chave_dia = collections.defaultdict(list)
    for r in fin0:
        if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and isinstance(r['DATA_HORA'], dt.datetime) and eh_num(r['RESULTADO']):
            por_chave_dia[chave(r) + (r['DATA_HORA'].date(),)].append(r)
    xb = None
    if tol >= 6:
        for r in sorted(fin0, key=lambda r: r['DATA_HORA'] if isinstance(r['DATA_HORA'], dt.datetime) else dt.datetime.min,
                        reverse=True):
            if r['ORIGEM_RESULTADO'] != 'INTERFACEAMENTO' or r['STATUS_ANALITICO'] != 'ATIVO' or \
                    up(r['ANALITO']) in usados_an or not isinstance(r['DATA_HORA'], dt.datetime):
                continue
            if r['DATA_HORA'].time() > dt.time(23, 50) or r['DATA_HORA'].date() >= ctx.hoje:
                continue
            m = r['DATA_HORA'] + dt.timedelta(minutes=5)
            viz = [o for o in por_chave_dia[chave(r) + (r['DATA_HORA'].date(),)]
                   if o is not r and abs((o['DATA_HORA'] - m).total_seconds()) / 60 <= 10]
            if not viz:
                xb = r
                break
    # (c) interface ATIVO no fim do dia; manual 00:02 do dia seguinte a menos de TOL atravessando a meia-noite
    yc = None
    max_part = max((r['DATA_HORA'] for r in fin0 if r['PARTICIPA_ESTATISTICA'] == SIM
                    and isinstance(r['DATA_HORA'], dt.datetime)), default=None)
    if tol >= 5 and max_part:
        lim = dt.time(23, 59, 59)
        for r in sorted(fin0, key=lambda r: r['DATA_HORA'] if isinstance(r['DATA_HORA'], dt.datetime) else dt.datetime.min,
                        reverse=True):
            if r['ORIGEM_RESULTADO'] != 'INTERFACEAMENTO' or r['STATUS_ANALITICO'] != 'ATIVO' or \
                    not isinstance(r['DATA_HORA'], dt.datetime) or up(r['ANALITO']) in usados_an or \
                    (xb is not None and up(r['ANALITO']) == up(xb['ANALITO'])):
                continue
            meia = meia_noite(r['DATA_HORA'].date() + dt.timedelta(days=1))
            m = meia + dt.timedelta(minutes=2)
            if (m - r['DATA_HORA']).total_seconds() / 60 > tol - 1 or r['DATA_HORA'].time() > lim:
                continue
            if m.date() >= ctx.hoje or m > max_part:
                continue
            viz = [o for o in por_chave_dia[chave(r) + (m.date(),)] if (o['DATA_HORA'] - meia).total_seconds() / 60 <= tol + 3]
            if not viz:
                yc = r
                break
    pre = {'a': ma is not None, 'b': xb is not None, 'c': yc is not None}
    est_a0 = estat_do_analito(q, ma['ANALITO']) if ma else None
    na0 = next((iint(l[2]) for l in est_a0 if iint(l[1]) == iint(ma['NIVEL'])), None) if ma else None
    r_ia = None
    if ma:
        k = int(str(ma['ID_REGISTRO'])[4:])
        r_ia, _ = ctx.inativar(f'man_{k}', 'QA ETL E13a: manual inativado pela chave digitada em minúsculas.')
        cel_a = ctx.valor(TB_INAT, 'ID_REGISTRO', r_ia)
        ctx.marcar('E13a', ma['ID_REGISTRO'])
    id_b = id_c = None
    if xb:
        ctx.inativar(xb['ID_REGISTRO'], 'QA ETL E13b: interface inativado; o manual corrige o valor.')
        _, id_b = ctx.manual(xb['DATA_HORA'] + dt.timedelta(minutes=5), xb['ANALITO'], xb['LOTE'], iint(xb['NIVEL']),
                             round(float(xb['RESULTADO']) * 1.02, 6), xb['EQUIPAMENTO'], matriz=xb['MATRIZ'])
        ctx.marcar('E13b', xb['ID_REGISTRO'], id_b)
    if yc:
        quando = meia_noite(yc['DATA_HORA'].date() + dt.timedelta(days=1)) + dt.timedelta(minutes=2)
        _, id_c = ctx.manual(quando, yc['ANALITO'], yc['LOTE'], iint(yc['NIVEL']), yc['RESULTADO'], yc['EQUIPAMENTO'],
                             matriz=yc['MATRIZ'])
        ctx.marcar('E13c', yc['ID_REGISTRO'], id_c)
    ctx.atualizar('E13 R1')
    fin1 = ler_final(ctx)
    f1, _ = por_id(fin1)
    qa = q.ler(TB_QA)
    if ma:
        ra = f1.get(ma['ID_REGISTRO'])
        est_a1 = estat_do_analito(q, ma['ANALITO'])
        na1 = next((iint(l[2]) for l in est_a1 if iint(l[1]) == iint(ma['NIVEL'])), None)
        reg('E13a Manual inativado pela chave "man_k": INATIVADO/X_VERMELHO, mesmo RUN, n−1 na Estatística',
            ra is not None and ra['STATUS_ANALITICO'] == 'INATIVADO' and ra['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
            and iint(ra['RUN']) == iint(ma['RUN']) and na0 is not None and na1 == na0 - 1 and cel_a == ma['ID_REGISTRO'],
            {'manual': resumo(ra), 'RUN_antes': ma['RUN'], 'n_estatistica': [na0, na1], 'celula': cel_a})
    else:
        reg('E13a Manual e inativação do mesmo ID -- pré-condição ausente: nenhum MAN_ ATIVO no recorte da Estatística',
            False, {'candidatos_E12': pref_ids})
    if xb:
        rb = f1.get(id_b)
        a09 = achados(qa, 'A09', id_b)
        reg('E13b Manual corrigindo interface INATIVADO (±5 min): CONFLITO_MANUAL com ID_RELACIONADO = interface, '
            'fora da estatística, e ALERTA A09 (decisão conservadora; hoje falha: prova D09b)',
            rb is not None and rb['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and rb['ID_RELACIONADO'] == xb['ID_REGISTRO']
            and rb['PARTICIPA_ESTATISTICA'] == NAO and bool(a09),
            {'manual': resumo(rb), 'interface': resumo(f1.get(xb['ID_REGISTRO'])), 'A09': [x['DETALHE'] for x in a09],
             'A01': len(achados(qa, 'A01', id_b))})
    else:
        reg('E13b Manual corrigindo interface inativado -- pré-condição ausente: TOL < 6 min ou nenhum interface '
            'isolado', False, {'TOL': tol})
    if yc:
        rc = f1.get(id_c)
        reg('E13c Conflito manual × interface atravessa a meia-noite (manual 00:02, interface no fim do dia anterior, '
            'dentro de ±TOL por DATA_HORA): CONFLITO_MANUAL (hoje falha: prova D09c)',
            rc is not None and rc['STATUS_ANALITICO'] == 'CONFLITO_MANUAL' and rc['ID_RELACIONADO'] == yc['ID_REGISTRO'],
            {'manual': resumo(rc), 'interface': resumo(f1.get(yc['ID_REGISTRO'])), 'TOL': tol,
             'achados_do_manual': sorted({x['CODIGO'] for x in qa if x['ID_REGISTRO'] == id_c})})
    else:
        # sem interface a menos de TOL da meia-noite no dado real: o caso e provado com DB_SEAC sintetico em
        # qa_etl_recebimento.py (E13c: interface S16 23:58 x manual 00:02) -- aqui so registra a remissao (INFO)
        RES.append({'teste': 'E13c coberto em qa_etl_recebimento.py (DB_SEAC sintético: interface 23:58 x manual 00:02)',
                    'resultado': 'INFO', 'evidencia': {'TOL': tol}})
    if ma and r_ia:
        ctx.apagar_linha(TB_INAT, r_ia)
        ctx.atualizar('E13 R2 apagar inativação')
        fin2 = ler_final(ctx)
        f2, _ = por_id(fin2)
        ra2 = f2.get(ma['ID_REGISTRO'])
        est_a2 = estat_do_analito(q, ma['ANALITO'])
        reg('E13a2 Apagar a inativação reabilita: ATIVO, mesmo ID, mesmo RUN, sem duplicar; Estatística do analito '
            'idêntica ao estado ativo (A..AU)',
            ra2 is not None and ra2['STATUS_ANALITICO'] == 'ATIVO' and iint(ra2['RUN']) == iint(ma['RUN'])
            and sum(1 for r in fin2 if r['ID_REGISTRO'] == ma['ID_REGISTRO']) == 1 and iguais(est_a2, est_a0),
            {'manual': resumo(ra2), 'estatistica_dif': difs_grade(est_a0, est_a2, 'Estat')})
        return reconciliar_global(ctx, 'E13', fin2)
    return reconciliar_global(ctx, 'E13', fin1)


# =============================================================================
#  E15 -- idempotencia do cenario completo
# =============================================================================
def snapshot_completo(ctx):
    q = ctx.q
    fin = ler_final(ctx)
    d, dup = por_id(fin)
    _, qa_tudo, qa_ids = snap_qa(q)
    return {'fin': d, 'dup': dup, 'n': len(fin), 'eng': snap_eng(q), 'estat': snap_estat(q),
            'qa': qa_tudo, 'qa_ids': qa_ids, 'lista': fin}


def e15(ctx):
    s0 = snapshot_completo(ctx)
    tempos, falhas = [], []
    for k in range(3):
        _, t = ctx.atualizar(f'E15 #{k + 1}')
        tempos.append(round(t, 1))
        s = snapshot_completo(ctx)
        d = diff_final(s0['fin'], s['fin'])
        ok = {'final': not d, 'n': s['n'] == s0['n'], 'sem_duplicados': not s['dup'],
              'eng_saida': iguais(s['eng']['grade'], s0['eng']['grade']) and iguais(s['eng']['cab'], s0['eng']['cab']),
              'engPainel': iguais(s['eng']['painel'], s0['eng']['painel']), 'engEstat': iguais(s['eng']['estat'], s0['eng']['estat']),
              'estatistica': iguais(s['estat'], s0['estat']), 'qa_multiconjunto': s['qa'] == s0['qa']}
        if not all(ok.values()):
            falhas.append({'refresh': k + 1, 'ok': ok, 'linhas_diferentes': dict(list(d.items())[:5]),
                           'eng_dif': difs_grade(s0['eng']['grade'], s['eng']['grade'], 'grade'),
                           'estat_dif': difs_grade(s0['estat'], s['estat'], 'Estat'),
                           'qa_novos': [list(x) for x in (s['qa_ids'] - s0['qa_ids'])][:5],
                           'qa_sumidos': [list(x) for x in (s0['qa_ids'] - s['qa_ids'])][:5]})
    reg('E15 Idempotência do cenário acumulado (E08–E14): 3 refreshes com saída IDÊNTICA -- final linha inteira, '
        'Eng_Saida, engPainel, engEstat, aba Estatística e tblQA_Integracao (multiconjunto); |Final| constante, nenhum '
        'ID novo, sumido ou repetido', not falhas and not s0['dup'],
        {'linhas_final': s0['n'], 'tempos_s': tempos, 'falhas': falhas, 'duplicados_iniciais': s0['dup'][:5]})
    return reconciliar_global(ctx, 'E15', s['lista'])


# =============================================================================
#  E16 -- matriz de rastreabilidade por ID
# =============================================================================
def e16(ctx, fin):
    q = ctx.q
    f, _ = por_id(fin)
    rec_ids = {r['ID_REGISTRO'] for r in q.ler(TB_REC)}
    man = q.ler(TB_MAN)
    inat = q.ler(TB_INAT)
    qa = q.ler(TB_QA)
    grupos = derivar_manuais(man, ctx.pref)
    venc, _ = vencedor_manual(grupos, set(f))
    lin_inat = collections.defaultdict(list)
    for n, r in enumerate(inat, 1):
        i = norm_id(r.get('ID_REGISTRO'), ctx.pref)
        if i:
            lin_inat[i].append((n, r.get(COL_LJ)))
    qa_por = collections.defaultdict(list)
    for x in qa:
        if x['ID_REGISTRO']:
            qa_por[x['ID_REGISTRO']].append(x['CODIGO'])
    pred, desc = filtro_estatistica(ctx)
    e = q.eng_saida()
    t = ctx.tela
    pos = {r: i for i, r in enumerate(e['runs'])}
    amostra = random.Random(57).sample(sorted(f), min(40, len(f)))
    alvo = sorted(set(ctx.ids) | set(amostra), key=str)
    linhas, prob = [], []
    for i in alvo:
        r = f.get(i)
        lm = None
        if i and '~L' in str(i):
            lm = iint(str(i).split('~L')[-1])
        elif i in grupos:
            lm = venc.get(i)
        if r is None:
            ach = qa_por.get(i, [])
            linhas.append({'ID': i, 'ORIGEM': None, 'NO_RECEBIMENTO': i in rec_ids, 'LINHA_MANUAL': lm,
                           'INATIVADO': None, 'LINHA_INAT': [n for n, _ in lin_inat.get(i, [])], 'STATUS': 'FORA_DA_FINAL',
                           'ACHADOS_QA': ach, 'CENARIOS': sorted(ctx.ids.get(i, []))})
            if not ach and lin_inat.get(i):
                prob.append((i, 'inativação de ID inexistente sem E03'))
            continue
        no_lj = 'FORA_DA_TELA'
        if up(r['ANALITO']) == t['analito'].upper() and cstr(r['LOTE']).strip() == t['lote'] and \
                (not ctx.bio or up(r['EQUIPAMENTO']) == up(t['equip'])):
            run = iint(r['RUN'])
            if run not in pos:
                no_lj = 'FORA_DA_JANELA'
            else:
                nv = (iint(r['NIVEL']) or 1) - 1
                sl = pos[run]
                if r['PARTICIPA_ESTATISTICA'] == SIM:
                    no_lj = 'PONTO' if e['val'][sl][nv] == r['RESULTADO'] else 'AUSENTE'
                elif r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO':
                    no_lj = 'X' if r['RESULTADO'] in e['x'][sl][nv * 3:nv * 3 + 3] else 'AUSENTE'
                else:
                    no_lj = 'NAO_PLOTA'
        destino = {'PARTICIPA': r['PARTICIPA_ESTATISTICA'] == SIM and r['TIPO_PLOTAGEM_LJ'] == 'NORMAL'
                   and r['STATUS_ANALITICO'] == 'ATIVO',
                   'X': r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO' and r['STATUS_ANALITICO'] == 'INATIVADO'
                   and r['REGISTRAR_RESULTADO_NO_LJ'] == SIM and r['PARTICIPA_ESTATISTICA'] == NAO,
                   'NAO_PLOTA': r['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR' and r['PARTICIPA_ESTATISTICA'] == NAO}
        motivo = r['MOTIVO_EXCLUSAO_AUTOMATICA'] or (('inativação registrada' if r['INATIVACAO_REGISTRADA'] == SIM else None))
        na_est = pred(r) is not None
        linha = {'ID': i, 'ORIGEM': r['ORIGEM_RESULTADO'], 'NO_RECEBIMENTO': i in rec_ids, 'LINHA_MANUAL': lm,
                 'INATIVADO': r['STATUS_ANALITICO'] == 'INATIVADO', 'LINHA_INAT': [n for n, _ in lin_inat.get(i, [])],
                 'REGISTRAR_LJ': r['REGISTRAR_RESULTADO_NO_LJ'], 'STATUS': r['STATUS_ANALITICO'],
                 'PARTICIPA': r['PARTICIPA_ESTATISTICA'], 'TIPO_PLOTAGEM': r['TIPO_PLOTAGEM_LJ'], 'RUN': iint(r['RUN']),
                 'NA_ESTATISTICA': SIM if na_est else NAO, 'NO_RECORTE_ESTATISTICA': pred(r, True) is not None,
                 'NO_LJ': no_lj, 'ACHADOS_QA': qa_por.get(i, []), 'DESTINO': [k for k, v in destino.items() if v],
                 'MOTIVO': motivo, 'CENARIOS': sorted(ctx.ids.get(i, []))}
        linhas.append(linha)
        if sum(destino.values()) != 1:
            prob.append((i, f'destino não único: {linha["DESTINO"]}'))
        if na_est and r['PARTICIPA_ESTATISTICA'] != SIM:
            prob.append((i, 'NA_ESTATISTICA=SIM com PARTICIPA=NÃO'))
        if r['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO' and not destino['X']:
            prob.append((i, 'X sem INATIVADO∧LJ'))
        if r['STATUS_ANALITICO'] != 'ATIVO' and not (r['STATUS_ANALITICO'] and (motivo or qa_por.get(i))):
            prob.append((i, 'não ATIVO sem motivo nem achado'))
        if (i in rec_ids) == (r['ORIGEM_RESULTADO'] == 'MANUAL'):
            prob.append((i, 'origem incoerente com o recebimento'))
        if no_lj == 'AUSENTE':
            prob.append((i, 'deveria estar no LJ em tela e não está'))
    os.makedirs(RESULTADOS, exist_ok=True)
    arq = os.path.join(RESULTADOS, f'etl_rastreabilidade_{ctx.produto}.json')
    with open(arq, 'w', encoding='utf-8') as fh:
        json.dump({'produto': ctx.produto, 'gerado_em': dt.datetime.now().isoformat(timespec='seconds'),
                   'nota': 'Só os IDs dos cenários do qa_etl e uma amostra fixa (semente 57) -- nunca a base inteira.',
                   'recorte_estatistica': desc, 'tela': {k: str(v) for k, v in t.items()},
                   'linhas': linhas, 'problemas': prob}, fh, ensure_ascii=False, indent=1, default=str)
    reg('E16 Matriz de rastreabilidade por ID: cada ID tem destino único e coerente (participa / X / não plota, com '
        'motivo ou achado); nenhum NA_ESTATISTICA=SIM com PARTICIPA=NÃO; nenhum X sem INATIVADO∧LJ'
        + ('' if len(ctx.ids) else ' -- pré-condição ausente: nenhum ID de cenário'),
        not prob and len(ctx.ids) > 0,
        {'arquivo': arq, 'ids': len(linhas), 'ids_de_cenario': len(ctx.ids), 'amostra': len(amostra),
         'problemas': prob[:15]})


# =============================================================================
def executar(produto, copia, parte):
    q = QA(produto, copia)
    ctx = None
    try:
        ctx = Ctx(q, produto)
        print(f'   PREFIXO={ctx.pref} GAP={ctx.gap} TOL={ctx.tol} parte={parte}', flush=True)
        ctx.atualizar('estado-base')
        fin = reconciliar_global(ctx, 'estado-base')
        if escolher_tela(ctx, fin) is None:
            reg('E10/E12/E14 tela do Painel -- pré-condição ausente: nenhum analito com resultado no lote em análise',
                False, {'lote': ctx.lote_painel()})
            return
        etapas = []
        if parte in ('todas', '1'):
            etapas += [('E08', e08), ('E09', e09), ('E10', e10), ('E11', e11)]
        if parte in ('todas', '2'):
            etapas += [('E12+E14', e12_e14), ('E13', e13)]
        for nome, fn in etapas:
            print(f'== {nome}', flush=True)
            try:
                r = fn(ctx, fin)
                fin = r if isinstance(r, list) else ler_final(ctx)
            except Exception as ex:                  # noqa: BLE001 -- a etapa falha, a suite segue e registra
                reg(f'{nome} exceção na execução da etapa', False, {'erro': repr(ex)[:400],
                                                                     'traceback': traceback.format_exc()[-1500:]})
                try:
                    voltar_tela(ctx)
                    fin = ler_final(ctx)
                except Exception:                    # noqa: BLE001
                    pass
        print('== E15', flush=True)
        try:
            fin = e15(ctx)
        except Exception as ex:                      # noqa: BLE001
            reg('E15 exceção na execução da etapa', False, {'erro': repr(ex)[:400], 'traceback': traceback.format_exc()[-1500:]})
        print('== E16', flush=True)
        try:
            e16(ctx, fin)
        except Exception as ex:                      # noqa: BLE001
            reg('E16 exceção na execução da etapa', False, {'erro': repr(ex)[:400], 'traceback': traceback.format_exc()[-1500:]})
    finally:
        if ctx is not None:
            RES.append({'teste': 'TEMPOS', 'resultado': 'INFO', 'evidencia': ctx.tempos})
            print('   tempos de atualização:', json.dumps(ctx.tempos, ensure_ascii=False), flush=True)
        q.fechar(salvar=False)


def main():
    args = [a for a in sys.argv[1:]]
    parte = 'todas'
    if '--parte' in args:
        i = args.index('--parte')
        parte = args[i + 1]
        del args[i:i + 2]
    if len(args) < 2 or parte not in ('todas', '1', '2'):
        print(__doc__)
        sys.exit(2)
    produto, entrada = args[0], os.path.abspath(args[1])
    saida = args[2] if len(args) > 2 else None
    h0 = sha256(entrada)
    raiz, ext = os.path.splitext(entrada)
    copia = f'{raiz}_qa_etl_tmp{ext}'
    shutil.copy2(entrada, copia)
    t0 = time.time()
    try:
        executar(produto, copia, parte)
    except Exception as ex:                          # noqa: BLE001 -- falha de abertura/infra: registra
        reg('E00 execução da suíte', False, {'erro': repr(ex)[:400], 'traceback': traceback.format_exc()[-1500:]})
    finally:
        for _ in range(5):
            try:
                if os.path.exists(copia):
                    os.remove(copia)
                break
            except OSError:
                time.sleep(2)
        h1 = sha256(entrada)
        reg('E00 Arquivo de entrada intacto: SHA-256 igual antes e depois (a suíte só abriu uma cópia descartável e '
            'nunca salvou)', h0 == h1,
            {'arquivo': entrada, 'sha256_antes': h0, 'sha256_depois': h1, 'copia_removida': not os.path.exists(copia),
             'duracao_total_min': round((time.time() - t0) / 60, 1)})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ETL (parte {parte}): {sum(1 for r in RES if r["resultado"] == "PASS")} PASS / '
          f'{len(falhas)} FAIL ===', flush=True)
    sys.exit(1 if falhas else 0)


if __name__ == '__main__':
    main()
