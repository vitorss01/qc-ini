# -*- coding: utf-8 -*-
"""instalar_adr083.py -- ADR-083: regras de Westgard do MESMO nivel na SEQUENCIA DE ATIVOS daquele nivel.

Uso:  python instalar_adr083.py <Bioquimica|Hematologia> <arquivo.xlsm> [--sem-salvar]

Decisao do usuario (22/10/2026; substitui a decisao 3 do ADR-081): o 2_2s entre corridas (oficial na Bioquimica) e os
complementares SAME_LEVEL comparam observacoes ATIVAS seguidas do nivel -- uma corrida so de outro nivel, ou um
resultado inativado, no meio NAO interrompe a regra. As regras ENTRE niveis (R_4s, 2_2s/2of3_2s dentro da corrida,
4_1s N2_R2, 8x N2_R4, 3_1s N3_R1, 6x N3_R2) continuam pareando pelo RUN. O motor recalcula TODO o historico, entao a
instalacao pode mudar o veredicto de corridas JA liberadas: este instalador gera o RELATORIO antes de trocar o VBA.

Idempotente (a 2a execucao registra 0 mudancas e o relatorio sai vazio). Passos:
  1. VARREDURA ANTES, com o VBA ANTIGO do arquivo: para cada lote com resultado elegivel (e, na Bioquimica, cada
     equipamento -- o escopo da aba Eventos_Westgard e UM lote x UM equipamento, todos os analitos, todo o historico do
     lote) roda mEstatistica.RegistrarEventosWestgard e le os eventos (OFICIAL e COMPLEMENTAR) da aba;
  2. instala todo o VBA atual das fontes (codigo_atual.py) -- EXIGE compilacao;
  3. VARREDURA DEPOIS, igual, com o VBA novo;
  4. RELATORIO (so quando o mEstatistica do arquivo ainda NAO tem o ADR-083 -- sem SeqDoNivel --, ou se a varredura
     achar diferenca): CSV ao lado do .xlsm,
     '<arquivo>_ADR083_mudancas_westgard.csv' (; e UTF-8), linhas EVENTO (evento OFICIAL que entrou/saiu),
     VEREDICTO (nivel x RUN que passou de OK para REJEITADO ou o contrario) e CONTAGEM (eventos por regra/detector/classe
     que mudaram, inclusive COMPLEMENTAR) -- so RUN, analito, lote, nivel, regra e antes/depois, sem valor de resultado;
     e UMA linha REGRA_WESTGARD_ALTERADA no Audit_Log com as contagens (nas mesmas condicoes);
     O CSV e gravado ANTES de salvar o .xlsm (CSV que nao grava = nada salvo; .xlsm que nao salva = CSV apagado); com
     --sem-salvar ele leva o sufixo _PREVIA e nunca ocupa o nome canonico;
  4b. aba Cfg_Westgard_Escopo: a Descricao dos detectores do mesmo nivel (ACROSS_RUN_SAME_LEVEL, SAME_LEVEL_Rn) passa a
     dizer "observacoes ativas seguidas" (o mesmo texto de scripts_fase3/materializar_cfg_westgard.py);
  5. o Painel volta ao lote/equipamento em que estava (analito e periodo nao sao tocados); o motor e recalculado
     (InvalidarCache + AtualizarEstatistica). A BI_Data NAO e tocada (pedido do usuario em 27/10/2026: "nao mexer em
     nada de bi/power bi por enquanto"); ela se atualiza quando o laboratorio rodar "Atualizar BI";
  5b. FOTO POR LOTE x EQUIPAMENTO (opcao completa do ADR-083, 27/10/2026): cria (idempotente) a aba tecnica
     Westgard_Fotos -- oculta (veryHidden) e protegida, como as outras abas tecnicas -- e grava nela, pelo proprio VBA
     (mEstatistica.GravarFotosWestgard), a foto dos eventos OFICIAIS de TODOS os lotes x equipamento: uma linha
     CONTROLE por lote x equipamento (impressao digital dos dados de que os eventos dependem + data da foto), uma
     EVENTO por evento OFICIAL (a chave estavel da trilha) e linhas CORRIDAS (corridas avaliadas, para separar
     RETROATIVO de rotina). Foto inteira (forcar) quando a aba e nova ou o mEstatistica do arquivo mudou; senao so os
     lotes cuja impressao digital mudou. NAO audita nada (a unica linha do instalador no Audit_Log continua sendo a
     REGRA_WESTGARD_ALTERADA). A partir dela o ATUALIZAR (mIntegracao.AtualizarDadosCore) audita a mudanca de
     veredicto em QUALQUER lote, nao so no do Painel;
  6. FUMACA: eventos antigos contidos nos novos (superconjunto, as 14 colunas) -- violacao e LOGADA e vai no CSV como
     evento removido, nao interrompe (o laboratorio decide); a aba final bate com a varredura do lote do Painel; a foto
     de cada lote x equipamento bate com a VARREDURA DEPOIS (os mesmos eventos OFICIAIS);
  7. devolve a protecao de cada aba e a da estrutura como estavam (Westgard_Fotos sempre protegida).
Qualquer falha de VBA/COM para ANTES de salvar.
"""
import collections
import csv
import datetime as dt
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import pqlib  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import codigo_atual  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

ABA_EV = 'Eventos_Westgard'
EV_NCOL = 14                                   # mEstatistica.EV_NCOL
# colunas (1-based) da aba Eventos_Westgard (ADR-045)
C_DATA, C_RUN, C_ANALITO, C_NIVEIS, C_REGRA, C_DETECTOR = 1, 2, 3, 4, 5, 6
C_ESCOPO, C_CLASSE, C_N, C_R, C_RUNINI, C_EVID, C_CLASSIF, C_ZMAX = 7, 8, 9, 10, 11, 12, 13, 14
SUFIXO_CSV = '_ADR083_mudancas_westgard.csv'
ABA_CFG_WG = 'Cfg_Westgard_Escopo'
C_CFG_DET, C_CFG_DESC = 4, 12                  # Detector e Descricao (materializar_cfg_westgard.CAB)
# = scripts_fase3/materializar_cfg_westgard.DESCRICAO (ADR-083): o --conferir de la continua batendo
DESC_CFG_WG = {
    'ACROSS_RUN_SAME_LEVEL': 'Observacoes ativas seguidas do mesmo nivel de controle (corrida so de outro nivel no '
                             'meio nao interrompe; ADR-083)',
    'SAME_LEVEL_R3': 'Complementar: 3 observacoes ativas seguidas do mesmo nivel (corrida so de outro nivel no '
                     'meio nao interrompe; ADR-083)',
    'SAME_LEVEL_R4': 'Complementar: 4 observacoes ativas seguidas do mesmo nivel (corrida so de outro nivel no '
                     'meio nao interrompe; ADR-083)',
    'SAME_LEVEL_R6': 'Complementar: 6 observacoes ativas seguidas do mesmo nivel (corrida so de outro nivel no '
                     'meio nao interrompe; ADR-083)',
    'SAME_LEVEL_R8': 'Complementar: 8 observacoes ativas seguidas do mesmo nivel (corrida so de outro nivel no '
                     'meio nao interrompe; ADR-083)',
}
CAB_CSV = ['TIPO', 'EQUIPAMENTO', 'LOTE', 'ANALITO', 'NIVEL', 'RUN', 'REGRA', 'DETECTOR', 'RUN_INICIAL',
           'CLASSE', 'ANTES', 'DEPOIS']
MUDANCAS = []
# ADR-083 (opcao completa): aba tecnica da foto por lote x equipamento (= mEstatistica.WGF_*)
ABA_FOTOS = 'Westgard_Fotos'
CAB_FOTOS = ['TIPO', 'LOTE', 'EQUIPAMENTO', 'CHAVE', 'DESCRICAO', 'DATA_FOTO']
TITULO_FOTOS = ('WESTGARD_FOTOS (ADR-083, opcao completa): foto dos eventos OFICIAIS de Westgard por lote x equipamento '
                '-- o ATUALIZAR compara com ela e audita no Audit_Log a mudanca de veredicto em qualquer lote. Aba '
                'tecnica, gravada so pelo VBA: nao editar.')
LOTE_DA_ABA = {}                               # (equip, lote) -> lote que a aba Eventos_Westgard mostrou na varredura


def log(*a):
    print(*a, flush=True)


def txt(v):
    """Como o CStr do VBA mostra um codigo: 8353.0 -> '8353'."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return '' if v is None else str(v).strip()


def run_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def norm_cel(v):
    """Valor de celula comparavel entre duas leituras (datas e floats estaveis)."""
    if hasattr(v, 'year'):
        return v.replace(tzinfo=None).strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(v, float):
        return round(v, 9)
    return '' if v is None else str(v)


def ler_coluna(lo, nome):
    return [r[0] for r in lo.ListColumns(nome).DataBodyRange.Value]


# ---------------------------------------------------------------- codigo VBA do arquivo x fontes
def _texto_componente(wb, comp):
    cm = wb.VBProject.VBComponents(comp).CodeModule
    n = cm.CountOfLines
    return cm.Lines(1, n) if n else ''


def vba_diferente(wb, produto):
    """Componentes cujo codigo no arquivo difere da fonte (a mesma conferencia do codigo_atual)."""
    d = 'bio' if produto.startswith('Bio') else 'hema'
    nomes = {c.Name for c in wb.VBProject.VBComponents}
    dif = []
    pares = [(m, codigo_atual._fonte(d, m + '.bas')) for m in codigo_atual.MODULOS + codigo_atual.NOVOS]
    pares += [(m, codigo_atual._fonte(d, m + '.cls')) for m in codigo_atual.CLASSES]
    for aba, arq in codigo_atual.abas(wb):
        cn = [ws.CodeName for ws in wb.Worksheets if ws.Name == aba]
        if cn:
            pares.append((cn[0], codigo_atual._fonte(d, arq)))
    pares.append((wb.CodeName, codigo_atual._fonte(d, codigo_atual.PASTA)))
    for comp, fonte in pares:
        if comp not in nomes:
            dif.append(comp)
            continue
        if codigo_atual.normalizar(_texto_componente(wb, comp)).lower() != \
                codigo_atual.normalizar(ii.codigo(fonte)).lower():
            dif.append(comp)
    return dif


# ---------------------------------------------------------------- alvos da varredura
def cfg(wb, chave):
    lo = pqlib.tabela(wb, 'tblConfigIntegracao')
    if lo is None:
        return ''
    for r in lo.DataBodyRange.Value:
        if str(r[0] or '').strip().upper() == chave:
            return str(r[1] or '').strip()
    return ''


def alvos_da_varredura(wb, bio, eq_hema=''):
    """[(equipamento ou None, lote)]: lotes com resultado ELEGIVEL (os outros nao tem evento por construcao: a serie do
    RegistrarEventosWestgard so tem resultado elegivel), por equipamento na Bioquimica (cada equipamento e uma serie
    propria); na Hematologia o equipamento da serie e o do mEstatistica.EquipFiltro (seletor, ou EQUIPAMENTO_PADRAO;
    '' = todos) e o seletor nao e tocado."""
    lo = pqlib.tabela(wb, 'tblCQ_Final')
    lotes = ler_coluna(lo, 'LOTE')
    part = ler_coluna(lo, 'PARTICIPA_ESTATISTICA')
    equips = ler_coluna(lo, 'EQUIPAMENTO')
    an = ler_coluna(lo, 'ANALITO')
    alvos = set()
    for L, p, e, a in zip(lotes, part, equips, an):
        if p != 'SIM' or not txt(L) or not txt(a):
            continue
        eq = txt(e).upper()
        if bio:
            if eq:
                alvos.add((eq, txt(L)))
        elif not eq_hema or eq == eq_hema:
            alvos.add((None, txt(L)))
    cad = set()
    try:
        cad = {txt(r[0]) for r in wb.Names('regLoteCol').RefersToRange.Value if txt(r[0])}
    except Exception:                                          # noqa: BLE001 -- cadastro ausente: so informa
        pass
    sem_res = sorted(cad - {L for _, L in alvos})
    return sorted(alvos, key=lambda x: (x[0] or '', x[1])), sem_res


# ---------------------------------------------------------------- celulas do Painel que a varredura troca
class Seletor:
    """loteSel (e selEquipamento na Bioquimica): guarda e devolve EXATAMENTE como estava."""

    def __init__(self, wb, bio):
        self.wb, self.bio = wb, bio
        self.cel = {'loteSel': wb.Names('loteSel').RefersToRange}
        if bio:
            self.cel['selEquipamento'] = wb.Names('selEquipamento').RefersToRange
        self.orig = {n: (c.HasFormula, c.Formula, c.Value) for n, c in self.cel.items()}

    def por(self, equip, lote):
        if self.bio and equip is not None:
            self.cel['selEquipamento'].Value = equip
        self.cel['loteSel'].Value = "'" + lote          # texto: "010" nao vira 10 (mLotes.GravarLoteSel)

    def devolver(self):
        for n, (tem_f, f, v) in self.orig.items():
            c = self.cel[n]
            if tem_f:
                c.Formula = f
            elif v in (None, ''):
                c.ClearContents()
            elif isinstance(v, str):
                c.Value = "'" + v
            else:
                c.Value = v
        for n, (tem_f, f, v) in self.orig.items():
            c = self.cel[n]
            if (c.HasFormula, txt(c.Value)) != (tem_f, txt(v)):
                raise SystemExit(f'Painel: {n} nao voltou ao valor de antes ({v!r} -> {c.Value!r}) -- nada foi salvo')


# ---------------------------------------------------------------- varredura
def ler_eventos(wb):
    """(lote, equipamento, [linhas de 14 colunas]) da aba Eventos_Westgard como esta agora."""
    ws = wb.Worksheets(ABA_EV)
    n = int(float(ws.Range('J2').Value or 0))
    lote, eq = txt(ws.Range('L2').Value), txt(ws.Range('N2').Value).upper()
    if n <= 0:
        return lote, eq, []
    v = ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, EV_NCOL)).Value
    return lote, eq, [list(r) for r in v if txt(r[C_REGRA - 1])]


def mesmo_lote(a, b):
    """O lote da aba (canonico do cadastro, mLotes.LoteCanonico) e o da base: "010" e "10" sao o mesmo numero."""
    if a == b:
        return True
    try:
        return float(a) == float(b)
    except ValueError:
        return False


def varrer(ex, wb, alvos, sel, rotulo):
    """{(equip, lote): [linhas]} com o VBA que estiver no arquivo."""
    t0 = time.time()
    ex.run(f"'{wb.Name}'!mEstatistica.InvalidarCache", teto=600)
    out = {}
    for k, (eq, lote) in enumerate(alvos, 1):
        sel.por(eq, lote)
        ex.run(f"'{wb.Name}'!mEstatistica.RegistrarEventosWestgard", teto=900)
        lt, eqa, linhas = ler_eventos(wb)
        if not mesmo_lote(lt, lote) or (eq is not None and eqa != eq):
            raise SystemExit(f'varredura {rotulo}: a aba Eventos_Westgard ficou com lote {lt!r}/{eqa!r}, esperado '
                             f'{lote!r}/{eq!r} -- nada foi salvo')
        out[(eq, lote)] = linhas
        LOTE_DA_ABA[(eq, lote)] = lt
        if k % 10 == 0:
            log(f'   {rotulo}: {k}/{len(alvos)} ({time.time() - t0:.0f} s)')
    s = time.time() - t0
    n_of = sum(1 for ls in out.values() for r in ls if txt(r[C_CLASSE - 1]) == 'OFICIAL')
    n_tot = sum(len(ls) for ls in out.values())
    log(f'   {rotulo}: {len(alvos)} lote(s)x equipamento em {s:.1f} s ({s / max(1, len(alvos)):.2f} s cada); '
        f'{n_tot} evento(s), {n_of} OFICIAL')
    return out, s


def chave_evento(eq, lote, r):
    return (eq or '', lote, txt(r[C_ANALITO - 1]).upper(), txt(r[C_NIVEIS - 1]), txt(r[C_REGRA - 1]),
            txt(r[C_DETECTOR - 1]), run_int(r[C_RUN - 1]))


def marcas(eq, lote, linhas):
    """(equip, lote, ANALITO, nivel, RUN) -> {regras} marcadas por evento OFICIAL. E o veredicto do motor: um evento
    oficial marca, na corrida de fechamento, exatamente os niveis da sua coluna Niveis (AvaliarWestgard)."""
    m = collections.defaultdict(set)
    for r in linhas:
        if txt(r[C_CLASSE - 1]) != 'OFICIAL':
            continue
        for nv in txt(r[C_NIVEIS - 1]).split(','):
            if nv.strip():
                m[(eq or '', lote, txt(r[C_ANALITO - 1]).upper(), int(float(nv)), run_int(r[C_RUN - 1]))].add(
                    txt(r[C_REGRA - 1]))
    return m


def comparar(antes, depois):
    linhas_csv, resumo = [], collections.Counter()
    nome_an = {}
    alvos = sorted(set(antes) | set(depois), key=lambda x: (x[0] or '', x[1]))
    violacoes = []
    for (eq, lote) in alvos:
        A, D = antes.get((eq, lote), []), depois.get((eq, lote), [])
        for r in A + D:
            nome_an[txt(r[C_ANALITO - 1]).upper()] = txt(r[C_ANALITO - 1])
        # superconjunto: cada linha antiga (14 colunas) continua na aba nova
        ca = collections.Counter(tuple(norm_cel(x) for x in r) for r in A)
        cd = collections.Counter(tuple(norm_cel(x) for x in r) for r in D)
        falta = ca - cd
        if falta:
            violacoes.append(((eq, lote), sum(falta.values())))
        # eventos OFICIAIS que entraram / sairam
        oa = {chave_evento(eq, lote, r): r for r in A if txt(r[C_CLASSE - 1]) == 'OFICIAL'}
        od = {chave_evento(eq, lote, r): r for r in D if txt(r[C_CLASSE - 1]) == 'OFICIAL'}
        for k in sorted(set(oa) ^ set(od), key=lambda x: tuple(str(c) for c in x)):
            r = od.get(k) or oa.get(k)
            novo = k in od
            resumo['evento_novo' if novo else 'evento_removido'] += 1
            linhas_csv.append(['EVENTO', eq or '', lote, nome_an.get(k[2], k[2]), k[3], k[6], k[4], k[5],
                               run_int(r[C_RUNINI - 1]), 'OFICIAL', '' if novo else 'OFICIAL',
                               'OFICIAL' if novo else ''])
        # veredicto (nivel x RUN)
        ma, md = marcas(eq, lote, A), marcas(eq, lote, D)
        for k in sorted(set(ma) | set(md), key=lambda x: tuple(str(c) for c in x)):
            va = 'REJEITADO' if ma.get(k) else 'OK'
            vd = 'REJEITADO' if md.get(k) else 'OK'
            if va != vd:
                resumo['ok_para_rejeitado' if vd == 'REJEITADO' else 'rejeitado_para_ok'] += 1
                regras = '+'.join(sorted(md.get(k) or ma.get(k)))
                linhas_csv.append(['VEREDICTO', eq or '', lote, nome_an.get(k[2], k[2]), k[3], k[4], regras, '', '',
                                   'OFICIAL', va, vd])
        # contagem por (analito, regra, detector, classe), inclusive COMPLEMENTAR
        def contar(ls):
            c = collections.Counter()
            for r in ls:
                c[(txt(r[C_ANALITO - 1]).upper(), txt(r[C_REGRA - 1]), txt(r[C_DETECTOR - 1]),
                   txt(r[C_CLASSE - 1]))] += 1
            return c
        qa_, qd_ = contar(A), contar(D)
        for k in sorted(set(qa_) | set(qd_)):
            if qa_.get(k, 0) != qd_.get(k, 0):
                linhas_csv.append(['CONTAGEM', eq or '', lote, nome_an.get(k[0], k[0]), '', '', k[1], k[2], '',
                                   k[3], qa_.get(k, 0), qd_.get(k, 0)])
        for (an_, rg, dt_, cl), v in qa_.items():
            resumo[f'{cl.lower()}_antes'] += v
        for (an_, rg, dt_, cl), v in qd_.items():
            resumo[f'{cl.lower()}_depois'] += v
    return linhas_csv, resumo, violacoes


def nome_csv(caminho_xlsm, linhas, previa=False):
    """Onde gravar o relatorio; None = nao gravar (ja existe um e o novo esta vazio)."""
    base = os.path.splitext(caminho_xlsm)[0] + SUFIXO_CSV
    if previa:
        # --sem-salvar: o .xlsm nao muda; a previa nunca ocupa o nome canonico (o que o Audit_Log de verdade cita)
        return base[:-4] + '_PREVIA.csv'
    if os.path.exists(base) and linhas:
        # nunca destroi o relatorio da 1a instalacao: um relatorio novo com linhas vai com data/hora no nome
        return base[:-4] + '_' + dt.datetime.now().strftime('%Y%m%d_%H%M%S') + '.csv'
    if os.path.exists(base):
        return None
    return base


def gravar_csv(alvo, linhas):
    if not alvo:
        return None
    with open(alvo, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(CAB_CSV)
        for ln in linhas:
            w.writerow(['' if x is None else x for x in ln])
    return alvo


def descricoes_cfg_westgard(wb):
    """Descricao dos detectores do mesmo nivel na aba Cfg_Westgard_Escopo (ADR-083). Devolve quantas mudaram (None =
    aba ausente). Idempotente: a 2a execucao nao muda nada."""
    ws = ii.sheet(wb, ABA_CFG_WG)
    if ws is None:
        return None
    n, r = 0, 4
    while txt(ws.Cells(r, 1).Value):
        det = txt(ws.Cells(r, C_CFG_DET).Value)
        if det in DESC_CFG_WG and txt(ws.Cells(r, C_CFG_DESC).Value) != DESC_CFG_WG[det]:
            ws.Cells(r, C_CFG_DESC).Value = DESC_CFG_WG[det]
            n += 1
        r += 1
    return n


def montar_aba_fotos(wb):
    """ADR-083 (opcao completa): a aba tecnica Westgard_Fotos -- titulo, contadores (H2 = linhas CONTROLE, J2 = linhas
    da foto; o VBA apaga J2 durante a gravacao) e cabecalho; colunas A:E em texto. Oculta (veryHidden). Idempotente:
    existindo, so confere o cabecalho. Devolve (aba, criada)."""
    ws = ii.sheet(wb, ABA_FOTOS)
    criada = ws is None
    if criada:
        ultima = wb.Worksheets(wb.Worksheets.Count)
        ws = wb.Worksheets.Add(None, ultima)
        ws.Name = ABA_FOTOS
        ws.Range('A1').Value = TITULO_FOTOS
        ws.Range('A1').Font.Bold = True
        ws.Range('G2').Value = 'Lotes x equip.:'
        ws.Range('I2').Value = 'Linhas da foto:'
        ws.Columns('A:E').NumberFormat = '@'
        ws.Columns('F').NumberFormat = 'dd/mm/yyyy hh:mm'
        for j, w in enumerate((11, 12, 16, 60, 90, 16), 1):
            ws.Columns(j).ColumnWidth = w
    cab = [txt(ws.Cells(3, j).Value) for j in range(1, len(CAB_FOTOS) + 1)]
    if cab != CAB_FOTOS:
        for j, c in enumerate(CAB_FOTOS, 1):
            ws.Cells(3, j).Value = c
        ws.Range(ws.Cells(3, 1), ws.Cells(3, len(CAB_FOTOS))).Font.Bold = True
    ws.Cells.Locked = True
    if ws.Visible != 2:
        ws.Visible = 2                                 # veryHidden, como as abas tecnicas na entrega
    return ws, criada


def ler_foto(wb):
    """({(LOTE, EQUIPAMENTO): set((ANALITO, niveis, regra, RUN))}, linhas CONTROLE) da foto. (None, 0) se a foto estiver
    incompleta (J2 vazio)."""
    ws = wb.Worksheets(ABA_FOTOS)
    n = txt(ws.Range('J2').Value)
    if not n:
        return None, 0
    n = int(float(n))
    out, n_ctrl = {}, 0
    if n <= 0:
        return out, 0
    v = ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, len(CAB_FOTOS))).Value
    for r in v:
        tipo, lote, eq, chave = txt(r[0]), txt(r[1]), txt(r[2]).upper(), txt(r[3])
        if tipo == 'CONTROLE':
            n_ctrl += 1
            out.setdefault((lote, eq), set())
        elif tipo == 'EVENTO':
            an, _lt, niv, regra, run_ = chave.rsplit('|', 4)
            out.setdefault((lote, eq), set()).add((an, niv, regra, txt(run_)))
    return out, n_ctrl


def conferir_foto(foto_ev, depois, bio, eq_hema):
    """A foto de cada lote x equipamento = os eventos OFICIAIS da VARREDURA DEPOIS (mesma base, mesmo VBA). Devolve
    (divergencias, alvos conferidos); divergencias vazia = coerente."""
    dif, n = [], 0
    for (eq, lote), linhas in depois.items():
        lt_aba = LOTE_DA_ABA.get((eq, lote), lote)
        if lt_aba != lote:
            # a aba mostrou o lote canonico do cadastro ("010" x "10"): a varredura nao e comparavel ao lote da base
            log(f'   foto x varredura: {eq or ""}/{lote} nao conferido (a aba mostrou o lote {lt_aba!r})')
            continue
        esp = {(txt(r[C_ANALITO - 1]).upper(), txt(r[C_NIVEIS - 1]), txt(r[C_REGRA - 1]), txt(run_int(r[C_RUN - 1])))
               for r in linhas if txt(r[C_CLASSE - 1]) == 'OFICIAL' and txt(r[C_REGRA - 1])}
        obt = foto_ev.get((lote, (eq if bio else eq_hema) or ''), set())
        n += 1
        if esp != obt:
            dif.append(((eq, lote), sorted(esp - obt)[:3], sorted(obt - esp)[:3]))
    return dif, n


def main(produto, caminho, salvar=True):
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:                                          # noqa: BLE001
        pass
    caminho = os.path.abspath(caminho)
    bio = produto.startswith('Bio')
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.Calculation = -4135
        if ii.sheet(wb, ABA_EV) is None:
            raise SystemExit(f'aba {ABA_EV} ausente -- nada foi alterado')

        dif0 = vba_diferente(wb, produto)
        # o ADR-083 ja esta no arquivo quando o motor do mesmo nivel ja anda pela sequencia de ativos (SeqDoNivel):
        # mEstatistica diferente por OUTRO motivo nao e a instalacao desta regra
        ja083 = 'Function SeqDoNivel(' in _texto_componente(wb, 'mEstatistica')
        troca = 'mEstatistica' in dif0 and not ja083
        log(f'0. VBA do arquivo x fontes: {len(dif0)} componente(s) diferente(s)'
            + (f' ({", ".join(dif0[:8])}{" ..." if len(dif0) > 8 else ""})' if dif0 else '')
            + ('; mEstatistica ANTIGO (sem o ADR-083): instalacao da regra' if troca else
               '; o mEstatistica do arquivo ja tem o ADR-083' if ja083 else '; mEstatistica ja e o da fonte'))

        eq_hema = ''
        if not bio:
            try:
                eq_hema = txt(ex.run(f"'{wb.Name}'!mEstatistica.EquipFiltro", teto=120)).upper()
            except Exception:                                  # noqa: BLE001 -- VBA antigo sem EquipFiltro
                eq_hema = cfg(wb, 'EQUIPAMENTO_PADRAO').upper()
        alvos, sem_res = alvos_da_varredura(wb, bio, eq_hema)
        log(f'   varredura: {len(alvos)} (equipamento, lote) com resultado elegivel'
            + (f'; {len(sem_res)} lote(s) do cadastro sem resultado (sem evento por construcao)' if sem_res else ''))
        if bio and len({e for e, _ in alvos}) > 1:
            log('   AVISO (ADR-083): a BI_Data (mBI.FlagsDoMotor) agrupa nucleo|ANALITO SEM equipamento -- num analito '
                'com dois equipamentos o 2_2s ACROSS da BI_Data pode divergir de Eventos_Westgard e do Painel (que sao '
                'por equipamento); este relatorio cobre so Eventos_Westgard')
        sel = Seletor(wb, bio)
        painel0 = (txt(sel.cel['loteSel'].Value), txt(sel.cel['selEquipamento'].Value) if bio else '')

        log('1. VARREDURA ANTES (VBA do arquivo)')
        antes, s_antes = varrer(ex, wb, alvos, sel, 'antes')

        log('2. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        codigo_atual.instalar(ex, wb, produto, log)

        log('3. VARREDURA DEPOIS (VBA novo)')
        depois, s_depois = varrer(ex, wb, alvos, sel, 'depois')

        linhas, resumo, violacoes = comparar(antes, depois)
        log(f'4. RELATORIO: {resumo.get("evento_novo", 0)} evento(s) OFICIAL novo(s), '
            f'{resumo.get("evento_removido", 0)} removido(s); veredicto (nivel x RUN) OK->REJEITADO '
            f'{resumo.get("ok_para_rejeitado", 0)}, REJEITADO->OK {resumo.get("rejeitado_para_ok", 0)}; '
            f'COMPLEMENTAR {resumo.get("complementar_antes", 0)} -> {resumo.get("complementar_depois", 0)}; '
            f'OFICIAL {resumo.get("oficial_antes", 0)} -> {resumo.get("oficial_depois", 0)}')
        if violacoes:
            log(f'   ATENCAO: SUPERCONJUNTO VIOLADO -- linha(s) antiga(s) que nao estao na aba nova: {violacoes[:10]} '
                '(as removidas de classe OFICIAL estao no CSV como EVENTO)')
        else:
            log('   superconjunto OK: toda linha antiga (14 colunas) continua na aba nova')
        if not dif0 and linhas:
            raise SystemExit('o VBA ja era o das fontes e mesmo assim as duas varreduras diferem (motor nao '
                             f'deterministico?): {linhas[:5]} -- nada foi salvo')

        mudou = troca or bool(linhas)          # 2a execucao: False (relatorio vazio, nada gravado, MUDANCAS = 0)
        csv_ = nome_csv(caminho, linhas, previa=not salvar) if mudou else None
        if mudou:
            log(f'   CSV (gravado antes de salvar o .xlsm): '
                f'{csv_ or "ja existia e o relatorio novo esta vazio -- preservado"} ({len(linhas)} linha(s))')
        else:
            log('   relatorio vazio (o mEstatistica ja era o do ADR-083): CSV nao regravado')

        nd = descricoes_cfg_westgard(wb)
        if nd is None:
            log(f'4b. {ABA_CFG_WG}: aba ausente (nada a atualizar)')
        else:
            log(f'4b. {ABA_CFG_WG}: {nd} descricao(oes) de detector do mesmo nivel atualizada(s)')
            if nd:
                MUDANCAS.append(f'{ABA_CFG_WG}: {nd} descricao(oes) (observacoes ativas seguidas)')

        log('5. Painel de volta; motor e BI recalculados')
        sel.devolver()
        ex.xl.Calculation = -4105
        # sempre: a varredura deixou Eventos_Westgard/Eng_Saida no ultimo lote varrido
        for mac in ('mEstatistica.InvalidarCache', 'mEstatistica.AtualizarEstatistica'):
            ex.run(f"'{wb.Name}'!{mac}", teto=900)
        ex.esperar()
        # 27/10/2026: a BI nao e tocada pelo instalador (usuario: "nao mexer em nada de bi/power bi por enquanto").
        # A entrega 27 rodou mBI.AtualizarBIData aqui e a BI_Data passou de um retrato antigo para o historico inteiro
        log('   BI_Data nao tocada (a BI fica com o laboratorio: botao "Atualizar BI")'
            + ('; o Westgard mudou -- a BI so reflete isso depois do "Atualizar BI"' if mudou else ''))

        if mudou:
            det = (('ADR-083: regras de Westgard do MESMO nivel na sequencia de ativos do nivel (entre niveis pelo RUN). '
                    if troca else 'ADR-083 (o arquivo ja tinha a regra; o VBA do motor mudou e o Westgard tambem). ') +
                   f'Varredura de {len(alvos)} lote(s)x equipamento, todo o historico: eventos OFICIAIS '
                   f'{resumo.get("oficial_antes", 0)} -> {resumo.get("oficial_depois", 0)} '
                   f'(+{resumo.get("evento_novo", 0)} / -{resumo.get("evento_removido", 0)}); veredictos nivel x RUN '
                   f'OK->REJEITADO {resumo.get("ok_para_rejeitado", 0)}, REJEITADO->OK '
                   f'{resumo.get("rejeitado_para_ok", 0)}; COMPLEMENTAR {resumo.get("complementar_antes", 0)} -> '
                   f'{resumo.get("complementar_depois", 0)}. Lista: {os.path.basename(csv_) if csv_ else "(CSV ja existente)"}')
            ex.run(f"'{wb.Name}'!mAuditoria.RegistrarLog", 'REGRA_WESTGARD_ALTERADA', det, teto=120)
            MUDANCAS.append('Audit_Log REGRA_WESTGARD_ALTERADA')

        log(f'5b. {ABA_FOTOS}: foto dos eventos OFICIAIS por lote x equipamento (sem auditar)')
        ws_fotos, fotos_criada = montar_aba_fotos(wb)
        forcar = fotos_criada or 'mEstatistica' in dif0
        t0 = time.time()
        r_foto = txt(ex.run(f"'{wb.Name}'!mEstatistica.GravarFotosWestgard", forcar, teto=1800))
        log(f'   {"aba criada; " if fotos_criada else ""}GravarFotosWestgard(forcar={forcar}): {r_foto} '
            f'({time.time() - t0:.1f} s)')
        if not r_foto.startswith('RECALC=') or 'FALHAS=' in r_foto:
            raise SystemExit(f'{ABA_FOTOS}: a foto nao foi gravada ({r_foto}) -- nada foi salvo')
        if fotos_criada:
            MUDANCAS.append(f'{ABA_FOTOS}: aba criada')
        if not r_foto.startswith('RECALC=0/') or 'FOTO_INICIAL' in r_foto:
            MUDANCAS.append(f'{ABA_FOTOS}: foto gravada ({r_foto})')

        log('6. FUMACA')
        lt, eqa, fin = ler_eventos(wb)
        lote_p = txt(ex.run(f"'{wb.Name}'!mLotes.LotePainel", teto=120))
        eq_p = txt(ex.run(f"'{wb.Name}'!mEstatistica.EquipFiltro", teto=120)).upper()
        if lt != lote_p or eqa != eq_p:
            raise SystemExit(f'fumaca: Eventos_Westgard do lote {lt!r}/{eqa!r}, o Painel esta em {lote_p!r}/{eq_p!r} '
                             f'(antes da varredura: {painel0}) -- nada foi salvo')
        alvo_p = next((k for k in depois if mesmo_lote(k[1], lt) and (k[0] or '') == (eqa if bio else '')), None)
        if alvo_p is not None:
            igual = collections.Counter(tuple(norm_cel(x) for x in r) for r in fin) == \
                collections.Counter(tuple(norm_cel(x) for x in r) for r in depois[alvo_p])
            if not igual:
                raise SystemExit('fumaca: a aba Eventos_Westgard final difere da varredura do mesmo lote '
                                 '-- nada foi salvo')
            log(f'   Eventos_Westgard do lote do Painel ({lt}{" / " + eqa if bio else ""}) = varredura: '
                f'{len(fin)} evento(s)')
        foto_ev, n_ctrl = ler_foto(wb)
        if foto_ev is None:
            raise SystemExit(f'fumaca: {ABA_FOTOS} incompleta (J2 vazio) -- nada foi salvo')
        dif_f, n_conf = conferir_foto(foto_ev, depois, bio, eq_hema)
        if dif_f:
            raise SystemExit(f'fumaca: a foto de {ABA_FOTOS} difere da varredura depois em {len(dif_f)} lote(s)x '
                             f'equipamento (so na varredura / so na foto): {dif_f[:5]} -- nada foi salvo')
        log(f'   {ABA_FOTOS} = varredura depois: {n_conf} lote(s)x equipamento conferido(s), {n_ctrl} linha(s) '
            f'CONTROLE, {sum(len(s) for s in foto_ev.values())} evento(s) OFICIAL(is)')
        log(f'   tempos: varredura antes {s_antes:.1f} s, depois {s_depois:.1f} s')
        log(f'MUDANCAS {len(MUDANCAS)}' + (': ' + '; '.join(MUDANCAS) if MUDANCAS else ''))

        # a aba da foto e tecnica: protegida como o Sync_Log, mesmo que nao estivesse (ou acabasse de ser criada)
        foto[ABA_FOTOS] = dict(foto_protecao(wb)[ABA_FOTOS], conteudo=True, objetos=False, cenarios=True)
        devolver_protecao(wb, foto)
        if not ws_fotos.ProtectContents:
            raise SystemExit(f'{ABA_FOTOS} ficou sem protecao -- nada foi salvo')
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if csv_:
            # ANTES de salvar: o .xlsm salvo cita este CSV (REGRA_WESTGARD_ALTERADA) e, com o VBA novo, uma 2a execucao
            # ja nao consegue gerar o relatorio (ele so existe enquanto o arquivo tem o VBA antigo)
            try:
                gravar_csv(csv_, linhas)
            except OSError as e:
                raise SystemExit(f'relatorio NAO gravado ({csv_}): {e} -- nada foi salvo')
            log(f'relatorio: {csv_}' + ('' if salvar else ' (--sem-salvar: previa; o .xlsm NAO foi gravado)'))
        if salvar:
            try:
                wb.Save()
            except Exception:
                if csv_:                                       # descreve uma mudanca que nao foi gravada
                    try:
                        os.remove(csv_)
                    except OSError:
                        pass
                raise
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], salvar='--sem-salvar' not in a)
