# -*- coding: utf-8 -*-
"""repro_e48_alvos.py -- Tester/QA 25/11/2026, VALIDAR entrega 48 (ADR-102: alvos do laboratorio por lote e periodo,
excecao temporaria e referencia bula/Unity). Script PROPRIO do Tester: NAO reaproveita o qa_alvos_manual.py (so a
infraestrutura qa_final.QA). Cobre o que a suite nao prova diretamente.

Uso (FASE 2, depois da entrega, com NENHUM outro Excel de teste aberto):
    python -u repro_e48_alvos.py <Bioquimica|Hematologia> <saida.json> [--origem <QC_xxx.xlsm>]

  --origem  arquivo de producao a COPIAR (padrao: QC_RAIZ\\QC_<produto>.xlsm ou a raiz do repositorio). A origem nunca e
            aberta pelo Excel: shutil.copy2 para %TEMP%\\qcw\\auditoria\\e48\\ e tudo roda na copia; nada e salvo.

Oraculo: Python puro (statistics.fmean / statistics.stdev, DP amostral n-1) sobre a tblCQ_Final lida pela COM
(matriz de valores), com os filtros do ADR-102 escritos aqui (PARTICIPA_ESTATISTICA = SIM, LOTE e ANALITO exatos sem
caixa, todos os equipamentos, DATA_HORA por dia entre De e Ate inclusive, vazio = sem limite). Nenhuma funcao VBA de
calculo e usada como referencia (ApuracaoTexto so aparece como evidencia).

Cenarios (estado: APROVADO / REPROVADO / NAO_EXECUTADO / BLOQUEADO):
  X00 pre-condicoes (secao instalada, lote x analito com chave unica, lote em tela = lote escolhido)
  X01 resultado INATIVADO dentro do periodo nao entra na apuracao (inativacao sintetica TESTE-AUDITORIA na copia +
      ATUALIZAR so entradas): N cai 1, Media/DP = oraculo sem ele
  X02 todos os equipamentos entram juntos (Bio: >= 2 equipamentos); o equipamento do Painel (L4, so Bio) nao muda a
      apuracao
  X03 datas inclusivas: Ate = dia D inclui os resultados de D com hora; Ate = D-1 exclui; De = D inclui D; vazio = sem
      limite
  X04 valores ESTATICOS: depois de Aplicar, mudar lote/datas/analito/lote em analise e rodar ATUALIZAR nao muda a
      linha do lote na LotesStore (C:H, W, X:Z, AA:AO) nem nenhuma C:H da LotesStore (filtros)
  X05 LJ e Westgard passam a usar o alvo gravado: LotesStore = oraculo; Calc (lsMedN/lsDPN), series Media/+3s do LJ;
      1_3s OFICIAL de Eventos_Westgard = oraculo |x - M| > 3*DP sobre os pontos do Eng_Saida
  X06 excecao SO DO DP (bula) seguida de "Aplicar" modo 3 (so o DP): a marca EXCECAO do DP sai, a Media fica,
      CONDICAO_N1 e o amarelo coerentes com a regra da marca por campo
  X07 trocar o analito (C3) tranca a digitacao liberada; analito trocado sem evento (spinner) -> gravar e recusado
  X08 referencia: resultado IGUAL ao LI ou ao LS e dentro (sem observacao); fora = PENDENTE (= oraculo); Westgard e
      LotesStore inalterados pela grade
  X09 resultado com observacao PENDENTE inativado -> RETIRADA no ATUALIZAR
  X10 resultado INATIVADO fora da faixa nunca gera observacao; LotesStore inalterada pela grade
  Y00 arquivo de origem intacto (SHA-256)
"""
import collections
import datetime as dt
import hashlib
import json
import os
import shutil
import statistics
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
TESTES = os.path.abspath(os.path.join(AQUI, '..', '..', 'integracao_dbseac', 'testes'))
sys.path.insert(0, TESTES)
sys.path.insert(0, os.path.join(TESTES, '..'))
from qa_final import QA  # noqa: E402
from instalar_integracao import SENHA  # noqa: E402   (senha so por import, nunca no texto)

TAG = 'TESTE-AUDITORIA E48'
R_ENT, R_NV1, R_MAN_TIT, R_MAN1, R_MAN_JUST, R_REF1, N_REF = 98, 100, 105, 107, 110, 115, 40
AMARELO = 156 * 65536 + 235 * 256 + 255          # RGB(255, 235, 156)
EXC = 'EXCECAO_TEMPORARIA'
RES = []


# ------------------------------------------------------------------ utilidades
def reg(cod, estado, desc, evid):
    RES.append({'id': cod, 'estado': estado, 'descricao': desc, 'evidencia': evid})
    print(f'{estado:13s} {cod} {desc} -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:700], flush=True)


def ok(cod, cond, desc, evid):
    reg(cod, 'APROVADO' if cond else 'REPROVADO', desc, evid)


def num(v):
    return (isinstance(v, (int, float)) and not isinstance(v, bool)
            and not (isinstance(v, int) and v < -2146820000) and v == v)


def txt(v):
    if isinstance(v, float) and v == v and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def U(v):
    return txt(v).upper()


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


def dia(v):
    if isinstance(v, dt.datetime):
        return v.replace(tzinfo=None).date()
    return None


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def nivel(r):
    try:
        return int(float(r.get('NIVEL')))
    except (TypeError, ValueError):
        return 0


def ativo(r):
    return U(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))


# ------------------------------------------------------------------ oraculo (independente da VBA)
def apurar(rows, lote, an, ini, fim, nlv):
    """{nv: (n, media, dp)} -- ADR-102 secao 1 (decisoes tecnicas)."""
    out = {}
    for nv in range(1, nlv + 1):
        xs = []
        for r in rows:
            if not ativo(r) or U(r.get('LOTE')) != U(lote) or U(r.get('ANALITO')) != U(an) or nivel(r) != nv:
                continue
            d = dia(r.get('DATA_HORA'))
            if ini is not None or fim is not None:
                if d is None:
                    continue
                if ini is not None and d < ini:
                    continue
                if fim is not None and d > fim:
                    continue
            xs.append(float(r['RESULTADO']))
        m = statistics.fmean(xs) if xs else None
        s = statistics.stdev(xs) if len(xs) >= 2 else None
        out[nv] = (len(xs), m, s)
    return out


def cond_n(n, s, rec, ideal):
    if n < 2:
        return 'BLOQUEADO'
    if not (num(s) and s > 0):
        return 'DP_ZERO'
    return 'TEMPORARIO' if n < rec else 'ACEITAVEL' if n < ideal else 'IDEAL'


def marcas(cond):
    """Marca por campo do CONDICAO_Nk (revisao 25/11 do ADR-102), escrita aqui a partir do texto do ADR."""
    c = U(cond)
    if not c:
        return '', ''
    if c == EXC:
        return 'EXCECAO', 'EXCECAO'
    if c == 'TEMPORARIO':
        return 'TEMPORARIO', 'TEMPORARIO'
    tm = td = ''
    for t in [x.strip() for x in c.split('+')]:
        if t == EXC + '_MEDIA':
            tm = 'EXCECAO'
        if t == EXC + '_DP':
            td = 'EXCECAO'
        if t == 'TEMPORARIO_MEDIA':
            tm = 'TEMPORARIO'
        if t == 'TEMPORARIO_DP':
            td = 'TEMPORARIO'
    return tm, td


def propria(cond):
    for t in [x.strip() for x in U(cond).split('+')]:
        if t in ('ACEITAVEL', 'IDEAL'):
            return t
    return ''


# ------------------------------------------------------------------ execucao
def executar(produto, origem, saida):
    h0 = sha(origem)
    pasta = os.path.join(os.environ.get('TEMP', '.'), 'qcw', 'auditoria', 'e48')
    os.makedirs(pasta, exist_ok=True)
    copia = os.path.join(pasta, f'repro_{produto}_{time.strftime("%Y%m%d_%H%M%S")}.xlsm')
    shutil.copy2(origem, copia)
    print('COPIA', copia, 'ORIGEM_SHA', h0[:12], flush=True)
    q = None
    try:
        q = QA(produto, copia)
        xl, nlv, bio = q.ex.xl, q.nlv, q.bio
        p, ls = q.ws('Painel'), q.ws('LotesStore')

        def run(m, *a, teto=900):
            return q.run(m if '.' in m else 'mAlvosRef.' + m, *a, teto=teto)

        def resp(s):
            run('DefinirRespostasTeste', s, teto=60)

        def esperar():
            try:
                q.ex.esperar()
            except Exception:                 # noqa: BLE001
                pass

        def escrever(rng, valor, eventos=True):
            ev0 = xl.EnableEvents
            xl.EnableEvents = eventos
            try:
                try:
                    rng.Value = valor
                except Exception:             # noqa: BLE001  (celula travada: desprotege a aba so na copia)
                    rng.Worksheet.Unprotect(SENHA)
                    rng.Value = valor
            finally:
                xl.EnableEvents = ev0
            esperar()

        def secao(lote, ini, fim):
            """B98/E98 sem evento, G98 com evento: UM Worksheet_Change -> mAlvosRef.SecaoMudou."""
            escrever(p.Range(f'B{R_ENT}'), lote, eventos=False)
            escrever(p.Range(f'E{R_ENT}'), dt.datetime.combine(ini, dt.time()) if ini else None, eventos=False)
            escrever(p.Range(f'G{R_ENT}'), dt.datetime.combine(fim, dt.time()) if fim else None, eventos=True)

        def vista():
            v = p.Range(f'B{R_NV1}:D{R_NV1 + nlv - 1}').Value
            return {nv: tuple(v[nv - 1]) for nv in range(1, nlv + 1)}

        def bate(vs, ora):
            div = []
            for nv in range(1, nlv + 1):
                n, m, s = ora[nv]
                vn, vm, vsd = vs[nv]
                if int(vn or 0) != n:
                    div.append((nv, 'N', vn, n))
                elif n >= 1 and not perto(vm, m):
                    div.append((nv, 'Media', vm, m))
                elif n >= 2 and not perto(vsd, s):
                    div.append((nv, 'DP', vsd, s))
            return div

        def ler_cq():
            return q.ler('tblCQ_Final')

        def ult_ls():
            return ls.Cells(ls.Rows.Count, 16).End(-4162).Row

        def foto_ls(ate=None):
            u = (ate or ult_ls()) + 3
            return [tuple(r) for r in ls.Range(f'A1:AO{u}').Value]

        def eventos_wg():
            ws = q.ws('Eventos_Westgard')
            n = int(float(ws.Range('J2').Value or 0))
            if n <= 0:
                return []
            return [tuple(r) for r in ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 14)).Value]

        def amarelo():
            return int(p.Range(f'I{R_NV1}').Interior.Color or 0) == AMARELO

        def estado_man():
            return txt(run('EstadoManualTexto', teto=60))

        def obs():
            o = q.ws('RefBulaObs')
            u = o.Cells(o.Rows.Count, 1).End(-4162).Row
            return [tuple(r) for r in o.Range(f'A2:Q{u}').Value] if u >= 2 else []

        def atualizar_entradas():
            t0 = time.time()
            r = txt(run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            esperar()
            return r, round(time.time() - t0, 1)

        # ============================================================ X00 pre-condicoes
        if not run('SecaoInstalada', teto=60):
            reg('X00', 'BLOQUEADO', 'secao do ADR-102 nao instalada na copia', {'copia': copia})
            return
        rec, ideal = [int(x) for x in txt(run('LimitesNTexto', teto=60)).split('|')]
        rows = ler_cq()
        espec = q.espec()
        catalogo = [txt(v[0]) for v in espec.Range('A4:A43').Value]
        cat_u = [U(x) for x in catalogo]
        u = ult_ls()
        chaves = collections.Counter(U(v[0]) for v in ls.Range(f'P2:P{u}').Value if txt(v[0]))
        grupos = collections.defaultdict(list)
        for r in rows:
            if num(r.get('RESULTADO')) and txt(r.get('LOTE')) and txt(r.get('ANALITO')):
                grupos[(U(r['LOTE']), U(r['ANALITO']))].append(r)
        cand = []
        for (lt, an), rs in grupos.items():
            if chaves.get(f'{lt}|{an}') != 1 or an not in cat_u:
                continue
            at = [r for r in rs if ativo(r)]
            n1 = sum(1 for r in at if nivel(r) == 1 and dia(r.get('DATA_HORA')))
            neq = len({U(r.get('EQUIPAMENTO')) for r in at})
            ninat = sum(1 for r in rs if not ativo(r) and nivel(r) == 1)
            if n1 < 8:
                continue
            cand.append(((neq >= 2) if bio else True, n1 >= rec + 5, ninat > 0, n1, lt, an))
        if not cand:
            reg('X00', 'BLOQUEADO', 'nenhum lote x analito com chave unica e N1 >= 8 ativos', {})
            return
        cand.sort(reverse=True)
        LT, AN = cand[0][4], cand[0][5]
        lote = next(txt(r['LOTE']) for r in grupos[(LT, AN)])
        analito = catalogo[cat_u.index(AN)]
        # lote em analise (H3) e analito do Painel = os do teste, para o LJ/Westgard do X05
        if U(run('mLotes.LoteEmTela', teto=60)) != LT:
            escrever(p.Range('H3'), "'" + lote, eventos=True)
            xl.Calculate()
        if U(q.nome('selAnalito').Value) != AN:
            escrever(p.Range('B3'), cat_u.index(AN) + 1, eventos=False)
            run('mEstatistica.PainelMudou', teto=900)
            esperar()
        rLs = [i for i, v in enumerate(ls.Range(f'P2:P{u}').Value, 2) if U(v[0]) == f'{LT}|{AN}'][0]
        pre = {'lote': lote, 'analito': analito, 'limites_N': [rec, ideal], 'linha_LotesStore': rLs,
               'candidato': cand[0][:4], 'lote_em_tela': txt(run('mLotes.LoteEmTela', teto=60)),
               'selAnalito': txt(q.nome('selAnalito').Value), 'copia': copia, 'origem_sha': h0[:12]}
        ok('X00', U(pre['lote_em_tela']) == LT and U(pre['selAnalito']) == AN,
           'pre-condicoes: lote x analito com chave unica; lote em tela e analito do Painel = os do teste', pre)

        def linha():
            return tuple(ls.Range(f'A{rLs}:AO{rLs}').Value[0])

        def datas_n1(rs):
            return sorted(dia(r['DATA_HORA']) for r in rs if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN
                          and nivel(r) == 1 and dia(r.get('DATA_HORA')))

        # ============================================================ X03 datas inclusivas
        try:
            ds = datas_n1(rows)
            com_hora = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and nivel(r) == 1
                        and isinstance(r.get('DATA_HORA'), dt.datetime) and r['DATA_HORA'].time() != dt.time()]
            D = dia(com_hora[len(com_hora) // 2]['DATA_HORA']) if com_hora else ds[len(ds) // 2]
            d0 = ds[0]
            casos = {'ate_D': (d0, D), 'ate_D-1': (d0, D - dt.timedelta(days=1)), 'de_D': (D, None),
                     'so_D': (D, D), 'sem_limite': (None, None)}
            ev3, div3 = {}, []
            for k, (a, b) in casos.items():
                secao(lote, a, b)
                ora = apurar(rows, lote, AN, a, b, nlv)
                vs = vista()
                d_ = bate(vs, ora)
                ev3[k] = {'periodo': [str(a), str(b)], 'N_vista': [vs[nv][0] for nv in vs],
                          'N_oraculo': [ora[nv][0] for nv in ora]}
                if d_:
                    div3.append((k, d_[:3]))
            n_D = sum(1 for x in ds if x == D)
            incl = ev3['ate_D']['N_oraculo'][0] - ev3['ate_D-1']['N_oraculo'][0] == n_D and n_D >= 1
            ok('X03', not div3 and incl and ev3['so_D']['N_oraculo'][0] == n_D,
               'datas inclusivas (dia D com hora entra em Ate = D e em De = D; sai em Ate = D-1; vazio = sem limite): '
               'apuracao do Painel = oraculo', {'D': str(D), 'resultados_N1_em_D': n_D, 'casos': ev3,
                                                'divergencias': div3, 'com_hora': bool(com_hora)})
        except Exception as e:                # noqa: BLE001
            reg('X03', 'REPROVADO', 'datas inclusivas: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X02 todos os equipamentos
        try:
            at = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN]
            por_eq = collections.Counter((U(r.get('EQUIPAMENTO')) or '(sem)', nivel(r)) for r in at)
            eqs = sorted({k[0] for k in por_eq})
            secao(lote, None, None)
            ora = apurar(rows, lote, AN, None, None, nlv)
            vs = vista()
            div = bate(vs, ora)
            ev2 = {'equipamentos': eqs, 'N_por_equipamento_nivel': {f'{a}|N{b}': c for (a, b), c in por_eq.items()},
                   'N_vista': [vs[nv][0] for nv in vs], 'N_oraculo': [ora[nv][0] for nv in ora], 'divergencias': div}
            l4_ok = True
            if bio and len(eqs) >= 2:
                l4_0 = p.Range('L4').Value
                escrever(p.Range('L4'), next(r.get('EQUIPAMENTO') for r in at), eventos=True)
                secao(lote, None, None)
                vs2 = vista()
                l4_ok = not bate(vs2, ora)
                ev2['L4_trocado_para'] = txt(p.Range('L4').Value)
                ev2['N_vista_com_L4'] = [vs2[nv][0] for nv in vs2]
                escrever(p.Range('L4'), l4_0, eventos=True)
                ev2['L4_restaurado'] = txt(p.Range('L4').Value)
            if bio and len(eqs) < 2:
                reg('X02', 'NAO_EXECUTADO', 'todos os equipamentos: o lote escolhido tem 1 equipamento so (dados)', ev2)
            else:
                maior = max(c for (a, b), c in por_eq.items() if b == 1)
                ok('X02', not div and l4_ok and (not bio or ora[1][0] > maior),
                   'todos os equipamentos entram juntos (N do lote > N de qualquer equipamento, Bio); o L4 do Painel '
                   'nao muda a apuracao', ev2)
        except Exception as e:                # noqa: BLE001
            reg('X02', 'REPROVADO', 'equipamentos: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X01 inativado no periodo nao entra
        inat_id, inat_val = None, None
        try:
            ds = datas_n1(rows)
            ini, fim = ds[0], ds[-1]
            alvo = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and nivel(r) == 1
                    and dia(r.get('DATA_HORA')) and ini < dia(r['DATA_HORA']) < fim and txt(r.get('ID_REGISTRO'))]
            if not alvo:
                alvo = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and nivel(r) == 1
                        and txt(r.get('ID_REGISTRO'))]
            x = alvo[len(alvo) // 2]
            inat_id, inat_val = txt(x['ID_REGISTRO']), float(x['RESULTADO'])
            ora_antes = apurar(rows, lote, AN, ini, fim, nlv)
            ja_inat = sum(1 for r in rows if not ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN
                          and num(r.get('RESULTADO')))
            q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': inat_id, 'ANALITO': txt(x['ANALITO']),
                                                            'MOTIVO': f'Outro: {TAG} (repro, so na copia)'})
            r_at, t_at = atualizar_entradas()
            rows = ler_cq()
            depois = [r for r in rows if txt(r.get('ID_REGISTRO')) == inat_id and U(r.get('ANALITO')) == AN]
            ficou_inativo = bool(depois) and all(U(r.get('PARTICIPA_ESTATISTICA')) != 'SIM' for r in depois)
            secao(lote, ini, fim)
            ora = apurar(rows, lote, AN, ini, fim, nlv)
            vs = vista()
            div = bate(vs, ora)
            ok('X01', r_at.startswith('OK|') and ficou_inativo and not div and ora[1][0] == ora_antes[1][0] - 1,
               'resultado inativado DENTRO do periodo nao entra (N1 cai 1; Media/DP = oraculo sem ele)',
               {'ID': inat_id, 'valor': inat_val, 'periodo': [str(ini), str(fim)], 'atualizar': [r_at[:80], t_at],
                'PARTICIPA_depois': [txt(r.get('PARTICIPA_ESTATISTICA')) for r in depois],
                'N1_antes_depois': [ora_antes[1][0], ora[1][0]], 'vista': vs, 'divergencias': div,
                'inativos_preexistentes_no_lote': ja_inat,
                'ApuracaoTexto_VBA': txt(run('ApuracaoTexto', lote, analito,
                                             float((dt.datetime.combine(ini, dt.time()) - dt.datetime(1899, 12, 30)).days),
                                             float((dt.datetime.combine(fim, dt.time()) - dt.datetime(1899, 12, 30)).days),
                                             teto=120))[:200]})
        except Exception as e:                # noqa: BLE001
            reg('X01', 'REPROVADO', 'inativado no periodo: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X05 Aplicar -> LotesStore, LJ e Westgard
        alvo_m = {}
        try:
            ds = datas_n1(rows)
            fim_a = next((d for d in sorted(set(ds)) if rec <= sum(1 for x in ds if x <= d) < ideal), ds[-1])
            ini_a = ds[0]
            antes = linha()
            m_old = {nv: (antes[2 + 2 * (nv - 1)], antes[3 + 2 * (nv - 1)]) for nv in range(1, nlv + 1)}
            secao(lote, ini_a, fim_a)
            ora = apurar(rows, lote, AN, ini_a, fim_a, nlv)
            resp('1|SIM|SIM')
            r_ap = txt(run('AplicarAlvos'))
            resp('')
            esperar()
            xl.Calculate()
            lin = linha()
            div = []
            for nv in range(1, nlv + 1):
                n, m, s = ora[nv]
                c = cond_n(n, s, rec, ideal)
                mv, sv = lin[2 + 2 * (nv - 1)], lin[3 + 2 * (nv - 1)]
                if c in ('BLOQUEADO', 'DP_ZERO'):
                    if (mv, sv) != m_old[nv]:
                        div.append((nv, 'nivel mantido mudou'))
                    continue
                if not (perto(mv, m) and perto(sv, s)):
                    div.append((nv, 'LotesStore != oraculo', mv, sv, m, s))
                if U(lin[26 + 5 * (nv - 1) + 3]) != c:
                    div.append((nv, 'CONDICAO', lin[26 + 5 * (nv - 1) + 3], c))
                alvo_m[nv] = (float(mv), float(sv))
            ok('X05a', r_ap.startswith('OK|') and not div,
               '"Aplicar Alvos ao Lote" modo 1: LotesStore C:H = oraculo do periodo; CONDICAO_Nk = regra do N',
               {'periodo': [str(ini_a), str(fim_a)], 'retorno': r_ap, 'oraculo': ora, 'antes': m_old,
                'divergencias': div, 'W': txt(lin[22])[:200]})
            # --- LJ: Calc (formulas com lsMedN/lsDPN) e series do grafico
            c_ = q.ws('Calc')
            f1 = [str(c_.Cells(1, k).Formula or '').replace(' ', '') for k in range(1, 131)]
            div_lj, ev_lj = [], {}
            for nv, (m, s) in alvo_m.items():
                cm = next((k + 1 for k, f in enumerate(f1) if f'>0,INDEX(lsMedN{nv},' in f), None)
                cs = next((k + 1 for k, f in enumerate(f1) if f'>0,INDEX(lsDPN{nv},' in f), None)
                vm = c_.Cells(1, cm).Value if cm else None
                vsd = c_.Cells(1, cs).Value if cs else None
                ser = {}
                try:
                    ch = p.ChartObjects(nv).Chart
                    for i in range(1, ch.SeriesCollection().Count + 1):
                        sr = ch.SeriesCollection(i)
                        try:
                            ser[str(sr.Name)] = [y for y in list(sr.Values) if num(y)]
                        except Exception:     # noqa: BLE001
                            pass
                except Exception as e:        # noqa: BLE001
                    ser = {'erro': repr(e)}
                ym, y3 = ser.get('Média', []), ser.get('+3s', [])
                ev_lj[nv] = {'Calc': [vm, vsd], 'serie_Media': ym[:2], 'serie_+3s': y3[:2], 'alvo': [m, s]}
                if not (perto(vm, m) and perto(vsd, s)):
                    div_lj.append((nv, 'Calc'))
                if not ym or not all(perto(y, m) for y in ym) or not y3 or not all(perto(y, m + 3 * s) for y in y3):
                    div_lj.append((nv, 'series do LJ'))
                if not perto(q.run('mLotes.MediaDoLote', lote, analito, nv, teto=60), m):
                    div_lj.append((nv, 'mLotes.MediaDoLote'))
            ok('X05b', bool(alvo_m) and not div_lj,
               'LJ usa o alvo gravado: Calc (lsMedN/lsDPN), series Media e +3s do grafico e mLotes.MediaDoLote',
               {'niveis': ev_lj, 'divergencias': div_lj})

            # --- Westgard: 1_3s OFICIAL = |x - M| > 3*DP nos pontos do Eng_Saida (oraculo aqui)
            def conferir_wg():
                eng = q.eng_saida()
                if U(eng.get('lote')) != LT or U(eng.get('analito')) != AN or not eng['n']:
                    return None, {'eng_lote': eng.get('lote'), 'eng_analito': eng.get('analito'), 'n': eng['n']}
                # o Eng_Saida e mesmo o valor do resultado? (confere contra a tblCQ_Final antes de usar)
                vals_run = collections.defaultdict(list)
                for r in rows:
                    if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and num(r.get('RUN')):
                        vals_run[(int(float(r['RUN'])), nivel(r))].append(float(r['RESULTADO']))
                tot = casam = 0
                ora_new, ora_old, uni = collections.defaultdict(set), collections.defaultdict(set), set()
                for i, run_ in enumerate(eng['runs']):
                    for nv in alvo_m:
                        x_ = eng['val'][i][nv - 1] if i < len(eng['val']) else None
                        if not num(x_):
                            continue
                        tot += 1
                        cands = vals_run.get((run_, nv), [])
                        if any(perto(x_, y) for y in cands):
                            casam += 1
                        if len(cands) != 1:
                            continue
                        uni.add((run_, nv))
                        m, s = alvo_m[nv]
                        if abs(x_ - m) > 3 * s:
                            ora_new[nv].add(run_)
                        mo, so = m_old[nv]
                        if num(mo) and num(so) and so > 0 and abs(x_ - mo) > 3 * so:
                            ora_old[nv].add(run_)
                if tot == 0 or casam < 0.9 * tot:
                    return None, {'pontos': tot, 'casam_com_tblCQ_Final': casam}
                evs = collections.defaultdict(set)
                for e_ in eventos_wg():
                    if (U(e_[2]) == AN and txt(e_[4]) == '1_3s' and U(e_[7]) == 'OFICIAL' and num(e_[1])
                            and txt(e_[3]) in [str(nv) for nv in alvo_m]):
                        nv = int(txt(e_[3]))
                        if (int(e_[1]), nv) in uni:
                            evs[nv].add(int(e_[1]))
                igual = all(evs[nv] == ora_new[nv] for nv in alvo_m)
                return igual, {'eventos_1_3s': {nv: sorted(evs[nv]) for nv in alvo_m},
                               'oraculo_alvo_novo': {nv: sorted(ora_new[nv]) for nv in alvo_m},
                               'oraculo_alvo_antigo': {nv: sorted(ora_old[nv]) for nv in alvo_m},
                               'discriminante': any(ora_new[nv] != ora_old[nv] for nv in alvo_m),
                               'pontos_unicos': len(uni)}
            igual, evw = conferir_wg()
            if igual is None:
                reg('X05c', 'NAO_EXECUTADO', 'Westgard: Eng_Saida nao e do lote/analito do teste ou nao confere com a '
                    'tblCQ_Final (oraculo nao aplicavel)', evw)
            elif igual:
                ok('X05c', True, 'Westgard (1_3s OFICIAL) usa o alvo gravado logo depois do Aplicar (= oraculo)', evw)
            else:
                r_at, t_at = atualizar_entradas()
                rows = ler_cq()
                igual2, evw2 = conferir_wg()
                reg('X05c', 'PARCIAL' if igual2 else 'REPROVADO',
                    'Westgard (1_3s OFICIAL) = oraculo com o alvo novo: ' +
                    ('so depois do ATUALIZAR' if igual2 else 'nem depois do ATUALIZAR'),
                    {'imediato': evw, 'apos_atualizar': evw2, 'atualizar': [r_at[:80], t_at]})
        except Exception as e:                # noqa: BLE001
            reg('X05', 'REPROVADO', 'Aplicar/LJ/Westgard: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X04 valores estaticos
        try:
            antes_lin = linha()
            antes_ch = [tuple(r) for r in ls.Range(f'C2:H{ult_ls()}').Value]
            outros = [txt(v[0]) for v in q.nome('regLoteCol').Value if txt(v[0]) and U(v[0]) != LT]
            ds = datas_n1(rows)
            passos = []
            if outros:
                secao(outros[0], ds[0], ds[-1])
                passos.append(f'B98 -> {outros[0]}')
            secao(lote, ds[len(ds) // 3], ds[-1])
            passos.append('datas trocadas')
            secao(lote, None, None)
            passos.append('datas vazias')
            outro_an = next((x for x in catalogo if x and U(x) != AN), None)
            if outro_an:
                escrever(p.Range('C3'), outro_an, eventos=True)
                escrever(p.Range('C3'), analito, eventos=True)
                passos.append(f'C3 -> {outro_an} -> {analito}')
            if outros:
                escrever(p.Range('H3'), "'" + outros[0], eventos=True)
                escrever(p.Range('H3'), "'" + lote, eventos=True)
                passos.append(f'H3 -> {outros[0]} -> {lote}')
            run('AtualizarTudo')
            run('VerificarReferencia')
            passos.append('AtualizarTudo + VerificarReferencia')
            lin_f = linha()
            ch_f = [tuple(r) for r in ls.Range(f'C2:H{ult_ls()}').Value]
            dif_ch = [i + 2 for i, (a, b) in enumerate(zip(antes_ch, ch_f)) if a != b]
            r_at, t_at = atualizar_entradas()
            lin_a = linha()
            dif_col = [i for i, (a, b) in enumerate(zip(antes_lin, lin_a)) if a != b and i >= 2]
            ok('X04', lin_f == antes_lin and not dif_ch and len(ch_f) == len(antes_ch) and r_at.startswith('OK|')
               and not dif_col,
               'valores ESTATICOS: trocar lote/datas/analito/lote em analise e o ATUALIZAR nao mudam o alvo gravado',
               {'passos': passos, 'linhas_C:H_diferentes_apos_filtros': dif_ch[:10],
                'colunas_da_linha_diferentes_apos_ATUALIZAR': dif_col, 'atualizar': [r_at[:80], t_at],
                'selAnalito': txt(q.nome('selAnalito').Value), 'lote_em_tela': txt(run('mLotes.LoteEmTela', teto=60))})
        except Exception as e:                # noqa: BLE001
            reg('X04', 'REPROVADO', 'estatico: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X06 excecao so do DP -> Aplicar modo 3
        try:
            if U(q.nome('selAnalito').Value) != AN:
                escrever(p.Range('B3'), cat_u.index(AN) + 1, eventos=False)
                run('mEstatistica.PainelMudou', teto=900)
            ds = datas_n1(rows)
            secao(lote, ds[0], ds[-1])
            run('BloquearManual', teto=60)
            antes = linha()
            c_ant = txt(antes[29])
            resp('SIM')
            r_lib = txt(run('LiberarManualCore'))
            resp('')
            dp_bula = round(float(antes[3]) * 1.11 + 0.0007, 6)
            escrever(p.Range(f'C{R_MAN1}'), dp_bula)
            escrever(p.Range(f'D{R_MAN1}'), 'Bula')
            escrever(p.Range(f'B{R_MAN_JUST}'), f'{TAG}: DP da bula por excecao, media propria mantida.')
            resp('SIM|SIM')
            r_g = txt(run('GravarManualCore'))
            resp('')
            exc = linha()
            tm0, td0 = marcas(antes[29])
            esp_exc_tm = tm0
            ev6 = {'liberar': r_lib, 'gravar': r_g, 'cond_antes': c_ant, 'cond_excecao': txt(exc[29]),
                   'dp': [antes[3], exc[3], dp_bula], 'media': [antes[2], exc[2]], 'amarelo_excecao': amarelo(),
                   'estado_depois_de_gravar': estado_man()}
            passo1 = (r_lib.startswith('OK|') and r_g.startswith('OK|1|') and perto(exc[3], dp_bula)
                      and exc[2] == antes[2] and marcas(exc[29]) == (esp_exc_tm, 'EXCECAO') and amarelo()
                      and estado_man() == 'TRANCADA')
            # Aplicar modo 3 no periodo inteiro
            ora = apurar(rows, lote, AN, ds[0], ds[-1], nlv)
            resp('3|SIM|SIM')
            r_3 = txt(run('AplicarAlvos'))
            resp('')
            dep = linha()
            c3 = cond_n(ora[1][0], ora[1][2], rec, ideal)
            tm3, td3 = marcas(dep[29])
            w3 = U(dep[22])
            esperado_amarelo = ('TEMPORARIO' in w3 or 'EXCECAO TEMPORARIA' in w3
                                or any(any(marcas(dep[29 + 5 * k])) for k in range(nlv)))
            esp_c = None
            if c3 not in ('BLOQUEADO', 'DP_ZERO'):
                td_esp = 'TEMPORARIO' if c3 == 'TEMPORARIO' else ''
                esp_c = (esp_exc_tm, td_esp)
            ev6.update({'aplicar_modo3': r_3, 'cond_N1_depois': txt(dep[29]), 'cond_do_N': c3,
                        'dp_depois': [dep[3], ora[1][2]], 'media_depois': [dep[2], exc[2]], 'W': txt(dep[22])[:260],
                        'amarelo': amarelo(), 'amarelo_esperado': esperado_amarelo})
            passo2 = (r_3.startswith('OK|') and r_3.endswith('|DP') and esp_c is not None
                      and (tm3, td3) == esp_c and perto(dep[3], ora[1][2]) and dep[2] == exc[2]
                      and 'DP (EXCECAO TEMPORARIA)' not in w3 and amarelo() == esperado_amarelo
                      and (esp_exc_tm or c3 == 'TEMPORARIO' or propria(dep[29]) == c3))
            ok('X06', passo1 and passo2,
               'excecao SO do DP (bula) e depois "Aplicar" modo 3: a marca EXCECAO do DP sai, a Media fica, condicao '
               'e amarelo coerentes', ev6)
        except Exception as e:                # noqa: BLE001
            reg('X06', 'REPROVADO', 'excecao so DP -> modo 3: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X07 troca de analito tranca
        try:
            outro_an = next(x for x in catalogo if x and U(x) != AN)
            resp('SIM')
            r_l = txt(run('LiberarManualCore'))
            resp('')
            lib = estado_man()
            escrever(p.Range('C3'), outro_an, eventos=True)
            depois_c3 = estado_man()
            trav = bool(p.Range(f'B{R_MAN1}').Locked) and not txt(p.Range(f'Z{R_MAN_TIT}').Value)
            escrever(p.Range('C3'), analito, eventos=True)
            if U(q.nome('selAnalito').Value) != AN:
                escrever(p.Range('B3'), cat_u.index(AN) + 1, eventos=False)
                run('mEstatistica.PainelMudou', teto=900)
            # spinner: B3 muda SEM Worksheet_Change; o gravar tem de recusar e trancar
            resp('SIM')
            r_l2 = txt(run('LiberarManualCore'))
            resp('')
            antes = linha()
            escrever(p.Range('B3'), cat_u.index(U(outro_an)) + 1, eventos=False)
            xl.Calculate()
            escrever(p.Range(f'C{R_MAN1}'), 1.2345, eventos=False)
            escrever(p.Range(f'D{R_MAN1}'), 'Bula', eventos=False)
            escrever(p.Range(f'B{R_MAN_JUST}'), f'{TAG}: tentativa com o analito trocado pelo spinner.', eventos=False)
            resp('SIM|SIM')
            r_g2 = txt(run('GravarManualCore'))
            resp('')
            sel_spin = txt(q.nome('selAnalito').Value)
            ok_spin = r_g2.startswith('ERRO|') and linha() == antes and estado_man() == 'TRANCADA'
            escrever(p.Range('B3'), cat_u.index(AN) + 1, eventos=False)
            run('mEstatistica.PainelMudou', teto=900)
            ok('X07', r_l.startswith('OK|') and lib.startswith('LIBERADA') and depois_c3 == 'TRANCADA' and trav
               and r_l2.startswith('OK|') and ok_spin,
               'trocar o analito (C3) tranca de novo; analito trocado pelo spinner (sem evento): gravar recusado, '
               'nada gravado, trancado', {'liberar': r_l, 'antes': lib, 'depois_C3': depois_c3, 'trancado': trav,
                                          'spinner_selAnalito': sel_spin, 'gravar_spinner': r_g2[:220]})
        except Exception as e:                # noqa: BLE001
            reg('X07', 'REPROVADO', 'troca de analito: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X08 referencia: igual ao limite = dentro
        try:
            if U(q.nome('selAnalito').Value) != AN:
                escrever(p.Range('B3'), cat_u.index(AN) + 1, eventos=False)
                run('mEstatistica.PainelMudou', teto=900)
            secao(lote, None, None)
            at1 = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and nivel(r) == 1]
            vals = sorted(float(r['RESULTADO']) for r in at1)
            k = max(1, len(vals) // 6)
            li = next(v for v in vals[k:] if v > vals[0])
            lsup = next(v for v in reversed(vals[:-k]) if v < vals[-1])
            rg = next((R_REF1 + i for i, v in enumerate(p.Range(f'A{R_REF1}:A{R_REF1 + N_REF - 1}').Value)
                       if U(v[0]) == AN), None)
            if rg is None:
                raise RuntimeError('analito fora da grade da referencia')
            foto_a, wg_a = foto_ls(), eventos_wg()
            st_a = [txt(p.Range(f'J{6 + i}').Value) for i in range(1, nlv + 1)]
            escrever(p.Range(f'B{rg}:D{rg}'), ((round((li + lsup) / 2, 6), li, lsup),), eventos=True)
            xl.Calculate()
            o = [l for l in obs() if U(l[2]) == LT and U(l[3]) == AN and txt(l[4]) == '1' and perto(l[7], li)
                 and perto(l[8], lsup)]
            pend = [l for l in o if U(l[10]) == 'PENDENTE']
            abx = sum(1 for x in vals if x < li)
            acm = sum(1 for x in vals if x > lsup)
            iguais = {txt(r['ID_REGISTRO']) for r in at1 if float(r['RESULTADO']) in (li, lsup)}
            com_obs_igual = [txt(l[1]) for l in o if txt(l[1]) in iguais]
            sit = collections.Counter(txt(l[9]) for l in pend)
            foto_b, wg_b = foto_ls(), eventos_wg()
            st_b = [txt(p.Range(f'J{6 + i}').Value) for i in range(1, nlv + 1)]
            ok('X08', sit.get('ABAIXO DO LI', 0) == abx and sit.get('ACIMA DO LS', 0) == acm and iguais
               and not com_obs_igual and foto_a == foto_b and wg_a == wg_b and st_a == st_b,
               'referencia: resultado IGUAL ao LI/LS fica dentro (sem observacao); abaixo/acima = PENDENTE (= oraculo); '
               'LotesStore A1:AO, Eventos_Westgard e status do LJ inalterados pela grade',
               {'LI': li, 'LS': lsup, 'oraculo_abaixo_acima': [abx, acm], 'pendentes': dict(sit),
                'IDs_iguais_ao_limite': sorted(iguais)[:10], 'com_obs_indevida': com_obs_igual[:10],
                'LotesStore_igual': foto_a == foto_b, 'Westgard_igual': wg_a == wg_b, 'status_LJ': [st_a, st_b]})
        except Exception as e:                # noqa: BLE001
            reg('X08', 'REPROVADO', 'referencia/limite: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})
            rg = None

        # ============================================================ X09 PENDENTE inativada -> RETIRADA
        try:
            pend = [l for l in obs() if U(l[2]) == LT and U(l[3]) == AN and U(l[10]) == 'PENDENTE' and txt(l[1])]
            if not pend:
                reg('X09', 'NAO_EXECUTADO', 'nenhuma observacao PENDENTE do lote para inativar', {})
            else:
                alvo = pend[0]
                pid = txt(alvo[1])
                q.escrever_linha('tblInativacao_NaoConformes', {'ID_REGISTRO': pid, 'ANALITO': txt(alvo[3]),
                                                                'MOTIVO': f'Outro: {TAG} (repro X09, so na copia)'})
                r_at, t_at = atualizar_entradas()
                rows = ler_cq()
                st_ = [U(l[10]) for l in obs() if U(l[0]) == U(alvo[0])]
                ok('X09', r_at.startswith('OK|') and st_ == ['RETIRADA'],
                   'observacao PENDENTE cujo resultado foi inativado vira RETIRADA no ATUALIZAR',
                   {'ID': pid, 'chave': txt(alvo[0]), 'status_depois': st_, 'atualizar': [r_at[:80], t_at]})
        except Exception as e:                # noqa: BLE001
            reg('X09', 'REPROVADO', 'RETIRADA por inativacao: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ X10 inativado fora da faixa nao gera observacao
        try:
            if inat_id is None or rg is None:
                reg('X10', 'NAO_EXECUTADO', 'sem o resultado inativado do X01 ou sem a linha da grade', {})
            else:
                at1 = [r for r in rows if ativo(r) and U(r['LOTE']) == LT and U(r['ANALITO']) == AN and nivel(r) == 1]
                vals = [float(r['RESULTADO']) for r in at1]
                li = inat_val + max(1e-6, abs(inat_val) * 1e-6)
                lsup = max(vals + [inat_val]) + 1.0
                foto_a = foto_ls()
                escrever(p.Range(f'C{rg}:D{rg}'), ((li, lsup),), eventos=True)
                o = [l for l in obs() if U(l[2]) == LT and U(l[3]) == AN and txt(l[4]) == '1' and perto(l[7], li)
                     and perto(l[8], lsup)]
                do_inat = [l for l in obs() if txt(l[1]) == inat_id and U(l[3]) == AN and U(l[10]) == 'PENDENTE']
                abx = sum(1 for x in vals if x < li)
                ok('X10', not do_inat and len([l for l in o if U(l[10]) == 'PENDENTE']) == abx and foto_a == foto_ls(),
                   'resultado INATIVADO abaixo do LI nao gera observacao; ativos abaixo = PENDENTE (= oraculo); '
                   'LotesStore inalterada', {'ID_inativado': inat_id, 'valor': inat_val, 'LI': li, 'LS': lsup,
                                             'pendentes_do_inativado': len(do_inat), 'oraculo_abaixo': abx,
                                             'pendentes': len([l for l in o if U(l[10]) == 'PENDENTE'])})
        except Exception as e:                # noqa: BLE001
            reg('X10', 'REPROVADO', 'inativado fora da faixa: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})
    except Exception as e:                    # noqa: BLE001
        reg('Z99', 'REPROVADO', 'execucao do repro sem excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        if q is not None:
            q.fechar(salvar=False)           # fecha so o Excel criado aqui (xlh: Quit; taskkill so do proprio PID)
    h1 = sha(origem)
    ok('Y00', h0 == h1, 'arquivo de origem (producao) intacto: SHA-256 igual antes e depois', {'antes': h0, 'depois': h1})
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump({'produto': produto, 'copia': copia, 'origem_sha': h0, 'resultados': RES}, f, ensure_ascii=False,
                  indent=1, default=str)
    cont = collections.Counter(r['estado'] for r in RES)
    print(f'\n=== repro_e48_alvos {produto}: ' + ', '.join(f'{k} {v}' for k, v in sorted(cont.items())) + ' ===',
          flush=True)
    return cont


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 2 or a[0] not in ('Bioquimica', 'Hematologia'):
        print(__doc__)
        sys.exit(2)
    prod, out = a[0], os.path.abspath(a[1])
    if '--origem' in a:
        org = os.path.abspath(a[a.index('--origem') + 1])
    else:
        org = os.path.join(os.environ.get('QC_RAIZ', RAIZ), f'QC_{prod}.xlsm')
    c = executar(prod, org, out)
    sys.exit(1 if c.get('REPROVADO') else 0)
