# -*- coding: utf-8 -*-
"""qa_fluxo_manual.py -- prova de ponta a ponta de que um resultado MANUAL digitado num dia qualquer se integra aos do
interfaceamento: mesma corrida, mesmo grafico, mesma estatistica, mesmo Westgard, mesma trilha.

Uso: python qa_fluxo_manual.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Pelo caminho real: escreve na aba Digitar Resultados com os eventos ligados (o Worksheet_Change gera o MAN_ e carimba) e
roda mIntegracao.AtualizarSoEntradas (o ATUALIZAR sem ler o DB_SEAC). MODO_FONTE = HISTORICO durante o teste (nada novo
chega no meio; vale tambem em casa, sem DB_SEAC). Tudo e lido da copia em tempo de execucao: analito, lote, equipamento,
corridas, alvo (Media/DP), filtro da Estatistica. Nada salvo.

  F01 escolha, sobre a base: analito do lote EM USO com corrida automatica C1 (mesmo RUN) que tem o N1 ATIVO uma unica
      vez e NAO tem o nivel t; sem automatico do nivel t a ate TOLERANCIA_CONFLITO_MIN do instante do manual (dia-1, dia,
      dia+1); o item anterior a ate GAP_CORRIDA_MIN e o seguinte (de outro bloco) a mais; RUN na janela do Eng_Saida.
      Tambem: C2 (F04), o dia D (F05) e o ultimo dia (F06). Sem C1: FAIL 'pre-condicao ausente'
  F02 o nivel t digitado 2 min depois do N1 de C1:
      a) ATIVO/SIM/NORMAL/MANUAL, MAN_ novo, USUARIO_MANUAL e RECEBIDO_EM preenchidos
      b) RUN e CORRIDA_NO_DIA do N1 de C1 (= oraculo qa_etl.oraculo_run); nenhum outro resultado muda
      c) Eng_Saida: a MESMA linha do RUN tem o N1 automatico e o manual no nivel t (valor e ID)
      d) corrida do lote = oraculo qa_corrida_lote; o nivel t ganha +1 dali em diante; o N1 nao muda
      e) QA: nenhum A01/E04/E12 do manual; A06 do manual se e somente se houver automatico do nivel t no dia
      f) Audit_Log: exatamente 1 linha RESULTADO_*: RESULTADO_MANUAL_INCLUIDO, 'ID MAN_x', RUN de C1, quem lancou
  F03 a) Estatistica: n+1, media/DP = calculo independente (amostral, n-1) da Principal com o filtro do produto e a
      media fechada (n*m + x)/(n+1); b) engPainel do nivel t = n/media/DP dos pontos do Eng_Saida com filtro = 1, n+1
  F04 R_4s entre automatico e manual no MESMO RUN: C2 com N1 em 1,2 <= |z| <= 2,0; o nivel t digitado do lado oposto
      (|z| < 3, sem 1_3s) com |z1 - zt| > 4 -> RR4s = 1 nos dois niveis, veredito REJEITADO, evento R_4s no RUN e
      WESTGARD_EVENTO_* na trilha
  F05 manual N1 numa corrida NOVA antes das automaticas do dia D (com um N1 automatico posterior inativado antes):
      a) as corridas do _GK/dia D andam +1 (RUN e CORRIDA_NO_DIA); nada fora dele muda; o manual e a corrida 1
      b) corrida do lote = oraculo em todos os niveis, sem lacuna, X entre as corridas certas
      c) DadosDoPonto do manual (ID, analito, nivel, corrida do lote, data/hora, valor, ATIVO, lote)
      d) o inativado continua INATIVADO/X_VERMELHO, com o RUN deslocado, e o X aparece nesse RUN
      e) Liberacao: um marcador gravado em C/D ao lado de um RUN continua ao lado do MESMO RUN depois de um manual num
         dia sem nenhum ATIVO do lote (RUN novo no meio da lista). ESPERADO HOJE: REPROVAR (as assinaturas ficam
         amarradas a POSICAO, mLotes.AtualizarListaLiberacao) -- risco de desalinhamento da Liberacao
      f) INFO: trilha de Westgard do deslocamento (REMOVIDO + NOVO do mesmo evento so por troca de RUN = achado)
  F06 manual N1 no ultimo dia, depois da ultima corrida, em alvo + 3,5 DP: 1_3s, veredito REJEITADO, engPainel e J7
      REJEITADO; corrigir o RESULTADO na propria linha para o alvo: OK no RUN, os outros vereditos voltam ao de antes,
      RESULTADO_MANUAL_ALTERADO (antes -> depois), USUARIO e REGISTRADO_EM da linha intactos
  F07 a) idempotencia: AtualizarSoEntradas 2x + ATUALIZAR completo -> Principal identica e nenhuma linha RESULTADO_*
      b) retirada: apagar as linhas (um ClearContents, como o Delete), a inativacao e o comentario -> RUN de todos os
         IDs, Principal, corrida do lote, engPainel, Estatistica, eventos de Westgard e Liberacao = estado inicial;
         1 RESULTADO_MANUAL_RETIRADO por manual
  F00 seguranca do teste: arquivo de entrada intacto (SHA-256 no inicio e no fim)
Sai com 1 se algum falhar.
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..'))
from qa_final import QA, reg, RES, NAO, stats, sha256_arquivo  # noqa: E402
from qa_inativar_ux import Leitor, curto, num_pt, ativos_em_ordem, consecutivos, x_do_id  # noqa: E402
import qa_corrida_lote as cl  # noqa: E402
import qa_manual_substitui as ms  # noqa: E402
import qa_westgard_sequencia as wsq  # noqa: E402
import qa_etl as etl  # noqa: E402

SIM = 'SIM'
SENHA = 'qcini2025'
TB_MAN = 'tblResultados_Manuais'
TB_INAT = 'tblInativacao_NaoConformes'
TB_COM = 'tblComentariosTecnicos'
TB_FIN = 'tblCQ_Final'
TB_QA = 'tblQA_Integracao'
ABA_LIB = 'Liberação'
MOT = 'Falha do interfaceamento'          # motivo real da rotina (lista da aba Digitar Resultados)
MOT_INAT = 'QA fluxo manual'
EPOCA = dt.datetime(1899, 12, 30)
NEF = 7                                    # campos por nivel no Eng_Saida (1=R13s, 3=RR4s, 7=veredito)
ENG_ALVO0 = 75                             # Eng_Saida BW..CB: (Media, DP) de N1, N2, N3 (mEstatistica.ENG_COL_ALVO0)
MARGEM_JANELA = 5                          # corridas: os passos acrescentam ate 3 corridas e a janela (NK) corta as antigas
G = lambda m: dt.timedelta(minutes=m)      # noqa: E731

txt, iint, perto = ms.txt, ms.iint, ms.perto
up, eh_num = etl.up, etl.eh_num


# ---------------------------------------------------------------- utilidades puras (sem COM)
def minuto(d):
    return d.replace(second=0, microsecond=0)


def serial_dia(d):
    return float((dt.datetime.combine(d, dt.time()) - EPOCA).days)


def perto_rel(a, b, tol=1e-9):
    if not eh_num(a) or not eh_num(b):
        return False
    return abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


def mediana(vals):
    v = sorted(vals)
    return v[len(v) // 2] if v else None


def eh_auto(r):
    return txt(r.get('ORIGEM_RESULTADO')).upper() == 'INTERFACEAMENTO'


def gk(r):
    """_GK do motor (sem nivel): equipamento, matriz, lote, analito, dia da DATA."""
    return (up(r['EQUIPAMENTO']), up(r['MATRIZ']), txt(r['LOTE']), up(r['ANALITO']), r['DATA'].date())


def real(r):
    """Ocupa corrida (mesma regra do oraculo do qa_etl): qualquer status de CORRIDA_REAL, com DATA_HORA, nao alvo de
    um manual que substitui."""
    return (r.get('STATUS_ANALITICO') in etl.CORRIDA_REAL and not etl.alvo_de_manual(r)
            and isinstance(r.get('DATA_HORA'), dt.datetime) and isinstance(r.get('DATA'), dt.datetime))


def ordem(r):
    return (r['DATA_HORA'], iint(r.get('NIVEL')) or 0, txt(r.get('ITEM_ID')), txt(r['ID_REGISTRO']))


def no_escopo(r, an, lote, eqf):
    return (up(r['ANALITO']) == up(an) and txt(r['LOTE']) == lote
            and (not eqf or up(r['EQUIPAMENTO']) == eqf))


def indice_conflito(fin_lista, an, lote):
    """(EQ, MZ, LOTE, AN, NIVEL) -> [(DATA_HORA, ID)] de todo resultado com valor, nao retransmissao, de qualquer
    status e origem -- mais estrito que o motor (que olha os automaticos e, entre manuais, o mesmo instante)."""
    idx = collections.defaultdict(list)
    for r in fin_lista:
        if up(r['ANALITO']) != up(an) or txt(r['LOTE']) != lote:
            continue
        if r['STATUS_ANALITICO'] == 'DUPLICIDADE_ORIGEM' or not isinstance(r.get('DATA_HORA'), dt.datetime) \
                or not eh_num(r['RESULTADO']):
            continue
        idx[(up(r['EQUIPAMENTO']), up(r['MATRIZ']), txt(r['LOTE']), up(r['ANALITO']), iint(r['NIVEL']))].append(
            (r['DATA_HORA'], r['ID_REGISTRO']))
    return idx


def conflito(idx, chave4, nivel, quando, tol):
    for dh, i in idx.get(tuple(chave4) + (nivel,), ()):
        if abs((dh - quando).total_seconds()) / 60.0 <= tol + 1:
            return i
    return None


def corridas_candidatas(reais_esc, idx, runs_jan, gap, tol, nlv, excluir):
    """Corridas automaticas (um RUN = um bloco) com o N1 ATIVO uma vez e algum nivel t faltando, onde o manual do nivel
    t, 2 min depois do N1 (ao minuto), entra no MESMO bloco sem unir com o seguinte e sem conflito."""
    pos_jan = {r: i for i, r in enumerate(runs_jan)}
    por_gk = collections.defaultdict(list)
    for r in reais_esc:
        por_gk[gk(r)].append(r)
    out = []
    for k, rows in por_gk.items():
        if k[4] in excluir:
            continue
        rows = sorted(rows, key=ordem)
        for b in etl.blocos_do_grupo(rows, gap):
            runs_b = {iint(r['RUN']) for r in b}
            if len(runs_b) != 1:
                continue
            run = next(iter(runs_b))
            if run is None or pos_jan.get(run, -1) < MARGEM_JANELA:
                continue
            if not all(eh_auto(r) for r in b):
                continue
            n1 = [r for r in b if iint(r['NIVEL']) == 1]
            if len(n1) != 1 or n1[0]['STATUS_ANALITICO'] != 'ATIVO' or not eh_num(n1[0]['RESULTADO']):
                continue
            n1 = n1[0]
            presentes = {iint(r['NIVEL']) for r in b}
            faltam = [t for t in range(2, nlv + 1) if t not in presentes]
            if not faltam:
                continue
            dh_m = minuto(n1['DATA_HORA'] + G(2))
            if dh_m.date() != k[4]:
                continue
            ids_b = {r['ID_REGISTRO'] for r in b}
            ant = [r for r in rows if r['DATA_HORA'] <= dh_m]
            seg = [r for r in rows if r['DATA_HORA'] > dh_m and r['ID_REGISTRO'] not in ids_b]
            if not ant or ant[-1]['ID_REGISTRO'] not in ids_b:
                continue
            if (dh_m - ant[-1]['DATA_HORA']).total_seconds() / 60.0 > gap - 0.5:
                continue
            if seg and (seg[0]['DATA_HORA'] - dh_m).total_seconds() / 60.0 <= gap + 0.5:
                continue
            ts = [t for t in faltam if not conflito(idx, k[:4], t, dh_m, tol)]
            if not ts:
                continue
            out.append({'run': run, 'gk': k, 'dia': k[4], 'n1': n1, 'niveis_t': ts, 'dh_m': dh_m,
                        'pos_janela': pos_jan[run], 'bloco': [r['ID_REGISTRO'] for r in b]})
    out.sort(key=lambda c: -c['run'])
    return out


def dias_f05(reais_esc, idx, runs_jan, gap, tol, excluir):
    """_GK/dia com >= 2 corridas, todas na janela, espaco para um manual N1 antes da 1a (fora do GAP e da TOL) e um N1
    automatico ATIVO numa corrida posterior (o que sera inativado)."""
    pos_jan = {r: i for i, r in enumerate(runs_jan)}
    por_gk = collections.defaultdict(list)
    for r in reais_esc:
        por_gk[gk(r)].append(r)
    out = []
    for k, rows in por_gk.items():
        if k[4] in excluir:
            continue
        rows = sorted(rows, key=ordem)
        blocos = etl.blocos_do_grupo(rows, gap)
        if len(blocos) < 2:
            continue
        runs_dia = {iint(r['RUN']) for r in rows}
        if None in runs_dia or any(pos_jan.get(rn, -1) < MARGEM_JANELA for rn in runs_dia):
            continue
        m = minuto(rows[0]['DATA_HORA'] - G(max(gap, tol) + 5))
        if m < etl.meia_noite(rows[0]['DATA_HORA']) + G(1):
            continue
        if conflito(idx, k[:4], 1, m, tol):
            continue
        xs = [r for b in blocos[1:] for r in b if iint(r['NIVEL']) == 1 and r['STATUS_ANALITICO'] == 'ATIVO'
              and eh_auto(r) and eh_num(r['RESULTADO'])]
        if not xs:
            continue
        out.append({'gk': k, 'dia': k[4], 'm': m, 'x': xs[-1], 'n_corridas': len(blocos),
                    'equip': rows[0]['EQUIPAMENTO'], 'matriz': rows[0].get('MATRIZ'), 'analito': rows[0]['ANALITO']})
    out.sort(key=lambda c: c['dia'], reverse=True)
    return out


def plano_f06(reais_esc, idx, gap, tol):
    """O ultimo dia do escopo: manual N1 depois da ultima corrida (> GAP do ultimo item e > TOL do ultimo N1)."""
    if not reais_esc:
        return None
    ult_dia = max(r['DATA'].date() for r in reais_esc)
    do_dia = [r for r in reais_esc if r['DATA'].date() == ult_dia]
    k = gk(max(do_dia, key=ordem))
    rows = sorted((r for r in do_dia if gk(r) == k), key=ordem)
    n1s = [r for r in rows if iint(r['NIVEL']) == 1]
    lim = rows[-1]['DATA_HORA'] + G(gap + 2)
    if n1s:
        lim = max(lim, n1s[-1]['DATA_HORA'] + G(tol + 2))
    m6 = etl._minuto_acima(lim)
    if m6.date() != ult_dia or conflito(idx, k[:4], 1, m6, tol):
        return None
    return {'gk': k, 'dia': ult_dia, 'm': m6, 'equip': rows[-1]['EQUIPAMENTO'], 'matriz': rows[-1].get('MATRIZ'),
            'analito': rows[-1]['ANALITO']}


def nucleo_lote(codigo):
    """Espelho de mEntrada.NucleoLote: 'QC-' + nucleo + 2 digitos de nivel -> nucleo; nucleo devolve ele mesmo."""
    s = txt(codigo)
    return s[3:-2] if s.upper().startswith('QC-') and len(s) > 5 else s


def run_de_dia(d, corrida):
    return int(d.strftime('%y%m%d')) * 100 + corrida


def dia_de_run(run):
    return dt.date(2000 + run // 1000000, (run // 10000) % 100, (run // 100) % 100)


def acoes(x):
    return [a.strip() for a in str(x.get('acao') or '').split('+')]


def so_resultado(aud):
    return [x for x in aud if any(a.startswith('RESULTADO_') for a in acoes(x))]


def manual_dict(quando, equip, lote, nivel, analito, valor, matriz=None):
    """A linha como o usuario digita (DATA, HORA ao minuto, ...); o ID MAN_ nasce no 1o campo e o carimbo tambem.
    MATRIZ = a do automatico da mesma chave (o usuario escolhe na lista): sem ela o PrepararLinha poe a MATRIZ_PADRAO,
    que pode nao ser a do lote, e o manual cairia em outro _GK (outro RUN) -- falso FAIL do F02/F05/F06."""
    v = {'DATA': dt.datetime.combine(quando.date(), dt.time()), 'HORA': (quando.hour * 60 + quando.minute) / 1440.0,
         'EQUIPAMENTO': equip}
    if txt(matriz):
        v['MATRIZ'] = txt(matriz)
    v.update({'LOTE': str(lote), 'NIVEL': int(nivel), 'ANALITO': analito, 'RESULTADO': float(valor), 'MOTIVO': MOT})
    return v


def como_num(v):
    """Numero de uma celula de configuracao (numero ou texto '10' / '10,5'); None se nao for numero."""
    if eh_num(v):
        return float(v)
    try:
        return float(str(v).strip().replace(',', '.'))
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- utilidades com COM
def apagar_linha(q, tabela, r):
    """Corpo de qa_etl.Ctx.apagar_linha: apaga a linha como o usuario (seleciona e Delete) -- UM ClearContents das
    celulas sem formula. Celula a celula, com o evento ligado, a aba manual geraria um MAN_ novo no meio da limpeza."""
    lo = q.lo(tabela)
    alvo = None
    for c in lo.ListColumns:
        cel = c.DataBodyRange.Cells(r, 1)
        if not cel.HasFormula:
            alvo = cel if alvo is None else q.ex.xl.Union(alvo, cel)
    if alvo is not None:
        ws = lo.Parent
        prot = bool(ws.ProtectContents)
        if prot:
            ws.Unprotect(SENHA)
        try:
            alvo.ClearContents()
        finally:
            if prot:
                ws.Protect(SENHA, False, True, True, True)


def valor_nome(q, n):
    try:
        v = q.wb.Names(n).RefersToRange.Value
    except Exception:                                     # noqa: BLE001
        v = q.wb.Application.Evaluate(n)
    if hasattr(v, 'year'):
        v = (v.replace(tzinfo=None) - EPOCA).days
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def filtro_estat(q):
    """O filtro da aba Estatistica (como o T17/T18 do qa_final): Bio = lote, periodo efetivo e equipamento do Painel;
    Hema = lote e anos B3/D3."""
    lote_est = txt(q.nome('Estat_Lote').Value)
    if q.bio:
        ini, fim = valor_nome(q, 'Estat_Ini_Efetiva'), valor_nome(q, 'Estat_Fim_Efetiva')
        equip = txt(q.ws('Painel').Range('L4').Value)

        def f(r):
            d = r['DATA']
            ds = (d.replace(tzinfo=None) - EPOCA).days if hasattr(d, 'year') else 0
            return (txt(r['LOTE']) == lote_est and (ini == 0 or ds >= ini) and (fim == 0 or ds < int(fim) + 1)
                    and (not equip or up(r['EQUIPAMENTO']) == up(equip)))
        return f, {'lote': lote_est, 'ini': ini, 'fim': fim, 'equip': equip}
    est = q.ws('Estatística')
    ano_de = iint(est.Range('B3').Value) or 0
    ano_ate = iint(est.Range('D3').Value) or ano_de

    def f(r):
        d = r['DATA']
        return txt(r['LOTE']) == lote_est and hasattr(d, 'year') and (ano_de == 0 or ano_de <= d.year <= ano_ate)
    return f, {'lote': lote_est, 'ano_de': ano_de, 'ano_ate': ano_ate}


def alvos_eng(q):
    """Eng_Saida BW..CB por linha: [M1, DP1, M2, DP2, M3, DP3]."""
    e = q.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    if n == 0:
        return []
    v = e.Range(e.Cells(3, ENG_ALVO0), e.Cells(2 + n, ENG_ALVO0 + 5)).Value
    return [list(x) for x in v] if isinstance(v[0], tuple) else [list(v)]


def alvo_do_nivel(alv, t):
    for linha in alv:
        m_, s_ = linha[2 * (t - 1)], linha[2 * (t - 1) + 1]
        if eh_num(m_) and eh_num(s_) and float(s_) > 0:
            return float(m_), float(s_)
    return None


def eng_painel(q):
    return [list(x) for x in q.nome('engPainel').Value]


def veredito_por_run(e, t):
    """RUN -> veredito (campo 7) do nivel t nas linhas com filtro do periodo = 1."""
    out = {}
    for i, rn in enumerate(e['runs']):
        if e['fil'][i] == 1:
            out[rn] = txt(e['flags'][i][(t - 1) * NEF + 6])
    return out


def serie_ponto(q, nivel):
    """Indice da serie de pontos do nivel no grafico do Painel ('Resultado', 'OK' ou 'Viola*'), como DadosDoPonto aceita."""
    nomes = [str(s.Name) for s in q.ws('Painel').ChartObjects(nivel).Chart.SeriesCollection()]
    for k, n in enumerate(nomes, 1):
        if n == 'Resultado':
            return k, nomes
    for k, n in enumerate(nomes, 1):
        if n == 'OK' or n.lower().startswith('viola'):
            return k, nomes
    return None, nomes


def lib_ler(q):
    return [list(r) for r in q.ws(ABA_LIB).Range('A4:F203').Value]


def lib_norm(lib):
    return [[wsq.norm_cel(x) for x in r] for r in lib]


def lib_runs(lib):
    return [iint(r[0]) for r in lib if iint(r[0])]


def lib_gravar(q, linha0, c, d):
    """Grava C/D de uma linha da Liberacao (marcador do teste), eventos desligados, aba desprotegida so para isso
    (senha da protecao da aba -- nenhuma credencial de usuario)."""
    ws = q.ws(ABA_LIB)
    prot = bool(ws.ProtectContents)
    ev = q.ex.xl.EnableEvents
    q.ex.xl.EnableEvents = False
    try:
        if prot:
            ws.Unprotect(SENHA)
        for col, val in ((3, c), (4, d)):
            cel = ws.Cells(4 + linha0, col)
            if val in (None, ''):
                cel.ClearContents()
            else:
                cel.Value = val
    finally:
        if prot:
            ws.Protect(SENHA, False, True, True, True)
        q.ex.xl.EnableEvents = ev


def estado_resultado(r):
    if r is None:
        return None
    return (iint(r['RUN']), iint(r['CORRIDA_NO_DIA']), r['STATUS_ANALITICO'], r['PARTICIPA_ESTATISTICA'],
            r['TIPO_PLOTAGEM_LJ'], r['RESULTADO'])


# ---------------------------------------------------------------- a suite
def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    nlv = q.nlv
    tempos = {}
    ctx = {'an': ''}
    manuais = []                        # (rotulo, linha da tabela, ID)
    extras = {}                         # linhas da inativacao/comentario do F05, marcador da Liberacao

    def run(macro, *a, teto=300):
        return q.run(macro, *a, teto=teto)

    def so_entradas(rot):
        t0 = time.time()
        r = str(run('mIntegracao.AtualizarSoEntradas', ctx['an'], teto=1500))
        tempos[rot] = round(time.time() - t0, 1)
        if not r.startswith('OK|'):
            raise RuntimeError(f'AtualizarSoEntradas ({rot}): {r[:400]}')
        return r

    def completo(rot):
        t0 = time.time()
        q.atualizar()
        tempos[rot] = round(time.time() - t0, 1)

    def escrever_manual(rot, valores):
        q.ex.xl.EnableEvents = True
        r = q.escrever_linha(TB_MAN, valores)
        lin = ms.linha_manual(q, r)
        mid = txt(lin.get('ID_REGISTRO'))
        manuais.append((rot, r, mid))
        return r, mid, lin

    try:
        # ------------------------------------------------------------ preparo: historico, motor VBA, lote em uso
        modo0 = q.cfg('MODO_FONTE', 'HISTORICO')
        motor0 = q.cfg('MOTOR_DADOS', 'VBA')
        completo('base_atualizar')
        gap0, tol0 = q.cfg('GAP_CORRIDA_MIN'), q.cfg('TOLERANCIA_CONFLITO_MIN')
        gap, tol = como_num(gap0), como_num(tol0)
        if gap is None or tol is None:
            raise RuntimeError(f'pré-condição ausente: GAP_CORRIDA_MIN={gap0!r} / TOLERANCIA_CONFLITO_MIN={tol0!r} '
                               'não numéricos na tblConfigIntegracao')
        lote_uso = txt(run('mDados.LoteAtivoCore'))
        try:
            lote_sel = txt(q.nome('loteSel').Value)
        except Exception:                             # noqa: BLE001 -- sem o nome: como H3 vazio (TODOS)
            lote_sel = ''
        lote_pai = txt(run('mLotes.LotePainel'))
        trocou_lote = False
        if lote_uso and (not lote_sel or not wsq.mesmo_lote(lote_pai, lote_uso)):
            # o Painel precisa estar no lote EM USO (a Liberacao so lista ele) e nunca em TODOS (H3 vazio: media/DP do
            # engPainel ficam vazios por desenho, ADR-086)
            pai = q.ws('Painel')
            try:
                pai.Unprotect(SENHA)
            except Exception:                         # noqa: BLE001
                pass
            q.ex.xl.EnableEvents = True
            pai.Range('H3').Value = "'" + lote_uso
            q.ex.esperar()
            q.ex.xl.Calculate()
            run('mEstatistica.RecalcularAnalitoAtual', teto=600)
            q.ex.xl.Calculate()
            trocou_lote = True
            lote_pai = txt(run('mLotes.LotePainel'))
        lote = lote_pai
        eqf = up(run('mEstatistica.EquipFiltro'))
        hoje = dt.date.today()
        fin_b = q.final_por_id()
        lista_b = list(fin_b.values())
        cadastro = [txt(v[0]) for v in q.espec().Range('A4:A43').Value]
        cad_up = [c.upper() for c in cadastro]

        # ------------------------------------------------------------ F01 escolha (analito, C1, C2, D, ultimo dia)
        cont = collections.Counter(up(r['ANALITO']) for r in lista_b if txt(r['LOTE']) == lote
                                   and r['PARTICIPA_ESTATISTICA'] == SIM and iint(r['NIVEL']) == 1
                                   and (not eqf or up(r['EQUIPAMENTO']) == eqf))
        ranking = [cadastro[cad_up.index(a)] for a, n in cont.most_common() if a in cad_up and n >= 10]

        def preparar_painel(an):
            esc = [r for r in lista_b if real(r) and no_escopo(r, an, lote, eqf)]
            if not esc:
                return None
            d_ini = min(r['DATA'].date() for r in esc)
            d_fim = max(r['DATA'].date() for r in esc)
            q.ex.xl.EnableEvents = False
            try:
                pai = q.ws('Painel')
                pai.Unprotect(SENHA)
                pai.Range('B3').Value = cad_up.index(up(an)) + 1
                # numero de serie: o pywin32 converte datetime ingenuo por fuso
                pai.Range('G3').Value = serial_dia(d_ini)
                pai.Range('G4').Value = serial_dia(d_fim)
            finally:
                q.ex.xl.EnableEvents = True
            run('mEstatistica.AtualizarEstatistica', teto=900)
            if q.bio:
                run('RecalcularEstatPeriodo', teto=900)
            return esc

        escolha, tentativas = None, []
        for an in ranking[:4]:
            esc = preparar_painel(an)
            if not esc:
                continue
            e_ = q.eng_saida()
            if up(e_['analito']) != up(an):
                tentativas.append({'analito': an, 'falha': f'Eng_Saida mostra {e_["analito"]!r}'})
                continue
            idx = indice_conflito(lista_b, an, lote)
            f06 = plano_f06(esc, idx, gap, tol)
            excl = {d for d in {r['DATA'].date() for r in esc} if d >= hoje}
            if f06:
                excl.add(f06['dia'])
            else:
                excl.add(max(r['DATA'].date() for r in esc))
            cands = corridas_candidatas(esc, idx, e_['runs'], gap, tol, nlv, excl)
            filt, _ = filtro_estat(q)
            cands.sort(key=lambda c: (not filt({'LOTE': lote, 'DATA': dt.datetime.combine(c['dia'], dt.time()),
                                                'EQUIPAMENTO': c['n1']['EQUIPAMENTO']}), -c['run']))
            alv = alvos_eng(q)
            m1 = run('mLotes.MediaDoLote', lote, an, 1)
            s1 = run('mLotes.DPDoLote', lote, an, 1)
            alvo1 = (float(m1), float(s1)) if eh_num(m1) and eh_num(s1) and float(s1) > 0 else alvo_do_nivel(alv, 1)
            score = (bool(cands), alvo1 is not None, bool(f06))
            tentativas.append({'analito': an, 'C1_candidatos': len(cands), 'alvo_N1': alvo1, 'F06': bool(f06),
                               'corridas_na_janela': e_['n']})
            if escolha is None or score > escolha['score']:
                escolha = {'an': an, 'esc': esc, 'idx': idx, 'f06': f06, 'excl': excl, 'cands': cands,
                           'alvo1': alvo1, 'score': score}
            if all(score):
                break
        if escolha is None or not escolha['cands']:
            reg('F01 Escolha: corrida automática C1 com o N1 ATIVO uma vez, sem o nível t, onde o manual do nível t entra '
                'no mesmo RUN (GAP/TOL/janela conferidos)', False,
                {'falha': 'pré-condição ausente: nenhum analito do lote em uso com corrida C1 que cumpra as condições',
                 'lote': lote, 'lote_em_uso': lote_uso, 'equip': eqf, 'GAP': gap, 'TOL': tol, 'tentativas': tentativas})
            raise RuntimeError('pré-condição ausente: sem C1 (F01)')
        an = escolha['an']
        ctx['an'] = an
        if up(q.eng_saida()['analito']) != up(an):
            preparar_painel(an)
        esc, idx, f06, cands = escolha['esc'], escolha['idx'], escolha['f06'], escolha['cands']
        so_entradas('base_painel')

        # ------------------------------------------------------------ estado inicial (o F07 volta a ele)
        fin0 = q.final_por_id()
        run0 = ms.runs(fin0)
        cq0 = ms.tabela_v2(q, TB_FIN)
        e0 = cl.eng(q, nlv)
        eng0 = q.eng_saida()
        p0 = q.painel_stats()
        est0 = {t: q.estat_motor(an, t) for t in range(1, nlv + 1)}
        of0 = {k for k in wsq.oficiais(wsq.eventos(q)) if k[0] == up(an)}
        lib0 = lib_ler(q)
        alv0 = alvos_eng(q)
        filt, filt_desc = filtro_estat(q)
        C1 = next((c for c in cands if c['run'] in eng0['runs']), None)

        cache_alvo = {}

        def alvo_t(t):
            """Alvo (Media, DP) do nivel no lote, lido da copia (mLotes); sem ele, o BW..CB do Eng_Saida."""
            if t not in cache_alvo:
                m_ = run('mLotes.MediaDoLote', lote, an, t)
                s_ = run('mLotes.DPDoLote', lote, an, t)
                par = (float(m_), float(s_)) if eh_num(m_) and eh_num(s_) and float(s_) > 0 else None
                cache_alvo[t] = par or alvo_do_nivel(alv0, t)
            return cache_alvo[t]

        # C2 (F04): outra corrida candidata, em outro dia, N1 com 1,2 <= |z| <= 2,0 pelo alvo da propria linha
        C2 = None
        if C1:
            melhor = None
            for c in cands:
                if c['dia'] == C1['dia'] or c['run'] not in eng0['runs']:
                    continue
                i = eng0['runs'].index(c['run'])
                v1 = eng0['val'][i][0]
                a_ = alv0[i] if i < len(alv0) else None
                if not eh_num(v1) or not a_ or not eh_num(a_[0]) or not eh_num(a_[1]) or float(a_[1]) <= 0:
                    continue
                z1 = (float(v1) - float(a_[0])) / float(a_[1])
                if not 1.2 <= abs(z1) <= 2.0:
                    continue
                t2 = next((t for t in c['niveis_t'] if alvo_t(t)), None)
                if t2 is None:
                    continue
                cand = dict(c, z1=z1, t=t2, alvo_n1=(float(a_[0]), float(a_[1])))
                if melhor is None or abs(abs(z1) - 1.8) < abs(abs(melhor['z1']) - 1.8):
                    melhor = cand
            C2 = melhor
        excl5 = set(escolha['excl']) | ({C1['dia']} if C1 else set()) | ({C2['dia']} if C2 else set())
        D5 = next(iter(dias_f05(esc, idx, eng0['runs'], gap, tol, excl5)), None)

        def resumo_c(c):
            if not c:
                return None
            return {'RUN': c['run'], 'dia': c['dia'], 'N1': c['n1']['ID_REGISTRO'], 'N1_hora': c['n1']['DATA_HORA'],
                    'manual_em': c['dh_m'], 'niveis_faltando_sem_conflito': c['niveis_t'], 'linha_no_eng': c['pos_janela'],
                    'bloco': c['bloco'], **({'z_N1': round(c['z1'], 3), 'nivel_t': c['t']} if 'z1' in c else {})}
        reg('F01 Escolha: corrida automática C1 com o N1 ATIVO uma vez, sem o nível t, onde o manual do nível t entra no '
            'mesmo RUN (sem automático do nível t a até TOL, anterior a até GAP, seguinte de outro bloco a mais de GAP, '
            'RUN na janela do Eng_Saida, fora de hoje e do último dia)',
            C1 is not None,
            {'analito': an, 'lote': lote, 'lote_em_uso': lote_uso, 'painel_trocado_para_o_lote_em_uso': trocou_lote,
             'equip': eqf, 'GAP': gap, 'TOL': tol, 'C1': resumo_c(C1), 'C2_F04': resumo_c(C2),
             'D_F05': D5 and {'dia': D5['dia'], 'manual_em': D5['m'], 'corridas': D5['n_corridas'],
                              'inativar': D5['x']['ID_REGISTRO']},
             'F06': f06 and {'dia': f06['dia'], 'manual_em': f06['m']}, 'alvo_N1': escolha['alvo1'],
             'filtro_estatistica': filt_desc, 'tentativas': tentativas,
             **({} if C1 else {'falha': 'pré-condição ausente: nenhuma C1 dentro da janela do Eng_Saida'})})
        if not C1:
            raise RuntimeError('pré-condição ausente: sem C1 (F01)')

        fin_atual = fin0

        # ============================================================ F02 nivel faltante na corrida automatica
        t = C1['niveis_t'][0]
        n1c = C1['n1']
        vals_t = [r['RESULTADO'] for r in fin0.values() if no_escopo(r, an, lote, eqf) and iint(r['NIVEL']) == t
                  and r['STATUS_ANALITICO'] == 'ATIVO' and eh_num(r['RESULTADO'])]
        v2 = mediana(vals_t)
        if v2 is None:
            raise RuntimeError(f'pré-condição ausente: nenhum ATIVO do nível {t} para o valor do manual (F02)')
        est_a = q.estat_motor(an, t)
        p_a = q.painel_stats()[t - 1]
        eng_a = q.eng_saida()
        cor_a = cl.eng(q, nlv)
        r_aud = ms.audit_ult(q)
        r2, M2, lin2 = escrever_manual('F02', manual_dict(C1['dh_m'], n1c['EQUIPAMENTO'], lote, t, n1c['ANALITO'], v2,
                                                          n1c.get('MATRIZ')))
        so_entradas('F02')
        fin1 = q.final_por_id()
        m2 = fin1.get(M2)
        ok2a = (m2 is not None and re.fullmatch(r'MAN_\d{4,}', M2) is not None and M2 not in fin0
                and m2['STATUS_ANALITICO'] == 'ATIVO' and m2['PARTICIPA_ESTATISTICA'] == SIM
                and m2['TIPO_PLOTAGEM_LJ'] == 'NORMAL' and txt(m2['ORIGEM_RESULTADO']).upper() == 'MANUAL'
                and bool(txt(m2.get('USUARIO_MANUAL'))) and m2.get('RECEBIDO_EM') not in (None, '')
                and iint(m2['NIVEL']) == t and perto(m2['RESULTADO'], v2))
        reg(f'F02a Nível {t} digitado na corrida automática: ATIVO/SIM/NORMAL/MANUAL, ID MAN_ novo, USUARIO_MANUAL e '
            'RECEBIDO_EM preenchidos', ok2a,
            {'linha_digitada': {k: lin2.get(k) for k in ('ID_REGISTRO', 'DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE',
                                                         'NIVEL', 'ANALITO', 'RESULTADO', 'USUARIO', 'REGISTRADO_EM')},
             'final': ms.resumo(m2), 'USUARIO_MANUAL': m2 and m2.get('USUARIO_MANUAL'),
             'RECEBIDO_EM': m2 and m2.get('RECEBIDO_EM')})

        n1_1 = fin1.get(n1c['ID_REGISTRO'])
        orc_run = etl.oraculo_run(list(fin1.values()), gap)
        par_m2 = (iint(m2['RUN']), iint(m2['CORRIDA_NO_DIA'])) if m2 else None
        par_n1 = (iint(n1_1['RUN']), iint(n1_1['CORRIDA_NO_DIA'])) if n1_1 else None
        div_orc = [(i, (iint(r['CORRIDA_NO_DIA']), iint(r['RUN'])), orc_run.get(i)) for i, r in fin1.items()
                   if real(r) and gk(r) == C1['gk'] and orc_run.get(i) != (iint(r['CORRIDA_NO_DIA']), iint(r['RUN']))]
        mudou2 = [i for i in fin0 if estado_resultado(fin0[i]) != estado_resultado(fin1.get(i))]
        reg('F02b O manual fica com o RUN e a CORRIDA_NO_DIA do N1 da corrida automática (= oráculo); nenhum outro '
            'resultado muda (RUN, corrida, status, participa, plotagem, valor)',
            m2 is not None and par_m2 == par_n1 == (C1['run'], iint(n1c['CORRIDA_NO_DIA'])) and not div_orc
            and not mudou2 and ms.runs({i: fin1[i] for i in fin0 if i in fin1}) == run0,
            {'manual': par_m2, 'N1': par_n1, 'C1': (C1['run'], iint(n1c['CORRIDA_NO_DIA'])),
             'divergencias_oraculo_no_dia': div_orc[:10], 'mudaram': mudou2[:10], 'n_mudaram': len(mudou2)})

        eng_d = q.eng_saida()
        L1 = Leitor(q).eng_por_run()
        i_c1 = eng_d['runs'].index(C1['run']) if C1['run'] in eng_d['runs'] else None
        lin_c1 = L1.get(C1['run'])
        ok2c = (i_c1 is not None and lin_c1 is not None and eng_d['runs'] == eng_a['runs']
                and perto(eng_d['val'][i_c1][0], n1c['RESULTADO']) and perto(eng_d['val'][i_c1][t - 1], v2)
                and lin_c1['id'][0] == n1c['ID_REGISTRO'] and lin_c1['id'][t - 1] == M2)
        reg(f'F02c Eng_Saida: a MESMA linha do RUN tem o N1 automático e o nível {t} manual (valor e ID); o eixo não ganha '
            'nem perde corrida', ok2c,
            {'RUN': C1['run'], 'linha': i_c1, 'valores': eng_d['val'][i_c1] if i_c1 is not None else None,
             'ids': lin_c1['id'] if lin_c1 else None, 'esperado': (n1c['ID_REGISTRO'], n1c['RESULTADO'], M2, v2),
             'eixo_igual': eng_d['runs'] == eng_a['runs']})

        cor_d = cl.eng(q, nlv)
        orc1 = cl.oraculo(list(fin1.values()), cor_d['analito'], cor_d['lote'], cor_d['equip'], nlv)
        dif2 = cl.comparar(cor_d, orc1, nlv)
        desl = []
        for i, rn in enumerate(cor_d['runs']):
            if rn not in cor_a['runs']:
                continue
            ant = cor_a['cor'][t - 1][cor_a['runs'].index(rn)]
            dep = cor_d['cor'][t - 1][i]
            if rn == C1['run']:
                continue
            if ant is not None and dep != ant + (1 if rn > C1['run'] else 0):
                desl.append((rn, ant, dep))
        reg(f'F02d Corrida do lote = oráculo independente em todos os níveis; o nível {t} ganha +1 dali em diante; o N1 '
            'não muda', i_c1 is not None and not dif2 and not desl and cor_d['cor'][0] == cor_a['cor'][0]
            and orc1[t - 1][0].get(C1['run']) is not None
            and cor_d['cor'][t - 1][i_c1] == orc1[t - 1][0].get(C1['run']),
            {'divergencias': dif2[:10], 'deslocamento_errado': desl[:10],
             'corrida_do_manual': cor_d['cor'][t - 1][i_c1] if i_c1 is not None else None,
             'N1_igual': cor_d['cor'][0] == cor_a['cor'][0]})

        qa1 = q.ler(TB_QA)
        ruins_qa = [(x['CODIGO'], x['DETALHE']) for x in qa1 if x['ID_REGISTRO'] == M2 and x['CODIGO'] in ('A01', 'E04', 'E12')]
        a06 = ms.achados(qa1, 'A06', M2)
        a06_esp = bool(m2) and any(
            eh_auto(r) and r['STATUS_ANALITICO'] in ('ATIVO', 'INATIVADO') and up(r['EQUIPAMENTO']) == up(m2['EQUIPAMENTO'])
            and up(r['MATRIZ']) == up(m2['MATRIZ']) and txt(r['LOTE']) == txt(m2['LOTE']) and iint(r['NIVEL']) == t
            and up(r['ANALITO']) == up(m2['ANALITO']) and hasattr(r.get('DATA'), 'year')
            and r['DATA'].date() == m2['DATA'].date() for r in fin1.values())
        reg('F02e QA: nenhum A01, E04 ou E12 do manual; A06 do manual se e somente se houver automático do mesmo nível no '
            'dia (mesmos equipamento/matriz/lote/analito)', not ruins_qa and bool(a06) == a06_esp,
            {'A01_E04_E12': ruins_qa, 'A06': a06, 'A06_esperado': a06_esp})

        aud2 = ms.audit_desde(q, r_aud)
        res2 = so_resultado(aud2)
        reg('F02f Audit_Log: exatamente 1 linha RESULTADO_*: RESULTADO_MANUAL_INCLUIDO, motivo "ID MAN_x", RUN da corrida, '
            'parecer com quem lançou',
            len(res2) == 1 and 'RESULTADO_MANUAL_INCLUIDO' in acoes(res2[0]) and res2[0]['motivo'] == f'ID {M2}'
            and iint(res2[0]['run']) == C1['run'] and 'manual lancado por' in res2[0]['parecer'],
            {'linhas_RESULTADO': res2, 'outras_acoes': sorted({a for x in aud2 for a in acoes(x)})})

        # ============================================================ F03 Estatistica e engPainel
        est_d = q.estat_motor(an, t)
        pop_a = [r['RESULTADO'] for r in fin0.values() if up(r['ANALITO']) == up(an) and iint(r['NIVEL']) == t
                 and r['PARTICIPA_ESTATISTICA'] == SIM and filt(r)]
        pop_d = [r['RESULTADO'] for r in fin1.values() if up(r['ANALITO']) == up(an) and iint(r['NIVEL']) == t
                 and r['PARTICIPA_ESTATISTICA'] == SIM and filt(r)]
        o_a, o_d = stats(pop_a), stats(pop_d)
        no_filtro = bool(m2) and filt(m2)
        ev3 = {'estatistica_antes': est_a, 'estatistica_depois': est_d, 'independente_antes': o_a,
               'independente_depois': o_d, 'manual_no_filtro': no_filtro, 'filtro': filt_desc}
        if not est_a or not est_d or est_a[0] in (None, '') or est_d[0] in (None, ''):
            ok3a = False
            ev3['falha'] = f'pré-condição ausente: a aba Estatística não tem a linha de {an} N{t}'
        elif not no_filtro:
            ok3a = False
            ev3['falha'] = 'pré-condição ausente: a data/equipamento do manual está fora do filtro da Estatística'
        else:
            n_a, m_a = int(est_a[0]), float(est_a[1])
            fechada = (n_a * m_a + v2) / (n_a + 1)
            dp_ok = (o_d[2] is None and est_d[2] in (None, '', 0, 0.0)) or perto_rel(est_d[2], o_d[2])
            ok3a = (n_a == o_a[0] and int(est_d[0]) == n_a + 1 == o_d[0] and perto_rel(est_d[1], o_d[1])
                    and perto_rel(est_d[1], fechada) and dp_ok)
            ev3['media_fechada'] = fechada
        reg(f'F03a Estatística N{t}: n+1; média e DP = cálculo independente (amostral, n−1) da Principal com o filtro do '
            'produto; média = (n·m + x)/(n+1)', ok3a, ev3)

        p_d = q.painel_stats()[t - 1]
        vals_p = [linha[t - 1] for linha, fl in zip(eng_d['val'], eng_d['fil']) if fl == 1 and isinstance(linha[t - 1], float)]
        n_p, m_p, s_p = stats(vals_p)
        fil_c1 = eng_d['fil'][i_c1] if i_c1 is not None else None
        ok3b = (fil_c1 == 1 and p_d[1] == n_p and p_a[1] not in (None, '') and int(p_d[1]) == int(p_a[1]) + 1
                and perto_rel(p_d[2], m_p, 2e-9) and (s_p is None or perto_rel(p_d[3], s_p, 2e-9)))
        reg(f'F03b engPainel N{t}: n+1 e n/média/DP = os pontos do Eng_Saida com filtro do período = 1 (recalculados à parte)',
            ok3b, {'antes': p_a, 'depois': p_d, 'independente': (n_p, m_p, s_p), 'filtro_da_linha_C1': fil_c1,
                   **({} if fil_c1 == 1 else {'falha': 'pré-condição ausente: o RUN de C1 está fora do período do Painel'})})
        fin_atual = fin1

        # ============================================================ F04 R_4s entre automatico e manual no mesmo RUN
        if C2 is None:
            # 23/11/2026 (fechamento): NAO_EXECUTADO, nao FAIL -- depende de haver nos DADOS uma corrida com o N1 nessa faixa
            # (a Hema tem as 3 corridas quase sempre completas); a regra entre automatico e manual ja e provada quando ha
            # (Bio 20/20). Como o qa_adr088, a falta da pre-condicao fica explicita e nao bloqueia o gate.
            RES.append({'teste': 'F04 R_4s entre automático e manual no mesmo RUN', 'resultado': 'NAO_EXECUTADO',
                        'evidencia': {'motivo': 'pré-condição ausente nos dados: nenhuma corrida C2 (condições do F01, '
                                                'outro dia) com o N1 em 1,2 <= |z| <= 2,0 e alvo do nível faltante',
                                      'alvo_N1': escolha['alvo1']}})
            print('NAO_EXECUTADO F04 -- pre-condicao ausente nos dados', flush=True)
        else:
            t2, run2 = C2['t'], C2['run']
            mt, st_ = alvo_t(t2)
            z1 = C2['z1']
            sinal = 1.0 if z1 > 0 else -1.0
            zt = z1 - sinal * 4.5
            zt = max(zt, -2.9) if sinal > 0 else min(zt, 2.9)
            v4 = round(mt + zt * st_, 6)
            zt_real = (v4 - mt) / st_
            r_aud = ms.audit_ult(q)
            r4, M4, _ = escrever_manual('F04', manual_dict(C2['dh_m'], C2['n1']['EQUIPAMENTO'], lote, t2,
                                                           C2['n1']['ANALITO'], v4, C2['n1'].get('MATRIZ')))
            so_entradas('F04')
            fin4 = q.final_por_id()
            m4 = fin4.get(M4)
            fl = wsq.flags_por_run(q).get(run2)
            evs = [r for r in wsq.eventos(q) if txt(r[wsq.C_REGRA - 1]) == 'R_4s' and iint(r[wsq.C_RUN - 1]) == run2
                   and up(r[wsq.C_ANALITO - 1]) == up(an)]
            # motivo da linha WESTGARD_EVENTO_* = 'regra detector' (mEstatistica.RegistrarMudancaWestgard/DescEventoWg):
            # so conta a linha do R_4s deste RUN (outra regra no mesmo RUN nao prova o R_4s)
            aud4 = [x for x in wsq.audit_desde(q, r_aud) if x['acao'].startswith('WESTGARD_EVENTO') and x['run'] == run2
                    and 'REMOVIDO' not in x['acao'] and (x['motivo'].split(' ') or [''])[0] == 'R_4s']
            ok4 = (m4 is not None and m4['STATUS_ANALITICO'] == 'ATIVO' and iint(m4['RUN']) == run2
                   and abs(zt_real) < 3 and abs(z1 - zt_real) > 4 and fl is not None
                   and iint(fl[2]) == 1 and iint(fl[(t2 - 1) * NEF + 2]) == 1
                   and txt(fl[6]) == 'REJEITADO' and txt(fl[(t2 - 1) * NEF + 6]) == 'REJEITADO'
                   and bool(evs) and bool(aud4))
            reg(f'F04 R_4s entre o N1 automático e o N{t2} manual no MESMO RUN: RR4s = 1 nos dois níveis, veredito '
                'REJEITADO, evento R_4s no RUN em Eventos_Westgard e WESTGARD_EVENTO_* na trilha', ok4,
                {'RUN': run2, 'manual': ms.resumo(m4), 'z_N1': round(z1, 3), f'z_N{t2}': round(zt_real, 3),
                 'alvo_N1_linha': C2['alvo_n1'], f'alvo_N{t2}': (mt, st_),
                 'flags_N1': fl[:NEF] if fl else None,
                 f'flags_N{t2}': fl[(t2 - 1) * NEF:t2 * NEF] if fl else None,
                 'eventos_R_4s': [[wsq.norm_cel(x) for x in r] for r in evs][:3], 'trilha_westgard': aud4[:6]})
            fin_atual = fin4

        # ============================================================ F05 manual em corrida NOVA antes das automaticas
        # valor dos manuais N1 do F05 e do F05e: a mediana dos N1 ATIVOS do escopo (lida da copia)
        v5 = mediana([r['RESULTADO'] for r in fin0.values() if no_escopo(r, an, lote, eqf) and iint(r['NIVEL']) == 1
                      and r['STATUS_ANALITICO'] == 'ATIVO' and eh_num(r['RESULTADO'])])
        if D5 is None:
            for nome in ('F05a', 'F05b', 'F05c', 'F05d'):
                reg(f'{nome} Manual numa corrida nova antes das automáticas de um dia', False,
                    {'falha': 'pré-condição ausente: nenhum _GK/dia com >= 2 corridas na janela, espaço antes da 1a '
                              '(fora do GAP e da TOL) e N1 automático ATIVO posterior'})
        else:
            X = D5['x']['ID_REGISTRO']
            fin_a5 = q.final_por_id()
            r_aud5 = ms.audit_ult(q)
            q.ex.xl.EnableEvents = True
            extras['inat'] = q.escrever_linha(TB_INAT, {'ID_REGISTRO': X, 'ANALITO': D5['x']['ANALITO'], 'MOTIVO': MOT_INAT})
            extras['com'] = q.escrever_linha(TB_COM, {'ID_REGISTRO': X, 'COMENTARIO_TECNICO':
                                                      'QA fluxo manual: inativado para conferir o X com o deslocamento.'})
            r5, M5, _ = escrever_manual('F05', manual_dict(D5['m'], D5['equip'], lote, 1, D5['analito'], v5,
                                                           D5.get('matriz')))
            so_entradas('F05')
            fin5 = q.final_por_id()
            m5 = fin5.get(M5)
            r_a, r_d = ms.runs(fin_a5), ms.runs(fin5)
            ruins5, n_desl = [], 0
            for i, (rn, cr) in r_a.items():
                ra = fin_a5[i]
                no_dia = rn is not None and cr is not None and real(ra) and gk(ra) == D5['gk']
                esp = (rn + 1, cr + 1) if no_dia else (rn, cr)
                n_desl += 1 if no_dia else 0
                if r_d.get(i) != esp:
                    ruins5.append({'ID': i, 'antes': (rn, cr), 'depois': r_d.get(i), 'esperado': esp})
            orc5 = etl.oraculo_run(list(fin5.values()), gap)
            run_m5 = run_de_dia(D5['dia'], 1)
            par_m5 = (iint(m5['CORRIDA_NO_DIA']), iint(m5['RUN'])) if m5 else None
            reg('F05a Manual N1 numa corrida NOVA antes das automáticas do dia: é a corrida 1 (= oráculo); todas as '
                'corridas do mesmo equipamento/matriz/lote/analito/dia andam +1 (RUN e CORRIDA_NO_DIA); nada fora delas '
                'muda', m5 is not None and m5['STATUS_ANALITICO'] == 'ATIVO' and par_m5 == (1, run_m5)
                and orc5.get(M5) == par_m5 and n_desl >= 2 and not ruins5,
                {'manual': ms.resumo(m5), 'oraculo_manual': orc5.get(M5), 'deslocados': n_desl,
                 'divergencias': ruins5[:10], 'n_divergencias': len(ruins5), 'dia': D5['dia']})

            e5 = cl.eng(q, nlv)
            o5 = cl.oraculo(list(fin5.values()), e5['analito'], e5['lote'], e5['equip'], nlv)
            dif5 = cl.comparar(e5, o5, nlv)
            L5 = Leitor(q).eng_por_run()
            consec = {t_ + 1: consecutivos(ativos_em_ordem(L5, t_)) for t_ in range(nlv)}
            reg('F05b Corrida do lote = oráculo em todos os níveis depois do deslocamento: sem lacuna e cada X entre as '
                'corridas certas', e5['n'] > 0 and run_m5 in e5['runs'] and not dif5 and all(consec.values()),
                {'divergencias': dif5[:10], 'consecutivos_por_nivel': consec, 'corridas_na_janela': e5['n'],
                 'RUN_do_manual_na_janela': run_m5 in e5['runs']})

            s_idx, nomes_s = serie_ponto(q, 1)
            pos5 = e5['runs'].index(run_m5) + 1 if run_m5 in e5['runs'] else None
            dp5 = str(run('mIntegracao.DadosDoPonto', 1, s_idx, pos5, teto=60)) if s_idx and pos5 else ''
            pt = dp5.split('|')
            cor5 = o5[0][0].get(run_m5)
            ok5c = (len(pt) == 9 and pt[0] == 'OK' and pt[1] == M5 and m5 is not None
                    and pt[2].upper() == up(m5['ANALITO']) and pt[3] == '1' and pt[4] == str(cor5)
                    and pt[5] == D5['m'].strftime('%d/%m/%Y %H:%M:%S') and perto(num_pt(pt[6]), v5)
                    and pt[7] == 'ATIVO' and pt[8] == txt(m5['LOTE']))
            reg('F05c DadosDoPonto do manual (série de pontos do N1, posição do RUN): ID, analito, nível, corrida do lote, '
                'data/hora, valor, ATIVO e lote', ok5c,
                {'retorno': dp5, 'serie': s_idx, 'series_do_grafico': nomes_s, 'ponto': pos5, 'corrida_esperada': cor5,
                 'data_hora_esperada': D5['m'].strftime('%d/%m/%Y %H:%M:%S'), 'valor_esperado': v5,
                 **({} if pos5 else {'falha': 'pré-condição ausente: o RUN do manual não está na janela do Eng_Saida'})})

            x5 = fin5.get(X)
            rx_esp = (r_a[X][0] + 1, r_a[X][1] + 1) if r_a.get(X) and r_a[X][0] is not None else None
            run_x, _, pos_x = x_do_id(L5, 0, X)
            reg('F05d O N1 automático inativado continua INATIVADO/X_VERMELHO com o RUN deslocado, e o X aparece nesse RUN',
                x5 is not None and x5['STATUS_ANALITICO'] == 'INATIVADO' and x5['TIPO_PLOTAGEM_LJ'] == 'X_VERMELHO'
                and x5['PARTICIPA_ESTATISTICA'] == NAO and rx_esp is not None and r_d.get(X) == rx_esp
                and run_x == rx_esp[0],
                {'inativado': ms.resumo(x5), 'RUN_esperado': rx_esp, 'X_no_grafico': (run_x, pos_x)})

            # F05f (INFO): a trilha de Westgard do deslocamento
            aud5 = wsq.audit_desde(q, r_aud5)
            wg5 = [x for x in aud5 if x['acao'].startswith('WESTGARD_EVENTO') and up(x['analito']) == up(an)]
            runs_dia = {rd[0] for i, rd in r_d.items() if rd[0] is not None and fin5[i].get('DATA') is not None
                        and real(fin5[i]) and gk(fin5[i]) == D5['gk']}
            pares = []
            for x in wg5:
                if 'REMOVIDO' in x['acao'] and x['run'] is not None and (x['run'] + 1) in runs_dia:
                    for y in wg5:
                        if y is not x and 'REMOVIDO' not in y['acao'] and y['run'] == x['run'] + 1 \
                                and y['nivel'] == x['nivel'] and y['motivo'].split(' ')[:2] == x['motivo'].split(' ')[:2]:
                            pares.append({'removido': x, 'novo': y})
            RES.append({'teste': 'F05f INFO Trilha de Westgard do deslocamento de RUN (REMOVIDO + NOVO do mesmo evento só '
                                 'por troca de RUN = achado: a chave do evento tem o RUN)',
                        'resultado': 'INFO',
                        'evidencia': {'linhas_WESTGARD': [{k: x[k] for k in ('acao', 'run', 'nivel', 'motivo')}
                                                          for x in wg5][:12],
                                      'pares_so_por_troca_de_RUN': len(pares), 'exemplos': pares[:3]}})
            print(f'INFO F05f -- {len(wg5)} linhas WESTGARD_EVENTO, {len(pares)} pares REMOVIDO+NOVO por troca de RUN',
                  flush=True)
            fin_atual = fin5

        # ---------------------------------------------------- F05e Liberacao
        # independente do D5: roda (ou reprova por pre-condicao propria) mesmo sem o dia do F05
        lib_a = lib_ler(q)
        lista_a = lib_runs(lib_a)
        nucleo = nucleo_lote(lote_uso)
        ev5e = {'lote_do_teste': lote, 'lote_em_uso': lote_uso, 'corridas_listadas': len(lista_a)}
        D0 = None
        if not wsq.mesmo_lote(lote, nucleo):
            ev5e['falha'] = 'pré-condição ausente: o lote do teste não é o lote em uso (a Liberação só lista ele)'
        elif len(lista_a) < 3:
            ev5e['falha'] = 'pré-condição ausente: a Liberação lista menos de 3 corridas'
        else:
            dias_ativos = {r['DATA'].date() for r in fin_atual.values() if r['STATUS_ANALITICO'] == 'ATIVO'
                           and hasattr(r.get('DATA'), 'year')
                           # como o VBA (NucleoLote(LOTE) = lote em uso), e tambem pelo valor ("010" = "10"): mais
                           # dias excluidos, nunca um dia com ATIVO do lote
                           and (nucleo_lote(r['LOTE']) == nucleo or wsq.mesmo_lote(nucleo_lote(r['LOTE']), nucleo))
                           and (not eqf or up(r['EQUIPAMENTO']) == eqf)}
            dias_esc = {r['DATA'].date() for r in fin_atual.values() if real(r) and no_escopo(r, an, lote, eqf)}
            # a chave (equipamento, matriz, lote, analito) mais frequente do escopo: a do manual do D0
            chave0 = collections.Counter(gk(r)[:4] for r in esc).most_common(1)[0][0]
            d_ = dia_de_run(lista_a[-1]) - dt.timedelta(days=1)
            primeiro = dia_de_run(lista_a[0])
            while d_ > primeiro:
                quando = dt.datetime.combine(d_, dt.time(12, 0))
                k_ins = sum(1 for x in lista_a if x < run_de_dia(d_, 1))
                if d_ < hoje and d_ not in dias_ativos and d_ not in dias_esc and k_ins <= min(len(lista_a) - 1, 198) \
                        and not conflito(idx, chave0, 1, quando, tol):
                    D0 = {'dia': d_, 'quando': quando, 'k': k_ins}
                    break
                d_ -= dt.timedelta(days=1)
            if D0 is None:
                ev5e['falha'] = ('pré-condição ausente: nenhum dia no período listado sem resultado ATIVO do lote em '
                                 'uso (para criar um RUN novo no meio da lista)')
        if D0 is None:
            reg('F05e Liberação: a assinatura (marcador em C/D) continua ao lado do MESMO RUN depois de um manual que '
                'cria um RUN novo no meio da lista', False, ev5e)
        else:
            j = D0['k']
            run_j = lista_a[j]
            extras['marcador'] = (j, lib_a[j][2], lib_a[j][3], run_j)
            marca = f'QA_FLUXO_MANUAL: marcador do RUN {run_j}'
            lib_gravar(q, j, 'QA', marca)
            r0 = next((r for r in esc if gk(r)[:4] == chave0), esc[-1])     # grafia real da chave (gk e maiuscula)
            r5e, M5e, _ = escrever_manual('F05e', manual_dict(D0['quando'], r0['EQUIPAMENTO'], lote, 1, r0['ANALITO'], v5,
                                                              r0.get('MATRIZ')))
            so_entradas('F05e')
            fin5e = q.final_por_id()
            m5e = fin5e.get(M5e)
            run_novo = iint(m5e['RUN']) if m5e else None
            lib_b = lib_ler(q)
            lista_b2 = lib_runs(lib_b)
            j2 = lista_b2.index(run_j) if run_j in lista_b2 else None
            ao_lado = (lib_b[j2][2], lib_b[j2][3]) if j2 is not None else None
            novos = sorted(set(lista_b2) - set(lista_a))
            sumiram = sorted(set(lista_a) - set(lista_b2))
            pre_ok = (m5e is not None and m5e['STATUS_ANALITICO'] == 'ATIVO' and run_novo in lista_b2
                      and novos == [run_novo] and j2 == j + 1)
            ok5e = pre_ok and ao_lado == ('QA', marca)
            ev5e.update({'dia_do_manual': D0['dia'], 'manual': ms.resumo(m5e), 'RUN_novo': run_novo,
                         'linha_do_marcador': 4 + j, 'RUN_do_marcador': run_j,
                         'RUN_ao_lado_do_marcador_depois': lib_b[j][0], 'conteudo_na_linha_do_marcador_depois':
                         (lib_b[j][2], lib_b[j][3]), 'linha_do_RUN_depois': None if j2 is None else 4 + j2,
                         'conteudo_ao_lado_do_RUN_depois': ao_lado, 'RUNs_novos_na_lista': novos,
                         'RUNs_que_sumiram': sumiram})
            if not pre_ok:
                ev5e['falha'] = ('pré-condição ausente: o manual não criou exatamente um RUN novo antes do RUN do '
                                 'marcador na lista')
            elif not ok5e:
                ev5e['falha'] = (f'risco de desalinhamento da Liberação: o marcador gravado ao lado do RUN {run_j} '
                                 f'(linha {4 + j}) ficou na linha {4 + j}, agora ao lado do RUN {lib_b[j][0]}; o RUN '
                                 f'{run_j} desceu para a linha {4 + j2} e ficou com {ao_lado}. As assinaturas (C:F e '
                                 'LiberStore) estão amarradas à POSIÇÃO e só A:B é reescrita '
                                 '(mLotes.AtualizarListaLiberacao)')
            reg('F05e Liberação: a assinatura (marcador em C/D) continua ao lado do MESMO RUN depois de um manual que '
                'cria um RUN novo no meio da lista', ok5e, ev5e)
            fin_atual = fin5e

        # ============================================================ F06 ultimo ponto e veredito
        alvo1 = escolha['alvo1']
        if f06 is None or alvo1 is None:
            for nome in ('F06a', 'F06b'):
                reg(f'{nome} Último ponto em alvo + 3,5 DP e correção', False,
                    {'falha': 'pré-condição ausente: ' + ('sem alvo (Média/DP) do N1 no lote' if alvo1 is None else
                                                          'sem espaço no último dia depois da última corrida (GAP/TOL)'),
                     'alvo_N1': alvo1})
        else:
            v6 = round(alvo1[0] + 3.5 * alvo1[1], 6)
            e6a = q.eng_saida()
            ver_a = veredito_por_run(e6a, 1)
            r_aud6 = ms.audit_ult(q)
            val6 = manual_dict(f06['m'], f06['equip'], lote, 1, f06['analito'], v6, f06.get('matriz'))
            if f06['m'] > dt.datetime.now() - G(5):
                # lancado depois da meia-noite (o ultimo dia e hoje): sem isso, DATA/HORA posterior ao lancamento
                val6['REGISTRADO_EM'] = dt.datetime.combine(f06['dia'] + dt.timedelta(days=1), dt.time(0, 10))
            r6, M6, _ = escrever_manual('F06', val6)
            so_entradas('F06_rejeita')
            fin6 = q.final_por_id()
            m6 = fin6.get(M6)
            run6 = iint(m6['RUN']) if m6 else None
            e6 = q.eng_saida()
            ver_d = veredito_por_run(e6, 1)
            fl6 = wsq.flags_por_run(q).get(run6)
            pv6 = eng_painel(q)[0][9]
            q.ex.xl.Calculate()
            j7 = txt(q.ws('Painel').Range('J7').Value)
            comuns = [rn for rn in ver_a if rn in ver_d and rn != run6]
            voltou_ok = [rn for rn in comuns if ver_a[rn] == 'REJEITADO' and ver_d[rn] != 'REJEITADO']
            k6 = sum(1 for rn in comuns if ver_d[rn] == 'REJEITADO') - sum(1 for rn in comuns if ver_a[rn] == 'REJEITADO') + 1
            ok6a = (m6 is not None and m6['STATUS_ANALITICO'] == 'ATIVO' and fl6 is not None and iint(fl6[0]) == 1
                    and txt(fl6[6]) == 'REJEITADO' and ver_d.get(run6) == 'REJEITADO' and not voltou_ok and k6 >= 1
                    and txt(pv6) == 'REJEITADO' and j7.upper().startswith('REJEITADO'))
            reg('F06a Manual N1 no último dia em alvo + 3,5 DP: 1_3s, veredito REJEITADO no RUN, nenhum outro veredito volta '
                'a OK, engPainel(N1) e J7 REJEITADO', ok6a,
                {'manual': ms.resumo(m6), 'alvo_N1': alvo1, 'valor': v6, 'flags_N1': fl6[:NEF] if fl6 else None,
                 'novos_rejeitados_k': k6, 'voltaram_a_OK': voltou_ok, 'engPainel_N1_col10': pv6, 'J7': j7,
                 'REGISTRADO_EM_escrito': 'REGISTRADO_EM' in val6})

            lin6a = ms.linha_manual(q, r6)
            r_aud6b = ms.audit_ult(q)
            alvo_v = round(alvo1[0], 6)
            q.ex.xl.EnableEvents = True
            ms.gravar_celula(q, r6, 'RESULTADO', alvo_v)
            so_entradas('F06_corrige')
            fin6b = q.final_por_id()
            m6b = fin6b.get(M6)
            e6b = q.eng_saida()
            ver_b = veredito_por_run(e6b, 1)
            fl6b = wsq.flags_por_run(q).get(run6)
            q.ex.xl.Calculate()
            j7b = txt(q.ws('Painel').Range('J7').Value)
            lin6b = ms.linha_manual(q, r6)
            ev6 = [x for x in ms.audit_desde(q, r_aud6b) if 'RESULTADO_MANUAL_ALTERADO' in acoes(x)
                   and x['motivo'] == f'ID {M6}']
            comuns_b = [rn for rn in ver_a if rn in ver_b and rn != run6]
            dif_ver = [(rn, ver_a[rn], ver_b[rn]) for rn in comuns_b if ver_a[rn] != ver_b[rn]]
            rej_b = sum(1 for v in ver_b.values() if v == 'REJEITADO')
            j7_ok = j7b.upper().startswith('OK') if rej_b == 0 else j7b.upper().startswith('REJEITADO')
            ok6b = (m6b is not None and iint(m6b['RUN']) == run6 and perto(m6b['RESULTADO'], alvo_v)
                    and fl6b is not None and iint(fl6b[0]) == 0 and txt(fl6b[6]) == 'OK' and not dif_ver and j7_ok
                    and len(ev6) == 1 and perto(ev6[0]['res_ant'], v6) and perto(ev6[0]['res_novo'], alvo_v)
                    and lin6b.get('USUARIO') == lin6a.get('USUARIO') and lin6b.get('REGISTRADO_EM') == lin6a.get('REGISTRADO_EM')
                    and lin6b.get('ID_REGISTRO') == M6)
            reg('F06b Corrigir o RESULTADO do manual para o alvo: OK no RUN, os outros vereditos voltam aos de antes do '
                'manual, J7 coerente; RESULTADO_MANUAL_ALTERADO (antes -> depois); mesmo ID, USUARIO e REGISTRADO_EM',
                ok6b, {'manual': ms.resumo(m6b), 'flags_N1': fl6b[:NEF] if fl6b else None,
                       'vereditos_diferentes_de_antes': dif_ver[:10], 'rejeitados_na_janela': rej_b, 'J7': j7b,
                       'trilha': ev6, 'carimbo_antes': {k: lin6a.get(k) for k in ('ID_REGISTRO', 'USUARIO', 'REGISTRADO_EM')},
                       'carimbo_depois': {k: lin6b.get(k) for k in ('ID_REGISTRO', 'USUARIO', 'REGISTRADO_EM')}})

        # ============================================================ F07 idempotencia e retirada
        cq_a = ms.tabela_v2(q, TB_FIN)
        r_aud7 = ms.audit_ult(q)
        so_entradas('F07_idem_1')
        so_entradas('F07_idem_2')
        completo('F07_idem_completo')
        cq_b = ms.tabela_v2(q, TB_FIN)
        res7 = so_resultado(ms.audit_desde(q, r_aud7))
        dif7 = [i for i in set(cq_a) | set(cq_b) if cq_a.get(i) != cq_b.get(i)]
        reg('F07a Idempotência: AtualizarSoEntradas 2x e o ATUALIZAR completo deixam a Principal idêntica e não gravam '
            'nenhuma linha RESULTADO_* na trilha', not dif7 and not res7,
            {'ids_diferentes': dif7[:10], 'trilha_RESULTADO': res7[:10], 'manuais': [m for _, _, m in manuais]})

        r_aud7b = ms.audit_ult(q)
        q.ex.xl.EnableEvents = True
        if 'inat' in extras:
            apagar_linha(q, TB_INAT, extras['inat'])
        if 'com' in extras:
            apagar_linha(q, TB_COM, extras['com'])
        for _, r_, _ in manuais:
            apagar_linha(q, TB_MAN, r_)
        if 'marcador' in extras:
            # o marcador acompanha o RUN (correcao da Liberacao, 23/11/2026): apagar na linha ATUAL do RUN dele
            jm, c0, d0, run_m = extras['marcador']
            lista_m = [iint(r[0]) for r in lib_ler(q)]
            lib_gravar(q, lista_m.index(run_m) if run_m in lista_m else jm, c0, d0)
        completo('F07_retirada')
        fin7 = q.final_por_id()
        cq7 = ms.tabela_v2(q, TB_FIN)
        run7 = ms.runs(fin7)
        dif_run = [i for i in set(run0) | set(run7) if run0.get(i) != run7.get(i)]
        dif_cq = [i for i in set(cq0) | set(cq7) if cq0.get(i) != cq7.get(i)]
        cab = [c.Name for c in q.lo(TB_FIN).ListColumns]
        cols_dif = sorted({cab[k] for i in dif_cq[:50] if i in cq0 and i in cq7
                           for k in range(min(len(cq0[i]), len(cab))) if cq0[i][k] != cq7[i][k]})
        e7 = cl.eng(q, nlv)
        dif_eng = cl.comparar(e7, cl.oraculo(list(fin7.values()), e7['analito'], e7['lote'], e7['equip'], nlv), nlv)
        p7 = q.painel_stats()
        est7 = {t_: q.estat_motor(an, t_) for t_ in range(1, nlv + 1)}
        of7 = {k for k in wsq.oficiais(wsq.eventos(q)) if k[0] == up(an)}
        lib7 = lib_ler(q)

        def igual_num(a, b):
            if a is None or b is None:
                return a == b
            return all((x == y) or (eh_num(x) and eh_num(y) and perto_rel(x, y)) for x, y in zip(a, b)) and len(a) == len(b)
        aud7 = ms.audit_desde(q, r_aud7b)
        retirados = {m: len([x for x in aud7 if 'RESULTADO_MANUAL_RETIRADO' in acoes(x) and x['motivo'] == f'ID {m}'])
                     for _, _, m in manuais}
        ok7b = (not dif_run and not dif_cq and not dif_eng and e7['cor'] == e0['cor'] and e7['runs'] == e0['runs']
                and all(igual_num(a, b) for a, b in zip(p7, p0)) and len(p7) == len(p0)
                and all(igual_num(est7[t_], est0[t_]) for t_ in est0) and of7 == of0
                and lib_norm(lib7) == lib_norm(lib0) and all(v == 1 for v in retirados.values()))
        reg('F07b Retirada (apagar as linhas manuais com um Delete, a inativação e o comentário): RUN de todos os IDs, '
            'Principal, corrida do lote, engPainel, Estatística, eventos de Westgard e Liberação = estado inicial; 1 '
            'RESULTADO_MANUAL_RETIRADO por manual', ok7b,
            {'RUN_diferente': dif_run[:10], 'n_RUN_diferente': len(dif_run), 'principal_ids_diferentes': dif_cq[:10],
             'n_principal_diferente': len(dif_cq), 'colunas_diferentes': cols_dif,
             'corrida_do_lote_divergencias': dif_eng[:5], 'corrida_do_lote_igual_ao_inicio': e7['cor'] == e0['cor'],
             'painel_inicio': p0, 'painel_fim': p7, 'estatistica_inicio': est0, 'estatistica_fim': est7,
             'westgard_oficiais_iguais': of7 == of0, 'eventos_so_no_inicio': sorted(map(str, of0 - of7))[:5],
             'eventos_so_no_fim': sorted(map(str, of7 - of0))[:5], 'liberacao_igual': lib_norm(lib7) == lib_norm(lib0),
             'liberacao_dif': [(4 + i, a, b) for i, (a, b) in enumerate(zip(lib_norm(lib0), lib_norm(lib7))) if a != b][:6],
             'retirados_por_manual': retirados,
             'reativado': [x for x in aud7 if 'RESULTADO_REATIVADO' in acoes(x)][:3]})
        q.cfg('MODO_FONTE', modo0)
        q.cfg('MOTOR_DADOS', motor0)
    except Exception as ex_geral:                    # noqa: BLE001
        reg('F99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:],
             'manuais_escritos': manuais, 'tempos_s': tempos})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('F00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1, 'tempos_s': tempos})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    n_ok = sum(1 for r in RES if r['resultado'] == 'PASS')
    print(f'\n=== {produto} fluxo do manual: {n_ok} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
    sys.exit(1 if falhas else 0)
