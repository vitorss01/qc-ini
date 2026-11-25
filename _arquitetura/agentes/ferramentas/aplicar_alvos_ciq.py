# -*- coding: utf-8 -*-
"""Alvo do laboratorio (Media/DP de cada LOTE x ANALITO x NIVEL) pelos PRIMEIROS resultados do CIQ.

Pedido do usuario (22/11/2026): para cada lote do cadastro (regLoteCol) presente na tblCQ_Final, cada analito do
catalogo (Especificações A4:A43) e cada nivel, a Media e o DP da LotesStore passam a ser a media aritmetica e o DP
amostral (n-1) dos N PRIMEIROS resultados com PARTICIPA_ESTATISTICA = 'SIM' do lote (todos os equipamentos), na ordem
(RUN, DATA_HORA, ID_REGISTRO). Com MIN..N-1 resultados usa todos; com menos de MIN nao mexe ('insuficiente'). Sem
exclusao de pontos. Sobrescreve o que houver (bula, digitado, TESTE do ADR-076, vazio).

Uso: python aplicar_alvos_ciq.py <Bioquimica|Hematologia> [--arquivo <xlsm>] [--simular] [--n 40] [--min 20]
                                 [--so-copia <saida.xlsm>] [--forcar]
     python aplicar_alvos_ciq.py --autoteste          (so o calculo, dados sinteticos, sem Excel)

ADR-102 (25/11/2026): a aplicacao em massa fica aposentada (alvo = escolha do analista no Painel); a ferramenta continua
para o --simular e como oraculo, e PULA as linhas lote|analito cuja origem (W) comeca com 'MANUAL' (contadas no JSON
em pulados_origem_manual). Adendo 25/11: a origem 'EXCECAO TEMPORARIA ...' (valores digitados no Painel: Media da bula
e DP do lote anterior) NAO e pulada -- e temporaria por definicao e deve ser substituida pelos valores do laboratorio;
como a aplicacao em massa esta aposentada, quem a substitui e o "Aplicar Alvos ao Lote" do Painel, e o --simular apenas
a mostra como candidata a troca.

  --arquivo   arquivo de entrada (padrao: QC_<produto>.xlsm na raiz do repositorio = a producao)
  --simular   calcula e grava o CSV/JSON em %TEMP%\\qcw\\alvos; NADA e gravado na planilha
  --so-copia  aplica numa copia e grava o resultado em <saida.xlsm>; a entrada NAO e trocada
  (sem flag)  aplica e troca a entrada, com backup _backup_pre_alvos_<stamp>/ ao lado dela
  --forcar    ignora a marca blAlvosCIQ (a ferramenta ja rodou neste arquivo)

Passos (modelo: aplicar_atualizar.py):
 1. Y00 sha256 da entrada; copia para %TEMP%\\qcw\\alvos (a entrada nunca e aberta pelo Excel);
 2. na copia (qa_final.QA): pre-condicoes (marca, mEtlSync.MotorVBA/mIntegracao.ModoHistorico perguntados a VBA --
    sem q.cfg, que desprotege a Cfg_Integracao --, catalogo, cadastro, blocos por lote, tblCQ_Final, LotesStore,
    Westgard_Fotos, Especificações); le tblCQ_Final e LotesStore A:Z (com os bits TESTE X..Z);
 3. calcula em Python (calcular_alvos, funcao pura) e grava o CSV alvos_ciq_<produto>_<stamp>.csv (UTF-8 BOM, ';');
    conta os EVENTOS da Westgard_Fotos por lote x equipamento (antes). Com --simular para aqui, sem salvar;
 4. currentUser = 'aplicar_alvos_ciq' (autoria em V da LotesStore e no Audit_Log); por linha lote|analito com algum
    nivel recalculado: mLotes.GravarParametroStore(lote, an, j, valor, EspecDoLote, origem) para j = 2nv-1 (Media)
    e 2nv (DP); mAuditoria.Auditar PARAMETRO_LOTE_ALTERADO por campo que mudou; depois de cada lote, confere W e o
    bit TESTE e corrige a mao o caso CStr(antes) = CStr(valor) (GravarParametroStore nao toca W nem o bit);
    PARAMETRO_LOTE_CALCULADO por lote; mLotesBlocos.EscreverValoresDoLote; marca blAlvosCIQ;
 5. InvalidarLotes, CarregarParametrosView(LoteEmTela), InvalidarParametros, AtualizarResumoLote; conferencias
    (store, trilha, blocos, ConferirLotesStore, leitura pela VBA);
 6. mIntegracao.AtualizarSoEntradas("") (teto 1800; sem DB_SEAC com MOTOR_DADOS = VBA ou MODO_FONTE = HISTORICO),
    que refaz o escopo do Westgard dos lotes com alvo novo e audita [RETROATIVO] (A07 le o retorno da foto no TEMPOS:
    o nucleo engole o erro dela e devolve OK|); recontagem da Westgard_Fotos; A08b recalcula o alvo sobre a
    tblCQ_Final DEPOIS do ATUALIZAR (entrada pendente mudaria os n primeiros: FAIL, nada salvo);
 7. currentUser = '', Save com eventos desligados, fecha, blindar_entrega.blindar + verificar, F00 sha da entrada
    inalterado; --so-copia: copia para a saida; sem flag: backup, troca e confere o sha.
Qualquer FAIL antes de salvar: nada e salvo nem trocado (saida 1). Encerra so o Excel que abriu.

ATENCAO: o repositorio e PUBLICO e o .gitignore ainda nao cobre _backup_pre_alvos_*/ nem alvos_ciq_*.csv/json.
"""
import argparse
import collections
import datetime
import json
import math
import os
import re
import shutil
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
INT = os.path.join(REPO, '_arquitetura', 'integracao_dbseac')

sys.path.insert(0, INT)
from instalar_integracao import SENHA  # noqa: E402 -- fonte unica da senha das abas (troca: OC-0067)
MARCA = 'blAlvosCIQ'
MODULO = 'aplicar_alvos_ciq'
PEDIDO = '22/11/2026'
CALC_MANUAL, CALC_AUTO = -4135, -4105

# Texto da origem (W da LotesStore, um por linha lote|analito). O BO de qa_blocos_lote pode importar estes padroes:
# so os niveis recalculados tem 'Nk: n=<n> (dd/mm/aaaa a dd/mm/aaaa)'; os mantidos dizem 'Nk: mantido (...)' sem 'n='.
PREFIXO_W = 'CIQ {data}: alvo do laboratorio = media/DP amostral dos primeiros resultados Ativos do lote ' \
            '(todos os equipamentos) -- '
RX_W_CIQ = re.compile(r'^CIQ (\d\d/\d\d/\d{4}): alvo do laboratorio = ')
RX_W_NIVEL = re.compile(r'N(\d): n=(\d+) \((\d\d/\d\d/\d{4}|\?) a (\d\d/\d\d/\d{4}|\?)\)')
RX_W_MANTIDO = re.compile(r'N(\d): mantido \(')


# =============================================================== calculo puro (sem Excel)
def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000) \
        and math.isfinite(float(v))


def txt(v):
    if isinstance(v, float) and math.isfinite(v) and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def chave(v):
    """LOTE e ANALITO comparados como a VBA (mLotes.ChaveLote): maiusculas, sem espacos; lote 8974.0 -> '8974'."""
    return txt(v).upper()


def bit_val(v):
    """Bit TESTE (X..Z) lido como a VBA le (mLotes.AlvoDeTeste: Val(CStr(v)) > 0): numero OU texto ('1', '3')."""
    if num(v):
        return int(v)
    try:
        return int(float(txt(v).replace(',', '.')))
    except (TypeError, ValueError):
        return 0


def _ord_num(v):
    if num(v):
        return (0, float(v), '')
    try:
        return (0, float(str(v).strip().replace(',', '.')), '')
    except (TypeError, ValueError):
        return (1, 0.0, txt(v))


def _ord_data(v):
    return (0, v) if isinstance(v, datetime.datetime) else (1, datetime.datetime.max)


def ordem(r):
    """(RUN, DATA_HORA, ID_REGISTRO): RUN primeiro, como mLotes.PrimeirosN; os empates de RUN pela hora e pelo id."""
    return (_ord_num(r.get('RUN')), _ord_data(r.get('DATA_HORA')), _ord_num(r.get('ID_REGISTRO')))


def media_dp(vals):
    """Media e DP amostral (n-1) em duas passadas, na ordem dada (a mesma aritmetica de mLotes.PrimeirosN)."""
    n = len(vals)
    if n == 0:
        return None, None
    soma = 0.0
    for x in vals:
        soma += x
    m = soma / n
    if n < 2:
        return m, None
    q = 0.0
    for x in vals:
        q += (x - m) ** 2
    return m, math.sqrt(q / (n - 1))


def calcular_alvos(linhas, nlv, n=40, minimo=20, lotes=None, analitos=None, ignorados=None):
    """linhas: dicts da tblCQ_Final (RUN, DATA_HORA, NIVEL, LOTE, ANALITO, RESULTADO, PARTICIPA_ESTATISTICA,
    EQUIPAMENTO, ID_REGISTRO). lotes/analitos: conjuntos de CHAVES (chave()) aceitas, None = todas.
    Devolve {(LOTE, ANALITO, nivel): resultado}; status 'calculado' (n_usado >= minimo e DP > 0), 'insuficiente'
    (n_usado < minimo) ou 'dp_zero'. ignorados (Counter opcional) recebe o que ficou fora e por que."""
    if not (2 <= minimo <= n):
        raise ValueError(f'precisa 2 <= minimo ({minimo}) <= n ({n})')
    ign = ignorados if ignorados is not None else collections.Counter()
    grupos = collections.defaultdict(list)
    for r in linhas:
        if txt(r.get('PARTICIPA_ESTATISTICA')).upper() != 'SIM':
            continue
        lt, an = chave(r.get('LOTE')), chave(r.get('ANALITO'))
        if lotes is not None and lt not in lotes:
            ign['lote_fora_do_cadastro'] += 1
            continue
        if analitos is not None and an not in analitos:
            ign['analito_fora_do_catalogo'] += 1
            continue
        try:
            nv = int(float(r.get('NIVEL')))
        except (TypeError, ValueError):
            ign['nivel_invalido'] += 1
            continue
        if not 1 <= nv <= nlv:
            ign['nivel_fora'] += 1
            continue
        if not num(r.get('RESULTADO')):
            ign['resultado_nao_numerico'] += 1
            continue
        grupos[(lt, an, nv)].append(r)
    out = {}
    for k, rs in grupos.items():
        rs = sorted(rs, key=ordem)
        todos = [float(r['RESULTADO']) for r in rs]
        usados = rs[:n]
        vu = todos[:n]
        m, s = media_dp(vu)
        if len(vu) < minimo:
            st, m, s = 'insuficiente', None, None
        elif s is None or s <= 0:
            st, m, s = 'dp_zero', None, None
        else:
            st = 'calculado'
        datas = [r.get('DATA_HORA') for r in usados if isinstance(r.get('DATA_HORA'), datetime.datetime)]
        mp, sp = media_dp(todos)
        eq = collections.OrderedDict()
        for i, r in enumerate(rs):
            e = chave(r.get('EQUIPAMENTO')) or '(sem)'
            d = eq.setdefault(e, {'n_lote': 0, 'soma': 0.0, 'n_usado': 0})
            d['n_lote'] += 1
            d['soma'] += todos[i]
            if i < n:
                d['n_usado'] += 1
        for d in eq.values():
            d['media_lote'] = d.pop('soma') / d['n_lote']
        out[k] = {'lote': k[0], 'analito': k[1], 'nivel': k[2], 'status': st, 'n_lote': len(rs), 'n_usado': len(vu),
                  'media': m, 'dp': s, 'data_ini': min(datas) if datas else None, 'data_fim': max(datas) if datas else None,
                  'run_ini': txt(usados[0].get('RUN')), 'run_fim': txt(usados[-1].get('RUN')),
                  'ids_usados': [txt(r.get('ID_REGISTRO')) for r in usados],
                  'media_periodo': mp, 'dp_periodo': sp, 'por_equip': dict(eq), 'valores': todos}
    return out


def pct_fora(valores, m, s, k):
    """% dos resultados com |x - m| > k*s (projecao de 1_2s / 1_3s); None sem parametro valido."""
    if not valores or not num(m) or not num(s) or s <= 0:
        return None
    return 100.0 * sum(1 for x in valores if abs(x - m) > k * s) / len(valores)


def classificar_origem(w, bit, m, s):
    """De onde vinha o par antes: TESTE (bit X..Z), vazio, CIQ (botao), digitado ou bula/outro."""
    if bit_val(bit) > 0:
        return 'TESTE'
    if txt(m) == '' and txt(s) == '':
        return 'vazio'
    w = txt(w)
    if w.upper().startswith('CIQ'):
        return 'CIQ'
    if 'digitado' in w.lower():
        return 'digitado'
    return 'bula/outro'


def origem_manual(w):
    """ADR-102: a linha lote|analito tem alvo gravado pelo analista no Painel ("Aplicar Alvos ao Lote": W comeca com
    'MANUAL'). A ferramenta NAO a toca (todos os niveis da linha): o que o analista gravou fica. A origem
    'EXCECAO TEMPORARIA' (adendo 25/11) NAO e manual neste sentido: e temporaria e deve ser substituida."""
    return txt(w).upper().startswith('MANUAL')


def selecionar_pares(lotes, cat, store, calc, nlv):
    """Pares (LOTE, AN, nv) a considerar, na ordem lote do cadastro x analito do catalogo x nivel. store: {'LOTE|AN':
    (linha, valores A:Z)}. Devolve (antes, pares, manuais): manuais = chaves 'LOTE|AN' puladas (ADR-102)."""
    antes, pares, manuais = {}, [], []
    for lt in lotes:
        for an in cat:
            linha = store.get(f'{lt}|{an}', (0, (None,) * 26))[1]
            if origem_manual(linha[22]):
                manuais.append(f'{lt}|{an}')
                continue
            for nv in range(1, nlv + 1):
                m, s = linha[2 + 2 * (nv - 1)], linha[3 + 2 * (nv - 1)]
                rot = classificar_origem(linha[22], linha[23 + nv - 1], m, s)
                antes[(lt, an, nv)] = (m, s, rot)
                if (lt, an, nv) in calc or rot != 'vazio':
                    pares.append((lt, an, nv))
    return antes, pares, manuais


def status_par(res, m_antes, s_antes):
    """aplicado | igual (valor novo identico ao gravado: so W e o bit TESTE mudam) | insuficiente | dp_zero."""
    if res is None:
        return 'insuficiente'
    if res['status'] != 'calculado':
        return res['status']
    if num(m_antes) and num(s_antes) and float(m_antes) == res['media'] and float(s_antes) == res['dp']:
        return 'igual'
    return 'aplicado'


def _data(d):
    return d.strftime('%d/%m/%Y') if isinstance(d, datetime.datetime) else '?'


def texto_origem(hoje, niveis, minimo):
    """W da linha lote|analito descrevendo TODOS os niveis. niveis: [(nv, status, res, rotulo_antes, n_lote)]."""
    partes = []
    for nv, st, res, rot, n_lote in niveis:
        if st in ('aplicado', 'igual'):
            partes.append(f'N{nv}: n={res["n_usado"]} ({_data(res["data_ini"])} a {_data(res["data_fim"])})')
        elif st == 'dp_zero':
            partes.append(f'N{nv}: mantido ({rot}; DP zero nos {res["n_usado"]} primeiros)')
        else:
            partes.append(f'N{nv}: mantido ({rot}; {n_lote} resultado(s) < {minimo})')
    return PREFIXO_W.format(data=hoje) + ' | '.join(partes)


def montar_linhas(calc, antes, pares, nlv, minimo):
    """Linhas do CSV. antes: {(LOTE, AN, nv): (m, s, rotulo)}; pares: [(LOTE, AN, nv)] na ordem desejada (lote do
    cadastro x analito do catalogo x nivel). Devolve (colunas, linhas, equipamentos)."""
    equips = sorted({e for r in calc.values() for e in r['por_equip']})
    cols = ['lote', 'analito', 'nivel', 'status', 'n_lote', 'n_usado', 'data_ini', 'data_fim', 'run_ini', 'run_fim',
            'origem_antes', 'media_antes', 'dp_antes', 'media_nova', 'dp_nova', 'deslocamento_em_dp_antes',
            'razao_dp_nova_antes', 'media_periodo_inteiro', 'dp_periodo_inteiro',
            'pct_fora_2dp_antes', 'pct_fora_3dp_antes', 'pct_fora_2dp_depois', 'pct_fora_3dp_depois']
    for e in equips:
        cols += [f'n_lote_{e}', f'media_lote_{e}', f'n_usado_{e}']
    cols += ['maior_dif_entre_equip', 'marca_dif_equip_maior_1dp']
    out = []
    for k in pares:
        res = calc.get(k)
        m0, s0, rot = antes.get(k, (None, None, 'vazio'))
        st = status_par(res, m0, s0)
        novo = st in ('aplicado', 'igual')
        m1, s1 = (res['media'], res['dp']) if novo else (m0, s0)
        ok0 = num(m0) and num(s0) and s0 > 0
        vals = res['valores'] if res else []
        lin = {'lote': k[0], 'analito': k[1], 'nivel': k[2], 'status': st,
               'n_lote': res['n_lote'] if res else 0, 'n_usado': res['n_usado'] if res else 0,
               'data_ini': _data(res['data_ini']) if res else '', 'data_fim': _data(res['data_fim']) if res else '',
               'run_ini': res['run_ini'] if res else '', 'run_fim': res['run_fim'] if res else '',
               'origem_antes': rot, 'media_antes': m0 if num(m0) else txt(m0), 'dp_antes': s0 if num(s0) else txt(s0),
               'media_nova': res['media'] if novo else None, 'dp_nova': res['dp'] if novo else None,
               'deslocamento_em_dp_antes': (res['media'] - m0) / s0 if novo and ok0 else None,
               'razao_dp_nova_antes': res['dp'] / s0 if novo and ok0 else None,
               'media_periodo_inteiro': res['media_periodo'] if res else None,
               'dp_periodo_inteiro': res['dp_periodo'] if res else None,
               'pct_fora_2dp_antes': pct_fora(vals, m0, s0, 2), 'pct_fora_3dp_antes': pct_fora(vals, m0, s0, 3),
               'pct_fora_2dp_depois': pct_fora(vals, m1, s1, 2), 'pct_fora_3dp_depois': pct_fora(vals, m1, s1, 3)}
        pe = res['por_equip'] if res else {}
        for e in equips:
            d = pe.get(e)
            lin[f'n_lote_{e}'] = d['n_lote'] if d else 0
            lin[f'media_lote_{e}'] = d['media_lote'] if d else None
            lin[f'n_usado_{e}'] = d['n_usado'] if d else 0
        medias = [d['media_lote'] for d in pe.values() if d['n_lote'] >= 2]
        dif = (max(medias) - min(medias)) if len(medias) >= 2 else None
        dp_ref = s1 if num(s1) and s1 > 0 else None
        lin['maior_dif_entre_equip'] = dif
        lin['marca_dif_equip_maior_1dp'] = '' if dif is None or dp_ref is None else ('SIM' if dif > dp_ref else 'nao')
        out.append(lin)
    return cols, out, equips


def retorno_foto_wg(r):
    """Retorno de mEstatistica.AuditarFotosWestgard dentro do resumo do ATUALIZAR (fim do TEMPOS, mIntegracao:188):
    '... westgard_fotos 1,23s [RECALC=k/N MUDANCAS=m]'. Devolve (texto entre colchetes, k ou None)."""
    mw = re.search(r'westgard_fotos \S+ \[(.*)\]\s*$', r or '', re.S)
    wgres = mw.group(1) if mw else ''
    mk = re.match(r'RECALC=(\d+)/(\d+)', wgres)
    return wgres, (int(mk.group(1)) if mk else None)


def fmt(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'SIM' if v else 'NAO'
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return f'{v:.10g}'.replace('.', ',')      # Excel pt-BR abre o CSV com ';' e virgula decimal
    return str(v).replace(';', ',').replace('\n', ' ')


def escrever_csv(caminho, cols, linhas):
    with open(caminho, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(';'.join(cols) + '\r\n')
        for lin in linhas:
            f.write(';'.join(fmt(lin.get(c)) for c in cols) + '\r\n')


# =============================================================== aplicacao no Excel (COM)
class Abortar(Exception):
    pass


def main(argv):
    ap = argparse.ArgumentParser(description='Alvo do laboratorio pelos primeiros resultados do CIQ')
    ap.add_argument('produto', choices=['Bioquimica', 'Hematologia'])
    ap.add_argument('--arquivo')
    ap.add_argument('--simular', action='store_true')
    ap.add_argument('--n', type=int, default=40)
    ap.add_argument('--min', dest='minimo', type=int, default=20)
    ap.add_argument('--so-copia', dest='so_copia')
    ap.add_argument('--forcar', action='store_true')
    a = ap.parse_args(argv)
    if not (2 <= a.minimo <= a.n):
        ap.error(f'precisa 2 <= --min ({a.minimo}) <= --n ({a.n})')
    if a.simular and a.so_copia:
        ap.error('--simular e --so-copia sao exclusivos')

    sys.path.insert(0, INT)
    sys.path.insert(0, os.path.join(INT, 'testes'))
    sys.path.insert(0, os.path.join(REPO, '_arquitetura', 'etapa1_multilote'))
    import pythoncom  # noqa: E402
    import blindar_entrega  # noqa: E402
    from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
    from qa_blocos_lote import blocos, cadastro, O_AN, NAN, perto  # noqa: E402
    E = pythoncom.Empty

    produto = a.produto
    entrada = os.path.abspath(a.arquivo) if a.arquivo else os.path.join(REPO, f'QC_{produto}.xlsm')
    modo = 'simular' if a.simular else ('so-copia' if a.so_copia else 'producao')
    saida = os.path.abspath(a.so_copia) if a.so_copia else None
    stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')
    hoje = datetime.date.today().strftime('%d/%m/%Y')
    pasta = os.path.join(os.environ['TEMP'], 'qcw', 'alvos')
    os.makedirs(pasta, exist_ok=True)
    base = f'alvos_ciq_{produto}_{stamp}'
    csv_tmp, json_tmp = os.path.join(pasta, base + '.csv'), os.path.join(pasta, base + '.json')
    rel = {'produto': produto, 'modo': modo, 'quando': datetime.datetime.now().isoformat(timespec='seconds'),
           'entrada': entrada, 'saida': saida, 'n': a.n, 'minimo': a.minimo, 'pedido': PEDIDO}

    def em_uso(p):
        d, b = os.path.dirname(p), os.path.basename(p)
        return [x for x in (os.path.join(d, '~$' + b), os.path.join(d, '~$' + b[2:])) if os.path.exists(x)]

    def gravar_json():
        rel['checks'] = RES
        with open(json_tmp, 'w', encoding='utf-8') as f:
            json.dump(rel, f, ensure_ascii=False, indent=1, default=str)

    def falhas():
        return [r['teste'] for r in RES if r['resultado'] != 'PASS']

    if not os.path.isfile(entrada):
        reg('Y00 sha256 da entrada no inicio', False, {'pre-condicao ausente': f'arquivo nao existe: {entrada}'})
        return 1
    if saida and os.path.normcase(saida) == os.path.normcase(entrada):
        reg('Y00 sha256 da entrada no inicio', False, {'erro': '--so-copia nao pode ser a propria entrada'})
        return 1
    h0 = sha256_arquivo(entrada)
    trava = em_uso(entrada)
    reg('Y00 sha256 da entrada no inicio (e ninguem com o arquivo aberto; na simulacao so informativo)',
        not trava or a.simular, {'arquivo': entrada, 'sha256': h0, 'arquivo_de_trava': trava})
    rel['sha_entrada'] = h0
    if trava and not a.simular:
        return 1
    copia = os.path.join(pasta, os.path.basename(entrada))   # mesmo nome: ThisWorkbook.Name vai para o Audit_Log
    shutil.copy2(entrada, copia)

    q = QA(produto, copia)
    salvou = False
    try:
        xl, wb, nlv = q.ex.xl, q.wb, q.nlv
        # ------------------------------------------------------------ A00 pre-condicoes
        pre, ev = [], {}
        try:
            ev['marca_antes'] = wb.Names(MARCA).RefersTo
        except Exception:
            ev['marca_antes'] = None
        if ev['marca_antes'] and not a.forcar and not a.simular:
            pre.append(f'ja aplicado neste arquivo (marca {MARCA} = {ev["marca_antes"]}); use --forcar para refazer')
        # a MESMA regra de mIntegracao.AtualizarDadosCore (semOrigem = hist Or (soEntradas And vba)), perguntada a
        # propria VBA: q.cfg desprotegeria a Cfg_Integracao e a copia seria salva (e trocada) assim
        try:
            ev['MotorVBA'] = bool(q.run('mEtlSync.MotorVBA', teto=60))
            ev['ModoHistorico'] = bool(q.run('mIntegracao.ModoHistorico', teto=60))
        except Exception as e:
            pre.append(f'pre-condicao ausente: mEtlSync.MotorVBA / mIntegracao.ModoHistorico ({repr(e)[:160]})')
        else:
            if not (ev['MotorVBA'] or ev['ModoHistorico']):
                pre.append('MOTOR_DADOS != VBA e MODO_FONTE != HISTORICO: AtualizarSoEntradas leria o DB_SEAC')
        abas = [w.Name for w in wb.Worksheets]
        # ADR-074: Analitos = blocos por lote e Especificações = motor; antes da migracao a Analitos e o motor
        for aba in ('LotesStore', 'Westgard_Fotos', 'Analitos', 'Especificações', 'Audit_Log'):
            if aba not in abas:
                pre.append(f'pre-condicao ausente: aba {aba}')
        catalogo, cad, bl0 = [], [], []
        try:
            catalogo = [txt(v[0]) for v in q.espec().Range('A4:A43').Value if txt(v[0])]
            cad = cadastro(q)
            if 'Analitos' in abas:
                bl0 = blocos(q.ws('Analitos'))
        except Exception as e:
            pre.append(f'pre-condicao ausente: catalogo/cadastro/blocos ilegiveis ({repr(e)[:160]})')
        if not catalogo:
            pre.append('pre-condicao ausente: catalogo de analitos vazio (Especificações A4:A43)')
        if not cad:
            pre.append('pre-condicao ausente: cadastro de lotes vazio (regLoteCol)')
        if not bl0:
            pre.append('pre-condicao ausente: nenhum bloco por lote na aba Analitos (ADR-074)')
        try:
            q.nome('currentUser')
        except Exception:
            pre.append('pre-condicao ausente: nome currentUser')
        try:
            rows = q.ler('tblCQ_Final')
        except Exception as e:
            rows = []
            pre.append(f'pre-condicao ausente: tblCQ_Final ilegivel ({repr(e)[:160]})')
        if not rows:
            pre.append('pre-condicao ausente: tblCQ_Final vazia')
        ev.update({'analitos_catalogo': len(catalogo), 'lotes_cadastro': len(cad), 'blocos_por_lote': len(bl0),
                   'linhas_tblCQ_Final': len(rows),
                   'nlv': nlv, 'problemas': pre})
        reg('A00 Pre-condicoes: marca, motor sem DB_SEAC, abas, catalogo, cadastro, currentUser, tblCQ_Final', not pre, ev)
        rel['precondicoes'] = ev
        if pre:
            raise Abortar('pre-condicoes')

        # ------------------------------------------------------------ calculo
        cat = collections.OrderedDict((chave(x), x) for x in catalogo)
        presentes = {chave(r.get('LOTE')) for r in rows}
        lotes = collections.OrderedDict((chave(x), x) for x in cad if chave(x) in presentes)
        ign = collections.Counter()
        calc = calcular_alvos(rows, nlv, a.n, a.minimo, set(lotes), set(cat), ign)

        def ler_store():
            ls = q.ws('LotesStore')
            ult = ls.Cells(ls.Rows.Count, 16).End(-4162).Row
            d = {}
            if ult >= 2:
                for i, linha in enumerate(ls.Range(f'A2:Z{ult}').Value, 2):
                    if txt(linha[15]):
                        d.setdefault(chave(linha[15]), (i, tuple(linha)))
            return d

        st0 = ler_store()
        antes, pares, manuais = selecionar_pares(lotes, cat, st0, calc, nlv)
        rel['pulados_origem_manual'] = {'linhas_lote_analito': len(manuais), 'exemplos': manuais[:20]}
        print(f'ADR-102: {len(manuais)} linha(s) lote|analito com origem MANUAL puladas (alvo do analista mantido)',
              flush=True)
        cols, linhas_csv, equips = montar_linhas(calc, antes, pares, nlv, a.minimo)
        for lin in linhas_csv:                                  # nomes como no cadastro/catalogo
            lin['lote'], lin['analito'] = lotes[lin['lote']], cat[lin['analito']]
        escrever_csv(csv_tmp, cols, linhas_csv)
        cont = collections.Counter(l['status'] for l in linhas_csv)
        rel.update({'csv': csv_tmp, 'json': json_tmp, 'lotes_considerados': list(lotes.values()),
                    'lotes_do_cadastro_sem_resultado': [x for x in cad if chave(x) not in presentes],
                    'pares_por_status': dict(cont), 'equipamentos': equips, 'linhas_ignoradas_no_calculo': dict(ign),
                    'marca_dif_equip_maior_1dp': sum(1 for l in linhas_csv if l['marca_dif_equip_maior_1dp'] == 'SIM')})
        print(f'CALCULO: {dict(cont)} -> {csv_tmp}', flush=True)

        # ------------------------------------------------------------ Westgard antes
        def contar_wg():
            w = q.ws('Westgard_Fotos')
            ult = w.Cells(w.Rows.Count, 1).End(-4162).Row
            c = collections.Counter()
            if ult >= 4:
                for t, lt, eq in w.Range(f'A4:C{ult}').Value:
                    if txt(t).upper() == 'EVENTO':
                        c[f'{txt(lt)}|{txt(eq)}'] += 1
            return c
        wg0 = contar_wg()

        tocar = collections.OrderedDict()                       # (lt, an) -> [(nv, status, res, rot, n_lote)]
        for k in pares:
            tocar.setdefault((k[0], k[1]), [])
        for (lt, an) in list(tocar):
            niv = []
            for nv in range(1, nlv + 1):
                res = calc.get((lt, an, nv))
                m0, s0, rot = antes[(lt, an, nv)]
                niv.append((nv, status_par(res, m0, s0), res, rot, res['n_lote'] if res else 0))
            if any(x[1] in ('aplicado', 'igual') for x in niv):
                tocar[(lt, an)] = niv
            else:
                del tocar[(lt, an)]
        rel['linhas_lote_analito_a_gravar'] = len(tocar)

        if a.simular:
            reg('A01 Simulacao: CSV gravado, nada gravado na planilha', os.path.isfile(csv_tmp),
                {'csv': csv_tmp, 'pares_por_status': dict(cont), 'linhas_a_gravar': len(tocar)})
            rel['westgard_eventos_antes'] = dict(wg0)
            raise Abortar('simulacao')

        # ------------------------------------------------------------ gravacao
        def definir_usuario(valor):
            rg = q.nome('currentUser')
            ws = rg.Worksheet
            prot = bool(ws.ProtectContents)
            if prot:
                ws.Unprotect(SENHA)
            rg.Value = valor
            if prot:
                ws.Protect(SENHA, False, True, True, True)       # o mesmo de mSeguranca.RestaurarProtecao
            return txt(rg.Value)

        aud = q.ws('Audit_Log')
        # a ultima linha COM EVENTO pela propria VBA: End(xlUp) para na linha fantasma do ListObject tblAuditoria
        # (mAuditoria.UltimaLinhaAudit) e as linhas novas ficariam fora da conferencia A03
        aud0 = int(q.run('mAuditoria.UltimaLinhaAudit', teto=60))
        xl.EnableEvents = False
        xl.Calculation = CALC_MANUAL                     # 2 x 1.200 gravacoes sem recalcular a pasta a cada uma
        usuario = definir_usuario(MODULO)
        parecer = f'Alvo do laboratorio pelos {a.n} primeiros resultados do CIQ (pedido do usuario de {PEDIDO})'
        n_alt, n_cstr, lotes_tocados, origens = 0, 0, collections.OrderedDict(), {}
        por_lote = collections.defaultdict(list)
        for (lt, an), niv in tocar.items():
            por_lote[lt].append((an, niv))
        ls = q.ws('LotesStore')
        for lt, itens in por_lote.items():
            lote = lotes[lt]
            resumo = collections.Counter()
            for an, niv in itens:
                nome_an = cat[an]
                esp = q.run('mLotesBlocos.EspecDoLote', lote, nome_an, teto=60)
                origem = texto_origem(hoje, niv, a.minimo)
                origens[(lt, an)] = origem
                for nv, st, res, rot, n_lote in niv:
                    resumo[st] += 1
                    if st not in ('aplicado', 'igual'):
                        continue
                    for j, campo, valor in ((2 * nv - 1, 'Media', res['media']), (2 * nv, 'DP', res['dp'])):
                        ant = q.run('mLotes.GravarParametroStore', lote, nome_an, j, valor,
                                    esp if esp is not None else E, origem, teto=120)
                        if not (num(ant) and float(ant) == valor):
                            q.run('mAuditoria.Auditar', 'CONFIGURACAO', 'PARAMETRO_LOTE_ALTERADO', MODULO, 0, E, '',
                                  lote, nv, nome_an, ant if ant not in (None, '') else E, valor, '', '',
                                  f'{campo} N{nv}', parecer, teto=60)
                            n_alt += 1
            # armadilha do CStr igual: W e o bit TESTE do nivel recalculado conferidos e corrigidos a mao
            st1 = ler_store()
            fix = []
            for an, niv in itens:
                i, linha = st1.get(f'{lt}|{an}', (0, None))
                if not i:
                    continue
                if txt(linha[22]) != origens[(lt, an)]:
                    fix.append((i, 23, origens[(lt, an)]))
                for nv, st, *_ in niv:
                    if st in ('aplicado', 'igual') and (bit_val(linha[23 + nv - 1]) > 0 or txt(linha[23 + nv - 1])):
                        fix.append((i, 24 + nv - 1, None))
            if fix:
                prot = bool(ls.ProtectContents)
                if prot:
                    ls.Unprotect(SENHA)
                for i, c, v in fix:
                    if v is None:
                        ls.Cells(i, c).ClearContents()
                    else:
                        ls.Cells(i, c).Value = v
                if prot:
                    ls.Protect(SENHA, False, True, True, True)
                q.run('mLotes.InvalidarLotes', teto=60)
                n_cstr += len(fix)
            ng = resumo['aplicado'] + resumo['igual']
            q.run('mAuditoria.Auditar', 'CONFIGURACAO', 'PARAMETRO_LOTE_CALCULADO', MODULO, 0, E, '', lote, 0, '', '',
                  f'{ng} par(es) analito/nivel', '', '', 'Media/DP',
                  f'Media/DP do lote {lote} = media/DP amostral dos {a.n} primeiros resultados Ativos do CIQ (todos os '
                  f'equipamentos; minimo {a.minimo}): {resumo["aplicado"]} par(es) alterado(s), {resumo["igual"]} '
                  f'igual(is) ao anterior, {resumo["insuficiente"] + resumo["dp_zero"]} mantido(s) nas linhas tocadas. '
                  f'Pedido do usuario de {PEDIDO}.', teto=60)
            q.run('mLotesBlocos.EscreverValoresDoLote', lote, teto=120)
            lotes_tocados[lote] = dict(resumo)
            print(f'  lote {lote}: {dict(resumo)}', flush=True)
        wb.Names.Add(MARCA, f'="{hoje} n{a.n} min{a.minimo}"')
        q.run('mLotes.InvalidarLotes', teto=60)
        lote_tela = txt(q.run('mLotes.LoteEmTela', teto=60))
        if lote_tela:
            q.run('mLotes.CarregarParametrosView', lote_tela, teto=300)
        q.run('mEstatistica.InvalidarParametros', teto=60)
        q.run('mLotes.AtualizarResumoLote', teto=60)
        xl.Calculation = CALC_AUTO
        xl.Calculate()
        rel.update({'campos_alterados_auditados': n_alt, 'correcoes_cstr_igual': n_cstr, 'lotes_tocados': lotes_tocados,
                    'lote_em_tela': lote_tela})

        # ------------------------------------------------------------ conferencias
        st2 = ler_store()
        div, n_exatos = [], 0
        tocadas = {f'{lt}|{an}' for lt, an in tocar}
        for (lt, an), niv in tocar.items():
            i, linha = st2.get(f'{lt}|{an}', (0, None))
            if not i:
                div.append((lt, an, 'linha ausente'))
                continue
            if txt(linha[22]) != origens[(lt, an)]:
                div.append((lt, an, 'W', txt(linha[22])[:80]))
            for nv, st, res, rot, _ in niv:
                m, s = linha[2 + 2 * (nv - 1)], linha[3 + 2 * (nv - 1)]
                if st in ('aplicado', 'igual'):
                    # tolerancia 1e-13 relativa: so absorve um eventual arredondamento do Excel na 15a casa;
                    # quantos bateram bit a bit vai na evidencia
                    if not (perto(m, res['media'], 1e-13) and perto(s, res['dp'], 1e-13)):
                        div.append((lt, an, nv, 'valor', m, s, res['media'], res['dp']))
                    elif float(m) == res['media'] and float(s) == res['dp']:
                        n_exatos += 1
                    if txt(linha[23 + nv - 1]):
                        div.append((lt, an, nv, 'bit TESTE', linha[23 + nv - 1]))
                else:
                    m0, s0, _r = antes[(lt, an, nv)]
                    if txt(m) != txt(m0) or txt(s) != txt(s0):
                        div.append((lt, an, nv, 'mantido mudou', m0, s0, m, s))
        outras = [k for k in set(st0) | set(st2) if k not in tocadas
                  and (st0.get(k, (0, None))[1] != st2.get(k, (0, None))[1])]
        reg('A02 LotesStore: cada par recalculado = media/DP calculados (exato), bit TESTE zerado, W = origem nova com '
            'todos os niveis; niveis mantidos sem mudanca; nenhuma outra linha da store mudou',
            bool(tocar) and not div and not outras,
            {'linhas_tocadas': len(tocar), 'divergencias': div[:8], 'outras_linhas_mudaram': outras[:8],
             'correcoes_cstr_igual': n_cstr, 'pares_iguais_bit_a_bit': n_exatos})

        ua = int(q.run('mAuditoria.UltimaLinhaAudit', teto=60))
        novas = aud.Range(f'A{aud0 + 1}:Y{ua}').Value if ua > aud0 else ()
        alt = [r for r in novas if txt(r[6]) == 'PARAMETRO_LOTE_ALTERADO' and txt(r[7]) == MODULO]
        clc = [r for r in novas if txt(r[6]) == 'PARAMETRO_LOTE_CALCULADO' and txt(r[7]) == MODULO]
        quem = {txt(r[24]) for r in alt + clc}
        reg('A03 Audit_Log: uma linha PARAMETRO_LOTE_ALTERADO por campo que mudou e uma PARAMETRO_LOTE_CALCULADO por '
            'lote, autoria aplicar_alvos_ciq',
            usuario == MODULO and len(alt) == n_alt and len(clc) == len(lotes_tocados) and quem <= {MODULO},
            {'currentUser': usuario, 'alteradas': [len(alt), n_alt], 'calculadas': [len(clc), len(lotes_tocados)],
             'usuarios': sorted(quem)})

        b = q.ws('Analitos')
        bl = {chave(k): h for k, h in blocos(b)}
        difb, sem_bloco = [], []
        for lote in lotes_tocados:
            h = bl.get(chave(lote))
            if not h:
                sem_bloco.append(lote)          # invariante do BA (todo lote do cadastro tem bloco): falta = FAIL
                continue
            for linha in b.Range(b.Cells(h + O_AN, 1), b.Cells(h + O_AN + NAN - 1, 10)).Value:
                an = chave(linha[0])
                if not an:
                    continue
                sv = st2.get(f'{chave(lote)}|{an}', (0, (None,) * 26))[1][2:8]
                for j in range(6):
                    v, s = linha[4 + j], sv[j]
                    if (num(s) and not perto(v, s)) or (not num(s) and v not in ('', None)):
                        difb.append((lote, an, j + 1, v, s))
        conf = txt(q.run('mLotes.ConferirLotesStore', teto=60))
        reg('A04 Blocos da aba Analitos (E:J) dos lotes tocados = LotesStore; ConferirLotesStore ok',
            bool(lotes_tocados) and not sem_bloco and not difb and conf.startswith('ok'),
            {'lotes_tocados': len(lotes_tocados), 'lotes_sem_bloco': sem_bloco[:8],
             'divergencias': difb[:8], 'ConferirLotesStore': conf[:200]})

        amostra = [(lt, an, nv, res) for (lt, an), niv in tocar.items() for nv, st, res, *_ in niv
                   if st in ('aplicado', 'igual')]
        passo = max(1, len(amostra) // 40)
        divv = []
        for lt, an, nv, res in amostra[::passo]:
            mv = q.run('mLotes.MediaDoLote', lotes[lt], cat[an], nv, teto=60)
            sv = q.run('mLotes.DPDoLote', lotes[lt], cat[an], nv, teto=60)
            tv = q.run('mLotes.AlvoDeTeste', lotes[lt], cat[an], nv, teto=60)
            if not (perto(mv, res['media'], 1e-12) and perto(sv, res['dp'], 1e-12)) or tv:
                divv.append((lotes[lt], cat[an], nv, mv, sv, tv))
        reg('A05 Leitura pela VBA (mLotes.MediaDoLote/DPDoLote/AlvoDeTeste) de uma amostra dos pares gravados = alvo '
            'novo, sem marca TESTE', bool(amostra) and not divv,
            {'conferidos': len(amostra[::passo]), 'gravados': len(amostra), 'divergencias': divv[:6]})

        try:
            marca = wb.Names(MARCA).RefersTo
        except Exception:
            marca = None
        reg('A06 Marca blAlvosCIQ gravada (string entre aspas)', bool(marca) and '"' in str(marca), {'marca': marca})

        # ------------------------------------------------------------ ATUALIZAR so das entradas (Westgard)
        if falhas():
            raise Abortar('conferencias antes do ATUALIZAR')
        xl.EnableEvents = True
        t0 = datetime.datetime.now()
        r = txt(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
        seg = (datetime.datetime.now() - t0).total_seconds()
        # AtualizarDadosCore ENGOLE o erro do Westgard (On Error Resume Next em AuditarFotosWestgard) e devolve 'OK|'
        # mesmo assim; o retorno da foto vai no fim do TEMPOS: '... westgard_fotos 1,23s [RECALC=k/N MUDANCAS=m]'.
        # 'OK|' sozinho seria falso PASS: exige RECALC sem FALHAS e, com campo alterado, k > 0 (a impressao digital
        # inclui Media/DP, mEstatistica.ImpressoesDigitaisWg)
        wgres, k_recalc = retorno_foto_wg(r)
        ok_wg = k_recalc is not None and 'FALHAS=' not in wgres and 'ERRO' not in wgres and (n_alt == 0 or k_recalc > 0)
        reg('A07 mIntegracao.AtualizarSoEntradas("") = OK e a foto do Westgard foi refeita sem falha '
            '(RECALC k/N, k > 0 com alvo alterado)', r.startswith('OK|') and ok_wg,
            {'westgard_fotos': wgres[:300], 'escopos_refeitos': k_recalc, 'campos_alterados': n_alt,
             'resumo': r[:300], 'segundos': round(seg)})
        rel.update({'atualizar_so_entradas': r[:2000], 'atualizar_segundos': round(seg)})
        wg1 = contar_wg()
        chaves_wg = sorted(set(wg0) | set(wg1))
        rel['westgard_eventos'] = {k: [wg0.get(k, 0), wg1.get(k, 0)] for k in chaves_wg}
        rel['westgard_eventos_total'] = [sum(wg0.values()), sum(wg1.values())]
        pl = collections.defaultdict(lambda: [0, 0])
        for k, (x0, x1) in rel['westgard_eventos'].items():
            pl[k.split('|')[0]][0] += x0
            pl[k.split('|')[0]][1] += x1
        rel['westgard_eventos_por_lote'] = dict(pl)
        reg('A08 Westgard_Fotos recontada por lote x equipamento (informativo: eventos antes -> depois)', True,
            {'total': rel['westgard_eventos_total'],
             'maiores_quedas': sorted(((wg0.get(k, 0) - wg1.get(k, 0), k) for k in chaves_wg), reverse=True)[:6]})
        aud_fim = int(q.run('mAuditoria.UltimaLinhaAudit', teto=60))
        rel['audit_log_linhas_novas'] = aud_fim - aud0

        # o alvo foi calculado sobre a tblCQ_Final ANTES do ATUALIZAR: se havia entrada pendente (inativacao, manual,
        # substituicao) ainda nao aplicada, o reprocessamento muda os Ativos e o alvo gravado nao corresponde mais
        # aos 'n primeiros' que o motor ve. Recalcula sobre a tabela de agora e exige o mesmo alvo em todo par gravado.
        rows2 = q.ler('tblCQ_Final')
        calc2 = calcular_alvos(rows2, nlv, a.n, a.minimo, set(lotes), set(cat))
        dif2 = []
        for (lt, an), niv in tocar.items():
            for nv, st, res, *_ in niv:
                if st not in ('aplicado', 'igual'):
                    continue
                r2 = calc2.get((lt, an, nv))
                if not r2 or r2['status'] != 'calculado' or r2['ids_usados'] != res['ids_usados'] \
                        or not (perto(r2['media'], res['media'], 1e-12) and perto(r2['dp'], res['dp'], 1e-12)):
                    dif2.append((lotes[lt], cat[an], nv, res['media'], res['dp'],
                                 r2['media'] if r2 else None, r2['dp'] if r2 else None))
        mudou_status = sum(1 for k in set(calc) | set(calc2)
                           if (calc.get(k) or {}).get('status') != (calc2.get(k) or {}).get('status'))
        reg('A08b tblCQ_Final depois do ATUALIZAR: os mesmos n primeiros e o mesmo alvo em todo par gravado (senao '
            'havia entrada pendente: rode o ATUALIZAR na producao antes e aplique de novo)', not dif2,
            {'linhas': [len(rows), len(rows2)], 'pares_divergentes': len(dif2), 'exemplos': dif2[:6],
             'pares_que_mudaram_de_status': mudou_status})

        # ------------------------------------------------------------ salvar a copia
        if falhas():
            raise Abortar('falha antes de salvar')
        sem_usuario = definir_usuario('')
        reg('A09 currentUser vazio antes de salvar (a identidade da sessao nao vai para o disco)', sem_usuario == '',
            {'currentUser': sem_usuario})
        if falhas():
            raise Abortar('currentUser')
        xl.EnableEvents = False
        wb.Save()
        salvou = True
    except Abortar as e:
        print(f'interrompido: {e}', flush=True)
    except Exception as e:
        reg('AER Erro inesperado durante a aplicacao (nada foi salvo)', False,
            {'erro': repr(e)[:400], 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.ex.fechar()

    rel['F00'] = h1 = sha256_arquivo(entrada)
    reg('F00 sha256 da entrada no fim = inicio (a entrada nunca foi aberta nem gravada)', h1 == h0,
        {'inicio': h0, 'fim': h1})
    if not salvou:
        gravar_json()
        print(f'{"SIMULACAO" if a.simular else "NADA APLICADO"}: CSV {csv_tmp} | JSON {json_tmp}', flush=True)
        return 1 if (falhas() or not a.simular) else 0

    erro_b = None
    try:                                    # blindar levanta SystemExit (copia somente leitura): o JSON sai mesmo assim
        blindar_entrega.blindar(copia)
        ok_b = bool(blindar_entrega.verificar(copia))
    except (Exception, SystemExit) as e:
        ok_b, erro_b = False, repr(e)[:300]
    reg('A10 Copia aplicada blindada (blindar_entrega.verificar)', ok_b, {'copia': copia, 'erro': erro_b})
    trava = em_uso(entrada)
    if falhas() or trava or sha256_arquivo(entrada) != h0:
        reg('A11 Troca', False, {'erro': 'falha, entrada em uso ou entrada mudou durante a aplicacao; nada foi trocado',
                                 'arquivo_de_trava': trava, 'copia': copia})
        gravar_json()
        return 1
    hc = sha256_arquivo(copia)
    if saida:
        destino_rel = os.path.dirname(saida)
        os.makedirs(destino_rel, exist_ok=True)
        shutil.copy2(copia, saida)
        hs = sha256_arquivo(saida)
        reg('A11 Copia aplicada gravada em --so-copia; a entrada nao foi trocada', hs == hc,
            {'saida': saida, 'sha256': hs})
    else:
        destino_rel = os.path.join(os.path.dirname(entrada), f'_backup_pre_alvos_{stamp}')
        os.makedirs(destino_rel)
        shutil.copy2(entrada, os.path.join(destino_rel, os.path.basename(entrada)))
        shutil.copy2(copia, entrada)
        hs = sha256_arquivo(entrada)
        reg('A11 Producao trocada pela copia aplicada (backup antes) e sha conferido', hs == hc,
            {'producao': entrada, 'sha_antes': h0, 'sha_depois': hs, 'backup': destino_rel})
        rel['backup'] = destino_rel
    rel['sha_resultado'] = hs
    gravar_json()
    for p in (csv_tmp, json_tmp):
        shutil.copy2(p, os.path.join(destino_rel, os.path.basename(p)))
    print(f'APLICADO {produto} ({modo}): sha {h0[:16]} -> {hs[:16]} | CSV/JSON em {destino_rel}', flush=True)
    return 1 if falhas() else 0


# =============================================================== autoteste (sem Excel)
def autoteste():
    import statistics
    falhou = []

    def ok(cod, cond, ev=''):
        print(('PASS ' if cond else 'FAIL ') + cod + (f' -- {ev}' if ev else ''), flush=True)
        if not cond:
            falhou.append(cod)

    t0 = datetime.datetime(2026, 10, 1, 8, 0)

    def lin(lote, an, nv, run, x, eq='XN1000', part='SIM', h=0, idr=None):
        return {'LOTE': lote, 'ANALITO': an, 'NIVEL': nv, 'RUN': run, 'RESULTADO': x, 'PARTICIPA_ESTATISTICA': part,
                'EQUIPAMENTO': eq, 'DATA_HORA': t0 + datetime.timedelta(days=int(run) % 1000, hours=h),
                'ID_REGISTRO': idr if idr is not None else f'ID{run}-{h}'}

    # T01 ordem por RUN: 45 resultados embaralhados (valor = posicao), os 40 primeiros sao RUN 1..40
    import random
    rnd = random.Random(7)
    rs = [lin('8974', 'GLI', 1, 26110100 + i, float(i)) for i in range(1, 46)]
    rnd.shuffle(rs)
    c = calcular_alvos(rs, 2)
    r = c[('8974', 'GLI', 1)]
    m_esp, s_esp = statistics.mean(range(1, 41)), statistics.stdev(range(1, 41))
    ok('T01 ordem por RUN, 40 primeiros de 45; media/DP amostral conhecidos (20,5 / 11,6905)',
       r['status'] == 'calculado' and r['n_usado'] == 40 and r['n_lote'] == 45 and abs(r['media'] - 20.5) < 1e-12
       and abs(r['dp'] - s_esp) < 1e-12 and abs(m_esp - 20.5) < 1e-12 and r['run_ini'] == '26110101'
       and r['run_fim'] == '26110140', (r['media'], r['dp'], s_esp))
    ok('T02 periodo = datas do 1o e do ultimo usado; DP do periodo inteiro = stdev dos 45',
       r['data_ini'] == t0 + datetime.timedelta(days=101) and r['data_fim'] == t0 + datetime.timedelta(days=140)
       and abs(r['dp_periodo'] - statistics.stdev(range(1, 46))) < 1e-12
       and abs(r['media_periodo'] - 23.0) < 1e-12, (r['data_ini'], r['data_fim'], r['dp_periodo']))

    # T03 20..39 usa todos; T04 < 20 insuficiente; T05 exatamente 20
    vals25 = [5.0 + 0.1 * ((i * 7) % 11) for i in range(25)]
    c = calcular_alvos([lin('A1', 'CREA', 2, 100 + i, v) for i, v in enumerate(vals25)], 2)
    r = c[('A1', 'CREA', 2)]
    ok('T03 25 resultados (20..39): usa todos', r['status'] == 'calculado' and r['n_usado'] == 25
       and abs(r['media'] - statistics.mean(vals25)) < 1e-12 and abs(r['dp'] - statistics.stdev(vals25)) < 1e-12)
    c = calcular_alvos([lin('A1', 'CREA', 1, 100 + i, float(i)) for i in range(19)], 2)
    r = c[('A1', 'CREA', 1)]
    ok('T04 19 resultados (< 20): insuficiente, sem media/DP', r['status'] == 'insuficiente' and r['media'] is None
       and r['dp'] is None and r['n_usado'] == 19)
    c = calcular_alvos([lin('A1', 'CREA', 1, 100 + i, float(i)) for i in range(20)], 2)
    ok('T05 exatamente 20: calculado', c[('A1', 'CREA', 1)]['status'] == 'calculado')

    # T06 empate de RUN na fronteira: RUN 40 tem 2 resultados; desempata por DATA_HORA, depois por ID_REGISTRO
    base = [lin('L', 'HGB', 1, 1000 + i, 10.0) for i in range(1, 40)]          # 39 iguais a 10
    base += [lin('L', 'HGB', 1, 1040, 999.0, h=5), lin('L', 'HGB', 1, 1040, 20.0, h=1)]
    c = calcular_alvos(list(reversed(base)), 3)
    r = c[('L', 'HGB', 1)]
    ok('T06 empate de RUN: entra o de DATA_HORA menor (20, nao 999)', r['n_usado'] == 40
       and abs(r['media'] - (39 * 10 + 20) / 40) < 1e-12, r['media'])
    base = [lin('L', 'HGB', 1, 1000 + i, 10.0) for i in range(1, 40)]
    base += [lin('L', 'HGB', 1, 1040, 999.0, h=1, idr=10), lin('L', 'HGB', 1, 1040, 20.0, h=1, idr=9)]
    c = calcular_alvos(base, 3)
    ok('T07 empate de RUN e DATA_HORA: desempata pelo ID_REGISTRO numerico (9 antes de 10)',
       abs(c[('L', 'HGB', 1)]['media'] - (39 * 10 + 20) / 40) < 1e-12)

    # T08 filtros: so SIM, lote/analito normalizados, nivel > nlv fora, texto fora, todos os equipamentos juntos
    rs = [lin(8974.0, ' gli ', 1, 200 + i, 100.0 + i, eq='DIMENSION 1' if i % 2 else 'DIMENSION 2') for i in range(30)]
    rs += [lin('8974', 'GLI', 1, 150 + i, 1e6, part='NÃO') for i in range(10)]
    rs += [lin('8974', 'GLI', 3, 300 + i, 1.0) for i in range(30)]
    rs += [lin('8974', 'GLI', 1, 400, 'texto'), lin('9999', 'GLI', 1, 401, 5.0), lin('8974', 'URE', 1, 402, 5.0)]
    ign = collections.Counter()
    c = calcular_alvos(rs, 2, lotes={'8974'}, analitos={'GLI'}, ignorados=ign)
    r = c.get(('8974', 'GLI', 1))
    ok('T08 so PARTICIPA=SIM; lote 8974.0 = "8974"; " gli " = GLI; N3 fora na Bio; texto, lote e analito fora',
       r is not None and r['n_lote'] == 30 and set(c) == {('8974', 'GLI', 1)} and ign['nivel_fora'] == 30
       and ign['resultado_nao_numerico'] == 1 and ign['lote_fora_do_cadastro'] == 1
       and ign['analito_fora_do_catalogo'] == 1, dict(ign))
    pe = r['por_equip']
    ok('T09 por equipamento: n e media de cada DIMENSION', pe['DIMENSION 1']['n_lote'] == 15
       and pe['DIMENSION 2']['n_lote'] == 15 and abs(pe['DIMENSION 1']['media_lote'] - statistics.mean(
           [100.0 + i for i in range(30) if i % 2])) < 1e-12, {k: (v['n_lote'], v['media_lote']) for k, v in pe.items()})

    # T10 DP zero (todos iguais) nao e gravado
    c = calcular_alvos([lin('Z', 'NA', 1, 500 + i, 140.0) for i in range(25)], 2)
    ok('T10 todos iguais: dp_zero, nao grava', c[('Z', 'NA', 1)]['status'] == 'dp_zero')

    # T11 status do par: TESTE com media igual -> 'igual'; diferente -> 'aplicado'
    c = calcular_alvos([lin('T', 'K', 1, 600 + i, float(i % 5)) for i in range(30)], 2)
    r = c[('T', 'K', 1)]
    ok('T11 status_par: igual (mesmo double) / aplicado / insuficiente',
       status_par(r, r['media'], r['dp']) == 'igual' and status_par(r, r['media'], r['dp'] * 1.5) == 'aplicado'
       and status_par(None, 1, 1) == 'insuficiente')

    # T12 W: todos os niveis, n e periodo legiveis pelo BO; mantido sem 'n='
    niv = [(1, 'aplicado', r, 'TESTE', 30), (2, 'igual', r, 'digitado', 30),
           (3, 'insuficiente', None, 'TESTE', 12)]
    w = texto_origem('22/11/2026', niv, 20)
    lidos = RX_W_NIVEL.findall(w)
    ok('T12 W cita todos os niveis: N1/N2 n=30 (periodo), N3 mantido (TESTE; 12 < 20)',
       RX_W_CIQ.match(w) is not None and [x[:2] for x in lidos] == [('1', '30'), ('2', '30')]
       and RX_W_MANTIDO.findall(w) == ['3'] and 'N3: mantido (TESTE; 12 resultado(s) < 20)' in w, w)

    # T13 linhas do CSV: deslocamento, razao, % fora, marca entre equipamentos
    rs = [lin('B', 'COL', 1, 700 + i, 100.0 + (i % 2) * 10, eq='D1' if i % 2 == 0 else 'D2') for i in range(40)]
    c = calcular_alvos(rs, 2)
    rr = c[('B', 'COL', 1)]
    ant = {('B', 'COL', 1): (90.0, 2.0, 'bula/outro'), ('B', 'COL', 2): (50.0, 1.0, 'TESTE')}
    cols, linhas, equips = montar_linhas(c, ant, [('B', 'COL', 1), ('B', 'COL', 2)], 2, 20)
    l1, l2 = linhas
    ok('T13 CSV: deslocamento = (105-90)/2, razao = dp/2, % fora antes 100 (alvo deslocado), marca D1/D2 > 1 DP',
       l1['status'] == 'aplicado' and abs(l1['deslocamento_em_dp_antes'] - 7.5) < 1e-12
       and abs(l1['razao_dp_nova_antes'] - rr['dp'] / 2) < 1e-12 and l1['pct_fora_2dp_antes'] == 100.0
       and l1['pct_fora_2dp_depois'] == 0.0 and l1['marca_dif_equip_maior_1dp'] == 'SIM'
       and equips == ['D1', 'D2'] and l1['n_usado_D1'] == 20, {k: l1[k] for k in cols[:22]})
    ok('T14 CSV: par sem resultado e com TESTE = insuficiente, depois = antes', l2['status'] == 'insuficiente'
       and l2['n_lote'] == 0 and l2['media_nova'] is None and l2['origem_antes'] == 'TESTE')
    ok('T15 classificar_origem: TESTE pelo bit, vazio, CIQ, digitado, bula',
       classificar_origem('x', 3, 1, 1) == 'TESTE' and classificar_origem('', None, None, '') == 'vazio'
       and classificar_origem('CIQ: 20 primeiros', None, 1, 1) == 'CIQ'
       and classificar_origem('digitado na aba Analitos (bloco do lote)', 0, 1, 1) == 'digitado'
       and classificar_origem('', None, 1, 1) == 'bula/outro')
    ok('T16 fmt: virgula decimal, vazio para None, sem ; no texto', fmt(1.5) == '1,5' and fmt(None) == ''
       and fmt(3) == '3' and fmt('a;b') == 'a,b')
    try:
        calcular_alvos([], 2, n=10, minimo=20)
        ok('T17 minimo > n recusado', False)
    except ValueError:
        ok('T17 minimo > n recusado', True)
    ok('T18 bit TESTE lido como a VBA (Val(CStr)): numero, texto, vazio',
       bit_val(3) == 3 and bit_val(1.0) == 1 and bit_val('2') == 2 and bit_val(None) == 0 and bit_val('') == 0
       and classificar_origem('x', '1', 1, 1) == 'TESTE')
    r_ok = ('OK|Atualização concluída.\r\n\r\nTempo: sync 1,0s · motor 9,8s · westgard_fotos 2,35s '
            '[RECALC=7/31 MUDANCAS=412]')
    r_err = 'OK|Atualização concluída. Tempo: motor 9,8s · westgard_fotos 0,01s [ERRO|Subscript out of range]'
    ok('T19 retorno da foto do Westgard no TEMPOS: RECALC k lido; ERRO engolido pelo nucleo sem k (A07 falha)',
       retorno_foto_wg(r_ok) == ('RECALC=7/31 MUDANCAS=412', 7)
       and retorno_foto_wg(r_err) == ('ERRO|Subscript out of range', None) and retorno_foto_wg('OK|x')[1] is None)
    # T20 ADR-102: linha com origem MANUAL e pulada (todos os niveis), contada; as outras seguem
    st = {'L1|GLI': (2, (None, None, 10.0, 1.0, 20.0, 2.0) + (None,) * 16
                     + ('MANUAL 25/11/2026 10:00 ana: N1 n=25 (01/10/2026 a 31/10/2026) | N2 n=12 (...) TEMPORARIO',
                        None, None, None)),
          'L1|URE': (3, (None, None, 5.0, 0.5) + (None,) * 18 + ('digitado na aba Analitos', None, None, None))}
    calc_t = {('L1', 'GLI', 1): {}, ('L1', 'URE', 1): {}}
    ant_t, par_t, man_t = selecionar_pares(['L1'], ['GLI', 'URE'], st, calc_t, 2)
    ok('T20 origem MANUAL (ADR-102): a linha L1|GLI e pulada em todos os niveis e contada; L1|URE segue; '
       'EXCECAO TEMPORARIA nao e pulada (temporaria: deve ser substituida)',
       man_t == ['L1|GLI'] and all(k[1] != 'GLI' for k in par_t) and ('L1', 'URE', 1) in par_t
       and origem_manual(' manual x') and not origem_manual('CIQ 22/11/2026: ...') and not origem_manual(None)
       and not origem_manual('EXCECAO TEMPORARIA 25/11/2026 10:00 ana: N1 media=BULA; DP=BULA'),
       {'pulados': man_t, 'pares': par_t})
    total = 20
    print(f'AUTOTESTE: {total - len(falhou)}/{total} PASS' + (f' -- FAIL {falhou}' if falhou else ''), flush=True)
    return 1 if falhou else 0


if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == '--autoteste':
        sys.exit(autoteste())
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1:]))
