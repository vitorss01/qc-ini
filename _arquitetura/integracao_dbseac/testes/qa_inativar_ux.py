# -*- coding: utf-8 -*-
"""QA do ADR-081 (usabilidade) -- inativar pelo grafico em UM passo: duplo clique, motivo, confirmar; o grafico ja
volta atualizado e renumerado. Roda numa COPIA instalada (com o layout do Eng_Saida do ADR-081: BK..BM corrida do
lote do ponto ativo, BN..BV posicao de cada X).

Uso: python qa_inativar_ux.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

Pelo caminho real, por automacao (Application.Interactive = False): as funcoes que a janela usa
(mIntegracao.DadosDoPonto, InativarPorId, ReativarPorId, AtualizarSoEntradas), a permissao (mSeguranca.PodeInativar,
com um usuario de teste criado nesta copia e login pelo DoLogin) e a fachada mUI.InativarPontoUI com a resposta do
teste no lugar da janela (mUI.DefinirRespostaTeste). Nada e salvo no arquivo.

  U01 DadosDoPonto = a linha da tblCQ_Final (ID, analito, data/hora, valor, lote, estado) e corrida = Eng_Saida BK..
      em todos os niveis
  U02 InativarPorId com X + AtualizarSoEntradas: INATIVADO / X_VERMELHO / PARTICIPA = NAO; as corridas ativas do nivel
      renumeradas sem buraco (as anteriores iguais, as seguintes -1) e o X numa posicao fracionaria entre as vizinhas
  U03 sem X: NAO_PLOTAR, renumera, nenhum X do ID
  U04 motivo vazio (ou so espacos e linhas em branco): "Informe o motivo da inativacao." e nada gravado
  U05 ID ja inativado: recusa; linha com o analito errado: corrigida no lugar (sem linha nova)
  U06 ANALITO gravado = o da tblCQ_Final (ID de outro analito que nao o da tela)
  U07 ReativarPorId: linha limpa, ATIVO, numeracao EXATAMENTE a de antes (Eng_Saida BK..BM), Audit_Log com o motivo
  U08 sem permissao (sem login; papel abaixo de PAPEL_MIN_INATIVAR): a fachada recusa e nada e gravado
  U09 motivo com acentos e simbolos (>= sigma c-cedilha a-til) preservado na aba e na Principal - Resultados
  U10 InativarPontoUI com a resposta do teste: grava, atualiza, caixa "tip" com o visto; no X a mesma fachada reativa;
      CANCELAR nao grava; sem gancho e com Interactive = False: "ERRO|sem interface" (sem travar)
  U11 digitar o ID na aba Inativar preenche o ANALITO sozinho
  U12 abas protegidas antes = depois; mIntegracao.AtualizarDadosAutomatico continua "OK|"
  U13 tempo da InativarPontoUI (confirmar -> grafico atualizado) <= 18 s Bio / 25 s Hema
  U14 progresso visivel (mUI.UltimoAvisoProgresso: "aguarde" antes da atualizacao) e confirmacao com o lugar do
      ponto ("X vermelho entre as corridas k e k+1"; ao reativar "voltou a ser a corrida n")
  U15 clique repetido ate 2 s depois de terminar: "IGNORADO|...", nada gravado, gancho nao fica armado
  U00 seguranca do teste: arquivo de entrada intacto
"""
import datetime as dt
import collections
import hashlib
import json
import os
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, NAO, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

SIM = 'SIM'
TB = 'tblInativacao_NaoConformes'
CAB = ['ID_REGISTRO', 'ANALITO', 'REGISTRAR - LJ', 'MOTIVO', 'DATA_INATIVACAO', 'USUARIO']
# colunas do Eng_Saida (mEstatistica, iguais na Bio e na Hema; constantes VBA nao saem pelo Application.Run)
C_VALOR0, C_X0, C_ID0, C_IDX0, C_CORL0, C_XPOS0 = 25, 30, 39, 42, 63, 66
TETO_S = {'bio': 18.0, 'hema': 25.0}      # Bio 15 -> 18 s em 27/10/2026: ADR-083 (foto do Westgard, ~0,8 s com 1 lote
                                          # mudado) + folga para a maquina do laboratorio com o Excel agendado aberto
BLOQUEIO_S = 2.0                              # mUI.BLOQUEIO_SEG: cliques ignorados depois de terminar
USUARIO_QA = 'QA_INATIVAR_UX'
SENHA_QA = 'QaInativarUx#2026'                # valor de teste: o usuario so existe nesta copia, nunca salva
VISTO = '✓'
ABAS_PROT = ['Inativar', 'Painel', 'Principal - Resultados', 'Audit_Log', 'Eng_Saida', 'QA_INTEGRACAO']


def curto(i):
    s = str(i)
    p = s.rfind('-')
    return s[p + 1:] if p >= 0 and s[p + 1:].isdigit() else s


def num_pt(t):
    try:
        return float(str(t).replace('.', '').replace(',', '.')) if ',' in str(t) else float(t)
    except (TypeError, ValueError):
        return None


def sha(t):
    return hashlib.sha256(t.encode('latin-1')).hexdigest()


def num(v):
    return v if isinstance(v, float) else None          # erro de celula chega como int pelo COM


def txt(v):
    """Como o CStr do VBA mostra um codigo: 8353.0 -> '8353'."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return '' if v is None else str(v).strip()


# ---------------------------------------------------------------- leituras
class Leitor:
    def __init__(self, q):
        self.q, self.eng = q, q.ws('Eng_Saida')

    def n(self):
        return int(self.eng.Range('I1').Value or 0)

    def bloco(self, c0, ncol):
        n = self.n()
        if n == 0:
            return []
        v = self.eng.Range(self.eng.Cells(3, c0), self.eng.Cells(2 + n, c0 + ncol - 1)).Value
        if not isinstance(v, tuple):
            return [[v]]
        return [list(r) for r in v]

    def eng_por_run(self):
        """RUN -> dict(cor=[corrida por nivel], id=[ID do ponto por nivel], xid/xpos=[[3 X] por nivel]) na ordem
        das linhas. Chave = RUN (coluna B): a linha da corrida pode mudar se uma corrida sair da janela."""
        q, n = self.q, self.n()
        if n == 0:
            return collections.OrderedDict()
        runs = [r[0] for r in self.bloco(2, 1)]
        cor = self.bloco(C_CORL0, q.nlv)
        ids = self.bloco(C_ID0, q.nlv)
        xid = self.bloco(C_IDX0, 3 * q.nlv)
        xpos = self.bloco(C_XPOS0, 3 * q.nlv)
        out = collections.OrderedDict()
        for i, rn in enumerate(runs):
            if rn in (None, ''):
                continue
            out[int(rn)] = {'linha': i, 'cor': [num(c) for c in cor[i]], 'id': [str(x or '') for x in ids[i]],
                            'xid': [[str(xid[i][t * 3 + s] or '') for s in range(3)] for t in range(q.nlv)],
                            'xpos': [[num(xpos[i][t * 3 + s]) for s in range(3)] for t in range(q.nlv)]}
        return out

    def cq(self, ids):
        """Linhas da tblCQ_Final dos IDs pedidos (sem ler a tabela inteira coluna a coluna)."""
        lo = self.q.lo('tblCQ_Final')
        cab = [c.Name for c in lo.ListColumns]
        col_id = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
        pos = {}
        for k, v in enumerate(col_id, start=1):
            if v[0] in ids and v[0] not in pos:
                pos[v[0]] = k
        out = {}
        for i, k in pos.items():
            linha = lo.DataBodyRange.Rows(k).Value[0]
            out[i] = {c: (x.replace(tzinfo=None) if hasattr(x, 'tzinfo') and getattr(x, 'tzinfo', None) else x)
                      for c, x in zip(cab, linha)}
        return out

    def inativar(self):
        """Linhas preenchidas da aba Inativar: {numero da linha: {coluna: valor}}."""
        lo = self.q.lo(TB)
        v = lo.DataBodyRange.Value
        out = {}
        for k, linha in enumerate(v, start=1):
            if any(x not in (None, '') for x in linha):
                out[k] = dict(zip(CAB, linha))
        return out

    def audit_ult(self):
        al = self.q.ws('Audit_Log')
        return al.Cells(al.Rows.Count, 1).End(-4162).Row

    def audit_desde(self, r0):
        al = self.q.ws('Audit_Log')
        ult = self.audit_ult()
        out = []
        for r in range(r0 + 1, ult + 1):
            out.append({'acao': al.Cells(r, 7).Value, 'motivo': al.Cells(r, 23).Value, 'parecer': al.Cells(r, 24).Value})
        return out

    def protecao(self):
        out = {}
        for n in ABAS_PROT:
            try:
                out[n] = bool(self.q.ws(n).ProtectContents)
            except Exception:                       # noqa: BLE001
                out[n] = None
        return out


def ativos_em_ordem(por_run, t):
    return [(rn, d['cor'][t]) for rn, d in por_run.items() if d['cor'][t] is not None]


def consecutivos(lista):
    v = [c for _, c in lista]
    return all(abs(b - a - 1) < 1e-9 for a, b in zip(v, v[1:])) and all(abs(c - round(c)) < 1e-9 for c in v)


def renumerou(antes, depois, t, run_alvo):
    """As corridas ativas do nivel t depois de tirar run_alvo: anteriores iguais, seguintes -1, sem buraco."""
    a = ativos_em_ordem(antes, t)
    d = dict(ativos_em_ordem(depois, t))
    runs = [rn for rn, _ in a]
    if run_alvo not in runs:
        return False, {'erro': 'o run do ponto nao tinha corrida antes'}
    k = runs.index(run_alvo)
    ruins = []
    for j, (rn, c) in enumerate(a):
        if rn == run_alvo:
            if rn in d:
                ruins.append((rn, 'continua com corrida', d[rn]))
            continue
        esp = c if j < k else c - 1
        if rn not in d or abs(d[rn] - esp) > 1e-9:
            ruins.append((rn, esp, d.get(rn)))
    lista = ativos_em_ordem(depois, t)
    ok = not ruins and consecutivos(lista) and len(lista) == len(a) - 1
    return ok, {'n_antes': len(a), 'n_depois': len(lista), 'divergencias': ruins[:10],
                'primeira_e_ultima_depois': [lista[0][1], lista[-1][1]] if lista else None}


def x_do_id(por_run, t, idf):
    for rn, d in por_run.items():
        for s in range(3):
            if d['xid'][t][s] == idf:
                return rn, s, d['xpos'][t][s]
    return None, None, None


# ---------------------------------------------------------------- sessao / configuracao
def criar_usuario(q, login, papel):
    us = q.ws('Usuarios')
    us.Unprotect(ii.SENHA)
    i = next((i for i in range(4, 54) if str(us.Cells(i, 1).Value or '').strip().upper() == login.upper()), None)
    if i is None:
        i = next(i for i in range(4, 54) if not us.Cells(i, 1).Value)
    us.Cells(i, 1).Value = login
    us.Cells(i, 2).Value = 'QA ADR-081'
    us.Cells(i, 3).Value = papel
    us.Cells(i, 4).Value = sha(SENHA_QA)


def fazer_login(q, login):
    q.nome('loginUser').Value = login
    q.nome('loginPass').Value = SENHA_QA
    q.run('mSeguranca.DoLogin', teto=600)
    return str(q.nome('currentUser').Value or '')


def cfg_gravar(q, chave, valor):
    """Grava CHAVE = valor em tblConfigIntegracao (cria a linha se faltar). Devolve o valor anterior."""
    lo = q.lo('tblConfigIntegracao')
    q.ws('Cfg_Integracao').Unprotect(ii.SENHA)
    ev = q.ex.xl.EnableEvents
    q.ex.xl.EnableEvents = False
    try:
        ks = [str(r[0] or '').strip().upper() for r in lo.DataBodyRange.Value]
        if chave in ks:
            cel = lo.DataBodyRange.Cells(ks.index(chave) + 1, 2)
            antes = cel.Value
        else:
            lin = lo.ListRows.Add()
            lin.Range.Cells(1, 1).Value = chave
            cel = lin.Range.Cells(1, 2)
            antes = None
        cel.Value = valor
    finally:
        q.ex.xl.EnableEvents = ev
    return antes


def preparar_painel(q):
    """Analito com mais resultados no lote em analise, periodo = todo o dado (como o qa_inativar)."""
    wb = q.wb
    lo = q.lo('tblCQ_Final')
    cab = [c.Name for c in lo.ListColumns]
    v = lo.DataBodyRange.Value
    ix = {c: cab.index(c) for c in ('ANALITO', 'LOTE', 'PARTICIPA_ESTATISTICA', 'EQUIPAMENTO', 'DATA')}
    v_lote = wb.Application.Evaluate('loteAnalise')
    v_lote = v_lote.Value if hasattr(v_lote, 'Value') else v_lote
    lote_an = str(int(v_lote)) if isinstance(v_lote, float) else str(v_lote).strip()
    eq_pai = str(q.ws('Painel').Range('L4').Value) if q.bio else None
    linhas = [r for r in v if str(r[ix['LOTE']]) == lote_an and r[ix['PARTICIPA_ESTATISTICA']] == SIM
              and (eq_pai is None or r[ix['EQUIPAMENTO']] == eq_pai)]
    cont = collections.Counter(r[ix['ANALITO']] for r in linhas)
    cadastro = [str(x[0] or '').strip() for x in q.espec().Range('A4:A43').Value]
    an_top = next((a for a, _ in cont.most_common() if a in cadastro), None)
    dts = [r[ix['DATA']] for r in linhas if r[ix['ANALITO']] == an_top]
    if dts:
        q.ex.xl.EnableEvents = False
        pai = q.ws('Painel')
        pai.Unprotect(ii.SENHA)
        pai.Range('B3').Value = cadastro.index(an_top) + 1
        serie = lambda d: float((d.replace(tzinfo=None) - dt.datetime(1899, 12, 30)).days)   # noqa: E731
        pai.Range('G3').Value = serie(min(dts))
        pai.Range('G4').Value = serie(max(dts))
        q.ex.xl.EnableEvents = True
        q.run('mEstatistica.AtualizarEstatistica')
    return cadastro


def series(q, nivel):
    ch = q.ws('Painel').ChartObjects(nivel).Chart
    nomes = [str(s.Name) for s in ch.SeriesCollection()]
    s_res = nomes.index('Resultado') + 1 if 'Resultado' in nomes else None
    s_x = [k + 1 for k, n in enumerate(nomes) if 'conforme' in n.lower()]
    return s_res, s_x


def tip(q, nivel):
    try:
        return str(q.ws('Painel').ChartObjects(nivel).Chart.Shapes('tip').TextFrame2.TextRange.Text)
    except Exception:                               # noqa: BLE001
        return ''


# ---------------------------------------------------------------- a suite
def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    L = Leitor(q)
    tempos = {}

    def run(macro, *a, teto=300):
        return str(q.run(macro, *a, teto=teto))

    def atualizar_so_entradas(rotulo):
        t0 = time.time()
        r = run('mIntegracao.AtualizarSoEntradas', teto=900)
        tempos[rotulo] = round(time.time() - t0, 1)
        if not r.startswith('OK|'):
            raise RuntimeError(f'AtualizarSoEntradas ({rotulo}): {r[:400]}')
        return r

    try:
        prot0 = L.protecao()
        cadastro = preparar_painel(q)
        base = L.eng_por_run()
        analito_tela = str(q.ws('Eng_Saida').Range('C1').Value or '')
        # pontos normais do N1 com corrida, na ordem das linhas; os do meio tem vizinhas dos dois lados
        pts = [(rn, d['linha'], d['id'][0], d['cor'][0]) for rn, d in base.items()
               if d['cor'][0] is not None and d['id'][0]]
        if len(pts) < 14:
            raise RuntimeError(f'pré-condição ausente: menos de 14 pontos N1 com corrida no LJ ({analito_tela}): {len(pts)}')
        m = len(pts) // 2
        P2, P3, P10, P5, P11 = pts[m], pts[m - 3], pts[m + 3], pts[m - 6], pts[m + 6]
        cq0 = L.cq({p[2] for p in (P2, P3, P10, P5, P11)})
        s_res, s_x = series(q, 1)
        if s_res is None or not s_x:
            raise RuntimeError('pré-condição ausente: o gráfico N1 não tem a série "Resultado" ou os X')
        inat0 = L.inativar()
        aud0 = L.audit_ult()

        # ------------------------------------------------------------ U08a sem login: a fachada recusa
        pode0 = q.run('mSeguranca.PodeInativar', teto=60)
        run('mUI.DefinirRespostaTeste', 'CONFIRMAR|QA ADR-081 sem login|1', teto=60)
        r8a = run('mUI.InativarPontoUI', 1, s_res, P10[1] + 1, teto=120)
        inat8a = L.inativar()

        # ------------------------------------------------------------ U01 DadosDoPonto nos niveis
        u01 = []
        for t in range(q.nlv):
            s_t, _ = series(q, t + 1)
            cand = [(rn, d) for rn, d in base.items() if d['cor'][t] is not None and d['id'][t]]
            if not cand or s_t is None:
                u01.append({'nivel': t + 1, 'ok': False, 'erro': 'sem ponto com corrida ou sem serie Resultado'})
                continue
            rn, d = cand[len(cand) // 2]
            dados = run('mIntegracao.DadosDoPonto', t + 1, s_t, d['linha'] + 1, teto=60)
            p = dados.split('|')
            r = L.cq({d['id'][t]}).get(d['id'][t], {})
            ok = (len(p) == 9 and p[0] == 'OK' and p[1] == d['id'][t].upper() and p[2] == txt(r.get('ANALITO'))
                  and p[3] == str(t + 1) and p[4] == str(int(d['cor'][t]))
                  and r.get('DATA_HORA') is not None and p[5] == r['DATA_HORA'].strftime('%d/%m/%Y %H:%M')
                  and num_pt(p[6]) is not None and abs(num_pt(p[6]) - float(r.get('RESULTADO'))) < 1e-9
                  and p[7] == ('INATIVADO' if r.get('STATUS_ANALITICO') == 'INATIVADO' else 'ATIVO')
                  and p[8] == txt(r.get('LOTE')))
            u01.append({'nivel': t + 1, 'ok': ok, 'DadosDoPonto': dados, 'corrida_BK': d['cor'][t],
                        'cq': {k: r.get(k) for k in ('ID_REGISTRO', 'ANALITO', 'DATA_HORA', 'RESULTADO', 'LOTE',
                                                      'STATUS_ANALITICO')}})
        d_lim = run('mIntegracao.DadosDoPonto', 1, 1, 1, teto=60)
        reg('U01 DadosDoPonto = a linha da Principal - Resultados (ID, analito, data/hora, valor, lote, estado) e '
            'corrida = Eng_Saida BK.. em todos os níveis; linha de limite = ERRO (não é resultado)',
            all(x['ok'] for x in u01) and len(u01) == q.nlv and d_lim.startswith('ERRO|'),
            {'niveis': u01, 'linha_de_limite': d_lim})

        # ------------------------------------------------------------ U04 motivo vazio
        r4 = run('mIntegracao.InativarPorId', P2[2], '   ', True, str(int(P2[3])), teto=60)
        inat4 = L.inativar()
        reg('U04 Motivo vazio: InativarPorId devolve ERRO e nada é gravado na aba Inativar',
            r4 == 'ERRO|Informe o motivo da inativação.' and inat4 == inat0, {'retorno': r4, 'aba_igual': inat4 == inat0})

        # ------------------------------------------------------------ U02 inativar com X + renumeracao
        mot2 = 'QA ADR-081: repetição por erro analítico'
        r2 = run('mIntegracao.InativarPorId', P2[2], mot2, True, str(int(P2[3])), teto=60)
        atualizar_so_entradas('U02')
        depois2 = L.eng_por_run()
        f2 = L.cq({P2[2]}).get(P2[2], {})
        ok_ren2, ev_ren2 = renumerou(base, depois2, 0, P2[0])
        rnx, sx, posx = x_do_id(depois2, 0, P2[2])
        ant = [c for rn, c in ativos_em_ordem(base, 0)]
        k2 = [rn for rn, _ in ativos_em_ordem(base, 0)].index(P2[0])
        viz = ant[k2 - 1] if k2 > 0 else 0.0          # corrida da vizinha anterior (nao muda)
        ok_x = posx is not None and viz < posx < viz + 1 and abs(posx - round(posx)) > 1e-9
        dados_x = run('mIntegracao.DadosDoPonto', 1, s_x[sx], depois2[rnx]['linha'] + 1, teto=60) if sx is not None else ''
        ok_dx = dados_x.startswith('OK|') and dados_x.split('|')[4] == f'entre {int(viz)} e {int(viz) + 1}' \
            and dados_x.split('|')[7] == 'INATIVADO'
        reg('U02 InativarPorId com X + AtualizarSoEntradas: INATIVADO, X_VERMELHO, PARTICIPA = NÃO; as corridas ativas '
            'do nível renumeradas sem buraco (anteriores iguais, seguintes -1) e o X numa posição fracionária entre as '
            'vizinhas ("entre k e k+1" no DadosDoPonto)',
            r2.startswith('OK|') and f2.get('STATUS_ANALITICO') == 'INATIVADO' and f2.get('TIPO_PLOTAGEM_LJ') == 'X_VERMELHO'
            and f2.get('PARTICIPA_ESTATISTICA') == NAO and ok_ren2 and ok_x and ok_dx,
            {'ID': P2[2], 'retorno': r2, 'corrida_antes': P2[3],
             'cq': {k: f2.get(k) for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'PARTICIPA_ESTATISTICA', 'MOTIVO_INATIVACAO')},
             'renumeracao': ev_ren2, 'X': {'run': rnx, 'slot': sx, 'posicao': posx, 'vizinha_anterior': viz},
             'DadosDoPonto_X': dados_x, 'tempo_s': tempos.get('U02')})

        # ------------------------------------------------------------ U05a ja inativado
        r5a = run('mIntegracao.InativarPorId', P2[2], 'QA de novo', True, '', teto=60)
        inat5a = L.inativar()
        n_p2 = sum(1 for x in inat5a.values() if str(x['ID_REGISTRO']) == P2[2])

        # ------------------------------------------------------------ U07 reativar: numeracao volta exata
        aud7 = L.audit_ult()
        mot7 = 'QA ADR-081: inativado por engano'
        r7 = run('mIntegracao.ReativarPorId', P2[2], mot7, teto=60)
        inat7 = L.inativar()
        atualizar_so_entradas('U07')
        depois7 = L.eng_por_run()
        f7 = L.cq({P2[2]}).get(P2[2], {})
        cor_b = {rn: d['cor'] for rn, d in base.items()}
        cor_7 = {rn: d['cor'] for rn, d in depois7.items()}
        dif7 = [(rn, cor_b.get(rn), cor_7.get(rn)) for rn in set(cor_b) | set(cor_7) if cor_b.get(rn) != cor_7.get(rn)]
        a7 = L.audit_desde(aud7)
        aud_ok7 = any(x['acao'] == 'REATIVACAO_PEDIDA' and x['motivo'] == mot7 and P2[2] in str(x['parecer']) for x in a7)
        reg('U07 ReativarPorId: a linha sai da aba Inativar, o resultado volta a ATIVO e a numeração das corridas volta '
            'EXATAMENTE à de antes (Eng_Saida BK..BM, todos os níveis); Audit_Log com o motivo',
            r7.startswith('OK|') and not any(str(x['ID_REGISTRO']) == P2[2] for x in inat7.values())
            and f7.get('STATUS_ANALITICO') == 'ATIVO' and f7.get('PARTICIPA_ESTATISTICA') == SIM and not dif7 and aud_ok7,
            {'retorno': r7, 'cq': {k: f7.get(k) for k in ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'PARTICIPA_ESTATISTICA')},
             'corridas_diferentes': dif7[:10], 'audit': a7[-6:], 'tempo_s': tempos.get('U07')})

        # ------------------------------------------------------------ U03 + U09 sem X, motivo com simbolos
        # ADR-084: texto livre em DUAS linhas, com espaco e linha em branco nas pontas (saem; a quebra do meio fica)
        mot3 = 'QA ADR-081: z ≥ 3σ; ação corretiva —\nrecalibração (Δ > ET)'
        r3 = run('mIntegracao.InativarPorId', curto(P3[2]), ' \r\n' + mot3.replace('\n', '\r\n') + ' \r\n ', False,
                 str(int(P3[3])), teto=60)
        inat3 = L.inativar()
        lin3 = next((x for x in inat3.values() if str(x['ID_REGISTRO']) == P3[2]), {})
        atualizar_so_entradas('U03')
        depois3 = L.eng_por_run()
        f3 = L.cq({P3[2]}).get(P3[2], {})
        ok_ren3, ev_ren3 = renumerou(base, depois3, 0, P3[0])
        rnx3, _, _ = x_do_id(depois3, 0, P3[2])
        reg('U03 Sem X (REGISTRAR - LJ desmarcado): NAO_PLOTAR, INATIVADO, PARTICIPA = NÃO, renumera sem buraco e o ID '
            'não aparece como X',
            r3.startswith('OK|') and lin3.get('REGISTRAR - LJ') is False and f3.get('STATUS_ANALITICO') == 'INATIVADO'
            and f3.get('TIPO_PLOTAGEM_LJ') == 'NAO_PLOTAR' and f3.get('PARTICIPA_ESTATISTICA') == NAO and ok_ren3
            and rnx3 is None,
            {'ID': P3[2], 'retorno': r3, 'linha': lin3, 'cq': {k: f3.get(k) for k in
             ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'PARTICIPA_ESTATISTICA')}, 'renumeracao': ev_ren3,
             'tempo_s': tempos.get('U03')})
        reg('U09 Motivo livre com acentos, símbolos (≥ σ Δ — ç ã) e quebra de linha preservado na aba Inativar e na '
            'Principal - Resultados; espaços e linhas em branco das pontas removidos (ADR-084)',
            lin3.get('MOTIVO') == mot3 and f3.get('MOTIVO_INATIVACAO') == mot3,
            {'enviado': mot3, 'aba': lin3.get('MOTIVO'), 'principal': f3.get('MOTIVO_INATIVACAO')})
        r3b = run('mIntegracao.ReativarPorId', P3[2], 'QA ADR-081: limpeza do teste', teto=60)
        atualizar_so_entradas('U03_reativar')
        if r3b.startswith('ERRO|'):
            raise RuntimeError('limpeza do U03 falhou: ' + r3b)

        # ------------------------------------------------------------ U05b linha com analito errado: corrigida
        outro = next(a for a in cadastro if a and a.upper() != str(cq0[P5[2]]['ANALITO']).upper())
        lo = q.lo(TB)
        ws_in = lo.Parent
        col_id = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
        r_err = next(i for i, v in enumerate(col_id, start=1) if v[0] in (None, '')
                     and all(lo.ListColumns(c).DataBodyRange.Cells(i, 1).Value in (None, '', False) for c in CAB))
        q.ex.xl.EnableEvents = False
        try:
            ws_in.Unprotect(ii.SENHA)
            lo.ListColumns('ID_REGISTRO').DataBodyRange.Cells(r_err, 1).Value = P5[2]
            lo.ListColumns('ANALITO').DataBodyRange.Cells(r_err, 1).Value = outro
            lo.ListColumns('MOTIVO').DataBodyRange.Cells(r_err, 1).Value = 'QA ADR-081: linha com analito errado'
        finally:
            q.ex.xl.EnableEvents = True
            q.run('mSeguranca.ReprotectAll', teto=120)
        r5b = run('mIntegracao.InativarPorId', curto(P5[2]), 'QA ADR-081: troca de lote ou de frasco do controle', True, '',
                  teto=60)
        inat5b = L.inativar()
        linhas_p5 = {k: x for k, x in inat5b.items() if str(x['ID_REGISTRO']) == P5[2]}
        reg('U05 ID já inativado: recusa ("já inativado", sem linha nova); ID na aba com o ANALITO errado (inativação '
            'que não valia, E10): a mesma linha é corrigida -- analito da Principal, motivo novo, caixa e carimbo',
            r5a.startswith('ERRO|Este resultado já está inativado (aba Inativar, linha ') and n_p2 == 1
            and r5b.startswith('OK|') and r5b.endswith('|corrigida') and list(linhas_p5) == [r_err]
            and linhas_p5[r_err]['ANALITO'] == cq0[P5[2]]['ANALITO'] and linhas_p5[r_err]['REGISTRAR - LJ'] is True
            and linhas_p5[r_err]['DATA_INATIVACAO'] not in (None, '') and linhas_p5[r_err]['USUARIO'] not in (None, ''),
            {'repetido': r5a, 'linhas_do_ID_P2': n_p2, 'corrigida': r5b, 'linha_errada': r_err,
             'analito_errado': outro, 'linhas_do_ID_P5': linhas_p5})
        run('mIntegracao.ReativarPorId', P5[2], 'QA ADR-081: limpeza do teste', teto=60)

        # ------------------------------------------------------------ U06 analito da Principal, nao o da tela
        lo_cq = q.lo('tblCQ_Final')
        cab_cq = [c.Name for c in lo_cq.ListColumns]
        v_cq = lo_cq.DataBodyRange.Value
        ia, ist, iid = cab_cq.index('ANALITO'), cab_cq.index('STATUS_ANALITICO'), cab_cq.index('ID_REGISTRO')
        ja_na_aba = {str(x['ID_REGISTRO']) for x in L.inativar().values()}
        outro_r = next((r for r in v_cq if r[ia] and str(r[ia]).strip().upper() != analito_tela.strip().upper()
                        and r[ist] == 'ATIVO' and str(r[iid] or '').strip() and '~' not in str(r[iid])
                        and str(r[iid]) not in ja_na_aba), None)
        if outro_r is None:
            raise RuntimeError('pré-condição ausente: nenhum resultado ATIVO de outro analito')
        r6 = run('mIntegracao.InativarPorId', outro_r[iid], 'QA ADR-081: erro de identificação', True, '', teto=60)
        lin6 = next((x for x in L.inativar().values() if str(x['ID_REGISTRO']) == outro_r[iid]), {})
        reg('U06 O ANALITO gravado é o do resultado na Principal - Resultados, não o analito em tela',
            r6.startswith('OK|') and lin6.get('ANALITO') == outro_r[ia] and str(lin6.get('ANALITO')) != analito_tela,
            {'ID': outro_r[iid], 'analito_em_tela': analito_tela, 'analito_gravado': lin6.get('ANALITO'),
             'analito_do_resultado': outro_r[ia]})
        run('mIntegracao.ReativarPorId', outro_r[iid], 'QA ADR-081: limpeza do teste', teto=60)

        # ------------------------------------------------------------ U11 digitar o ID preenche o ANALITO
        r11 = q.escrever_linha(TB, {'ID_REGISTRO': curto(P11[2])})
        lin11 = {c: lo.ListColumns(c).DataBodyRange.Cells(r11, 1).Value for c in CAB}
        lo.ListColumns('ID_REGISTRO').DataBodyRange.Cells(r11, 1).ClearContents()     # evento: linha inteira sai
        lin11b = {c: lo.ListColumns(c).DataBodyRange.Cells(r11, 1).Value for c in CAB}
        reg('U11 Digitar só o número do ID na aba Inativar preenche o ANALITO sozinho (o da Principal - Resultados); '
            'apagar o ID limpa a linha',
            lin11['ID_REGISTRO'] == P11[2] and lin11['ANALITO'] == cq0[P11[2]]['ANALITO']
            and all(v in (None, '') for v in lin11b.values()),
            {'digitado': curto(P11[2]), 'linha': lin11, 'depois_de_apagar': lin11b})

        # ------------------------------------------------------------ login (usuario de teste) e U08b
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        min_antes = cfg_gravar(q, 'PAPEL_MIN_INATIVAR', '')
        sess = fazer_login(q, USUARIO_QA)
        pode1 = q.run('mSeguranca.PodeInativar', teto=60)
        cfg_gravar(q, 'PAPEL_MIN_INATIVAR', 'ADM')
        pode2 = q.run('mSeguranca.PodeInativar', teto=60)
        inat8b0 = L.inativar()
        run('mUI.DefinirRespostaTeste', 'CONFIRMAR|QA ADR-081 papel baixo|1', teto=60)
        r8b = run('mUI.InativarPontoUI', 1, s_res, P10[1] + 1, teto=120)
        inat8b = L.inativar()
        cfg_gravar(q, 'PAPEL_MIN_INATIVAR', '' if min_antes is None else min_antes)
        pode3 = q.run('mSeguranca.PodeInativar', teto=60)
        reg('U08 Sem permissão a fachada recusa e nada é gravado: sem login (PodeInativar falso) e com papel abaixo de '
            'PAPEL_MIN_INATIVAR (ANALISTA com mínimo ADM); chave vazia = qualquer usuário logado',
            pode0 is False and r8a.startswith('ERRO|') and inat8a == inat0 and sess == USUARIO_QA
            and pode1 is True and pode2 is False and r8b.startswith('ERRO|') and inat8b == inat8b0
            and pode3 is (min_antes in (None, '') or str(min_antes).strip().upper() in ('ANALISTA', 'TÉCNICO', 'TECNICO')),
            {'sem_login': {'PodeInativar': pode0, 'retorno': r8a, 'aba_igual': inat8a == inat0},
             'sessao': sess, 'PodeInativar_chave_vazia': pode1, 'PodeInativar_minimo_ADM': pode2,
             'papel_baixo': {'retorno': r8b, 'aba_igual': inat8b == inat8b0}, 'chave_original': min_antes})

        # ------------------------------------------------------------ U10 + U13 + U14 + U15 a fachada da janela
        aud10 = L.audit_ult()
        base10 = L.eng_por_run()
        lin_p10 = base10[P10[0]]['linha'] + 1
        cor_p10 = base10[P10[0]]['cor'][0]
        at10 = ativos_em_ordem(base10, 0)
        k10 = [rn for rn, _ in at10].index(P10[0])
        viz10 = int(at10[k10 - 1][1]) if k10 > 0 else 0          # vizinha anterior: nao muda com a inativacao
        onde10 = f'X vermelho entre as corridas {viz10} e {viz10 + 1}' if viz10 >= 1 else 'X vermelho antes da corrida 1'
        run('mUI.DefinirRespostaTeste', 'CANCELAR', teto=60)
        r10c = run('mUI.InativarPontoUI', 1, s_res, lin_p10, teto=120)
        inat10c = L.inativar()
        mot10 = 'QA ADR-081: falha de equipamento ou de calibração'
        run('mUI.DefinirRespostaTeste', f'CONFIRMAR|{mot10}|1', teto=60)
        t0 = time.time()
        r10 = run('mUI.InativarPontoUI', 1, s_res, lin_p10, teto=900)
        tempos['U10_inativar_UI'] = round(time.time() - t0, 1)
        prog10 = run('mUI.UltimoAvisoProgresso', teto=60)
        tip10 = tip(q, 1)
        f10 = L.cq({P10[2]}).get(P10[2], {})
        depois10 = L.eng_por_run()
        ok_ren10, ev_ren10 = renumerou(base10, depois10, 0, P10[0])
        rnx10, sx10, _ = x_do_id(depois10, 0, P10[2])
        time.sleep(BLOQUEIO_S + 0.5)                 # passou o bloqueio de cliques repetidos (mUI.BLOQUEIO_SEG)
        # no X vermelho a mesma fachada reativa
        r10r, r15 = '', ''
        if sx10 is not None:
            run('mUI.DefinirRespostaTeste', 'CONFIRMAR|QA ADR-081: investigação concluída|0', teto=60)
            t0 = time.time()
            r10r = run('mUI.InativarPontoUI', 1, s_x[sx10], depois10[rnx10]['linha'] + 1, teto=900)
            tempos['U10_reativar_UI'] = round(time.time() - t0, 1)
            # U15: o clique repetido que chega logo depois de terminar (fila do Excel) e ignorado e nao grava
            run('mUI.DefinirRespostaTeste', 'CONFIRMAR|QA ADR-081: clique repetido|1', teto=60)
            r15 = run('mUI.InativarPontoUI', 1, s_res, lin_p10, teto=120)
        prog10r = run('mUI.UltimoAvisoProgresso', teto=60)
        tip10r = tip(q, 1)
        inat15 = L.inativar()
        f10r = L.cq({P10[2]}).get(P10[2], {})
        cor_10r = {rn: d['cor'] for rn, d in L.eng_por_run().items()}
        cor_10 = {rn: d['cor'] for rn, d in base10.items()}
        a10 = L.audit_desde(aud10)
        # sem gancho e sem interface: nao abre a janela, nao trava (e prova que o gancho do clique ignorado nao ficou)
        time.sleep(BLOQUEIO_S + 0.5)
        t0 = time.time()
        r10s = run('mUI.InativarPontoUI', 1, s_res, lin_p10, teto=60)
        t_sem = round(time.time() - t0, 1)
        ok10 = (r10c.startswith('CANCELADO|') and inat10c == inat8b
                and r10.startswith('OK|' + VISTO) and 'inativado' in r10 and tip10.startswith(VISTO)
                and curto(P10[2]) in tip10 and 'BIO-' not in tip10 and 'HEM-' not in tip10
                and f10.get('STATUS_ANALITICO') == 'INATIVADO' and f10.get('TIPO_PLOTAGEM_LJ') == 'X_VERMELHO'
                and f10.get('MOTIVO_INATIVACAO') == mot10 and ok_ren10
                and r10r.startswith('OK|' + VISTO) and 'reativado' in r10r and tip10r.startswith(VISTO)
                and f10r.get('STATUS_ANALITICO') == 'ATIVO' and cor_10r == cor_10
                and any(x['acao'] == 'INATIVACAO_PEDIDA' for x in a10) and any(x['acao'] == 'REATIVACAO_PEDIDA' for x in a10)
                and r10s == 'ERRO|sem interface' and t_sem < 30)
        reg('U10 InativarPontoUI (resposta do teste no lugar da janela): CANCELAR não grava; CONFIRMAR grava, atualiza, '
            'renumera e escreve a confirmação "✓ ID … inativado — …" (ID curto) na caixa do gráfico; no X a mesma fachada '
            'reativa ("✓ ID … reativado") e a numeração volta; sem gancho com Interactive = False: "ERRO|sem interface" '
            'sem travar', ok10,
            {'cancelar': r10c, 'inativar': r10, 'tip': tip10, 'cq': {k: f10.get(k) for k in
             ('STATUS_ANALITICO', 'TIPO_PLOTAGEM_LJ', 'MOTIVO_INATIVACAO')}, 'renumeracao': ev_ren10,
             'X': {'run': rnx10, 'slot': sx10}, 'reativar': r10r, 'tip_reativar': tip10r,
             'cq_reativado': f10r.get('STATUS_ANALITICO'), 'numeracao_voltou': cor_10r == cor_10,
             'audit': [x['acao'] for x in a10][-8:], 'sem_interface': r10s, 'tempo_sem_interface_s': t_sem})
        teto = TETO_S['bio' if q.bio else 'hema']
        t_ui = tempos.get('U10_inativar_UI')
        reg(f'U13 Tempo de InativarPontoUI (confirmar → gráfico atualizado) <= {teto:.0f} s',
            t_ui is not None and t_ui <= teto and (tempos.get('U10_reativar_UI') or 0) <= teto,
            {'tempos_s': tempos, 'teto_s': teto})
        volta = f'voltou a ser a corrida {int(cor_p10)}'
        reg('U14 Progresso visível e confirmação com o lugar do ponto: "⏳ Inativando o ID … aguarde" na caixa antes da '
            'atualização (⏳ Reativando… no X); a confirmação diz "X vermelho entre as corridas k e k+1; as seguintes foram '
            'renumeradas" e, ao reativar, "voltou a ser a corrida n"',
            prog10.startswith('\u23f3') and f'Inativando o ID {curto(P10[2])}' in prog10 and 'aguarde' in prog10
            and onde10 in r10 and 'as seguintes foram renumeradas' in r10 and onde10 in tip10
            and prog10r.startswith('\u23f3') and f'Reativando o ID {curto(P10[2])}' in prog10r
            and volta in r10r and volta in tip10r,
            {'progresso_inativar': prog10, 'progresso_reativar': prog10r, 'esperado_onde': onde10,
             'confirmacao': r10, 'esperado_reativar': volta, 'confirmacao_reativar': r10r})
        reg('U15 Clique repetido até 2 s depois de terminar (o que ficou na fila do Excel) é ignorado: "IGNORADO|…", nada '
            'gravado na aba nem na Principal, e o gancho não fica armado (a chamada seguinte, sem gancho, dá "sem interface")',
            r15.startswith('IGNORADO|') and not any(str(x['ID_REGISTRO']) == P10[2] for x in inat15.values())
            and f10r.get('STATUS_ANALITICO') == 'ATIVO' and r10s == 'ERRO|sem interface',
            {'clique_repetido': r15, 'aba_tem_o_ID': any(str(x['ID_REGISTRO']) == P10[2] for x in inat15.values()),
             'cq': f10r.get('STATUS_ANALITICO'), 'chamada_seguinte': r10s})

        # ------------------------------------------------------------ U12 protecao e ATUALIZAR completo
        t0 = time.time()
        r12 = run('mIntegracao.AtualizarDadosAutomatico', teto=1500)
        t12 = round(time.time() - t0, 1)
        prot1 = L.protecao()
        reg('U12 As abas protegidas antes continuam protegidas depois (Inativar, Painel, Principal, Audit_Log, '
            'Eng_Saida, QA_INTEGRACAO) e mIntegracao.AtualizarDadosAutomatico continua "OK|"',
            all(prot1.get(k) for k, v in prot0.items() if v) and r12.startswith('OK|'),
            {'antes': prot0, 'depois': prot1, 'atualizar': r12[:300], 'tempo_atualizar_s': t12})
    except Exception as ex_geral:                    # noqa: BLE001
        reg('U99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(ex_geral).__name__}: {ex_geral}', 'traceback': traceback.format_exc()[-1500:],
             'tempos_s': tempos})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('U00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} inativar UX (ADR-081): {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3])
    sys.exit(1 if falhas else 0)
