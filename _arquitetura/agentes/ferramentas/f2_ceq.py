# -*- coding: utf-8 -*-
"""F2 do Usuario Simulado -- CEQ: registro, validacoes, preenchimento automatico, correcao, isolamento e o efeito
no bias, no ET e no Sigma (so do analito certo).

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    set PYTHONIOENCODING=utf-8
    python -u f2_ceq.py <copia.xlsm> <pasta_saida> [<analito_canonico> <nome_no_CAP> <analito_controle>
                                                   [<nome_fora_do_mapa> [<unidade>]]]
    (padrao: Glicose / "Glucose, serum" / Ureia / "Glucose" / mg/dL)
    Variaveis opcionais (3o passe, 01/11/2026):
      F2_LOTE_ALT=<lote>      liga o passo I9: Painel H3 -> outro lote -> H3 vazio (ADR-086) -> lote original;
                              confere que o CEQ (G, T, AC, AD) nao depende do lote e que o ET/Sigma acompanham.
      F2_SEM_MAPA=<nome>|<canonico>  liga o passo I10: primeiro CEQ de um analito que nao tem CEQ nem entrada no
                              mapa W:Y (nome do provedor e, depois, o proprio nome canonico).

Sempre numa COPIA da producao. A copia nunca e salva; dois retratos sao gravados por SaveCopyAs (eventos desligados)
em <pasta_saida>: f2_S1.xlsm (depois da inclusao) e f2_S2.xlsm (depois da correcao), para a calculadora independente
(calc_independente.py) refazer as 62 linhas sem Excel. O Excel e encerrado so pelo PID criado aqui (finally).

Passos (cada um grava o que viu em <pasta_saida>/f2.json; um passo que falha nao derruba os outros):
  P0  abre (qa_final.QA), cria um usuario ANALISTA de teste e faz login (caminho do usuario); Painel C3 = analito;
      retrato B0 (Estatistica A14:AU80 e A90:F110, K5, Painel, EQA_Base, tabela do CAP, Audit_Log);
      exploracao das telas do CEQ: botoes (OnAction), protecao/bloqueio das celulas, validacao de dados, eventos.
  P1  preenchimento automatico e listas da janela (DistintosCEQ, ValorRegistrado, RegistrosDaRodada).
  P2  validacoes do nucleo da janela (mCEQForm.SalvarCEQ, o que o botao "Salvar e proximo" chama): duplicidade de
      uma chave REAL, vazio, texto, limites invertidos, ano, SD negativo, avaliacao vazia, provedor invalido;
      a tabela do provedor tem de ficar identica.
  P3  INCLUSAO: rodada nova "TESTE-AUDITORIA R1" (ano EP em uso) do analito escolhido, resultado 110 / alvo 100
      (bias +10 %). Confere EQA_Base, Audit_Log, a lista da rodada, a Estatistica (so as linhas do analito mudam),
      o Painel e o recalculo independente do bias (duas etapas, ADR-035), do ET e do Sigma.
  P4  CORRECAO: a janela recusa a mesma chave ("edite a linha existente"); edicao direta da linha na tabela do
      provedor (resultado 105, bias +5 %) com os eventos ligados; o que muda sozinho; depois o botao Consolidar da
      EQA_Base; recalculo independente; trilha de auditoria da edicao.
  P5  ISOLAMENTO: Controllab do mesmo analito; ano diferente do Ano EP; nome do analito fora do mapa; registro sem
      valor-alvo; Ano EP = ano novo / TODOS; equipamento do Painel; provedor do analito trocado no Painel (M4).
  P6  PERMISSAO: logout e SalvarCEQ sem sessao.
  P7  REVERSAO: apaga as linhas TESTE-AUDITORIA como o usuario (limpa as celulas digitadas), consolida e confere que a
      Estatistica voltou bit a bit ao B0.
Marcacao dos dados de teste: rodada/amostra/arquivo com "TESTE-AUDITORIA"; valores redondos sinteticos (110, 105,
120, 130, 150). Nenhuma senha e escrita aqui (vem da suite de testes); nenhum caminho de rede e impresso.
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01 (frente F2).
"""
import datetime as dt
import hashlib
import json
import os
import statistics
import sys
import time
import traceback

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
sys.path.insert(0, os.path.join(TESTES, '..'))
from qa_final import QA  # noqa: E402
from qa_etl_recebimento import SENHA  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

COPIA = os.path.abspath(sys.argv[1])
SAIDA = os.path.abspath(sys.argv[2])
ANALITO = sys.argv[3] if len(sys.argv) > 3 else 'Glicose'
NOME_CAP = sys.argv[4] if len(sys.argv) > 4 else 'Glucose, serum'
CONTROLE = sys.argv[5] if len(sys.argv) > 5 else 'Ureia'
NOME_FORA = sys.argv[6] if len(sys.argv) > 6 else 'Glucose'
UNIDADE = sys.argv[7] if len(sys.argv) > 7 else 'mg/dL'
LOTE_ALT = os.environ.get('F2_LOTE_ALT', '').strip()
SEM_MAPA = os.environ.get('F2_SEM_MAPA', '').strip()
USUARIO = 'TA_F2_ANALISTA'
MARCA = 'TESTE-AUDITORIA'
OUT = {'passos': {}, 'erros': [], 'pids': []}
COLS = [None]


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def col_letra(i):  # 1-based
    s = ''
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def limpo(v):
    if isinstance(v, dt.datetime):
        return v.replace(tzinfo=None).isoformat()
    if isinstance(v, (list, tuple)):
        return [limpo(x) for x in v]
    return v


def sha(c):
    return hashlib.sha256(open(c, 'rb').read()).hexdigest()[:12]


def gravar():
    with open(os.path.join(SAIDA, 'f2.json'), 'w', encoding='utf-8') as f:
        json.dump(OUT, f, ensure_ascii=False, indent=1, default=str)


def passo(nome):
    def deco(fn):
        def run(*a, **k):
            log('==>', nome)
            try:
                r = fn(*a, **k)
                OUT['passos'][nome] = limpo(r)
            except Exception as e:  # noqa: BLE001
                OUT['passos'][nome] = {'ERRO': repr(e), 'trace': traceback.format_exc()[-2500:]}
                OUT['erros'].append(f'{nome}: {e!r}')
                log('ERRO', nome, repr(e))
            gravar()
            return OUT['passos'][nome]
        return run
    return deco


# ------------------------------------------------------------------------------------------------ leitura
class L:
    def __init__(self, q):
        self.q = q
        self.xl = q.ex.xl
        self.est = q.ws('Estatística')
        self.pn = q.ws('Painel')
        self.base = q.ws('EQA_Base')

    def estat(self):
        v = self.est.Range('A14:AU80').Value
        out = {}
        for i, r in enumerate(v):
            if r[0] not in (None, ''):
                out[f'{str(r[0]).strip()}|{int(r[1]) if isinstance(r[1], float) else r[1]}'] = \
                    {'linha': 14 + i, **{col_letra(j + 1): r[j] for j in range(len(r))}}
        return out

    def resumo(self):
        return [list(r) for r in self.est.Range('A90:F110').Value]

    def k5(self):
        return self.est.Range('K5').Value

    def painel(self):
        p = self.pn
        d = {c: p.Range(c).Value for c in ('C3', 'H3', 'L4', 'M4', 'G3', 'G4')}
        d['B7:I8'] = [list(r) for r in p.Range('B7:I8').Value]
        for rg in ('C46:C47', 'J46:J47', 'M46:M47'):
            d[rg] = [r[0] for r in p.Range(rg).Value]
        return d

    def eqa_base(self):
        n = int(self.xl.WorksheetFunction.CountA(self.base.Range('A2:A5001')))
        rows = [list(r) for r in self.base.Range(f'A2:U{1 + max(n, 1)}').Value] if n else []
        return {'n': n, 'carimbo': self.base.Range('Z1').Value, 'rows': rows}

    def tabela(self, prov='CAP'):
        lo = self.lo_prov(prov)
        body = [list(r) for r in lo.DataBodyRange.Value] if lo.ListRows.Count else []
        return {'linhas': lo.ListRows.Count, 'r0': lo.DataBodyRange.Row, 'body': body}

    def lo_prov(self, prov):
        nome = 'tblEQA_CAP_Dados' if prov == 'CAP' else 'tblEQA_Controllab_Dados'
        for ws in self.q.wb.Worksheets:
            for t in ws.ListObjects:
                if t.Name == nome:
                    return t
        raise RuntimeError(nome + ' ausente')

    def audit_ult(self):
        al = self.q.ws('Audit_Log')
        return al.Cells(al.Rows.Count, 1).End(-4162).Row

    def audit_desde(self, r0):
        al = self.q.ws('Audit_Log')
        ult = self.audit_ult()
        if ult <= r0:
            return []
        cab = [c for c in al.Range(al.Cells(1, 1), al.Cells(1, 30)).Value[0]]
        out = []
        for r in al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 30)).Value:
            out.append({str(cab[i]): r[i] for i in range(30) if r[i] not in (None, '') and cab[i]})
        return out


def diff_estat(a, b):
    out = []
    for k in sorted(set(a) | set(b)):
        ra, rb = a.get(k, {}), b.get(k, {})
        for c in sorted(set(ra) | set(rb), key=lambda x: (len(x), x)):
            if c == 'linha':
                continue
            if ra.get(c) != rb.get(c):
                out.append({'chave': k, 'col': c, 'antes': ra.get(c), 'depois': rb.get(c)})
    return out


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def bias_independente(rows, canon, prov, ano):
    """Recalculo independente do bias do ET/Sigma (definicao do projeto: mCEQ.BiasEQ, ADR-035/071/072, conferida no
    codigo): linhas da EQA_Base do analito canonico, Uso_Analitico <> NAO, provedor do analito, ano EP explicito
    (ou TODOS = todos os anos), rodada TODAS. bias_i = (Resultado - Alvo)/Alvo*100 RECALCULADO aqui; etapa 1 = media
    dos |bias| de cada rodada (ano|rodada); etapa 2 = media das medias. Sinal: a mesma consolidacao com sinal."""
    grupos = {}
    for r in rows:
        if str(r[4] or '').strip().upper() != canon.upper():
            continue
        if str(r[19] or '').strip().upper() == 'NAO':
            continue
        if str(r[0] or '').strip().upper() != prov.upper():
            continue
        if not num(r[1]):
            continue
        if ano != 'TODOS' and int(r[1]) != int(ano):
            continue
        res, alvo = r[6], r[7]
        if not (num(res) and num(alvo)) or alvo == 0:
            continue
        b = (res - alvo) / alvo * 100
        grupos.setdefault(f'{int(r[1])}|{r[2]}', []).append(b)
    if not grupos:
        return {'abs': 'SEM EP', 'sinal': 'SEM EP', 'n': 0, 'rodadas': 0, 'por_rodada': {}}
    m_abs = {k: statistics.fmean(abs(x) for x in v) for k, v in grupos.items()}
    m_sig = {k: statistics.fmean(v) for k, v in grupos.items()}
    return {'abs': statistics.fmean(m_abs.values()), 'sinal': statistics.fmean(m_sig.values()),
            'n': sum(len(v) for v in grupos.values()), 'rodadas': len(grupos),
            'por_rodada': {k: [m_abs[k], len(grupos[k])] for k in grupos}}


def dr(a, b):
    if not (num(a) and num(b)):
        return None if a == b else 'tipo'
    if a == b:
        return 0.0
    return abs(a - b) / max(abs(a), abs(b))


def conferir_analito(est, ind, cvp, analito):
    """Compara as linhas do analito na Estatistica com o independente: G, T, AC, AD; H = 1,65*F + |bias|;
    L = (ETp - |bias|)/CVp (CVp da linha de base: nao depende do CEQ)."""
    out = []
    for nv in (1, 2):
        r = est.get(f'{analito}|{nv}')
        if not r:
            continue
        g = ind['abs']
        h = 1.65 * r['F'] + g if num(g) and num(r['F']) else None
        sg = (r['K'] - g) / cvp[nv] if num(g) and num(r['K']) and cvp.get(nv) else None
        for col, qc, iv in (('G', r['G'], g), ('T', r['T'], ind['sinal']), ('AC', r['AC'], ind['n']),
                            ('AD', r['AD'], ind['rodadas']), ('H', r['H'], h), ('L', r['L'], sg)):
            d = dr(qc, iv)
            out.append({'nivel': nv, 'col': col, 'qc': qc, 'ind': iv, 'dif_rel': d,
                        'ok': d is not None and d != 'tipo' and d <= 1e-9})
    return out


# ------------------------------------------------------------------------------------------------ acoes
def salvar_ceq(q, **kw):
    base = {'prov': 'CAP', 'rod': f'{MARCA} R1', 'ano': '2025', 'an': NOME_CAP, 'am': 'TA-01', 'res': '110',
            'alvo': '100', 'sd': '2', 'sdi': '5', 'linf': '90', 'lsup': '110', 'aval': 'Acceptable', 'nlabs': '50',
            'unid': UNIDADE, 'pag': '1', 'arq': f'{MARCA}.pdf'}
    base.update(kw)
    campos = [base[k] for k in ('prov', 'rod', 'ano', 'an', 'am', 'res', 'alvo', 'sd', 'sdi', 'linf', 'lsup',
                                'aval', 'nlabs', 'unid', 'pag', 'arq')]
    t0 = time.time()
    r = str(q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.SalvarCEQ", campos))
    return r, round(time.time() - t0, 1)


def consolidar(q, macro):
    t0 = time.time()
    q.run(macro, teto=900)
    q.ex.esperar()
    return round(time.time() - t0, 1)


def set_celula(q, ws, cel, v):
    try:
        ws.Unprotect(SENHA)
    except Exception:  # noqa: BLE001
        pass
    q.ex.xl.EnableEvents = True
    ws.Range(cel).Value = v
    q.ex.esperar()
    q.ex.xl.Calculate()


def retrato(lx, rotulo, painel=True):
    d = {'estat': lx.estat(), 'k5': lx.k5(), 'resumo': lx.resumo()}
    if painel:
        d['painel'] = lx.painel()
    log('retrato', rotulo, 'K5 =', str(d['k5'])[:120])
    return d


def save_copy(q, nome):
    alvo = os.path.join(SAIDA, nome)
    ev = q.ex.xl.EnableEvents
    q.ex.xl.EnableEvents = False
    try:
        q.wb.SaveCopyAs(alvo)
    finally:
        q.ex.xl.EnableEvents = ev
    return alvo


# ------------------------------------------------------------------------------------------------ principal
def main():
    os.makedirs(SAIDA, exist_ok=True)
    OUT['copia'] = os.path.basename(COPIA)
    OUT['sha_copia_antes'] = sha(COPIA)
    OUT['analito'], OUT['nome_cap'], OUT['controle'] = ANALITO, NOME_CAP, CONTROLE
    q = QA('Bioquimica', COPIA)
    OUT['pids'].append(q.ex.pid)
    log('EXCEL_PID', q.ex.pid)
    gravar()
    lx = L(q)
    S = {}
    try:
        @passo('P0_abertura_exploracao')
        def p0():
            r = {}
            criar_usuario(q, USUARIO, 'ANALISTA')
            r['login'] = fazer_login(q, USUARIO)
            r['papel'] = q.nome('currentPapel').Value
            q.ex.xl.EnableEvents = True
            q.ex.xl.Calculation = -4105
            # telas do CEQ: botoes, protecao, bloqueio, validacao, eventos
            telas = {}
            for nome in ('EQA.CAP_Dados', 'EQA.CAP_Resumo', 'EQA.Controllab_Dados', 'EQA.Controllab_Resumo',
                         'EQA_Base', 'EQA.CAP_Nao_Aceitaveis', 'Início', 'Estatística'):
                ws = q.ws(nome)
                t = {'protegida': bool(ws.ProtectContents), 'visivel': ws.Visible, 'botoes': []}
                for shp in ws.Shapes:
                    b = {'nome': shp.Name, 'celula': None, 'acao': None, 'texto': None}
                    try:
                        b['celula'] = shp.TopLeftCell.Address
                    except Exception:  # noqa: BLE001
                        pass
                    try:
                        b['acao'] = shp.OnAction
                    except Exception:  # noqa: BLE001
                        pass
                    try:
                        b['texto'] = shp.TextFrame2.TextRange.Text
                    except Exception:  # noqa: BLE001
                        pass
                    if b['acao'] or b['texto']:
                        t['botoes'].append(b)
                try:
                    cm = q.wb.VBProject.VBComponents(ws.CodeName).CodeModule
                    n = cm.CountOfLines
                    cod = cm.Lines(1, n) if n else ''
                    t['eventos'] = [ln.strip() for ln in cod.splitlines() if ln.strip().startswith(('Private Sub',
                                                                                                    'Public Sub'))]
                except Exception as e:  # noqa: BLE001
                    t['eventos'] = f'nao lido: {e!r}'
                telas[nome] = t
            r['telas'] = telas
            lo = lx.lo_prov('CAP')
            ws = lo.Parent
            r0 = lo.DataBodyRange.Row
            ult = r0 + lo.ListRows.Count - 1
            cols = {}
            for c in range(1, 19):
                cel = ws.Cells(r0, lo.Range.Column + c - 1)
                vazia = ws.Cells(ult + 1, lo.Range.Column + c - 1)
                v = {'cab': lo.HeaderRowRange.Cells(1, c).Value, 'locked_linha1': bool(cel.Locked),
                     'locked_apos_tabela': bool(vazia.Locked), 'formula': bool(cel.HasFormula)}
                try:
                    v['validacao'] = [cel.Validation.Type, cel.Validation.Formula1]
                except Exception:  # noqa: BLE001
                    v['validacao'] = None
                cols[col_letra(c)] = v
            r['tabela_cap'] = {'r0': r0, 'linhas': lo.ListRows.Count, 'ultima': ult, 'colunas': cols,
                               'estado_Z2': ws.Range('Z2').Value, 'carimbo_Z1': ws.Range('Z1').Value}
            r['frmCEQ_instalado'] = any(c.Name == 'frmCEQ' for c in q.wb.VBProject.VBComponents)
            # Painel no analito testado (como o usuario: C3 pela lista)
            set_celula(q, lx.pn, 'C3', ANALITO)
            try:
                q.run('mEstatistica.AtualizarEstatistica', teto=900)
            except Exception as e:  # noqa: BLE001
                r['aviso_atualizar_estat'] = repr(e)
            S['B0'] = retrato(lx, 'B0')
            S['base0'] = lx.eqa_base()
            S['cap0'] = lx.tabela('CAP')
            S['ctl0'] = lx.tabela('Controllab')
            r['B0_painel'] = S['B0']['painel']
            r['B0_k5'] = S['B0']['k5']
            r['B0_linhas_analito'] = {k: v for k, v in S['B0']['estat'].items() if k.startswith(ANALITO + '|')
                                      or k.startswith(CONTROLE + '|')}
            r['EQA_Base'] = {'n': S['base0']['n'], 'carimbo': S['base0']['carimbo']}
            # CVp de cada nivel na linha de base (L = (K - G)/CVp; o CV pooled nao depende do CEQ)
            S['cvp'] = {}
            for nv in (1, 2):
                rr = S['B0']['estat'][f'{ANALITO}|{nv}']
                S['cvp'][nv] = (rr['K'] - rr['G']) / rr['L']
            r['cvp_derivado'] = S['cvp']
            ind = bias_independente(S['base0']['rows'], ANALITO, 'CAP', 2025)
            r['B0_conferencia'] = conferir_analito(S['B0']['estat'], ind, S['cvp'], ANALITO)
            r['B0_independente'] = ind
            return r
        p0()

        @passo('P1_preenchimento_listas')
        def p1():
            run = lambda m, *a: q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.{m}", *a)  # noqa: E731
            rod = run('DistintosCEQ', 'CAP', 2, True)
            ana = run('DistintosCEQ', 'CAP', 4, False)
            ava = run('DistintosCEQ', 'CAP', 12, False)
            rod_ctl = run('DistintosCEQ', 'Controllab', 2, True)
            return {'rodadas_CAP': list(rod or []), 'analitos_CAP': list(ana or []), 'avaliacoes_CAP': list(ava or []),
                    'rodadas_Controllab': list(rod_ctl or []),
                    'ano_de_C-C 2025': run('ValorRegistrado', 'CAP', 2, 'C-C 2025', 3),
                    'unidade_de_' + NOME_CAP: run('ValorRegistrado', 'CAP', 4, NOME_CAP, 14),
                    'unidade_de_' + ANALITO: run('ValorRegistrado', 'CAP', 4, ANALITO, 14),
                    'ano_de_rodada_inexistente': run('ValorRegistrado', 'CAP', 2, 'XYZ', 3),
                    'regs_C-C_2025': len(run('RegistrosDaRodada', 'CAP', 'C-C 2025', '2025') or []),
                    'regs_C-C_2025_ano_2026': run('RegistrosDaRodada', 'CAP', 'C-C 2025', '2026')}
        p1()

        @passo('P2_validacoes')
        def p2():
            casos = {
                'V01 duplicidade de chave REAL (C-A 2025, CHM-01)': dict(rod='C-A 2025', am='CHM-01', res='999'),
                'V02 resultado vazio': dict(am='TA-V02', res=''),
                'V03 resultado texto': dict(am='TA-V03', res='abc'),
                'V04 limite inferior > superior': dict(am='TA-V04', linf='120', lsup='80'),
                'V05 ano com 2 digitos': dict(am='TA-V05', ano='25'),
                'V06 SD negativo': dict(am='TA-V06', sd='-1'),
                'V07 avaliacao vazia': dict(am='TA-V07', aval=''),
                'V08 provedor invalido': dict(am='TA-V08', prov='PNCQ'),
                'V09 rodada vazia': dict(am='TA-V09', rod=''),
                'V10 amostra vazia': dict(am=''),
                'V11 analito vazio': dict(am='TA-V11', an=''),
                'V12 alvo texto': dict(am='TA-V12', alvo='cem'),
            }
            out = {}
            antes = lx.tabela('CAP')
            for k, kw in casos.items():
                out[k] = salvar_ceq(q, **kw)[0]
            depois = lx.tabela('CAP')
            return {'retornos': out, 'tabela_identica': antes == depois,
                    'linhas_antes_depois': [antes['linhas'], depois['linhas']]}
        p2()

        @passo('P3_inclusao')
        def p3():
            a0 = lx.audit_ult()
            ret, seg = salvar_ceq(q)  # rodada nova TESTE-AUDITORIA R1, ano 2025, 110 / 100
            q.ex.esperar()
            S['B1'] = retrato(lx, 'B1')
            S['base1'] = lx.eqa_base()
            cap1 = lx.tabela('CAP')
            linha = None
            if ret.startswith('OK|'):
                linha = int(ret.split('|')[1])
            row_tab = [x for x in cap1['body'] if str(x[1]) == f'{MARCA} R1']
            row_base = [x for x in S['base1']['rows'] if str(x[2]) == f'{MARCA} R1']
            ind = bias_independente(S['base1']['rows'], ANALITO, 'CAP', 2025)
            dif = diff_estat(S['B0']['estat'], S['B1']['estat'])
            outros = [d for d in dif if not d['chave'].startswith(ANALITO + '|')]
            regs = q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.RegistrosDaRodada", 'CAP', f'{MARCA} R1', '2025')
            S['linha_R1'] = linha
            cap_res = q.ws('EQA.CAP_Resumo').UsedRange.Value
            no_resumo = any(f'{MARCA}' in str(c) for r in cap_res for c in r if c)
            return {'retorno': ret, 'segundos': seg, 'linha_tabela': row_tab, 'linha_EQA_Base': row_base,
                    'carimbo': S['base1']['carimbo'], 'trilha': lx.audit_desde(a0), 'lista_da_rodada': regs,
                    'K5': [S['B0']['k5'], S['B1']['k5']], 'independente': ind,
                    'conferencia': conferir_analito(S['B1']['estat'], ind, S['cvp'], ANALITO),
                    'mudancas_analito': [d for d in dif if d['chave'].startswith(ANALITO + '|')],
                    'mudancas_outros_analitos': outros,
                    'painel_antes': S['B0']['painel'], 'painel_depois': S['B1']['painel'],
                    'resumo_6sigma_mudou': S['B0']['resumo'] != S['B1']['resumo'],
                    'resumo_CAP_mostra_teste': no_resumo,
                    'S1': save_copy(q, 'f2_S1.xlsm')}
        p3()

        @passo('P4_correcao')
        def p4():
            r = {}
            r['janela_mesma_chave'] = salvar_ceq(q, res='105')[0]
            lo = lx.lo_prov('CAP')
            ws = lo.Parent
            lin = S.get('linha_R1')
            r['linha'] = lin
            cel = ws.Cells(lin, lo.Range.Column + 5)   # F = Seu Resultado
            r['protegida'] = bool(ws.ProtectContents)
            r['celula_bloqueada'] = bool(cel.Locked)
            a0 = lx.audit_ult()
            q.ex.xl.EnableEvents = True
            escrita = None
            try:
                cel.Value = 105          # o que o usuario digita (celula da tabela, eventos ligados)
                escrita = 'sem desproteger'
            except Exception as e:  # noqa: BLE001
                r['erro_escrita_protegida'] = repr(e)
                ws.Unprotect(SENHA)
                cel.Value = 105
                escrita = 'precisou desproteger (passo tecnico)'
            r['escrita'] = escrita
            q.ex.esperar()
            q.ex.xl.Calculate()
            r['O_bias_na_tabela'] = ws.Cells(lin, lo.Range.Column + 14).Value
            r['estado_Z2_apos_editar'] = ws.Range('Z2').Value
            r['carimbo_Z1_apos_editar'] = ws.Range('Z1').Value
            b2a = retrato(lx, 'B2a (editado, sem consolidar)')
            base2a = lx.eqa_base()
            r['EQA_Base_linha_apos_editar'] = [x for x in base2a['rows'] if str(x[2]) == f'{MARCA} R1']
            r['estat_mudou_sem_consolidar'] = diff_estat(S['B1']['estat'], b2a['estat'])
            r['trilha_da_edicao'] = lx.audit_desde(a0)
            # botao Consolidar da EQA_Base (o que a tela manda clicar)
            acao = None
            for shp in q.ws('EQA_Base').Shapes:
                try:
                    oa = str(shp.OnAction or '')
                except Exception:  # noqa: BLE001
                    oa = ''
                try:
                    tx = str(shp.TextFrame2.TextRange.Text or '')
                except Exception:  # noqa: BLE001
                    tx = ''
                if oa and ('EQABASE' in oa.upper() or 'CONSOLID' in tx.upper()):
                    acao = oa
                    r['texto_botao_consolidar'] = tx
            r['acao_botao_consolidar'] = acao
            macro = acao.split('!')[-1] if acao else 'mEQA.AtualizarEQABase'
            if '.' not in macro:
                macro = 'mEQA.' + macro if macro == 'AtualizarEQABase' else macro
            a1 = lx.audit_ult()
            r['segundos_consolidar'] = consolidar(q, macro)
            r['macro_executada'] = macro
            S['B2'] = retrato(lx, 'B2')
            S['base2'] = lx.eqa_base()
            r['trilha_consolidar'] = lx.audit_desde(a1)
            r['EQA_Base_linha'] = [x for x in S['base2']['rows'] if str(x[2]) == f'{MARCA} R1']
            ind = bias_independente(S['base2']['rows'], ANALITO, 'CAP', 2025)
            r['independente'] = ind
            r['conferencia'] = conferir_analito(S['B2']['estat'], ind, S['cvp'], ANALITO)
            dif = diff_estat(S['B1']['estat'], S['B2']['estat'])
            r['mudancas_analito'] = [d for d in dif if d['chave'].startswith(ANALITO + '|')]
            r['mudancas_outros_analitos'] = [d for d in dif if not d['chave'].startswith(ANALITO + '|')]
            r['K5'] = S['B2']['k5']
            r['painel'] = S['B2']['painel']
            r['estado_Z2_apos_consolidar'] = ws.Range('Z2').Value
            r['lista_da_rodada'] = q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.RegistrosDaRodada", 'CAP', f'{MARCA} R1', '2025')
            r['S2'] = save_copy(q, 'f2_S2.xlsm')
            return r
        p4()

        @passo('P5_isolamento')
        def p5():
            r = {}
            ref = S['B2']

            def comparar(rotulo, acao):
                a0 = lx.audit_ult()
                ret = acao()
                q.ex.esperar()
                b = retrato(lx, rotulo, painel=False)
                base = lx.eqa_base()
                dif = diff_estat(ref['estat'], b['estat'])
                return {'retorno': ret, 'carimbo': base['carimbo'], 'K5': b['k5'], 'mudancas': dif,
                        'n_mudancas': len(dif), 'trilha': [t.get('Ação') or t.get('ACAO') or t for t in lx.audit_desde(a0)],
                        '_b': b, '_base': base}

            # I2 Controllab do mesmo analito (provedor do analito = CAP)
            x = comparar('I2 Controllab', lambda: salvar_ceq(q, prov='Controllab', rod=f'{MARCA} CTL1', an=ANALITO,
                                                           am='TA-C1', res='120', alvo='100'))
            x['linha_EQA_Base'] = [y for y in x.pop('_base')['rows'] if str(y[2]) == f'{MARCA} CTL1']
            x.pop('_b')
            r['I2_controllab'] = x
            # I3 CAP de outro ano (2026) com Ano EP = 2025
            x = comparar('I3 ano 2026', lambda: salvar_ceq(q, rod=f'{MARCA} R2', ano='2026', am='TA-R2',
                                                         res='130', alvo='100'))
            x.pop('_b')
            x['linha_EQA_Base'] = [y for y in x.pop('_base')['rows'] if str(y[2]) == f'{MARCA} R2']
            r['I3_outro_ano'] = x
            # I4 nome fora do mapa
            x = comparar('I4 nome fora do mapa', lambda: salvar_ceq(q, an=NOME_FORA, am='TA-03', res='150', alvo='100'))
            x.pop('_b')
            x['linha_EQA_Base'] = [y for y in x.pop('_base')['rows'] if str(y[4] or '') == '' and str(y[3]) == NOME_FORA]
            r['I4_nome_fora_do_mapa'] = x
            # I5 sem valor-alvo
            x = comparar('I5 sem alvo', lambda: salvar_ceq(q, am='TA-04', res='100', alvo=''))
            x.pop('_b')
            x['linha_EQA_Base'] = [y for y in x.pop('_base')['rows'] if str(y[5]) == 'TA-04']
            r['I5_sem_alvo'] = x
            S['base5'] = lx.eqa_base()
            ref5 = retrato(lx, 'ref5', painel=True)
            r['painel_ref5'] = ref5['painel']
            # I6 Ano EP = 2026 e TODOS (filtro do usuario na Estatistica N4)
            est = lx.est
            n4 = est.Range('N4').Value
            r['N4_original'] = n4
            for ano in (2026, 'TODOS'):
                set_celula(q, est, 'N4', ano)
                b = retrato(lx, f'N4={ano}', painel=False)
                ind = bias_independente(S['base5']['rows'], ANALITO, 'CAP', ano)
                ctl = bias_independente(S['base5']['rows'], CONTROLE, 'CAP', ano)
                dif = diff_estat(ref5['estat'], b['estat'])
                r[f'I6_N4_{ano}'] = {
                    'K5': b['k5'], 'independente': ind,
                    'conferencia': conferir_analito(b['estat'], ind, S['cvp'], ANALITO),
                    'controle_independente': ctl,
                    'controle_conferencia': conferir_analito(b['estat'], ctl, {1: None, 2: None}, CONTROLE),
                    'n_mudancas': len(dif),
                    'analitos_que_mudaram_G': sorted({d['chave'] for d in dif if d['col'] == 'G'}),
                    'G_de_todos': {k: v['G'] for k, v in b['estat'].items()}}
            set_celula(q, est, 'N4', n4)
            b = retrato(lx, 'N4 restaurado', painel=False)
            r['I6_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            # I7 equipamento do Painel
            l4 = lx.pn.Range('L4').Value
            outro = 'DIMENSION 1' if str(l4).strip().upper() == 'DIMENSION 2' else 'DIMENSION 2'
            set_celula(q, lx.pn, 'L4', outro)
            try:
                q.run('mEstatistica.AtualizarEstatistica', teto=900)
            except Exception as e:  # noqa: BLE001
                r['aviso_I7'] = repr(e)
            b = retrato(lx, f'L4={outro}', painel=True)
            dif = diff_estat(ref5['estat'], b['estat'])
            r['I7_equipamento'] = {
                'de_para': [l4, outro],
                'G_T_AC_AD_iguais_todos': all(d['col'] not in ('G', 'T', 'AC', 'AD') for d in dif),
                'colunas_que_mudaram': sorted({d['col'] for d in dif}),
                'linha_analito_depois': {k: v for k, v in b['estat'].items() if k.startswith(ANALITO + '|')},
                'painel': b['painel']}
            set_celula(q, lx.pn, 'L4', l4)
            try:
                q.run('mEstatistica.AtualizarEstatistica', teto=900)
            except Exception as e:  # noqa: BLE001
                r['aviso_I7b'] = repr(e)
            b = retrato(lx, 'L4 restaurado', painel=False)
            r['I7_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            # I8 provedor do analito no Painel (M4) -> Controllab
            esp = q.espec()
            lin_esp = None
            for i in range(4, 44):
                if str(esp.Cells(i, 1).Value or '').strip().upper() == ANALITO.upper():
                    lin_esp = i
            r['AR_antes'] = esp.Cells(lin_esp, 44).Value if lin_esp else None
            a0 = lx.audit_ult()
            set_celula(q, lx.pn, 'M4', 'Controllab')
            b = retrato(lx, 'M4=Controllab', painel=True)
            dif = diff_estat(ref5['estat'], b['estat'])
            r['I8_provedor_controllab'] = {
                'AR_depois': esp.Cells(lin_esp, 44).Value if lin_esp else None,
                'M4_depois': lx.pn.Range('M4').Value, 'M4_formula': lx.pn.Range('M4').Formula,
                'linha_analito': {k: {c: v[c] for c in ('G', 'H', 'L', 'T', 'AC', 'AD', 'R', 'S')}
                                  for k, v in b['estat'].items() if k.startswith(ANALITO + '|')},
                'outros_analitos_mudaram': sorted({d['chave'] for d in dif if not d['chave'].startswith(ANALITO + '|')}),
                'painel': b['painel'], 'trilha': lx.audit_desde(a0)}
            set_celula(q, lx.pn, 'M4', r['AR_antes'] or 'CAP')
            b = retrato(lx, 'M4 restaurado', painel=False)
            r['I8_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            r['AR_restaurado'] = esp.Cells(lin_esp, 44).Value if lin_esp else None
            # I9 lote do Painel (H3): outro lote, H3 vazio (TODOS, ADR-086), lote original
            if LOTE_ALT:
                h3 = lx.pn.Range('H3').Value
                r['I9_H3_original'] = h3
                for rot, val in (('outro_lote', LOTE_ALT), ('H3_vazio', '')):
                    t0 = time.time()
                    set_celula(q, lx.pn, 'H3', val if val else None)
                    try:
                        q.run('mEstatistica.AtualizarEstatistica', teto=900)
                    except Exception as e:  # noqa: BLE001
                        r[f'aviso_I9_{rot}'] = repr(e)
                    b = retrato(lx, f'H3={val!r}', painel=True)
                    dif = diff_estat(ref5['estat'], b['estat'])
                    la = {k: v for k, v in b['estat'].items() if k.startswith(ANALITO + '|')}
                    conf = []
                    for k, v in la.items():
                        g, f, h = v.get('G'), v.get('F'), v.get('H')
                        hi = 1.65 * f + g if num(f) and num(g) else None
                        conf.append({'chave': k, 'F': f, 'G': g, 'H': h, 'H_ind': hi, 'dif_rel_H': dr(h, hi),
                                     'L': v.get('L'), 'L_ref5': ref5['estat'][k].get('L'),
                                     'C_D_E': [v.get('B'), v.get('C'), v.get('D'), v.get('E')]})
                    r[f'I9_{rot}'] = {
                        'H3': lx.pn.Range('H3').Value, 'segundos': round(time.time() - t0, 1), 'K5': b['k5'],
                        'G_T_AC_AD_iguais_todos': all(d['col'] not in ('G', 'T', 'AC', 'AD') for d in dif),
                        'mud_GTACAD': [d for d in dif if d['col'] in ('G', 'T', 'AC', 'AD')][:20],
                        'colunas_que_mudaram': sorted({d['col'] for d in dif}),
                        'n_mudancas': len(dif), 'conferencia_analito': conf,
                        'cab_estat': [lx.est.Range('A3:P5').Value], 'painel': b['painel']}
                set_celula(q, lx.pn, 'H3', h3)
                try:
                    q.run('mEstatistica.AtualizarEstatistica', teto=900)
                except Exception as e:  # noqa: BLE001
                    r['aviso_I9_restaura'] = repr(e)
                b = retrato(lx, 'H3 restaurado', painel=True)
                dif = diff_estat(ref5['estat'], b['estat'])
                r['I9_restaurado_identico'] = dif == []
                r['I9_restaurado_dif'] = dif[:20]
                r['I9_painel_restaurado_igual'] = b['painel'] == ref5['painel']
            # I10 primeiro CEQ de um analito sem CEQ e sem entrada no mapa
            if SEM_MAPA:
                ref = ref5  # noqa: F841 (comparar le 'ref' na hora da chamada)
                nome_prov, canon = SEM_MAPA.split('|')
                for i, an in enumerate((nome_prov, canon)):
                    x = comparar(f'I10 {an}', lambda an=an, i=i: salvar_ceq(q, an=an, am=f'TA-H{i + 1}',
                                                                            res='110', alvo='100', unid='%'))
                    b = x.pop('_b')
                    x['linha_EQA_Base'] = [y for y in x.pop('_base')['rows'] if str(y[5]) == f'TA-H{i + 1}']
                    x['linha_canonico'] = {k: {c: v.get(c) for c in ('B', 'F', 'G', 'H', 'L', 'T', 'AC', 'AD')}
                                           for k, v in b['estat'].items() if k.startswith(canon + '|')}
                    x['mudancas'] = x['mudancas'][:20]
                    r[f'I10_{i + 1}_{an}'] = x
                r['I10_mapear'] = [q.ex.xl.Run(f"'{q.wb.Name}'!mEQA.MapearAnalito", 'CAP', a)
                                   for a in (nome_prov, canon)]
            return r
        p5()

        @passo('P6_permissao')
        def p6():
            r = {}
            q.run('mSeguranca.Logout', teto=300)
            r['sessao_apos_logout'] = q.run('mSeguranca.SessaoAtiva', teto=60)
            r['usuario_celula'] = q.nome('currentUser').Value
            r['salvar_sem_sessao'] = salvar_ceq(q, am='TA-SEMLOGIN', res='100', alvo='100')[0]
            r['aviso'] = 'AbrirCEQ nao foi chamado (abre janela modal e travaria a automacao); o nucleo e o mesmo'
            # volta a sessao para a reversao
            r['relogin'] = fazer_login(q, USUARIO)
            return r
        p6()

        @passo('P7_reversao')
        def p7():
            r = {}
            apagadas = []
            for prov in ('CAP', 'Controllab'):
                lo = lx.lo_prov(prov)
                ws = lo.Parent
                try:
                    ws.Unprotect(SENHA)
                except Exception:  # noqa: BLE001
                    pass
                body = lo.DataBodyRange.Value
                for i, row in enumerate(body):
                    if any(MARCA in str(c or '') for c in row[:18]) or str(row[3] or '') == NOME_FORA \
                            or str(row[4] or '').startswith('TA-'):
                        rr = lo.DataBodyRange.Row + i
                        apagadas.append([prov, rr, list(row[:6])])
                        for c in list(range(2, 15)) + [17, 18]:     # B..N, Q, R (A = provedor fica como estava)
                            ws.Cells(rr, lo.Range.Column + c - 1).ClearContents()
            r['linhas_limpas'] = apagadas
            r['segundos_consolidar'] = consolidar(q, 'mEQA.AtualizarEQABase')
            b = retrato(lx, 'B_final', painel=True)
            base = lx.eqa_base()
            dif = diff_estat(S['B0']['estat'], b['estat'])
            r['estat_identica_B0'] = dif == []
            r['diferencas_vs_B0'] = dif[:40]
            r['K5_igual_B0'] = b['k5'] == S['B0']['k5']
            r['EQA_Base_linhas_iguais_B0'] = base['rows'] == S['base0']['rows']
            r['EQA_Base_n'] = [S['base0']['n'], base['n']]
            r['carimbo_final'] = base['carimbo']
            r['painel_igual_B0'] = b['painel'] == S['B0']['painel']
            return r
        p7()
    finally:
        q.fechar(salvar=False)
        OUT['sha_copia_depois'] = sha(COPIA)
        gravar()
        log('fim; copia intacta =', OUT['sha_copia_antes'] == OUT['sha_copia_depois'])


if __name__ == '__main__':
    main()
