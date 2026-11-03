# -*- coding: utf-8 -*-
"""F2 do Usuario Simulado na HEMATOLOGIA -- CEQ: registro, validacoes, preenchimento automatico, correcao, isolamento
e o efeito no bias, no ET e no Sigma (so do analito certo). Adaptacao de f2_ceq.py (Bioquimica) ao layout da Hema.

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    set PYTHONIOENCODING=utf-8
    python -u f2_ceq_hema.py <copia.xlsm> <pasta_saida> [<analito_canonico> <nome_no_CAP> <analito_controle> <lote_alt> <unidade>]
    (padrao: RBC / "Red Blood Cell Count" / WBC / 619611 / "x 10E12/L")

Diferencas da Hematologia (conferidas no codigo e nas formulas em 31/10/2026):
  - Estatistica A14:AU97 = 28 analitos x 3 niveis, todas as linhas sempre (nao dependem do Painel);
  - provedor do CEQ UNICO e global em Estatistica!R4 (eqProvedor), Ano EP N4, Rodada P4 (a Bio usa Especificacoes AR);
  - nao ha equipamento no Painel (equipamento unico, Cfg_Integracao B11); o Painel tem o lote em H3 e o periodo do
    CIQ e Ano De/Ate na Estatistica (B3/D3);
  - Painel B7:J9 (3 niveis); G7:G9, I7:I9, J/M/O45:47, C/H61:63 leem a Estatistica; H7:H9 e calculado no Painel;
  - a tabela do CAP tem 24 colunas (S..X auxiliares) e NENHUMA linha vazia: o SalvarCEQ acrescenta uma linha.

Sempre numa COPIA da producao. A copia nunca e salva; dois retratos sao gravados por SaveCopyAs (eventos desligados)
em <pasta_saida>: f2h_S1.xlsm (depois da inclusao) e f2h_S2.xlsm (depois da correcao), para calc_independente.py.
O Excel e encerrado so pelo PID criado aqui (finally).

Passos (gravados em <pasta_saida>/f2h.json; um passo que falha nao derruba os outros):
  P0  abre (qa_final.QA), cria um usuario ANALISTA de teste e faz login; Painel C3 = analito; retrato B0; exploracao
      das telas (botoes/OnAction, protecao, Locked, validacao, eventos).
  P1  preenchimento automatico e listas da janela (DistintosCEQ, ValorRegistrado, RegistrosDaRodada).
  P2  validacoes do nucleo da janela (mCEQForm.SalvarCEQ): duplicidade de chave REAL e 11 casos invalidos.
  P3  INCLUSAO: rodada "TESTE-AUDITORIA R1" (2025) do analito, 110 / 100 (bias +10 %). Confere EQA_Base, Audit_Log,
      lista da rodada, Estatistica (so as linhas do analito), Painel e o recalculo independente (bias, ET, Sigma).
  P4  CORRECAO: a janela recusa a mesma chave; edicao direta da celula (105, bias +5 %) com eventos ligados; o que
      muda sozinho; depois o botao Consolidar; recalculo independente; trilha.
  P5  ISOLAMENTO: Controllab do mesmo analito; CAP de 2026 com Ano EP 2025; nome fora do mapa; sem alvo;
      Ano EP = 2026 / TODOS; periodo do CIQ (Ano De/Ate = 2026); lote do Painel (H3); provedor global R4 = Controllab.
  P6  PERMISSAO: logout e SalvarCEQ sem sessao.
  P7  REVERSAO: limpa as linhas de teste como o usuario faria, consolida, confere bit a bit contra o B0.
Marcacao dos dados de teste: rodada/amostra/arquivo com "TESTE-AUDITORIA" (amostras TA-*); valores redondos
sinteticos (110, 105, 120, 130, 150, 100). Nenhuma senha e escrita aqui (vem da suite); nenhum caminho de rede.
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01 (frente F2, Hematologia).
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
ANALITO = sys.argv[3] if len(sys.argv) > 3 else 'RBC'
NOME_CAP = sys.argv[4] if len(sys.argv) > 4 else 'Red Blood Cell Count'
CONTROLE = sys.argv[5] if len(sys.argv) > 5 else 'WBC'
LOTE_ALT = sys.argv[6] if len(sys.argv) > 6 else '619611'
UNIDADE = sys.argv[7] if len(sys.argv) > 7 else 'x 10E12/L'
USUARIO = 'TA_F2H_ANALISTA'
MARCA = 'TESTE-AUDITORIA'
NIV = (1, 2, 3)
EST_RG = 'A14:AU97'
OUT = {'passos': {}, 'erros': [], 'pids': []}


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def col_letra(i):
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
    if isinstance(v, dict):
        return {k: limpo(x) for k, x in v.items()}
    return v


def sha(c):
    return hashlib.sha256(open(c, 'rb').read()).hexdigest()[:12]


def gravar():
    with open(os.path.join(SAIDA, 'f2h.json'), 'w', encoding='utf-8') as f:
        json.dump(limpo(OUT), f, ensure_ascii=False, indent=1, default=str)


def passo(nome):
    def deco(fn):
        def run(*a, **k):
            log('==>', nome)
            try:
                OUT['passos'][nome] = limpo(fn(*a, **k))
            except Exception as e:  # noqa: BLE001
                OUT['passos'][nome] = {'ERRO': repr(e), 'trace': traceback.format_exc()[-2500:]}
                OUT['erros'].append(f'{nome}: {e!r}')
                log('ERRO', nome, repr(e))
            gravar()
            return OUT['passos'][nome]
        return run
    return deco


class L:
    def __init__(self, q):
        self.q = q
        self.xl = q.ex.xl
        self.est = q.ws('Estatística')
        self.pn = q.ws('Painel')
        self.base = q.ws('EQA_Base')

    def estat(self):
        v = self.est.Range(EST_RG).Value
        out = {}
        for i, r in enumerate(v):
            if r[0] not in (None, ''):
                out[f'{str(r[0]).strip()}|{int(r[1]) if isinstance(r[1], float) else r[1]}'] = \
                    {'linha': 14 + i, **{col_letra(j + 1): r[j] for j in range(len(r))}}
        return out

    def filtros(self):
        e = self.est
        return {c: e.Range(c).Value for c in ('B3', 'D3', 'B4', 'N4', 'P4', 'R4', 'K5', 'L11', 'N11', 'P11')}

    def painel(self):
        p = self.pn
        d = {c: p.Range(c).Value for c in ('C3', 'H3', 'K3', 'K4', 'G3', 'G4', 'O4', 'Q4')}
        d['B7:J9'] = [list(r) for r in p.Range('B7:J9').Value]
        for rg in ('C45:C47', 'J45:J47', 'M45:M47', 'O45:O47', 'C61:C63', 'H61:H63'):
            d[rg] = [r[0] for r in p.Range(rg).Value]
        return d

    def eqa_base(self):
        n = int(self.xl.WorksheetFunction.CountA(self.base.Range('A2:A5001')))
        rows = [list(r) for r in self.base.Range(f'A2:U{1 + max(n, 1)}').Value] if n else []
        return {'n': n, 'carimbo': self.base.Range('Z1').Value, 'rows': rows}

    def lo_prov(self, prov):
        nome = 'tblEQA_CAP_Dados' if prov == 'CAP' else 'tblEQA_Controllab_Dados'
        for ws in self.q.wb.Worksheets:
            for t in ws.ListObjects:
                if t.Name == nome:
                    return t
        raise RuntimeError(nome + ' ausente')

    def tabela(self, prov='CAP'):
        lo = self.lo_prov(prov)
        body = [list(r) for r in lo.DataBodyRange.Value] if lo.ListRows.Count else []
        return {'linhas': lo.ListRows.Count, 'r0': lo.DataBodyRange.Row, 'body': body}

    def audit_ult(self):
        al = self.q.ws('Audit_Log')
        return al.Cells(al.Rows.Count, 1).End(-4162).Row

    def audit_desde(self, r0):
        al = self.q.ws('Audit_Log')
        ult = self.audit_ult()
        if ult <= r0:
            return []
        cab = list(al.Range(al.Cells(1, 1), al.Cells(1, 30)).Value[0])
        out = []
        for r in al.Range(al.Cells(r0 + 1, 1), al.Cells(ult, 30)).Value:
            out.append({str(cab[i]): r[i] for i in range(30) if r[i] not in (None, '') and cab[i]})
        return out


def diff_estat(a, b):
    out = []
    for k in sorted(set(a) | set(b)):
        ra, rb = a.get(k, {}), b.get(k, {})
        for c in sorted(set(ra) | set(rb), key=lambda x: (len(x), x)):
            if c != 'linha' and ra.get(c) != rb.get(c):
                out.append({'chave': k, 'col': c, 'antes': ra.get(c), 'depois': rb.get(c)})
    return out


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def bias_independente(rows, canon, prov, ano):
    """mCEQ.BiasEQ (ADR-035/071/072), conferido no codigo: EQA_Base do analito canonico, Uso <> NAO, provedor =
    Estatistica R4, ano EP explicito (TODOS = todos os anos), rodada TODAS. bias_i recalculado de Resultado e Alvo;
    etapa 1 media dos |bias| por ano|rodada; etapa 2 media das medias. Sinal: a mesma com o bias com sinal."""
    grupos = {}
    for r in rows:
        if str(r[4] or '').strip().upper() != canon.upper():
            continue
        if str(r[19] or '').strip().upper() == 'NAO':
            continue
        if str(r[0] or '').strip().upper() != str(prov).upper():
            continue
        if not num(r[1]):
            continue
        if ano != 'TODOS' and int(r[1]) != int(ano):
            continue
        res, alvo = r[6], r[7]
        if not (num(res) and num(alvo)) or alvo == 0:
            continue
        grupos.setdefault(f'{int(r[1])}|{r[2]}', []).append((res - alvo) / alvo * 100)
    if not grupos:
        return {'abs': 'SEM EP', 'sinal': 'SEM EP', 'n': 'SEM EP', 'rodadas': 'SEM EP', 'por_rodada': {}}
    m_abs = {k: statistics.fmean(abs(x) for x in v) for k, v in grupos.items()}
    m_sig = {k: statistics.fmean(v) for k, v in grupos.items()}
    return {'abs': statistics.fmean(m_abs.values()), 'sinal': statistics.fmean(m_sig.values()),
            'n': sum(len(v) for v in grupos.values()), 'rodadas': len(grupos),
            'por_rodada': {k: [m_abs[k], m_sig[k], len(grupos[k])] for k in grupos}}


def dr(a, b):
    if not (num(a) and num(b)):
        return 0.0 if a == b else 'tipo'
    if a == b:
        return 0.0
    return abs(a - b) / max(abs(a), abs(b))


def conferir_analito(est, ind, cvp, analito):
    """G, T, AC, AD; H = 1,65*F + |bias|; L = (ETp - |bias|)/CVp (CVp derivado da linha de base: nao depende do CEQ).
    Com SEM EP, o esperado e G/T/AC/AD = 'SEM EP' e H/L vazios."""
    out = []
    for nv in NIV:
        r = est.get(f'{analito}|{nv}')
        if not r:
            continue
        g = ind['abs']
        if num(g):
            h = 1.65 * r['F'] + abs(g) if num(r['F']) else ''
            sg = (r['K'] - abs(g)) / cvp[nv] if num(r['K']) and cvp.get(nv) else ('-' if r['K'] == '-' else None)
        else:
            h, sg = '', ('-' if r['K'] == '-' else '')
        for col, qc, iv in (('G', r['G'], g), ('T', r['T'], ind['sinal']), ('AC', r['AC'], ind['n']),
                            ('AD', r['AD'], ind['rodadas']), ('H', r['H'], h), ('L', r['L'], sg)):
            if iv is None:
                continue
            d = dr(qc, iv)
            out.append({'nivel': nv, 'col': col, 'qc': qc, 'ind': iv, 'dif_rel': d,
                        'ok': d != 'tipo' and d <= 1e-9})
    return out


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
    d = {'estat': lx.estat(), 'filtros': lx.filtros()}
    if painel:
        d['painel'] = lx.painel()
    log('retrato', rotulo, 'K5 =', str(d['filtros']['K5'])[:120])
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


def so(dif, analito, dentro=True):
    return [d for d in dif if d['chave'].startswith(analito + '|') == dentro]


def main():
    os.makedirs(SAIDA, exist_ok=True)
    OUT['copia'] = os.path.basename(COPIA)
    OUT['sha_copia_antes'] = sha(COPIA)
    OUT['analito'], OUT['nome_cap'], OUT['controle'], OUT['lote_alt'] = ANALITO, NOME_CAP, CONTROLE, LOTE_ALT
    q = QA('Hematologia', COPIA)
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
            telas = {}
            for nome in ('EQA.CAP_Dados', 'EQA.CAP_Resumo', 'EQA.Controllab_Dados', 'EQA.Controllab_Resumo',
                         'EQA_Base', 'EQA.CAP_Nao_Aceitaveis', 'Início', 'Estatística'):
                ws = q.ws(nome)
                t = {'protegida': bool(ws.ProtectContents), 'botoes': []}
                for shp in ws.Shapes:
                    b = {'nome': shp.Name, 'celula': None, 'acao': None, 'texto': None}
                    for k, f in (('celula', lambda: shp.TopLeftCell.Address), ('acao', lambda: shp.OnAction),
                                 ('texto', lambda: shp.TextFrame2.TextRange.Text)):
                        try:
                            b[k] = f()
                        except Exception:  # noqa: BLE001
                            pass
                    if b['acao'] or b['texto']:
                        t['botoes'].append(b)
                try:
                    cm = q.wb.VBProject.VBComponents(ws.CodeName).CodeModule
                    n = cm.CountOfLines
                    cod = cm.Lines(1, n) if n else ''
                    t['eventos'] = [ln.strip() for ln in cod.splitlines()
                                    if ln.strip().startswith(('Private Sub', 'Public Sub'))]
                except Exception as e:  # noqa: BLE001
                    t['eventos'] = f'nao lido: {e!r}'
                telas[nome] = t
            r['telas'] = telas
            for prov in ('CAP', 'Controllab'):
                lo = lx.lo_prov(prov)
                ws = lo.Parent
                r0 = lo.DataBodyRange.Row
                ult = r0 + lo.ListRows.Count - 1
                cols = {}
                for c in range(1, lo.ListColumns.Count + 1):
                    cel = ws.Cells(r0, lo.Range.Column + c - 1)
                    vazia = ws.Cells(ult + 1, lo.Range.Column + c - 1)
                    v = {'cab': lo.HeaderRowRange.Cells(1, c).Value, 'locked_linha1': bool(cel.Locked),
                         'locked_apos_tabela': bool(vazia.Locked), 'formula': bool(cel.HasFormula)}
                    try:
                        v['validacao'] = [cel.Validation.Type, cel.Validation.Formula1]
                    except Exception:  # noqa: BLE001
                        v['validacao'] = None
                    cols[col_letra(c)] = v
                body = lo.DataBodyRange.Value
                provs_col_a = {}
                for row in body:
                    provs_col_a[str(row[0])] = provs_col_a.get(str(row[0]), 0) + 1
                r['tabela_' + prov] = {'r0': r0, 'linhas': lo.ListRows.Count, 'ultima': ult, 'colunas': cols,
                                       'Z2': ws.Range('Z2').Value, 'Z3': ws.Range('Z3').Value,
                                       'provedor_na_coluna_A': provs_col_a,
                                       'protegida': bool(ws.ProtectContents)}
            r['frmCEQ_instalado'] = any(c.Name == 'frmCEQ' for c in q.wb.VBProject.VBComponents)
            set_celula(q, lx.pn, 'C3', ANALITO)
            S['B0'] = retrato(lx, 'B0')
            S['base0'] = lx.eqa_base()
            S['cap0'] = lx.tabela('CAP')
            S['ctl0'] = lx.tabela('Controllab')
            S['prov'] = S['B0']['filtros']['R4']
            S['ano'] = S['B0']['filtros']['N4']
            r['B0_filtros'] = S['B0']['filtros']
            r['B0_painel'] = S['B0']['painel']
            r['B0_linhas'] = {k: v for k, v in S['B0']['estat'].items()
                              if k.startswith(ANALITO + '|') or k.startswith(CONTROLE + '|')}
            r['EQA_Base'] = {'n': S['base0']['n'], 'carimbo': S['base0']['carimbo'],
                             'provedores': sorted({str(x[0]) for x in S['base0']['rows']})}
            S['cvp'] = {}
            for nv in NIV:
                rr = S['B0']['estat'][f'{ANALITO}|{nv}']
                S['cvp'][nv] = (rr['K'] - abs(rr['G'])) / rr['L'] if num(rr['L']) else None
            r['cvp_derivado'] = S['cvp']
            r['cvp_painel_C45_C47'] = S['B0']['painel']['C45:C47']
            ind = bias_independente(S['base0']['rows'], ANALITO, S['prov'], S['ano'])
            r['B0_independente'] = ind
            r['B0_conferencia'] = conferir_analito(S['B0']['estat'], ind, S['cvp'], ANALITO)
            return r
        p0()

        @passo('P1_preenchimento_listas')
        def p1():
            run = lambda m, *a: q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.{m}", *a)  # noqa: E731
            rods = list(run('DistintosCEQ', 'CAP', 2, True) or [])
            return {'rodadas_CAP': rods, 'analitos_CAP': list(run('DistintosCEQ', 'CAP', 4, False) or []),
                    'avaliacoes_CAP': list(run('DistintosCEQ', 'CAP', 12, False) or []),
                    'provedor_col_A_CAP': list(run('DistintosCEQ', 'CAP', 1, False) or []),
                    'rodadas_Controllab': list(run('DistintosCEQ', 'Controllab', 2, True) or []),
                    'ano_de_FH9-C 2025': run('ValorRegistrado', 'CAP', 2, 'FH9-C 2025', 3),
                    'unidade_de_' + NOME_CAP: run('ValorRegistrado', 'CAP', 4, NOME_CAP, 14),
                    'unidade_de_' + ANALITO: run('ValorRegistrado', 'CAP', 4, ANALITO, 14),
                    'ano_de_rodada_inexistente': run('ValorRegistrado', 'CAP', 2, 'XYZ', 3),
                    'regs_FH9-C_2025': len(run('RegistrosDaRodada', 'CAP', 'FH9-C 2025', '2025') or []),
                    'regs_FH9-C_2025_ano_2026': run('RegistrosDaRodada', 'CAP', 'FH9-C 2025', '2026'),
                    'regs_rodada_1': len(run('RegistrosDaRodada', 'CAP', '1', '') or [])}
        p1()

        @passo('P2_validacoes')
        def p2():
            casos = {
                'V01 duplicidade de chave REAL (FH9-A 2025, Basophils, FH9-01)':
                    dict(rod='FH9-A 2025', an='Basophils', am='FH9-01', res='999'),
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
                'V13 duplicidade com caixa/espacos diferentes (fh9-a 2025, basophils , fh9-01)':
                    dict(rod='fh9-a 2025', an='basophils ', am=' fh9-01', res='999'),
            }
            antes = lx.tabela('CAP')
            out = {k: salvar_ceq(q, **kw)[0] for k, kw in casos.items()}
            depois = lx.tabela('CAP')
            return {'retornos': out, 'tabela_identica': antes == depois,
                    'linhas_antes_depois': [antes['linhas'], depois['linhas']]}
        p2()

        @passo('P3_inclusao')
        def p3():
            a0 = lx.audit_ult()
            ret, seg = salvar_ceq(q)
            q.ex.esperar()
            S['B1'] = retrato(lx, 'B1')
            S['base1'] = lx.eqa_base()
            cap1 = lx.tabela('CAP')
            S['linha_R1'] = int(ret.split('|')[1]) if ret.startswith('OK|') else None
            row_tab = [x for x in cap1['body'] if str(x[1]) == f'{MARCA} R1']
            row_base = [x for x in S['base1']['rows'] if str(x[2]) == f'{MARCA} R1']
            ind = bias_independente(S['base1']['rows'], ANALITO, S['prov'], S['ano'])
            dif = diff_estat(S['B0']['estat'], S['B1']['estat'])
            regs = q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.RegistrosDaRodada", 'CAP', f'{MARCA} R1', '2025')
            res_ws = q.ws('EQA.CAP_Resumo').UsedRange.Value
            return {'retorno': ret, 'segundos': seg, 'linhas_tabela_antes_depois': [S['cap0']['linhas'], cap1['linhas']],
                    'linha_tabela': row_tab, 'linha_EQA_Base': row_base,
                    'carimbo': S['base1']['carimbo'], 'Z3': lx.lo_prov('CAP').Parent.Range('Z3').Value,
                    'trilha': lx.audit_desde(a0), 'lista_da_rodada': regs,
                    'K5': [S['B0']['filtros']['K5'], S['B1']['filtros']['K5']], 'independente': ind,
                    'conferencia': conferir_analito(S['B1']['estat'], ind, S['cvp'], ANALITO),
                    'mudancas_analito': so(dif, ANALITO), 'mudancas_outros_analitos': so(dif, ANALITO, False),
                    'painel_antes': S['B0']['painel'], 'painel_depois': S['B1']['painel'],
                    'resumo_CAP_mostra_teste': any(MARCA in str(c) for rr in res_ws for c in rr if c),
                    'S1': save_copy(q, 'f2h_S1.xlsm')}
        p3()

        @passo('P4_correcao')
        def p4():
            r = {'janela_mesma_chave': salvar_ceq(q, res='105')[0]}
            lo = lx.lo_prov('CAP')
            ws = lo.Parent
            lin = S.get('linha_R1')
            r['linha'] = lin
            cel = ws.Cells(lin, lo.Range.Column + 5)
            r['protegida'] = bool(ws.ProtectContents)
            r['celula_bloqueada'] = bool(cel.Locked)
            a0 = lx.audit_ult()
            q.ex.xl.EnableEvents = True
            try:
                cel.Value = 105
                r['escrita'] = 'sem desproteger (COM passa por UserInterfaceOnly)'
            except Exception as e:  # noqa: BLE001
                r['erro_escrita_protegida'] = repr(e)
                ws.Unprotect(SENHA)
                cel.Value = 105
                r['escrita'] = 'precisou desproteger (passo tecnico)'
            q.ex.esperar()
            q.ex.xl.Calculate()
            r['O_bias_na_tabela'] = ws.Cells(lin, lo.Range.Column + 14).Value
            r['Z2_apos_editar'] = ws.Range('Z2').Value
            r['Z3_apos_editar'] = ws.Range('Z3').Value
            b2a = retrato(lx, 'B2a (editado, sem consolidar)')
            base2a = lx.eqa_base()
            r['EQA_Base_linha_apos_editar'] = [x for x in base2a['rows'] if str(x[2]) == f'{MARCA} R1']
            r['estat_mudou_sem_consolidar'] = diff_estat(S['B1']['estat'], b2a['estat'])
            r['painel_mudou_sem_consolidar'] = b2a['painel'] != S['B1']['painel']
            r['trilha_da_edicao'] = lx.audit_desde(a0)
            acao = None
            for shp in ws.Shapes:
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
                macro = 'mEQA.' + macro
            a1 = lx.audit_ult()
            r['segundos_consolidar'] = consolidar(q, macro)
            r['macro_executada'] = macro
            S['B2'] = retrato(lx, 'B2')
            S['base2'] = lx.eqa_base()
            r['trilha_consolidar'] = lx.audit_desde(a1)
            r['EQA_Base_linha'] = [x for x in S['base2']['rows'] if str(x[2]) == f'{MARCA} R1']
            ind = bias_independente(S['base2']['rows'], ANALITO, S['prov'], S['ano'])
            r['independente'] = ind
            r['conferencia'] = conferir_analito(S['B2']['estat'], ind, S['cvp'], ANALITO)
            dif = diff_estat(S['B1']['estat'], S['B2']['estat'])
            r['mudancas_analito'] = so(dif, ANALITO)
            r['mudancas_outros_analitos'] = so(dif, ANALITO, False)
            r['K5'] = S['B2']['filtros']['K5']
            r['painel'] = S['B2']['painel']
            r['Z2_apos_consolidar'] = ws.Range('Z2').Value
            r['lista_da_rodada'] = q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.RegistrosDaRodada", 'CAP', f'{MARCA} R1',
                                               '2025')
            r['S2'] = save_copy(q, 'f2h_S2.xlsm')
            return r
        p4()

        @passo('P5_isolamento')
        def p5():
            r = {}
            ref = S['B2']

            def comparar(rotulo, acao, filtro):
                a0 = lx.audit_ult()
                ret = acao()
                q.ex.esperar()
                b = retrato(lx, rotulo, painel=False)
                base = lx.eqa_base()
                dif = diff_estat(ref['estat'], b['estat'])
                return {'retorno': ret, 'carimbo': base['carimbo'], 'K5': b['filtros']['K5'],
                        'n_mudancas': len(dif), 'mudancas': dif[:30],
                        'trilha': [t.get('Ação') or t.get('ACAO') or t for t in lx.audit_desde(a0)],
                        'linha_EQA_Base': [y for y in base['rows'] if filtro(y)]}

            r['I2_controllab'] = comparar('I2 Controllab', lambda: salvar_ceq(
                q, prov='Controllab', rod=f'{MARCA} CTL1', an=ANALITO, am='TA-C1', res='120', alvo='100'),
                lambda y: str(y[2]) == f'{MARCA} CTL1')
            r['I3_outro_ano'] = comparar('I3 ano 2026', lambda: salvar_ceq(
                q, rod=f'{MARCA} R2', ano='2026', am='TA-R2', res='130', alvo='100'),
                lambda y: str(y[2]) == f'{MARCA} R2')
            fora = NOME_CAP.rsplit(' ', 1)[0] if ' ' in NOME_CAP else NOME_CAP + ' X'
            r['I4_nome_fora_do_mapa'] = comparar('I4 nome fora do mapa', lambda: salvar_ceq(
                q, an=fora, am='TA-03', res='150', alvo='100'), lambda y: str(y[5]) == 'TA-03')
            r['I4_nome_usado'] = fora
            r['I5_sem_alvo'] = comparar('I5 sem alvo', lambda: salvar_ceq(q, am='TA-04', res='100', alvo=''),
                                        lambda y: str(y[5]) == 'TA-04')
            S['base5'] = lx.eqa_base()
            ref5 = retrato(lx, 'ref5', painel=True)
            S['ref5'] = ref5
            r['painel_ref5'] = ref5['painel']
            # I5b: o que a linha do analito mostra em S (limites) e AC depois do sem-alvo
            r['I5_linha_analito'] = {k: {c: v[c] for c in ('G', 'S', 'AC', 'AD')}
                                     for k, v in ref5['estat'].items() if k.startswith(ANALITO + '|')}
            est = lx.est
            n4 = est.Range('N4').Value
            for ano in (2026, 'TODOS'):
                set_celula(q, est, 'N4', ano)
                b = retrato(lx, f'N4={ano}', painel=False)
                ind = bias_independente(S['base5']['rows'], ANALITO, S['prov'], ano)
                ctl = bias_independente(S['base5']['rows'], CONTROLE, S['prov'], ano)
                dif = diff_estat(ref5['estat'], b['estat'])
                r[f'I6_N4_{ano}'] = {
                    'K5': b['filtros']['K5'], 'independente': ind,
                    'conferencia': conferir_analito(b['estat'], ind, S['cvp'], ANALITO),
                    'controle_independente': ctl,
                    'controle_conferencia': conferir_analito(b['estat'], ctl, {}, CONTROLE),
                    'n_mudancas': len(dif),
                    'analitos_G_numerico': sorted({k.split('|')[0] for k, v in b['estat'].items() if num(v['G'])}),
                    'G_de_todos': {k: v['G'] for k, v in b['estat'].items()}}
            set_celula(q, est, 'N4', n4)
            b = retrato(lx, 'N4 restaurado', painel=False)
            r['I6_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            # I7a periodo do CIQ (Ano De/Ate = 2026): o bias do CEQ nao pode mudar
            b3, d3 = est.Range('B3').Value, est.Range('D3').Value
            set_celula(q, est, 'B3', 2026)
            set_celula(q, est, 'D3', 2026)
            try:
                q.run('mEstatistica.AtualizarEstatisticaAba', teto=900)
            except Exception as e:  # noqa: BLE001
                r['aviso_I7a'] = repr(e)
            q.ex.xl.Calculate()
            b = retrato(lx, 'B3:D3=2026', painel=True)
            dif = diff_estat(ref5['estat'], b['estat'])
            r['I7a_periodo_CIQ_2026'] = {
                'G_T_AC_AD_iguais_todos': all(d['col'] not in ('G', 'T', 'AC', 'AD', 'R', 'S') for d in dif),
                'colunas_que_mudaram': sorted({d['col'] for d in dif}),
                'linha_analito_depois': {k: {c: v[c] for c in ('C', 'D', 'E', 'F', 'G', 'H', 'L', 'T', 'AC')}
                                         for k, v in b['estat'].items() if k.startswith(ANALITO + '|')},
                'painel': b['painel']}
            set_celula(q, est, 'B3', b3)
            set_celula(q, est, 'D3', d3)
            try:
                q.run('mEstatistica.AtualizarEstatisticaAba', teto=900)
            except Exception as e:  # noqa: BLE001
                r['aviso_I7a_b'] = repr(e)
            q.ex.xl.Calculate()
            b = retrato(lx, 'B3:D3 restaurado', painel=False)
            r['I7a_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            # I7b lote do Painel (H3): o bias do CEQ nao pode mudar
            h3 = lx.pn.Range('H3').Value
            nf = lx.pn.Range('H3').NumberFormat
            novo = LOTE_ALT if nf == '@' or not isinstance(h3, str) else "'" + LOTE_ALT
            set_celula(q, lx.pn, 'H3', novo)
            q.ex.xl.Calculate()
            b = retrato(lx, f'H3={LOTE_ALT}', painel=True)
            dif = diff_estat(ref5['estat'], b['estat'])
            r['I7b_lote'] = {
                'de_para': [h3, lx.pn.Range('H3').Value, nf],
                'G_T_AC_AD_iguais_todos': all(d['col'] not in ('G', 'T', 'AC', 'AD', 'R', 'S') for d in dif),
                'colunas_que_mudaram': sorted({d['col'] for d in dif}),
                'linha_analito_depois': {k: {c: v[c] for c in ('C', 'D', 'E', 'F', 'G', 'H', 'L', 'T', 'AC')}
                                         for k, v in b['estat'].items() if k.startswith(ANALITO + '|')},
                'painel': b['painel'], 'B4': b['filtros']['B4']}
            volta = h3 if nf == '@' or not isinstance(h3, str) else "'" + h3
            set_celula(q, lx.pn, 'H3', volta)
            q.ex.xl.Calculate()
            b = retrato(lx, 'H3 restaurado', painel=True)
            r['I7b_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            r['I7b_painel_restaurado_identico'] = b['painel'] == ref5['painel']
            # I8 provedor global R4 = Controllab
            r4 = est.Range('R4').Value
            a0 = lx.audit_ult()
            set_celula(q, est, 'R4', 'Controllab')
            b = retrato(lx, 'R4=Controllab', painel=True)
            dif = diff_estat(ref5['estat'], b['estat'])
            r['I8_provedor_controllab'] = {
                'R4': est.Range('R4').Value, 'K5': b['filtros']['K5'],
                'analitos_G_numerico': sorted({k.split('|')[0] for k, v in b['estat'].items() if num(v['G'])}),
                'linha_analito': {k: {c: v[c] for c in ('G', 'H', 'L', 'T', 'AC', 'AD', 'R', 'S')}
                                  for k, v in b['estat'].items() if k.startswith(ANALITO + '|')},
                'n_linhas_mudaram': len({d['chave'] for d in dif}),
                'painel': b['painel'], 'trilha': lx.audit_desde(a0)}
            set_celula(q, est, 'R4', r4)
            b = retrato(lx, 'R4 restaurado', painel=True)
            r['I8_restaurado_identico'] = diff_estat(ref5['estat'], b['estat']) == []
            return r
        p5()

        @passo('P6_permissao')
        def p6():
            r = {}
            q.run('mSeguranca.Logout', teto=300)
            r['sessao_apos_logout'] = q.run('mSeguranca.SessaoAtiva', teto=60)
            r['usuario_celula'] = q.nome('currentUser').Value
            r['salvar_sem_sessao'] = salvar_ceq(q, am='TA-SEMLOGIN', res='100', alvo='100')[0]
            r['aviso'] = 'AbrirCEQ nao foi chamado (janela modal travaria a automacao); o nucleo e o mesmo'
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
                    if any(MARCA in str(c or '') for c in row[:18]) or str(row[4] or '').startswith('TA-'):
                        rr = lo.DataBodyRange.Row + i
                        apagadas.append([prov, rr, list(row[:6])])
                        for c in list(range(2, 15)) + [17, 18]:
                            ws.Cells(rr, lo.Range.Column + c - 1).ClearContents()
            r['linhas_limpas'] = apagadas
            r['segundos_consolidar'] = consolidar(q, 'mEQA.AtualizarEQABase')
            b = retrato(lx, 'B_final', painel=True)
            base = lx.eqa_base()
            dif = diff_estat(S['B0']['estat'], b['estat'])
            r['estat_identica_B0'] = dif == []
            r['diferencas_vs_B0'] = dif[:40]
            r['K5_igual_B0'] = b['filtros']['K5'] == S['B0']['filtros']['K5']
            r['EQA_Base_linhas_iguais_B0'] = base['rows'] == S['base0']['rows']
            r['EQA_Base_n'] = [S['base0']['n'], base['n']]
            r['carimbo_final'] = base['carimbo']
            r['painel_igual_B0'] = b['painel'] == S['B0']['painel']
            if not r['painel_igual_B0']:
                r['painel_dif'] = {k: [S['B0']['painel'][k], b['painel'][k]] for k in b['painel']
                                   if b['painel'][k] != S['B0']['painel'][k]}
            r['linhas_tabela_CAP_final'] = lx.tabela('CAP')['linhas']
            return r
        p7()
    finally:
        q.fechar(salvar=False)
        OUT['sha_copia_depois'] = sha(COPIA)
        gravar()
        log('fim; copia intacta =', OUT['sha_copia_antes'] == OUT['sha_copia_depois'])


if __name__ == '__main__':
    main()
