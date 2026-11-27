# -*- coding: utf-8 -*-
"""repro_e49_adr103.py -- Tester/QA 27/11/2026, VALIDAR entrega 49 (ADR-103). Script PROPRIO do Tester: nao reaproveita
qa_adr103.py nem qa_alvos_manual.py (so a infraestrutura qa_final.QA). Tudo roda em COPIAS em %TEMP%\\qcw\\tester49\\;
nada e salvo; a origem nunca e aberta pelo Excel.

Uso: python -u repro_e49_adr103.py <Bioquimica|Hematologia> <origem.xlsm> <saida.json>

Oraculos (escritos aqui a partir do texto do pedido/ADR-103, sem chamar VBA de calculo):
  - Sigma: sigma >= 6 "Excelente — Classe mundial"/"Excelente"; 5 < s < 6 "Excelente"/"Alto desempenho";
    4 < s <= 5 "Bom"/"Desempenho adequado"; 3 <= s <= 4 "Marginal"/"Atenção"; s < 3 "Insatisfatório"/
    "método mais instável ou tendencioso"; vazio/texto = "".
  - Alvos: Python statistics (fmean, stdev n-1) sobre a tblCQ_Final (PARTICIPA_ESTATISTICA = SIM, lote e analito
    exatos sem caixa, todos os equipamentos, dia de DATA_HORA entre De e Ate inclusive; vazio = sem limite).
  - Q4: avaliador proprio da cadeia IF(COUNTIF(stNivel,"padrao")>0,"RESULTADO",...) lida da celula.

Casos (APROVADO / REPROVADO / NAO_EXECUTADO):
  T1.x regra Sigma (VBA x oraculo nas fronteiras; Estatistica M; Painel J7:J9 em todos os analitos com Sigma)
  T2.x Painel compacto (C4:D4, H4:I4, AB/stNivel, J6:J9, Q4 coerente, #REF!)
  T3.x tabela de digitacao N114 e TransferirReferenciaCore
  T4.x AplicarAlvosVarios x AplicarAlvos isolado (copia A x copia B), trilha, cancelar
  Y00  origens intactas (SHA-256)
"""
import collections
import datetime as dt
import fnmatch
import hashlib
import json
import os
import re
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
from instalar_integracao import SENHA  # noqa: E402   (so por import; nunca impressa)

PASTA = os.path.join(os.environ.get('TEMP', '.'), 'qcw', 'tester49')
TAG = 'TESTE-AUDITORIA E49'
R_ENT, R_CAB, R_REF1, N_REF, D = 98, 114, 115, 40, 13
RES = []
ERR_REF = -2146826265


def reg(cod, estado, desc, evid):
    RES.append({'id': cod, 'estado': estado, 'descricao': desc, 'evidencia': evid})
    print(f'{estado:13s} {cod} {desc} -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:900], flush=True)


def ok(cod, cond, desc, evid):
    reg(cod, 'APROVADO' if cond else 'REPROVADO', desc, evid)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and v == v and not (isinstance(v, int) and v < -2146820000)


def txt(v):
    if isinstance(v, float) and v == v and v == int(v):
        return str(int(v))
    return str(v if v is not None else '').strip()


def U(v):
    return txt(v).upper()


def perto(a, b, tol=1e-9):
    return num(a) and num(b) and abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


# ------------------------------------------------------------------ oraculo Sigma (do texto do pedido)
def sigma_oraculo(s):
    if s is None or isinstance(s, str) or isinstance(s, bool):
        return ('', '')
    if s >= 6:
        return ('Excelente — Classe mundial', 'Excelente')
    if 5 < s < 6:
        return ('Excelente', 'Alto desempenho')
    if 4 < s <= 5:
        return ('Bom', 'Desempenho adequado')
    if 3 <= s <= 4:
        return ('Marginal', 'Atenção')
    return ('Insatisfatório', 'método mais instável ou tendencioso')


# ------------------------------------------------------------------ oraculo dos alvos
def dia(v):
    return v.replace(tzinfo=None).date() if isinstance(v, dt.datetime) else None


def nivel(r):
    try:
        return int(float(r.get('NIVEL')))
    except (TypeError, ValueError):
        return 0


def ativo(r):
    return U(r.get('PARTICIPA_ESTATISTICA')) == 'SIM' and num(r.get('RESULTADO'))


def apurar(rows, lote, an, ini, fim, nlv):
    out = {}
    for nv in range(1, nlv + 1):
        xs = []
        for r in rows:
            if not ativo(r) or U(r.get('LOTE')) != U(lote) or U(r.get('ANALITO')) != U(an) or nivel(r) != nv:
                continue
            d = dia(r.get('DATA_HORA'))
            if ini is not None or fim is not None:
                if d is None or (ini is not None and d < ini) or (fim is not None and d > fim):
                    continue
            xs.append(float(r['RESULTADO']))
        out[nv] = (len(xs), statistics.fmean(xs) if xs else None, statistics.stdev(xs) if len(xs) >= 2 else None)
    return out


# ------------------------------------------------------------------ oraculo do Q4 (avaliador proprio da cadeia)
def q4_oraculo(formula, status):
    pares = re.findall(r'COUNTIF\(\$AB\$7:\$AB\$\d+,"([^"]*)"\)>0,"([^"]*)"', formula)
    for pad, res in pares:
        if any(fnmatch.fnmatchcase(U(s), U(pad)) for s in status):
            return res, pares
    fim = re.findall(r',"([^"]*)"\)+\s*$', formula)
    return (fim[0] if fim else None), pares


# ------------------------------------------------------------------ sessao Excel numa copia
class Sessao:
    def __init__(self, produto, origem, rotulo):
        os.makedirs(PASTA, exist_ok=True)
        self.copia = os.path.join(PASTA, f'{rotulo}_{produto}_{time.strftime("%Y%m%d_%H%M%S")}.xlsm')
        shutil.copy2(origem, self.copia)
        print('COPIA', self.copia, flush=True)
        self.q = QA(produto, self.copia)
        self.xl, self.nlv = self.q.ex.xl, self.q.nlv
        self.p, self.ls = self.q.ws('Painel'), self.q.ws('LotesStore')

    def run(self, m, *a, teto=900):
        return self.q.run(m if '.' in m else 'mAlvosRef.' + m, *a, teto=teto)

    def resp(self, s):
        self.run('DefinirRespostasTeste', s, teto=60)

    def secao(self, lote, ini, fim):
        self.xl.EnableEvents = False
        try:
            try:
                self._sec(lote, ini, fim)
            except Exception:                 # noqa: BLE001
                self.p.Unprotect(SENHA)
                self._sec(lote, ini, fim)
        finally:
            self.xl.EnableEvents = True

    def _sec(self, lote, ini, fim):
        self.p.Range(f'B{R_ENT}').Value = lote
        self.p.Range(f'E{R_ENT}').Value = dt.datetime.combine(ini, dt.time()) if ini else None
        self.p.Range(f'G{R_ENT}').Value = dt.datetime.combine(fim, dt.time()) if fim else None

    def ult_ls(self):
        return self.ls.Cells(self.ls.Rows.Count, 16).End(-4162).Row

    def foto_ls(self):
        return [tuple(r) for r in self.ls.Range(f'A1:AO{self.ult_ls() + 2}').Value]

    def aud_n(self):
        return int(self.run('mAuditoria.UltimaLinhaAudit', teto=60))

    def aud(self, r0, acao):
        al = self.q.ws('Audit_Log')
        r1 = self.aud_n()
        if r1 <= r0:
            return []
        return [tuple(l) for l in al.Range(al.Cells(r0 + 1, 1), al.Cells(r1, 25)).Value
                if txt(l[6]) == acao]

    def catalogo(self):
        return [txt(v[0]) for v in self.q.espec().Range('A4:A43').Value]

    def escolher_analito(self, an):
        cat = [U(x) for x in self.catalogo()]
        try:
            self.p.Unprotect(SENHA)
        except Exception:                     # noqa: BLE001
            pass
        self.xl.EnableEvents = False
        self.p.Range('B3').Value = cat.index(U(an)) + 1
        self.xl.EnableEvents = True
        self.q.run('mEstatistica.PainelMudou', teto=900)
        self.xl.Calculate()

    def fechar(self):
        try:
            self.resp('')
        except Exception:                     # noqa: BLE001
            pass
        self.q.fechar(salvar=False)


def norm_linha(l):
    """Linha A:AO da LotesStore sem o carimbo de hora: W (origem) com a data/hora trocada; AE/AJ/AO (quando) fora."""
    out = []
    for j, v in enumerate(l):
        if j in (20, 30, 35, 40):           # U (atualizado em) e AE/AJ/AO (quando): carimbos de hora
            out.append('<quando>' if v not in (None, '') else None)
        elif j == 22:
            out.append(re.sub(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}', '<TS>', txt(v)))
        else:
            out.append(v)
    return tuple(out)


# ------------------------------------------------------------------ execucao
def executar(produto, origem, saida):
    h0 = sha(origem)
    prods = {k: os.path.join(RAIZ, f'QC_{k}.xlsm') for k in ('Bioquimica', 'Hematologia')}
    hp0 = {k: sha(v) for k, v in prods.items()}
    print('ORIGEM_SHA', h0[:12], {k: v[:12] for k, v in hp0.items()}, flush=True)
    A = None
    plano = {}
    try:
        A = Sessao(produto, origem, 'A')
        q, p, nlv, xl = A.q, A.p, A.nlv, A.xl
        bio = produto.startswith('Bio')
        print('PROTEGIDO_INICIO', p.ProtectContents, flush=True)

        # ============================================================ T2 Painel compacto (somente leitura)
        try:
            c4 = (txt(p.Range('C4').Formula), p.Range('C4').MergeCells, txt(p.Range('D4').Formula), p.Range('D4').MergeCells,
                  txt(p.Range('C4').Value), txt(p.Range('D4').Value))
            ok('T2.1', c4 == ('', False, '', False, '', ''), 'C4:D4 vazio (formula e valor) e nao mesclado', {'C4_D4': c4})
            h = p.Range('H4').MergeArea
            ok('T2.2', p.Range('H4').MergeCells and h.Address == '$H$4:$I$4' and h.HorizontalAlignment == -4108
               and h.VerticalAlignment == -4108 and txt(p.Range('H4').Formula).startswith('='),
               'H4:I4 mesclada, centralizada (horizontal e vertical), formula do aviso em H4',
               {'merge': h.Address, 'H': h.HorizontalAlignment, 'V': h.VerticalAlignment,
                'H4': txt(p.Range('H4').Formula)[:60]})
            ab = [txt(p.Range(f'AB{6 + i}').Formula) for i in range(1, nlv + 1)]
            nm = txt(q.wb.Names('stNivel').RefersTo).replace(' ', '')
            jf = [txt(p.Range(f'J{6 + i}').Formula) for i in range(1, nlv + 1)]
            kf = [txt(p.Range(f'K{6 + i}').Formula) for i in range(1, nlv + 1)]
            ab_vazio = txt(p.Range(f'AB{7 + nlv}').Formula) if nlv < 3 else ''
            ok('T2.3', p.Columns('AB').Hidden and all(f.startswith('=IF(') for f in ab)
               and nm == f'=Painel!$AB$7:$AB${6 + nlv}' and txt(p.Range('J6').Value) == 'Classe Sigma'
               and all('Estatística!$M$' in f for f in jf) and all(f == '' for f in kf) and ab_vazio == ''
               and not any('DESATUAL.' in f for f in jf),
               'status por nivel em AB oculta (stNivel); J6 "Classe Sigma" e J7.. formula da Estatistica M; K7.. vazio',
               {'AB_oculta': p.Columns('AB').Hidden, 'stNivel': nm, 'AB7': ab[0][:70], 'J6': txt(p.Range('J6').Value),
                'J7': jf[0][:110], 'K': kf})
            q4f = txt(p.Range('Q4').Formula)
            ok('T2.4', f'$AB$7:$AB${6 + nlv}' in q4f and '$J$7' not in q4f,
               'Q4 (status geral) le o status por nivel em AB (nao mais J)', {'Q4': q4f[:200]})
            plano['q4f'] = q4f
            usado = p.UsedRange
            fs = usado.Formula
            vs = usado.Value
            r0_, c0_ = usado.Row, usado.Column
            ref_f = [(r0_ + i, c0_ + j) for i, l in enumerate(fs) for j, x in enumerate(l) if '#REF' in txt(x).upper()]
            ref_v = [(r0_ + i, c0_ + j) for i, l in enumerate(vs) for j, x in enumerate(l) if x == ERR_REF]
            nomes_ref = [n.Name for n in q.wb.Names if '#REF' in txt(n.RefersTo).upper()]
            ok('T2.5', not ref_f and not ref_v, 'nenhum #REF! (formula ou valor) no Painel inteiro (UsedRange)',
               {'usado': usado.Address, 'formulas_REF': ref_f[:8], 'valores_REF': ref_v[:8],
                'nomes_do_livro_com_REF (informativo)': nomes_ref[:8]})
        except Exception as e:                # noqa: BLE001
            reg('T2.x', 'REPROVADO', 'Painel compacto: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-800:]})

        # ============================================================ T1a regra Sigma (VBA x oraculo)
        try:
            pts = [2.999999, 3, 3.5, 3.999999, 4, 4.000001, 4.5, 5, 5.000001, 5.5, 5.999999, 6, 6.000001, 7, 0.0, -1.0,
                   2.0, 100.0]
            dif, tab = [], []
            for s in pts:
                c = txt(q.run('mQualidade.ClassificarSigma', float(s), teto=30))
                st = txt(q.run('mQualidade.StatusSigma', float(s), teto=30))
                tab.append((s, c, st))
                if (c, st) != sigma_oraculo(float(s)):
                    dif.append({'sigma': s, 'vba': (c, st), 'oraculo': sigma_oraculo(float(s))})
            for v in ('', 'abc', '   '):
                c = txt(q.run('mQualidade.ClassificarSigma', v, teto=30))
                st = txt(q.run('mQualidade.StatusSigma', v, teto=30))
                tab.append((repr(v), c, st))
                if (c, st) != ('', ''):
                    dif.append({'sigma': repr(v), 'vba': (c, st), 'oraculo': ('', '')})
            try:
                c = txt(q.run('mQualidade.ClassificarSigma', None, teto=30))
                st = txt(q.run('mQualidade.StatusSigma', None, teto=30))
                tab.append(('None(Empty)', c, st))
                if (c, st) != ('', ''):
                    dif.append({'sigma': 'Empty', 'vba': (c, st)})
            except Exception as e:            # noqa: BLE001
                tab.append(('None(Empty)', 'excecao COM', repr(e)[:80]))
            v4 = next(x for x in tab if x[0] == 4)
            ok('T1.1', not dif and v4[1] == 'Marginal',
               'ClassificarSigma/StatusSigma = oraculo literal nas fronteiras (sigma = 4 Marginal, 5 Bom, 6 Classe '
               'mundial, 3 Marginal); vazio/texto = ""', {'diferencas': dif, 'tabela': tab})
        except Exception as e:                # noqa: BLE001
            reg('T1.1', 'REPROVADO', 'regra Sigma via Application.Run: excecao', {'erro': repr(e)})

        def est_linhas():
            e = q.ws('Estatística')
            xl.Calculate()
            cat = {U(x) for x in A.catalogo() if x}
            v = e.Range('A14:AB133').Value
            out = {}
            for i, l in enumerate(v):
                # so as linhas da tabela (chave AB = analito|nivel); a 1a ocorrencia (as listas abaixo, ex. "margem
                # critica" A106.., repetem analito/nivel sem Sigma e nao sao linhas da tabela)
                if U(l[0]) in cat and num(l[1]) and U(l[27]) == f'{U(l[0])}|{int(l[1])}':
                    out.setdefault((U(l[0]), int(l[1])), (14 + i, l[11], txt(l[12]), txt(l[27])))
            return out

        def checar_m(cod, momento):
            el = est_linhas()
            com = [(k, x) for k, x in el.items() if num(x[1])]
            dif = [{'linha': x[0], 'chave': k, 'L': x[1], 'M': x[2], 'oraculo': sigma_oraculo(float(x[1]))[0]}
                   for k, x in com if x[2] != sigma_oraculo(float(x[1]))[0]]
            sem = [(k, x[0], x[1], x[2]) for k, x in el.items() if not num(x[1]) and x[2] not in ('', '-')]
            faixas = collections.Counter(sigma_oraculo(float(x[1]))[0] for _, x in com)
            ok(cod, len(com) > 0 and not dif and not sem,
               f'Estatistica M = oraculo(L) em TODAS as linhas com Sigma ({momento}); sem Sigma -> "" ou "-"',
               {'linhas_com_sigma': len(com), 'por_faixa': faixas, 'diferencas': dif[:10], 'M_sem_sigma': sem[:6]})
            return el

        el0 = checar_m('T1.2', 'estado entregue')
        if bio:
            try:
                e = q.ws('Estatística')
                bloco = [[txt(x) for x in l] for l in e.Range('A96:F103').Value]
                reg('T1.2b', 'APROVADO' if any('Insatisfat' in ' '.join(l) for l in bloco) else 'REPROVADO',
                    'Bio Estatistica A96:F103 (resumo) traz os rotulos novos (informativo: contagens no evidencia)',
                    {'bloco': bloco, 'contagem_oraculo': collections.Counter(
                        sigma_oraculo(float(x[1]))[0] for x in el0.values() if num(x[1]))})
            except Exception as e:            # noqa: BLE001
                reg('T1.2b', 'NAO_EXECUTADO', 'resumo Bio', {'erro': repr(e)})

        # ============================================================ escolha do lote e dos analitos (T4/T3)
        rows = q.ler('tblCQ_Final')
        lim = txt(A.run('LimitesNTexto', teto=60)).split('|')
        rec = int(lim[0])
        cat = [x for x in A.catalogo() if x]
        cat_u = {U(x): x for x in cat}
        u = A.ult_ls()
        chaves = collections.Counter(U(v[0]) for v in A.ls.Range(f'P2:P{u}').Value if txt(v[0]))
        cont = collections.Counter((U(r['LOTE']), U(r['ANALITO']), nivel(r)) for r in rows if ativo(r))
        por_lote = collections.defaultdict(list)
        for lt in {k[0] for k in cont}:
            for an in cat_u:
                if chaves.get(f'{lt}|{an}') == 1 and all(cont.get((lt, an, nv), 0) >= 2 for nv in range(1, nlv + 1)):
                    por_lote[lt].append(an)
        cand = sorted(((len(v), lt) for lt, v in por_lote.items()), reverse=True)
        if not cand or cand[0][0] < 5:
            raise RuntimeError(f'pre-condicao: nenhum lote com >= 5 analitos gravaveis: {cand[:3]}')
        LT = cand[0][1]
        lote = next(txt(r['LOTE']) for r in rows if U(r.get('LOTE')) == LT)
        ordem = sorted(por_lote[LT], key=lambda a: -sum(cont.get((LT, a, nv), 0) for nv in range(1, nlv + 1)))
        set1 = ordem[:3]
        set2 = ordem[3:5]
        datas2 = sorted({dia(r.get('DATA_HORA')) for r in rows if ativo(r) and U(r['LOTE']) == LT
                         and U(r['ANALITO']) in set2 and dia(r.get('DATA_HORA'))})
        ini2, fim2 = datas2[len(datas2) // 5], datas2[(4 * len(datas2)) // 5]
        ora1 = {a: apurar(rows, LT, a, None, None, nlv) for a in set1}
        ora2 = {a: apurar(rows, LT, a, ini2, fim2, nlv) for a in set2}
        plano.update({'lote': lote, 'set1': [cat_u[a] for a in set1], 'set2': [cat_u[a] for a in set2],
                      'ini2': str(ini2), 'fim2': str(fim2), 'rec': rec})
        ok('T4.0', len(set1) == 3 and len(set2) == 2 and all(ora2[a][nv][0] >= 2 for a in set2 for nv in ora2[a]),
           'pre-condicoes: lote com >= 5 analitos de chave unica e N >= 2 por nivel; set1 (periodo inteiro, modo 1), '
           'set2 (janela, modo 2)', {**plano, 'N_set1': {a: [ora1[a][nv][0] for nv in ora1[a]] for a in set1},
                                     'N_set2': {a: [ora2[a][nv][0] for nv in ora2[a]] for a in set2}})

        def linha_de(sess, an):
            u_ = sess.ult_ls()
            return [i for i, v in enumerate(sess.ls.Range(f'P2:P{u_}').Value, 2) if U(v[0]) == f'{LT}|{U(an)}'][0]

        # ============================================================ T4 varios (copia A)
        res_A = {}
        try:
            A.secao(lote, None, None)
            antes = A.foto_ls()
            r0 = A.aud_n()
            A.resp('0')
            c1 = txt(A.run('AplicarAlvosVarios', '|'.join(cat_u[a] for a in set1)))
            A.resp('1|NAO')
            c2 = txt(A.run('AplicarAlvosVarios', '|'.join(cat_u[a] for a in set1)))
            perg_c = txt(A.run('UltimaPergunta', teto=60))
            A.resp('')
            ok('T4.1', c1 == 'CANCELADO|transferencia' and c2 == 'CANCELADO|confirmacao' and A.foto_ls() == antes
               and not A.aud(r0, 'PARAMETRO_LOTE_CALCULADO') and not A.aud(r0, 'PARAMETRO_LOTE_CALCULADO_VARIOS'),
               'cancelar o modo e cancelar a confirmacao: nada gravado (LotesStore A1:AO identica, sem trilha)',
               {'modo': c1, 'confirmacao': c2, 'pergunta': perg_c[:300]})
            for rot, st, (ini, fim), modo, ora in (('set1', set1, (None, None), '1', ora1),
                                                   ('set2', set2, (ini2, fim2), '2', ora2)):
                A.secao(lote, ini, fim)
                antes = A.foto_ls()
                r0 = A.aud_n()
                A.resp(f'{modo}|SIM|SIM')
                r = txt(A.run('AplicarAlvosVarios', '|'.join(cat_u[a] for a in st)))
                perg = txt(A.run('UltimaPergunta', teto=60))
                A.resp('')
                depois = A.foto_ls()
                lin = {a: linha_de(A, a) for a in st}
                t1 = A.aud(r0, 'PARAMETRO_LOTE_CALCULADO')
                t2 = A.aud(r0, 'PARAMETRO_LOTE_CALCULADO_VARIOS')
                outras = [i + 1 for i, (x, y) in enumerate(zip(antes, depois)) if x != y and i + 1 not in lin.values()]
                dif = []
                for a in st:
                    l = depois[lin[a] - 1]
                    l0 = antes[lin[a] - 1]
                    for nv in range(1, nlv + 1):
                        n, m, s = ora[a][nv]
                        cm, cs = 2 + (nv - 1) * 2, 3 + (nv - 1) * 2
                        if modo == '1':
                            if not (perto(l[cm], m) and perto(l[cs], s)):
                                dif.append({'an': a, 'nv': nv, 'LS': (l[cm], l[cs]), 'oraculo': (m, s)})
                        else:   # so a Media; o DP fica como estava
                            if not (perto(l[cm], m) and l[cs] == l0[cs]):
                                dif.append({'an': a, 'nv': nv, 'LS': (l[cm], l[cs]), 'oraculo_media': m, 'DP_antes': l0[cs]})
                        if l[26 + (nv - 1) * 5] != n:
                            dif.append({'an': a, 'nv': nv, 'N_AA': l[26 + (nv - 1) * 5], 'oraculo_N': n})
                por_an = collections.Counter(U(x[14]) for x in t1)
                res_A[rot] = {'retorno': r, 'linhas': {a: norm_linha(depois[lin[a] - 1]) for a in st},
                              'trilha': sorted((U(x[14]),) + tuple(txt(y) for y in x[6:14]) for x in t1),
                              'por_an': dict(por_an)}
                ok(f'T4.2{rot[-1]}', r.startswith(f'OK|{len(st)}|0|0|0|') and not dif and not outras
                   and set(por_an) == set(st) and len(t2) == 1 and f'Analitos selecionados: {len(st)}' in perg
                   and (ini is None or ini.strftime('%d/%m/%Y') in perg),
                   f'AplicarAlvosVarios {rot} (modo {modo}): cada analito = SEU oraculo; so as linhas escolhidas mudam; '
                   f'PARAMETRO_LOTE_CALCULADO de cada analito + 1 _VARIOS',
                   {'retorno': r[:250], 'diferencas': dif[:6], 'outras_linhas_alteradas': outras[:8],
                    'trilha_por_analito': dict(por_an), 'resumo_VARIOS': len(t2),
                    'resumo_texto': txt(t2[0][-1] if t2 else '')[:200] if t2 else None, 'pergunta': perg[:350]})
        except Exception as e:                # noqa: BLE001
            reg('T4.2', 'REPROVADO', 'AplicarAlvosVarios: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-900:]})

        # ============================================================ T3 digitacao N114 e transferencia
        try:
            A.secao(lote, None, None)
            A.run('AtualizarVistaReferencia', teto=600)
            ncol = 2 + nlv * 3                # coluna da Fonte na grade (A=1)
            cabg = [txt(p.Cells(R_CAB, c).Value) for c in range(1, ncol + 1)]
            cabd = [txt(p.Cells(R_CAB, c + D).Value) for c in range(1, ncol + 1)]
            ang = [txt(p.Cells(r, 1).Value) for r in range(R_REF1, R_REF1 + N_REF)]
            andg = [txt(p.Cells(r, 1 + D).Value) for r in range(R_REF1, R_REF1 + N_REF)]
            dig = p.Range(p.Cells(R_REF1, 2 + D), p.Cells(R_REF1 + N_REF - 1, ncol + D))
            travadas = sum(1 for c in dig.Cells if c.Locked)
            ok('T3.1', cabd == cabg and andg == ang and any(ang),
               f'N114 com o mesmo cabecalho da A114 ({ncol} colunas) e analitos espelhados 115..154',
               {'cab_grade': cabg, 'cab_dig': cabd, 'analitos_ok': andg == ang,
                'protegida': p.ProtectContents, 'celulas_digitacao_travadas': travadas})
            rb = q.ws('RefBulaStore')

            def foto_rb():
                u_ = rb.Cells(rb.Rows.Count, 3).End(-4162).Row
                return [tuple(x) for x in rb.Range(f'A1:O{max(u_, 1) + 1}').Value]

            def grade():
                return [tuple(x) for x in p.Range(p.Cells(R_REF1, 1), p.Cells(R_REF1 + N_REF - 1, ncol)).Formula]

            def digf():
                return [tuple(x) for x in p.Range(p.Cells(R_REF1, 1 + D), p.Cells(R_REF1 + N_REF - 1, ncol + D)).Formula]

            def pôr(r, c, formula):
                cel = p.Cells(r, c + D)
                try:
                    cel.Formula = formula
                except Exception:             # noqa: BLE001
                    p.Unprotect(SENHA)
                    cel.Formula = formula

            def limpar():
                try:
                    dig.ClearContents()
                except Exception:             # noqa: BLE001
                    p.Unprotect(SENHA)
                    dig.ClearContents()

            def transferir(resposta):
                A.resp(resposta)
                r = txt(A.run('TransferirReferenciaCore', teto=600))
                pg = txt(A.run('UltimaPergunta', teto=60))
                A.resp('')
                return r, pg

            def estado():
                return grade(), foto_rb(), A.foto_ls()

            limpar()
            e0 = estado()
            r_vaz, _ = transferir('SIM')
            ok('T3.2', r_vaz.startswith('ERRO|') and 'vazia' in r_vaz and estado() == e0,
               'digitacao vazia: recusa, nada muda', {'retorno': r_vaz[:120]})
            # tres analitos do lote (linhas da grade); o 1o com FORMULA na origem (Media e Fonte)
            lin_t = [r for r in range(R_REF1, R_REF1 + N_REF) if U(p.Cells(r, 1).Value) in set1 + set2][:3]
            r1, r2, r3 = lin_t
            g0 = grade()
            xl.EnableEvents = True            # como o usuario digita (eventos ligados)
            pôr(r1, 2, '=50+0.5')             # N1 Media por formula -> deve chegar 50,5
            pôr(r1, 3, '45')
            pôr(r1, 4, '=SUM(50,6)')          # outra formula -> deve chegar 56
            fmt_fonte = txt(p.Cells(r1, ncol + D).NumberFormat)
            pôr(r1, ncol, 'Bula E49')          # Fonte e coluna de TEXTO (@): formula digitada vira texto "=..." (T3.6)
            c2_ = 2 + 3                       # N2 Media
            pôr(r2, c2_, '200')
            pôr(r2, c2_ + 1, '190')
            pôr(r2, c2_ + 2, '210')
            pôr(r2, ncol, TAG)
            e1 = estado()
            d1 = digf()
            r0 = A.aud_n()
            r_nao, pg_nao = transferir('NAO')
            ok('T3.3', r_nao == 'CANCELADO|confirmacao' and estado() == e1 and digf() == d1
               and not A.aud(r0, 'REFERENCIA_BULA_ALTERADA'),
               'cancelar a confirmacao: grade, RefBulaStore, LotesStore e Audit_Log iguais',
               {'retorno': r_nao, 'pergunta': pg_nao[:250]})
            havia = any(txt(g0[r - R_REF1][c - 1]) not in ('',) and txt(g0[r - R_REF1][c - 1]) != v for r, c, v in
                        ((r1, 2, '50.5'), (r1, 3, '45'), (r1, 4, '56'), (r2, c2_, '200'), (r2, c2_ + 1, '190'),
                         (r2, c2_ + 2, '210')))
            r_sim, pg_sim = transferir('SIM')
            g1 = grade()
            esper = {(r1, 2): 50.5, (r1, 3): 45.0, (r1, 4): 56.0, (r2, c2_): 200.0, (r2, c2_ + 1): 190.0,
                     (r2, c2_ + 2): 210.0}
            val_ok = all(perto(p.Cells(r, c).Value, v) for (r, c), v in esper.items())
            fonte_ok = (txt(p.Cells(r1, ncol).Value) == 'Bula E49' and txt(p.Cells(r2, ncol).Value) == TAG)
            alvos = list(esper) + [(r1, ncol), (r2, ncol)]
            com_formula = [(r, c, txt(p.Cells(r, c).Formula)) for r, c in alvos if p.Cells(r, c).HasFormula
                           or txt(p.Cells(r, c).Formula).startswith('=')]
            mudou_fora = [(R_REF1 + i, j + 1) for i, (a_, b_) in enumerate(zip(g0, g1)) for j, (x, y) in
                          enumerate(zip(a_, b_)) if x != y and (R_REF1 + i, j + 1) not in alvos]
            st_rb = {U(x[2]): x for x in foto_rb()[1:] if txt(x[2])}
            rb_ok = all(perto(st_rb.get(f'{LT}|{U(p.Cells(r, 1).Value)}', [None] * 15)[3 + (c - 2)], v)
                        for (r, c), v in esper.items())
            links = q.wb.LinkSources(1)
            trilha = A.aud(r0, 'REFERENCIA_BULA_ALTERADA')
            ok('T3.4', r_sim == 'OK|8|2' and val_ok and fonte_ok and not com_formula and not mudou_fora and rb_ok
               and digf() == d1 and A.foto_ls() == e1[2] and not links and len(trilha) >= 6
               and (('SUBSTITU' in pg_sim) == havia),
               'transferir: SO VALORES (formula da origem chega como valor, sem formula/vinculo no destino), '
               'so as celulas preenchidas, RefBulaStore do lote, origem intacta, LotesStore igual, trilha',
               {'retorno': r_sim, 'valores_ok': val_ok, 'fonte_ok': fonte_ok, 'destino_com_formula': com_formula,
                'outras_celulas_da_grade_mudaram': mudou_fora[:8], 'refbulastore_ok': rb_ok,
                'origem_intacta': digf() == d1, 'links_externos': links, 'trilha': len(trilha),
                'grade_tinha_valor_diferente': havia, 'formato_fonte_digitacao': fmt_fonte, 'aviso_substituicao': 'SUBSTITU' in pg_sim})
            # substituicao: muda 1 valor; os outros iguais nao contam como substituidos
            pôr(r1, 4, '57')
            e2 = estado()
            r_sub, pg_sub = transferir('SIM')
            m_ = re.search(r'(\d+) valor\(es\) já transferido', pg_sub)
            ok('T3.5', r_sub.startswith('OK|') and 'SUBSTITU' in pg_sub and m_ and m_.group(1) == '1'
               and perto(p.Cells(r1, 4).Value, 57.0),
               'valor ja transferido e alterado: a confirmacao avisa 1 SUBSTITUIDO; confirmado, vai para a grade',
               {'retorno': r_sub, 'pergunta_fim': pg_sub[-180:]})
            # invalidos: cada um recusa TUDO (inclusive a linha valida nova r3)
            pôr(r3, 2, '10')
            pôr(r3, 3, '9')
            pôr(r3, 4, '11')
            casos = []
            for rot, cel, f_, chave in (('texto em coluna numerica', (r2, c2_), 'abc', 'número'),
                                        ('LI > LS na digitacao', (r1, 3), '999', 'LI maior'),
                                        ('LI digitado > LS da grade', (r1, 3), '60', 'LI maior'),
                                        ('Fonte comecando com "="', (r2, ncol), "'=SOMA(1)", 'Fonte')):
                ant = p.Cells(cel[0], cel[1] + D).Formula
                if rot == 'LI digitado > LS da grade':
                    pôr(r1, 4, '')            # LS fica o da grade (57)
                pôr(cel[0], cel[1], f_)
                e_ = estado()
                r0 = A.aud_n()
                rr_, _ = transferir('SIM')
                casos.append({'caso': rot, 'retorno': rr_[:150], 'nada_mudou': estado() == e_,
                              'sem_trilha': not A.aud(r0, 'REFERENCIA_BULA_ALTERADA'),
                              'ok': rr_.startswith('ERRO|') and chave in rr_ and estado() == e_
                              and not A.aud(r0, 'REFERENCIA_BULA_ALTERADA')})
                pôr(cel[0], cel[1], ant)
                if rot == 'LI digitado > LS da grade':
                    pôr(r1, 4, '57')
            r3_vazia = all(txt(p.Cells(r3, c).Value) == txt(e2[0][r3 - R_REF1][c - 1]) for c in (2, 3, 4)) or True
            ok('T3.6', all(c['ok'] for c in casos),
               'valores invalidos recusam a transferencia INTEIRA (a linha valida nova tambem nao vai); nada muda',
               {'casos': casos, 'linha_valida_nova': r3})
            ok('T3.7', A.foto_ls() == e1[2], 'a transferencia nunca toca a LotesStore (A1:AO identica do inicio ao fim)', {})
            limpar()
        except Exception as e:                # noqa: BLE001
            reg('T3.x', 'REPROVADO', 'digitacao/transferencia: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-900:]})

        # ============================================================ T1c/T2.6 Painel J7:J9 e Q4 por analito
        try:
            el = checar_m('T1.3', 'depois dos alvos aplicados')
            t0 = time.time()
            difJ, difQ, vist, nsig = [], [], 0, 0
            for an in cat:
                if time.time() - t0 > 1500:
                    break
                A.escolher_analito(an)
                vist += 1
                st = [txt(p.Range(f'AB{6 + i}').Value) for i in range(1, nlv + 1)]
                q4v = txt(p.Range('Q4').Value)
                q4o, _ = q4_oraculo(plano['q4f'], st)
                if q4o is not None and q4v != q4o:
                    difQ.append({'analito': an, 'status': st, 'Q4': q4v, 'oraculo': q4o})
                elq = est_linhas()
                for nv in range(1, nlv + 1):
                    x = elq.get((U(an), nv))
                    jv = txt(p.Range(f'J{6 + nv}').Value)
                    iv = p.Range(f'I{6 + nv}').Value
                    fv = txt(p.Range(f'F{6 + nv}').Value)
                    if x is None or not num(x[1]):
                        if jv not in ('',) and fv != '-':
                            difJ.append({'analito': an, 'nv': nv, 'J': jv, 'L': x[1] if x else None})
                        continue
                    nsig += 1
                    esperado = '' if fv == '-' else sigma_oraculo(float(x[1]))[0]
                    if jv != esperado or (fv != '-' and not perto(iv, x[1])) or (fv != '-' and jv != x[2]):
                        difJ.append({'analito': an, 'nv': nv, 'J': jv, 'I': iv, 'L': x[1], 'M': x[2],
                                     'oraculo': esperado, 'F': fv})
            ok('T1.4', nsig > 0 and not difJ,
               'Painel J7.. (Classe Sigma) = oraculo(L) = Estatistica M e I7.. = L, em todos os analitos/niveis com Sigma',
               {'analitos_visitados': vist, 'de': len(cat), 'niveis_com_sigma': nsig, 'diferencas': difJ[:10],
                'seg': round(time.time() - t0)})
            ok('T2.6', vist > 0 and not difQ, 'Q4 coerente com os status por nivel (AB) em todos os analitos visitados',
               {'analitos': vist, 'diferencas': difQ[:8]})
        except Exception as e:                # noqa: BLE001
            reg('T1.4', 'REPROVADO', 'Painel por analito: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-900:]})
    except Exception as e:                    # noqa: BLE001
        reg('Z99', 'REPROVADO', 'execucao (copia A) sem excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-1500:]})
    finally:
        if A is not None:
            A.fechar()

    # ================================================================ T4.3 copia B: botao de UM analito, isolado
    if plano.get('set1') and 'res_A' in dir() and res_A:
        B = None
        try:
            B = Sessao(produto, origem, 'B')
            difs = {}
            for rot, nomes, (ini, fim), modo in (('set1', plano['set1'], (None, None), '1'),
                                                 ('set2', plano['set2'], (dt.date.fromisoformat(plano['ini2']),
                                                                          dt.date.fromisoformat(plano['fim2'])), '2')):
                if rot not in res_A:
                    continue
                linhas, trilha = {}, []
                for an in nomes:
                    B.escolher_analito(an)
                    B.secao(plano['lote'], ini, fim)
                    r0 = B.aud_n()
                    B.resp(f'{modo}|SIM|SIM')
                    r = txt(B.run('AplicarAlvos'))
                    B.resp('')
                    sel = U(B.q.nome('selAnalito').Value)
                    u_ = B.ult_ls()
                    ln = [i for i, v in enumerate(B.ls.Range(f'P2:P{u_}').Value, 2)
                          if U(v[0]) == f'{U(plano["lote"])}|{U(an)}'][0]
                    linhas[U(an)] = norm_linha(B.ls.Range(f'A{ln}:AO{ln}').Value[0])
                    trilha += [(U(x[14]),) + tuple(txt(y) for y in x[6:14]) for x in B.aud(r0, 'PARAMETRO_LOTE_CALCULADO')]
                    if not r.startswith('OK|') or sel != U(an):
                        difs.setdefault(rot, []).append({'analito': an, 'retorno': r[:150], 'selAnalito': sel})
                for an, la in res_A[rot]['linhas'].items():
                    lb = linhas.get(an)
                    if la != lb:
                        difs.setdefault(rot, []).append({'analito': an, 'colunas_diferentes': [
                            (j, la[j], lb[j] if lb else None) for j in range(len(la)) if not lb or la[j] != lb[j]][:8]})
                if sorted(trilha) != res_A[rot]['trilha']:
                    difs.setdefault(rot, []).append({'trilha_A': res_A[rot]['trilha'][:6], 'trilha_B': sorted(trilha)[:6]})
            ok('T4.3', not difs,
               'cada analito do AplicarAlvosVarios (copia A) = o que o botao de UM analito grava isolado (copia B): '
               'LotesStore A:AO (sem o carimbo de hora) e trilha PARAMETRO_LOTE_CALCULADO identicas',
               {'diferencas': difs, 'comparados': {k: list(v['linhas']) for k, v in res_A.items()}})
        except Exception as e:                # noqa: BLE001
            reg('T4.3', 'REPROVADO', 'copia B: excecao', {'erro': repr(e), 'trace': traceback.format_exc()[-1200:]})
        finally:
            if B is not None:
                B.fechar()
    else:
        reg('T4.3', 'NAO_EXECUTADO', 'copia B nao rodada: a copia A nao produziu o resultado do AplicarAlvosVarios', {})

    h1 = sha(origem)
    hp1 = {k: sha(v) for k, v in prods.items()}
    ok('Y00', h0 == h1 and hp0 == hp1, 'origem e producoes da raiz intactas (SHA-256)',
       {'origem': [h0[:16], h1[:16]], 'producoes': {k: [hp0[k][:16], hp1[k][:16]] for k in prods}})
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump({'produto': produto, 'origem_sha': h0, 'plano': plano, 'resultados': RES}, f, ensure_ascii=False,
                  indent=1, default=str)
    cont = collections.Counter(r['estado'] for r in RES)
    print(f'\n=== repro_e49_adr103 {produto}: ' + ', '.join(f'{k} {v}' for k, v in sorted(cont.items())) + ' ===',
          flush=True)
    return cont


if __name__ == '__main__':
    a = sys.argv[1:]
    if len(a) < 3 or a[0] not in ('Bioquimica', 'Hematologia'):
        print(__doc__)
        sys.exit(2)
    c = executar(a[0], os.path.abspath(a[1]), os.path.abspath(a[2]))
    sys.exit(1 if c.get('REPROVADO') else 0)
