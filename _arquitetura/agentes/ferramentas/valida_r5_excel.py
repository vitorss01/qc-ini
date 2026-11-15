# -*- coding: utf-8 -*-
"""valida_r5_excel.py -- Tester/QA, ciclo 2026-11-15_reteste05 (+ parte 2), modo VALIDAR do ADR-095 (entrega 42).

Itens: OC-0012 (aviso do corte na Hema; regressao na Bio), OC-0068 (listas de escolha), OC-0071 (cabecalho da
Estatistica da Bio), OC-0072 (ano literal), mais as conferencias A (R5-01 e a janela da Estatistica da Bio) e
B (P-03: linha manual copiada e colada com o mesmo ID).

Escrito pelo Tester/QA. NAO reaproveita os scripts do Usuario Simulado (reteste05*_usuario_*.py, xml_formatos_listas)
nem as suites qa_adr09x. Reaproveita so a infraestrutura: qa_final.QA (abre a copia, revela as abas, imprime
EXCEL_PID), repro_comum (meu), Vigia (meu, valida_adr090_091: le e fecha caixas do MEU PID, sem imagem) e
qa_inativar_ux.criar_usuario/fazer_login (usuario de teste so na copia).

Oraculo proprio (Python puro sobre o Value2 da tblCQ_Final da copia):
  * resultado que conta = PARTICIPA_ESTATISTICA = SIM, RESULTADO numerico, analito, equipamento do Painel (Eng_Saida!S1)
    e lote (H3; vazio = todos os lotes);
  * corrida = RUN (com lote) ou o par (RUN, LOTE) (H3 vazio, ADR-080/086); data da corrida = data da 1a linha;
  * corridas ordenadas por (data, RUN); janela = as 180 corridas elegiveis mais recentes ate o Ate;
  * cortadas = corridas do periodo antes da janela; AF1 = n do periodo (por nivel) ate o Ate.
Para a Estatistica da Bio (conferencia A): n/media por nivel para varias janelas (lote inteiro, ano de B5..D5,
De..Ate, De..ultimo dado), para dizer qual janela a grade usa.

Uso: python -u valida_r5_excel.py <Bioquimica|Hematologia> <base|nova> [--manual]
  base = copia do backup da entrega 41; nova = copia da producao atual (entrega 42).
  --manual (so nova): manuais de teste, colagem de linha (P-03) e ATUALIZAR DADOS.
Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-15_reteste05/tester/; nada e salvo; fechar() encerra so o PID
que o QA criou; nenhuma captura de tela. Dados de teste: usuario TESTE_AUDITORIA_R5V, valores sinteticos
123,456 / 98,765 / 111,111 (Bio) e 12,345 / 11,111 (Hema), data 14/11/2026, OBS/COMENTARIO TESTE-AUDITORIA.
"""
import collections
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import statistics as stt
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402
from valida_adr090_091 import Vigia  # noqa: E402  (minha, ciclo reteste03)
sys.path.insert(0, C.TESTES)
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = sys.argv[1]
ORIG = sys.argv[2]
MANUAL = '--manual' in sys.argv
SO = sys.argv[sys.argv.index('--so') + 1].split(',') if '--so' in sys.argv else None
LOTE_BIO = sys.argv[sys.argv.index('--lote') + 1] if '--lote' in sys.argv else '8974'


def quer(n):
    return SO is None or n in SO
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
CICLO = '2026-11-15_reteste05'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO, 'tester')
BACKUP = {'Bioquimica': '_backup_pre_integracao_2026-11-15_0011', 'Hematologia': '_backup_pre_integracao_2026-11-15_0458'}
USR = 'TESTE_AUDITORIA_R5V'
SEGREDOS = [x for x in (os.environ.get('USERNAME'), os.environ.get('COMPUTERNAME')) if x and len(x) > 3]
try:
    import win32api
    SEGREDOS.append(win32api.GetUserNameEx(3))
except Exception:  # noqa: BLE001
    pass
SEGREDOS = [x for x in SEGREDOS if x and len(x) > 3]
EPOCA = dt.datetime(1899, 12, 30)


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    return t(v).upper()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def ser(d):
    return float((d - EPOCA).days)


def dser(x):
    return (EPOCA + dt.timedelta(days=float(x))).strftime('%d/%m/%Y %H:%M:%S') if num(x) else t(x)


def rel(a, b):
    try:
        return abs(a - b) / max(abs(a), abs(b), 1e-300)
    except Exception:  # noqa: BLE001
        return None


class Log(C.Log):
    def gravar(self):
        s = json.dumps(self.R, ensure_ascii=False, indent=1, default=str)
        for x in SEGREDOS:
            s = re.sub(re.escape(x), '<omitido>', s, flags=re.I)
        open(self.saida, 'w', encoding='utf-8').write(s)


def secao(L, nome, fn, *a):
    t0 = time.time()
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    try:
        r = fn(*a)
    except Exception as e:  # noqa: BLE001
        r = {'ERRO': repr(e), 'trace': traceback.format_exc()[-2500:]}
    if isinstance(r, dict):
        r['_s'] = round(time.time() - t0, 1)
    L[nome] = r
    return r


# ============================================================================================ tblCQ_Final (Value2)
class Final:
    COLS = ('ID_REGISTRO', 'ANALITO', 'NIVEL', 'LOTE', 'EQUIPAMENTO', 'RUN', 'DATA_HORA', 'RESULTADO',
            'PARTICIPA_ESTATISTICA', 'TIPO_PLOTAGEM_LJ', 'STATUS_ANALITICO', 'ORIGEM_RESULTADO',
            'MOTIVO_EXCLUSAO_AUTOMATICA')

    def __init__(self, qa):
        lo = qa.lo('tblCQ_Final')
        self.n = lo.ListRows.Count
        tudo = lo.DataBodyRange.Value2
        self.hash = hashlib.sha256(repr(tudo).encode()).hexdigest()[:16]
        cab = [str(c.Name) for c in lo.ListColumns]
        ix = {c: cab.index(c) for c in self.COLS if c in cab}
        self.rows = [{c: l[i] for c, i in ix.items()} for l in tudo]


def oraculo_painel(fin, an, lote, eq, de, ate, nlv, nk=180):
    """Reproducao independente da janela de 180 corridas, do corte e do n do periodo (ver cabecalho)."""
    todos = lote == ''
    eleg, xs = [], []
    for r in fin.rows:
        if up(r['ANALITO']) != up(an):
            continue
        if eq and up(r['EQUIPAMENTO']) != up(eq):
            continue
        if not todos and t(r['LOTE']) != lote:
            continue
        if up(r['PARTICIPA_ESTATISTICA']) == 'SIM':
            eleg.append(r)
        elif t(r['TIPO_PLOTAGEM_LJ']) == 'X_VERMELHO':
            xs.append(r)

    def key(r):
        return (t(r['RUN']), t(r['LOTE'])) if todos else t(r['RUN'])
    data = collections.OrderedDict()
    for r in eleg + xs:
        k = key(r)
        if k not in data:
            data[k] = float(r['DATA_HORA']) if num(r['DATA_HORA']) else 0.0
    ks = set(key(r) for r in eleg)
    def runn(k):
        try:
            return int(float(k[0] if todos else k))
        except Exception:  # noqa: BLE001
            return 0
    ordem = sorted(data, key=lambda k: (data[k], runn(k)))

    def passa(d):
        if d <= 0:
            return True
        if de and int(d) < de:
            return False
        if ate and int(d) > ate:
            return False
        return True
    fim = len(ordem)
    if ate:
        while fim >= 1 and int(data[ordem[fim - 1]]) > ate:
            fim -= 1
    ate_fim = ordem[:fim]
    elegs = [k for k in ate_fim if k in ks]
    jan = set(elegs[-nk:])
    ini = ate_fim.index(elegs[-nk]) if len(elegs) >= nk else 0
    cort = sum(1 for k in ate_fim[:ini] if passa(data[k]))
    so_x_jan = sum(1 for k in ate_fim[ini:] if k not in ks)
    n_eleg_jan = min(len(elegs), nk)
    pular_x = max(0, so_x_jan - (nk - n_eleg_jan))      # corridas so-X sem posicao na janela (D10 do QA-ETL-001)
    per = collections.Counter()
    janv = collections.defaultdict(list)
    for r in eleg:
        if not num(r['RESULTADO']):
            continue
        k = key(r)
        nv = int(r['NIVEL'] or 0)
        if k in data and k in ate_fim and passa(data[k]):
            per[nv] += 1
        if k in jan:
            janv[nv].append(float(r['RESULTADO']))
    corr_per = sum(1 for k in ate_fim if k in ks and passa(data[k]))
    af1 = ' · '.join(f'N{i} {per.get(i, 0)}' for i in range(1, nlv + 1))
    niveis = []
    for i in range(1, nlv + 1):
        v = janv.get(i, [])
        niveis.append({'nivel': i, 'n': len(v), 'media': stt.fmean(v) if v else None,
                       'dp': stt.stdev(v) if len(v) > 1 else None})
    return {'corridas_periodo': corr_per, 'corridas_ate_fim': len(elegs), 'cortadas': cort,
            'so_X_na_janela': so_x_jan, 'pular_X': pular_x, 'M1_esperado': cort + pular_x, 'AF1': af1 if (cort + pular_x) else '', 'AF1_sempre': af1, 'janela': niveis}


def stats_janela(fin, an, lote, eq, nivel, de=None, ate=None, anos=None):
    """n/media/DP por janela de datas (por resultado, dia inteiro) ou por ano (anos = (anoDe, anoAte))."""
    v = []
    for r in fin.rows:
        if up(r['ANALITO']) != up(an) or up(r['PARTICIPA_ESTATISTICA']) != 'SIM' or not num(r['RESULTADO']):
            continue
        if eq and up(r['EQUIPAMENTO']) != up(eq):
            continue
        if lote and t(r['LOTE']) != lote:
            continue
        if int(r['NIVEL'] or 0) != nivel:
            continue
        d = r['DATA_HORA']
        if not num(d):
            continue
        if de and int(d) < de:
            continue
        if ate and int(d) > ate:
            continue
        if anos and not (anos[0] <= (EPOCA + dt.timedelta(days=float(d))).year <= anos[1]):
            continue
        v.append(float(r['RESULTADO']))
    return {'n': len(v), 'media': stt.fmean(v) if v else None, 'dp': stt.stdev(v) if len(v) > 1 else None}


# ============================================================================================ telas
def aplicar(qa, p, passos):
    for c, v in passos:
        if v is None:
            p.Range(c).MergeArea.ClearContents()
        elif isinstance(v, dt.datetime):
            p.Range(c).Value = ser(v)          # data como serial (OC-0066: o pywin32 desloca datetime pelo fuso)
        else:
            p.Range(c).Value = v
        qa.ex.esperar()
        time.sleep(0.4)
    time.sleep(1.0)
    qa.ex.esperar()


def celinfo(c):
    return {'f': t(c.Formula)[:200], 'txt': t(c.Text), 'v2': c.Value2, 'fmt': t(c.NumberFormat)}


def eng_hash(eng):
    """Hash do Eng_Saida A1:AF190 (o que o LJ plota) sem o G1 (carimbo Now)."""
    v = [list(x) for x in eng.Range('A1:AF190').Value2]
    v[0][6] = None
    return hashlib.sha256(repr(v).encode()).hexdigest()[:16]


def ler_cenario(qa, fin):
    p, eng, e = qa.ws('Painel'), qa.ws('Eng_Saida'), qa.ws('Estatística')
    filtro = {c: p.Range(c).Value2 for c in ('C3', 'H3', 'K3', 'K4', 'G3', 'G4') + (('L4',) if BIO else ())}
    de, ate = qa.nome('filtroDe').Value2, qa.nome('filtroAte').Value2
    de = int(de) if num(de) and de > 0 else None
    ate = int(ate) if num(ate) and ate > 0 else None
    an = t(qa.nome('selAnalito').Value)
    lote = t(eng.Range('E1').Value).lstrip("'")
    todos = bool(eng.Range('AD1').Value)
    eq = t(eng.Range('S1').Value)
    o = oraculo_painel(fin, an, '' if todos else lote, eq, de, ate, NLV)
    if o['cortadas'] + o['pular_X'] == 0:
        o['AF1'] = ''
    eng_p = [list(x) for x in qa.nome('engPainel').Value2]
    out = {'filtro': filtro, 'filtroDe': dser(de) if de else None, 'filtroAte': dser(ate) if ate else None,
           'analito': an, 'lote': lote, 'todos': todos, 'equip': eq,
           'M1': eng.Range('M1').Value2, 'AF1': t(eng.Range('AF1').Value), 'A39': t(p.Range('A39').Value),
           'engPainel': eng_p, 'oraculo': o,
           'painel_A6Q12': [list(x) for x in p.Range('A6:Q12').Value2],
           'calc': hashlib.sha256(repr(qa.ws('Calc').Range('A1:AZ190').Value2).encode()).hexdigest()[:16],
           'eng_saida': eng_hash(eng)}
    m1 = int(out['M1'] or 0)
    out['confere'] = {
        'M1_igual_cortadas': m1 == o['M1_esperado'],
        'AF1_igual_oraculo': out['AF1'] == o['AF1'],
        'n_janela_igual': [(eng_p[i][1] if len(eng_p[i]) > 1 else None, o['janela'][i]['n']) for i in range(NLV)],
        'media_rel': [rel(eng_p[i][2], o['janela'][i]['media']) if num(eng_p[i][2]) and o['janela'][i]['media'] is not None
                      else None for i in range(NLV)]}
    if BIO:
        hdr = {a: celinfo(e.Range(a)) for a in ('A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'A5', 'B5', 'C5', 'D5',
                                                'E5', 'F5', 'G5', 'H5', 'I5', 'J5')}
        out['estat_hdr'] = hdr
        g = e.Range('A14:AU133').Value2
        out['estat_hash'] = hashlib.sha256(repr(g).encode()).hexdigest()[:16]
        out['estat_grade'] = [list(x) for x in g]
        txt12 = ' '.join(t(e.Cells(r, k).Text) for r in range(1, 13) for k in range(1, 22))
        out['estat_1900'] = '1900' in txt12
        out['estat_yyyy'] = 'yyyy' in txt12.lower()
        # conferencia A: janelas candidatas para o n da grade (Glicose ou o analito do Painel)
        lin = {}
        for i, l in enumerate(g, 14):
            if l[0] and num(l[1]):
                lin.setdefault(f'{t(l[0])}|{int(l[1])}', (i, l[2], l[3]))
        h4 = t(e.Range('H4').Value).lstrip("'")
        b5, d5 = e.Range('B5').Value2, e.Range('D5').Value2
        anoDe = (EPOCA + dt.timedelta(days=b5)).year if num(b5) and b5 > 0 else 0
        anoAte = (EPOCA + dt.timedelta(days=d5)).year if num(d5) and d5 > 0 else anoDe
        jan = {}
        for nv in range(1, NLV + 1):
            k = f'{an}|{nv}'
            gr = lin.get(k)
            jan[k] = {'grade_linha_n_media': gr,
                      'lote_inteiro': stats_janela(fin, an, h4, eq, nv),
                      'ano_B5_D5': stats_janela(fin, an, h4, eq, nv, anos=(anoDe, anoAte)) if anoDe else None,
                      'De_Ate_painel': stats_janela(fin, an, h4, eq, nv, de=de, ate=ate) if (de or ate) else None}
        out['estat_janelas'] = {'H4_lote': h4, 'anoDe': anoDe, 'anoAte': anoAte, 'por_nivel': jan}
    return out


def cenarios(fin=None):
    D = dt.datetime
    if BIO:
        g = collections.Counter(t(r['ANALITO']) for r in (fin.rows if fin else [])
                                if up(r['EQUIPAMENTO']) == 'DIMENSION 2' and up(r['PARTICIPA_ESTATISTICA']) == 'SIM'
                                and up(r['ANALITO']) != 'GLICOSE')
        an2 = g.most_common(1)[0][0] if g else 'Ureia'
        # mes de 2026 com mais resultados de Glicose / DIMENSION 1 / lote 8974 (para as janelas so De / De-Ate / so Ate)
        meses = collections.Counter()
        for r in (fin.rows if fin else []):
            if up(r['ANALITO']) == 'GLICOSE' and up(r['EQUIPAMENTO']) == 'DIMENSION 1' and t(r['LOTE']) == LOTE_BIO                     and up(r['PARTICIPA_ESTATISTICA']) == 'SIM' and num(r['DATA_HORA']):
                d = EPOCA + dt.timedelta(days=float(r['DATA_HORA']))
                if d.year == 2026:
                    meses[d.month] += 1
        mes = max(meses, key=lambda m: (meses[m], m)) if meses else 6
        nm = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'][mes - 1]
        d1 = D(2026, mes, 1)
        d2 = (D(2026, mes + 1, 1) if mes < 12 else D(2027, 1, 1)) - dt.timedelta(days=1)
        return [
            ('B1_Tudo', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'" + LOTE_BIO), ('K4', 'Tudo (sem filtro)')]),
            ('B2_Set2026', [('K3', 2026), ('K4', 'Set')]),
            ('B3_AnoInteiro', [('K4', 'Ano inteiro')]),
            ('B4_H3vazio_Tudo', [('H3', None), ('K4', 'Tudo (sem filtro)')]),
            ('B5_H3vazio_Ano2026', [('K3', 2026), ('K4', 'Ano inteiro')]),
            ('B6_so_De_Set', [('H3', "'" + LOTE_BIO), ('K3', 2026), ('K4', 'Set'), ('G4', None)]),
            ('B6b_mes_com_dados', [('K4', nm)]),
            ('B6c_so_De_mes', [('G4', None)]),
            ('B7_DeAte_digitado_mes', [('G3', d1), ('G4', d2)]),
            ('B8_so_Ate_mes', [('G3', None)]),
            ('B9_DIM2_outro_H3vazio_Tudo', [('L4', 'DIMENSION 2'), ('C3', an2), ('H3', None),
                                            ('K4', 'Tudo (sem filtro)')]),
            ('B10_volta_Tudo', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'" + LOTE_BIO), ('K4', 'Tudo (sem filtro)')]),
        ]
    return [
        ('H1_WBC_todos_2026', [('C3', 'WBC'), ('H3', None), ('K3', 2026), ('K4', 'Ano inteiro')]),
        ('H2_WBC_todos_Tudo', [('K4', 'Tudo (sem filtro)')]),
        ('H3_WBC_todos_DeAte', [('G3', D(2026, 1, 1)), ('G4', D(2026, 10, 31))]),
        ('H4_PLT_todos_Tudo', [('C3', 'PLT'), ('K4', 'Tudo (sem filtro)')]),
        ('H5_RBC_todos_2025', [('C3', 'RBC'), ('K3', 2025), ('K4', 'Ano inteiro')]),
        ('H6_WBC_lote528211_Tudo', [('C3', 'WBC'), ('H3', "'528211"), ('K4', 'Tudo (sem filtro)')]),
        ('H7_WBC_todos_Out2026', [('H3', None), ('K3', 2026), ('K4', 'Out')]),
    ]


def s_painel(qa, fin):
    p = qa.ws('Painel')
    p.Activate()
    qa.ex.esperar()
    out = {}
    for nome, passos in cenarios(fin):
        print('   ', nome, time.strftime('%H:%M:%S'), flush=True)
        try:
            aplicar(qa, p, passos)
            out[nome] = ler_cenario(qa, fin)
        except Exception as e:  # noqa: BLE001
            out[nome] = {'ERRO': repr(e), 'trace': traceback.format_exc()[-1500:]}
    # a formula do A39 (os dois ramos do aviso) -- o ramo "do lote" nao tem dado real com > 180 corridas
    out['A39_formula_trecho'] = re.findall(r'Período inteiro.{0,160}', t(p.Range('A39').Formula))
    if BIO:
        e = qa.ws('Estatística')
        lin = None
        for r in range(14, 134):
            if up(e.Cells(r, 1).Value) == 'GLICOSE':
                lin = r
                break
        if lin:
            out['estat_formulas_Glicose'] = {e.Cells(lin, c).Address.replace('$', ''): t(e.Cells(lin, c).Formula)[:300]
                                              for c in range(1, 14)}
        nomes = {}
        for nm in ('engEstat', 'engPainel', 'Data_Inicio_Efetiva', 'Data_Fim_Efetiva'):
            try:
                nomes[nm] = t(qa.wb.Names(nm).RefersTo)
            except Exception as ex:  # noqa: BLE001
                nomes[nm] = f'(nao existe: {ex.__class__.__name__})'
        out['nomes'] = nomes
    return out


# ============================================================================================ OC-0072
def s_ano_literal(qa):
    xl = qa.ex.xl
    vis = [w.Name for w in qa.wb.Worksheets if w.Visible == -1]
    out = {'abas_visiveis': vis, 'achados_valores': [], 'achados_formulas': [], 'alvos': {}}
    for nm in vis:
        ws = qa.ws(nm)
        for lookin, chave in ((-4163, 'achados_valores'), (-4123, 'achados_formulas')):
            try:
                c = ws.Cells.Find(What='yyyy', LookIn=lookin, LookAt=2, MatchCase=False)
            except Exception as e:  # noqa: BLE001
                out[chave].append({'aba': nm, 'erro': repr(e)})
                continue
            vistos = 0
            if c is not None:
                primeiro = c.Address
                while True:
                    out[chave].append({'aba': nm, 'cel': c.Address.replace('$', ''), 'txt': t(c.Text)[:80],
                                       'fmt': t(c.NumberFormat), 'f': t(c.Formula)[:120]})
                    vistos += 1
                    c = ws.Cells.FindNext(After=c)
                    if c is None or c.Address == primeiro or vistos >= 15:
                        break
    alvos = [('Registros', 'E4'), ('Registros', 'F4'), ('Configuração', 'D26'), ('Configuração', 'AD26'),
             ('Eng_Saida', 'AC3'), ('Westgard_Fotos', 'F3'), ('LotesStore', 'U1')]
    alvos += [('Estatística', 'E8'), ('Estatística', 'F9'), ('Estatística', 'C10')] if BIO else \
        [('Eventos_Westgard', 'A4'), ('Eventos_Westgard', 'A5003')]
    for aba, a in alvos:
        try:
            c = qa.ws(aba).Range(a)
            out['alvos'][f'{aba}!{a}'] = {'txt': t(c.Text), 'v2': c.Value2, 'fmt': t(c.NumberFormat),
                                          'fmt_local': t(c.NumberFormatLocal), 'oculta': bool(c.EntireColumn.Hidden)}
        except Exception as e:  # noqa: BLE001
            out['alvos'][f'{aba}!{a}'] = repr(e)
    # Sync_Log (Bio): ultima linha com data
    if BIO:
        try:
            sl = qa.ws('Sync_Log')
            ult = sl.Cells(sl.Rows.Count, 1).End(-4162).Row
            out['alvos'][f'Sync_Log!A{ult}'] = {'txt': t(sl.Cells(ult, 1).Text), 'fmt': t(sl.Cells(ult, 1).NumberFormat)}
        except Exception as e:  # noqa: BLE001
            out['alvos']['Sync_Log'] = repr(e)
    # Registros E4/F4 digitados (o usuario filtra por data): Text depois de digitar, e o filtro so muda a vista
    rg = qa.ws('Registros')
    lo = rg.ListObjects(1)
    h0 = hashlib.sha256(repr(lo.DataBodyRange.Value2).encode()).hexdigest()[:16]
    try:
        rg.Range('E4').Value = ser(dt.datetime(2026, 9, 1))
        qa.ex.esperar()
        rg.Range('F4').Value = ser(dt.datetime(2026, 9, 30))
        qa.ex.esperar()
        time.sleep(1)
        out['Registros_E4F4_digitados'] = {'E4': t(rg.Range('E4').Text), 'F4': t(rg.Range('F4').Text),
                                           'linhas_visiveis': visiveis(lo),
                                           'hash_tabela_igual': hashlib.sha256(repr(lo.DataBodyRange.Value2).encode())
                                           .hexdigest()[:16] == h0}
    finally:
        rg.Range('E4').ClearContents()
        rg.Range('F4').ClearContents()
        qa.ex.esperar()
    out['n_achados_valores'] = len([x for x in out['achados_valores'] if 'cel' in x])
    out['n_achados_formulas'] = len([x for x in out['achados_formulas'] if 'cel' in x])
    return out


def visiveis(lo):
    try:
        return int(lo.DataBodyRange.Columns(1).SpecialCells(12).Count)    # xlCellTypeVisible
    except Exception:  # noqa: BLE001
        return 0


# ============================================================================================ OC-0068
def s_listas(qa, senha):
    xl = qa.ex.xl
    sep = t(xl.International[5 - 1])
    out = {'separador_lista': sep}
    lo = qa.lo('tblResultados_Manuais')
    _, r = C.linha_livre_manual(qa)
    alvos = []
    for col in (('EQUIPAMENTO', 'MATRIZ', 'MOTIVO') if BIO else ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO')):
        alvos.append((f'Digitar Resultados {col}', lo.ListColumns(col).DataBodyRange.Cells(r, 1)))
    alvos.append(('Registros H4', qa.ws('Registros').Range('H4')))
    if BIO:
        alvos.append(('Painel L4', qa.ws('Painel').Range('L4')))
    invalidos = ['DIMENSION 1,DIMENSION 2', 'DIMENSION 1;DIMENSION 2', 'DIMENSION 3', 'Sim,Não', 'Talvez',
                 'Falha do interfaceamento,Resultado não transmitido', 'Outro: TESTE-AUDITORIA', 'XN2000', 'PLASMA']
    for nome, cel in alvos:
        v = cel.Validation
        d = {}
        try:
            d = {'tipo': v.Type, 'formula1': t(v.Formula1), 'estilo': v.AlertStyle, 'mostra_erro': bool(v.ShowError),
                 'titulo': t(v.ErrorTitle), 'msg': t(v.ErrorMessage), 'em_celula': bool(v.InCellDropdown)}
        except Exception as e:  # noqa: BLE001
            out[nome] = {'erro': repr(e)}
            continue
        f1 = d['formula1']
        itens = [x for x in f1.split(sep)] if not f1.startswith('=') else None
        d['itens'] = itens
        d['n_itens'] = len(itens) if itens else None
        antes = cel.Formula
        xl.EnableEvents = False
        try:
            try:
                cel.Value = antes if antes else None      # celula destravada? (sem tirar a protecao da aba)
            except Exception:  # noqa: BLE001
                d['precisou_desproteger'] = True
                cel.Parent.Unprotect(senha)
            aceito, recus = {}, {}
            for it in (itens or []):
                cel.Value = it
                aceito[it] = bool(cel.Validation.Value)
            for it in invalidos:
                if itens and it in itens:
                    continue
                cel.Value = it
                recus[it] = bool(cel.Validation.Value)
            d['validos_aceitos'] = aceito
            d['invalidos_aceitos'] = recus
        finally:
            cel.Formula = antes
            xl.EnableEvents = True
        d['ok'] = itens is not None and all(aceito.values()) and not any(recus.values())
        out[nome] = d
    # Registros H4 "so muda a vista": hash da tabela e linhas visiveis com Sim / Não / vazio (eventos ligados)
    rg = qa.ws('Registros')
    lor = rg.ListObjects(1)
    h0 = hashlib.sha256(repr(lor.DataBodyRange.Value2).encode()).hexdigest()[:16]
    fin_h0 = hashlib.sha256(repr(qa.lo('tblCQ_Final').DataBodyRange.Value2).encode()).hexdigest()[:16]
    vis = {'inicio': visiveis(lor)}
    itens_h4 = out.get('Registros H4', {}).get('itens') or []
    for it in itens_h4:
        rg.Range('H4').Value = it
        qa.ex.esperar()
        time.sleep(1)
        vis[it] = visiveis(lor)
    rg.Range('H4').ClearContents()
    qa.ex.esperar()
    time.sleep(1)
    vis['limpo'] = visiveis(lor)
    out['Registros_H4_vista'] = {'linhas_visiveis': vis, 'nome_tabela': lor.Name,
                                 'hash_tabela_igual': hashlib.sha256(repr(lor.DataBodyRange.Value2).encode())
                                 .hexdigest()[:16] == h0,
                                 'hash_tblCQ_Final_igual': hashlib.sha256(repr(qa.lo('tblCQ_Final').DataBodyRange.Value2)
                                                                         .encode()).hexdigest()[:16] == fin_h0}
    return out


# ============================================================================================ manuais / P-03
def manual_cols(qa):
    lo = qa.lo('tblResultados_Manuais')
    return lo, [str(c.Name) for c in lo.ListColumns]


def ler_linha(qa, r):
    lo, cab = manual_cols(qa)
    rg = lo.ListRows(r).Range
    v = rg.Value2[0]
    out = {}
    for i, c in enumerate(cab):
        out[c] = t(v[i]) if not num(v[i]) else v[i]
        tx = t(rg.Cells(1, i + 1).Text)
        if tx != t(v[i]):
            out[c + '_txt'] = tx
    return out


def escrever(qa, r, col, v):
    lo, _ = manual_cols(qa)
    cel = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
    if isinstance(v, str):
        cel.FormulaLocal = v
    else:
        cel.Value = v
    qa.ex.esperar()
    time.sleep(0.2)


def ref_lote(fin, an, eq, nivel):
    """O lote do resultado mais recente (ATIVO, interfaceamento) do analito no equipamento e nivel."""
    c = [r for r in fin.rows if up(r['ANALITO']) == up(an) and up(r['PARTICIPA_ESTATISTICA']) == 'SIM'
         and int(r['NIVEL'] or 0) == nivel and (not eq or up(r['EQUIPAMENTO']) == up(eq)) and num(r['DATA_HORA'])
         and up(r['ORIGEM_RESULTADO']).startswith('INTERF')]
    c.sort(key=lambda r: r['DATA_HORA'])
    return t(c[-1]['LOTE']), dser(c[-1]['DATA_HORA']), len(c)


def serie(qa, fin, passos):
    p = qa.ws('Painel')
    aplicar(qa, p, passos)
    return ler_cenario(qa, fin)


def s_manual(qa, senha, L):
    out = {}
    fin0 = Final(qa)
    lo, cab = manual_cols(qa)
    out['colunas'] = cab
    an = 'Glicose' if BIO else 'WBC'
    eqA = 'DIMENSION 1' if BIO else ''
    loteA, ultA, nA = ref_lote(fin0, an, eqA, 1)
    out['referencia_A'] = {'analito': an, 'equip': eqA or '(painel)', 'lote': loteA, 'ultimo': ultA, 'n_ativos': nA}
    passosA = ([('L4', 'DIMENSION 1')] if BIO else []) + [('C3', an), ('H3', "'" + loteA), ('K3', 2026), ('K4', 'Nov')]
    antesA = serie(qa, fin0, passosA)
    out['serie_A_antes'] = {k: antesA[k] for k in ('engPainel', 'oraculo', 'M1', 'filtroDe', 'filtroAte')}
    if BIO:
        loteB, ultB, nB = ref_lote(fin0, an, 'DIMENSION 2', 1)
        out['referencia_B_DIM2'] = {'lote': loteB, 'ultimo': ultB, 'n_ativos': nB}
        passosB = [('L4', 'DIMENSION 2'), ('C3', an), ('H3', "'" + loteB), ('K3', 2026), ('K4', 'Nov')]
        antesB = serie(qa, fin0, passosB)
        out['serie_B_antes'] = {k: antesB[k] for k in ('engPainel', 'oraculo', 'M1')}
    # ---- entrada pela Digitar Resultados (botao) com eventos ligados
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as e:  # noqa: BLE001
        out['IrDigitarResultados'] = repr(e)
    dr = qa.ws('Digitar Resultados')
    dr.Activate()
    qa.ex.xl.EnableEvents = True
    out['EnableEvents'] = bool(qa.ex.xl.EnableEvents)
    sep = t(qa.ex.xl.International[5 - 1])
    _, r0 = C.linha_livre_manual(qa)
    f_eq = t(lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r0, 1).Validation.Formula1).split(sep)
    f_ma = t(lo.ListColumns('MATRIZ').DataBodyRange.Cells(r0, 1).Validation.Formula1).split(sep)
    f_mo = t(lo.ListColumns('MOTIVO').DataBodyRange.Cells(r0, 1).Validation.Formula1).split(sep)
    out['itens_menu'] = {'EQUIPAMENTO': f_eq, 'MATRIZ': f_ma, 'MOTIVO': f_mo}
    obs = next((c for c in cab if any(k in c.upper() for k in ('OBS', 'COMENT'))), None)
    linhas = {}

    def novo(rot, campos):
        _, r = C.linha_livre_manual(qa)
        for col, v in campos:
            escrever(qa, r, col, v)
        if obs:
            escrever(qa, r, obs, 'TESTE-AUDITORIA R5V ' + rot)
        linhas[rot] = r
        return r
    vig = Vigia(qa.ex.pid, teto=60)
    vig.start()
    # M1 (so Bio, OC-0068): EQUIPAMENTO e MATRIZ escolhidos do MENU (o 2o item = DIMENSION 2), MOTIVO do menu
    if BIO:
        eq_menu = f_eq[1] if len(f_eq) > 1 else f_eq[0]
        novo('M1_equip_menu', [('DATA', '14/11/2026'), ('HORA', '09:07:05'), ('EQUIPAMENTO', eq_menu),
                               ('MATRIZ', f_ma[0]), ('LOTE', "'" + loteB), ('NIVEL', '1'), ('ANALITO', an),
                               ('RESULTADO', '123,456'), ('MOTIVO', f_mo[-1])])
    # M2: linha de origem da colagem (P-03)
    val2 = '98,765' if BIO else '12,345'
    campos2 = [('DATA', '14/11/2026'), ('HORA', '10:20:30')] + ([('EQUIPAMENTO', 'DIMENSION 1')] if BIO else []) + \
        [('LOTE', "'" + loteA), ('NIVEL', '1'), ('ANALITO', an), ('RESULTADO', val2), ('MOTIVO', f_mo[-1])]
    r2 = novo('M2_origem', campos2)
    # P-03 a) colar sem editar: copia a linha inteira da tabela para a proxima livre (eventos ligados)
    _, r3 = C.linha_livre_manual(qa)
    lo.ListRows(r2).Range.Copy(lo.ListRows(r3).Range)
    qa.ex.esperar()
    time.sleep(0.5)
    qa.ex.xl.CutCopyMode = False
    linhas['M3_colada_sem_editar'] = r3
    # P-03 b) colar e editar: copia de novo e troca NIVEL e RESULTADO
    _, r4 = C.linha_livre_manual(qa)
    lo.ListRows(r2).Range.Copy(lo.ListRows(r4).Range)
    qa.ex.esperar()
    time.sleep(0.5)
    qa.ex.xl.CutCopyMode = False
    escrever(qa, r4, 'NIVEL', '2')
    escrever(qa, r4, 'RESULTADO', '111,111' if BIO else '11,111')
    linhas['M4_colada_editada'] = r4
    vig.parar.set()
    out['caixas_na_entrada'] = vig.vistos
    out['EnableEvents_depois'] = bool(qa.ex.xl.EnableEvents)
    out['linhas'] = linhas
    out['antes_atualizar'] = {k: ler_linha(qa, r) for k, r in linhas.items()}
    try:
        out['qcUltimoIdManual'] = t(qa.wb.Names('qcUltimoIdManual').RefersTo)
    except Exception as e:  # noqa: BLE001
        out['qcUltimoIdManual'] = repr(e)
    # ---- ATUALIZAR DADOS (nucleo do botao)
    vig2 = Vigia(qa.ex.pid, teto=1500)
    vig2.start()
    t0 = time.time()
    try:
        msg, seg = qa.atualizar()
        out['atualizar'] = {'msg': msg[:3000], 's': round(seg, 1)}
    except Exception as e:  # noqa: BLE001
        out['atualizar'] = {'ERRO': repr(e), 's': round(time.time() - t0, 1)}
    vig2.parar.set()
    out['caixas_no_atualizar'] = vig2.vistos
    out['depois_atualizar'] = {k: ler_linha(qa, r) for k, r in linhas.items()}
    fin1 = Final(qa)
    ids = {out['antes_atualizar'][k].get('ID_REGISTRO') for k in linhas}
    out['tblCQ_Final_manuais'] = [{c: (dser(r[c]) if c == 'DATA_HORA' else r[c]) for c in Final.COLS if c in r}
                                  for r in fin1.rows if any(t(r['ID_REGISTRO']).startswith(i) for i in ids if i)]
    try:
        qai = qa.ler('tblQA_Integracao')
        out['QA_Integracao_ids'] = [{k: t(v)[:300] for k, v in x.items() if v not in (None, '')} for x in qai
                                    if any(i and i in ' '.join(t(v) for v in x.values()) for i in ids)][:12]
    except Exception as e:  # noqa: BLE001
        out['QA_Integracao_ids'] = repr(e)
    out['tblCQ_Final_linhas'] = [fin0.n, fin1.n]
    # ---- series depois (as mesmas leituras do Painel), conferidas a mao
    depoisA = serie(qa, fin1, passosA)
    out['serie_A_depois'] = {k: depoisA[k] for k in ('engPainel', 'oraculo', 'M1', 'confere')}
    if BIO:
        depoisB = serie(qa, fin1, passosB)
        out['serie_B_depois'] = {k: depoisB[k] for k in ('engPainel', 'oraculo', 'M1', 'confere')}
        # a mao: n e media do N1 do DIMENSION 2 antes + 123,456
        a = antesB['oraculo']['janela'][0]
        n0, m0 = a['n'], a['media'] or 0
        out['B_N1_esperado_a_mao'] = {'n': n0 + 1, 'media': (n0 * m0 + 123.456) / (n0 + 1),
                                      'painel_depois': depoisB['engPainel'][0][:4]}
    # a mao: serie A, N1 ganha 98,765 (so a origem) e N2 ganha o editado (111,111) SE ele entrar
    a1, a2 = antesA['oraculo']['janela'][0], antesA['oraculo']['janela'][1]
    v1 = float(val2.replace(',', '.'))
    v2 = 111.111 if BIO else 11.111
    out['A_esperado_a_mao'] = {
        'N1_so_origem': {'n': a1['n'] + 1, 'media': (a1['n'] * (a1['media'] or 0) + v1) / (a1['n'] + 1)},
        'N2_com_editado': {'n': a2['n'] + 1, 'media': (a2['n'] * (a2['media'] or 0) + v2) / (a2['n'] + 1)},
        'N2_sem_editado': {'n': a2['n'], 'media': a2['media']},
        'painel_depois': [x[:4] for x in depoisA['engPainel']]}
    return out


def salvar_copia(qa, rot):
    """SaveCopyAs na pasta do Tester (a copia em uso nunca e salva): o XML do estado em uso vai ao valida_r5_xml."""
    destino = os.path.join(C.PASTA, 'xml', f'{PROD}_{ORIG}_{rot}_SaveCopyAs.xlsm')
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    if os.path.exists(destino):
        os.remove(destino)
    qa.wb.SaveCopyAs(destino)
    return {'arquivo': os.path.basename(destino), 'sha12': C.sha(destino)}


# ============================================================================================ main
def main():
    os.makedirs(C.PASTA, exist_ok=True)
    if ORIG == 'base':
        src = os.path.join(C.REPO, BACKUP[PROD], f'QC_{PROD}.xlsm')
    else:
        src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    h0 = C.sha(src)
    suf = ('_manual' if MANUAL else '') + (('_' + '_'.join(SO)) if SO else '') + (('_lote' + LOTE_BIO) if LOTE_BIO != '8974' else '')
    copia = os.path.join(C.PASTA, f'{PROD}_{ORIG}{suf}_{time.strftime("%H%M%S")}.xlsm')
    shutil.copy2(src, copia)
    L = Log(os.path.join(C.PASTA, f'valida_{PROD}_{ORIG}{suf}.json'))
    L['ambiente'] = {'origem': os.path.relpath(src, C.REPO), 'sha_origem_12': h0, 'sha_copia_12': C.sha(copia),
                     'copia': os.path.basename(copia)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    try:
        fin = Final(qa)
        L['tblCQ_Final'] = {'linhas': fin.n, 'hash': fin.hash}
        try:
            criar_usuario(qa, USR, 'ANALISTA')
            L['login'] = fazer_login(qa, USR)
        except Exception as e:  # noqa: BLE001
            L['login'] = {'ERRO': repr(e)}
        qa.ex.xl.EnableEvents = True
        if quer('O72'):
            secao(L, 'O72_ano_literal', s_ano_literal, qa)
            secao(L, 'O72_savecopyas_login', salvar_copia, qa, 'login')
        if quer('O68'):
            secao(L, 'O68_listas', s_listas, qa, senha)
        if quer('painel'):
            secao(L, 'painel', s_painel, qa, fin)
        if MANUAL:
            secao(L, 'manual_P03', s_manual, qa, senha, L)
            secao(L, 'O72_depois_atualizar', s_ano_literal, qa)
            secao(L, 'O72_savecopyas_atualizar', salvar_copia, qa, 'atualizar')
    except Exception as e:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(e), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['origem_intacta'] = {'sha_antes': h0, 'sha_depois': C.sha(src), 'igual': h0 == C.sha(src)}
    try:
        os.remove(copia)
        L['copia_apagada'] = True
    except Exception as e:  # noqa: BLE001
        L['copia_apagada'] = repr(e)
    print('FIM', PROD, ORIG, flush=True)


if __name__ == '__main__':
    t0 = time.time()
    main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
