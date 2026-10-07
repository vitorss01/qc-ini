# -*- coding: utf-8 -*-
"""QA do ADR-071 -- selecao de rodadas do CEQ e Sigma com CV pooled -- numa COPIA instalada.

Uso: python qa_rodadas_cv.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json] [--foto antes.json]
     python qa_rodadas_cv.py --gerar-foto <Bioquimica|Hematologia> <arquivo_antes.xlsm> <antes.json>

Tudo e recalculado aqui em Python direto da EQA_Base (outra implementacao da regra escrita no ADR-071) e
comparado com o que o VBA devolve -- celula a celula na aba e chamada a chamada nas funcoes:
  R00 (com --foto) regressao contra o arquivo ANTES: com TODAS, G/H/R/S/T/AC/AD e AF..AU identicos;
  R01 regressao com TODAS contra a regra ANTIGA (um ano vigente, rotulo da rodada) em todas as linhas;
  R02 dado real: ACUMULADAS / ULTIMAS 3 / ULTIMAS 6 = TODAS quando so existe um ano; K5 diz quantas existem;
  R03 listas = rodadas EXISTENTES por provedor (sem simulacao, sem linha fora do calculo) e validacoes;
  R04 Painel I7..I9 = Estatistica!L do analito em tela e Cfg_PlanoQC!B1 = MIN (D2);
  R05-R08 CEQ SINTETICO injetado na copia (CAP 2024/2025/2026 com rotulos repetidos 'A'/'B', rodada Uso=NAO,
     Controllab 2025/2026 com o mesmo rotulo '1'): 1 rodada, conjunto que atravessa anos, ACUMULADAS, ULTIMAS
     3/6, NRODADAS por ano|rodada, Controllab nunca misturado, provedor vazio = SEM EP, texto invalido = SEM EP
     com aviso em K5, item sem dado avisado, listas, incerteza (AJ..AU) sem seguir a selecao nova (D3).
Nada e salvo no arquivo testado.
"""
import functools
import json
import os
import re
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES  # noqa: E402

SENHA = 'qcini2025'
TOL = 1e-9
FIXOS = ['TODAS', 'ACUMULADAS', 'ULTIMAS 3', 'ULTIMAS 6']
LISTAS = {'lstAnosCAP': 'AF', 'lstAnosCTL': 'AG', 'lstRodadasCAP': 'AH', 'lstRodadasCTL': 'AI', 'lstRodadasEQA': 'AJ'}
# colunas que NAO dependem do Sigma (com TODAS tem de ficar identicas ao arquivo antes do ADR-071)
COLS_IGUAIS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'Q', 'R', 'S', 'T', 'AB', 'AC', 'AD',
               'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU']
COLS_SIGMA = ['L', 'M', 'N', 'O', 'P', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
VIES = ['AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU']
SEM_EP = 'SEM EP'


def ci(letra):
    n = 0
    for ch in letra:
        n = n * 26 + ord(ch) - 64
    return n - 1                                     # indice 0 na linha A..AU


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=TOL):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def igual_valor(a, b):
    if num(a) and num(b):
        return perto(a, b)
    return (a if a is not None else '') == (b if b is not None else '')


# ------------------------------------------------------------------ regra do VBA, escrita de novo
def isnum_vba(v):
    """IsNumeric do VBA: celula vazia (Empty) conta como numero; texto so se tiver cara de numero."""
    if v is None:
        return True
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    s = str(v).strip()
    if not s:
        return False
    try:
        float(s.replace(',', '.'))
        return True
    except ValueError:
        return False


def val(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0.0
    return float(str(v).replace(',', '.')) if isinstance(v, str) else float(v)


def txt(v):
    """CStr do VBA para o que a EQA_Base guarda (2025.0 -> '2025')."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def U(v):
    return txt(v).strip().upper()


def livre(v):
    return U(v) in ('', 'TODOS', 'TODAS')


def teto(ano):
    if ano is None or not isnum_vba(ano) or not txt(ano).strip():
        return 32767
    return int(val(ano))


def modo(rodada):
    """(modo, n, conjunto, erro): -1 invalido, 0 TODAS, 1 rotulo, 2 conjunto, 3 ACUMULADAS, 4 ULTIMAS n."""
    s = U(rodada)
    if s in ('', 'TODAS', 'TODOS'):
        return 0, 0, None, ''
    if s == 'ACUMULADAS':
        return 3, 0, None, ''
    if s[:7] in ('ULTIMAS', 'ÚLTIMAS'):
        a = s[7:].strip()
        if a.isdigit() and len(a) <= 3 and 1 <= int(a) <= 64:
            return 4, int(a), None, ''
        return -1, 0, None, 'ULTIMAS n'
    if '|' not in s and ';' not in s:
        return 1, 0, None, ''
    conj = {}
    for tok in s.split(';'):
        tok = tok.strip()
        if not tok:
            continue
        if '|' not in tok:
            return -1, 0, None, 'item sem ano'
        a, r = tok.split('|', 1)
        a, r = a.strip(), r.strip()
        if not (a.isdigit() and len(a) == 4) or not r or '|' in r:
            return -1, 0, None, 'item invalido'
        conj[a + '|' + r] = 1
    if not conj:
        return -1, 0, None, 'conjunto vazio'
    return 2, 0, conj, ''


def chave(r):
    if not isnum_vba(r[1]) or not txt(r[1]).strip():
        return ''
    return f'{int(val(r[1]))}|{U(r[2])}'


def antes(a1, r1, a2, r2):
    if a1 != a2:
        return a1 < a2
    if r1.isdigit() and r2.isdigit() and len(r1) <= 9 and len(r2) <= 9:
        return int(r1) < int(r2)
    return r1 < r2


def existentes(rows, prov):
    """Rodadas 'ANO|ROTULO' com linha utilizavel (Uso <> NAO, canonico, |bias| preenchido), mais recente primeiro."""
    out = {}
    for r in rows:
        if not U(r[4]) or U(r[19]) == 'NAO':
            continue
        if r[16] is None or not isnum_vba(r[16]):
            continue
        ch = chave(r)
        if not ch:
            continue
        if prov and U(r[0]) != prov.upper():
            continue
        out.setdefault(ch, (int(val(r[1])), U(r[2])))
    cmp = lambda x, y: -1 if antes(x[1][0], x[1][1], y[1][0], y[1][1]) else (1 if antes(y[1][0], y[1][1], x[1][0], x[1][1]) else 0)
    return [k for k, _ in sorted(out.items(), key=functools.cmp_to_key(cmp), reverse=True)]


class Oraculo:
    def __init__(self, rows):
        self.rows = rows
        self._ult = {}

    def casa(self, r, analito, prov, rodada, aref, md):
        if U(r[4]) != analito.strip().upper() or U(r[19]) == 'NAO':
            return False
        m, n, conj, _ = md
        if m < 0:
            return False
        if m >= 2:
            if livre(prov) or U(r[0]) != U(prov):
                return False
            if m == 2:
                return chave(r) in conj
            if m == 4:
                return chave(r) in self.ultimas(prov, aref, n)
            return True
        if not livre(prov) and U(r[0]) != U(prov):
            return False
        if not livre(rodada) and U(r[2]) != U(rodada):
            return False
        return True

    def ultimas(self, prov, aref, n):
        k_ = (U(prov), aref, n)
        if k_ not in self._ult:
            lst = [k for k in existentes(self.rows, U(prov)) if int(k.split('|')[0]) <= aref]
            self._ult[k_] = set(lst[:n])
        return self._ult[k_]

    def selecao(self, analito, ano, prov, rodada, col_vig, col):
        md = modo(rodada)
        aref = teto(ano)
        aref_vig = 32767 if md[0] == 2 else aref
        vig = None
        for r in self.rows:
            if self.casa(r, analito, prov, rodada, aref, md) and isnum_vba(r[1]) and isnum_vba(r[col_vig]):
                a = int(val(r[1])) if txt(r[1]).strip() else 0
                if a <= aref_vig and (vig is None or a > vig):
                    vig = a
        if vig is None:
            return None
        out = []
        for r in self.rows:
            if not (self.casa(r, analito, prov, rodada, aref, md) and isnum_vba(r[1]) and isnum_vba(r[col])):
                continue
            a = int(val(r[1])) if txt(r[1]).strip() else 0
            if (md[0] >= 2 and a <= vig) or (md[0] < 2 and a == vig):
                out.append(r)
        return out

    def bias(self, analito, ano, prov, rodada, metrica, legado=False):
        col = 15 if metrica == 'SIGNED' else 16
        linhas = self.selecao(analito, ano, prov, rodada, 16, col)
        if not linhas:
            return SEM_EP
        grupos = {}
        for r in linhas:
            k = U(r[2]) if legado else chave(r)
            g = grupos.setdefault(k, [0.0, 0.0, 0])
            if isnum_vba(r[16]):
                g[0] += val(r[16])
            if isnum_vba(r[15]):
                g[1] += val(r[15])
            g[2] += 1
        if metrica == 'N':
            return sum(g[2] for g in grupos.values())
        if metrica == 'NRODADAS':
            return len(grupos)
        i = 1 if metrica == 'SIGNED' else 0
        return sum(g[i] / g[2] for g in grupos.values()) / len(grupos)

    def sdi(self, analito, ano, prov, rodada):
        linhas = self.selecao(analito, ano, prov, rodada, 9, 9)
        if not linhas:
            return None
        return max(abs(val(r[9])) for r in linhas), len(linhas)

    def limites(self, analito, ano, prov, rodada):
        linhas = self.selecao(analito, ano, prov, rodada, 6, 6)
        if not linhas:
            return SEM_EP
        dentro = fora = sem = 0
        for r in linhas:
            if not isnum_vba(r[10]) or not isnum_vba(r[11]):
                sem += 1
            elif val(r[6]) < val(r[10]) or val(r[6]) > val(r[11]):
                fora += 1
            else:
                dentro += 1
        if fora:
            return f'NAO OK ({fora} fora dos limites)'
        if sem:
            return f'OK ({dentro} dentro; {sem} sem limite)'
        return f'OK ({dentro} dentro dos limites)'

    def resumo(self, prov, ano, rodada):
        """('INVALIDA'|'PROVEDOR'|None, n_rodadas, n_linhas) -- a contagem do K5."""
        md = modo(rodada)
        if md[0] < 0:
            return 'INVALIDA', 0, 0
        if md[0] >= 2 and livre(prov):
            return 'PROVEDOR', 0, 0
        aref = teto(ano)
        vig = None
        if md[0] < 2:
            for r in self.rows:
                if U(r[4]) and self.casa(r, U(r[4]), prov, rodada, aref, md) and isnum_vba(r[1]) and isnum_vba(r[16]):
                    a = int(val(r[1])) if txt(r[1]).strip() else 0
                    if a <= aref and (vig is None or a > vig):
                        vig = a
        rods, n = set(), 0
        for r in self.rows:
            if not (U(r[4]) and self.casa(r, U(r[4]), prov, rodada, aref, md) and isnum_vba(r[1]) and isnum_vba(r[16])):
                continue
            a = int(val(r[1])) if txt(r[1]).strip() else 0
            if md[0] < 2 and a != vig:
                continue
            if md[0] in (3, 4) and a > aref:
                continue
            n += 1
            rods.add(chave(r))
        return None, len(rods), n


def status_sdi_ok(texto, esperado):
    """Compara 'OK (|SDI| max 1,23 em 15 amostra(s))' / 'FORA (|SDI| max 2,50 > 2)' com (max, n)."""
    if esperado is None:
        return texto == SEM_EP
    mx, n = esperado
    m = re.match(r'^(OK|FORA) \(\|SDI\| max ([0-9.,]+)', str(texto))
    if not m:
        return False
    s = m.group(2)
    if ',' in s:                                     # Format$ do VBA no idioma da instalacao (pt-BR: 1,23)
        s = s.replace('.', '').replace(',', '.')
    if abs(float(s) - mx) > 0.0051:
        return False
    if (m.group(1) == 'OK') != (mx <= 2.0):
        return False
    return m.group(1) == 'FORA' or f' em {n} amostra' in str(texto)


# ------------------------------------------------------------------ leitura da pasta
def ler_base(q):
    b = q.wb.Worksheets('EQA_Base')
    ult = max(b.Cells(b.Rows.Count, 1).End(-4162).Row, b.Cells(b.Rows.Count, 5).End(-4162).Row, 2)
    return [list(r) for r in b.Range(b.Cells(2, 1), b.Cells(ult, 21)).Value]


def linhas_tabela(e, ult):
    vals = e.Range(f'A14:AU{ult}').Value
    return [(14 + i, list(v)) for i, v in enumerate(vals)]


def ultima_linha(e):
    ult = 14
    for r in range(14, 400):
        if str(e.Cells(r, 28).Formula or '').startswith('='):
            ult = r
        elif r > ult + 3:
            break
    return ult


def provedores_linha(q, bio):
    """Provedor que a formula de cada linha passa ao mCEQ: Bio = Analitos!AR do analito; Hema = eqProvedor."""
    if not bio:
        p = q.wb.Names('eqProvedor').RefersToRange.Value
        return lambda an: p
    a = q.wb.Worksheets('Analitos')
    mapa = {}
    for nome, pv in zip(a.Range('A4:A43').Value, a.Range('AR4:AR43').Value):
        if nome[0]:
            mapa[str(nome[0]).strip().upper()] = pv[0] if pv[0] not in (None,) else ''
    return lambda an: mapa.get(str(an).strip().upper(), 'CAP')


def foto(q):
    e = q.wb.Worksheets('Estatística')
    ult = ultima_linha(e)
    return {'produto': q.produto, 'ult': ult, 'linhas': [v for _, v in linhas_tabela(e, ult)],
            'K5': str(e.Range('K5').Text)}


def recalc(q):
    q.ex.esperar()
    q.ex.xl.Calculate()
    q.ex.esperar()


def definir(q, e, a, v):
    q.ex.xl.EnableEvents = False
    try:
        e.Range(a).Value = v
    finally:
        q.ex.xl.EnableEvents = True
    recalc(q)


# ------------------------------------------------------------------ testes
def r00_foto(q, e, ult, foto_antes):
    if not foto_antes:
        print('   R00 sem --foto: regressao contra o arquivo antes nao executada (R01 cobre com a regra antiga)', flush=True)
        return
    agora = [v for _, v in linhas_tabela(e, ult)]
    dif, mud = [], {c: 0 for c in COLS_SIGMA}
    if foto_antes['ult'] != ult:
        dif.append(('ultima linha', foto_antes['ult'], ult))
    for i, (a0, a1) in enumerate(zip(foto_antes['linhas'], agora)):
        for c in COLS_IGUAIS:
            if not igual_valor(a0[ci(c)], a1[ci(c)]):
                dif.append((14 + i, c, a0[ci(c)], a1[ci(c)]))
        for c in COLS_SIGMA:
            if not igual_valor(a0[ci(c)], a1[ci(c)]):
                mud[c] += 1
    reg('R00 Regressão contra o arquivo ANTES do ADR-071 (mesma base, TODAS): G, H, R, S, T, AC, AD e AF..AU idênticos '
        'célula a célula; só o que depende do Sigma (L e M..P, U..AA) muda',
        not dif, {'linhas': len(agora), 'colunas_conferidas': len(COLS_IGUAIS), 'diferencas': dif[:8],
                  'mudaram_por_depender_do_sigma': mud, 'K5_antes': foto_antes['K5'][:90], 'K5_agora': str(e.Range('K5').Text)[:110]})


def r01_legado(q, e, ult, orc, prov_de):
    ano = q.wb.Names('eqAnoEP').RefersToRange.Value
    rod = q.wb.Names('eqRodada').RefersToRange.Value
    dif, n = [], 0
    for r, v in linhas_tabela(e, ult):
        an = v[0]
        if an in (None, '', 0):
            continue
        n += 1
        pv = prov_de(an)
        for c, met in (('G', 'ABS'), ('T', 'SIGNED'), ('AC', 'N'), ('AD', 'NRODADAS')):
            esp = orc.bias(str(an), ano, pv, rod, met, legado=True)
            if not igual_valor(v[ci(c)], esp) and not (num(v[ci(c)]) and num(esp) and perto(v[ci(c)], esp, 1e-9)):
                dif.append((r, an, c, v[ci(c)], esp))
        if not status_sdi_ok(v[ci('R')], orc.sdi(str(an), ano, pv, rod)):
            dif.append((r, an, 'R', v[ci('R')], orc.sdi(str(an), ano, pv, rod)))
        esp_s = orc.limites(str(an), ano, pv, rod)
        if v[ci('S')] != esp_s:
            dif.append((r, an, 'S', v[ci('S')], esp_s))
    reg('R01 Regressão com TODAS: G (|bias| ABS em 2 etapas), T, AC, AD, R e S = regra ANTIGA recalculada da EQA_Base '
        '(um ano vigente <= Ano EP, rodada pelo rótulo) em todas as linhas',
        U(rod) in ('', 'TODAS', 'TODOS') and not dif and n > 0,
        {'rodada': rod, 'ano_ep': ano, 'linhas': n, 'divergencias': dif[:8]})


def colunas(e, ult, cs):
    out = {}
    for r, v in linhas_tabela(e, ult):
        if v[0] in (None, '', 0):
            continue
        out[r] = tuple(v[ci(c)] for c in cs)
    return out


def r02_real(q, e, ult, orc, bio):
    p4 = q.wb.Names('eqRodada').RefersToRange
    rod0 = p4.Value
    base = colunas(e, ult, ('G', 'T', 'AC', 'AD'))
    ano = q.wb.Names('eqAnoEP').RefersToRange.Value
    prov = q.wb.Names('eqProvedor').RefersToRange.Value
    anos = sorted({int(k.split('|')[0]) for k in existentes(orc.rows, '' if bio else U(prov)) if int(k.split('|')[0]) <= teto(ano)})
    ev, ok = {}, True
    for texto in ('ACUMULADAS', 'ULTIMAS 3', 'ULTIMAS 6'):
        t0 = time.perf_counter()
        definir(q, e, 'P4', texto)
        dtm = round(time.perf_counter() - t0, 2)
        agora = colunas(e, ult, ('G', 'T', 'AC', 'AD'))
        dif = [(r, base[r], agora[r]) for r in base if not all(igual_valor(a, b) for a, b in zip(base[r], agora[r]))]
        k5 = str(e.Range('K5').Text)
        _, nrod, nlin = orc.resumo(prov, ano, texto)
        k5_ok = f'{nrod} rodada(s), {nlin} linha(s)' in k5 and not k5.startswith('#')
        if texto == 'ULTIMAS 6' and nrod < 6:
            k5_ok = k5_ok and f'(so {nrod} existe(m) com dado)' in k5
        ev[texto] = {'segundos': dtm, 'diferencas': dif[:4], 'K5': k5[:140]}
        # so um ano de CEQ real: os modos novos tem de dar exatamente o que TODAS da
        if len(anos) == 1:
            ok = ok and not dif
        ok = ok and k5_ok
    definir(q, e, 'P4', rod0)
    reg('R02 Dado real (só um ano de CEQ): ACUMULADAS, ULTIMAS 3 e ULTIMAS 6 dão exatamente o G/T/AC/AD de TODAS; '
        'K5 conta rodadas e linhas pela mesma regra e avisa quando existem menos de 6 rodadas (3 ≠ 6)',
        ok, {'anos_ceq': anos, 'modos': ev})


def r03_listas(q, e, orc, bio, codigo, rotulo):
    cfg = q.wb.Worksheets('Configuração')
    dif = []
    for nome, prov in (('lstRodadasCAP', 'CAP'), ('lstRodadasCTL', 'CONTROLLAB'), ('lstRodadasEQA', '')):
        col = LISTAS[nome]
        vals = [str(v[0]) for v in cfg.Range(f'{col}2:{col}101').Value if v[0] not in (None, '')]
        esp = FIXOS + existentes(orc.rows, prov)
        if [x.upper() for x in vals] != [x.upper() for x in esp]:
            dif.append((nome, vals[:8], esp[:8]))
        try:
            rr = q.wb.Names(nome).RefersToRange
            if rr.Rows.Count != max(1, len(vals)):
                dif.append((nome, 'nome nao cobre a lista', rr.Address, len(vals)))
        except Exception as ex:                       # noqa: BLE001
            dif.append((nome, 'nome ausente', str(ex)[:60]))
    for nome, prov in (('lstAnosCAP', 'CAP'), ('lstAnosCTL', 'CONTROLLAB')):
        col = LISTAS[nome]
        vals = [int(v[0]) for v in cfg.Range(f'{col}2:{col}101').Value if v[0] not in (None, '')]
        esp = sorted({int(k.split('|')[0]) for k in existentes(orc.rows, prov)})
        if vals != esp:
            dif.append((nome, vals, esp))
    anos_ceq = [int(v[0]) for v in cfg.Range('AA2:AA50').Value if v[0] not in (None, '')]
    esp_ceq = sorted({int(k.split('|')[0]) for k in existentes(orc.rows, '')}) or [0]
    if anos_ceq != esp_ceq:
        dif.append(('lstAnosCEQ', anos_ceq, esp_ceq))
    # rodadas que so tem linha Uso = NAO (simulacao) nunca entram
    so_nao = set()
    tem_uso = set()
    for r in orc.rows:
        if U(r[4]) and chave(r):
            (so_nao if U(r[19]) == 'NAO' else tem_uso).add(chave(r))
    simul = sorted(so_nao - tem_uso)
    listadas = {str(x).upper() for row in cfg.Range('AH2:AJ101').Value for x in row if x}
    vaz = [s for s in simul if s in listadas]
    if vaz:
        dif.append(('simulacao na lista', vaz))
    # validacoes (lidas no idioma da instalacao: so conferimos que apontam para as listas)
    p4 = e.Range('P4').Validation
    f1 = str(p4.Formula1)
    val_ok = p4.Type == 3 and p4.AlertStyle == 2 and 'lstRodadas' in f1
    n4f = str(e.Range('N4').Validation.Formula1)
    val_ok = val_ok and 'lstAnos' in n4f
    ev = {'P4': f1[:80], 'N4': n4f[:80]}
    if not bio:
        r4 = e.Range('R4').Validation
        val_ok = val_ok and r4.Type == 3 and r4.IgnoreBlank is False and 'CAP' in str(r4.Formula1)
        val_ok = val_ok and 'lstRodadasCTL' in f1 and 'lstRodadasCAP' in f1
        ev['R4'] = [str(r4.Formula1), r4.IgnoreBlank]
    else:
        val_ok = val_ok and 'lstRodadasEQA' in f1
    k5 = str(e.Range('K5').Text)
    k5f = str(e.Range('K5').Formula)
    val_ok = val_ok and k5f.upper().startswith('=RESUMOFILTROEQ(') and 'TEXT(' not in k5f.upper() and 'linha(s)' in k5
    reg(f'{codigo} {rotulo}: listas de seleção = rodadas EXISTENTES por provedor (Uso ≠ NAO, analito canônico, |bias|), '
        f'itens fixos TODAS/ACUMULADAS/ULTIMAS 3/ULTIMAS 6 antes; nunca rodada de simulação; validações de N4/P4 '
        f'(estilo Aviso) apontam para as listas{"" if bio else ", R4 sem vazio"}; K5 = ResumoFiltroEQ (sem TEXTO)',
        not dif and val_ok, {'divergencias': dif[:6], 'simulacao_existente': simul[:6], 'validacoes': ev,
                              'K5': k5[:120]})


def r04_painel(q, e, ult, bio):
    nlv = 2 if bio else 3
    p = q.wb.Worksheets('Painel')
    c = q.wb.Worksheets('Cfg_PlanoQC')
    try:
        p.Unprotect(SENHA)
    except Exception:
        pass
    cad = [str(v[0]).strip() for v in q.wb.Worksheets('Analitos').Range('A4:A43').Value if v[0]]
    chaves = {str(v[ci('AB')]): (r, v) for r, v in linhas_tabela(e, ult)}
    sem_etp = [a for a in cad if (chaves.get(f'{a}|1') or (0, [None] * 47))[1][ci('K')] == '-']
    alvo = cad[:3] + sem_etp[:1]
    b30 = p.Range('B3').Value
    dif, ev = [], {}
    for an in alvo:
        q.ex.xl.EnableEvents = False
        try:
            p.Range('B3').Value = cad.index(an) + 1
        finally:
            q.ex.xl.EnableEvents = True
        q.run('mEstatistica.PainelMudou')
        recalc(q)
        iv = [p.Range(f'I{6 + k}').Value for k in range(1, nlv + 1)]
        lv = [(chaves.get(f'{an}|{k}') or (0, None))[0] for k in range(1, nlv + 1)]
        lv = [e.Range(f'L{r}').Value if r else '' for r in lv]
        b1 = c.Range('B1').Value
        nums = [x for x in iv if num(x)]
        b1_esp = min(nums) if nums else ('-' if '-' in iv else 'SEM DADOS')
        ok = all(igual_valor(a, b) for a, b in zip(iv, lv)) and igual_valor(b1, b1_esp)
        ev[an] = {'I': iv, 'L': lv, 'B1': b1}
        if not ok:
            dif.append((an, iv, lv, b1, b1_esp))
    q.ex.xl.EnableEvents = False
    try:
        p.Range('B3').Value = b30
    finally:
        q.ex.xl.EnableEvents = True
    q.run('mEstatistica.PainelMudou')
    recalc(q)
    rot = str(p.Range('I6').Text)
    reg('R04 Painel: Sigma por nível (I7..I9) = Estatística!L do analito em tela (CV pooled de N meses; sem bias = '
        'vazio, nunca bias 0) e Cfg_PlanoQC!B1 (sigmaDoPlano) = MIN dos níveis; "-" sem ETp (ADR-068)',
        not dif and 'CVp' in rot, {'analitos': ev, 'rotulo_I6': rot, 'divergencias': dif[:4]})


# ------------------------------------------------------------------ CEQ sintetico
SINT = [('CAP', 2024, 'A', [2.0, -4.0], 'SIM'), ('CAP', 2024, 'B', [6.0, 8.0], 'SIM'),
        ('CAP', 2026, 'A', [-1.0, 3.0], 'SIM'), ('CAP', 2026, 'B', [5.0, 5.0], 'SIM'),
        ('CAP', 2026, 'Z', [99.0, 99.0], 'NAO'),
        ('Controllab', 2025, '1', [40.0, 42.0], 'SIM'), ('Controllab', 2026, '1', [60.0, 62.0], 'SIM')]


def injetar(q, analito):
    b = q.wb.Worksheets('EQA_Base')
    try:
        b.Unprotect(SENHA)
    except Exception:
        pass
    ult = max(b.Cells(b.Rows.Count, 1).End(-4162).Row, b.Cells(b.Rows.Count, 5).End(-4162).Row)
    linhas = []
    for prov, ano, rod, bs, uso in SINT:
        for i, x in enumerate(bs, 1):
            am = f'QA71-{ano}{rod}{i}'
            linhas.append((prov, ano, rod, 'QA ' + analito, analito, am, 100 + x, 100.0, 2.0, x / 2, 80.0, 120.0,
                           'Acceptable', 'ACEITO', '', x, abs(x), '', 'QA ADR-071 sintetico', uso,
                           f'{prov}|{ano}|{rod}|{analito}|{am}'))
    b.Range(b.Cells(ult + 1, 1), b.Cells(ult + len(linhas), 21)).Value = tuple(linhas)
    b.Cells(1, 26).Value = 'consolidado: QA ADR-071 sintetico ' + time.strftime('%H:%M:%S')
    q.run('mCEQ.InvalidarEQ')
    return len(linhas)


def r05_funcoes(q, orc, analito):
    casos = [('TODAS', 2026, 'CAP'), ('TODAS', 2025, 'CAP'), ('A', 2026, 'CAP'), ('A', 2025, 'CAP'),
             ('2024|A', 2026, 'CAP'), ('2024|A; 2026|A', 2026, 'CAP'), (' 2024|a ;2026|b; ', 2000, 'CAP'),
             ('ACUMULADAS', 2026, 'CAP'), ('ACUMULADAS', 2025, 'CAP'), ('ULTIMAS 3', 2026, 'CAP'),
             ('ULTIMAS 6', 2026, 'CAP'), ('ULTIMAS 3', 2025, 'CAP'), ('últimas 2', 2026, 'CAP'),
             ('ULTIMAS 1', 2026, 'CAP'),
             ('ACUMULADAS', 2026, 'Controllab'), ('ULTIMAS 1', 2026, 'Controllab'), ('TODAS', 2026, 'Controllab'),
             ('2025|1; 2026|1', 2026, 'Controllab'),
             ('ACUMULADAS', 2026, ''), ('2024|A', 2026, ''), ('ULTIMAS 3', 2026, 'TODOS'), ('TODAS', 2026, ''),
             ('2025|', 2026, 'CAP'), ('C-A 2025; C-B 2025', 2026, 'CAP'), ('ULTIMAS 0', 2026, 'CAP'),
             ('ULTIMAS X', 2026, 'CAP'), ('2025|A|B', 2026, 'CAP'), ('25|A', 2026, 'CAP')]
    dif, ev = [], {}
    for rod, ano, prov in casos:
        linha = {}
        for met in ('ABS', 'SIGNED', 'N', 'NRODADAS'):
            got = q.run('mCEQ.BiasEQ', analito, ano, met, prov, rod)
            esp = orc.bias(analito, ano, prov, rod, met)
            linha[met] = got
            if not igual_valor(got, esp):
                dif.append((rod, ano, prov, met, got, esp))
        ev[f'{rod!r}/{ano}/{prov or "-"}'] = linha
    # afirmacoes diretas (alem do oraculo)
    diretas = {
        'rotulo repetido em anos diferentes = 2 rodadas': q.run('mCEQ.BiasEQ', analito, 2026, 'NRODADAS', 'CAP', '2024|A; 2026|A') == 2,
        'Controllab ACUMULADAS so com Controllab (4 amostras)': q.run('mCEQ.BiasEQ', analito, 2026, 'N', 'Controllab', 'ACUMULADAS') == 4,
        'CAP nunca leva o Controllab (bias 40-62 ausente)': all(
            (not num(q.run('mCEQ.BiasEQ', analito, 2026, 'ABS', 'CAP', t))) or q.run('mCEQ.BiasEQ', analito, 2026, 'ABS', 'CAP', t) < 30
            for t in ('ACUMULADAS', 'ULTIMAS 6', '2025|1; 2026|1')),
        'Uso=NAO (2026|Z, bias 99) nunca entra': q.run('mCEQ.BiasEQ', analito, 2026, 'ABS', 'CAP', '2026|Z') == SEM_EP,
        'provedor vazio nos modos novos = SEM EP (D1)': q.run('mCEQ.BiasEQ', analito, 2026, 'ABS', '', 'ACUMULADAS') == SEM_EP,
        'texto invalido = SEM EP (nunca #VALOR!)': q.run('mCEQ.BiasEQ', analito, 2026, 'ABS', 'CAP', '2025|') == SEM_EP,
        'conjunto atravessando anos ignora o teto': num(q.run('mCEQ.BiasEQ', analito, 2000, 'ABS', 'CAP', '2024|A; 2026|B')),
    }
    # SDI e limites seguem a mesma selecao (ano por item / acumulado)
    for rod in ('2024|A; 2026|A', 'ACUMULADAS', 'ULTIMAS 3'):
        got_n = q.run('mCEQ.SDIeq', analito, 2026, 'N', 'CAP', rod)
        esp = orc.sdi(analito, 2026, 'CAP', rod)
        if (esp is None and got_n != SEM_EP) or (esp is not None and got_n != esp[1]):
            dif.append(('SDIeq N', rod, got_n, esp))
        got_s = q.run('mCEQ.StatusLimitesEQ', analito, 2026, 'CAP', rod)
        if got_s != orc.limites(analito, 2026, 'CAP', rod):
            dif.append(('StatusLimitesEQ', rod, got_s, orc.limites(analito, 2026, 'CAP', rod)))
    # D3: a incerteza nao segue a selecao nova
    d3 = []
    for rod in ('ULTIMAS 3', 'ACUMULADAS', '2024|A; 2026|B', '2025|'):
        for met in ('N', 'NRODADAS', 'MEDIA', 'SITUACAO', 'UBIAS'):
            a_ = q.run('mCEQ.ViesEQ', analito, 2026, met, 'CAP', rod)
            b_ = q.run('mCEQ.ViesEQ', analito, 2026, met, 'CAP', 'TODAS')
            if not igual_valor(a_, b_):
                d3.append((rod, met, a_, b_))
    reg('R05 CEQ sintético (CAP 2024/2025/2026 com rótulos repetidos, rodada Uso=NAO, Controllab 2025/2026): BiasEQ '
        'ABS/SIGNED/N/NRODADAS = oráculo em 1 rodada, conjunto (atravessando anos), ACUMULADAS, ULTIMAS n, TODAS e rótulo '
        '(modos antigos intactos); NRODADAS conta ano|rodada; CAP e Controllab nunca misturados; provedor vazio e texto '
        'inválido = SEM EP; SDI e limites pela mesma seleção; incerteza (ViesEQ) igual a TODAS nos modos novos (D3)',
        not dif and all(diretas.values()) and not d3,
        {'casos': len(casos), 'divergencias': dif[:8], 'diretas': diretas, 'D3_divergencias': d3[:6],
         'exemplos': dict(list(ev.items())[:6])})


def r06_resumo(q, orc):
    casos = [('CAP', 2026, 'ULTIMAS 3', 'ultimas 3'), ('CAP', 2026, 'ULTIMAS 6', 'ultimas 6'),
             ('CAP', 2025, 'ACUMULADAS', 'acumuladas ate 2025'), ('CAP', 2026, 'ACUMULADAS', 'acumuladas ate 2026'),
             ('CAP', 2026, '2024|A; 2024|Q', 'SEM DADO: 2024|Q'), ('Controllab', 2026, 'ULTIMAS 1', 'ultimas 1'),
             ('CAP', 2026, 'TODAS', 'media de todas as rodadas'), ('CAP', 2026, '2025|', 'INVALIDA'),
             ('', 2026, 'ACUMULADAS', 'provedor nao definido'), ('CAP', 2026, 'C-A 2025; C-B 2025', 'item sem ano')]
    dif, ev = [], {}
    for prov, ano, rod, palavra in casos:
        t = str(q.run('mCEQ.ResumoFiltroEQ', prov, ano, rod))
        tipo, nrod, nlin = orc.resumo(prov, ano, rod)
        ok = palavra in t and f'{nrod} rodada(s), {nlin} linha(s) analiticas' in t
        ev[f'{prov or "-"}/{ano}/{rod}'] = t[:150]
        if not ok:
            dif.append((prov, ano, rod, t[:160], (tipo, nrod, nlin)))
    reg('R06 K5 (ResumoFiltroEQ) descreve o modo e conta rodadas e linhas analíticas pela regra do cálculo; texto inválido '
        'diz INVALIDA, item do conjunto sem dado aparece como SEM DADO, provedor vazio nos modos novos é avisado',
        not dif, {'textos': ev, 'divergencias': dif[:5]})


def r07_listas_sint(q, e, orc, bio):
    q.run('mDados.AtualizarListasAno')
    r03_listas(q, e, orc, bio, 'R07', 'CEQ sintético (CAP 2024/2026, Controllab, rodada Uso=NAO)')


def r08_aba(q, e, ult, orc, bio, prov_de):
    n4, p4 = e.Range('N4'), e.Range('P4')
    n4f, p40 = n4.Formula, p4.Value
    r4v = e.Range('R4').Value if not bio else None
    definir(q, e, 'N4', 2026)
    definir(q, e, 'P4', 'TODAS')
    viesT = colunas(e, ult, VIES)
    t0 = time.perf_counter()
    definir(q, e, 'P4', 'ULTIMAS 3')
    seg = round(time.perf_counter() - t0, 2)
    dif = []
    for r, v in linhas_tabela(e, ult):
        an = v[0]
        if an in (None, '', 0):
            continue
        pv = prov_de(an)
        for c, met in (('G', 'ABS'), ('T', 'SIGNED'), ('AC', 'N'), ('AD', 'NRODADAS')):
            esp = orc.bias(str(an), 2026, pv, 'ULTIMAS 3', met)
            if not igual_valor(v[ci(c)], esp):
                dif.append((r, an, c, v[ci(c)], esp))
    k5 = str(e.Range('K5').Text)
    prov_k5 = q.wb.Names('eqProvedor').RefersToRange.Value
    _, nrod, nlin = orc.resumo(prov_k5, 2026, 'ULTIMAS 3')
    k5_ok = 'ultimas 3' in k5 and f'{nrod} rodada(s), {nlin} linha(s)' in k5
    viesU = colunas(e, ult, VIES)
    d3 = [(r, viesT[r], viesU[r]) for r in viesT if not all(igual_valor(a, b) for a, b in zip(viesT[r], viesU[r]))]
    # texto invalido: SEM EP em G e aviso em K5
    definir(q, e, 'P4', '2025|')
    g_inv = [v[ci('G')] for r, v in linhas_tabela(e, ult) if v[0] not in (None, '', 0)]
    k5_inv = str(e.Range('K5').Text)
    inv_ok = all(g == SEM_EP for g in g_inv) and 'INVALIDA' in k5_inv
    # provedor vazio nos modos novos (so a Hematologia tem o provedor global)
    prov_ok, k5_prov = True, ''
    if not bio:
        definir(q, e, 'R4', '')
        definir(q, e, 'P4', 'ACUMULADAS')
        g_pv = [v[ci('G')] for r, v in linhas_tabela(e, ult) if v[0] not in (None, '', 0)]
        k5_prov = str(e.Range('K5').Text)
        prov_ok = all(g == SEM_EP for g in g_pv) and 'provedor nao definido' in k5_prov
        definir(q, e, 'R4', r4v)
    definir(q, e, 'P4', p40)
    q.ex.xl.EnableEvents = False
    try:
        n4.Formula = n4f
    finally:
        q.ex.xl.EnableEvents = True
    recalc(q)
    reg('R08 Aba Estatística com o CEQ sintético: P4 = "ULTIMAS 3" (Ano EP 2026) recalcula G/T/AC/AD de TODAS as linhas '
        'igual ao oráculo (provedor de cada linha), K5 conta igual, a incerteza AF..AU não muda (D3); P4 inválido = SEM EP '
        'em todas as linhas com aviso em K5' + ('' if bio else '; provedor vazio + ACUMULADAS = SEM EP com aviso (D1)'),
        not dif and k5_ok and not d3 and inv_ok and prov_ok,
        {'segundos_troca_P4': seg, 'divergencias': dif[:6], 'K5': k5[:140], 'esperado_K5': [nrod, nlin],
         'incerteza_mudou': d3[:3], 'K5_invalido': k5_inv[:120], 'K5_sem_provedor': k5_prov[:140]})


def executar(produto, caminho, saida, foto_antes=None):
    q = QA(produto, caminho)
    bio = produto.startswith('Bio')
    try:
        q.ex.xl.Calculation = -4105
        q.run('mIncerteza.RecalcularIncerteza')
        q.ex.esperar()
        e = q.wb.Worksheets('Estatística')
        e.Unprotect(SENHA)
        ult = ultima_linha(e)
        orc = Oraculo(ler_base(q))
        prov_de = provedores_linha(q, bio)
        r00_foto(q, e, ult, foto_antes)
        r01_legado(q, e, ult, orc, prov_de)
        r02_real(q, e, ult, orc, bio)
        r03_listas(q, e, orc, bio, 'R03', 'Dado real')
        r04_painel(q, e, ult, bio)
        # ---- CEQ sintetico (so nesta copia; nada e salvo)
        cand = [v[0] for _, v in linhas_tabela(e, ult) if v[0] not in (None, '', 0) and num(v[ci('G')])]
        analito = 'WBC' if (not bio and 'WBC' in cand) else cand[0]
        n = injetar(q, str(analito))
        orc = Oraculo(ler_base(q))
        print(f'   CEQ sintetico: {n} linhas injetadas para {analito}', flush=True)
        r05_funcoes(q, orc, str(analito))
        r06_resumo(q, orc)
        r07_listas_sint(q, e, orc, bio)
        r08_aba(q, e, ult, orc, bio, prov_de)
    finally:
        q.fechar()
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [x for x in RES if x['resultado'] == 'FAIL']
    print(f'\n=== {produto} rodadas/CV: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


def gerar_foto(produto, caminho, destino):
    q = QA(produto, caminho)
    try:
        q.ex.xl.Calculation = -4105
        q.run('mIncerteza.RecalcularIncerteza')
        q.ex.esperar()
        f = foto(q)
    finally:
        q.fechar()
    with open(destino, 'w', encoding='utf-8') as fp:
        json.dump(f, fp, ensure_ascii=False, default=str)
    print(f'foto gravada: {destino} ({len(f["linhas"])} linhas)', flush=True)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--gerar-foto':
        gerar_foto(a[1], os.path.abspath(a[2]), a[3])
        sys.exit(0)
    fa = None
    if '--foto' in a:
        i = a.index('--foto')
        with open(a[i + 1], encoding='utf-8') as fp:
            fa = json.load(fp)
        del a[i:i + 2]
    falhas = executar(a[0], os.path.abspath(a[1]), a[2] if len(a) > 2 else None, fa)
    sys.exit(1 if falhas else 0)
