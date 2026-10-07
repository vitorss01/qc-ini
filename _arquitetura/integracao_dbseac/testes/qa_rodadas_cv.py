# -*- coding: utf-8 -*-
"""QA do ADR-071 -- selecao de rodadas do CEQ e Sigma com CV pooled -- numa COPIA instalada.

Uso: python qa_rodadas_cv.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json] [--antes <arquivo_antes.xlsm>]
                             [--foto antes.json] [--sem-r00]
     python qa_rodadas_cv.py --gerar-foto <Bioquimica|Hematologia> <arquivo_antes.xlsm> <antes.json>

  --antes ..... o arquivo de ANTES desta entrega (o entregar.py passa a producao original, copiada antes de
                instalar; tambem pela variavel de ambiente QC_PROD_ORIGINAL). Nunca e aberto: a suite abre uma
                copia propria dele. Sem --antes/--foto o R00 FALHA (a regressao do gate nao passa calada);
                --sem-r00 pula explicitamente (uso manual).

Tudo e recalculado aqui em Python direto da EQA_Base (outra implementacao da regra escrita no ADR-071) e
comparado com o que o VBA devolve -- celula a celula na aba e chamada a chamada nas funcoes:
  R00 regressao contra o arquivo ANTES (P4 = TODAS nas DUAS copias): G/H/R/S/T/AC/AD, AF..AU e o resto que nao
      depende do Sigma identicos; coluna cuja formula mudou nesta entrega (ADR-068/069...) e celula cujo valor muda
      pela regra "valor numerico de verdade" (|bias|/limite vazio, revisao 07/10/2026) ficam fora, listadas;
  R01 regressao com TODAS (forcado NA COPIA DE TESTE) contra a regra ANTIGA (um ano vigente, rotulo da rodada);
  R02 dado real: ACUMULADAS / ULTIMAS 3 / ULTIMAS 6 = TODAS quando so existe um ano; K5 diz quantas existem;
  R03 listas = rodadas EXISTENTES por provedor (sem simulacao, sem linha fora do calculo) e validacoes;
  R04 Painel I7..I9 = Estatistica!L, Cfg_PlanoQC!B1 = MIN (D2), O3 avisa "SIGMA EM k DE N NIVEIS" quando um nivel
      com dados fica sem Sigma, e no bloco DESEMPENHO o CVp ao lado fecha a conta do Sigma;
  R05-R08 CEQ SINTETICO injetado na copia (CAP 2024/2025/2026 com rotulos repetidos 'A'/'B', rodada Uso=NAO,
     rodada SEM ALVO (|bias| vazio), segundo analito com outra familia de survey no mesmo ano (LN2-x), Controllab
     2025/2026 com o mesmo rotulo '1' -- o Controllab real do analito e retirado da copia): 1 rodada, conjunto que
     atravessa anos, ACUMULADAS, ULTIMAS n POR ANALITO, NRODADAS por ano|rodada, Controllab nunca misturado,
     provedor vazio = SEM EP, texto invalido = SEM EP com aviso em K5, K5 com ano vigente por analito, listas,
     incerteza (AJ..AU) sem seguir a selecao nova (D3);
  R09 depois de mudar a EQA_Base, o fim do AtualizarEQABase (mCEQ.MarcarCEQSujo + Calculate) atualiza G e K5;
  R10 mais de 64 rodadas (ACUMULADAS): todas entram, nenhuma truncada;
  R11 celula com valor de erro na EQA_Base (#DIV/0! no |bias|, #N/D na rodada): o cache monta, as outras linhas
      calculam, K5 nao diz "ilegivel" e as listas atualizam sem erro.
Nada e salvo no arquivo testado.
"""
import functools
import json
import os
import re
import shutil
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
COLS_CEQ = ['G', 'R', 'S', 'T', 'AC', 'AD']          # o R00 nao vale sem compara-las
VIES = ['AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU']
SEM_EP = 'SEM EP'


def ci(letra):
    n = 0
    for ch in letra:
        n = n * 26 + ord(ch) - 64
    return n - 1                                     # indice 0 na linha A..AU


def is_err(v):
    """Valor de erro de celula como o pywin32 devolve (int muito negativo)."""
    return isinstance(v, int) and not isinstance(v, bool) and v < -2146820000


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not is_err(v)


def perto(a, b, tol=TOL):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def igual_valor(a, b):
    if num(a) and num(b):
        return perto(a, b)
    return (a if a is not None else '') == (b if b is not None else '')


# ------------------------------------------------------------------ regra do VBA, escrita de novo
def isnum_vba(v, vazio_conta=False):
    """mCEQ.NumOk (revisao 07/10/2026): numero de verdade = IsNumeric E nao vazio, nunca booleano nem erro.
    vazio_conta=True reproduz a regra ANTIGA (IsNumeric do VBA: celula vazia contava como numero 0) -- so para o
    R00 saber quais celulas mudam por essa correcao."""
    if is_err(v):
        return False
    if v is None:
        return vazio_conta
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
    if v is None or is_err(v) or (isinstance(v, str) and not v.strip()):
        return 0.0
    return float(str(v).replace(',', '.')) if isinstance(v, str) else float(v)


def txt(v):
    """CStr do VBA para o que a EQA_Base guarda (2025.0 -> '2025'); valor de erro = '' (mCEQ.TxtEQ)."""
    if v is None or is_err(v):
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def U(v):
    return txt(v).strip().upper()


def livre(v):
    return U(v) in ('', 'TODOS', 'TODAS')


def teto(ano):
    if ano is None or not isnum_vba(ano):
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
    if not isnum_vba(r[1]):
        return ''
    return f'{int(val(r[1]))}|{U(r[2])}'


def antes(a1, r1, a2, r2):
    if a1 != a2:
        return a1 < a2
    if r1.isdigit() and r2.isdigit() and len(r1) <= 9 and len(r2) <= 9:
        return int(r1) < int(r2)
    return r1 < r2


def existentes(rows, prov, analito=None):
    """Rodadas 'ANO|ROTULO' com linha utilizavel (Uso <> NAO, canonico, |bias| numerico de verdade), mais recente
    primeiro; com analito, so as DELE (ULTIMAS n e por analito)."""
    out = {}
    for r in rows:
        if not U(r[4]) or U(r[19]) == 'NAO':
            continue
        if analito is not None and U(r[4]) != U(analito):
            continue
        if not isnum_vba(r[16]):
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
    def __init__(self, rows, vazio_conta=False):
        self.rows = rows
        self.vazio = vazio_conta
        self._ult = {}

    def num(self, v):
        return isnum_vba(v, self.vazio)

    def ano(self, r):
        """Ano da linha que entra (None = nao entra). Regra antiga: IsNumeric(Empty) -> ano 0."""
        if not self.num(r[1]):
            return None
        return int(val(r[1])) if txt(r[1]).strip() else 0

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
                return chave(r) in self.ultimas(prov, analito, aref, n)
            return True
        if not livre(prov) and U(r[0]) != U(prov):
            return False
        if not livre(rodada) and U(r[2]) != U(rodada):
            return False
        return True

    def ultimas(self, prov, analito, aref, n):
        k_ = (U(prov), U(analito), aref, n)
        if k_ not in self._ult:
            lst = [k for k in existentes(self.rows, U(prov), analito) if int(k.split('|')[0]) <= aref]
            self._ult[k_] = set(lst[:n])
        return self._ult[k_]

    def ano_vigente(self, analito, ano, prov, rodada, col_vig):
        md = modo(rodada)
        aref = teto(ano)
        aref_vig = 32767 if md[0] == 2 else aref
        vig = None
        for r in self.rows:
            if self.casa(r, analito, prov, rodada, aref, md) and self.num(r[col_vig]):
                a = self.ano(r)
                if a is not None and a <= aref_vig and (vig is None or a > vig):
                    vig = a
        return vig

    def selecao(self, analito, ano, prov, rodada, col_vig, col):
        md = modo(rodada)
        aref = teto(ano)
        vig = self.ano_vigente(analito, ano, prov, rodada, col_vig)
        if vig is None:
            return None
        out = []
        for r in self.rows:
            if not (self.casa(r, analito, prov, rodada, aref, md) and self.num(r[col])):
                continue
            a = self.ano(r)
            if a is None:
                continue
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
            if self.num(r[16]):
                g[0] += val(r[16])
            if self.num(r[15]):
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
            if not self.num(r[10]) or not self.num(r[11]):
                sem += 1
            elif val(r[6]) < val(r[10]) or val(r[6]) > val(r[11]):
                fora += 1
            else:
                dentro += 1
        if fora:
            return f'NAO OK ({fora} fora dos limites)'
        if not dentro:
            return f'NAO AVALIADO ({sem} sem limite do provedor)'
        if sem:
            return f'OK ({dentro} dentro; {sem} sem limite)'
        return f'OK ({dentro} dentro dos limites)'

    def resumo(self, prov, ano, rodada):
        """('INVALIDA'|'PROVEDOR'|None, n_rodadas, n_linhas, extra) -- a contagem do K5, ANALITO A ANALITO com o
        recorte do BiasEQ de cada um (ano vigente dele nos modos de antes; ULTIMAS n dele)."""
        md = modo(rodada)
        extra = {'ncurto': 0, 'nanal': 0, 'anos_vig': []}
        if md[0] < 0:
            return 'INVALIDA', 0, 0, extra
        if md[0] >= 2 and livre(prov):
            return 'PROVEDOR', 0, 0, extra
        aref = teto(ano)
        rods, n, anos_vig = set(), 0, set()
        for an in sorted({U(r[4]) for r in self.rows if U(r[4])}):
            vig = self.ano_vigente(an, ano, prov, rodada, 16) if md[0] < 2 else None
            rods_an = set()
            for r in self.rows:
                if U(r[4]) != an:
                    continue
                if not (self.casa(r, an, prov, rodada, aref, md) and self.num(r[16])):
                    continue
                a = self.ano(r)
                if a is None:
                    continue
                if md[0] < 2 and a != vig:
                    continue
                if md[0] in (3, 4) and a > aref:
                    continue
                n += 1
                rods.add(chave(r))
                rods_an.add(chave(r))
            if rods_an:
                extra['nanal'] += 1
                if md[0] < 2:
                    anos_vig.add(vig)
                if md[0] == 4 and len(rods_an) < md[1]:
                    extra['ncurto'] += 1
        extra['anos_vig'] = sorted(anos_vig)
        return None, len(rods), n, extra


def k5_esperado(orc, prov, ano, rodada):
    """Trechos que o K5 tem de conter para o recorte (contagem + avisos por analito)."""
    tipo, nrod, nlin, ex = orc.resumo(prov, ano, rodada)
    partes = [f'{nrod} rodada(s), {nlin} linha(s)']
    md = modo(rodada)
    if md[0] == 4 and ex['ncurto']:
        partes.append(f'({ex["ncurto"]} de {ex["nanal"]} analito(s) com menos de {md[1]} com dado)')
    if 0 <= md[0] < 2 and len(ex['anos_vig']) > 1:
        partes.append('ano vigente por analito')
    return partes, (tipo, nrod, nlin, ex)


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


def forcar_todas(q, e):
    """R01/R02/R00 comparam com TODAS: a suite fixa P4 = TODAS na SUA copia (a producao pode estar em ULTIMAS 3,
    num conjunto... -- o recurso entregue). Nada e salvo. Devolve o valor que estava."""
    p4 = q.wb.Names('eqRodada').RefersToRange
    antes_ = p4.Value
    if U(antes_) not in ('TODAS', ''):
        q.ex.xl.EnableEvents = False
        try:
            p4.Value = 'TODAS'
        finally:
            q.ex.xl.EnableEvents = True
    return antes_


def foto(q):
    e = q.wb.Worksheets('Estatística')
    ult = ultima_linha(e)
    return {'produto': q.produto, 'ult': ult, 'linhas': [v for _, v in linhas_tabela(e, ult)],
            'formulas': [list(r) for r in e.Range(f'A14:AU{ult}').FormulaR1C1],
            'cab': list(e.Range('A13:AU13').Value[0]), 'K5': str(e.Range('K5').Text),
            'P4': str(q.wb.Names('eqRodada').RefersToRange.Value)}


def foto_de_arquivo(produto, original, pasta):
    """Foto do arquivo de ANTES: abre uma COPIA (o original nunca e aberto), fixa P4 = TODAS e recalcula tudo com
    o VBA dele."""
    tmp = os.path.join(pasta, 'r00_antes_' + os.path.basename(original))
    shutil.copy2(original, tmp)
    q = QA(produto, tmp)
    try:
        q.ex.xl.Calculation = -4105
        e = q.wb.Worksheets('Estatística')
        try:
            e.Unprotect(SENHA)
        except Exception:                            # noqa: BLE001
            pass
        p40 = forcar_todas(q, e)
        try:
            q.run('mIncerteza.RecalcularIncerteza')
        except Exception:                            # noqa: BLE001 -- arquivo anterior ao ADR-064
            pass
        q.ex.esperar()
        q.ex.xl.CalculateFull()
        q.ex.esperar()
        f = foto(q)
        f['P4_original'] = str(p40)
    finally:
        q.fechar()
    try:
        os.remove(tmp)
    except OSError:
        pass
    return f


# ------------------------------------------------------------------ testes
def r00_foto(q, e, ult, foto_antes, orc, prov_de, motivo_sem=None):
    if not foto_antes:
        if motivo_sem == 'pulado':
            print('   R00 pulado explicitamente (--sem-r00)', flush=True)
            return
        reg('R00 Regressão contra o arquivo ANTES desta entrega: NÃO executada -- falta o arquivo de antes '
            '(--antes <arquivo> ou QC_PROD_ORIGINAL; o entregar.py passa a produção original)', False,
            {'motivo': motivo_sem or 'sem --antes/--foto', 'como_pular': '--sem-r00 (só uso manual)'})
        return
    q.ex.xl.CalculateFull()
    q.ex.esperar()
    agora = [v for _, v in linhas_tabela(e, ult)]
    f_agora = [list(r) for r in e.Range(f'A14:AU{ult}').FormulaR1C1]
    cab_agora = list(e.Range('A13:AU13').Value[0])
    dif, mud = [], {c: 0 for c in COLS_SIGMA}
    fora = {}                                         # coluna -> motivo de nao comparar
    if foto_antes['ult'] != ult:
        dif.append(('ultima linha', foto_antes['ult'], ult))
    cab0 = foto_antes.get('cab')
    f0 = foto_antes.get('formulas')
    for c in COLS_IGUAIS:
        if cab0 is not None and str(cab0[ci(c)] or '') != str(cab_agora[ci(c)] or ''):
            fora[c] = f'cabeçalho mudou: {str(cab0[ci(c)])[:30]!r} -> {str(cab_agora[ci(c)])[:30]!r}'
    # celulas cujo valor muda pela regra "valor numerico de verdade" (|bias|/limite vazio): antiga x nova
    velho = Oraculo(orc.rows, vazio_conta=True)
    ano = q.wb.Names('eqAnoEP').RefersToRange.Value
    documentada = set()
    for i, v in enumerate(agora):
        an = v[0]
        if an in (None, '', 0):
            continue
        pv = prov_de(an)
        for c, met in (('G', 'ABS'), ('T', 'SIGNED'), ('AC', 'N'), ('AD', 'NRODADAS')):
            if not igual_valor(velho.bias(str(an), ano, pv, 'TODAS', met, legado=True),
                               orc.bias(str(an), ano, pv, 'TODAS', met, legado=True)):
                documentada.add((i, c))
                if c == 'G':
                    documentada.add((i, 'H'))
        if velho.sdi(str(an), ano, pv, 'TODAS') != orc.sdi(str(an), ano, pv, 'TODAS'):
            documentada.add((i, 'R'))
        if velho.limites(str(an), ano, pv, 'TODAS') != orc.limites(str(an), ano, pv, 'TODAS'):
            documentada.add((i, 'S'))
    formula_mudou = {}
    comparadas = {c: 0 for c in COLS_IGUAIS}
    for i, (a0, a1) in enumerate(zip(foto_antes['linhas'], agora)):
        for c in COLS_IGUAIS:
            if c in fora:
                continue
            if f0 is not None and i < len(f0) and str(f0[i][ci(c)]) != str(f_agora[i][ci(c)]):
                formula_mudou[c] = formula_mudou.get(c, 0) + 1
                continue
            if (i, c) in documentada:
                continue
            comparadas[c] += 1
            if not igual_valor(a0[ci(c)], a1[ci(c)]):
                dif.append((14 + i, c, a0[ci(c)], a1[ci(c)]))
        for c in COLS_SIGMA:
            if not igual_valor(a0[ci(c)], a1[ci(c)]):
                mud[c] += 1
    for c, n in formula_mudou.items():
        fora.setdefault(c, f'fórmula mudou nesta entrega em {n} linha(s) (ADR posterior ao arquivo de antes)')
    sem_ceq = [c for c in COLS_CEQ if comparadas[c] == 0]
    reg('R00 Regressão contra o arquivo ANTES desta entrega (P4 = TODAS nas duas cópias): G, H, R, S, T, AC, AD, AF..AU e '
        'o resto que não depende do Sigma idênticos célula a célula; só o que depende do Sigma (L, M..P, U..AA) muda; '
        'coluna cuja fórmula mudou nesta entrega e célula da correção "|bias|/limite vazio não é número" ficam fora, '
        'listadas',
        not dif and not sem_ceq,
        {'linhas': len(agora), 'diferencas': dif[:8], 'n_diferencas': len(dif),
         'celulas_comparadas': sum(comparadas.values()), 'fora_da_regressao': fora,
         'mudanca_documentada_numok': sorted({c for _, c in documentada}), 'n_documentada': len(documentada),
         'colunas_ceq_sem_comparacao': sem_ceq, 'mudaram_por_depender_do_sigma': mud,
         'P4_antes': foto_antes.get('P4_original', foto_antes.get('P4')),
         'K5_antes': foto_antes['K5'][:90], 'K5_agora': str(e.Range('K5').Text)[:110]})


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
    reg('R01 Regressão com TODAS (fixado na cópia de teste): G (|bias| ABS em 2 etapas), T, AC, AD, R e S = regra ANTIGA '
        'recalculada da EQA_Base (um ano vigente <= Ano EP, rodada pelo rótulo; número = IsNumeric E não vazio) em '
        'todas as linhas',
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
    rod0 = p4.Value                                  # TODAS (forcar_todas)
    base = colunas(e, ult, ('G', 'T', 'AC', 'AD'))
    ano = q.wb.Names('eqAnoEP').RefersToRange.Value
    prov = q.wb.Names('eqProvedor').RefersToRange.Value
    anos = sorted({int(k.split('|')[0]) for k in existentes(orc.rows, '' if bio else U(prov)) if int(k.split('|')[0]) <= teto(ano)})
    ev, ok = {}, U(rod0) in ('', 'TODAS')
    for texto in ('ACUMULADAS', 'ULTIMAS 3', 'ULTIMAS 6'):
        t0 = time.perf_counter()
        definir(q, e, 'P4', texto)
        dtm = round(time.perf_counter() - t0, 2)
        agora = colunas(e, ult, ('G', 'T', 'AC', 'AD'))
        dif = [(r, base[r], agora[r]) for r in base if not all(igual_valor(a, b) for a, b in zip(base[r], agora[r]))]
        k5 = str(e.Range('K5').Text)
        partes, _ = k5_esperado(orc, prov, ano, texto)
        k5_ok = all(p_ in k5 for p_ in partes) and not k5.startswith('#')
        ev[texto] = {'segundos': dtm, 'diferencas': dif[:4], 'K5': k5[:160], 'K5_esperado': partes}
        # so um ano de CEQ real: os modos novos tem de dar exatamente o que TODAS da
        if len(anos) == 1:
            ok = ok and not dif
        ok = ok and k5_ok
    definir(q, e, 'P4', rod0)
    reg('R02 Dado real (só um ano de CEQ): ACUMULADAS, ULTIMAS 3 e ULTIMAS 6 dão exatamente o G/T/AC/AD de TODAS; '
        'K5 conta rodadas e linhas pela mesma regra e avisa quantos analitos têm menos de n rodadas (3 ≠ 6)',
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
        f'itens fixos TODAS/ACUMULADAS/ULTIMAS 3/ULTIMAS 6 antes; nunca rodada de simulação nem sem alvo; validações de '
        f'N4/P4 (estilo Aviso) apontam para as listas{"" if bio else ", R4 sem vazio"}; K5 = ResumoFiltroEQ (sem TEXTO)',
        not dif and val_ok, {'divergencias': dif[:6], 'simulacao_existente': simul[:6], 'validacoes': ev,
                              'K5': k5[:120]})


def o3_esperado(iv, ev_):
    """Texto do rotulo O3: aviso quando um nivel com dados (E ou I numerico, I <> "-") ficou sem Sigma."""
    k = sum(1 for x in iv if num(x))
    n = sum(1 for x, e_ in zip(iv, ev_) if x != '-' and (num(e_) or num(x)))
    if 0 < k < n:
        return f'SIGMA EM {k} DE {n} N'
    return 'SIGMA DO PLANO'


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
    try:
        cvp_rng = q.wb.Names('Sigma_CVp').RefersToRange
    except Exception:                                 # noqa: BLE001
        cvp_rng = None
    b30 = p.Range('B3').Value
    dif, ev = [], {}

    def mostrar(an):
        q.ex.xl.EnableEvents = False
        try:
            p.Range('B3').Value = cad.index(an) + 1
        finally:
            q.ex.xl.EnableEvents = True
        q.run('mEstatistica.PainelMudou')
        recalc(q)

    def conferir(an, guardar=True):
        iv = [p.Range(f'I{6 + k}').Value for k in range(1, nlv + 1)]
        evv = [p.Range(f'E{6 + k}').Value for k in range(1, nlv + 1)]
        rows = [(chaves.get(f'{an}|{k}') or (0, None))[0] for k in range(1, nlv + 1)]
        lv = [e.Range(f'L{r}').Value if r else '' for r in rows]
        b1 = c.Range('B1').Value
        nums = [x for x in iv if num(x)]
        b1_esp = min(nums) if nums else ('-' if '-' in iv else 'SEM DADOS')
        o3 = str(p.Range('O3').Value or '')
        o3_esp = o3_esperado(iv, evv)
        cvp = [cvp_rng.Cells(k, 1).Value for k in range(1, nlv + 1)] if cvp_rng is not None else [None] * nlv
        fecha = []
        for k, r in enumerate(rows):
            if not r:
                continue
            kk, gg, ll = e.Range(f'K{r}').Value, e.Range(f'G{r}').Value, lv[k]
            if num(ll) and num(kk) and num(gg):
                fecha.append(num(cvp[k]) and perto((kk - abs(gg)) / cvp[k], ll, 1e-9))
        ok = (all(igual_valor(a, b) for a, b in zip(iv, lv)) and igual_valor(b1, b1_esp) and o3_esp in o3
              and cvp_rng is not None and all(fecha))
        if guardar:
            ev[an] = {'I': iv, 'L': lv, 'B1': b1, 'O3': o3, 'CVp': cvp, 'sigma_fecha_com_CVp': fecha}
        if not ok:
            dif.append((an, iv, lv, b1, b1_esp, o3, o3_esp, cvp, fecha))
        return iv, evv
    for an in cad[:3] + sem_etp[:1]:
        mostrar(an)
        conferir(an)
    # nivel com dados SEM Sigma: janela de 1 mes (lotes n >= 20 somem) -- procura um analito com nivel faltando
    meses0 = e.Range('L11').Value
    parcial = None
    definir(q, e, 'L11', 1)
    q.run('mIncerteza.RecalcularIncerteza')
    q.ex.esperar()
    for an in cad[:14]:
        mostrar(an)
        iv = [p.Range(f'I{6 + k}').Value for k in range(1, nlv + 1)]
        evv = [p.Range(f'E{6 + k}').Value for k in range(1, nlv + 1)]
        if o3_esperado(iv, evv) != 'SIGMA DO PLANO':
            parcial = an
            conferir(an)
            ev['parcial (1 mes): ' + an] = ev.pop(an)
            break
    definir(q, e, 'L11', meses0)
    q.run('mIncerteza.RecalcularIncerteza')
    q.ex.esperar()
    # caso deterministico (achado 7): um nivel COM dados perde o Sigma -- I desse nivel vazio nesta copia (a
    # formula volta logo depois) --> O3 avisa "SIGMA EM k DE N NIVEIS" e B1 continua o MIN dos que tem Sigma
    forcado = {}
    for an in cad[:6]:
        mostrar(an)
        iv = [p.Range(f'I{6 + k}').Value for k in range(1, nlv + 1)]
        evv = [p.Range(f'E{6 + k}').Value for k in range(1, nlv + 1)]
        alvos = [k for k in range(nlv) if num(iv[k]) and num(evv[k])]
        if len(alvos) < 2:
            continue
        k0 = alvos[-1]
        cel = p.Range(f'I{7 + k0}')
        f_ = cel.Formula
        q.ex.xl.EnableEvents = False
        try:
            cel.Value = ''
        finally:
            q.ex.xl.EnableEvents = True
        recalc(q)
        iv2 = list(iv)
        iv2[k0] = ''
        o3f_ = str(p.Range('O3').Value or '')
        b1f_ = c.Range('B1').Value
        esp_o3 = o3_esperado(iv2, evv)
        esp_b1 = min(x for x in iv2 if num(x))
        forcado = {'analito': an, 'nivel_sem_sigma': k0 + 1, 'O3': o3f_, 'O3_esperado': esp_o3, 'B1': b1f_,
                   'B1_esperado': esp_b1, 'ok': esp_o3 != 'SIGMA DO PLANO' and esp_o3 in o3f_ and igual_valor(b1f_, esp_b1)}
        q.ex.xl.EnableEvents = False
        try:
            cel.Formula = f_
        finally:
            q.ex.xl.EnableEvents = True
        recalc(q)
        forcado['O3_depois_de_restaurar'] = str(p.Range('O3').Value or '')
        forcado['ok'] = forcado['ok'] and forcado['O3_depois_de_restaurar'] == o3_esperado(iv, evv)
        break
    q.ex.xl.EnableEvents = False
    try:
        p.Range('B3').Value = b30
    finally:
        q.ex.xl.EnableEvents = True
    q.run('mEstatistica.PainelMudou')
    recalc(q)
    rot = str(p.Range('I6').Text)
    o3f = str(p.Range('O3').Formula)
    reg('R04 Painel: Sigma por nível (I7..I9) = Estatística!L do analito em tela (CV pooled de N meses; sem bias = '
        'vazio, nunca bias 0); Cfg_PlanoQC!B1 = MIN dos níveis; O3 avisa "SIGMA EM k DE N NÍVEIS" quando um nível com '
        'dados fica sem Sigma (o MIN o ignora); no bloco DESEMPENHO o CVp ao lado fecha a conta (ETp − |Bias|)/CVp = '
        'Sigma; "-" sem ETp (ADR-068)',
        not dif and 'CVp' in rot and 'SIGMA EM' in o3f and bool(forcado.get('ok')),
        {'analitos': ev, 'rotulo_I6': rot, 'caso_parcial_janela_1_mes': parcial, 'nivel_sem_sigma_forcado': forcado,
         'divergencias': dif[:4]})


# ------------------------------------------------------------------ CEQ sintetico
# (indice do analito: 0 = principal, 1 = segundo analito; provedor, ano, rodada, |bias| por amostra (None = sem alvo), uso)
SINT = [(0, 'CAP', 2024, 'A', [2.0, -4.0], 'SIM'), (0, 'CAP', 2024, 'B', [6.0, 8.0], 'SIM'),
        (0, 'CAP', 2026, 'A', [-1.0, 3.0], 'SIM'), (0, 'CAP', 2026, 'B', [5.0, 5.0], 'SIM'),
        (0, 'CAP', 2026, 'Z', [99.0, 99.0], 'NAO'),
        (0, 'CAP', 2026, 'C', [None, None], 'SIM'),            # rodada digitada SEM ALVO: |bias| vazio
        (0, 'CAP', 2027, 'A', [None], 'SIM'),                  # ano que so tem linha sem alvo
        (0, 'Controllab', 2025, '1', [40.0, 42.0], 'SIM'), (0, 'Controllab', 2026, '1', [60.0, 62.0], 'SIM'),
        # outra familia de survey no mesmo ano, so do 2o analito ('LN2' fica depois de 'A'/'B'/'C' no StrComp)
        (1, 'CAP', 2026, 'LN2-A', [1.0, 2.0], 'SIM'), (1, 'CAP', 2026, 'LN2-B', [3.0, 4.0], 'SIM'),
        (1, 'CAP', 2026, 'LN2-C', [5.0, 6.0], 'SIM')]


def linha_sint(prov, ano, rod, analito, i, x, uso):
    am = f'QA71-{ano}{rod}{i}'
    if x is None:                                    # resultado lancado, alvo ainda nao: O/P = "" -> vazio
        return (prov, ano, rod, 'QA ' + analito, analito, am, 100.0, None, None, None, 80.0, 120.0,
                '', 'NAO AVALIADO', '', None, None, '', 'QA ADR-071 sintetico (sem alvo)', uso,
                f'{prov}|{ano}|{rod}|{analito}|{am}')
    return (prov, ano, rod, 'QA ' + analito, analito, am, 100 + x, 100.0, 2.0, x / 2, 80.0, 120.0,
            'Acceptable', 'ACEITO', '', x, abs(x), '', 'QA ADR-071 sintetico', uso,
            f'{prov}|{ano}|{rod}|{analito}|{am}')


def escrever_base(q, linhas, rotulo):
    b = q.wb.Worksheets('EQA_Base')
    try:
        b.Unprotect(SENHA)
    except Exception:
        pass
    ult = max(b.Cells(b.Rows.Count, 1).End(-4162).Row, b.Cells(b.Rows.Count, 5).End(-4162).Row)
    b.Range(b.Cells(ult + 1, 1), b.Cells(ult + len(linhas), 21)).Value = tuple(linhas)
    b.Cells(1, 26).Value = f'consolidado: {rotulo} ' + time.strftime('%H:%M:%S') + f' {time.perf_counter():.6f}'
    return ult + 1


def injetar(q, analitos):
    """CEQ sintetico. ISOLADO: as linhas Controllab REAIS do analito principal saem desta copia (nada e salvo), para
    que as afirmacoes sobre o Controllab nao dependam de o laboratorio ainda nao ter consolidado nenhuma."""
    b = q.wb.Worksheets('EQA_Base')
    try:
        b.Unprotect(SENHA)
    except Exception:
        pass
    ult = max(b.Cells(b.Rows.Count, 1).End(-4162).Row, b.Cells(b.Rows.Count, 5).End(-4162).Row)
    removidas = 0
    if ult >= 2:
        vals = b.Range(b.Cells(2, 1), b.Cells(ult, 21)).Value
        for i, r in enumerate(vals, start=2):
            if U(r[0]) == 'CONTROLLAB' and U(r[4]) == U(analitos[0]):
                b.Range(b.Cells(i, 1), b.Cells(i, 21)).ClearContents()
                removidas += 1
    linhas = []
    for ia, prov, ano, rod, bs, uso in SINT:
        for i, x in enumerate(bs, 1):
            linhas.append(linha_sint(prov, ano, rod, analitos[ia], i, x, uso))
    escrever_base(q, linhas, 'QA ADR-071 sintetico')
    q.run('mCEQ.InvalidarEQ')
    return len(linhas), removidas


def r05_funcoes(q, orc, analito, analito2):
    casos = [('TODAS', 2026, 'CAP'), ('TODAS', 2025, 'CAP'), ('TODAS', 2027, 'CAP'), ('A', 2026, 'CAP'), ('A', 2025, 'CAP'),
             ('A', 2027, 'CAP'), ('C', 2026, 'CAP'),
             ('2024|A', 2026, 'CAP'), ('2024|A; 2026|A', 2026, 'CAP'), (' 2024|a ;2026|b; ', 2000, 'CAP'),
             ('2026|C', 2026, 'CAP'), ('ACUMULADAS', 2026, 'CAP'), ('ACUMULADAS', 2025, 'CAP'),
             ('ACUMULADAS', 2027, 'CAP'), ('ULTIMAS 64', 2027, 'CAP'), ('ULTIMAS 3', 2026, 'CAP'),
             ('ULTIMAS 6', 2026, 'CAP'), ('ULTIMAS 3', 2025, 'CAP'), ('últimas 2', 2026, 'CAP'),
             ('ULTIMAS 1', 2026, 'CAP'),
             ('ACUMULADAS', 2026, 'Controllab'), ('ULTIMAS 1', 2026, 'Controllab'), ('TODAS', 2026, 'Controllab'),
             ('2025|1; 2026|1', 2026, 'Controllab'),
             ('ACUMULADAS', 2026, ''), ('2024|A', 2026, ''), ('ULTIMAS 3', 2026, 'TODOS'), ('TODAS', 2026, ''),
             ('2025|', 2026, 'CAP'), ('C-A 2025; C-B 2025', 2026, 'CAP'), ('ULTIMAS 0', 2026, 'CAP'),
             ('ULTIMAS X', 2026, 'CAP'), ('2025|A|B', 2026, 'CAP'), ('25|A', 2026, 'CAP')]
    casos2 = [('ULTIMAS 3', 2026, 'CAP'), ('ULTIMAS 1', 2026, 'CAP'), ('ULTIMAS 6', 2026, 'CAP'), ('TODAS', 2026, 'CAP'),
              ('ACUMULADAS', 2026, 'CAP')]
    dif, ev = [], {}
    for an, lst in ((analito, casos), (analito2, casos2)):
        for rod, ano, prov in lst:
            linha = {}
            for met in ('ABS', 'SIGNED', 'N', 'NRODADAS'):
                got = q.run('mCEQ.BiasEQ', an, ano, met, prov, rod)
                esp = orc.bias(an, ano, prov, rod, met)
                linha[met] = got
                if not igual_valor(got, esp):
                    dif.append((an, rod, ano, prov, met, got, esp))
            ev[f'{an}/{rod!r}/{ano}/{prov or "-"}'] = linha
    b = lambda *a: q.run('mCEQ.BiasEQ', *a)   # noqa: E731
    # afirmacoes diretas (alem do oraculo)
    diretas = {
        'rotulo repetido em anos diferentes = 2 rodadas': b(analito, 2026, 'NRODADAS', 'CAP', '2024|A; 2026|A') == 2,
        'Controllab ACUMULADAS so com Controllab (4 amostras sinteticas; o real foi isolado)':
            b(analito, 2026, 'N', 'Controllab', 'ACUMULADAS') == 4,
        'CAP nunca leva o Controllab (bias 40-62 ausente)': all(
            (not num(b(analito, 2026, 'ABS', 'CAP', t))) or b(analito, 2026, 'ABS', 'CAP', t) < 30
            for t in ('ACUMULADAS', 'ULTIMAS 6', '2025|1; 2026|1')),
        'Uso=NAO (2026|Z, bias 99) nunca entra': b(analito, 2026, 'ABS', 'CAP', '2026|Z') == SEM_EP,
        'provedor vazio nos modos novos = SEM EP (D1)': b(analito, 2026, 'ABS', '', 'ACUMULADAS') == SEM_EP,
        'texto invalido = SEM EP (nunca #VALOR!)': b(analito, 2026, 'ABS', 'CAP', '2025|') == SEM_EP,
        'conjunto atravessando anos ignora o teto': num(b(analito, 2000, 'ABS', 'CAP', '2024|A; 2026|B')),
        # revisao 07/10/2026
        'ULTIMAS 3 e por analito: outra familia (LN2-x 2026) nao tira o analito principal': (
            b(analito, 2026, 'NRODADAS', 'CAP', 'ULTIMAS 3') == 3 and b(analito2, 2026, 'NRODADAS', 'CAP', 'ULTIMAS 3') == 3
            and num(b(analito, 2026, 'ABS', 'CAP', 'ULTIMAS 3'))),
        'rodada sem alvo (|bias| vazio) nao entra em TODAS (2026: so A e B)': b(analito, 2026, 'NRODADAS', 'CAP', 'TODAS') == 2,
        'rodada so sem alvo = SEM EP, nunca 0': b(analito, 2026, 'ABS', 'CAP', '2026|C') == SEM_EP,
        'ano so com linha sem alvo nao vira vigente (Ano EP 2027 -> 2026)': b(analito, 2027, 'ANO', 'CAP', 'TODAS') == 2026,
        'ACUMULADAS = ULTIMAS 64 com rodada sem alvo na base': igual_valor(
            b(analito, 2027, 'ABS', 'CAP', 'ACUMULADAS'), b(analito, 2027, 'ABS', 'CAP', 'ULTIMAS 64')),
    }
    # SDI e limites seguem a mesma selecao (ano por item / acumulado)
    for rod in ('2024|A; 2026|A', 'ACUMULADAS', 'ULTIMAS 3', 'TODAS'):
        for an in (analito, analito2):
            got_n = q.run('mCEQ.SDIeq', an, 2026, 'N', 'CAP', rod)
            esp = orc.sdi(an, 2026, 'CAP', rod)
            if (esp is None and got_n != SEM_EP) or (esp is not None and got_n != esp[1]):
                dif.append(('SDIeq N', an, rod, got_n, esp))
            got_s = q.run('mCEQ.StatusLimitesEQ', an, 2026, 'CAP', rod)
            if got_s != orc.limites(an, 2026, 'CAP', rod):
                dif.append(('StatusLimitesEQ', an, rod, got_s, orc.limites(an, 2026, 'CAP', rod)))
    # D3: a incerteza nao segue a selecao nova
    d3 = []
    for rod in ('ULTIMAS 3', 'ACUMULADAS', '2024|A; 2026|B', '2025|'):
        for met in ('N', 'NRODADAS', 'MEDIA', 'SITUACAO', 'UBIAS'):
            a_ = q.run('mCEQ.ViesEQ', analito, 2026, met, 'CAP', rod)
            b_ = q.run('mCEQ.ViesEQ', analito, 2026, met, 'CAP', 'TODAS')
            if not igual_valor(a_, b_):
                d3.append((rod, met, a_, b_))
    reg('R05 CEQ sintético (CAP 2024/2025/2026/2027 com rótulos repetidos, rodada Uso=NAO, rodada SEM ALVO, 2º analito '
        'com a família LN2-x no mesmo ano, Controllab 2025/2026 isolado do real): BiasEQ ABS/SIGNED/N/NRODADAS = oráculo '
        'em 1 rodada, conjunto (atravessando anos), ACUMULADAS, ULTIMAS n POR ANALITO, TODAS e rótulo; |bias| vazio '
        'nunca vale 0 nem torna o ano vigente; NRODADAS conta ano|rodada; CAP e Controllab nunca misturados; provedor '
        'vazio e texto inválido = SEM EP; SDI e limites pela mesma seleção; incerteza (ViesEQ) igual a TODAS (D3)',
        not dif and all(diretas.values()) and not d3,
        {'casos': len(casos) + len(casos2), 'divergencias': dif[:8], 'diretas': diretas, 'D3_divergencias': d3[:6],
         'exemplos': dict(list(ev.items())[:6])})


def r06_resumo(q, orc):
    casos = [('CAP', 2026, 'ULTIMAS 3', 'ultimas 3 rodada(s) de cada analito'), ('CAP', 2026, 'ULTIMAS 6', 'ultimas 6'),
             ('CAP', 2025, 'ACUMULADAS', 'acumuladas ate 2025'), ('CAP', 2026, 'ACUMULADAS', 'acumuladas ate 2026'),
             ('CAP', 2027, 'ACUMULADAS', 'acumuladas ate 2027'),
             ('CAP', 2026, '2024|A; 2024|Q', 'SEM DADO: 2024|Q'), ('CAP', 2026, '2026|C', 'SEM DADO: 2026|C'),
             ('Controllab', 2026, 'ULTIMAS 1', 'ultimas 1'),
             ('CAP', 2026, 'TODAS', 'media de todas as rodadas'), ('CAP', 2027, 'TODAS', 'media de todas as rodadas'),
             ('CAP', 2026, '2025|', 'INVALIDA'),
             ('', 2026, 'ACUMULADAS', 'provedor nao definido'), ('CAP', 2026, 'C-A 2025; C-B 2025', 'item sem ano')]
    dif, ev = [], {}
    for prov, ano, rod, palavra in casos:
        t = str(q.run('mCEQ.ResumoFiltroEQ', prov, ano, rod))
        partes, (tipo, nrod, nlin, ex) = k5_esperado(orc, prov, ano, rod)
        partes = [p_ + (' analiticas' if p_.endswith('linha(s)') else '') for p_ in partes]
        ok = palavra in t and all(p_ in t for p_ in partes)
        ev[f'{prov or "-"}/{ano}/{rod}'] = t[:170]
        if not ok:
            dif.append((prov, ano, rod, t[:170], partes))
    # o caso que motivou a revisao: anos vigentes diferentes por analito (2026 so para os sinteticos)
    _, (_, _, _, ex26) = k5_esperado(orc, 'CAP', 2026, 'TODAS')
    reg('R06 K5 (ResumoFiltroEQ) descreve o modo e conta rodadas e linhas analíticas ANALITO A ANALITO pela regra do '
        'cálculo (ano vigente de cada analito; ULTIMAS n de cada analito, dizendo quantos têm menos de n); |bias| vazio '
        'não conta; texto inválido diz INVALIDA, item do conjunto sem dado aparece como SEM DADO, provedor vazio nos '
        'modos novos é avisado',
        not dif and len(ex26['anos_vig']) > 1, {'textos': ev, 'divergencias': dif[:5], 'anos_vigentes_TODAS_2026': ex26['anos_vig']})


def r07_listas_sint(q, e, orc, bio):
    q.run('mDados.AtualizarListasAno')
    r03_listas(q, e, orc, bio, 'R07', 'CEQ sintético (CAP 2024/2026, LN2-x, Controllab, rodada Uso=NAO e sem alvo)')


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
    partes, (_, nrod, nlin, _) = k5_esperado(orc, prov_k5, 2026, 'ULTIMAS 3')
    k5_ok = 'ultimas 3' in k5 and all(p_ in k5 for p_ in partes)
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
        'igual ao oráculo (provedor de cada linha, ULTIMAS por analito), K5 conta igual, a incerteza AF..AU não muda '
        '(D3); P4 inválido = SEM EP em todas as linhas com aviso em K5'
        + ('' if bio else '; provedor vazio + ACUMULADAS = SEM EP com aviso (D1)'),
        not dif and k5_ok and not d3 and inv_ok and prov_ok,
        {'segundos_troca_P4': seg, 'divergencias': dif[:6], 'K5': k5[:160], 'esperado_K5': partes,
         'incerteza_mudou': d3[:3], 'K5_invalido': k5_inv[:120], 'K5_sem_provedor': k5_prov[:140]})


def linha_do(e, ult, analito, nivel=1):
    for r, v in linhas_tabela(e, ult):
        if str(v[0] or '').strip().upper() == str(analito).strip().upper() and int(v[1] or 0) == nivel:
            return r
    return None


def r09_consolidar(q, e, ult, analito, prov_de):
    """Achado 6: o fim do AtualizarEQABase (MarcarCEQSujo + Calculate do RecalcularIncerteza) atualiza G e K5."""
    r = linha_do(e, ult, analito)
    ano = q.wb.Names('eqAnoEP').RefersToRange.Value
    rod = q.wb.Names('eqRodada').RefersToRange.Value
    prov_k5 = q.wb.Names('eqProvedor').RefersToRange.Value
    pv = prov_de(analito)
    g0, k50 = e.Range(f'G{r}').Value, str(e.Range('K5').Text)
    a = int(teto(ano)) if teto(ano) != 32767 else 2026
    escrever_base(q, [linha_sint(pv or 'CAP', a, 'QA-R09', analito, i, 50.0, 'SIM') for i in (1, 2)], 'QA R09')
    recalc(q)
    g_parado = e.Range(f'G{r}').Value
    q.run('mCEQ.MarcarCEQSujo')
    recalc(q)
    orc = Oraculo(ler_base(q))
    g1, k51 = e.Range(f'G{r}').Value, str(e.Range('K5').Text)
    esp = orc.bias(str(analito), ano, pv, rod, 'ABS')
    partes, _ = k5_esperado(orc, prov_k5, ano, rod)
    try:
        cm = q.wb.VBProject.VBComponents('mEQA').CodeModule
        codigo = cm.Lines(1, cm.CountOfLines)
        chama = 'mCEQ.MarcarCEQSujo' in codigo and codigo.index('mCEQ.MarcarCEQSujo') < codigo.index('mIncerteza.RecalcularIncerteza', codigo.index('mCEQ.MarcarCEQSujo'))
    except Exception as ex:                          # noqa: BLE001
        chama = f'nao lido: {ex}'
    reg('R09 Depois de mudar a EQA_Base, o fim do AtualizarEQABase (mCEQ.MarcarCEQSujo antes do Calculate do '
        'RecalcularIncerteza) atualiza G e K5 sem mexer em N4/P4/R4 (sem ele, G ficava com o conjunto antigo)',
        chama is True and igual_valor(g1, esp) and all(p_ in k51 for p_ in partes) and k51 != k50,
        {'analito': analito, 'G_antes': g0, 'G_so_com_Calculate': g_parado, 'G_depois': g1, 'G_esperado': esp,
         'K5_antes': k50[:120], 'K5_depois': k51[:120], 'K5_esperado': partes, 'AtualizarEQABase_chama': chama})


def r10_muitas(q, analito):
    """Achado 3: mais de 64 rodadas (provedor sintetico QA64, 70 rodadas num ano) -- nenhuma truncada."""
    linhas = [linha_sint('QA64', 2020, str(k), analito, 1, k / 10.0, 'SIM') for k in range(1, 71)]
    escrever_base(q, linhas, 'QA R10')
    q.run('mCEQ.InvalidarEQ')
    orc = Oraculo(ler_base(q))
    got = {m: q.run('mCEQ.BiasEQ', analito, 2026, m, 'QA64', 'ACUMULADAS') for m in ('ABS', 'N', 'NRODADAS')}
    esp = {m: orc.bias(analito, 2026, 'QA64', 'ACUMULADAS', m) for m in got}
    conj = '; '.join(f'2020|{k}' for k in range(1, 71))
    got_c = q.run('mCEQ.BiasEQ', analito, 2026, 'NRODADAS', 'QA64', conj)
    k5 = str(q.run('mCEQ.ResumoFiltroEQ', 'QA64', 2026, 'ACUMULADAS'))
    reg('R10 Mais de 64 rodadas (ACUMULADAS e conjunto de 70 itens): todas entram -- NRODADAS = 70, média = oráculo; '
        'nada truncado em silêncio (antes: MAX_ROD = 64 pela ordem física da EQA_Base)',
        got['NRODADAS'] == 70 and got_c == 70 and all(igual_valor(got[m], esp[m]) for m in got) and '70 rodada(s)' in k5,
        {'obtido': got, 'esperado': esp, 'conjunto_70': got_c, 'K5': k5[:140]})


def r11_erro(q, analito, analito2):
    """Achado 4: celula com valor de erro na EQA_Base nao derruba o cache."""
    b = q.wb.Worksheets('EQA_Base')
    ult = max(b.Cells(b.Rows.Count, 1).End(-4162).Row, b.Cells(b.Rows.Count, 5).End(-4162).Row)
    vals = b.Range(b.Cells(2, 1), b.Cells(ult, 21)).Value
    alvo = [i for i, r in enumerate(vals, start=2) if U(r[4]) == U(analito) and U(r[0]) == 'CAP'
            and str(r[18] or '').startswith('QA ADR-071 sintetico') and num(r[16])]
    b.Cells(alvo[0], 17).Formula = '=1/0'            # |bias| = #DIV/0! (colado sobre a formula da aba de digitacao)
    b.Cells(alvo[1], 3).Formula = '=NA()'           # rodada = #N/D
    b.Cells(1, 26).Value = 'consolidado: QA R11 ' + time.strftime('%H:%M:%S')
    q.run('mCEQ.InvalidarEQ')
    orc = Oraculo(ler_base(q))
    dif = []
    for an in (analito, analito2):
        for rod in ('TODAS', 'ACUMULADAS', 'ULTIMAS 3'):
            for met in ('ABS', 'N', 'NRODADAS'):
                got = q.run('mCEQ.BiasEQ', an, 2026, met, 'CAP', rod)
                esp = orc.bias(an, 2026, 'CAP', rod, met)
                if not igual_valor(got, esp):
                    dif.append((an, rod, met, got, esp))
    k5 = str(q.run('mCEQ.ResumoFiltroEQ', 'CAP', 2026, 'ACUMULADAS'))
    try:
        q.run('mDados.AtualizarListasAno')
        listas = 'ok'
    except Exception as ex:                          # noqa: BLE001
        listas = f'erro: {ex}'[:160]
    reg('R11 Célula com valor de erro na EQA_Base (#DIV/0! no |bias|, #N/D na rodada): o cache monta, as demais linhas '
        'calculam igual ao oráculo (erro = não número; rodada com erro = rótulo vazio), K5 não diz "ilegível" e as '
        'listas atualizam sem erro (antes: CStr dava erro 13 e todo o CEQ ficava SEM EP)',
        not dif and 'ilegivel' not in k5 and 'indisponivel' not in k5 and listas == 'ok',
        {'linhas_com_erro': alvo[:2], 'divergencias': dif[:6], 'K5': k5[:140], 'AtualizarListasAno': listas})


def executar(produto, caminho, saida, foto_antes=None, motivo_sem=None):
    q = QA(produto, caminho)
    bio = produto.startswith('Bio')
    try:
        q.ex.xl.Calculation = -4105
        e = q.wb.Worksheets('Estatística')
        e.Unprotect(SENHA)
        p4_original = forcar_todas(q, e)
        if U(p4_original) not in ('TODAS', ''):
            print(f'   P4 da copia era {p4_original!r}: fixado em TODAS (so nesta copia)', flush=True)
        q.run('mIncerteza.RecalcularIncerteza')
        q.ex.esperar()
        ult = ultima_linha(e)
        orc = Oraculo(ler_base(q))
        prov_de = provedores_linha(q, bio)
        r00_foto(q, e, ult, foto_antes, orc, prov_de, motivo_sem)
        r01_legado(q, e, ult, orc, prov_de)
        r02_real(q, e, ult, orc, bio)
        r03_listas(q, e, orc, bio, 'R03', 'Dado real')
        r04_painel(q, e, ult, bio)
        # ---- CEQ sintetico (so nesta copia; nada e salvo)
        cand = [v[0] for _, v in linhas_tabela(e, ult) if v[0] not in (None, '', 0) and num(v[ci('G')])]
        analito = 'WBC' if (not bio and 'WBC' in cand) else cand[0]
        analito2 = next(a for a in cand if U(a) != U(analito))
        n, removidas = injetar(q, [str(analito), str(analito2)])
        orc = Oraculo(ler_base(q))
        print(f'   CEQ sintetico: {n} linhas injetadas para {analito}/{analito2}; {removidas} linha(s) Controllab real '
              f'de {analito} retirada(s) da copia', flush=True)
        r05_funcoes(q, orc, str(analito), str(analito2))
        r06_resumo(q, orc)
        r07_listas_sint(q, e, orc, bio)
        r08_aba(q, e, ult, orc, bio, prov_de)
        r09_consolidar(q, e, ult, str(analito), prov_de)
        r10_muitas(q, str(analito))
        r11_erro(q, str(analito), str(analito2))
    finally:
        q.fechar()
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [x for x in RES if x['resultado'] == 'FAIL']
    print(f'\n=== {produto} rodadas/CV: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


def gerar_foto(produto, caminho, destino):
    f = foto_de_arquivo(produto, caminho, os.path.dirname(os.path.abspath(destino)))
    with open(destino, 'w', encoding='utf-8') as fp:
        json.dump(f, fp, ensure_ascii=False, default=str)
    print(f'foto gravada: {destino} ({len(f["linhas"])} linhas)', flush=True)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--gerar-foto':
        gerar_foto(a[1], os.path.abspath(a[2]), a[3])
        sys.exit(0)
    fa, motivo = None, None
    if '--foto' in a:
        i = a.index('--foto')
        with open(a[i + 1], encoding='utf-8') as fp:
            fa = json.load(fp)
        del a[i:i + 2]
    antes_arq = os.environ.get('QC_PROD_ORIGINAL') or None
    if '--antes' in a:
        i = a.index('--antes')
        antes_arq = a[i + 1]
        del a[i:i + 2]
    if '--sem-r00' in a:
        a.remove('--sem-r00')
        motivo = 'pulado'
    elif fa is None and antes_arq:
        if os.path.exists(antes_arq):
            try:
                fa = foto_de_arquivo(a[0], os.path.abspath(antes_arq), os.path.dirname(os.path.abspath(a[1])))
            except Exception as ex:                  # noqa: BLE001
                motivo = f'foto do arquivo de antes falhou: {type(ex).__name__}: {ex}'[:300]
        else:
            motivo = f'arquivo de antes nao existe: {antes_arq}'
    falhas = executar(a[0], os.path.abspath(a[1]), a[2] if len(a) > 2 else None, fa, motivo)
    sys.exit(1 if falhas else 0)
