# -*- coding: utf-8 -*-
"""QA da INCERTEZA DE MEDICAO (ADR-064) -- validacao INDEPENDENTE numa COPIA instalada.

Uso: python qa_incerteza.py <Bioquimica|Hematologia> <copia.xlsm> [saida.json]

Tudo o que a aba Estatistica e o Painel mostram e RECALCULADO aqui em Python a partir dos dados
crus (tblCQ_Final, EQA_Base e a digitacao do CAP) -- outra implementacao, mesma regra escrita no
ADR-064 -- e comparado celula a celula. Nada e salvo no arquivo testado.
"""
import json
import math
import os
import sys
import time
from collections import defaultdict

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES  # noqa: E402

TOL = 1e-6
N_LOTE_MIN, DIAS_MIN, DIAS_VAL, GL_MIN, GL_VAL = 20, 90, 180, 30, 100


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, int) and v < -2146820000)


def perto(a, b, tol=TOL):
    if not (num(a) and num(b)):
        return False
    return abs(a - b) <= tol * max(1.0, abs(b))


def dia(v):
    if hasattr(v, 'toordinal'):
        return float(v.toordinal() - 693594)           # serial do Excel (1900)
    return float(int(v))


def coluna(lo, nome):
    return [r[0] for r in lo.ListColumns(nome).DataBodyRange.Value]


def esperado_ciq(q, eq_filtro, ini, fim, exc):
    """u(Rw) por analito|nivel, lotes separados (n >= 20) e agrupados por gl."""
    lo = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
    cols = {c: coluna(lo, c) for c in ('DATA_HORA', 'NIVEL', 'LOTE', 'ANALITO', 'RESULTADO', 'PARTICIPA_ESTATISTICA',
                                        'EQUIPAMENTO')}
    agg = defaultdict(lambda: defaultdict(list))
    cont = defaultdict(lambda: [0, 0])
    for i in range(len(cols['ANALITO'])):
        an = str(cols['ANALITO'][i] or '').strip()
        if not an:
            continue
        if eq_filtro and str(cols['EQUIPAMENTO'][i] or '').strip().upper() != eq_filtro:
            continue
        dh = cols['DATA_HORA'][i]
        if dh is None or not hasattr(dh, 'toordinal'):
            continue
        d = dia(dh)
        if d < ini or d > fim:
            continue
        if any(a <= d <= b for a, b in exc):
            continue
        k = (an.upper(), str(cols['NIVEL'][i]).strip().replace('.0', ''))
        cont[k][0] += 1
        if str(cols['PARTICIPA_ESTATISTICA'][i] or '').strip().upper() != 'SIM':
            cont[k][1] += 1
            continue
        v = cols['RESULTADO'][i]
        if not num(v):
            continue
        lote = str(cols['LOTE'][i] or '').strip()
        if not lote:
            continue
        agg[k][lote].append((float(v), d))
    out = {}
    for k, lotes in agg.items():
        s_cv = s_s = gl = ntot = soma = 0.0
        nl = fora = 0
        dmin, dmax, cvs = None, None, []
        for lt, xs in lotes.items():
            n = len(xs)
            if n < N_LOTE_MIN:
                fora += 1
                continue
            vals = [x for x, _ in xs]
            m = sum(vals) / n
            if m == 0:                                   # lote todo em zero: CV indefinido, fica fora
                fora += 1
                continue
            sd = math.sqrt(max(0.0, sum((x - m) ** 2 for x in vals) / (n - 1)))
            cv = sd / abs(m) * 100
            s_cv += (n - 1) * cv * cv
            cvs.append(cv)
            s_s += (n - 1) * sd * sd
            gl += n - 1
            ntot += n
            soma += sum(vals)
            nl += 1
            a, b = min(d for _, d in xs), max(d for _, d in xs)
            dmin = a if dmin is None else min(dmin, a)
            dmax = b if dmax is None else max(dmax, b)
        dias = (dmax - dmin) if nl else None
        val = ('INSUFICIENTE' if nl == 0 or dias < DIAS_MIN or gl < GL_MIN else
               'PROVISORIA' if dias < DIAS_VAL or gl < GL_VAL else 'VALIDA')
        out[k] = {'cv': math.sqrt(s_cv / gl) if gl else None, 'gl': gl, 'lotes': nl, 'fora': fora, 'dias': dias,
                  'media': soma / ntot if ntot else None, 'validade': val,
                  'pexcl': cont[k][1] / cont[k][0] * 100 if cont[k][0] else None}
    return out


def esperado_ceq(q, analito, ano_ref, provedor, rodada):
    """Media com sinal, EP, amostras, rodadas, Nordtest -- direto da EQA_Base e da digitacao do CAP."""
    b = q.wb.Worksheets('EQA_Base')
    ult = b.Cells(b.Rows.Count, 5).End(-4162).Row
    d = b.Range(b.Cells(2, 1), b.Cells(max(2, ult), 21)).Value
    def txt(x):                                   # = mEQA.Txt: Trim(CStr(v)) -- 2025.0 vira "2025"
        if isinstance(x, float) and x.is_integer():
            return str(int(x))
        return '' if x is None else str(x).strip()

    nlabs = {}
    for aba, prov in (('EQA.CAP_Dados', 'CAP'), ('EQA.Controllab_Dados', 'Controllab')):
        w = q.wb.Worksheets(aba)
        u = w.Cells(w.Rows.Count, 4).End(-4162).Row
        if u < 2:
            continue
        for r in w.Range(w.Cells(2, 1), w.Cells(u, 18)).Value:
            if r[3] in (None, ''):
                continue
            ch = '|'.join([prov, txt(r[2]), txt(r[1]), txt(r[3]), txt(r[4])])     # provedor|ano|rodada|analito|amostra
            if num(r[12]) and r[12] > 0:
                nlabs[ch.upper()] = float(r[12])

    def casa(r):
        if str(r[4] or '').strip().upper() != analito.upper():
            return False
        if str(r[19] or '').strip().upper() == 'NAO':
            return False
        if provedor and str(provedor).upper() not in ('', 'TODOS', 'TODAS') and str(r[0]).strip().upper() != str(provedor).upper():
            return False
        if rodada and str(rodada).upper() not in ('', 'TODOS', 'TODAS') and str(r[2]).strip().upper() != str(rodada).upper():
            return False
        return True
    anos = [int(r[1]) for r in d if casa(r) and num(r[1]) and num(r[15]) and int(r[1]) <= int(ano_ref)]
    if not anos:
        return None
    av = max(anos)
    for anomin in (av, av - 1):                   # ano vigente; sem o criterio, inclui o ano anterior
        bs, rods, ucr, vistos = [], set(), [], set()
        for r in d:
            if not (casa(r) and num(r[1]) and anomin <= int(r[1]) <= av and num(r[15])):
                continue
            if str(r[13] or '').strip().upper() == 'NAO AVALIADO':
                continue
            if not (num(r[6]) and num(r[7])) or r[6] == 0 or r[7] == 0:
                continue
            ch = str(r[20] or '').strip().upper()   # coluna Chave da EQA_Base (a mesma que o VBA usa)
            if ch:
                if ch in vistos:
                    continue
                vistos.add(ch)
            bs.append(float(r[15]))
            rods.add((int(r[1]), str(r[2]).strip().upper()))
            if num(r[8]) and r[8] >= 0 and ch in nlabs:
                ucr.append(100 * float(r[8]) / abs(float(r[7])) / math.sqrt(nlabs[ch]))
        if len(bs) >= 6 and len(rods) >= 2:
            break
    m = len(bs)
    if m == 0:
        return None
    media = sum(bs) / m
    dp = math.sqrt(max(0.0, sum((x - media) ** 2 for x in bs) / (m - 1))) if m >= 2 else None
    ep = dp / math.sqrt(m) if dp is not None else None
    rms = math.sqrt(sum(x * x for x in bs) / m)
    ok = m >= 6 and len(rods) >= 2
    uc = sum(ucr) / len(ucr) if ucr else None
    return {'media': media, 'ep': ep, 'n': m, 'rodadas': len(rods), 'ok': ok,
            'ubias': math.sqrt(rms ** 2 + uc ** 2) if (ok and uc is not None) else None,
            'detectavel': (abs(media) > 2 * ep) if (ok and ep is not None) else None}


def executar(produto, caminho, saida):
    q = QA(produto, caminho)
    e = q.wb.Worksheets('Estatística')
    a = q.wb.Worksheets('Analitos')
    p = q.wb.Worksheets('Painel')
    bio = produto.startswith('Bio')
    try:
        q.ex.xl.Calculation = -4105
        t0 = time.perf_counter()
        q.run('mIncerteza.RecalcularIncerteza')
        q.ex.esperar()
        print(f'   recalculo inicial {time.perf_counter() - t0:.1f}s', flush=True)
        for tentativa in range(5):
            # logo depois do recalculo o COM ja devolveu "Membro nao encontrado" uma vez (transitorio, 04/10/2026)
            try:
                e = q.wb.Worksheets('Estatística')
                ini, fim = dia(e.Range('AG11').Value), dia(e.Range('AI11').Value)
                break
            except Exception:
                if tentativa == 4:
                    raise
                time.sleep(2)
                q.ex.esperar()
        exc = []
        if bio:
            for r in e.Range('Estat_Exclusoes').Value:
                if r[0] and r[1]:
                    x, y = dia(r[0]), dia(r[1])
                    if y >= x:
                        exc.append((x, y))
        try:
            eqf = str(q.wb.Names('selEquipamento').RefersToRange.Value or '').strip().upper()
        except Exception:
            eqf = ''
        lo_ = q.wb.Worksheets('Principal - Resultados').ListObjects('tblCQ_Final')
        ultima = max(dia(v) for v, pt in zip(coluna(lo_, 'DATA'), coluna(lo_, 'PARTICIPA_ESTATISTICA'))
                     if hasattr(v, 'toordinal') and str(pt or '').strip().upper() == 'SIM')
        reg('I00 Janela da incerteza = 12 meses móveis terminando no último resultado que PARTICIPA da estatística '
            '(uma data futura digitada por engano num resultado manual não desloca a janela)',
            fim == ultima and ini == dia_menos_12m(fim), {'janela': [e.Range('AG11').Text, e.Range('AI11').Text],
                                                          'ultimo_resultado': ultima})

        esp = esperado_ciq(q, eqf, ini, fim, exc)
        linhas = []
        for r in range(14, 134):                     # so a tabela analito|nivel: a chave AB existe
            an, ch = e.Range(f'A{r}').Value, str(e.Range(f'AB{r}').Value or '')
            if not an or an == 0 or '|' not in ch:
                continue
            linhas.append((r, str(an), str(e.Range(f'B{r}').Value).replace('.0', '')))
        dif, cont, cont_u = [], {'VALIDA': 0, 'PROVISORIA': 0, 'INSUFICIENTE': 0}, [0, 0]
        for r, an, nv in linhas:
            x = esp.get((an.upper(), nv))
            vals = {c: e.Range(f'{c}{r}').Value for c in ('AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AO', 'AP')}
            val = x['validade'] if x else 'INSUFICIENTE'
            cont[val] += 1
            if val == 'INSUFICIENTE':
                if vals['AF'] not in ('', None) or vals['AO'] != 'Insuficiente':
                    dif.append((r, an, nv, 'insuficiente exibido', vals['AF'], vals['AO']))
                continue
            if not perto(vals['AF'], x['cv']):
                dif.append((r, an, nv, 'u(Rw)', vals['AF'], x['cv'], val, x['gl'], x['dias']))
                continue
            if vals['AG'] != x['gl'] or vals['AH'] != x['lotes'] or vals['AI'] != x['dias']:
                dif.append((r, an, nv, 'gl/lotes/dias', (vals['AG'], vals['AH'], vals['AI']), (x['gl'], x['lotes'], x['dias'])))
            sit = str(vals['AP'])
            if num(vals['AJ']):                       # Nordtest: uc = raiz(u(Rw)^2 + u(bias)^2)
                cont_u[0] += 1
                if not (perto(vals['AK'], math.sqrt(vals['AF'] ** 2 + vals['AJ'] ** 2)) and perto(vals['AL'], 2 * vals['AK'])):
                    dif.append((r, an, nv, 'uc/U', vals['AK'], vals['AL']))
                elif not perto(vals['AM'], vals['AL'] * x['media'] / 100):
                    dif.append((r, an, nv, 'U unidade', vals['AM'], vals['AL'] * x['media'] / 100))
                if (val == 'PROVISORIA') != sit.startswith('PROVISÓRIA'):
                    dif.append((r, an, nv, 'situacao', sit, val))
            else:                                     # sem u(bias): U nao estimada, motivo na situacao
                cont_u[1] += 1
                if vals['AK'] not in ('', None) or vals['AL'] not in ('', None) or not sit.startswith('U não estimada: '):
                    dif.append((r, an, nv, 'sem u(bias)', vals['AK'], vals['AL'], sit))
        reg(f'I01 u(Rw), gl, lotes, dias conferem com o recálculo INDEPENDENTE (lotes n ≥ 20 agrupados por gl) em todas as '
            f'{len(linhas)} linhas; uc = raiz(u(Rw)² + u(bias)²), U = 2 uc e U na unidade quando há u(bias) do CEQ; '
            f'sem u(bias), U não estimada e o motivo na situação',
            not dif, {'linhas': len(linhas), 'validade': cont, 'com_U_sem_U': cont_u, 'divergencias': dif[:8]})

        # meta e classe
        cvi_col = 'O' if bio else 'M'
        dif = []
        for r, an, nv in linhas:
            rr = [i for i in range(4, 44) if str(a.Range(f'A{i}').Value or '').strip() == an]
            cvi = a.Range(f'{cvi_col}{rr[0]}').Value if rr else None
            fonte, cvtp = e.Range(f'I{r}').Value, e.Range(f'J{r}').Value
            cvim = cvi if num(cvi) and cvi > 0 else None       # so o CVI cadastrado (sem 2 x CVTp)
            meta, classe, uc = e.Range(f'AN{r}').Value, e.Range(f'AO{r}').Value, e.Range(f'AK{r}').Value
            if cvim is None:
                if meta not in ('', None) or (num(uc) and classe != 'Sem meta'):
                    dif.append((r, an, nv, 'sem CVI', meta, classe))
                continue
            if not perto(meta, 0.5 * cvim):
                dif.append((r, an, nv, 'meta', meta, 0.5 * cvim))
            if num(uc):
                q_ = uc / cvim
                cl = 'Ótimo' if q_ <= .25 else 'Desejável' if q_ <= .5 else 'Mínimo' if q_ <= .75 else 'Não atende'
                if classe != cl:
                    dif.append((r, an, nv, 'classe', classe, cl, round(q_, 3)))
        reg('I02 Meta u = 0,50 × CVI (EFLM; só o CVI cadastrado na aba Analitos; nunca do TEa CLIA) e classe '
            'Ótimo/Desejável/Mínimo/Não atende pelos cortes 0,25/0,50/0,75 × CVI', not dif, {'divergencias': dif[:8]})

        # ADR-067: sem certificado de calibrador -- nenhum campo, nome ou formula de u(cal)
        tem_nome = any(n_.Name == 'ucalFabricante' for n_ in q.wb.Names)
        cab_ucal = [c for c in range(1, 80) if 'u(cal)' in str(a.Cells(3, c).Value or '')]
        f_ucal = [c for c in range(32, 48) if 'ucal' in str(e.Cells(14, c).Formula).lower()]
        reg('I03 Sem u(cal) (ADR-067): nome ucalFabricante, campo na Analitos e referências nas fórmulas removidos; '
            'a incerteza usa só CIQ + CEQ', not tem_nome and not cab_ucal and not f_ucal,
            {'nome': tem_nome, 'cabecalho_analitos': cab_ucal, 'formulas': f_ucal})

        # CEQ
        ano, prov_f, rod = e.Range('N4').Value, e.Range('L4').Value, e.Range('P4').Value
        dif, vistos, amostra = [], set(), {}
        for r, an, nv in linhas:
            if an in vistos:
                continue
            vistos.add(an)
            prov = prov_f
            if bio:
                rr = [i for i in range(4, 44) if str(a.Range(f'A{i}').Value or '').strip() == an]
                prov = a.Range(f'AR{rr[0]}').Value if rr and a.Range(f'AR{rr[0]}').Value else 'CAP'
            x = esperado_ceq(q, an, ano, prov, rod)
            mv, tri, ub = e.Range(f'AR{r}').Value, str(e.Range(f'AS{r}').Value), e.Range(f'AJ{r}').Value
            if x is None:
                if num(mv) or not tri.startswith('não verificável'):
                    dif.append((an, 'sem CEQ', mv, tri))
                continue
            if not perto(mv, x['media']):
                dif.append((an, 'media', mv, x['media']))
            if not x['ok']:
                ok_tri = tri.startswith('não verificável')
            elif not x['detectavel']:
                ok_tri = tri == 'não detectável'
            else:
                ok_tri = tri.startswith(('detectável', 'RELEVANTE'))
            if not ok_tri:
                dif.append((an, 'triagem', tri, x))
            if x['ubias'] is not None:
                if not perto(ub, x['ubias']):
                    dif.append((an, 'u(bias)', ub, x['ubias']))
            elif ub not in ('', None):
                dif.append((an, 'u(bias) sem criterio', ub))
            if len(amostra) < 4:
                amostra[an] = {'media': round(x['media'], 3), 'n': x['n'], 'rodadas': x['rodadas'], 'triagem': tri}
        reg('I04 Viés do CEQ recalculado da EQA_Base: média com sinal, triagem (≥ 6 amostras de ≥ 2 rodadas; '
            'detectável se |média| > 2 EP) e u(bias) Nordtest = raiz(RMS² + u(Cref)²), u(Cref) = DP do grupo / raiz(nº de '
            'laboratórios), só com o critério',
            not dif, {'analitos': len(vistos), 'exemplos': amostra, 'divergencias': dif[:6]})

        # Painel = Estatistica
        nlv = 2 if bio else 3
        r0 = [r for r in range(40, 160) if str(p.Cells(r, 1).Value or '').startswith('INCERTEZA DE MEDIÇÃO')][0]
        sel = str(q.wb.Names('selAnalito').RefersToRange.Value)
        dif = []
        for k in range(1, nlv + 1):
            rr = [x[0] for x in linhas if x[1] == sel and x[2] == str(k)]
            if not rr:
                continue
            for pc, ec in (('C', 'AF'), ('E', 'AL'), ('H', 'AM'), ('J', 'AN'), ('K', 'AO'), ('M', 'AP'), ('O', 'AR'), ('P', 'AS')):
                pv, ev = p.Range(f'{pc}{r0 + 2 + k}').Value, e.Range(f'{ec}{rr[0]}').Value
                if pv != ev and not (num(pv) and perto(pv, ev)):
                    dif.append((k, pc, pv, ev))
        nota = str(p.Cells(r0 + 1, 1).Value)
        reg('I05 Bloco "INCERTEZA DE MEDIÇÃO" do Painel mostra, por nível, exatamente a linha da Estatística do analito '
            'selecionado, com a janela na nota', not dif and nota.startswith('Janela') and e.Range('AG11').Text in nota,
            {'linha': r0, 'analito': sel, 'nota': nota[:90], 'divergencias': dif[:6]})

        # sem erro de formula
        err = []
        for r, an, nv in linhas:
            for c in range(32, 48):
                t = e.Cells(r, c).Text
                if t.startswith('#'):
                    err.append((r, c, t))
        for r in range(r0, r0 + 4 + nlv):
            for c in range(1, 22):
                if p.Cells(r, c).Text.startswith('#'):
                    err.append(('Painel', r, c))
        reg('I06 Nenhum erro de fórmula (#VALOR!, #NOME?...) nas colunas da incerteza nem no bloco do Painel',
            not err, {'erros': err[:8]})

        # CEQ da Hematologia e texto acentuado do mEQA
        if not bio:
            wbc = [x[0] for x in linhas if x[1] == 'WBC'][0]
            ac, t_, rs = e.Range(f'AC{wbc}').Value, e.Range(f'T{wbc}').Value, str(e.Range(f'R{wbc}').Value)
            reg('I07 Hematologia: as colunas do CEQ (bias com sinal, nº de amostras e rodadas, status SDI e limites) usam o '
                'provedor do filtro como o bias de G -- antes liam Analitos!AR, que só existe na Bioquímica, e davam SEM EP',
                num(ac) and ac > 0 and num(t_) and not rs.startswith('SEM EP'),
                {'WBC': {'N_EQA': ac, 'bias_sinal': t_, 'status_SDI': rs[:40]}})
        pad = {s_: q.run('mEQA.PadronizarStatus', s_) for s_ in ('Aceitável', 'Não aceito', 'NÃO CONFORME', 'Insatisfatório',
                                                                 'Acceptable')}
        reg('I08 Padronização do status do CEQ reconhece rótulos acentuados (Controllab em português) -- na Bioquímica '
            'o texto estava corrompido ("ACEITÃVEL")',
            pad == {'Aceitável': 'ACEITO', 'Não aceito': 'NAO ACEITO', 'NÃO CONFORME': 'NAO ACEITO',
                    'Insatisfatório': 'NAO ACEITO', 'Acceptable': 'ACEITO'}, pad)

        # linhas sem analito cadastrado (A = 0): nada calculado
        fant = [r for r in range(14, 134) if e.Range(f'A{r}').Value == 0]
        cheias = [(r, c) for r in fant for c in range(32, 48) if e.Cells(r, c).Value not in ('', None)]
        reg('I10 Linhas sem analito cadastrado (A = 0) não calculam nem mostram incerteza', not cheias,
            {'linhas_sem_analito': len(fant), 'celulas_preenchidas': cheias[:6]})

        # sem u(bias): o motivo dito e o do CEQ (insuficiente / sem u(Cref) / sem CEQ)
        motivos = {}
        for r, an, nv in linhas:
            s_ = str(e.Range(f'AP{r}').Value or '')
            if s_.startswith('U não estimada: '):
                motivos[s_[16:].split(' (')[0]] = motivos.get(s_[16:].split(' (')[0], 0) + 1
        conhecidos = ('CEQ insuficiente', 'sem u(Cref): faltam DP do grupo e nº de laboratórios no CEQ',
                      'sem CEQ para o analito', 'sem CEQ utilizável para o analito', 'sem u(bias) do CEQ')
        reg('I11 Sem u(bias), a situação diz o motivo do CEQ (insuficiente, sem u(Cref), sem CEQ) -- U não é inventada '
            'com u(Rw) sozinha', all(m_ in conhecidos for m_ in motivos), {'motivos': motivos})

        # decisao manual de Uso_Analitico na EQA_Base sobrevive a consolidacao
        b_ = q.wb.Worksheets('EQA_Base')
        try:
            b_.Unprotect('qcini2025')
        except Exception:
            pass
        ult_ = b_.Cells(b_.Rows.Count, 21).End(-4162).Row
        alvo_l = next(i for i in range(2, ult_ + 1) if str(b_.Cells(i, 21).Value or '').strip())
        chave = str(b_.Cells(alvo_l, 21).Value)
        antes = str(b_.Cells(alvo_l, 20).Value)
        novo = 'NAO' if antes.upper() == 'SIM' else 'SIM'
        b_.Cells(alvo_l, 20).Value = novo
        q.run('mEQA.AtualizarEQABase')
        ult_ = b_.Cells(b_.Rows.Count, 21).End(-4162).Row
        achou = [i for i in range(2, ult_ + 1) if str(b_.Cells(i, 21).Value) == chave]
        depois = str(b_.Cells(achou[0], 20).Value) if achou else None
        carimbo = ' '.join(str(b_.Cells(1, c).Value or '') for c in range(22, 30))
        if achou:
            b_.Cells(achou[0], 20).Value = antes
        reg('I12 Decisão manual de Uso_Analitico na EQA_Base sobrevive a uma nova consolidação do CEQ '
            '(antes era apagada sem aviso)', depois == novo and 'uso manual mantido' in carimbo,
            {'chave': chave, 'antes': antes, 'manual': novo, 'depois': depois, 'carimbo': carimbo.strip()[-80:]})

        if not bio:
            k5 = str(e.Range('K5').Value or '')
            pv = q.wb.Names('eqProvedor').RefersToRange
            reg('I13 Hematologia: provedor do CEQ definido (CAP, editável, célula própria fora da nota mesclada); '
                'antes eqProvedor caía no meio de F4:M4 e ficava vazio ("todos os provedores", 0 linhas em K5)',
                str(pv.Value) == 'CAP' and pv.MergeArea.Count == 1 and not pv.Locked and ' 0 linha' not in k5,
                {'eqProvedor': pv.Address, 'valor': pv.Value, 'K5': k5[:120]})

        # desempenho do recalculo
        t0 = time.perf_counter()
        q.run('mIncerteza.RecalcularIncerteza')
        q.ex.esperar()
        dt = time.perf_counter() - t0
        reg('I09 Recalcular toda a incerteza (depois de ATUALIZAR DADOS) leva menos de 10 s', dt < 10,
            {'segundos': round(dt, 2)})
    finally:
        q.fechar()
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [x for x in RES if x['resultado'] == 'FAIL']
    print(f'\n=== {produto} incerteza: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


def dia_menos_12m(fim):
    """EDATE(fim, -12) + 1, em serial do Excel."""
    import datetime as dt
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(fim))
    y, m = d.year - 1, d.month
    import calendar
    dd = min(d.day, calendar.monthrange(y, m)[1])
    return float((dt.date(y, m, dd) - dt.date(1899, 12, 30)).days + 1)


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
