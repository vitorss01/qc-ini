# -*- coding: utf-8 -*-
"""repro_dados.py -- Tester/QA, ciclo 2026-10-31_ciclo01, modo REPRODUZIR.

Conferencias SO DE LEITURA (openpyxl, sem Excel) sobre COPIAS da producao. Cobre as ocorrencias que sao leitura de dado:

  D1  valores <= 0 ATIVOS (Bio: BIO-371183/371184; varredura da tblCQ_Final; Hema: varredura)
  D2  suspeitos de estudos/cv_suspeitos_<produto>.csv ainda ativos; IDs citados; CVp por lote e Sigma
  D3  bias do CEQ com amostras de resultado 0 (Bilirrubina direta): BiasEQ (com) x filtro do ViesEQ (sem)
  D4  lotes que nao sao do analito (urina 5659/6353/6354, PCR 8353, HbA1c no 8974) ativos em analitos sericos
  D5  formula do ET do Painel (H7:H9) e G/H da Estatistica para analitos sem CEQ
  D6  protecao da aba EQA.CAP_Dados e Locked das colunas digitaveis
  D7  (Hema) alvo do bloco do lote 522611 x media medida (z em DP do alvo)
  D8  (Hema) mapa W:Y da EQA_Base: nomes que mapeiam para o vizinho (Lymphs -> LYMPH%)

Uso: python -u repro_dados.py <Bio_copia.xlsm> <Hema_copia.xlsm> <saida.json>
Nada e gravado nas copias (load_workbook read_only).

Metodo independente: nao usa os scripts do Usuario Simulado. O estimador do bias segue mCEQ.BiasEQ (ADR-035: media
por rodada dos |bias|, depois media das rodadas); o "sem zeros" aplica a regra do mCEQ.ViesEQ (xlab = 0 ou alvo = 0
fora). O CVp segue mIncerteza.IncertezaCIQ (lotes com n >= 20, raiz(soma (n-1) CV^2 / soma (n-1))).
"""
import csv
import datetime as dt
import json
import math
import os
import statistics as st
import sys
from collections import defaultdict

import openpyxl

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
ESTUDOS = os.path.join(RAIZ, '_arquitetura', 'integracao_dbseac', 'estudos')


def cq_final(wb):
    ws = wb['Principal - Resultados']
    it = ws.iter_rows(min_row=6, values_only=True)
    cab = list(next(it))
    out = []
    for r in it:
        if r[0] is None:
            continue
        out.append(dict(zip(cab, r)))
    return out


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def ativo(r):
    return r.get('PARTICIPA_ESTATISTICA') == 'SIM' and r.get('STATUS_ANALITICO') == 'ATIVO'


def mdcv(vals):
    if len(vals) < 2:
        return {'n': len(vals)}
    m = st.fmean(vals)
    s = st.stdev(vals)
    return {'n': len(vals), 'media': m, 'dp': s, 'cv': s / abs(m) * 100 if m else None}


def cv_pooled(linhas, ini, fim):
    porlote = defaultdict(list)
    for r in linhas:
        d = r['DATA']
        if isinstance(d, dt.datetime) and ini <= d <= fim and num(r['RESULTADO']):
            porlote[str(r['LOTE'])].append(float(r['RESULTADO']))
    sq = gl = 0.0
    lotes = {}
    for lt, v in porlote.items():
        n = len(v)
        if n < 20:
            lotes[lt] = {'n': n, 'fora': 'n<20'}
            continue
        m = st.fmean(v)
        if m == 0:
            continue
        cv = st.stdev(v) / abs(m) * 100
        lotes[lt] = {'n': n, 'cv': cv}
        sq += (n - 1) * cv * cv
        gl += n - 1
    return (math.sqrt(sq / gl) if gl else None), lotes


def estat_linhas(wb, ncol=48):
    ws = wb['Estatística']
    out = {}
    for i, r in enumerate(ws.iter_rows(min_row=14, max_row=120, max_col=ncol, values_only=True), 14):
        if r[0] and r[1] is not None and f'{r[0]}|{r[1]}' not in out:   # 1a ocorrencia = tabela principal
            out[f'{r[0]}|{r[1]}'] = {'linha': i, 'C_n': r[2], 'D_media': r[3], 'E_dp': r[4], 'F_cv': r[5],
                                     'G_bias': r[6], 'H_et': r[7], 'K_etp': r[10], 'L_sigma': r[11],
                                     'M_status': r[12]}
    return out


# ----------------------------------------------------------------------------------------------- Bioquimica
def bio(path):
    R = {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    cq = cq_final(wb)
    porid = {r['ID_REGISTRO']: r for r in cq}
    R['linhas_cq'] = len(cq)

    # D1 ----------------------------------------------------------------- valores <= 0 ATIVOS
    campos = ('ID_REGISTRO', 'ORIGEM_RESULTADO', 'EQUIPAMENTO', 'DATA_HORA', 'RUN', 'NIVEL', 'LOTE', 'ANALITO',
              'ANALITO_CADASTRADO', 'RESULTADO', 'FLAG', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA')
    R['D1_ids'] = {i: {k: str(porid[i][k]) for k in campos} for i in ('BIO-371183', 'BIO-371184') if i in porid}
    naopos = [r for r in cq if ativo(r) and num(r['RESULTADO']) and r['RESULTADO'] <= 0]
    R['D1_ativos_le0'] = {
        'total': len(naopos),
        'negativos': sum(1 for r in naopos if r['RESULTADO'] < 0),
        'zeros': sum(1 for r in naopos if r['RESULTADO'] == 0),
        'cadastrados': sum(1 for r in naopos if r['ANALITO_CADASTRADO'] == 'SIM'),
        'origem': dict(sorted(defaultdict(int, {o: sum(1 for r in naopos if r['ORIGEM_RESULTADO'] == o)
                                                for o in {r['ORIGEM_RESULTADO'] for r in naopos}}).items())),
        'data_min': str(min(r['DATA'] for r in naopos)) if naopos else None,
        'data_max': str(max(r['DATA'] for r in naopos)) if naopos else None,
        'por_analito': dict(sorted(((a, sum(1 for r in naopos if r['ANALITO'] == a))
                                    for a in {r['ANALITO'] for r in naopos}), key=lambda x: -x[1])),
    }
    # Glicose N2 lote 8977 DIM1 outubro: com e sem os dois
    sel = [r for r in cq if ativo(r) and r['ANALITO'] == 'Glicose' and str(r['LOTE']) == '8977'
           and r['EQUIPAMENTO'] == 'DIMENSION 1' and r['NIVEL'] == 2
           and dt.datetime(2026, 10, 1) <= r['DATA'] <= dt.datetime(2026, 10, 31)]
    R['D1_glicose_n2_out'] = {'com': mdcv([r['RESULTADO'] for r in sel]),
                              'sem_371183_371184': mdcv([r['RESULTADO'] for r in sel
                                                         if r['ID_REGISTRO'] not in ('BIO-371183', 'BIO-371184')])}

    # D2 ----------------------------------------------------------------- suspeitos ainda ativos + Sigma
    susp = list(csv.DictReader(open(os.path.join(ESTUDOS, 'cv_suspeitos_bioquimica.csv'), encoding='utf-8-sig'),
                               delimiter=';'))
    ainda = [s['id'] for s in susp if s['id'] in porid and ativo(porid[s['id']])]
    R['D2_suspeitos'] = {'lista': len(susp), 'ainda_ativos': len(ainda),
                         'nao_encontrados': [s['id'] for s in susp if s['id'] not in porid]}
    ids = ('BIO-319726', 'BIO-232288', 'BIO-232289', 'BIO-268082', 'BIO-215815', 'BIO-215820')
    R['D2_ids'] = {i: {k: str(porid[i][k]) for k in campos} for i in ids if i in porid}
    est = estat_linhas(wb)
    ini, fim = dt.datetime(2026, 4, 29), dt.datetime(2026, 10, 28)   # janela do Sigma (Estatistica N11:P11)
    eq_painel = wb['Painel']['L4'].value          # o CVp do Sigma e do equipamento do Painel (conferido: DIM2)
    R['D2_equip_painel'] = eq_painel
    g1 = [r for r in cq if ativo(r) and r['ANALITO'] == 'Glicose' and r['NIVEL'] == 1 and r['EQUIPAMENTO'] == eq_painel]
    cvp_com, lotes_com = cv_pooled(g1, ini, fim)
    cvp_sem, lotes_sem = cv_pooled([r for r in g1 if r['ID_REGISTRO'] != 'BIO-319726'], ini, fim)
    e = est.get('Glicose|1', {})
    sig_sem = None
    if num(e.get('K_etp')) and num(e.get('G_bias')) and cvp_sem:
        sig_sem = (e['K_etp'] - abs(e['G_bias'])) / cvp_sem
    R['D2_glicose_n1_sigma'] = {'estatistica': e, 'cvp_com': cvp_com, 'lotes_com': lotes_com,
                                'sigma_recalc_com': ((e['K_etp'] - abs(e['G_bias'])) / cvp_com)
                                if num(e.get('K_etp')) and num(e.get('G_bias')) and cvp_com else None,
                                'cvp_sem_319726': cvp_sem, 'lotes_sem': lotes_sem, 'sigma_sem_319726': sig_sem}
    st_sig = defaultdict(int)
    for k, v in est.items():
        st_sig[str(v['M_status'])] += 1
    R['D2_status_sigma'] = dict(st_sig)
    pt = est.get('Proteína total|1', {}), est.get('Proteína total|2', {})
    R['D2_proteina_total'] = {'N1': pt[0], 'N2': pt[1]}

    # D3 ----------------------------------------------------------------- bias com/sem resultado 0
    ws = wb['EQA_Base']
    it = ws.iter_rows(min_row=1, max_col=21, values_only=True)
    cab = list(next(it))
    base = [dict(zip(cab, r)) for r in it if r[0]]
    R['D3_cab'] = cab
    bd = [b for b in base if (b[cab[4]] or '') == 'Bilirrubina direta' and str(b[cab[0]]).upper() == 'CAP'
          and str(b[cab[1]]) == '2025' and str(b[cab[19]]).upper() == 'SIM']

    def bias2(linhas, filtro):
        rod = defaultdict(list)
        for b in linhas:
            if not num(b[cab[16]]):
                continue
            if filtro and (b[cab[6]] == 0 or b[cab[7]] == 0):
                continue
            rod[f'{b[cab[1]]}|{b[cab[2]]}'].append(float(b[cab[16]]))
        med = {k: st.fmean(v) for k, v in rod.items()}
        return (st.fmean(med.values()) if med else None), {k: (len(rod[k]), med[k]) for k in med}

    com, rcom = bias2(bd, False)
    sem, rsem = bias2(bd, True)
    zeros = [{'chave': b[cab[20]], 'res': b[cab[6]], 'alvo': b[cab[7]], 'bias_abs': b[cab[16]]}
             for b in bd if b[cab[6]] == 0 or b[cab[7]] == 0]
    e1, e2 = est.get('Bilirrubina direta|1', {}), est.get('Bilirrubina direta|2', {})
    out = {'n_amostras': len(bd), 'zeros': zeros, 'bias_com': com, 'rodadas_com': rcom, 'bias_sem': sem,
           'rodadas_sem': rsem, 'estat_G20': e1.get('G_bias'), 'estat_N1': e1, 'estat_N2': e2}
    if num(e1.get('F_cv')) and sem is not None:
        cvp1 = (e1['K_etp'] - abs(e1['G_bias'])) / e1['L_sigma']
        cvp2 = (e2['K_etp'] - abs(e2['G_bias'])) / e2['L_sigma']
        out['sem_zeros'] = {'ET_N1': 1.65 * e1['F_cv'] + sem, 'ET_N2': 1.65 * e2['F_cv'] + sem,
                            'Sigma_N1': (e1['K_etp'] - sem) / cvp1, 'Sigma_N2': (e2['K_etp'] - sem) / cvp2,
                            'CVp_N1_deduzido': cvp1, 'CVp_N2_deduzido': cvp2}
        out['com_zeros_recalc'] = {'ET_N1': 1.65 * e1['F_cv'] + com}
    R['D3_bilirrubina_direta'] = out

    # D4 ----------------------------------------------------------------- lotes que nao sao do analito
    def resumo(f):
        ls = [r for r in cq if f(r)]
        at = [r for r in ls if ativo(r)]
        return {'linhas': len(ls), 'ativas': len(at),
                'matriz': sorted({str(r['MATRIZ']) for r in ls}),
                'por_nivel': {n: mdcv([r['RESULTADO'] for r in at if r['NIVEL'] == n and num(r['RESULTADO'])])
                              for n in (1, 2)},
                'exemplos': [{k: str(r[k]) for k in ('ID_REGISTRO', 'DATA_HORA', 'EQUIPAMENTO', 'NIVEL',
                                                      'ANALITO_ORIGEM', 'RESULTADO', 'MATRIZ', 'STATUS_ANALITICO')}
                             for r in at[:3]]}
    R['D4_5659_glicose'] = resumo(lambda r: str(r['LOTE']) == '5659' and r['ANALITO'] == 'Glicose')
    R['D4_6353_ureia'] = resumo(lambda r: str(r['LOTE']) == '6353' and r['ANALITO'] == 'Ureia')
    R['D4_6354_creatinina'] = resumo(lambda r: str(r['LOTE']) == '6354' and r['ANALITO'] == 'Creatinina')
    R['D4_8353_naoPCR'] = resumo(lambda r: str(r['LOTE']) == '8353' and r['ANALITO'] != 'Proteína C-reativa')
    R['D4_8974_hba1c'] = resumo(lambda r: str(r['LOTE']) == '8974' and r['ANALITO'] == 'Hemoglobina glicada')
    R['D4_ids'] = {i: {k: str(porid[i][k]) for k in campos + ('MATRIZ',)}
                   for i in ('BIO-230952', 'BIO-230958', 'BIO-355958', 'BIO-369955', 'BIO-369956') if i in porid}
    # varredura lote x analito: lotes cujo analito "dono" e outro (mais de 90% das linhas do lote em outro analito
    # nao serve: o lote de controle multianalito e normal). Criterio simples: lotes de urina conhecidos
    urina = defaultdict(lambda: defaultdict(int))
    for r in cq:
        if str(r['LOTE']) in ('5659', '6353', '6354') and ativo(r):
            urina[str(r['LOTE'])][r['ANALITO']] += 1
    R['D4_lotes_urina_ativos_por_analito'] = {k: dict(v) for k, v in urina.items()}

    # D5 ----------------------------------------------------------------- formula do ET do Painel
    R['D5_estat_sem_ceq'] = {k: est.get(k) for k in ('Hemoglobina glicada|1', 'Hemoglobina glicada|2',
                                                      'Proteína C-reativa|1', 'Proteína C-reativa|2')}
    sem_ep = [k for k, v in est.items() if v['G_bias'] == 'SEM EP']
    R['D5_linhas_sem_ep'] = sem_ep
    wb.close()
    wbf = openpyxl.load_workbook(path, read_only=True, data_only=False)
    p = wbf['Painel']
    R['D5_formulas'] = {c: p[c].value for c in ('G7', 'H7', 'H8')}
    e = wbf['Estatística']
    R['D5_formula_estat_H14'] = e['H14'].value
    wbf.close()
    return R


# ----------------------------------------------------------------------------------------------- Hematologia
def hema(path):
    R = {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    cq = cq_final(wb)
    R['linhas_cq'] = len(cq)
    naopos = [r for r in cq if ativo(r) and num(r['RESULTADO']) and r['RESULTADO'] <= 0]
    R['D1_ativos_le0'] = {'total': len(naopos),
                          'por_origem': {o: sum(1 for r in naopos if r['ORIGEM_RESULTADO'] == o)
                                         for o in {r['ORIGEM_RESULTADO'] for r in naopos}},
                          'por_analito': {a: sum(1 for r in naopos if r['ANALITO'] == a)
                                          for a in {r['ANALITO'] for r in naopos}}}
    est = estat_linhas(wb)

    # D7 ----------------------------------------------------------------- alvo do lote 522611
    ws = wb['Analitos']
    blk = list(ws.iter_rows(min_row=1, max_row=40, max_col=10, values_only=True))
    R['D7_bloco_lote'] = blk[1][1]
    alvos = {}
    for r in blk[5:]:
        if r[0] and isinstance(r[4], (int, float)):
            alvos[r[0]] = r[4:10]
    medidas = defaultdict(list)
    for r in cq:
        if ativo(r) and str(r['LOTE']) == '522611' and num(r['RESULTADO']):
            medidas[(r['ANALITO'], r['NIVEL'])].append(float(r['RESULTADO']))
    pares = []
    for a, v in alvos.items():
        for n in (1, 2, 3):
            m_alvo, dp_alvo = v[2 * (n - 1)], v[2 * (n - 1) + 1]
            x = medidas.get((a, n), [])
            if not x or not num(m_alvo) or not num(dp_alvo) or dp_alvo == 0:
                continue
            z = (st.fmean(x) - m_alvo) / dp_alvo
            pares.append({'analito': a, 'nivel': n, 'alvo': m_alvo, 'dp_alvo': dp_alvo, 'n': len(x),
                          'media_medida': st.fmean(x), 'z': z})
    pares.sort(key=lambda p: -abs(p['z']))
    R['D7_pares_z_gt5'] = [p for p in pares if abs(p['z']) > 5]
    R['D7_n_pares'] = len(pares)
    # mesmo teste nos outros lotes com bloco? os blocos ficam na LotesStore; aqui so o lote em tela (B2)

    # D8 ----------------------------------------------------------------- mapa W:Y
    ws = wb['EQA_Base']
    mapa = []
    for r in ws.iter_rows(min_row=2, max_row=201, min_col=23, max_col=25, values_only=True):
        if r[1]:
            mapa.append(r)
    R['D8_mapa_n'] = len(mapa)
    R['D8_mapa_provedores'] = sorted({str(r[0]) for r in mapa})
    R['D8_lymph'] = [r for r in mapa if 'lymph' in str(r[1]).lower()]
    # pares "nome curto -> % / nome com Absolute -> #"
    R['D8_pares_pct_abs'] = [(r[1], r[2]) for r in mapa if str(r[2]).endswith('%') or str(r[2]).endswith('#')]

    # D5 ----------------------------------------------------------------- ET do Painel
    R['D5_estat_sem_ceq'] = {k: est.get(k) for k in ('RET%|1', 'RET%|2', 'RET%|3', 'RDW-SD|1', 'RDW-SD|2',
                                                      'RDW-SD|3')}
    R['D5_linhas_sem_ep'] = [k for k, v in est.items() if v['G_bias'] == 'SEM EP']
    wb.close()
    wbf = openpyxl.load_workbook(path, read_only=True, data_only=False)
    p = wbf['Painel']
    R['D5_formulas'] = {c: p[c].value for c in ('G7', 'H7', 'H8', 'H9')}
    wbf.close()
    return R


def protecao(path):
    """D6: protecao da aba e Locked das colunas A..R da tabela do CAP (precisa do modo nao read_only)."""
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True, keep_links=False)
    ws = wb['EQA.CAP_Dados']
    out = {'sheet_protegida': bool(ws.protection.sheet), 'tem_senha': bool(ws.protection.password)}
    ult = ws.max_row
    lock = {}
    for col in 'ABCDEFGHIJKLMNOPQR':
        lock[col] = ws[f'{col}{ult}'].protection.locked
    out['linha_amostrada'] = ult
    out['locked_A_R'] = lock
    out['validacoes'] = len(ws.data_validations.dataValidation)
    wb.close()
    return out


if __name__ == '__main__':
    bio_p, hema_p, saida = sys.argv[1:4]
    res = {'gerado_em': dt.datetime.now().isoformat(timespec='seconds')}
    print('Bio...', flush=True)
    res['Bioquimica'] = bio(bio_p)
    print('Hema...', flush=True)
    res['Hematologia'] = hema(hema_p)
    print('protecao...', flush=True)
    res['Bioquimica']['D6_protecao_cap'] = protecao(bio_p)
    res['Hematologia']['D6_protecao_cap'] = protecao(hema_p)
    json.dump(res, open(saida, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
    print('OK', saida)
