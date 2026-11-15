# -*- coding: utf-8 -*-
"""reteste04_usuario_f2.py -- Usuario Simulado, ciclo 2026-11-14_reteste04, frente F2, modo RETESTE (Bioquimica).

Refaz numa COPIA NOVA da producao atual (entrega 41) o fluxo original de OC-0040, OC-0043, OC-0056/ADR-093, OC-0057,
OC-0058, OC-0059 (documento), OC-0066/ADR-094 e ADR-091, com o fluxo vizinho de cada um, e o pedido do ciclo:
incluir e corrigir um resultado de CEQ TESTE-AUDITORIA e conferir que bias, ET e Sigma do analito certo (e so dele)
mudam como deveriam.

Secoes (uma copia, um Excel, nada salvo na copia de trabalho; retratos SaveCopyAs com eventos desligados):
  estaticas  (arquivo recem-aberto, sem login) OC-0040: K3/K4 do Painel em mes, trimestre, "Tudo", De/Ate
             atravessando o ano; cabecalho da Estatistica (A4, B4, D4, H5, J5, travas) e n/media/DP/CV da Glicose
             contra contagem independente da tblCQ_Final. Retratos S_mes / S_trim para a calculadora.
             OC-0043: J7:J8 -- largura util medida, fonte efetiva, ShrinkToFit, fator de reducao de cada forma curta
             e estados reais provocados pelo filtro (REJEITADO, PER. VAZIO, SEM CORRIDA) com o Q4.
  ceq        (login ANALISTA de teste) Proteina total / CAP "Protein, total, serum": retrato S0 -> SalvarCEQ
             (7,7 sobre alvo 7 = +10 %) -> S1 -> CorrigirCEQ (7,35 = +5 %) -> S2 -> vizinhos: Controllab do mesmo
             analito e CAP de outro ano -> S3 -> duplicidade -> ExcluirCEQ de tudo -> S4 (tem de voltar ao S0).
             Diferenca celula a celula da Estatistica A14:AU133 entre retratos (so o analito pode mudar).
  manuais    Digitar Resultados (eventos ligados): Creatinina no lote de URINA (ADR-093, aceito como CRE2), analito
             de soro sem codigo recebido na urina (recusado, A13), -1 (IMPLAUSIVEL), 8794 (LOTE_FORA_DO_CQ),
             8353 Glicose (SO_ANALITOS -> LOTE_OUTRO_MATERIAL), 8353 PCR (aceito), Glicose comum (controle).
             Oraculo proprio do de/para (tblDeParaAnalitos x interfaceamento da tblCQ_Final).
             + inativar um ponto real (mIntegracao.InativarPorId, o nucleo da janela) + comentario tecnico
             (GravarComentarioTecnico, acao COMENTAR) -> ATUALIZAR DADOS (AtualizarDadosAutomatico, le o DB_SEAC so
             para leitura) -> trilha do Audit_Log por manual, QA I05/A13 -> 2o ATUALIZAR (so entradas) sem trilha
             repetida.
  adr091     oraculo proprio das regras da tblLotesMaterial sobre a tblCQ_Final + Painel (8353 PCR x Glicose).
  adr094     varredura de TODAS as tabelas (ListObjects) de todas as abas: nenhum texto "####"; data-hora com
             segundos onde o valor tem hora; so data onde a coluna e de data; linhas novas desta sessao (inativacao,
             comentario, manual, Sync_Log INICIO/FIM, QA).
  sem_sessao Logout -> SalvarCEQ/CorrigirCEQ/ExcluirCEQ devolvem ERRO|0 e nada muda (OC-0057); AbrirCEQ sem login;
             vizinho: login de novo -> Corrigir e Excluir funcionam.
  calc       (sem Excel) calc_independente.py sobre os retratos.

Regras: so COPIA em %TEMP%/qcw/auditoria/2026-11-14_reteste04/; fechar() encerra so o PID do QA; nenhuma captura de
tela (caixas lidas pelo texto das janelas do PID desta sessao). Dados de teste marcados TESTE-AUDITORIA.

  parte 2 (--parte2, 2a copia nova): CEQ de Glicose (analito com CIQ no periodo, para o ET mudar tambem):
             G0 -> SalvarCEQ 110/100 -> G1 -> CorrigirCEQ 105 -> G2 -> linha de mapa TESTE-AUDITORIA do Controllab
             para Glicose (EQA_Base W:Y, apagada depois) + amostra Controllab 200/100 -> G3 (nao pode mudar: o
             provedor da Glicose e o CAP, Especificacoes!AR) -> ExcluirCEQ de tudo -> G4 (= G0); depois manual +
             inativar + comentar + ATUALIZAR e a varredura do ADR-094; Painel nos lotes com regra (8353, 8794, 6354).

Uso: python -u reteste04_usuario_f2.py            (Excel + calc)
     python -u reteste04_usuario_f2.py --parte2   (2a copia: Glicose/Controllab, ADR-094, Painel dos lotes; + calc G*)
     python -u reteste04_usuario_f2.py --so-calc  (so a calculadora sobre os retratos ja gravados)
Saida: %TEMP%/qcw/auditoria/2026-11-14_reteste04/Bioquimica_F2.json, Bioquimica_F2_p2.json (+ calc_*.json)
Nota (1a rodada, 14/11): o oraculo do de/para considerava "tem de/para" qualquer linha da matriz; o codigo
(mEtlCQ ParaQcini) exige ANALITO_QCINI preenchido. Corrigido aqui; o veredito foi refeito sobre o retrato S0.
"""
import collections
import datetime as dt
import json
import os
import re
import statistics as st
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Bioquimica']
import valida_adr090_091 as V  # noqa: E402  (so infraestrutura: Audit, Vigia, ler/escrever manual)
sys.argv = _argv
import repro_comum as C  # noqa: E402
sys.path.insert(0, C.TESTES)
import auditoria_abas  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = 'Bioquimica'
CICLO = '2026-11-14_reteste04'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
USR = 'TESTE_AUDITORIA_F2'
TAG = 'TESTE-AUDITORIA'
MOTIVO = 'Outro: TESTE-AUDITORIA'
AN, NOME, UN = 'Proteína total', 'Protein, total, serum', 'g/dL'
ROD = 'TESTE-AUDITORIA F2 R1'
t, up, num = V.t, V.up, V.num
RETR = {}
MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


def retrato(qa, nome):
    caminho = os.path.join(C.PASTA, f'{PROD}_F2_{nome}.xlsm')
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    try:
        qa.wb.SaveCopyAs(caminho)
    finally:
        qa.ex.xl.EnableEvents = ev
    RETR[nome] = caminho
    return os.path.basename(caminho)


def estat_all(qa):
    v = qa.ws('Estatística').Range('A14:AU133').Value
    return [list(r) for r in v]


def col_letra(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def diff_estat(a, b):
    out = []
    for i, (ra, rb) in enumerate(zip(a, b)):
        for j, (x, y) in enumerate(zip(ra, rb)):
            if x != y:
                out.append({'cel': f'{col_letra(j)}{14 + i}', 'analito': t(ra[0]), 'nivel': t(ra[1]),
                            'antes': x if not isinstance(x, dt.datetime) else str(x),
                            'depois': y if not isinstance(y, dt.datetime) else str(y)})
    return out


def linhas_an(est, an):
    cols = {'C_n': 2, 'D_media': 3, 'E_dp': 4, 'F_cv': 5, 'G_bias': 6, 'H_et': 7, 'K_etp': 10, 'L_sigma': 11,
            'M_class': 12, 'T_sinal': 19, 'AC_nEQA': 28, 'AD_nrod': 29}
    return {f'{t(r[0])}|{t(r[1])}': {k: r[i] for k, i in cols.items()} for r in est if t(r[0]) == an}


def painel_an(qa):
    p = qa.ws('Painel')
    v = p.Range('A7:J8').Value
    cab = ['nivel', 'n', 'media', 'dp', 'cv', 'etp', 'bias', 'et', 'sigma', 'status']
    return {'C3': t(p.Range('C3').Value), 'niveis': [dict(zip(cab, [x if not isinstance(x, dt.datetime) else str(x)
                                                                    for x in l])) for l in v]}


def por(qa, ws, addr, valor):
    """Filtro do Painel como o usuario: celula livre da aba protegida, eventos ligados."""
    try:
        C.por(qa, ws, addr, valor, local=False)
        return 'ok'
    except Exception as ex:  # noqa: BLE001
        return 'ERRO ' + repr(ex)[:120]


# ================================================================================================ OC-0040 / OC-0043
def cab_estat(qa):
    e, p = qa.ws('Estatística'), qa.ws('Painel')
    cab = {}
    for a in ('A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'A5', 'B5', 'C5', 'D5', 'G5', 'H5', 'I5', 'J5'):
        c = e.Range(a)
        cab[a] = {'f': t(c.Formula)[:140], 'txt': t(c.Text)[:120], 'locked': bool(c.Locked)}
    blob = ' '.join(t(e.Cells(r, c).Text) for r in range(1, 8) for c in range(1, 16))
    return {'cab': cab, 'tem_Ano_0': bool(re.search(r'\bAno 0\b', blob)) or t(e.Range('D4').Text) == '0',
            'Painel': {a: t(p.Range(a).Text) for a in ('C3', 'H3', 'L4', 'K3', 'K4', 'G3', 'G4', 'I3')}}


def ind_periodo(fin, an, lote, eq, ini, fim):
    out = {}
    for nv in (1, 2):
        vals = [float(r['RESULTADO']) for r in fin
                if t(r['ANALITO']) == an and t(r['LOTE']) == lote and up(r['EQUIPAMENTO']) == up(eq)
                and r['PARTICIPA_ESTATISTICA'] == 'SIM' and num(r['RESULTADO']) and int(r['NIVEL'] or 0) == nv
                and isinstance(r['DATA_HORA'], dt.datetime) and (ini is None or r['DATA_HORA'].date() >= ini)
                and (fim is None or r['DATA_HORA'].date() <= fim)]
        out[nv] = C.mdcv(vals) if vals else {'n': 0}
    return out


def comparar_estat(qa, ind, an):
    est = linhas_an(estat_all(qa), an)
    res = {}
    ok = True
    for nv in (1, 2):
        q = est.get(f'{an}|{nv}', {})
        i = ind[nv]
        d = {'qc': [q.get('C_n'), q.get('D_media'), q.get('E_dp'), q.get('F_cv')],
             'ind': [i.get('n'), i.get('media'), i.get('dp'), i.get('cv')]}
        d['dif_rel'] = [C.rel(a, b) if num(a) and num(b) else None for a, b in zip(d['qc'], d['ind'])]
        d['ok'] = (d['qc'][0] == d['ind'][0]) and all(x is None or x <= 1e-9 for x in d['dif_rel'][1:])
        ok = ok and d['ok']
        res[f'N{nv}'] = d
    res['ok'] = ok
    return res


def s_o40(qa, fin):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    out = {'Estatistica_protegida': bool(e.ProtectContents), 'Painel_protegido': bool(p.ProtectContents)}
    try:
        out['K4_lista'] = t(p.Range('K4').Validation.Formula1)[:300]
    except Exception as ex:  # noqa: BLE001
        out['K4_lista'] = 'sem validacao ' + repr(ex)[:60]
    out['0_aberto'] = cab_estat(qa)
    passos = {}
    for c, v in (('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8977")):
        passos[c] = por(qa, p, c, v)
    # mes
    passos['K3'] = por(qa, p, 'K3', 2026)
    passos['K4_Set'] = por(qa, p, 'K4', 'Set')
    time.sleep(1)
    m = cab_estat(qa)
    m['indep'] = comparar_estat(qa, ind_periodo(fin, 'Glicose', '8977', 'DIMENSION 1', dt.date(2026, 9, 1),
                                                dt.date(2026, 9, 30)), 'Glicose')
    m['painel_engPainel'] = qa.painel_stats()
    m['retrato'] = retrato(qa, 'S_mes')
    out['1_mes_Set2026'] = m
    # trimestre
    passos['K4_T3'] = por(qa, p, 'K4', 'T3')
    time.sleep(1)
    tr = cab_estat(qa)
    tr['indep'] = comparar_estat(qa, ind_periodo(fin, 'Glicose', '8977', 'DIMENSION 1', dt.date(2026, 7, 1),
                                                 dt.date(2026, 9, 30)), 'Glicose')
    tr['retrato'] = retrato(qa, 'S_trim')
    out['2_trim_T3_2026'] = tr
    # trimestre = dados brutos do trimestre, nao media dos meses
    meses = [ind_periodo(fin, 'Glicose', '8977', 'DIMENSION 1', dt.date(2026, k, 1),
                         (dt.date(2026, k + 1, 1) - dt.timedelta(days=1))) for k in (7, 8, 9)]
    out['2_trim_media_dos_meses_N1'] = st.fmean([x[1]['media'] for x in meses if x[1].get('n', 0) > 0]) \
        if any(x[1].get('n', 0) for x in meses) else None
    # Tudo (sem filtro)
    passos['K4_Tudo'] = por(qa, p, 'K4', 'Tudo (sem filtro)')
    time.sleep(1)
    tu = cab_estat(qa)
    tu['indep'] = comparar_estat(qa, ind_periodo(fin, 'Glicose', '8977', 'DIMENSION 1', None, None), 'Glicose')
    out['3_tudo'] = tu
    # vizinho: De/Ate atravessando o ano
    passos['G3'] = por(qa, p, 'G3', dt.datetime(2025, 12, 1))
    passos['G4'] = por(qa, p, 'G4', dt.datetime(2026, 1, 31))
    time.sleep(1)
    out['4_DeAte_2025_2026'] = cab_estat(qa)
    # vizinho: tentar escolher a Visao na Estatistica (o que o rotulo antigo sugeria)
    b4 = e.Range('B4')
    v0 = b4.Value
    try:
        b4.Value = 'MENSAL'
        out['5_B4_MENSAL'] = 'ACEITO (celula livre) -> restaurado'
        qa.ex.xl.EnableEvents = False
        b4.Value = v0
        qa.ex.xl.EnableEvents = True
    except Exception as ex:  # noqa: BLE001
        out['5_B4_MENSAL'] = 'RECUSADO: ' + repr(ex)[:160]
    # volta ao ano inteiro 2026
    passos['K3b'] = por(qa, p, 'K3', 2026)
    passos['K4_ano'] = por(qa, p, 'K4', 'Ano inteiro')
    out['6_ano_inteiro_2026'] = cab_estat(qa)
    out['passos'] = passos
    c = [out[k] for k in ('0_aberto', '1_mes_Set2026', '2_trim_T3_2026', '3_tudo', '4_DeAte_2025_2026',
                          '6_ano_inteiro_2026')]
    out['nenhum_Ano_0'] = not any(x['tem_Ano_0'] for x in c)
    return out


def s_o43(qa, fin):
    p = qa.ws('Painel')
    out = {}
    cel = {}
    curtas = ['REJEITADO', 'ALERTA', 'OK', 'DESATUAL.', 'BASE VAZIA', 'SEM MÉD/DP', 'AGUARDA 1ª', 'SEM CORRIDA',
              'PER. VAZIO', 'SEM PONTO']
    for a in ('J7', 'J8'):
        c = p.Range(a)
        ma = c.MergeArea
        larg = float(ma.Width) - 3.0
        f = c.DisplayFormat.Font
        med = {}
        for s in curtas:
            w = auditoria_abas.largura_pt(s, f.Name, bool(f.Bold), float(f.Size))
            med[s] = {'pt': round(w, 1), 'fator': round(min(1.0, larg / w), 3)}
        cel[a] = {'merge': t(ma.Address), 'largura_util_pt': round(larg, 1), 'fonte': [t(f.Name), bool(f.Bold),
                  float(f.Size)], 'shrink': bool(c.ShrinkToFit), 'wrap': bool(c.WrapText), 'medidas': med,
                  'formula': t(c.Formula)[:400]}
    out['celulas'] = cel
    out['pior_fator'] = min(m['fator'] for x in cel.values() for m in x['medidas'].values())
    # estados reais provocados pelo filtro
    est = {}

    def foto(rot):
        time.sleep(1)
        est[rot] = {'J7': t(p.Range('J7').Text), 'J8': t(p.Range('J8').Text), 'Q4': t(p.Range('Q4').Text),
                    'J7_cor': p.Range('J7').DisplayFormat.Interior.Color, 'A39': t(p.Range('A39').Text)[:200],
                    'filtro': {a: t(p.Range(a).Text) for a in ('C3', 'H3', 'L4', 'K3', 'K4')}}
    for c_, v in (('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8977"), ('K3', 2026), ('K4', 'Out')):
        por(qa, p, c_, v)
    foto('Glicose 8977 DIM1 Out/2026 (ciclo 01: REJEITADO)')
    por(qa, p, 'K3', 2025)
    por(qa, p, 'K4', 'Jan')
    foto('Glicose 8977 DIM1 Jan/2025 (periodo sem corrida)')
    por(qa, p, 'K4', 'Tudo (sem filtro)')
    por(qa, p, 'H3', "'6354")
    foto('Glicose lote 6354 (urina) Tudo')
    out['estados'] = est
    out['ok_medida'] = all(x['shrink'] for x in cel.values()) and out['pior_fator'] >= 0.70
    return out


# ================================================================================================ CEQ
def campos(prov='CAP', an=NOME, res='7,7', alvo='7', am='TA-F2-01', ano='2025', rod=ROD, un=UN):
    return [prov, rod, ano, an, am, res, alvo, '0,2', '1', '6,3', '7,7', 'Acceptable', '50', un, '1',
            f'{TAG}.pdf']


def base_teste(qa):
    b = qa.ws('EQA_Base')
    v = b.Range('A1:U5001').Value
    cab = [t(x) for x in v[0]]
    return [{k: (x if not isinstance(x, dt.datetime) else str(x)) for k, x in zip(cab, r) if x not in (None, '')}
            for r in v[1:] if any(TAG in t(x) for x in r)]


def s_ceq(qa, aud):
    out = {}
    e, p = qa.ws('Estatística'), qa.ws('Painel')
    esp = qa.espec()
    cat = esp.Range('A4:AR43').Value
    linha_an = next((r for r in cat if t(r[0]) == AN), None)
    out['especificacoes'] = {'AR_provedor': t(linha_an[43]) if linha_an else None,
                             'S_fonte': t(linha_an[18]) if linha_an else None,
                             'T_etp': linha_an[19] if linha_an else None}
    out['Estatistica_N4_anoEP'] = t(e.Range('N4').Value)
    out['Estatistica_P4_rodada'] = t(e.Range('P4').Value)
    ano = out['Estatistica_N4_anoEP'] if re.fullmatch(r'\d{4}', out['Estatistica_N4_anoEP']) else '2025'
    out['ano_usado'] = ano
    por(qa, p, 'C3', AN)
    qa.ex.xl.Calculate()
    E0 = estat_all(qa)
    out['S0'] = {'estat': linhas_an(E0, AN), 'painel': painel_an(qa), 'retrato': retrato(qa, 'S0')}
    aud.novos()
    # ---- incluir (a janela: Confirmacoes -> Salvar)
    out['aviso_analito'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', NOME))
    out['confirmacoes'] = t(qa.run('mCEQForm.ConfirmacoesCEQ', campos(ano=ano), 0))
    r1 = t(qa.run('mCEQForm.SalvarCEQ', campos(ano=ano)))
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    E1a = estat_all(qa)
    qa.ex.xl.Calculate()
    E1 = estat_all(qa)
    cap = qa.ws('EQA.CAP_Dados')
    out['S1'] = {'salvar': r1, 'linha': lin, 'gravado_F_G': [cap.Range(f'F{lin}').Value, cap.Range(f'G{lin}').Value]
                 if lin else None, 'descricao': t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', lin)) if lin else '',
                 'lista_rodada': [[t(x) for x in l] for l in (qa.run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, ano)
                                                               or [])],
                 'base_teste': base_teste(qa), 'Z3': t(cap.Range('Z3').Value)[:160],
                 'estat_sem_Calculate_igual_com': E1a == E1,
                 'estat': linhas_an(E1, AN), 'painel': painel_an(qa), 'diff_S0': diff_estat(E0, E1),
                 'trilha': [x['txt'][:400] for x in aud.novos()], 'retrato': retrato(qa, 'S1')}
    # ---- corrigir (a mesma linha; 7,35 = +5 %)
    r2 = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(ano=ano, res='7,35'))) if lin else 'sem linha'
    qa.ex.xl.Calculate()
    E2 = estat_all(qa)
    out['S2'] = {'corrigir': r2, 'gravado_F': cap.Range(f'F{lin}').Value if lin else None,
                 'base_teste': base_teste(qa), 'estat': linhas_an(E2, AN), 'painel': painel_an(qa),
                 'diff_S1': diff_estat(E1, E2), 'trilha': [x['txt'][:500] for x in aud.novos()],
                 'retrato': retrato(qa, 'S2')}
    # ---- vizinhos: o mesmo analito no Controllab e no CAP de outro ano
    extra = []
    rc = t(qa.run('mCEQForm.SalvarCEQ', campos(prov='Controllab', res='14', am='TA-F2-CTL')))
    if rc.startswith('OK|'):
        extra.append(('Controllab', int(rc.split('|')[1])))
    ano_out = str(int(ano) - 1)
    ra = t(qa.run('mCEQForm.SalvarCEQ', campos(res='14', am='TA-F2-ANO', ano=ano_out)))
    if ra.startswith('OK|'):
        extra.append(('CAP', int(ra.split('|')[1])))
    qa.ex.xl.Calculate()
    E3 = estat_all(qa)
    out['S3_vizinhos'] = {'controllab': rc, 'cap_outro_ano': [ano_out, ra], 'base_teste': base_teste(qa),
                          'diff_S2': diff_estat(E2, E3), 'estat': linhas_an(E3, AN),
                          'trilha': [x['txt'][:300] for x in aud.novos()], 'retrato': retrato(qa, 'S3')}
    # ---- duplicidade
    out['duplicidade'] = t(qa.run('mCEQForm.SalvarCEQ', campos(ano=ano, res='7,35')))[:400]
    # ---- excluir tudo
    exc = {}
    for prov, ln in reversed(extra):
        exc[f'{prov}|{ln}'] = t(qa.run('mCEQForm.ExcluirCEQ', prov, ln))[:160]
    exc[f'CAP|{lin}'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    qa.ex.xl.Calculate()
    E4 = estat_all(qa)
    out['S4_excluido'] = {'excluir': exc, 'base_teste': base_teste(qa), 'diff_S0': diff_estat(E0, E4),
                          'igual_S0_bit_a_bit': E0 == E4, 'painel': painel_an(qa),
                          'trilha': [x['txt'][:300] for x in aud.novos()], 'retrato': retrato(qa, 'S4')}
    return out


# ================================================================================================ manuais + ATUALIZAR
def oraculo_depara(qa, fin):
    dp = qa.ler('tblDeParaAnalitos')
    regras = [r for r in qa.ler('tblLotesMaterial') if t(r.get('LOTE'))]
    mz_lote = {t(r['LOTE']): up(r['MATRIZ']) for r in regras if up(r['REGRA']) == 'MATRIZ'}
    cod = collections.defaultdict(set)
    tem_mz = set()
    for d in dp:
        if t(d.get('ANALITO_ORIGEM')):
            if t(d.get('ANALITO_QCINI')):
                cod[t(d['ANALITO_QCINI'])].add(t(d['ANALITO_ORIGEM']))
                # ADR-093 cond. 2 (mEtlCQ ParaQcini): "tem de/para" = linha da matriz com ANALITO_QCINI preenchido;
                # a linha "sem cadastro" (QCINI vazio, caso da URINA) nao conta
                tem_mz.add((up(d.get('MATRIZ')), t(d['ANALITO_ORIGEM'])))
    recebido = {(up(r['MATRIZ']), t(r['ANALITO_ORIGEM'])) for r in fin if r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'}

    def previsto(lote, an):
        mz = mz_lote.get(lote)
        if not mz:
            return {'regra_MATRIZ': False}
        cs = cod.get(an, set())
        if len(cs) != 1:
            return {'matriz': mz, 'codigos': sorted(cs), 'aceito': False, 'motivo': 'ambiguo ou sem codigo'}
        c = next(iter(cs))
        if (mz, c) in tem_mz:
            return {'matriz': mz, 'codigo': c, 'aceito': False, 'motivo': 'codigo ja tem de/para na matriz'}
        if (mz, c) not in recebido:
            return {'matriz': mz, 'codigo': c, 'aceito': False, 'motivo': 'nunca recebido nessa matriz'}
        return {'matriz': mz, 'codigo': c, 'aceito': True}
    return previsto, regras, cod, recebido


def s_manuais(qa, fin, aud, L):
    out = {}
    p = qa.ws('Painel')
    previsto, regras, cod, recebido = oraculo_depara(qa, fin)
    out['regras_tblLotesMaterial'] = [{k: t(v) for k, v in r.items() if t(v)} for r in regras]
    esp = qa.espec()
    catalogo = [t(x[0]) for x in esp.Range('A4:A43').Value if t(x[0])]
    # escolha do analito recusado (codigo unico, nunca recebido na urina) e do ambiguo, pelo oraculo
    lote_u = '6354'
    cand = {a: previsto(lote_u, a) for a in catalogo}
    out['oraculo_6354'] = {a: v for a, v in cand.items() if v.get('aceito') or a in ('Creatinina', 'Glicose')}
    recus = next((a for a, v in cand.items() if not v.get('aceito') and v.get('motivo') == 'nunca recebido nessa matriz'
                  and a not in ('Glicose',)), None)
    ambig = next((a for a, v in cand.items() if v.get('motivo') == 'ambiguo ou sem codigo' and len(v.get('codigos', []))
                  > 1), None)
    out['escolhidos'] = {'aceito': 'Creatinina', 'recusado_nunca_recebido': recus, 'ambiguo': ambig}
    por(qa, p, 'L4', 'DIMENSION 1')
    por(qa, p, 'C3', 'Glicose')
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)[:120]
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == 'Glicose' and int(r['NIVEL'] or 0) == 1 and t(r['LOTE']) == '8977'
                    and isinstance(r['DATA_HORA'], dt.datetime) and t(r['EQUIPAMENTO']) == 'DIMENSION 1'],
                   key=lambda r: r['DATA_HORA'])
    R = reais[-25]
    dia = R['DATA_HORA'].strftime('%d/%m/%Y')
    out['dia_dos_manuais'] = dia
    qa.ex.xl.EnableEvents = True
    linhas = {}

    def novo(rot, lote, an, res, hora, nivel='1'):
        _, r = C.linha_livre_manual(qa)
        for col, v in (('DATA', dia), ('HORA', hora), ('EQUIPAMENTO', 'DIMENSION 1'), ('LOTE', "'" + lote),
                       ('NIVEL', nivel), ('ANALITO', an), ('RESULTADO', res), ('MOTIVO', MOTIVO)):
            V.escrever_celula(qa, r, col, v)
        linhas[rot] = r
    novo('A_controle_Glicose_8977', '8977', 'Glicose', '123,456', '04:11')
    novo('B_6354_Creatinina', lote_u, 'Creatinina', '12,345', '04:13')
    if recus:
        novo('C_6354_recusado', lote_u, recus, '1,234', '04:15')
    if ambig:
        novo('C2_6354_ambiguo', lote_u, ambig, '1,234', '04:16')
    novo('D_implausivel_8977', '8977', 'Glicose', '-1', '04:17')
    novo('E_8794_fora_do_cq', '8794', 'Glicose', '50', '04:19')
    novo('F_8353_Glicose_so_analitos', '8353', 'Glicose', '123,456', '04:21')
    novo('G_8353_PCR', '8353', 'Proteína C-reativa', '1,234', '04:23')
    # ---- inativar um ponto real e comentar outro (nucleo da janela do Painel)
    alvo_inat = reais[-40]
    alvo_com = reais[-41]
    out['inativar'] = {'id': t(alvo_inat['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.InativarPorId', t(alvo_inat['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: inativação de teste (F2)', True, '', teto=120))[:200]}
    out['comentar'] = {'id': t(alvo_com['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.GravarComentarioTecnico', t(alvo_com['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: comentário de teste (F2)', 'NÃO', teto=120))[:200]}
    aud.novos()
    sync0 = qa.lo('tblSyncLog').ListRows.Count
    vg = V.Vigia(qa.ex.pid, teto=600)
    vg.start()
    t0 = time.time()
    try:
        r_at = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    finally:
        vg.parar.set()
    out['atualizar1'] = {'ret': r_at[:600], 's': round(time.time() - t0, 1), 'caixas': vg.vistos[:5],
                         'sync_linhas': [sync0, qa.lo('tblSyncLog').ListRows.Count]}
    tr1 = aud.novos()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in leit.items()}
    out['ids'] = ids
    finid = qa.final_por_id()
    qai = qa.ler('tblQA_Integracao')
    res = {}
    for k, v in leit.items():
        i = ids[k]
        f = finid.get(i, {})
        res[k] = {'STATUS_aba': v['STATUS'][0], 'DETALHE_aba': v['DETALHE'][1][:260], 'RUN_aba': v['RUN'][1],
                  'final': {c: t(f.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ANALITO',
                                                     'ANALITO_ORIGEM', 'ANALITO_QCINI', 'BLOCO', 'MATRIZ',
                                                     'MOTIVO_EXCLUSAO_AUTOMATICA', 'REGISTRAR_RESULTADO_NO_LJ')},
                  'QA': [f"{t(x.get('CODIGO'))}|{t(x.get('SEVERIDADE'))}|"
                         + ' | '.join(t(z) for z in x.values() if t(z))[:220]
                         for x in qai if any(t(z) == i for z in x.values())],
                  'trilha': [x['acao'] for x in tr1 if x['acao'] and re.search(re.escape(i) + r'\b', x['txt'])],
                  'trilha_txt': [re.sub(r'\b[A-Z0-9_.-]*\\[^ |]*', '<usr>', x['txt'])[:420] for x in tr1
                                 if x['acao'] and re.search(re.escape(i) + r'\b', x['txt'])]}
    out['resultado'] = res
    out['previsto'] = {'B_6354_Creatinina': previsto(lote_u, 'Creatinina'),
                       'C_6354_recusado': previsto(lote_u, recus) if recus else None,
                       'C2_6354_ambiguo': previsto(lote_u, ambig) if ambig else None}
    ci = {c: t(finid.get(out['inativar']['id'], {}).get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
                                                                       'DATA_INATIVACAO', 'MOTIVO_INATIVACAO')}
    out['inativado_no_final'] = ci
    out['trilha_outros'] = [x['acao'] for x in tr1 if x['acao'] and not any(
        re.search(re.escape(i) + r'\b', x['txt']) for i in ids.values() if i)]
    # 2o ATUALIZAR (so entradas): nenhuma trilha de manual repetida
    t0 = time.time()
    r2 = t(qa.run('mIntegracao.AtualizarSoEntradas', teto=1500))
    tr2 = aud.novos()
    out['atualizar2'] = {'ret': r2[:200], 's': round(time.time() - t0, 1),
                         'trilha_manual_repetida': [x['acao'] for x in tr2 if x['acao'] and 'MANUAL' in x['acao']]}
    # o manual de urina entra na Creatinina do soro? (Painel/Estatistica)
    L['_ids_manuais'] = ids
    return out


# ================================================================================================ ADR-091
def s_adr091(qa):
    out = {}
    fin = qa.ler('tblCQ_Final')
    regras = [r for r in qa.ler('tblLotesMaterial') if t(r.get('LOTE'))]
    esp = qa.espec()
    catalogo = {up(x[0]) for x in esp.Range('A4:A43').Value if t(x[0])}
    por_lote = collections.defaultdict(list)
    for r in fin:
        por_lote[t(r['LOTE'])].append(r)
    precede = set(V.PRECEDE_LOTE) | {'IMPLAUSIVEL'}
    det, viol = {}, []
    for rg in regras:
        lt, tipo = t(rg['LOTE']), up(rg['REGRA'])
        rs = por_lote.get(lt, [])
        c = collections.Counter(r['STATUS_ANALITICO'] for r in rs)
        if tipo == 'FORA_DO_CQ':
            bad = [r['ID_REGISTRO'] for r in rs if r['STATUS_ANALITICO'] != 'LOTE_FORA_DO_CQ'
                   and r['STATUS_ANALITICO'] not in precede]
            part = [r['ID_REGISTRO'] for r in rs if r['PARTICIPA_ESTATISTICA'] == 'SIM']
            if bad or part:
                viol.append({'lote': lt, 'regra': tipo, 'status_errado': bad[:5], 'participa': part[:5]})
        elif tipo == 'SO_ANALITOS':
            lista = {up(a) for a in t(rg.get('ANALITOS')).split(';') if a.strip()}
            fora = [r for r in rs if up(r['ANALITO']) not in lista]
            dentro = [r for r in rs if up(r['ANALITO']) in lista]
            bad = [r['ID_REGISTRO'] for r in fora if r['STATUS_ANALITICO'] != 'LOTE_OUTRO_MATERIAL'
                   and r['STATUS_ANALITICO'] not in precede]
            bad += [r['ID_REGISTRO'] for r in fora if r['PARTICIPA_ESTATISTICA'] == 'SIM']
            bad2 = [r['ID_REGISTRO'] for r in dentro if r['STATUS_ANALITICO'].startswith('LOTE_')]
            if bad or bad2:
                viol.append({'lote': lt, 'regra': tipo, 'fora_da_lista_errado': bad[:5], 'da_lista_excluido': bad2[:5]})
            c = {'dentro': dict(collections.Counter(r['STATUS_ANALITICO'] for r in dentro)),
                 'fora': dict(collections.Counter(r['STATUS_ANALITICO'] for r in fora))}
        elif tipo == 'MATRIZ':
            mz = up(rg['MATRIZ'])
            bad = [r['ID_REGISTRO'] for r in rs if up(r['MATRIZ']) != mz]
            ser = [r['ID_REGISTRO'] for r in rs if up(r['ANALITO']) in catalogo and r['PARTICIPA_ESTATISTICA'] == 'SIM']
            if bad or ser:
                viol.append({'lote': lt, 'regra': tipo, 'matriz_errada': bad[:5], 'analito_serico_participa': ser[:5]})
            c = {'status': dict(c), 'matriz': dict(collections.Counter(up(r['MATRIZ']) for r in rs)),
                 'analitos': dict(collections.Counter(t(r['ANALITO']) for r in rs).most_common(6))}
        det[f'{tipo} {lt}'] = {'linhas': len(rs), 'resumo': c if isinstance(c, dict) else dict(c)}
    out['regras'] = len(regras)
    out['por_regra'] = det
    out['violacoes'] = viol
    # Painel: 8353 (SO_ANALITOS) -- PCR participa, Glicose nao
    p = qa.ws('Painel')
    pain = {}
    for an in ('Proteína C-reativa', 'Glicose'):
        por(qa, p, 'L4', 'DIMENSION 1')
        por(qa, p, 'C3', an)
        por(qa, p, 'H3', "'8353")
        por(qa, p, 'K4', 'Tudo (sem filtro)')
        time.sleep(1)
        indep = collections.Counter(int(r['NIVEL'] or 0) for r in fin if t(r['LOTE']) == '8353' and t(r['ANALITO']) == an
                                    and up(r['EQUIPAMENTO']) == 'DIMENSION 1' and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                                    and num(r['RESULTADO']))
        pain[an] = {'H3': t(p.Range('H3').Text), 'C3': t(p.Range('C3').Text), 'L4': t(p.Range('L4').Text),
                    'H4_aviso': t(p.Range('H4').Text)[:160], 'K4': t(p.Range('K4').Text),
                    'J7_J8': [t(p.Range('J7').Text), t(p.Range('J8').Text)], 'engPainel': qa.painel_stats(),
                    'E1_lote_motor': t(qa.ws('Eng_Saida').Range('E1').Value), 'indep_participa_por_nivel': dict(indep)}
    out['painel_8353'] = pain
    # Principal: os status chegam a aba do usuario (tblCQ_Final) -- contagem por status nos lotes com regra
    pr = qa.ws('Principal - Resultados')
    out['Principal_A5'] = t(pr.Range('A5').Text)[:220]
    out['status_LOTE_total'] = dict(collections.Counter(r['STATUS_ANALITICO'] for r in fin
                                                        if r['STATUS_ANALITICO'].startswith('LOTE_')))
    out['ok'] = not viol
    return out


# ================================================================================================ ADR-094
RX_DH = re.compile(r'^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$')
RX_D = re.compile(r'^\d{2}/\d{2}/\d{4}$')


def s_adr094(qa, ids_extra):
    out = {'tabelas': {}, 'hashes': [], 'sem_segundos': [], 'formato_ruim': []}
    for ws in qa.wb.Worksheets:
        for lo in ws.ListObjects:
            try:
                n = lo.ListRows.Count
            except Exception:
                continue
            if n == 0 or lo.DataBodyRange is None:
                continue
            nome = t(lo.Name)
            info = {}
            for lc in lo.ListColumns:
                cn = t(lc.Name)
                rng = lc.DataBodyRange
                vals = rng.Value
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                idx = [i for i, x in enumerate(vals) if x[0] not in (None, '')]
                if not idx:
                    # coluna vazia: a celula nao tem texto -- nada a conferir
                    continue
                datas = [i for i in idx if isinstance(vals[i][0], dt.datetime)]
                amostra = sorted(set([idx[0], idx[-1]] + idx[::max(1, len(idx) // 40)][:40]))
                if datas:
                    amostra = sorted(set(amostra + [datas[0], datas[-1]] + datas[::max(1, len(datas) // 40)][:40]))
                ruins = []
                for i in amostra:
                    c = rng.Cells(i + 1, 1)
                    tx = t(c.Text)
                    if tx and set(tx) <= {'#'}:
                        out['hashes'].append(f'{ws.Name}!{nome}[{cn}] linha {i + 1}: "{tx}" (largura {c.ColumnWidth})')
                        ruins.append(i)
                        continue
                    v = vals[i][0]
                    if isinstance(v, dt.datetime):
                        vv = v.replace(tzinfo=None)
                        tem_hora = (vv.hour, vv.minute, vv.second) != (0, 0, 0)
                        if tem_hora and not RX_DH.match(tx) and cn.upper() not in ('HORA',):
                            out['sem_segundos'].append(f'{ws.Name}!{nome}[{cn}] linha {i + 1}: "{tx}" fmt '
                                                       f'"{t(c.NumberFormat)}"')
                        elif not tem_hora and not (RX_D.match(tx) or RX_DH.match(tx)) and cn.upper() != 'HORA':
                            out['formato_ruim'].append(f'{ws.Name}!{nome}[{cn}] linha {i + 1}: "{tx}" fmt '
                                                       f'"{t(c.NumberFormat)}"')
                if datas:
                    info[cn] = {'n_datas': len(datas), 'amostra': len(amostra),
                                'fmt': t(rng.Cells(datas[-1] + 1, 1).NumberFormat),
                                'ultimo_txt': t(rng.Cells(datas[-1] + 1, 1).Text)}
            if info:
                out['tabelas'][f'{ws.Name}!{nome}'] = info
    # linhas novas desta sessao
    novas = {}
    lo = qa.lo('tblSyncLog')
    cab = [t(c.Name) for c in lo.ListColumns]
    ult = lo.ListRows.Count
    novas['Sync_Log_ultima'] = {c: t(lo.ListColumns(c).DataBodyRange.Cells(ult, 1).Text) for c in cab
                                if c in ('INICIO', 'FIM', 'DATA_MIN', 'DATA_MAX', 'MODO', 'RESULTADO', 'STATUS')}
    fin = qa.lo('tblCQ_Final')
    idc = fin.ListColumns('ID_REGISTRO').DataBodyRange.Value
    pos = {t(x[0]): i + 1 for i, x in enumerate(idc)}
    for rot, i in ids_extra.items():
        if i in pos:
            novas[f'tblCQ_Final {rot} {i}'] = {c: t(fin.ListColumns(c).DataBodyRange.Cells(pos[i], 1).Text)
                                               for c in ('DATA', 'DATA_HORA', 'RECEBIDO_EM', 'DATA_INATIVACAO')}
    for tb, col in (('tblInativacao_NaoConformes', 'DATA_INATIVACAO'), ('tblComentariosTecnicos', 'DATA'),
                    ('tblResultados_Manuais', 'REGISTRADO_EM')):
        try:
            l2 = qa.lo(tb)
            v = l2.ListColumns(col).DataBodyRange.Value
            k = max((i for i, x in enumerate(v) if x[0] not in (None, '')), default=None)
            if k is not None:
                c = l2.ListColumns(col).DataBodyRange.Cells(k + 1, 1)
                novas[f'{tb}[{col}] ultima'] = {'txt': t(c.Text), 'valor': str(C.limpo(v[k][0])),
                                                'largura': c.ColumnWidth}
        except Exception as ex:  # noqa: BLE001
            novas[f'{tb}[{col}]'] = 'ERRO ' + repr(ex)[:80]
    # DB A5 e Inicio
    for aba, cel in (('DB', 'A5'), ('Início', None)):
        try:
            w = qa.ws(aba)
            if cel:
                novas[f'{aba}!{cel}'] = t(w.Range(cel).Text)[:200]
            else:
                ur = w.UsedRange.Value
                txt = [t(x) for l in ur for x in l if 'último' in t(x).lower() or 'ultimo' in t(x).lower()]
                novas[f'{aba} ultimo'] = txt[:4]
                hs = []
                for c in w.UsedRange.Cells:
                    tx = t(c.Text)
                    if tx and set(tx) <= {'#'}:
                        hs.append(t(c.Address))
                novas[f'{aba} ####'] = hs[:10]
        except Exception as ex:  # noqa: BLE001
            novas[aba] = 'ERRO ' + repr(ex)[:80]
    out['linhas_novas'] = novas
    out['ok'] = not out['hashes'] and not out['sem_segundos'] and not out['formato_ruim']
    return out


# ================================================================================================ OC-0057
def s_sem_sessao(qa, aud):
    out = {}
    r = t(qa.run('mCEQForm.SalvarCEQ', campos(am='TA-F2-SES', res='7,7')))
    lin = int(r.split('|')[1]) if r.startswith('OK|') else 0
    out['registro_com_sessao'] = r[:120]
    cap = qa.ws('EQA.CAP_Dados')
    qa.run('mSeguranca.Logout')
    qa.ex.esperar()
    out['SessaoAtiva'] = qa.run('mSeguranca.SessaoAtiva')
    aud.novos()
    f0 = cap.Range(f'F{lin}').Value if lin else None
    out['salvar'] = t(qa.run('mCEQForm.SalvarCEQ', campos(am='TA-F2-SES2')))[:160]
    out['corrigir'] = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(am='TA-F2-SES', res='7,0')))[:160] if lin else ''
    out['F_antes_depois'] = [f0, cap.Range(f'F{lin}').Value if lin else None]
    out['excluir'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    out['A_depois'] = t(cap.Range(f'A{lin}').Value) if lin else ''
    out['trilha'] = [x['acao'] for x in aud.novos()]
    out['abrir_sem_login'] = V.abrir_ceq_sem_sessao(qa)
    # vizinho: login de novo -> Corrigir e Excluir voltam a funcionar
    out['relogin'] = fazer_login(qa, USR)
    aud.novos()
    out['corrigir_com_sessao'] = t(qa.run('mCEQForm.CorrigirCEQ', lin, campos(am='TA-F2-SES', res='7,0')))[:160] \
        if lin else ''
    out['F_apos_corrigir'] = cap.Range(f'F{lin}').Value if lin else None
    out['excluir_com_sessao'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160] if lin else ''
    out['trilha_com_sessao'] = [x['acao'] for x in aud.novos()]
    out['ok'] = (out['salvar'].startswith('ERRO|0|') and out['corrigir'].startswith('ERRO|0|')
                 and out['excluir'].startswith('ERRO|0|') and out['F_antes_depois'][0] == out['F_antes_depois'][1]
                 and not [a for a in out['trilha'] if a and a.startswith('CEQ_')] and out['abrir_sem_login']['ok'])
    out['ok_vizinho'] = (out['corrigir_com_sessao'].startswith('OK|') and out['excluir_com_sessao'].startswith('OK|'))
    return out


# ================================================================================================ calc
def calc(nomes=('S_mes', 'S_trim', 'S0', 'S1', 'S2', 'S3', 'S4'), analitos=(AN, 'Glicose'), saida='calc'):
    import calc_independente as CI
    res = {}
    for nome in nomes:
        cam = os.path.join(C.PASTA, f'{PROD}_F2_{nome}.xlsm')
        if not os.path.exists(cam):
            continue
        d, c, meta, linhas = CI.auditar(cam, 1e-9)
        dif = [{'linha': L_['linha'], 'analito': L_['analito'], 'nivel': L_['nivel'], 'col': k, 'qc': x['qc'],
                'ind': x['ind'], 'dr': x['dif_rel']} for L_ in linhas for k, x in L_['comp'].items()
               if x['status'] != 'OK']
        tot = collections.Counter('OK' if x['status'] == 'OK' else 'DIF' for L_ in linhas for x in L_['comp'].values())
        alvo = {f"{L_['analito']}|{L_['nivel']}": {k: x['qc'] for k, x in L_['comp'].items()}
                for L_ in linhas if L_['analito'] in analitos}
        bias = {}
        for a in analitos:
            b = c.bias_ceq(a)
            bias[a] = ({'abs': b['abs'], 'signed': b['signed'], 'n': b['n'], 'nrod': b['nrod'],
                        'rodadas': b['rodadas']} if b else None)
        res[nome] = {'meta': {k: v for k, v in meta.items() if k in ('lote', 'equipamento', 'visao', 'janela',
                                                                       'janela_cache', 'ano_ep', 'rodada',
                                                                       'janela_sigma')},
                     'totais': dict(tot), 'divergencias': dif[:40], 'n_div': len(dif), 'linhas_alvo': alvo,
                     'bias_ind': bias, 'colunas_comparadas': sorted({k for L_ in linhas for k in L_['comp']})}
        json.dump(res[nome], open(os.path.join(C.PASTA, f'{saida}_{nome}.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
        print('calc', nome, dict(tot), flush=True)
    return res


# ================================================================================================ parte 2
def s_ceq_glicose(qa, aud, senha):
    """Bias, ET e Sigma de um analito COM CIQ no periodo (o ET precisa do CV do periodo) + provedor."""
    out = {}
    e, p = qa.ws('Estatística'), qa.ws('Painel')
    an = 'Glicose'
    for c_, v in (('L4', 'DIMENSION 1'), ('C3', an), ('H3', "'8977"), ('K3', 2026), ('K4', 'Ano inteiro')):
        por(qa, p, c_, v)
    esp = qa.espec()
    lin_an = next((r for r in esp.Range('A4:AR43').Value if t(r[0]) == an), None)
    out['especificacoes'] = {'AR_provedor': t(lin_an[43]), 'S_fonte': t(lin_an[18]), 'T_etp': lin_an[19]}
    b = qa.ws('EQA_Base')
    mapa = [(t(m[0]), t(m[1]), t(m[2])) for m in C.bloco(b, 'W2:Y600') if t(m[1])]
    nome_cap = next((nm for pv, nm, cn in mapa if up(pv) == 'CAP' and cn == an), None)
    out['nome_CAP'] = nome_cap
    out['mapa_Controllab'] = [(nm, cn) for pv, nm, cn in mapa if up(pv) == 'CONTROLLAB'][:10]
    ano = t(e.Range('N4').Value)
    ano = ano if re.fullmatch(r'\d{4}', ano) else '2025'
    out['ano'] = ano
    rod = 'TESTE-AUDITORIA F2 R2'

    def cg(prov='CAP', nm=nome_cap, res='110', am='TA-F2-G1'):
        return [prov, rod, ano, nm, am, res, '100', '2', '1', '90', '110', 'Acceptable', '50', 'mg/dL', '1',
                f'{TAG}.pdf']
    qa.ex.xl.Calculate()
    E0 = estat_all(qa)
    out['G0'] = {'estat': linhas_an(E0, an), 'painel': painel_an(qa), 'retrato': retrato(qa, 'G0')}
    aud.novos()
    r1 = t(qa.run('mCEQForm.SalvarCEQ', cg()))
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    qa.ex.xl.Calculate()
    E1 = estat_all(qa)
    out['G1'] = {'salvar': r1, 'estat': linhas_an(E1, an), 'painel': painel_an(qa), 'diff_G0': diff_estat(E0, E1),
                 'trilha': [x['acao'] for x in aud.novos()], 'retrato': retrato(qa, 'G1')}
    r2 = t(qa.run('mCEQForm.CorrigirCEQ', lin, cg(res='105'))) if lin else 'sem linha'
    qa.ex.xl.Calculate()
    E2 = estat_all(qa)
    out['G2'] = {'corrigir': r2, 'estat': linhas_an(E2, an), 'painel': painel_an(qa), 'diff_G1': diff_estat(E1, E2),
                 'trilha': [x['acao'] for x in aud.novos()], 'retrato': retrato(qa, 'G2')}
    # provedor: um nome do Controllab mapeado para Glicose (linha de mapa TESTE-AUDITORIA na EQA_Base W:Y da copia)
    b.Unprotect(senha)
    v = b.Range('W2:Y600').Value
    livre = next(i for i, l in enumerate(v, 2) if not any(l))
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre}').Value = (('Controllab', 'TESTE-AUDITORIA GLU CTL', an),)
    finally:
        qa.ex.xl.EnableEvents = ev
    out['mapa_teste_linha'] = livre
    out['aviso_ctl'] = t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'Controllab', 'TESTE-AUDITORIA GLU CTL'))
    rc = t(qa.run('mCEQForm.SalvarCEQ', cg(prov='Controllab', nm='TESTE-AUDITORIA GLU CTL', res='200', am='TA-F2-GC')))
    lin_c = int(rc.split('|')[1]) if rc.startswith('OK|') else 0
    qa.ex.xl.Calculate()
    E3 = estat_all(qa)
    out['G3_controllab'] = {'salvar': rc, 'base_teste': base_teste(qa), 'diff_G2': diff_estat(E2, E3),
                            'estat': linhas_an(E3, an), 'trilha': [x['acao'] for x in aud.novos()],
                            'retrato': retrato(qa, 'G3')}
    exc = {}
    if lin_c:
        exc[f'Controllab|{lin_c}'] = t(qa.run('mCEQForm.ExcluirCEQ', 'Controllab', lin_c))[:120]
    if lin:
        exc[f'CAP|{lin}'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:120]
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre}').ClearContents()
    finally:
        qa.ex.xl.EnableEvents = ev
    qa.ex.xl.Calculate()
    E4 = estat_all(qa)
    out['G4_excluido'] = {'excluir': exc, 'base_teste': base_teste(qa), 'diff_G0': diff_estat(E0, E4),
                          'igual_G0_bit_a_bit': E0 == E4, 'retrato': retrato(qa, 'G4')}
    return out


def s_atualizar_p2(qa, fin, aud):
    out = {}
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)[:120]
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == 'Glicose' and int(r['NIVEL'] or 0) == 1 and t(r['LOTE']) == '8977'
                    and isinstance(r['DATA_HORA'], dt.datetime) and t(r['EQUIPAMENTO']) == 'DIMENSION 1'],
                   key=lambda r: r['DATA_HORA'])
    R = reais[-25]
    dia = R['DATA_HORA'].strftime('%d/%m/%Y')
    qa.ex.xl.EnableEvents = True
    _, r = C.linha_livre_manual(qa)
    for col, v in (('DATA', dia), ('HORA', '04:11'), ('EQUIPAMENTO', 'DIMENSION 1'), ('LOTE', "'8977"), ('NIVEL', '1'),
                   ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)):
        V.escrever_celula(qa, r, col, v)
    out['inativar'] = {'id': t(reais[-40]['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.InativarPorId', t(reais[-40]['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: inativação de teste (F2)', True, '', teto=120))[:200]}
    out['comentar'] = {'id': t(reais[-41]['ID_REGISTRO']),
                       'ret': t(qa.run('mIntegracao.GravarComentarioTecnico', t(reais[-41]['ID_REGISTRO']),
                                       'TESTE-AUDITORIA: comentário de teste (F2)', 'NÃO', teto=120))[:200]}
    t0 = time.time()
    out['atualizar'] = {'ret': t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))[:300],
                        's': round(time.time() - t0, 1)}
    out['manual'] = V.ler_manual(qa, r)['ID_REGISTRO'][0]
    aud.novos()
    return out


def s_painel_lotes(qa):
    p = qa.ws('Painel')
    out = {}
    try:
        out['C3_lista'] = t(p.Range('C3').Validation.Formula1)[:200]
    except Exception as ex:  # noqa: BLE001
        out['C3_lista'] = 'sem validacao ' + repr(ex)[:60]

    def foto():
        time.sleep(1)
        d = {a: t(p.Range(a).Text)[:220] for a in ('C3', 'H3', 'L4', 'K4', 'H4', 'J7', 'J8', 'Q4', 'A39')}
        d['engPainel'] = qa.painel_stats()
        d['M1_cortadas'] = t(qa.ws('Eng_Saida').Range('M1').Value)
        return d
    por(qa, p, 'L4', 'DIMENSION 1')
    por(qa, p, 'K4', 'Tudo (sem filtro)')
    por(qa, p, 'C3', 'Proteína C-reativa')
    por(qa, p, 'H3', "'8353")
    out['1_PCR_8353'] = foto()
    por(qa, p, 'C3', 'Glicose')
    out['2_Glicose_com_8353'] = foto()
    por(qa, p, 'H3', "'8794")
    out['3_lote_8794_FORA_DO_CQ'] = foto()
    por(qa, p, 'C3', 'Creatinina')
    por(qa, p, 'H3', "'6354")
    out['4_Creatinina_6354_URINA'] = foto()
    por(qa, p, 'H3', "'8977")
    por(qa, p, 'C3', 'Glicose')
    out['5_volta_Glicose_8977'] = foto()
    return out


def parte2():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F2_p2.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F2_p2.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        aud = V.Audit(qa)
        secao(L, 'ceq_glicose', s_ceq_glicose, qa, aud, senha)
        at = secao(L, 'atualizar_p2', s_atualizar_p2, qa, fin, aud)
        extra = {}
        if isinstance(at, dict) and at.get('inativar'):
            extra = {'manual': at.get('manual'), 'inativado': at['inativar']['id'], 'comentado': at['comentar']['id']}
        secao(L, 'adr094', s_adr094, qa, extra)
        secao(L, 'painel_lotes', s_painel_lotes, qa)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    del senha
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM excel', flush=True)


def main():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F2.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F2.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        secao(L, 'O40', s_o40, qa, fin)
        secao(L, 'O43', s_o43, qa, fin)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        aud = V.Audit(qa)
        secao(L, 'ceq', s_ceq, qa, aud)
        man = secao(L, 'manuais', s_manuais, qa, fin, aud, L)
        secao(L, 'adr091', s_adr091, qa)
        extra = dict(L.R.get('_ids_manuais', {}))
        if isinstance(man, dict) and man.get('inativar'):
            extra['inativado'] = man['inativar']['id']
            extra['comentado'] = man['comentar']['id']
        secao(L, 'adr094', s_adr094, qa, extra)
        secao(L, 'O57_sem_sessao', s_sem_sessao, qa, aud)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    del senha
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM excel', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    if '--parte2' in sys.argv:
        parte2()
        r = calc(('G0', 'G1', 'G2', 'G3', 'G4'), ('Glicose', AN), 'calcG')
        json.dump(r, open(os.path.join(C.PASTA, f'{PROD}_F2_calcG.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
    else:
        if '--so-calc' not in sys.argv:
            main()
        r = calc()
        json.dump(r, open(os.path.join(C.PASTA, f'{PROD}_F2_calc.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1, default=str)
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
