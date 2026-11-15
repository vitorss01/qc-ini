# -*- coding: utf-8 -*-
"""reteste04_usuario_f1.py -- Usuario Simulado, ciclo 2026-11-14_reteste04, frente F1, modo RETESTE (Bioquimica).

Refaz o fluxo ORIGINAL das ocorrencias OC-0012, 0019, 0021, 0023, 0024, 0031, 0033, 0034, 0035, 0036, 0041, 0044,
0050, 0052, 0054, 0055 numa copia NOVA da producao atual (entrega 41) e o fluxo VIZINHO de cada uma:
  - OC-0012: janela de 180 corridas (Glicose 8974 DIM1 "Tudo") + oraculo proprio da janela (180 RUNs mais recentes com
    resultado que participa) + vizinho: um mes so (sem corte, sem aviso, n do Painel = contagem do mes).
  - OC-0041: K3/K4 depois de digitar De/Ate + vizinho: De/Ate atravessando o ano (K3 vazio; Estatistica D4 sem "Ano 0").
  - OC-0019/0021/0023/0024: manuais digitados na aba Digitar Resultados (eventos ligados) + ATUALIZAR DADOS pelo nucleo
    do botao (mIntegracao.AtualizarDadosAutomatico, que le o DB_SEAC de producao so para leitura). Vizinhos: RESULTADO
    texto, IMPLAUSIVEL (trilha FORA_DA_ESTATISTICA, ADR-092), lote de urina + "Creatinina" (ADR-093), EQUIPAMENTO
    digitado (sem ambar).
  - OC-0031/0033/0034/0052/0054/0055: nucleo da janela do CEQ (mCEQForm.*) com login de ANALISTA de teste. Vizinhos:
    duplicidade (texto da mensagem sem ano repetido), Corrigir/Excluir sem sessao (ADR-092 OC-0057), edicao direta
    com selecao.
  - OC-0035: GUIA lido como usuario (secoes 3, 5, 6) e botoes citados procurados no arquivo.
  - OC-0036/0044/0050: formatos (Estatistica AT, alinhamento A:B, Principal A5).
Infraestrutura reaproveitada (sem copiar oraculos para os pontos novos): qa_final.QA (abre a copia e revela as abas,
imprime EXCEL_PID), repro_comum, e funcoes de leitura/escrita/vigia do valida_adr090_091 (Tester, reteste03).

Regras: so COPIA em %TEMP%/qcw/auditoria/2026-11-14_reteste04/; nada e salvo; fechar() encerra so o PID do QA;
nenhuma captura de tela (caixas lidas pelo texto das janelas do PID desta sessao). Dados de teste: manuais com MOTIVO
"Outro: TESTE-AUDITORIA", valores 123,456 / 111,111 / 12,345 / -1 / "abc"; CEQ rodada "TESTE-AUDITORIA C-A 2025".

Uso: python -u reteste04_usuario_f1.py   (saida JSON em %TEMP%/qcw/auditoria/2026-11-14_reteste04/Bioquimica_F1.json)
     python -u reteste04_usuario_f1.py --complemento   (2a copia nova: OC-0012 vizinho com um mes que TEM dados e H3
     vazio; validacao de dados do EQUIPAMENTO (OC-0019); OC-0052 com um par de teste por prefixo literal no mapa da
     EQA_Base W:Y da copia, apagado em seguida; Estatistica B5/D5 com o Painel em "Tudo")
     python -u reteste04_usuario_f1.py --validacao   (3a copia nova: o que a validacao de dados (lista) das colunas
     EQUIPAMENTO, MATRIZ e MOTIVO da Digitar Resultados aceita, e o manual gravado com o unico item que o menu oferece)
"""
import collections
import datetime as dt
import os
import re
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_argv = sys.argv[:]
sys.argv = [sys.argv[0], 'Bioquimica']
import valida_adr090_091 as V  # noqa: E402
sys.argv = _argv
import repro_comum as C  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = 'Bioquimica'
CICLO = '2026-11-14_reteste04'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
USR = 'TESTE_AUDITORIA_R4'
MOTIVO = 'Outro: TESTE-AUDITORIA'
t, up, num = V.t, V.up, V.num


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


# ------------------------------------------------------------------------------------------------ OC-0012
def janela_indep(fin, an, lote, eq, de, ate, nk=180):
    """Oraculo proprio: RUNs (corridas) com resultado que participa, ate o fim do periodo; as nk mais recentes."""
    por_run = collections.defaultdict(collections.Counter)
    total = collections.Counter()
    for r in fin:
        if up(r['ANALITO']) != up(an) or r['PARTICIPA_ESTATISTICA'] != 'SIM' or not num(r['RESULTADO']):
            continue
        if lote is not None and t(r['LOTE']) != t(lote):
            continue
        if eq and up(r['EQUIPAMENTO']) != up(eq):
            continue
        d = r['DATA_HORA']
        if not isinstance(d, dt.datetime):
            continue
        if isinstance(ate, dt.datetime) and d.date() > ate.date():
            continue
        dentro = not (isinstance(de, dt.datetime) and d.date() < de.date())
        if dentro:
            total[int(r['NIVEL'] or 0)] += 1
        por_run[(t(r['RUN']), dentro)][int(r['NIVEL'] or 0)] += 1
    runs = sorted({k[0] for k in por_run})
    jan = set(runs[-nk:])
    n_jan = collections.Counter()
    for (rn, dentro), c in por_run.items():
        if rn in jan:
            n_jan.update(c)
    cortadas_periodo = len({rn for (rn, dentro) in por_run if dentro and rn not in jan})
    return {'runs_total': len(runs), 'n_periodo': dict(total), 'n_janela': dict(n_jan), 'runs_cortadas_no_periodo':
            cortadas_periodo}


def painel_estado(qa, fin):
    p, eng = qa.ws('Painel'), qa.ws('Eng_Saida')
    de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
    de = C.limpo(de) if isinstance(de, dt.datetime) else None
    ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
    todos = eng.Range('AD1').Value is True
    an = t(qa.nome('selAnalito').Value)
    lote = None if todos else t(eng.Range('E1').Value).lstrip("'")
    eq = t(eng.Range('S1').Value)
    ps = qa.painel_stats()
    ind = janela_indep(fin, an, lote, eq, de, ate)
    return {'filtro': {'analito': an, 'lote': lote, 'equip': eq, 'de': str(de), 'ate': str(ate),
                       'K3': t(p.Range('K3').Value), 'K4': t(p.Range('K4').Value)},
            'M1_cortadas': int(eng.Range('M1').Value or 0), 'AF1': t(eng.Range('AF1').Value),
            'A39': t(p.Range('A39').Value), 'painel_n_media_dp_cv': ps,
            'J7_J8': [t(p.Range('J7').Value), t(p.Range('J8').Value)], 'indep': ind}


def s_O12(qa, fin):
    p = qa.ws('Painel')
    out = {}
    for c, v in (('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K4', 'Tudo (sem filtro)')):
        C.por(qa, p, c, v, local=False)
    time.sleep(1)
    o = painel_estado(qa, fin)
    ind = o['indep']
    af1_esp = ' · '.join(f'N{k} {ind["n_periodo"].get(k, 0)}' for k in (1, 2))
    o['AF1_esperado'] = af1_esp
    o['n_janela_painel'] = [o['painel_n_media_dp_cv'][k][1] for k in range(2)]
    o['ok'] = (o['M1_cortadas'] > 0 and 'indicadores do Painel' in o['A39'] and o['AF1'] == af1_esp
               and f'Período inteiro: {af1_esp}' in o['A39']
               and o['n_janela_painel'] == [ind['n_janela'].get(1, 0), ind['n_janela'].get(2, 0)])
    out['original_Glicose_8974_DIM1_Tudo'] = o
    # Estatistica: o n do periodo inteiro que o aviso manda olhar
    est = C.estat_tabela(qa)
    out['estatistica_Glicose'] = {k: [v['n'], v['media'], v['cv']] for k, v in est.items() if k.startswith('Glicose|')}
    out['estatistica_cabecalho'] = {a: t(qa.ws('Estatística').Range(a).Text)[:120] for a in
                                    ('A4', 'B4', 'C4', 'D4', 'E4', 'B5', 'D5', 'H5', 'J5')}
    # vizinho: um mes so (K3 2026, K4 Out) -> sem corte e sem aviso
    C.por(qa, p, 'K3', 2026, local=False)
    C.por(qa, p, 'K4', 'Out', local=False)
    time.sleep(1)
    v = painel_estado(qa, fin)
    v['n_janela_painel'] = [v['painel_n_media_dp_cv'][k][1] for k in range(2)]
    v['ok'] = (v['M1_cortadas'] == 0 and 'fora do gráfico' not in v['A39']
               and v['n_janela_painel'] == [v['indep']['n_periodo'].get(1, 0), v['indep']['n_periodo'].get(2, 0)])
    out['vizinho_mes_Out2026'] = v
    out['ok'] = o['ok'] and v['ok']
    return out


# ------------------------------------------------------------------------------------------------ OC-0041
def s_O41(qa, senha):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    C.por(qa, p, 'K3', 2026, local=False)
    C.por(qa, p, 'K4', 'S1', local=False)
    passos = []
    p.Protect(senha, True, True)
    try:
        def foto(rot):
            qa.ex.esperar()
            time.sleep(0.5)
            passos.append({'passo': rot, 'K3': t(p.Range('K3').Value), 'K4': t(p.Range('K4').Value),
                           'G3': str(p.Range('G3').Value)[:10], 'G4': str(p.Range('G4').Value)[:10],
                           'Estat_D4': t(e.Range('D4').Text), 'Estat_C4_E4': t(e.Range('C4').Text) + ' | '
                           + t(e.Range('E4').Text), 'tem_Ano_0': bool(re.search(r'\bAno 0\b', ' '.join(
                               t(e.Cells(r, c).Text) for r in range(2, 7) for c in range(1, 12))))})
        foto('K3 2026 + K4 S1')
        p.Range('G3').Value = dt.datetime(2026, 9, 1)
        p.Range('G4').Value = dt.datetime(2026, 10, 31)
        foto('De/Até 01/09–31/10/2026 digitado')
        p.Range('K4').Value = 'Out'
        foto('K4 Out')
        # vizinho: De/Ate atravessando o ano
        p.Range('G3').Value = dt.datetime(2025, 12, 1)
        p.Range('G4').Value = dt.datetime(2026, 1, 31)
        foto('De/Até 01/12/2025–31/01/2026 (vizinho)')
        p.Range('K3').Value = 2026
        p.Range('K4').Value = 'Ano inteiro'
        foto('K3 2026 + K4 Ano inteiro (vizinho)')
    finally:
        p.Unprotect(senha)
    ok = (passos[1]['K4'] == 'Personalizado' and passos[1]['K3'] == '2026'
          and passos[2]['G3'] == '2026-10-01' and passos[2]['G4'] == '2026-10-31' and passos[2]['K4'] == 'Out')
    ok_viz = (passos[3]['K4'] == 'Personalizado' and passos[3]['K3'] == '' and not passos[3]['tem_Ano_0']
              and passos[4]['G3'] == '2026-01-01' and passos[4]['G4'] == '2026-12-31')
    return {'passos': passos, 'ok': ok, 'ok_vizinho': ok_viz}


# ------------------------------------------------------------------------------------------------ manuais
def s_manuais(qa, fin, aud):
    out = {}
    p = qa.ws('Painel')
    lo, cab = V.manual_cols(qa)
    out['colunas'] = cab
    C.por(qa, p, 'L4', 'DIMENSION 2', local=False)
    C.por(qa, p, 'C3', 'Glicose', local=False)
    sel_eq = t(qa.nome('selEquipamento').Value)
    out['selEquipamento'] = sel_eq
    # o usuario chega pela navegacao do botao
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)
    dr = qa.ws('Digitar Resultados')
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)
    out['instrucao_A3'] = t(dr.Range('A3').Value)[:500]
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == 'Glicose' and int(r['NIVEL'] or 0) == 1
                    and isinstance(r['DATA_HORA'], dt.datetime) and t(r['EQUIPAMENTO']) == 'DIMENSION 1'],
                   key=lambda r: r['DATA_HORA'])
    R = reais[-30]
    dR = R['DATA_HORA']
    out['referencia_real'] = {k: t(R[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO',
                                                   'DATA_HORA', 'RUN')}
    dia = dR.strftime('%d/%m/%Y')
    lote = "'" + t(R['LOTE'])
    qa.ex.xl.EnableEvents = True
    linhas = {}

    def novo(rot, campos):
        _, r = C.linha_livre_manual(qa)
        for col, v in campos:
            V.escrever_celula(qa, r, col, v)
        linhas[rot] = r
        return r

    def eqcel(r):
        c = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r, 1)
        return {'valor': t(c.Value), 'fundo': c.Interior.Color, 'fonte': c.Font.Color,
                'nota': t(c.Comment.Text()) if c.Comment is not None else ''}

    # M1 (OC-0019 + OC-0023 original): sem EQUIPAMENTO e sem NIVEL
    r1 = novo('M1_sem_nivel_sem_equip', [('DATA', dia), ('HORA', '03:17'), ('LOTE', lote), ('ANALITO', 'Glicose'),
                                         ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
    out['O19_auto'] = eqcel(r1)
    # M1b (vizinho OC-0019): EQUIPAMENTO automatico, o usuario NAO corrige (fica DIMENSION 2)
    r1b = novo('M1b_equip_auto_mantido', [('DATA', dia), ('HORA', '03:19'), ('LOTE', lote), ('NIVEL', '2'),
                                          ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
    out['O19_auto_M1b'] = eqcel(r1b)
    # o usuario ve o ambar e digita o certo no M1
    V.escrever_celula(qa, r1, 'EQUIPAMENTO', 'DIMENSION 1')
    out['O19_digitado'] = eqcel(r1)
    try:
        out['O19_validacao_DIMENSION1'] = bool(lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r1, 1).Validation.Value)
    except Exception as ex:  # noqa: BLE001
        out['O19_validacao_DIMENSION1'] = repr(ex)[:80]
    # M2 conflito, M8 sem HORA, M7 implausivel, M9 resultado texto, M6 urina + Creatinina
    hm = (dR + dt.timedelta(minutes=5)).strftime('%H:%M')
    base = [('EQUIPAMENTO', 'DIMENSION 1'), ('MOTIVO', MOTIVO)]
    novo('M2_conflito', [('DATA', dia), ('HORA', hm), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                         ('RESULTADO', '111,111')] + base)
    novo('M8_sem_hora', [('DATA', dia), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                         ('RESULTADO', '123,456')] + base)
    novo('M7_implausivel', [('DATA', dia), ('HORA', '03:29'), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                            ('RESULTADO', '-1')] + base)
    novo('M9_resultado_texto', [('DATA', dia), ('HORA', '03:31'), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                                ('RESULTADO', 'abc')] + base)
    novo('M6_6354_Creatinina', [('DATA', dia), ('HORA', '03:27'), ('LOTE', "'6354"), ('NIVEL', '1'),
                                ('ANALITO', 'Creatinina'), ('RESULTADO', '12,345')] + base)
    aud.novos()
    t0 = time.time()
    r_at1 = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    out['atualizar1'] = {'ret': r_at1[:400], 's': round(time.time() - t0, 1)}
    trilha1 = aud.novos()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in leit.items()}
    out['ids'] = ids
    fin1 = qa.final_por_id()
    out['manuais_1'] = {k: {'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'DETALHE_txt': v['DETALHE'][1][:220],
                            'EQUIPAMENTO': v['EQUIPAMENTO'][0],
                            'final': {c: t(fin1.get(ids[k], {}).get(c)) for c in
                                      ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ANALITO', 'MATRIZ', 'EQUIPAMENTO',
                                       'RUN')}}
                        for k, v in leit.items()}
    t1 = collections.defaultdict(list)
    tr = []
    for x in trilha1:
        if not x['acao']:
            continue
        k = next((k for k, i in ids.items() if i and re.search(re.escape(i) + r'\b', x['txt'])), None)
        t1[k].append(x['acao'])
        tr.append({'acao': x['acao'], 'manual': k, 'txt': re.sub(r'\b[A-Z0-9_.-]*\\[^ |]*', '<usr>', x['txt'])[:400]})
    out['trilha_1'] = tr
    out['trilha_por_manual'] = dict(t1)
    rq = t(qa.run('mIntegracao.ResumoQA'))
    qai = qa.ler('tblQA_Integracao')
    err = collections.Counter(x['CODIGO'] for x in qai if x['SEVERIDADE'] == 'ERRO')
    ngov = sum(err.get(c, 0) for c in ('E01', 'E02', 'E03', 'E10', 'E11'))
    lin = rq.split('\r\n')
    out['O24'] = {'resumo': lin, 'erros_por_codigo': dict(err), 'gov': ngov, 'len': len(rq),
                  'E04_ids': [' | '.join(t(v) for v in x.values() if v)[:200] for x in qai if x['CODIGO'] == 'E04'][:6],
                  'gov_linhas': [' | '.join(t(v) for v in x.values() if v)[:200] for x in qai
                                 if x['SEVERIDADE'] == 'ERRO' and x['CODIGO'] in ('E01', 'E02', 'E03', 'E10', 'E11')][:6]}
    out['O24']['ok'] = (err.get('E04', 0) > 0
                        and any(x.startswith(f'• {err["E04"]} resultado(s) manual(is) incompleto(s)') for x in lin)
                        and 'Há ERRO DE GOVERNANÇA' not in rq
                        and (('governança' in rq) == (ngov > 0)))
    # completar o NIVEL do M1 e ATUALIZAR de novo
    V.escrever_celula(qa, linhas['M1_sem_nivel_sem_equip'], 'NIVEL', '1')
    t0 = time.time()
    r_at2 = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    out['atualizar2'] = {'ret': r_at2[:400], 's': round(time.time() - t0, 1)}
    trilha2 = aud.novos()
    m1 = V.ler_manual(qa, linhas['M1_sem_nivel_sem_equip'])
    fin2 = qa.final_por_id()
    out['M1_depois'] = {'STATUS': m1['STATUS'][0], 'RUN_txt': m1['RUN'][1], 'DETALHE_txt': m1['DETALHE'][1],
                        'final': {c: t(fin2.get(ids['M1_sem_nivel_sem_equip'], {}).get(c)) for c in
                                  ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'EQUIPAMENTO', 'RUN')}}
    out['trilha_2'] = [{'acao': x['acao'], 'txt': x['txt'][:300]} for x in trilha2
                       if x['acao'] and re.search(re.escape(ids['M1_sem_nivel_sem_equip']) + r'\b', x['txt'])]
    out['trilha_2_outros_manuais'] = [x['acao'] for x in trilha2 if x['acao'] and 'MANUAL' in x['acao']
                                      and not re.search(re.escape(ids['M1_sem_nivel_sem_equip']) + r'\b', x['txt'])]
    # varredura RUN/DETALHE "0" e sem acento na aba
    n = lo.ListRows.Count
    vr, vd = lo.ListColumns('RUN').DataBodyRange, lo.ListColumns('DETALHE').DataBodyRange
    vi = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    zeros, sem_acento = [], []
    for i in range(1, min(n, 400) + 1):
        if vi[i - 1][0] in (None, ''):
            continue
        a, b = t(vr.Cells(i, 1).Text), t(vd.Cells(i, 1).Text)
        if a == '0' or b == '0':
            zeros.append((t(vi[i - 1][0]), a, b))
        if re.search(r'obrigatorio|invalido|lancamento|ambiguo|numerico|lancado', b):
            sem_acento.append((t(vi[i - 1][0]), b[:120]))
    out['varredura'] = {'linhas': n, 'com_zero': zeros[:10], 'sem_acento': sem_acento[:10]}
    m = out['manuais_1']
    out['O19_ok'] = (out['O19_auto']['valor'] == sel_eq == 'DIMENSION 2' and out['O19_auto']['fundo'] == V.AMBAR
                     and 'Preenchido pelo Painel com DIMENSION 2' in out['O19_auto']['nota']
                     and out['O19_digitado']['valor'] == 'DIMENSION 1' and out['O19_digitado']['fundo'] != V.AMBAR
                     and not out['O19_digitado']['nota'] and 'EQUIPAMENTO' in out['instrucao_A3'].upper())
    out['O21_ok'] = (all(m[k]['RUN_txt'] == '' for k in ('M1_sem_nivel_sem_equip', 'M2_conflito', 'M8_sem_hora'))
                     and 'obrigatório' in m['M1_sem_nivel_sem_equip']['DETALHE_txt']
                     and out['M1_depois']['RUN_txt'] not in ('', '0') and out['M1_depois']['DETALHE_txt'] in ('',)
                     and not zeros and not sem_acento)
    out['O23_ok'] = (m['M1_sem_nivel_sem_equip']['STATUS'] == 'MANUAL_INCOMPLETO'
                     and t1.get('M1_sem_nivel_sem_equip') == ['RESULTADO_MANUAL_RECUSADO']
                     and m['M8_sem_hora']['STATUS'] == 'MANUAL_INCOMPLETO'
                     and t1.get('M8_sem_hora') == ['RESULTADO_MANUAL_RECUSADO']
                     and m['M2_conflito']['STATUS'] == 'CONFLITO_MANUAL'
                     and t1.get('M2_conflito') == ['RESULTADO_MANUAL_EM_CONFLITO']
                     and out['M1_depois']['STATUS'] == 'ATIVO'
                     and [x['acao'] for x in out['trilha_2']] == ['RESULTADO_MANUAL_INCLUIDO'])
    out['vizinho_ok'] = {
        'M7_implausivel_trilha_fora': t1.get('M7_implausivel') == ['RESULTADO_MANUAL_FORA_DA_ESTATISTICA'],
        'M9_texto_incompleto': m['M9_resultado_texto']['STATUS'] == 'MANUAL_INCOMPLETO'
        and t1.get('M9_resultado_texto') == ['RESULTADO_MANUAL_RECUSADO'],
        'M6_urina_CRE2': m['M6_6354_Creatinina']['final'].get('ANALITO') == 'CRE2',
        'M1b_equip_auto_DIM2': m['M1b_equip_auto_mantido']['final'].get('EQUIPAMENTO') == 'DIMENSION 2',
        'segundo_atualizar_sem_trilha_repetida': not out['trilha_2_outros_manuais']}
    return out


# ------------------------------------------------------------------------------------------------ CEQ extra
def s_ceq_vizinhos(qa, aud):
    """Duplicidade (OC-0055 vizinho) e edicao direta COM selecao (OC-0054 vizinho)."""
    out = {}
    c1 = C.campos_ceq('CAP', V.ROD, V.ANO, V.NOME_CEQ, 'TA-71', '104', unid=V.UN_CEQ)
    r1 = t(qa.run('mCEQForm.SalvarCEQ', c1))
    r2 = t(qa.run('mCEQForm.SalvarCEQ', c1))
    out['salvar'] = r1[:200]
    out['duplicado'] = r2[:400]
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    out['descricao'] = t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', lin)) if lin else ''
    out['dup_ok'] = r2.startswith('ERRO|5|') and '2025 2025' not in r2 and 'C-A 2025' in r2
    # edicao direta com selecao antes (caminho normal)
    cap = qa.ws('EQA.CAP_Dados')
    aud.novos()
    if lin:
        cap.Activate()
        cap.Range(f'F{lin}').Select()
        qa.ex.esperar()
        cap.Range(f'F{lin}').Value = 106
        qa.ex.esperar()
        tr = aud.novos()
        out['edicao_com_selecao'] = [x['txt'][:300] for x in tr]
        out['Z3_depois'] = t(cap.Range('Z3').Value)[:140]
        out['edicao_ok'] = (len(tr) == 1 and tr[0]['acao'] == 'CEQ_ALTERADO' and '104 → 106' in tr[0]['txt'])
        cap.Range('A1').Select()
        out['excluir'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:200]
        aud.novos()
    return out


def s_fim(qa, aud):
    """Logout + vizinho OC-0033 (Corrigir/Excluir sem sessao, ADR-092 OC-0057) + AbrirCEQ sem sessao."""
    out = {}
    c1 = C.campos_ceq('CAP', V.ROD, V.ANO, V.NOME_CEQ, 'TA-81', '108', unid=V.UN_CEQ)
    r1 = t(qa.run('mCEQForm.SalvarCEQ', c1))
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    out['registro_com_sessao'] = r1[:120]
    qa.run('mSeguranca.Logout')
    qa.ex.esperar()
    out['SessaoAtiva'] = qa.run('mSeguranca.SessaoAtiva')
    aud.novos()
    out['salvar'] = t(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', V.ROD, V.ANO, V.NOME_CEQ, 'TA-82', '110',
                                                                  unid=V.UN_CEQ)))[:160]
    if lin:
        cap = qa.ws('EQA.CAP_Dados')
        f0 = t(cap.Range(f'F{lin}').Value)
        out['corrigir'] = t(qa.run('mCEQForm.CorrigirCEQ', lin, C.campos_ceq('CAP', V.ROD, V.ANO, V.NOME_CEQ, 'TA-81',
                                                                             '115', unid=V.UN_CEQ)))[:160]
        out['F_antes_depois'] = [f0, t(cap.Range(f'F{lin}').Value)]
        out['excluir'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:160]
        out['A_depois'] = t(cap.Range(f'A{lin}').Value)
    out['audit'] = [x['acao'] for x in aud.novos()]
    out['abrir'] = V.abrir_ceq_sem_sessao(qa)
    out['ok'] = (out['salvar'].startswith('ERRO|0|') and out.get('corrigir', '').startswith('ERRO|0|')
                 and out.get('excluir', '').startswith('ERRO|0|') and len(set(out.get('F_antes_depois', [1, 2]))) == 1
                 and not [a for a in out['audit'] if a and a.startswith('CEQ_')] and out['abrir']['ok'])
    return out


def main():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F1.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F1.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p, e = qa.ws('Painel'), qa.ws('Estatística')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        # sem sessao (arquivo recem-aberto): OC-0033 original
        secao(L, 'O33_sem_sessao_inicio', V.s_sem_sessao, qa)
        secao(L, 'O33_abrir_sem_sessao', V.abrir_ceq_sem_sessao, qa)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        secao(L, 'estaticas', V.s_estaticas, qa, senha)
        secao(L, 'O35_guia', V.s_guia, qa)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        secao(L, 'O12', s_O12, qa, fin)
        secao(L, 'O41', s_O41, qa, senha)
        aud = V.Audit(qa)
        secao(L, 'manuais', s_manuais, qa, fin, aud)
        secao(L, 'ceq', V.s_ceq, qa, senha, aud)
        secao(L, 'ceq_vizinhos', s_ceq_vizinhos, qa, aud)
        secao(L, 'fim_sem_sessao', s_fim, qa, aud)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM', flush=True)


MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']


def c_O12(qa, fin):
    p = qa.ws('Painel')
    out = {}
    cont = collections.Counter(r['DATA_HORA'].strftime('%Y-%m') for r in fin
                               if t(r['ANALITO']) == 'Glicose' and t(r['LOTE']) == '8974'
                               and t(r['EQUIPAMENTO']) == 'DIMENSION 1' and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                               and isinstance(r['DATA_HORA'], dt.datetime))
    ym, n = cont.most_common(1)[0]
    out['meses_8974_DIM1'] = dict(sorted(cont.items()))
    for c, v in (('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', "'8974"), ('K3', int(ym[:4])),
                 ('K4', MESES[int(ym[5:]) - 1])):
        C.por(qa, p, c, v, local=False)
    time.sleep(1)
    v = painel_estado(qa, fin)
    v['n_janela_painel'] = [v['painel_n_media_dp_cv'][k][1] for k in range(2)]
    v['ok'] = (v['M1_cortadas'] == 0 and 'fora do gráfico' not in v['A39']
               and v['n_janela_painel'] == [v['indep']['n_periodo'].get(1, 0), v['indep']['n_periodo'].get(2, 0)])
    out['vizinho_mes_com_dados'] = v
    # H3 vazio (TODOS os lotes) + Tudo
    C.por(qa, p, 'K4', 'Tudo (sem filtro)', local=False)
    p.Range('H3').MergeArea.ClearContents()
    qa.ex.esperar()
    time.sleep(1.5)
    w = painel_estado(qa, fin)
    w['n_janela_painel'] = [w['painel_n_media_dp_cv'][k][1] for k in range(2)]
    ind = w['indep']
    af1 = ' · '.join(f'N{k} {ind["n_periodo"].get(k, 0)}' for k in (1, 2))
    w['AF1_esperado'] = af1
    w['ok'] = (w['AF1'] == af1 and f'Período inteiro: {af1}' in w['A39'] and 'indicadores do Painel' in w['A39']
               and w['n_janela_painel'] == [ind['n_janela'].get(1, 0), ind['n_janela'].get(2, 0)])
    out['vizinho_H3_vazio_Tudo'] = w
    e = qa.ws('Estatística')
    out['estat_B5_D5_tudo'] = {a: [t(e.Range(a).Formula), t(e.Range(a).Text), t(e.Range(a).NumberFormat)]
                               for a in ('B5', 'D5', 'H5', 'J5')}
    out['painel_G3_G4'] = [t(p.Range('G3').Formula), t(p.Range('G4').Formula)]
    return out


def c_O19(qa):
    lo, _ = V.manual_cols(qa)
    cel = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(1, 1)
    out = {}
    try:
        vv = cel.Validation
        out['tipo'] = vv.Type
        out['formula1'] = t(vv.Formula1)
        out['showerror'] = bool(vv.ShowError)
        out['alertstyle'] = vv.AlertStyle
        out['errormessage'] = t(vv.ErrorMessage)[:200]
        out['incelldropdown'] = bool(vv.InCellDropdown)
    except Exception as ex:  # noqa: BLE001
        out['erro'] = repr(ex)[:120]
    _, r = C.linha_livre_manual(qa)
    c2 = lo.ListColumns('EQUIPAMENTO').DataBodyRange.Cells(r, 1)
    qa.ex.xl.EnableEvents = False
    res = {}
    try:
        for txt in ('DIMENSION 1', 'DIMENSION 2', 'DIMENSION 1,DIMENSION 2', 'DIMENSION 9'):
            c2.Value = txt
            try:
                res[txt] = bool(c2.Validation.Value)
            except Exception as ex:  # noqa: BLE001
                res[txt] = repr(ex)[:60]
        c2.ClearContents()
        for col in ('LOTE', 'ANALITO', 'NIVEL'):
            cc = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
            try:
                out[f'{col}_formula1'] = t(cc.Validation.Formula1)[:80]
            except Exception as ex:  # noqa: BLE001
                out[f'{col}_formula1'] = repr(ex)[:60]
    finally:
        qa.ex.xl.EnableEvents = True
    out['validation_value'] = res
    try:
        nomes = [t(n.Name) for n in qa.wb.Names if 'equip' in t(n.Name).lower()]
        out['nomes_equip'] = {n: t(qa.wb.Names(n).RefersTo)[:120] for n in nomes[:8]}
    except Exception as ex:  # noqa: BLE001
        out['nomes_equip'] = repr(ex)[:80]
    return out


def c_O52(qa, senha):
    b = qa.ws('EQA_Base')
    try:
        b.Unprotect(senha)
    except Exception:
        pass
    v = b.Range('W2:Y600').Value
    livre = next(i for i, l in enumerate(v, 2) if not any(l))
    ev = qa.ex.xl.EnableEvents
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre + 1}').Value = (('CAP', 'TESTE-AUDITORIA MC', 'Glicose'),
                                                   ('CAP', 'TESTE-AUDITORIA MCX', 'Ureia'))
    finally:
        qa.ex.xl.EnableEvents = ev
    out = {'linhas_teste': [livre, livre + 1],
           'aviso_literal': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'TESTE-AUDITORIA MC')),
           'aviso_maior': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'TESTE-AUDITORIA MCX')),
           'aviso_palavra_real': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Lactate')),
           'aviso_sem_par_real': t(qa.run('mCEQForm.AvisoAnalitoCEQ', 'CAP', 'Glucose, serum'))}
    qa.ex.xl.EnableEvents = False
    try:
        b.Range(f'W{livre}:Y{livre + 1}').ClearContents()
    finally:
        qa.ex.xl.EnableEvents = ev
    out['ok'] = ('“TESTE-AUDITORIA MCX” (Ureia)' in out['aviso_literal'] and 'existe também' not in out['aviso_maior']
                 and 'Lactate Dehydrogenase' in out['aviso_palavra_real'])
    return out


def complemento():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F1_comp.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F1_comp.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p, e = qa.ws('Painel'), qa.ws('Estatística')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        L['linhas_tblCQ_Final'] = len(fin)
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        secao(L, 'O12_vizinhos', c_O12, qa, fin)
        secao(L, 'O19_validacao', c_O19, qa)
        secao(L, 'O52_par_literal', c_O52, qa, senha)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM', flush=True)

def v_listas(qa):
    """Validacao de dados (lista) das colunas da Digitar Resultados: o que o Excel aceita (Range.Validation.Value)."""
    lo, _ = V.manual_cols(qa)
    _, r = C.linha_livre_manual(qa)
    casos = {'EQUIPAMENTO': ['DIMENSION 1', 'DIMENSION 2', 'DIMENSION 1,DIMENSION 2'],
             'MATRIZ': ['SORO', 'URINA'],
             'MOTIVO': ['Falha do interfaceamento', 'Outro', 'Resultado não transmitido']}
    out = {}
    qa.ex.xl.EnableEvents = False
    try:
        for col, vals in casos.items():
            c = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
            vv = c.Validation
            d = {'formula1': t(vv.Formula1)[:160], 'alertstyle': vv.AlertStyle, 'showerror': bool(vv.ShowError),
                 'msg': t(vv.ErrorMessage)[:80]}
            for x in vals:
                c.Value = x
                d[x] = bool(c.Validation.Value)
            d[t(vv.Formula1)] = None
            c.Value = t(vv.Formula1)
            d['valor = formula1 inteira'] = bool(c.Validation.Value)
            c.ClearContents()
            out[col] = d
    finally:
        qa.ex.xl.EnableEvents = True
    return out


def v_consequencia(qa, fin):
    """O que acontece se o usuario escolhe o unico item que o menu oferece ("DIMENSION 1,DIMENSION 2")."""
    p = qa.ws('Painel')
    C.por(qa, p, 'L4', 'DIMENSION 1', local=False)
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and t(r['ANALITO']) == 'Glicose'
                    and int(r['NIVEL'] or 0) == 1 and isinstance(r['DATA_HORA'], dt.datetime)
                    and t(r['EQUIPAMENTO']) == 'DIMENSION 1'], key=lambda r: r['DATA_HORA'])
    R = reais[-30]
    dia = R['DATA_HORA'].strftime('%d/%m/%Y')
    _, r = C.linha_livre_manual(qa)
    for col, v in [('DATA', dia), ('HORA', '03:41'), ('EQUIPAMENTO', 'DIMENSION 1,DIMENSION 2'),
                   ('LOTE', "'" + t(R['LOTE'])), ('NIVEL', '1'), ('ANALITO', 'Glicose'), ('RESULTADO', '123,456'),
                   ('MOTIVO', 'Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro')]:
        V.escrever_celula(qa, r, col, v)
    ret = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    m = V.ler_manual(qa, r)
    idm = m['ID_REGISTRO'][0]
    f = qa.final_por_id().get(idm, {})
    qai = [' | '.join(t(v) for v in x.values() if v)[:220] for x in qa.ler('tblQA_Integracao')
           if any(t(v) == idm for v in x.values())]
    return {'ret': ret[:300], 'manual': {k: m[k][0] for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'STATUS', 'DETALHE', 'RUN')},
            'final': {c: t(f.get(c)) for c in ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'EQUIPAMENTO', 'RUN')},
            'qa': qai}


def validacao_listas():
    src, copia, h0 = C.copiar(PROD)
    destino = os.path.join(C.PASTA, f'{PROD}_F1_val.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    os.replace(copia, destino)
    L = C.Log(os.path.join(C.PASTA, f'{PROD}_F1_val.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod_12': h0, 'copia': os.path.basename(destino)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, destino)
    L['excel_pid'] = qa.ex.pid
    try:
        p = qa.ws('Painel')
        p.Activate()
        qa.ex.esperar()
        fin = qa.ler('tblCQ_Final')
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        try:
            p.Unprotect(senha)
        except Exception:
            pass
        secao(L, 'listas', v_listas, qa)
        secao(L, 'consequencia', v_consequencia, qa, fin)
    except Exception as ex:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    h1 = C.sha(src)
    L['producao_intacta'] = {'sha_antes': h0, 'sha_depois': h1, 'igual': h0 == h1}
    print('FIM', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    if '--complemento' in sys.argv:
        complemento()
    elif '--validacao' in sys.argv:
        validacao_listas()
    else:
        main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
