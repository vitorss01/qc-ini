# -*- coding: utf-8 -*-
"""reteste04_usuario_f1_hema.py -- Usuario Simulado, ciclo 2026-11-14_reteste04, frente F1, modo RETESTE (Hematologia).

Irmao do reteste04_usuario_f1.py (Bioquimica). Refaz o fluxo ORIGINAL das ocorrencias OC-0012, 0021, 0023, 0024,
0031, 0033, 0034, 0035, 0036, 0041, 0050, 0052, 0054, 0055 numa copia NOVA da producao atual da Hematologia e o fluxo
VIZINHO de cada uma:
  - OC-0012: WBC, H3 vazio (todos os lotes), K3 2026 + "Ano inteiro" (cenario do parecer_cq G11 / reteste03) com
    oraculo proprio da janela (as 180 RUNs mais recentes com resultado que participa) + vizinhos: o lote de WBC com mais
    dados em "Tudo (sem filtro)"; um mes so (sem corte, sem aviso); cabecalho B5/D5 da Estatistica com "Tudo" (R4-02).
  - OC-0041: K3/K4 depois de digitar De/Ate + vizinho: De/Ate atravessando o ano.
  - OC-0021/0023/0024: manuais digitados na aba Digitar Resultados (navegacao pelo botao, eventos ligados) + ATUALIZAR
    DADOS pelo nucleo do botao (mIntegracao.AtualizarDadosAutomatico, que le o DB_SEAC de producao so para leitura).
    Vizinhos: RESULTADO texto, IMPLAUSIVEL (trilha FORA_DA_ESTATISTICA, ADR-092), EQUIPAMENTO automatico (Hema: sem
    ambar), segundo ATUALIZAR sem trilha repetida, validacao de dados (lista) das colunas EQUIPAMENTO/MATRIZ/MOTIVO
    (R4-01 da Bio, que la ficou "nao testada na Hema").
  - OC-0031/0033/0034/0052/0054/0055: nucleo da janela do CEQ (mCEQForm.*) com login de ANALISTA de teste (funcoes
    do valida_adr090_091 com produto Hematologia). Vizinhos: duplicidade, edicao direta com selecao,
    Corrigir/Excluir sem sessao.
  - OC-0035: GUIA lido (secoes 3, 5, 6) e botoes citados procurados no arquivo.
  - OC-0036/0050: formatos (Estatistica AT, Principal A5) + OC-0049 (linhas vazias) de carona.

Regras: so COPIA em %TEMP%/qcw/auditoria/2026-11-14_reteste04/; nada e salvo; fechar() encerra so o PID do QA;
nenhuma captura de tela (caixas lidas pelo texto das janelas do PID desta sessao). Dados de teste: manuais com MOTIVO
"Outro: TESTE-AUDITORIA", valores 123,456 / 111,111 / -1 / "abc"; CEQ rodada "TESTE-AUDITORIA C-A 2025", amostras TA-*.

Uso: python -u reteste04_usuario_f1_hema.py   (saida %TEMP%/qcw/auditoria/2026-11-14_reteste04/Hematologia_F1.json)
     python -u reteste04_usuario_f1_hema.py --complemento   (2a copia nova, so leitura de tela: o que o Painel mostra
     nas linhas 7-9 no modo TODOS com corte, o que a Estatistica mostra quando o aviso manda olha-la, e onde estao os
     botoes "Consolidar" (EQA_Base x abas de dados), para conferir o texto do Z3 e do GUIA)

Oraculo da janela (OC-0012), conferido depois da 1a rodada: no modo TODOS a "corrida" do motor e o par (RUN, LOTE) com
resultado que participa (ADR-080, corrida do lote); por RUN so, o oraculo da 278 -> 272 cortadas e a janela +1 por nivel.
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
sys.argv = [sys.argv[0], 'Hematologia']
import valida_adr090_091 as V  # noqa: E402
sys.argv = _argv
import repro_comum as C  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login  # noqa: E402

PROD = 'Hematologia'
NLV = 3
CICLO = '2026-11-14_reteste04'
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
USR = 'TESTE_AUDITORIA_R4'
MOTIVO = 'Outro: TESTE-AUDITORIA'
t, up, num = V.t, V.up, V.num
assert not V.BIO


def secao(L, nome, fn, *a):
    print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
    return V.secao(L, nome, fn, *a)


# ------------------------------------------------------------------------------------------------ OC-0012
def janela_indep(fin, an, lote, de, ate, nk=180):
    """Oraculo proprio: RUNs com resultado que participa, ate o fim do periodo; as nk mais recentes."""
    por_run = collections.defaultdict(collections.Counter)
    total = collections.Counter()
    for r in fin:
        if up(r['ANALITO']) != up(an) or r['PARTICIPA_ESTATISTICA'] != 'SIM' or not num(r['RESULTADO']):
            continue
        if lote is not None and t(r['LOTE']) != t(lote):
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
    cortadas = len({rn for (rn, dentro) in por_run if dentro and rn not in jan})
    return {'runs_total': len(runs), 'n_periodo': dict(total), 'n_janela': dict(n_jan), 'runs_cortadas_no_periodo':
            cortadas}


def painel_estado(qa, fin):
    p, eng = qa.ws('Painel'), qa.ws('Eng_Saida')
    de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
    de = C.limpo(de) if isinstance(de, dt.datetime) else None
    ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
    todos = eng.Range('AD1').Value is True
    an = t(qa.nome('selAnalito').Value)
    lote = None if todos else t(eng.Range('E1').Value).lstrip("'")
    ps = qa.painel_stats()
    ind = janela_indep(fin, an, lote, de, ate)
    af1_esp = ' · '.join(f'N{k} {ind["n_periodo"].get(k, 0)}' for k in range(1, NLV + 1))
    return {'filtro': {'analito': an, 'lote': lote, 'todos': todos, 'de': str(de), 'ate': str(ate),
                       'H3': t(p.Range('H3').Value), 'K3': t(p.Range('K3').Value), 'K4': t(p.Range('K4').Value)},
            'M1_cortadas': int(eng.Range('M1').Value or 0), 'AF1': t(eng.Range('AF1').Value), 'AF1_esperado': af1_esp,
            'A39': t(p.Range('A39').Value), 'painel_n_media_dp_cv': ps,
            'n_janela_painel': [ps[k][1] for k in range(NLV)],
            'J7_J9': [t(p.Range(f'J{7 + k}').Value) for k in range(NLV)], 'indep': ind}


def ok_corte(o):
    ind = o['indep']
    return (o['M1_cortadas'] > 0 and 'indicadores do Painel' in o['A39'] and o['AF1'] == o['AF1_esperado']
            and f'Período inteiro: {o["AF1_esperado"]}' in o['A39']
            and o['n_janela_painel'] == [ind['n_janela'].get(k, 0) for k in range(1, NLV + 1)])


def ok_sem_corte(o):
    ind = o['indep']
    return (o['M1_cortadas'] == 0 and 'fora do gráfico' not in o['A39']
            and o['n_janela_painel'] == [ind['n_periodo'].get(k, 0) for k in range(1, NLV + 1)])


def s_O12(qa, fin):
    p = qa.ws('Painel')
    e = qa.ws('Estatística')
    out = {}
    # original (parecer_cq G11 / reteste03 Hema): WBC, TODOS os lotes, Ano inteiro 2026
    for c, v in (('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro')):
        C.por(qa, p, c, v, local=False)
    p.Range('H3').MergeArea.ClearContents()
    qa.ex.esperar()
    time.sleep(1.5)
    o = painel_estado(qa, fin)
    o['ok'] = ok_corte(o)
    out['original_WBC_TODOS_2026'] = o
    out['estatistica_WBC'] = {k: [v['n'], v['media'], v['cv']] for k, v in C.estat_tabela(qa).items()
                              if k.startswith('WBC|')}
    # vizinho 1: o lote de WBC com mais dados, "Tudo (sem filtro)"
    cont = collections.Counter(t(r['LOTE']) for r in fin if t(r['ANALITO']) == 'WBC'
                               and r['PARTICIPA_ESTATISTICA'] == 'SIM')
    lote, n = cont.most_common(1)[0]
    out['lotes_WBC'] = dict(cont.most_common(8))
    for c, v in (('C3', 'WBC'), ('H3', "'" + lote), ('K4', 'Tudo (sem filtro)')):
        C.por(qa, p, c, v, local=False)
    time.sleep(1.5)
    w = painel_estado(qa, fin)
    w['ok'] = ok_corte(w) if w['indep']['runs_cortadas_no_periodo'] > 0 else ok_sem_corte(w)
    w['estatistica_WBC'] = {k: [v['n'], v['media'], v['cv']] for k, v in C.estat_tabela(qa).items()
                            if k.startswith('WBC|')}
    w['estat_cabecalho'] = {a: [t(e.Range(a).Formula)[:80], t(e.Range(a).Text)[:80]]
                            for a in ('B4', 'D4', 'B5', 'D5', 'H5', 'J5')}
    out['vizinho_lote_Tudo'] = w
    # vizinho 2: um mes com dados desse lote
    meses = collections.Counter(r['DATA_HORA'].strftime('%Y-%m') for r in fin
                                if t(r['ANALITO']) == 'WBC' and t(r['LOTE']) == lote
                                and r['PARTICIPA_ESTATISTICA'] == 'SIM' and isinstance(r['DATA_HORA'], dt.datetime))
    ym = meses.most_common(1)[0][0]
    out['meses_lote'] = dict(sorted(meses.items()))
    C.por(qa, p, 'K3', int(ym[:4]), local=False)
    C.por(qa, p, 'K4', MESES[int(ym[5:]) - 1], local=False)
    time.sleep(1.5)
    m = painel_estado(qa, fin)
    m['ok'] = ok_sem_corte(m)
    out['vizinho_mes'] = m
    out['ok'] = o['ok'] and w['ok'] and m['ok']
    return out


MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']


# ------------------------------------------------------------------------------------------------ OC-0041
def s_O41(qa, senha):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    C.por(qa, p, 'C3', 'WBC', local=False)
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
                           'Estat_D4': t(e.Range('D4').Text), 'Estat_B4': t(e.Range('B4').Text),
                           'tem_Ano_0': bool(re.search(r'\bAno 0\b', ' '.join(
                               t(e.Cells(r, c).Text) for r in range(2, 7) for c in range(1, 12))))})
        foto('K3 2026 + K4 S1')
        p.Range('G3').Value = dt.datetime(2026, 9, 1)
        p.Range('G4').Value = dt.datetime(2026, 10, 31)
        foto('De/Até 01/09–31/10/2026 digitado')
        p.Range('K4').Value = 'Out'
        foto('K4 Out')
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


# ------------------------------------------------------------------------------------------------ listas
def s_listas(qa):
    """Validacao de dados (lista) das colunas EQUIPAMENTO/MATRIZ/MOTIVO: o que o Excel aceita (Validation.Value)."""
    lo, cab = V.manual_cols(qa)
    _, r = C.linha_livre_manual(qa)
    out = {}
    qa.ex.xl.EnableEvents = False
    try:
        for col in ('EQUIPAMENTO', 'MATRIZ', 'MOTIVO', 'LOTE', 'ANALITO', 'NIVEL'):
            if col not in cab:
                out[col] = 'coluna ausente'
                continue
            c = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
            try:
                vv = c.Validation
                f1 = t(vv.Formula1)
                d = {'tipo': vv.Type, 'formula1': f1[:200], 'alertstyle': vv.AlertStyle, 'showerror': bool(vv.ShowError),
                     'msg': t(vv.ErrorMessage)[:100], 'aceita': {}}
            except Exception as ex:  # noqa: BLE001
                out[col] = 'sem validacao ' + repr(ex)[:60]
                continue
            if not f1.startswith('='):
                itens = [x.strip() for x in f1.split(',')] if ',' in f1 else [f1]
                for x in itens + ([f1] if ',' in f1 else []):
                    c.Value = x
                    try:
                        d['aceita'][x] = bool(c.Validation.Value)
                    except Exception as ex:  # noqa: BLE001
                        d['aceita'][x] = repr(ex)[:60]
                c.ClearContents()
            out[col] = d
    finally:
        qa.ex.xl.EnableEvents = True
    return out


# ------------------------------------------------------------------------------------------------ manuais
def s_manuais(qa, fin, aud):
    out = {}
    p = qa.ws('Painel')
    lo, cab = V.manual_cols(qa)
    out['colunas'] = cab
    C.por(qa, p, 'C3', 'WBC', local=False)
    sel_eq = t(qa.nome('selEquipamento').Value)
    out['selEquipamento'] = sel_eq
    try:
        qa.run('mIntegracao.IrDigitarResultados', teto=120)
    except Exception as ex:  # noqa: BLE001
        out['ir_erro'] = repr(ex)
    dr = qa.ws('Digitar Resultados')
    out['aba_ativa'] = t(qa.wb.ActiveSheet.Name)
    out['celula_ativa'] = t(qa.ex.xl.ActiveCell.Address)
    out['instrucao_A3'] = t(dr.Range('A3').Value)[:600]
    reais = sorted([r for r in fin if r['STATUS_ANALITICO'] == 'ATIVO' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                    and t(r['ANALITO']) == 'WBC' and int(r['NIVEL'] or 0) == 1
                    and isinstance(r['DATA_HORA'], dt.datetime)], key=lambda r: r['DATA_HORA'])
    R = reais[-30]
    dR = R['DATA_HORA']
    out['referencia_real'] = {k: t(R[k]) for k in ('ID_REGISTRO', 'EQUIPAMENTO', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO',
                                                   'DATA_HORA', 'RUN')}
    dia = dR.strftime('%d/%m/%Y')
    lote = "'" + t(R['LOTE'])
    eqR = t(R['EQUIPAMENTO'])
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

    # M1 (OC-0023 original): sem NIVEL e sem EQUIPAMENTO (carimbo do Painel)
    r1 = novo('M1_sem_nivel', [('DATA', dia), ('HORA', '03:17'), ('LOTE', lote), ('ANALITO', 'WBC'),
                               ('RESULTADO', '123,456'), ('MOTIVO', MOTIVO)])
    out['equip_auto_M1'] = eqcel(r1)
    hm = (dR + dt.timedelta(minutes=5)).strftime('%H:%M')
    base = [('EQUIPAMENTO', eqR), ('MOTIVO', MOTIVO)]
    novo('M2_conflito', [('DATA', dia), ('HORA', hm), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'WBC'),
                         ('RESULTADO', '111,111')] + base)
    novo('M8_sem_hora', [('DATA', dia), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'WBC'),
                         ('RESULTADO', '123,456')] + base)
    novo('M7_implausivel', [('DATA', dia), ('HORA', '03:29'), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'WBC'),
                            ('RESULTADO', '-1')] + base)
    novo('M9_resultado_texto', [('DATA', dia), ('HORA', '03:31'), ('LOTE', lote), ('NIVEL', '1'), ('ANALITO', 'WBC'),
                                ('RESULTADO', 'abc')] + base)
    out['linhas'] = linhas
    out['antes_atualizar'] = {k: {c: V.ler_manual(qa, r)[c][1] for c in ('ID_REGISTRO', 'STATUS', 'RUN', 'DETALHE',
                                                                          'EQUIPAMENTO')} for k, r in linhas.items()}
    aud.novos()
    t0 = time.time()
    r_at1 = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    out['atualizar1'] = {'ret': r_at1[:600], 's': round(time.time() - t0, 1)}
    trilha1 = aud.novos()
    leit = {k: V.ler_manual(qa, r) for k, r in linhas.items()}
    ids = {k: v['ID_REGISTRO'][0] for k, v in leit.items()}
    out['ids'] = ids
    fin1 = qa.final_por_id()
    out['manuais_1'] = {k: {'STATUS': v['STATUS'][0], 'RUN_txt': v['RUN'][1], 'DETALHE_txt': v['DETALHE'][1][:240],
                            'EQUIPAMENTO': v['EQUIPAMENTO'][0],
                            'final': {c: t(fin1.get(ids[k], {}).get(c)) for c in
                                      ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'ANALITO', 'EQUIPAMENTO', 'RUN',
                                       'TIPO_PLOTAGEM_LJ')}}
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
                  'E04': [' | '.join(t(v) for v in x.values() if v)[:200] for x in qai if x['CODIGO'] == 'E04'][:6],
                  'gov_linhas': [' | '.join(t(v) for v in x.values() if v)[:200] for x in qai
                                 if x['SEVERIDADE'] == 'ERRO' and x['CODIGO'] in ('E01', 'E02', 'E03', 'E10', 'E11')][:6],
                  'msg_atualizar_tem_ERRO_GOV': 'ERRO DE GOVERNANÇA' in r_at1.upper()}
    out['O24']['ok'] = (err.get('E04', 0) > 0
                        and any(x.startswith(f'• {err["E04"]} resultado(s) manual(is) incompleto(s)') for x in lin)
                        and 'Há ERRO DE GOVERNANÇA' not in rq
                        and (('governança' in rq) == (ngov > 0)))
    # completar o NIVEL do M1 e ATUALIZAR de novo
    V.escrever_celula(qa, linhas['M1_sem_nivel'], 'NIVEL', '1')
    t0 = time.time()
    r_at2 = t(qa.run('mIntegracao.AtualizarDadosAutomatico', teto=1500))
    out['atualizar2'] = {'ret': r_at2[:400], 's': round(time.time() - t0, 1)}
    trilha2 = aud.novos()
    m1 = V.ler_manual(qa, linhas['M1_sem_nivel'])
    fin2 = qa.final_por_id()
    i1 = ids['M1_sem_nivel']
    out['M1_depois'] = {'STATUS': m1['STATUS'][0], 'RUN_txt': m1['RUN'][1], 'DETALHE_txt': m1['DETALHE'][1],
                        'final': {c: t(fin2.get(i1, {}).get(c)) for c in
                                  ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'EQUIPAMENTO', 'RUN')}}
    out['trilha_2'] = [{'acao': x['acao'], 'txt': re.sub(r'\b[A-Z0-9_.-]*\\[^ |]*', '<usr>', x['txt'])[:300]}
                       for x in trilha2 if x['acao'] and re.search(re.escape(i1) + r'\b', x['txt'])]
    out['trilha_2_outros_manuais'] = [x['acao'] for x in trilha2 if x['acao'] and 'MANUAL' in x['acao']
                                      and not re.search(re.escape(i1) + r'\b', x['txt'])]
    rq2 = t(qa.run('mIntegracao.ResumoQA'))
    out['O24_depois'] = rq2.split('\r\n')
    # varredura RUN/DETALHE "0" e sem acento na aba
    n = lo.ListRows.Count
    vr, vd = lo.ListColumns('RUN').DataBodyRange, lo.ListColumns('DETALHE').DataBodyRange
    vi = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    zeros, sem_acento, detalhes = [], [], collections.Counter()
    for i in range(1, min(n, 400) + 1):
        if vi[i - 1][0] in (None, ''):
            continue
        a, b = t(vr.Cells(i, 1).Text), t(vd.Cells(i, 1).Text)
        detalhes[b[:70]] += 1
        if a == '0' or b == '0':
            zeros.append((t(vi[i - 1][0]), a, b))
        if re.search(r'obrigatorio|invalido|lancamento|ambiguo|numerico|lancado|implausivel|grafico|estatistica', b):
            sem_acento.append((t(vi[i - 1][0]), b[:140]))
    out['varredura'] = {'linhas': n, 'com_zero': zeros[:10], 'sem_acento': sem_acento[:10],
                        'detalhes_distintos': dict(detalhes.most_common(12))}
    m = out['manuais_1']
    out['O21_ok'] = (all(m[k]['RUN_txt'] == '' for k in ('M1_sem_nivel', 'M2_conflito', 'M8_sem_hora'))
                     and 'obrigatório' in m['M1_sem_nivel']['DETALHE_txt']
                     and out['M1_depois']['RUN_txt'] not in ('', '0') and out['M1_depois']['DETALHE_txt'] == ''
                     and not zeros and not sem_acento)
    out['O23_ok'] = (m['M1_sem_nivel']['STATUS'] == 'MANUAL_INCOMPLETO'
                     and t1.get('M1_sem_nivel') == ['RESULTADO_MANUAL_RECUSADO']
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
        'M1_equip_auto_sem_ambar': out['equip_auto_M1']['valor'] == sel_eq and out['equip_auto_M1']['fundo'] != V.AMBAR,
        'M1_ativo_no_final_participa': out['M1_depois']['final'].get('PARTICIPA_ESTATISTICA') == 'SIM',
        'segundo_atualizar_sem_trilha_repetida': not out['trilha_2_outros_manuais']}
    return out


# ------------------------------------------------------------------------------------------------ CEQ extra
def s_ceq_vizinhos(qa, aud):
    out = {}
    c1 = C.campos_ceq('CAP', V.ROD, V.ANO, V.NOME_CEQ, 'TA-71', '104', unid=V.UN_CEQ)
    r1 = t(qa.run('mCEQForm.SalvarCEQ', c1))
    r2 = t(qa.run('mCEQForm.SalvarCEQ', c1))
    out['salvar'] = r1[:200]
    out['duplicado'] = r2[:400]
    lin = int(r1.split('|')[1]) if r1.startswith('OK|') else 0
    out['descricao'] = t(qa.run('mCEQForm.DescricaoCEQ', 'CAP', lin)) if lin else ''
    out['dup_ok'] = r2.startswith('ERRO|') and '2025 2025' not in r2 and 'C-A 2025' in r2
    cap = qa.ws('EQA.CAP_Dados')
    aud.novos()
    if lin:
        cap.Activate()
        cap.Range(f'F{lin}').Select()
        qa.ex.esperar()
        cap.Range(f'F{lin}').Value = 106
        qa.ex.esperar()
        tr = aud.novos()
        out['edicao_com_selecao'] = [re.sub(r'\b[A-Z0-9_.-]*\\[^ |]*', '<usr>', x['txt'])[:300] for x in tr]
        out['Z3_depois'] = t(cap.Range('Z3').Value)[:160]
        out['edicao_ok'] = (len(tr) == 1 and tr[0]['acao'] == 'CEQ_ALTERADO' and '104 → 106' in tr[0]['txt']
                            and 'NÃO CONSOLIDADA' in out['Z3_depois'].upper())
        cap.Range('A1').Select()
        out['excluir'] = t(qa.run('mCEQForm.ExcluirCEQ', 'CAP', lin))[:200]
        aud.novos()
    return out


def s_fim(qa, aud):
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
        L['ultimo_recebido'] = str(max((r['DATA_HORA'] for r in fin if isinstance(r['DATA_HORA'], dt.datetime)),
                                       default=None))
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
        secao(L, 'listas_validacao', s_listas, qa)
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


def c_tela(qa, fin):
    p, e = qa.ws('Painel'), qa.ws('Estatística')
    out = {}
    for c, v in (('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro')):
        C.por(qa, p, c, v, local=False)
    p.Range('H3').MergeArea.ClearContents()
    qa.ex.esperar()
    time.sleep(1.5)
    out['painel_A6_J9_texto'] = [[t(p.Cells(r, c).Text) for c in range(1, 11)] for r in range(6, 10)]
    out['A39'] = t(p.Range('A39').Value)
    out['estat_cabecalho'] = {a: t(e.Range(a).Text)[:120] for a in ('A3', 'B3', 'C3', 'D3', 'A4', 'B4', 'C4', 'K5')}
    est = C.estat_tabela(qa)
    out['estat_WBC'] = {k: [v['linha'], v['n'], v['media'], v['cv']] for k, v in est.items() if k.startswith('WBC|')}
    botoes = []
    for ws in qa.wb.Worksheets:
        try:
            for sh in ws.Shapes:
                tx = ''
                for f in (lambda: sh.TextFrame2.TextRange.Text, lambda: sh.TextFrame.Characters().Text,
                          lambda: sh.AlternativeText):
                    try:
                        tx = t(f())
                    except Exception:
                        tx = ''
                    if tx:
                        break
                try:
                    oa = t(sh.OnAction)
                except Exception:
                    oa = ''
                if 'CONSOLID' in (tx + oa).upper():
                    botoes.append({'aba': t(ws.Name), 'nome': t(sh.Name), 'texto': tx[:80], 'macro': oa[-60:],
                                   'visivel': bool(sh.Visible), 'tipo': sh.Type})
        except Exception:
            pass
    out['botoes_consolidar'] = botoes
    out['abas_visiveis'] = [t(w.Name) for w in qa.wb.Worksheets if w.Visible == -1][:60]
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
        criar_usuario(qa, USR, 'ANALISTA')
        L['login'] = fazer_login(qa, USR)
        for w in (p, e):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        secao(L, 'tela', c_tela, qa, fin)
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
    else:
        main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
