# -*- coding: utf-8 -*-
"""reteste01_usuario.py -- Usuario Simulado, reteste do ciclo 01 (ADR-088). Independente dos scripts do Tester.

Uso (um Excel por vez; cada etapa abre e fecha o SEU Excel; nunca em paralelo com outro agente):
    set PYTHONIOENCODING=utf-8
    python -u reteste01_usuario.py <Bioquimica|Hematologia> [etapas]
    etapas (padrao todas, nesta ordem): f1a,f1lc,f1g,prep,f4,calc

Pasta de trabalho: %TEMP%\\qcw\\auditoria\\2026-11-04_reteste01\\usuario\\<produto>\\ (so copias; a producao e
copiada com shutil.copy2 e nunca aberta; o sha256 curto dela e conferido antes e depois de cada etapa).

  f1a   f1_entrada_ciq.py sessao A numa copia nova: congela o DB_SEAC, ATUALIZAR (o 1o depois da entrega), conferencias
        do automatico (A06: resultados <= 0) e as validacoes do manual (V13 = DATA "31/02/2026" colada).
  f1lc  OC-0001: sessao E (Bio: "dimension 1" / "glicose") ou D (Hema: "xn1000" / "wbc"), copia nova.
  f1g   OC-0004 manual: sessao G (N1 "-1" e N2 "0"), copia nova (na Bio tambem, com --produto Bioquimica).
  prep  copia nova -> CAMINHO_DB_SEAC = o congelado da f1a -> ATUALIZAR -> confere IMPLAUSIVEL/A11/A12/participantes
        <= 0 -> salva SO a copia (eventos desligados) -> Painel (como na etapa painel).
  painel copia da copia preparada: Painel no H3 vazio (OC-0013) e com analito SEM EP (OC-0002), pelas celulas de filtro.
  f4    f4_periodos_lotes.py (Bio) ou f4_hema.py (Hema) numa copia da copia preparada (OC-0002, 0003, 0013 e vizinhos).
  calc  calc_independente.py (sem Excel) na copia preparada: OC-0006 (Bilirrubina direta, IG/NRBC) e OC-0009 (B103).
Saidas: <pasta>\\*.json e *.log. Senha: importada da suite de testes, nunca escrita aqui. Nada de caminho de rede.
Autor: AGENTE_USUARIO_SIMULADO, 2026-11-04_reteste01.
"""
import collections
import datetime as dt
import hashlib
import json
import os
import pickle
import shutil
import subprocess
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
PROD = sys.argv[1]
ETAPAS = (sys.argv[2] if len(sys.argv) > 2 else 'f1a,f1lc,f1g,prep,painel,f4,calc').split(',')
W = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste01', 'usuario', PROD)
os.makedirs(W, exist_ok=True)
ORIGEM = os.path.join(REPO, f'QC_{PROD}.xlsm')
ENV = dict(os.environ, PYTHONIOENCODING='utf-8')
BIO = PROD == 'Bioquimica'


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()[:12]


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def copia(nome):
    cp = os.path.join(W, nome)
    shutil.copy2(ORIGEM, cp)
    return cp


def sub(cmd, nome_log, teto=7200):
    with open(os.path.join(W, nome_log), 'w', encoding='utf-8') as lg:
        rc = subprocess.call([sys.executable, '-u'] + cmd, stdout=lg, stderr=subprocess.STDOUT, env=ENV, timeout=teto)
    log('rc', rc, nome_log)
    return rc


F1 = os.path.join(AQUI, 'f1_entrada_ciq.py')
SA = os.path.join(W, 'f1_A.json')
# D/E/G usam so A3_alvo/A3_valores: o reteste refaz o alvo do CICLO 01 (Bio 8977/DIMENSION 1/30-10; Hema 625311/30-10),
# gravado de antemao em alvo_ciclo01.json; sem ele, o alvo da sessao A de hoje (que cai no mes corrente).
SA_ALVO = os.path.join(W, 'alvo_ciclo01.json') if os.path.exists(os.path.join(W, 'alvo_ciclo01.json')) else SA
PREP = os.path.join(W, f'{PROD}_prep.xlsm')


def etapa_prep():
    sys.path.insert(0, AQUI)
    import f1_entrada_ciq as f1
    f1.P.update(f1.PRODUTOS[PROD])
    cp = copia(f'{PROD}_prep.xlsm')
    out = {'origem_sha12': sha(ORIGEM), 'copia_sha12_antes': sha(cp)}
    q = f1.QA(PROD, cp)
    out['excel_pid'] = q.ex.pid
    print('EXCEL_PID', q.ex.pid, flush=True)
    try:
        cong = os.path.join(W, 'DB_SEAC_congelado.xlsm')
        out['dbseac_congelado_sha12'] = sha(cong)
        q.cfg('CAMINHO_DB_SEAC', cong)
        t0 = time.time()
        r, _ = q.atualizar(teto=1800)
        out['atualizar'] = {'s': round(time.time() - t0, 1), 'resumo': str(r)[:600]}
        fin = q.ler('tblCQ_Final')
        out['status'] = dict(collections.Counter(x['STATUS_ANALITICO'] for x in fin))
        out['participa_le0'] = [x['ID_REGISTRO'] for x in fin if x['PARTICIPA_ESTATISTICA'] == 'SIM'
                                and isinstance(x['RESULTADO'], (int, float)) and x['RESULTADO'] <= 0]
        impl = [x for x in fin if x['STATUS_ANALITICO'] == 'IMPLAUSIVEL']
        out['implausivel_amostra'] = [f1.resumo_f(x) for x in impl[:3]]
        out['implausivel_flags'] = dict(collections.Counter((x['PARTICIPA_ESTATISTICA'], x.get('REGISTRAR_LJ', '(sem coluna)'),
                                                             x['TIPO_PLOTAGEM_LJ']) for x in impl))
        qa = f1.qa_tab(q)
        out['qa_codigos'] = dict(collections.Counter(f"{x.get('SEVERIDADE')}|{x.get('CODIGO')}" for x in qa))
        out['qa_a12'] = [x for x in qa if x.get('CODIGO') == 'A12'][:6]
        out['qa_a11_ids'] = sorted({x.get('ID_REGISTRO') for x in qa if x.get('CODIGO') == 'A11'})
        out['a11_igual_implausivel'] = set(out['qa_a11_ids']) == {x['ID_REGISTRO'] for x in impl}
        est = q.ws('Estatística')
        if BIO:
            out['B98_B103'] = [est.Range(f'B{i}').Value for i in range(98, 104)]
            out['A98_A103'] = [est.Range(f'A{i}').Value for i in range(98, 104)]
        lin = {}
        for i in range(14, 100):
            a = est.Cells(i, 1).Value
            if a not in (None, '', 0) and any(k in str(a).upper() for k in ('BILIRRUBINA DIRETA', 'IG', 'NRBC')):
                lin[f'{i}'] = [a, est.Cells(i, 2).Value, est.Cells(i, 7).Value, est.Cells(i, 8).Value,
                               est.Cells(i, 12).Value, est.Cells(i, 29).Value, est.Cells(i, 30).Value]
        out['estat_bili_ig_nrbc'] = lin   # A, B(nivel), G bias, H ET, L Sigma, AC n, AD nrod
        # salva SO a copia (com eventos o Workbook_BeforeSave cancela o Save)
        q.ex.xl.EnableEvents = False
        antes = os.path.getmtime(cp)
        q.ex.wb.Save()
        q.ex.xl.EnableEvents = True
        out['save_ok'] = os.path.getmtime(cp) != antes
        painel_usuario(q, f1, out)
    except Exception as ex:
        out['erro'] = f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2000:]}'
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    out['origem_sha12_depois'] = sha(ORIGEM)
    with open(os.path.join(W, 'prep.json'), 'w', encoding='utf-8') as f:
        json.dump(f1.j(out), f, ensure_ascii=False, indent=1)
    log('prep', {k: out.get(k) for k in ('atualizar', 'status', 'save_ok', 'B98_B103')})


def painel_usuario(q, f1, out):
    """Painel como o usuario: H3 vazio (OC-0013) e analito SEM EP (OC-0002)."""
    # ---- Painel como o usuario: H3 vazio (OC-0013) e analito SEM EP (OC-0002)
    p = q.ws('Painel')
    q.ex.xl.EnableEvents = True            # o Worksheet_Change do Painel e o caminho do usuario
    try:
        p.Unprotect(f1.SENHA)
    except Exception:
        pass
    nl = 2 if BIO else 3
    rng = 'A6:L%d' % (6 + nl)

    def tela(rot):
        v = p.Range(rng).Value
        out[rot] = {'H3': p.Range('H3').Value, 'C3': p.Range('C3').Value, 'L4': p.Range('L4').Value,
                    'G6': p.Range('G6').Value, 'H6': p.Range('H6').Value, 'I6': p.Range('I6').Value,
                    'linhas': [list(x) for x in v[1:]],
                    'H_formula': p.Range('H7').Formula, 'A39': p.Range('A39').Value}
    tela('painel_inicio')
    h3_0 = p.Range('H3').Value
    p.Range('H3').MergeArea.ClearContents()   # H3 e celula mesclada
    q.ex.esperar()
    tela('painel_h3_vazio')
    p.Range('H3').Value = "'" + str(h3_0) if h3_0 not in (None, '') else ''
    q.ex.esperar()
    tela('painel_h3_volta')
    # analito sem CEQ: Bio HbA1c (DIMENSION 2, lote 8590, ADR-088); Hema IG% (SEM EP pelo ADR-088)
    if BIO:
        for cel, v in (('L4', 'DIMENSION 2'), ('C3', 'Hemoglobina glicada'), ('H3', "'8590")):
            p.Range(cel).Value = v
            q.ex.esperar()
    else:
        p.Range('C3').Value = 'IG%'
        q.ex.esperar()
    tela('painel_sem_ep')


def etapa_painel():
    sys.path.insert(0, AQUI)
    import f1_entrada_ciq as f1
    f1.P.update(f1.PRODUTOS[PROD])
    cp = os.path.join(W, f'{PROD}_painel.xlsm')
    shutil.copy2(PREP, cp)
    out = {'copia_de': os.path.basename(PREP), 'sha12': sha(cp)}
    q = f1.QA(PROD, cp)
    out['excel_pid'] = q.ex.pid
    print('EXCEL_PID', q.ex.pid, flush=True)
    try:
        painel_usuario(q, f1, out)
    except Exception as ex:
        out['erro'] = f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2000:]}'
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    with open(os.path.join(W, 'painel.json'), 'w', encoding='utf-8') as f:
        json.dump(f1.j(out), f, ensure_ascii=False, indent=1)
    log('painel', out.get('erro', 'ok'))


def gerar_pkl(xlsm, pkl):
    import openpyxl
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    ws = wb['Principal - Resultados']
    it = ws.iter_rows(min_row=6, values_only=True)
    cab = [c for c in next(it)]
    rows = []
    for r in it:
        if r[0] in (None, ''):
            break
        rows.append(dict(zip(cab, r)))
    wb.close()
    with open(pkl, 'wb') as f:
        pickle.dump({'cab': cab, 'rows': rows}, f)
    return len(rows)


log('inicio', PROD, ETAPAS, 'origem', sha(ORIGEM))
for et in ETAPAS:
    h = sha(ORIGEM)
    log('etapa', et)
    if et == 'f1a':
        sub([F1, 'A', copia(f'{PROD}_F1_A.xlsm'), W, SA, '--produto', PROD], 'f1_A.log')
    elif et == 'f1lc':
        s = 'E' if BIO else 'D'
        sub([F1, s, copia(f'{PROD}_F1_{s}.xlsm'), W, os.path.join(W, f'f1_{s}.json'), SA_ALVO, '--produto', PROD], f'f1_{s}.log')
    elif et == 'f1g':
        sub([F1, 'G', copia(f'{PROD}_F1_G.xlsm'), W, os.path.join(W, 'f1_G.json'), SA_ALVO, '--produto', PROD], 'f1_G.log')
    elif et == 'prep':
        etapa_prep()
    elif et == 'painel':
        etapa_painel()
    elif et == 'f4':
        c4 = os.path.join(W, f'{PROD}_F4.xlsm')
        shutil.copy2(PREP, c4)
        pkl = os.path.join(W, 'cq.pkl')
        log('pkl linhas', gerar_pkl(c4, pkl))
        if BIO:
            sub([os.path.join(AQUI, 'f4_periodos_lotes.py'), c4, pkl, os.path.join(W, 'f4.json')], 'f4.log')
        else:
            sub([os.path.join(AQUI, 'f4_hema.py'), c4, pkl, W], 'f4h.log')
    elif et == 'calc':
        sub([os.path.join(AQUI, 'calc_independente.py'), PREP, '--json', os.path.join(W, 'calc.json')], 'calc.log', 1800)
    if sha(ORIGEM) != h:
        log('ATENCAO: a producao mudou durante a etapa', et)
log('fim', PROD, 'origem', sha(ORIGEM))
