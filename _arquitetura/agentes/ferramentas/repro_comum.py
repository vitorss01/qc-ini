# -*- coding: utf-8 -*-
"""repro_comum.py -- utilitarios do Tester/QA para as reproducoes do ciclo 2026-10-31_ciclo01 (modo REPRODUZIR).

Escrito pelo Tester/QA (nao reaproveita os scripts do Usuario Simulado). Reaproveita so a infraestrutura oficial:
`qa_final.QA` (abre a copia, revela as abas, imprime EXCEL_PID) e `xlh` (watchdog do Application.Run).

Regras: so copias em %TEMP%/qcw/auditoria/2026-10-31_ciclo01/tester/; nada e salvo; fechar() encerra so o Excel que o
QA criou; nenhuma captura de tela.
"""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import statistics as st
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
import qa_final  # noqa: E402

PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-10-31_ciclo01', 'tester')
MOTIVO = 'Outro: TESTE-AUDITORIA'


def senha_planilha():
    """A senha de protecao das abas, lida do instalador (nao repetida neste arquivo versionado)."""
    t = open(os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'camada_dados.py'), encoding='utf-8').read()
    return re.search(r"^SENHA\s*=\s*'([^']+)'", t, re.M).group(1)


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()[:12]


def copiar(produto):
    os.makedirs(PASTA, exist_ok=True)
    src = os.path.join(REPO, f'QC_{produto}.xlsm')
    dst = os.path.join(PASTA, f'{produto}_repro.xlsm')
    h0 = sha(src)
    shutil.copy2(src, dst)
    return src, dst, h0


class Log:
    def __init__(self, saida):
        self.saida = saida
        self.R = {'inicio': dt.datetime.now().isoformat(timespec='seconds')}

    def __setitem__(self, k, v):
        self.R[k] = v
        print('##', k, json.dumps(v, ensure_ascii=False, default=str)[:900], flush=True)
        self.gravar()

    def __getitem__(self, k):
        return self.R[k]

    def gravar(self):
        json.dump(self.R, open(self.saida, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)


# ----------------------------------------------------------------------------------------------- celulas
def por(qa, ws, addr, valor, local=True):
    """Escreve numa celula com os eventos ligados (como o usuario) e espera o Excel."""
    c = ws.Range(addr)
    if local and isinstance(valor, str):
        c.FormulaLocal = valor
    else:
        c.Value = valor
    qa.ex.esperar()
    time.sleep(0.3)


def val(ws, addr):
    v = ws.Range(addr).Value
    return v


def bloco(ws, addr):
    v = ws.Range(addr).Value
    if not isinstance(v, tuple):
        return [[v]]
    return [list(x) for x in v]


def limpo(v):
    if hasattr(v, 'tzinfo') and getattr(v, 'tzinfo', None):
        return v.replace(tzinfo=None)
    return v


# ----------------------------------------------------------------------------------------------- tabelas
def final(qa):
    return qa.ler('tblCQ_Final')


def manuais(qa):
    return [r for r in qa.ler('tblResultados_Manuais') if r.get('ID_REGISTRO')]


def linha_livre_manual(qa):
    lo = qa.lo('tblResultados_Manuais')
    ids = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    dts = lo.ListColumns('DATA').DataBodyRange.Value
    for i in range(len(ids)):
        if ids[i][0] in (None, '') and dts[i][0] in (None, ''):
            return lo, i + 1
    raise RuntimeError('tblResultados_Manuais cheia')


def escrever_manual(qa, campos):
    """campos: lista ordenada [(coluna, texto_digitado)]. Escreve celula a celula com FormulaLocal e eventos ligados
    (o Worksheet_Change carimba ID, EQUIPAMENTO padrao, MATRIZ e REGISTRADO_EM). Devolve (linha, ID, linha lida)."""
    lo, r = linha_livre_manual(qa)
    for col, v in campos:
        cel = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
        if isinstance(v, str):
            cel.FormulaLocal = v
        else:
            cel.Value = v
        qa.ex.esperar()
    cab = [c.Name for c in lo.ListColumns]
    lida = dict(zip(cab, [limpo(x) for x in lo.ListRows(r).Range.Value[0]]))
    return r, lida.get('ID_REGISTRO'), lida


def validacao(qa, coluna, texto):
    """Veredicto da validacao de dados do Excel (Range.Validation.Value) para o texto digitado numa linha livre.
    Eventos desligados; a celula e limpa depois."""
    lo, r = linha_livre_manual(qa)
    cel = lo.ListColumns(coluna).DataBodyRange.Cells(r, 1)
    qa.ex.xl.EnableEvents = False
    try:
        cel.FormulaLocal = texto
        try:
            ok = bool(cel.Validation.Value)
        except Exception as e:
            ok = f'sem validacao ({e.__class__.__name__})'
        cel.ClearContents()
    finally:
        qa.ex.xl.EnableEvents = True
    return ok


def qa_integracao(qa, ids):
    out = []
    for r in qa.ler('tblQA_Integracao'):
        txt = ' '.join(str(v) for v in r.values() if v is not None)
        if any(i in txt for i in ids):
            out.append(r)
    return out


def estat_tabela(qa, ultima=133):
    """Tabela principal da Estatistica (1a ocorrencia de cada analito|nivel), colunas A..AU."""
    est = qa.ws('Estatística')
    v = est.Range(f'A14:AU{ultima}').Value
    out = {}
    for i, l in enumerate(v, 14):
        if l[0] and isinstance(l[1], (int, float)):
            k = f'{l[0]}|{int(l[1])}'
            if k not in out:
                out[k] = {'linha': i, 'n': l[2], 'media': l[3], 'dp': l[4], 'cv': l[5], 'G_bias': l[6], 'H_et': l[7],
                          'K_etp': l[10], 'L_sigma': l[11], 'M_status': l[12], 'T_sinal': l[19], 'AC_nEQA': l[28]}
    return out


def painel(qa, nlv):
    p = qa.ws('Painel')
    v = bloco(p, f'A7:J{6 + nlv}')
    cab = ['nivel', 'n', 'media', 'dp', 'cv', 'etp', 'bias', 'et', 'sigma', 'status']
    filtro = {c: limpo(val(p, c)) for c in ('C3', 'H3', 'K3', 'K4', 'G3', 'G4', 'L4', 'M4', 'H4')}
    return {'filtro': filtro, 'niveis': [dict(zip(cab, l)) for l in v], 'A39': val(p, 'A39')}


def eng_runs(qa, runs, nlv):
    eng = qa.ws('Eng_Saida')
    n = int(eng.Range('I1').Value or 0)
    if n == 0:
        return {}
    cab = [str(c) for c in bloco(eng, 'A2:AL2')[0]]
    linhas = bloco(eng, f'A3:AL{2 + n}')
    out = {}
    for l in linhas:
        try:
            rn = int(l[1])
        except Exception:
            continue
        if rn in runs:
            out[rn] = {cab[i] if cab[i] not in ('None', '') else f'c{i + 1}': limpo(x) for i, x in enumerate(l)
                       if x not in (None, '')}
    return out


def mdcv(v):
    v = [float(x) for x in v]
    if len(v) < 2:
        return {'n': len(v), 'media': (v[0] if v else None)}
    m = st.fmean(v)
    s = st.stdev(v)
    return {'n': len(v), 'media': m, 'dp': s, 'cv': s / abs(m) * 100}


def pop(linhas, analito, lote, nivel, equip=None, ini=None, fim=None, maiusc=False):
    """Populacao independente da tblCQ_Final (PARTICIPA=SIM). maiusc=True compara analito/equipamento sem caixa
    (como o motor); False compara exato (como o ETL)."""
    def eq(a, b):
        return (str(a).strip().upper() == str(b).strip().upper()) if maiusc else (a == b)
    out = []
    for r in linhas:
        if r.get('PARTICIPA_ESTATISTICA') != 'SIM':
            continue
        if not eq(r['ANALITO'], analito) or str(r['LOTE']) != str(lote) or int(r['NIVEL'] or 0) != nivel:
            continue
        if equip and not eq(r['EQUIPAMENTO'], equip):
            continue
        d = r['DATA']
        if ini and (not isinstance(d, dt.datetime) or d < ini):
            continue
        if fim and (not isinstance(d, dt.datetime) or d > fim):
            continue
        out.append(r)
    return out


def um_por_run(linhas):
    """Reconstrucao "um valor por RUN e nivel" (o Eng_Saida guarda um so): o ultimo pela DATA_HORA vence."""
    d = {}
    for r in sorted(linhas, key=lambda x: (x['DATA_HORA'] or dt.datetime.min)):
        d[r['RUN']] = r['RESULTADO']
    return list(d.values())


def rel(a, b):
    try:
        return abs(a - b) / max(abs(a), abs(b), 1e-300)
    except Exception:
        return None


def campos_ceq(prov, rod, ano, an, am, res, alvo='100', sd='2', sdi='5', linf='90', lsup='110', aval='Acceptable',
               nlabs='50', unid='', pag='', arq='TESTE-AUDITORIA.pdf'):
    return [prov, rod, ano, an, am, res, alvo, sd, sdi, linf, lsup, aval, nlabs, unid, pag, arq]


def eqa_teste(qa):
    """Linhas da EQA_Base com TESTE-AUDITORIA na chave."""
    b = qa.ws('EQA_Base')
    v = bloco(b, 'A1:U5001')
    cab = [str(c) for c in v[0]]
    out = [dict(zip(cab, l)) for l in v[1:] if l[0] and 'TESTE-AUDITORIA' in str(l[20])]
    return out, b.Range('Z1').Value
