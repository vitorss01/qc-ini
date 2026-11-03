# -*- coding: utf-8 -*-
"""F1 do Usuario Simulado -- entrada de CIQ manual e automatica, e equivalencia entre os dois caminhos.

Uso (um Excel por vez; nunca em paralelo com outro agente que use Excel):
    python f1_entrada_ciq.py A <copia_A.xlsm> <pasta_trabalho> <saida_A.json>
    python f1_entrada_ciq.py B <copia_B.xlsm> <pasta_trabalho> <saida_B.json> <saida_A.json>
    python f1_entrada_ciq.py C <copia_C.xlsm> <pasta_trabalho> <saida_C.json> BIO-<x>,BIO-<y>
      (C = SUBSTITUI_ID, ADR-082: x substituido com MOTIVO, y sem MOTIVO; precisa do DB_SEAC congelado da sessao A)
    python f1_entrada_ciq.py D <copia_D.xlsm> <pasta_trabalho> <saida_D.json> <saida_A.json> --produto Hematologia
      (D = equipamento/analito digitado em minusculas: RUN x motor do Painel; acrescentada na F1 da Hema)
    python f1_entrada_ciq.py E <copia_E.xlsm> <pasta_trabalho> <saida_E.json> <saida_A.json>
      (E = a mesma pergunta da D na Bioquimica: "dimension 1" / "glicose" num manual N2; 3o passe da F1 da Bio)
    python f1_entrada_ciq.py F <copia_F.xlsm> <pasta_trabalho> <saida_F.json> <saida_A.json> --produto Hematologia
      (F = o que a validacao de dados do Excel aceita quando o usuario DIGITA (Range.Validation.Value), e manuais
       com espaco sobrando no analito/equipamento; 2o passe da F1 da Hema)
    python f1_entrada_ciq.py G <copia_G.xlsm> <pasta_trabalho> <saida_G.json> <saida_A.json> --produto Hematologia
      (G = resultado implausivel digitado no manual: -1 e 0, MOTIVO com a marca TESTE-AUDITORIA; 2o passe da Hema)

O que cada sessao faz (sempre numa COPIA da producao; nada e salvo; o Excel e encerrado so pelo PID criado aqui):
  A  1. le as telas de entrada (aba Digitar Resultados: colunas, validacao de dados, formato, bloqueio);
     2. AUTOMATICO: congela o DB_SEAC (copia de arquivo, somente leitura da origem) em <pasta_trabalho>,
        aponta CAMINHO_DB_SEAC da copia para ele (como o administrador faria) e roda o ATUALIZAR
        (mIntegracao.AtualizarDadosAutomatico = nucleo do botao, sem dialogo). Confere, contra a origem lida
        por openpyxl: todo ID da origem recebido como BIO-<id>, as colunas de cada ID novo, a associacao
        equipamento/analito/lote/nivel/data/hora/valor na tblCQ_Final, duplicidade, Sync_Log e QA;
        uma 2a atualizacao com o mesmo arquivo (idempotencia);
     3. MANUAL valido (N1 e N2 do mesmo analito/lote/equipamento, numa hora livre de um dia ja passado),
        digitado na tblResultados_Manuais celula a celula com os eventos ligados; ATUALIZAR; confere a linha
        gravada, a tblCQ_Final, o Painel (engPainel), a Estatistica, o Eng_Saida (ponto do LJ e Westgard) e
        recalcula n/media/DP/CV de forma independente (statistics);
     4. VALIDACOES: vazio, texto, virgula/ponto, data futura, analito/nivel/lote/equipamento inexistente,
        duplicidade manual x manual e manual x interfaceamento, hora/data invalidas, equipamento vazio.
  B  mesmo DB_SEAC congelado (linha de base) e depois um DB_SEAC SINTETICO com os MESMOS dois resultados do
     passo A3 (IDs 9000001/9000002, clonando um resultado real do mesmo analito/lote/nivel); confere que a
     estatistica, o RUN, a corrida e o Westgard sao os mesmos do caminho manual.

Produto: --produto Bioquimica (padrao, comportamento identico ao do 1o/2o passe da Bio) ou --produto Hematologia
(acrescentado no ciclo 2026-10-31_ciclo01 para a F1 da Hema). Diferencas da Hema: prefixo HEM, analito WBC, 3 niveis,
um equipamento so (EQUIPAMENTO_PADRAO; o Painel nao tem seletor L4), a Estatistica filtra por Ano De/Ate (B3/D3) e
pelo lote B4 (nao pelo periodo do Painel) e o Painel com Periodo "Tudo" usa o historico inteiro do lote; V07 usa o
nivel 4; V11 usa o equipamento "XN9999" (espera-se o ALERTA A10); V15 = equipamento vazio recebe o padrao.
    python f1_entrada_ciq.py A <copia> <pasta> <saida.json> --produto Hematologia

Marcacao dos dados de teste: MOTIVO "Outro: TESTE-AUDITORIA" e valores com final ,111 / ,222 (sinteticos).
Na Hema (resultados com ate 2 casas): media arredondada a 1 casa + 0,011 / 0,022 / 0,033 (3 casas = sintetico).
O caminho do DB_SEAC nunca e impresso (vira <UNC>). Senha: importada da suite de testes, nunca escrita aqui.
Autor: AGENTE_USUARIO_SIMULADO, ciclo 2026-10-31_ciclo01.
"""
import collections
import datetime as dt
import hashlib
import json
import math
import os
import shutil
import statistics
import sys
import time
import traceback
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
TESTES = os.path.join(REPO, '_arquitetura', 'integracao_dbseac', 'testes')
sys.path.insert(0, TESTES)
from qa_final import QA  # noqa: E402
from qa_etl_recebimento import SENHA  # noqa: E402
import gerar_seac_sintetico as gss  # noqa: E402

MARCA = 'Outro: TESTE-AUDITORIA'
EPOCH = dt.datetime(1899, 12, 30)
OUT = {'passos': [], 'erros': []}

PRODUTOS = {
    'Bioquimica': {'nome': 'Bioquimica', 'pref': 'BIO', 'analito': 'Glicose', 'nlv': 2, 'bio': True},
    'Hematologia': {'nome': 'Hematologia', 'pref': 'HEM', 'analito': 'WBC', 'nlv': 3, 'bio': False},
}
P = dict(PRODUTOS['Bioquimica'])


def valor_sintetico(media, n):
    """Bio: media inteira + 0,111 * n (comportamento original). Hema: media com 1 casa + 0,011 * n."""
    if P['bio']:
        return round(media, 0) + (0.111 if n == 1 else 0.222 if n == 2 else 0.333)
    return round(round(media, 1) + 0.011 * n, 3)


def txt_valor(v):
    return f'{v:.3f}'.replace('.', ',')


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def serie(d):
    if isinstance(d, dt.datetime):
        d = d.replace(tzinfo=None)
        return (d - EPOCH).total_seconds() / 86400.0
    return float((dt.datetime(d.year, d.month, d.day) - EPOCH).days)


def ingenuo(v):
    if isinstance(v, dt.datetime):
        return dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second, v.microsecond)
    return v


def _chaves(v):
    if isinstance(v, dict):
        return {(k if isinstance(k, (str, int, float, bool)) or k is None else ' | '.join(map(str, k))
                 if isinstance(k, tuple) else str(k)): _chaves(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_chaves(x) for x in v]
    return v


def j(v):
    return json.loads(json.dumps(_chaves(v), ensure_ascii=False, default=str))


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()[:12]


# ------------------------------------------------------------------ origem congelada
def congelar_dbseac(caminho, destino):
    for t in range(8):
        try:
            shutil.copy2(caminho, destino + '.tmp')
            with zipfile.ZipFile(destino + '.tmp') as z:
                z.testzip()
            os.replace(destino + '.tmp', destino)
            return
        except Exception as e:
            log('copia do DB_SEAC falhou (tentativa', t + 1, '):', type(e).__name__)
            time.sleep(15)
    raise RuntimeError('nao consegui congelar o DB_SEAC')


def ler_origem(caminho, tabela):
    import pickle
    cache = caminho + '.origem.pkl'
    if os.path.exists(cache) and os.path.getmtime(cache) > os.path.getmtime(caminho):
        with open(cache, 'rb') as f:
            return pickle.load(f)
    r = _ler_origem(caminho, tabela)
    with open(cache, 'wb') as f:
        pickle.dump(r, f)
    return r


def _ler_origem(caminho, tabela):
    import openpyxl
    from openpyxl.utils import range_boundaries
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=False, keep_vba=False, keep_links=False)
    try:
        for ws in wb.worksheets:
            if tabela in ws.tables:
                c0, l0, c1, l1 = range_boundaries(ws.tables[tabela].ref)
                linhas = list(ws.iter_rows(min_row=l0, max_row=l1, min_col=c0, max_col=c1, values_only=True))
                cab = [str(c).strip() if c is not None else '' for c in linhas[0]]
                return cab, [dict(zip(cab, l)) for l in linhas[1:]]
    finally:
        wb.close()
    raise RuntimeError('tabela da origem nao encontrada')


# ------------------------------------------------------------------ leitura da copia
def ler_tab(q, nome):
    return q.ler(nome)


def sync_ultimas(q, k=3):
    sl = q.ws('Sync_Log')
    cab = [sl.Cells(3, c).Value for c in range(1, 29)]
    ult = sl.Cells(sl.Rows.Count, 1).End(-4162).Row
    out = []
    for r in range(max(4, ult - k + 1), ult + 1):
        v = sl.Range(sl.Cells(r, 1), sl.Cells(r, 28)).Value[0]
        out.append({c: (ingenuo(x) if isinstance(x, dt.datetime) else x) for c, x in zip(cab, v)})
    return out


def audit_desde(q, linha0):
    al = q.ws('Audit_Log')
    ult = al.Cells(al.Rows.Count, 1).End(-4162).Row
    if ult <= linha0:
        return ult, []
    cab = [al.Cells(1, c).Value for c in range(1, 15)]
    # cabecalho pode nao estar na linha 1: guarda as linhas cruas
    v = al.Range(al.Cells(linha0 + 1, 1), al.Cells(ult, 14)).Value
    return ult, [[(str(x) if x is not None else None) for x in r] for r in v]


def audit_ult(q):
    al = q.ws('Audit_Log')
    return al.Cells(al.Rows.Count, 1).End(-4162).Row


def qa_tab(q):
    return [{k: (str(v) if v is not None else None) for k, v in r.items()} for r in q.ler('tblQA_Integracao')]


def painel(q):
    p = q.ws('Painel')
    d = {a: p.Range(a).Value for a in ('B3', 'C3', 'G3', 'G4', 'H3', 'H4', 'K3', 'K4', 'L4', 'J7', 'J8', 'J9')}
    d = {k: (ingenuo(v) if isinstance(v, dt.datetime) else v) for k, v in d.items()}
    try:
        d['engPainel'] = [list(l) for l in q.wb.Names('engPainel').RefersToRange.Value]
    except Exception as e:
        d['engPainel'] = f'erro {e}'
    return d


def estat_linhas(q, analito):
    est = q.ws('Estatística')
    out = []
    for linha in est.Range('A14:H133' if P['bio'] else 'A14:H133').Value:
        if str(linha[0] or '').strip().upper() == analito.upper():
            out.append(list(linha))
    filtro = {}
    if not P['bio']:
        for a in ('B3', 'D3', 'B4'):
            filtro['Estatística!' + a] = est.Range(a).Value
        return {'linhas': out, 'filtro': filtro}
    for n in ('Estat_Lote', 'Estat_Ini_Efetiva', 'Estat_Fim_Efetiva'):
        try:
            v = q.wb.Names(n).RefersToRange.Value
        except Exception:
            v = q.wb.Application.Evaluate(n)
        filtro[n] = ingenuo(v) if isinstance(v, dt.datetime) else v
    return {'linhas': out, 'filtro': filtro}


def eng(q):
    e = q.ws('Eng_Saida')
    n = int(e.Range('I1').Value or 0)
    cab = [e.Cells(2, c).Value for c in range(1, 40)]
    linhas = []
    if n:
        v = e.Range(e.Cells(3, 1), e.Cells(2 + n, 39)).Value
        linhas = [list(r) for r in v]
    return {'n': n, 'cab': cab, 'linhas': linhas, 'C1': e.Range('C1').Value, 'E1': e.Range('E1').Value,
            'S1': e.Range('S1').Value}


def eng_run(q, run):
    e = eng(q)
    for r in e['linhas']:
        try:
            if int(r[1]) == int(run):
                return {'cab': e['cab'], 'linha': r}
        except (TypeError, ValueError):
            pass
    return None


def stats_ind(vals):
    n = len(vals)
    if n == 0:
        return {'n': 0}
    m = statistics.fmean(vals)
    sd = statistics.stdev(vals) if n > 1 else None
    return {'n': n, 'media': m, 'dp': sd, 'cv': (sd / m * 100.0) if (sd is not None and m) else None}


def pop(fin, analito, lote, equip, nivel, ini, fim, anos=None):
    """anos=(de, ate): filtro da Estatistica da Hema (ANO de DATA_HORA entre B3 e D3, src/hema EstatBasica)."""
    out = []
    for r in fin:
        if (r['ANALITO'] == analito and str(r['LOTE']) == str(lote) and r['EQUIPAMENTO'] == equip
                and int(r['NIVEL'] or 0) == nivel and r['PARTICIPA_ESTATISTICA'] == 'SIM'):
            if anos is not None:
                dh = r['DATA_HORA']
                if hasattr(dh, 'year') and anos[0] <= dh.year <= anos[1] and isinstance(r['RESULTADO'], (int, float)):
                    out.append(float(r['RESULTADO']))
                continue
            d = r['DATA']
            if hasattr(d, 'year'):
                s = serie(ingenuo(d))
                if (ini is None or s >= ini) and (fim is None or s < int(fim) + 1):
                    out.append(float(r['RESULTADO']))
    return out


def niveis():
    return tuple(range(1, P['nlv'] + 1))


def set_painel(q, lote=None, equip=None, analito=None, k3=None, k4=None, ano_estat=None):
    """Muda os filtros do Painel como o usuario (celula de entrada, eventos ligados). Hema: sem L4 (equipamento
    unico); ano_estat preenche Ano De/Ate (B3/D3) da aba Estatistica, que e o filtro dela."""
    p = q.ws('Painel')
    try:
        p.Unprotect(SENHA)
    except Exception:
        pass
    q.ex.xl.EnableEvents = True
    for cel, v in (('H3', lote), ('L4', equip if P['bio'] else None), ('C3', analito), ('K3', k3), ('K4', k4)):
        if v is not None:
            p.Range(cel).Value = v
            q.ex.esperar()
    if ano_estat is not None and not P['bio']:
        est = q.ws('Estatística')
        try:
            est.Unprotect(SENHA)
        except Exception:
            pass
        est.Range('B3').Value = ano_estat
        est.Range('D3').Value = ano_estat
    try:
        q.run('mEstatistica.AtualizarEstatistica')
        if P['bio']:
            q.run('RecalcularEstatPeriodo')
    except Exception as e:
        OUT['erros'].append(f'recalculo do painel: {e}')


# ------------------------------------------------------------------ escrita na aba Digitar Resultados
def linha_livre(q):
    lo = q.lo('tblResultados_Manuais')
    a = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
    b = lo.ListColumns('DATA').DataBodyRange.Value
    for i in range(len(a)):
        if a[i][0] in (None, '') and b[i][0] in (None, ''):
            return i + 1
    raise RuntimeError('tabela manual cheia')


def digitar(q, campos):
    """campos: lista de (coluna, valor, modo) com modo 'v' (Value) ou 'l' (FormulaLocal = o que o usuario digita
    no Excel pt-BR). Uma celula por vez, eventos ligados (o Worksheet_Change carimba)."""
    q.ex.xl.EnableEvents = True
    lo = q.lo('tblResultados_Manuais')
    r = linha_livre(q)
    for col, v, modo in campos:
        cel = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
        if modo == 'l':
            cel.FormulaLocal = v
        else:
            cel.Value = v
    return r


def linha_manual(q, r):
    lo = q.lo('tblResultados_Manuais')
    cols = [c.Name for c in lo.ListColumns]
    v = lo.DataBodyRange.Rows(r).Value[0]
    d = {}
    for c, x in zip(cols, v):
        d[c] = ingenuo(x) if isinstance(x, dt.datetime) else x
    # o que a celula mostra (texto formatado)
    d['_texto_RESULTADO'] = lo.ListColumns('RESULTADO').DataBodyRange.Cells(r, 1).Text
    d['_fmt_RESULTADO'] = lo.ListColumns('RESULTADO').DataBodyRange.Cells(r, 1).NumberFormat
    d['_texto_DATA'] = lo.ListColumns('DATA').DataBodyRange.Cells(r, 1).Text
    d['_texto_HORA'] = lo.ListColumns('HORA').DataBodyRange.Cells(r, 1).Text
    return d


def final_por_id(q):
    return {r['ID_REGISTRO']: r for r in q.ler('tblCQ_Final')}


CAMPOS_F = ('ID_REGISTRO', 'ORIGEM_RESULTADO', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'ANALITO_CADASTRADO',
            'DATA', 'DATA_HORA', 'RESULTADO', 'RUN', 'CORRIDA_NO_DIA', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
            'TIPO_PLOTAGEM_LJ', 'MOTIVO_EXCLUSAO_AUTOMATICA', 'ID_RELACIONADO', 'MOTIVO_MANUAL', 'USUARIO_MANUAL')


def resumo_f(r):
    if r is None:
        return None
    return {k: (ingenuo(r.get(k)) if isinstance(r.get(k), dt.datetime) else r.get(k)) for k in CAMPOS_F}


def atualizar(q, rotulo):
    t0 = time.time()
    try:
        r, _ = q.atualizar(teto=1800)
        erro = None
    except TimeoutError:
        raise
    except Exception as e:
        r, erro = None, str(e)
    t = round(time.time() - t0, 1)
    OUT['passos'].append({'atualizacao': rotulo, 's': t, 'erro': erro, 'resumo': (r or '')[:600]})
    log('ATUALIZAR', rotulo, t, 's', 'OK' if erro is None else 'ERRO ' + erro[:300])
    return erro, r, t


def inspecionar_telas(q):
    ws = q.ws('Digitar Resultados')
    lo = q.lo('tblResultados_Manuais')
    info = {}
    for col in ('DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO', 'MOTIVO', 'SUBSTITUI_ID'):
        try:
            v = lo.ListColumns(col).DataBodyRange.Cells(1, 1).Validation
            d = {'tipo': v.Type}
            for a in ('Operator', 'Formula1', 'Formula2', 'ErrorMessage', 'ShowError'):
                try:
                    d[a] = getattr(v, a)
                except Exception:
                    d[a] = None
        except Exception as e:
            d = f'sem validacao ({type(e).__name__})'
        info[col + '_validacao'] = d
    info['instrucoes'] = [ws.Range(f'A{i}').Value for i in range(1, 5)]
    for nome in ('lstLotes', 'lstAnalitos'):
        try:
            rg = q.wb.Names(nome).RefersToRange
            info[nome] = [x[0] for x in rg.Value if x[0] not in (None, '')]
        except Exception as e:
            info[nome] = f'erro {e}'
    # formulas de RUN/STATUS/DETALHE (como a linha mostra o que foi salvo)
    for c in ('RUN', 'STATUS', 'DETALHE'):
        info['formula_' + c] = lo.ListColumns(c).DataBodyRange.Cells(1, 1).Formula
    info['protegida'] = ws.ProtectContents
    info['congelar_paineis'] = None
    return info


# ================================================================== sessao A
def sessao_a(copia, pasta, saida):
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        OUT['telas'] = inspecionar_telas(q)
        cam = q.cfg('CAMINHO_DB_SEAC')
        tabela = q.cfg('TABELA_ORIGEM')
        OUT['cfg'] = {k: q.cfg(k) for k in ('MODO_FONTE', 'MOTOR_DADOS', 'PREFIXO_ID', 'EQUIPAMENTO_PADRAO',
                                             'MATRIZ_PADRAO', 'GAP_CORRIDA_MIN', 'TOLERANCIA_CONFLITO_MIN', 'NIVEIS')}
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        if not os.path.exists(congelado):
            congelar_dbseac(cam, congelado)
        OUT['dbseac_congelado'] = {'arquivo': os.path.basename(congelado), 'sha12': sha(congelado),
                                   'origem': '<UNC>', 'mtime': time.ctime(os.path.getmtime(congelado))}
        log('origem congelada', OUT['dbseac_congelado'])
        cab_o, orig = ler_origem(congelado, tabela)
        log('origem lida:', len(orig), 'linhas')
        rec0 = q.ler('tblDB_Recebimento')
        fin0 = final_por_id(q)
        al0 = audit_ult(q)
        q.cfg('CAMINHO_DB_SEAC', congelado)          # o administrador aponta para a copia congelada
        erro1, res1, t1 = atualizar(q, 'A1 DB_SEAC congelado')
        rec1 = q.ler('tblDB_Recebimento')
        fin1 = final_por_id(q)
        sync1 = sync_ultimas(q, 1)
        qa1 = qa_tab(q)
        al1, aud1 = audit_desde(q, al0)
        # ---------------- conferencia contra a origem
        pref = P['pref']
        ids_o = {}
        sem_id_o = []
        for o in orig:
            io = o.get('ID_ORIGEM')
            try:
                k = f'{pref}-{int(float(io))}'
                if float(io) != int(float(io)):
                    raise ValueError
                ids_o.setdefault(k, []).append(o)
            except (TypeError, ValueError):
                sem_id_o.append({c: o.get(c) for c in ('ANALITO', 'NIVEL', 'DATA_HORA', 'VALOR', 'ID_ORIGEM')})
        rec0_ids = {r['ID_REGISTRO'] for r in rec0}
        rec1_por = collections.defaultdict(list)
        for r in rec1:
            rec1_por[r['ID_REGISTRO']].append(r)
        novos = sorted(set(rec1_por) - rec0_ids)
        nao_recebidos = sorted(set(ids_o) - set(rec1_por))
        dup_o = {k: len(v) for k, v in ids_o.items() if len(v) > 1}
        dup_rec = {k: len(v) for k, v in rec1_por.items() if len(v) > 1}
        # colunas de cada novo: origem x recebimento
        def norm(v):
            v = ingenuo(v)
            if isinstance(v, dt.datetime):
                return v.replace(microsecond=0).isoformat()
            if isinstance(v, dt.date):
                return dt.datetime(v.year, v.month, v.day).isoformat()
            if isinstance(v, bool):
                return str(v)
            if isinstance(v, (int, float)) and float(v).is_integer():
                return str(int(v))
            if isinstance(v, float):
                return repr(v)
            if v is None or (isinstance(v, str) and v.strip() == ''):
                return None
            return str(v).strip()
        difs_rec = []
        for k in novos:
            o = ids_o.get(k, [None])[0]
            if o is None:
                difs_rec.append({'id': k, 'problema': 'novo no recebimento sem linha na origem'})
                continue
            r = rec1_por[k][0]
            for c in cab_o:
                if not c:
                    continue
                if norm(o.get(c)) != norm(r.get(c)):
                    difs_rec.append({'id': k, 'coluna': c, 'origem': norm(o.get(c)), 'recebido': norm(r.get(c))})
        # associacao na final
        difs_fin, st_novos, mapa_an = [], collections.Counter(), collections.Counter()
        exemplos = []
        for k in novos:
            o = ids_o.get(k, [None])[0]
            f = fin1.get(k)
            if f is None:
                difs_fin.append({'id': k, 'problema': 'recebido e ausente na tblCQ_Final'})
                continue
            st_novos[f['STATUS_ANALITICO']] += 1
            mapa_an[(o.get('ANALITO'), f['ANALITO'], f['ANALITO_CADASTRADO'])] += 1
            pares = [('EQUIPAMENTO', 'EQUIPAMENTO'), ('LOTE', 'LOTE'), ('NIVEL', 'NIVEL'), ('ANALITO', 'ANALITO_ORIGEM'),
                     ('DATA', 'DATA'), ('DATA_HORA', 'DATA_HORA'), ('VALOR', 'RESULTADO'), ('MATRIZ', 'MATRIZ'),
                     ('ID_ORIGEM', 'ID_ORIGEM'), ('ID_AMOSTRA', 'ID_AMOSTRA')]
            for co, cf in pares:
                if co not in cab_o:
                    continue
                a, b = norm(o.get(co)), norm(f.get(cf))
                if co == 'VALOR' and f['STATUS_ANALITICO'] == 'SEM_VALOR':
                    continue
                if a != b:
                    difs_fin.append({'id': k, 'origem': co, 'final': cf, 'v_origem': a, 'v_final': b})
            if f['ORIGEM_RESULTADO'] != 'INTERFACEAMENTO':
                difs_fin.append({'id': k, 'problema': 'ORIGEM_RESULTADO', 'v': f['ORIGEM_RESULTADO']})
            if len(exemplos) < 3:
                exemplos.append({'origem': {c: norm(o.get(c)) for c in cab_o if c}, 'final': resumo_f(f)})
        # linhas antigas intactas
        def tup(r):
            return tuple(norm(r.get(c)) for c in [c for c in rec0[0].keys()])
        antes = collections.Counter(tup(r) for r in rec0)
        depois = collections.Counter(tup(r) for r in rec1 if r['ID_REGISTRO'] in rec0_ids)
        mudaram = list((antes - depois).elements())[:3]
        # status/RUN dos antigos
        mud_fin = [k for k, r in fin0.items() if k in fin1 and (fin1[k]['RUN'], fin1[k]['STATUS_ANALITICO']) != (r['RUN'], r['STATUS_ANALITICO'])]
        nao_ativos = [resumo_f(fin1[k]) for k in novos if k in fin1 and fin1[k]['STATUS_ANALITICO'] != 'ATIVO'][:15]
        dh_o = [ingenuo(o.get('DATA_HORA')) for o in orig if isinstance(o.get('DATA_HORA'), dt.datetime)]
        # plausibilidade (achado OC-1 da Bio): valor <= 0 ativo na estatistica, nos novos e na base inteira
        def nao_pos(r):
            v = r.get('RESULTADO')
            return isinstance(v, (int, float)) and v <= 0 and r.get('PARTICIPA_ESTATISTICA') == 'SIM'
        OUT['A1_plausibilidade'] = {
            'novos_ativos_valor_nao_positivo': [resumo_f(fin1[k]) for k in novos if k in fin1 and nao_pos(fin1[k])][:20],
            'n_base_ativos_valor_nao_positivo': sum(1 for r in fin1.values() if nao_pos(r)),
            'base_amostra': [resumo_f(r) for r in fin1.values() if nao_pos(r)][:20],
            'equipamentos_novos': dict(collections.Counter(fin1[k]['EQUIPAMENTO'] for k in novos if k in fin1)),
            'equipamentos_base': dict(collections.Counter(r['EQUIPAMENTO'] for r in fin1.values()))}
        OUT['A1_automatico'] = {
            'erro': erro1, 'resumo': res1, 't_s': t1,
            'origem_linhas': len(orig), 'origem_ids_validos': len(ids_o), 'origem_sem_id': sem_id_o[:5],
            'n_origem_sem_id': len(sem_id_o), 'origem_dh_max': max(dh_o) if dh_o else None,
            'origem_colunas': cab_o, 'ids_repetidos_origem': dict(list(dup_o.items())[:10]), 'n_ids_repetidos_origem': len(dup_o),
            'recebimento': [len(rec0), len(rec1)], 'novos': len(novos), 'novos_amostra': novos[:5] + novos[-5:],
            'nao_recebidos': nao_recebidos[:10], 'n_nao_recebidos': len(nao_recebidos),
            'dup_recebimento': dup_rec, 'divergencias_origem_x_recebimento': difs_rec[:10], 'n_div_rec': len(difs_rec),
            'divergencias_origem_x_final': difs_fin[:10], 'n_div_fin': len(difs_fin),
            'status_novos': dict(st_novos), 'nao_ativos_novos': nao_ativos,
            'mapa_analito_origem_final': [list(k) + [v] for k, v in mapa_an.items()],
            'antigos_alterados_no_recebimento': mudaram, 'n_antigos_RUN_STATUS_mudaram': len(mud_fin),
            'antigos_RUN_STATUS_mudaram_amostra': [(k, (fin0[k]['RUN'], fin0[k]['STATUS_ANALITICO']), (fin1[k]['RUN'], fin1[k]['STATUS_ANALITICO'])) for k in mud_fin[:8]],
            'final': len(fin1), 'final_ids_unicos': len(set(fin1)),
            'sync_log': sync1, 'qa_codigos': dict(collections.Counter((r.get('SEVERIDADE'), r.get('CODIGO')) for r in qa1)),
            'qa_amostra': qa1[:12], 'audit_novas': len(aud1), 'audit_acoes': dict(collections.Counter(r[2] for r in aud1 if r)),
            'exemplos': exemplos}
        # idempotencia
        erro2, res2, t2 = atualizar(q, 'A2 mesmo arquivo')
        rec2 = q.ler('tblDB_Recebimento')
        fin2 = final_por_id(q)
        est1 = {k: (r['RUN'], r['STATUS_ANALITICO'], r['RESULTADO']) for k, r in fin1.items()}
        est2 = {k: (r['RUN'], r['STATUS_ANALITICO'], r['RESULTADO']) for k, r in fin2.items()}
        OUT['A2_idempotencia'] = {'erro': erro2, 'resumo': res2, 't_s': t2, 'recebimento': len(rec2),
                                  'final_igual': est1 == est2, 'sync_log': sync_ultimas(q, 1)}

        # ---------------- MANUAL valido
        fin = list(fin2.values())
        analito = P['analito']
        out_m = alvo_manual(q, fin, analito)
        OUT['A3_alvo'] = out_m
        lote, equip, d_alvo, hora_alvo = out_m['lote'], out_m['equip'], out_m['dia'], out_m['hora']
        anos = None if P['bio'] else (d_alvo.year, d_alvo.year)
        set_painel(q, lote=lote, equip=equip, analito=analito, k3=d_alvo.year,
                   k4=None if P['bio'] else 'Tudo (sem filtro)', ano_estat=d_alvo.year)
        p0 = painel(q)
        e0 = estat_linhas(q, analito)
        OUT['A3_painel_antes'] = p0
        OUT['A3_estat_antes'] = e0
        ini = serie(p0['G3']) if isinstance(p0['G3'], dt.datetime) else None
        fim = serie(p0['G4']) if isinstance(p0['G4'], dt.datetime) else None
        OUT['A3_ind_antes'] = {n: stats_ind(pop(fin, analito, lote, equip, n, ini, fim)) for n in niveis()}
        if anos:
            OUT['A3_ind_estat_antes'] = {n: stats_ind(pop(fin, analito, lote, equip, n, None, None, anos)) for n in niveis()}
        vals = {n: valor_sintetico(out_m['medias'][str(n)], n) for n in niveis()}
        OUT['A3_valores'] = {f'N{n}': vals[n] for n in niveis()}
        al0 = audit_ult(q)
        hh = hora_alvo.hour * 60 + hora_alvo.minute
        # o usuario: DATA, HORA, LOTE, NIVEL, ANALITO, RESULTADO (com virgula), MOTIVO -- EQUIPAMENTO vazio
        rs = []
        for n in niveis():
            hn = hh + (n - 1)
            rs.append(digitar(q, [('DATA', d_alvo.strftime('%d/%m/%Y'), 'l'), ('HORA', f'{hn // 60:02d}:{hn % 60:02d}', 'l'),
                                  ('LOTE', str(lote), 'l'), ('NIVEL', str(n), 'l'), ('ANALITO', analito, 'l'),
                                  ('RESULTADO', txt_valor(vals[n]), 'l'), ('MOTIVO', MARCA, 'l')]))
        OUT['A3_linhas_antes_atualizar'] = [linha_manual(q, r) for r in rs]
        erro3, res3, t3 = atualizar(q, 'A3 manual ' + '+'.join(f'N{n}' for n in niveis()))
        fin3 = final_por_id(q)
        lm = [linha_manual(q, r) for r in rs]
        ids_m = [x['ID_REGISTRO'] for x in lm]
        p1 = painel(q)
        e1 = estat_linhas(q, analito)
        fin3l = list(fin3.values())
        ini1 = serie(p1['G3']) if isinstance(p1['G3'], dt.datetime) else None
        fim1 = serie(p1['G4']) if isinstance(p1['G4'], dt.datetime) else None
        ind1 = {n: stats_ind(pop(fin3l, analito, lote, equip, n, ini1, fim1)) for n in niveis()}
        f1 = fin3.get(ids_m[0])
        OUT['A3_manual'] = {
            'erro': erro3, 'resumo': res3, 't_s': t3, 'linhas_digitadas': lm,
            'final': [resumo_f(fin3.get(i)) for i in ids_m], 'painel_depois': p1, 'estat_depois': e1,
            'ind_depois': ind1,
            'ind_estat_depois': ({n: stats_ind(pop(fin3l, analito, lote, equip, n, None, None, anos)) for n in niveis()}
                                 if anos else None),
            'eng_run': eng_run(q, f1['RUN']) if f1 and f1['RUN'] else None,
            'eng_cab': eng(q)['cab'],
            'dica_ponto': [q.run('mIntegracao.DadosDoId', i) for i in ids_m],
            'qa': [r for r in qa_tab(q) if r.get('ID_REGISTRO') in ids_m],
            'audit': audit_desde(q, al0)[1][-10:], 'sync_log': sync_ultimas(q, 1)}
        # corrida do lote e westgard do ponto (Eng_Saida inteiro guardado para comparar com B)
        OUT['A3_eng_completo'] = eng(q)['linhas'][-15:]

        # ---------------- VALIDACOES
        validacoes(q, analito, lote, equip, d_alvo, hh, fin3, out_m)
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


def alvo_manual(q, fin, analito):
    """Lote/equipamento com mais resultados do analito no mes mais recente; dia passado com uma hora livre
    (nenhum resultado do analito/lote/equipamento a menos de 90 min) depois da ultima corrida do dia."""
    hoje = dt.datetime.now()
    rec = [r for r in fin if r['ANALITO'] == analito and r['PARTICIPA_ESTATISTICA'] == 'SIM'
           and isinstance(r['DATA_HORA'], dt.datetime)]
    ult = max(ingenuo(r['DATA_HORA']) for r in rec)
    mes = [r for r in rec if ingenuo(r['DATA_HORA']).year == ult.year and ingenuo(r['DATA_HORA']).month == ult.month]
    try:
        lotes = q.wb.Names('lstLotes').RefersToRange.Value
        cad = {str(int(x[0])) if isinstance(x[0], float) else str(x[0]) for x in lotes if x[0] not in (None, '')}
    except Exception:
        cad = {str(r['LOTE']) for r in rec}       # sem a lista: qualquer lote com resultado
    cont = collections.Counter((str(r['LOTE']), r['EQUIPAMENTO']) for r in mes if str(r['LOTE']) in cad and int(r['NIVEL']) == 1)
    (lote, equip), _ = cont.most_common(1)[0]
    grupo = [r for r in rec if str(r['LOTE']) == lote]
    grp_eq = [r for r in grupo if r['EQUIPAMENTO'] == equip]
    dias = sorted({ingenuo(r['DATA_HORA']).date() for r in grp_eq}, reverse=True)
    escolha = None
    for d in dias:
        if d >= hoje.date():
            continue
        dh_dia = sorted(ingenuo(r['DATA_HORA']) for r in grupo if ingenuo(r['DATA_HORA']).date() == d)
        cand = dh_dia[-1] + dt.timedelta(minutes=90)
        cand = cand.replace(second=0, microsecond=0)
        if cand.date() == d and cand.hour <= 22:
            escolha = (d, cand.time())
            break
    if escolha is None:
        raise RuntimeError('sem hora livre para o manual')
    por_n = {n: [float(r['RESULTADO']) for r in grp_eq if int(r['NIVEL']) == n] for n in niveis()}
    n1, n2 = por_n[1], por_n[2]
    # resultado real mais recente do grupo (para o teste de conflito com o interfaceamento)
    ref = max((r for r in grp_eq if int(r['NIVEL']) == 1 and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'),
              key=lambda r: ingenuo(r['DATA_HORA']))
    return {'analito': analito, 'lote': lote, 'equip': equip, 'dia': escolha[0], 'hora': escolha[1],
            'media_n1': statistics.fmean(n1[-20:]), 'media_n2': statistics.fmean(n2[-20:]),
            'medias': {str(n): statistics.fmean(v[-20:]) for n, v in por_n.items()},
            'contagem_lote_equip_mes': dict(cont.most_common(4)),
            'ref_interface': resumo_f(ref)}


def validacoes(q, analito, lote, equip, d_alvo, hh, fin3, out_m):
    d = d_alvo.strftime('%d/%m/%Y')
    amanha = (dt.date.today() + dt.timedelta(days=1)).strftime('%d/%m/%Y')
    if P['bio']:
        outro_eq = 'DIMENSION 2' if equip == 'DIMENSION 1' else 'DIMENSION 1'
        eq_falso, nv_falso, vv = 'DIMENSION 9', '3', '90,111'
    else:
        outro_eq = None
        eq_falso, nv_falso = 'XN9999', '4'
        vv = txt_valor(OUT['A3_valores']['N1'])
    ref = out_m['ref_interface']
    rdh = ref['DATA_HORA']
    rconf = rdh + dt.timedelta(minutes=5)
    # dia anterior ao alvo, hora livre presumida (o teste confere pelo status)
    d_m1 = (d_alvo - dt.timedelta(days=1)).strftime('%d/%m/%Y')
    h = lambda m: f'{(m // 60) % 24:02d}:{m % 60:02d}'
    casos = [
        ('V01 só analito e resultado (campos vazios)', [('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V02 resultado em texto "abc"', [('DATA', d, 'l'), ('HORA', h(hh + 3), 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                          ('ANALITO', analito, 'l'), ('RESULTADO', 'abc', 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V03 resultado com ponto "90.5" (Excel pt-BR)', [('DATA', d, 'l'), ('HORA', h(hh + 6), 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                                         ('ANALITO', analito, 'l'), ('RESULTADO', '90.5', 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V04 resultado com ponto "1.5" (Excel pt-BR)', [('DATA', d, 'l'), ('HORA', h(hh + 9), 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                                        ('ANALITO', analito, 'l'), ('RESULTADO', '1.5', 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V05 data futura (amanhã)', [('DATA', amanha, 'l'), ('HORA', '08:00', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                       ('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V06 analito inexistente "Glicoze"', [('DATA', d_m1, 'l'), ('HORA', '23:30', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                                ('ANALITO', 'Glicoze', 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        (f'V07 nível inexistente {nv_falso}', [('DATA', d_m1, 'l'), ('HORA', '23:31', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', nv_falso, 'l'),
                                      ('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V08 lote inexistente "99999"', [('DATA', d_m1, 'l'), ('HORA', '23:32', 'l'), ('LOTE', '99999', 'l'), ('NIVEL', '1', 'l'),
                                           ('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V09 duplicidade manual x manual (cópia exata do MAN N1 do passo A3)',
         [('DATA', d, 'l'), ('HORA', h(hh), 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'), ('ANALITO', analito, 'l'),
          ('RESULTADO', f"{OUT['A3_valores']['N1']:.3f}".replace('.', ','), 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V10 duplicidade manual x interfaceamento (+5 min, mesmo valor)',
         [('DATA', rconf.strftime('%d/%m/%Y'), 'l'), ('HORA', rconf.strftime('%H:%M'), 'l'), ('EQUIPAMENTO', ref['EQUIPAMENTO'], 'l'),
          ('LOTE', str(ref['LOTE']), 'l'), ('NIVEL', '1', 'l'), ('ANALITO', analito, 'l'),
          ('RESULTADO', f"{float(ref['RESULTADO'])}".replace('.', ','), 'l'), ('MOTIVO', MARCA, 'l')]),
        (f'V11 equipamento inexistente "{eq_falso}"', [('DATA', d_m1, 'l'), ('HORA', '23:33', 'l'), ('EQUIPAMENTO', eq_falso, 'l'),
                                                        ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'), ('ANALITO', analito, 'l'),
                                                        ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V12 hora inválida "25:99"', [('DATA', d_m1, 'l'), ('HORA', '25:99', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                        ('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V13 data inválida "31/02/2026"', [('DATA', '31/02/2026', 'l'), ('HORA', '08:00', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                                             ('ANALITO', analito, 'l'), ('RESULTADO', vv, 'l'), ('MOTIVO', MARCA, 'l')]),
        ('V14 lote digitado como número 8977 vs texto (lote com zero à esquerda não testado)', None),
    ]
    casos = [c for c in casos if c[1] is not None]
    linhas = {}
    al0 = audit_ult(q)
    for nome, campos in casos:
        try:
            linhas[nome] = digitar(q, campos)
        except Exception as e:
            linhas[nome] = None
            OUT['erros'].append(f'{nome}: escrita falhou: {e}')
    if P['bio']:
        # V15: equipamento vazio com o Painel em OUTRO equipamento
        set_painel(q, equip=outro_eq)
        try:
            linhas[f'V15 equipamento vazio com o Painel em {outro_eq}'] = digitar(
                q, [('DATA', d_m1, 'l'), ('HORA', '23:34', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '1', 'l'),
                    ('ANALITO', analito, 'l'), ('RESULTADO', '90,111', 'l'), ('MOTIVO', MARCA, 'l')])
        except Exception as e:
            OUT['erros'].append(f'V15: {e}')
        set_painel(q, equip=equip)
    else:
        # V15 (Hema): equipamento vazio recebe o EQUIPAMENTO_PADRAO (sem seletor no Painel)
        # V16 (Hema): nivel 3 com resultado com virgula e 2 casas, como o equipamento manda (caso comum)
        try:
            linhas['V15 equipamento vazio (Hema: deve virar o EQUIPAMENTO_PADRAO)'] = digitar(
                q, [('DATA', d_m1, 'l'), ('HORA', '23:34', 'l'), ('LOTE', str(lote), 'l'), ('NIVEL', '2', 'l'),
                    ('ANALITO', analito, 'l'), ('RESULTADO', txt_valor(OUT['A3_valores']['N2']), 'l'), ('MOTIVO', MARCA, 'l')])
        except Exception as e:
            OUT['erros'].append(f'V15: {e}')
        try:
            linhas['V16 equipamento digitado em minúsculas "xn1000"'] = digitar(
                q, [('DATA', d_m1, 'l'), ('HORA', '23:36', 'l'), ('EQUIPAMENTO', equip.lower(), 'l'), ('LOTE', str(lote), 'l'),
                    ('NIVEL', '3', 'l'), ('ANALITO', analito, 'l'), ('RESULTADO', txt_valor(OUT['A3_valores']['N3']), 'l'),
                    ('MOTIVO', MARCA, 'l')])
        except Exception as e:
            OUT['erros'].append(f'V16: {e}')
    antes = {n: (linha_manual(q, r) if r else None) for n, r in linhas.items()}
    erro, res, t = atualizar(q, 'A4 validações')
    fin = final_por_id(q)
    qa = qa_tab(q)
    resultado = {}
    for n, r in linhas.items():
        if not r:
            continue
        lm = linha_manual(q, r)
        i = lm['ID_REGISTRO']
        cand = [k for k in fin if k == i or str(k).startswith(f'MAN_INVALIDO_L{r}') or str(k).startswith(f'{i}~')]
        resultado[n] = {'digitado_antes_de_atualizar': antes[n], 'linha_depois': lm,
                        'final': [resumo_f(fin[k]) for k in cand],
                        'qa': [x for x in qa if x.get('ID_REGISTRO') in cand]}
    p = painel(q)
    OUT['A4_validacoes'] = {'erro': erro, 'resumo': res, 't_s': t, 'casos': resultado, 'painel': p,
                            'audit': audit_desde(q, al0)[1][-20:],
                            'qa_codigos': dict(collections.Counter((x.get('SEVERIDADE'), x.get('CODIGO')) for x in qa)),
                            'qa_agregados': [x for x in qa if x.get('CODIGO') in ('A10', 'I01', 'A06')][:20],
                            'manual_a3_depois': [resumo_f(fin.get(x['ID_REGISTRO'])) for x in OUT['A3_manual']['final'] if x]}


# ================================================================== sessao B
def sessao_b(copia, pasta, saida, saida_a):
    A = json.load(open(saida_a, encoding='utf-8'))
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        erro1, res1, t1 = atualizar(q, 'B1 DB_SEAC congelado')
        alvo = A['A3_alvo']
        lote, equip, analito = alvo['lote'], alvo['equip'], alvo['analito']
        dia = dt.date.fromisoformat(alvo['dia'])
        hora = dt.time.fromisoformat(alvo['hora'])
        rec = q.ler('tblDB_Recebimento')
        cab = [c for c in rec[0].keys() if c not in ('ID_REGISTRO', 'RECEBIDO_EM')]
        # o molde e o resultado real mais recente do mesmo analito (nome do QC_INI, via tblCQ_Final), lote,
        # equipamento e nivel; a linha clonada e a do RECEBIMENTO (codigo do analito na origem, ex. GLUC)
        fin0 = final_por_id(q)
        rec_id = {r['ID_REGISTRO']: r for r in rec}
        mold = {}
        for n in niveis():
            m = [r for r in fin0.values() if r['ANALITO'] == analito and str(r['LOTE']) == lote and r['EQUIPAMENTO'] == equip
                 and int(r['NIVEL'] or 0) == n and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
                 and isinstance(r['DATA_HORA'], dt.datetime) and r['ID_REGISTRO'] in rec_id]
            mold[n] = rec_id[max(m, key=lambda r: ingenuo(r['DATA_HORA']))['ID_REGISTRO']]
        v = A['A3_valores']
        dh1 = dt.datetime.combine(dia, hora)
        linhas = [{'_modelo': mold[n], 'ID_ORIGEM': 9000000 + n, 'ITEM_ID': 9000000 + n, 'DATA': dia,
                   'DATA_HORA': dh1 + dt.timedelta(minutes=n - 1), 'VALOR': v[f'N{n}'], 'ID_AMOSTRA': 'TESTE-AUDITORIA'}
                  for n in niveis()]
        sint = os.path.join(pasta, 'DB_SEAC_sintetico_F1.xlsx')
        esc = gss.gerar(sint, P['nome'], cab, mold[1], linhas, tabela=q.cfg('TABELA_ORIGEM'))
        q.cfg('CAMINHO_DB_SEAC', sint)
        erro2, res2, t2 = atualizar(q, f"B2 sintético com os {P['nlv']} resultados do A3")
        fin = final_por_id(q)
        set_painel(q, lote=lote, equip=equip, analito=analito, k3=dia.year,
                   k4=None if P['bio'] else 'Tudo (sem filtro)', ano_estat=dia.year)
        p1 = painel(q)
        e1 = estat_linhas(q, analito)
        ini1 = serie(p1['G3']) if isinstance(p1['G3'], dt.datetime) else None
        fim1 = serie(p1['G4']) if isinstance(p1['G4'], dt.datetime) else None
        finl = list(fin.values())
        ids = [f"{P['pref']}-{9000000 + n}" for n in niveis()]
        anos = None if P['bio'] else (dia.year, dia.year)
        f1 = fin.get(ids[0])
        OUT['B'] = {'erro1': erro1, 't1': t1, 'erro2': erro2, 'resumo2': res2, 't2': t2,
                    'escritas': [{k: x.get(k) for k in ('ID_ORIGEM', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'DATA', 'DATA_HORA', 'VALOR', 'BLOCO', 'UNIDADE')} for x in esc],
                    'final': [resumo_f(fin.get(i)) for i in ids], 'painel': p1, 'estat': e1,
                    'ind': {n: stats_ind(pop(finl, analito, lote, equip, n, ini1, fim1)) for n in niveis()},
                    'ind_estat': ({n: stats_ind(pop(finl, analito, lote, equip, n, None, None, anos)) for n in niveis()}
                                  if anos else None),
                    'eng_run': eng_run(q, f1['RUN']) if f1 and f1['RUN'] else None,
                    'eng_completo': eng(q)['linhas'][-15:],
                    'dica_ponto': [q.run('mIntegracao.DadosDoId', i) for i in ids],
                    'sync_log': sync_ultimas(q, 1),
                    'qa': [r for r in qa_tab(q) if r.get('ID_REGISTRO') in ids]}
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


# ================================================================== sessao C (acrescentada no 2o passe do ciclo)
def sessao_c(copia, pasta, saida, ids_alvo):
    """SUBSTITUI_ID (ADR-082, GUIA §3) pelo caminho do usuario, num caso real: o usuario passa o mouse no ponto,
    anota o numero do ID do interfaceamento e o digita na coluna SUBSTITUI_ID de uma linha livre; confere o
    preenchimento automatico; digita RESULTADO e MOTIVO; ATUALIZAR. Tambem: ID inexistente (nada preenchido) e
    SUBSTITUI_ID sem MOTIVO (MANUAL_INCOMPLETO). ids_alvo = 'BIO-x,BIO-y' (x substituido com motivo; y sem motivo).
    Confere manual ATIVO / automatico SUBSTITUIDO_MANUAL, QA I04/E04, Audit_Log, dica do ponto e a estatistica do
    Painel e da Estatistica contra o recalculo independente."""
    x_id, y_id = ids_alvo.split(',')
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        erro1, res1, t1 = atualizar(q, 'C1 DB_SEAC congelado (linha de base)')
        fin0 = final_por_id(q)
        x0, y0 = fin0.get(x_id), fin0.get(y_id)
        OUT['C_alvos_antes'] = [resumo_f(x0), resumo_f(y0)]
        analito, lote, equip = x0['ANALITO'], str(x0['LOTE']), x0['EQUIPAMENTO']
        dhx = ingenuo(x0['DATA_HORA'])
        set_painel(q, lote=lote, equip=equip, analito=analito, k3=dhx.year,
                   k4=None if P['bio'] else 'Tudo (sem filtro)', ano_estat=dhx.year)
        anos = None if P['bio'] else (dhx.year, dhx.year)
        p0 = painel(q)
        ini = serie(p0['G3']) if isinstance(p0['G3'], dt.datetime) else None
        fim = serie(p0['G4']) if isinstance(p0['G4'], dt.datetime) else None
        OUT['C_painel_antes'] = p0
        OUT['C_estat_antes'] = estat_linhas(q, analito)
        OUT['C_ind_antes'] = {n: stats_ind(pop(list(fin0.values()), analito, lote, equip, n, ini, fim)) for n in niveis()}
        if anos:
            OUT['C_ind_estat_antes'] = {n: stats_ind(pop(list(fin0.values()), analito, lote, equip, n, None, None, anos))
                                        for n in niveis()}
        al0 = audit_ult(q)
        num = lambda i: i.split('-', 1)[1]
        # C2: o usuario digita so o numero do ID (como aparece na dica do ponto)
        r_x = digitar(q, [('SUBSTITUI_ID', num(x_id), 'l')])
        auto_x = linha_manual(q, r_x)
        if P['bio']:
            valor_novo, valor_y = 270.111, 270.222
        else:
            # Hema: perto da media do lote no nivel do resultado substituido (1 casa + 0,011 / 0,022)
            mx_ = OUT['C_ind_antes'][int(x0['NIVEL'])]['media']
            my_ = OUT['C_ind_antes'][int(y0['NIVEL'])]['media']
            valor_novo, valor_y = valor_sintetico(mx_, 1), valor_sintetico(my_, 2)
        lo = q.lo('tblResultados_Manuais')
        lo.ListColumns('RESULTADO').DataBodyRange.Cells(r_x, 1).FormulaLocal = txt_valor(valor_novo)
        lo.ListColumns('MOTIVO').DataBodyRange.Cells(r_x, 1).FormulaLocal = MARCA
        # C3: ID inexistente
        r_n = digitar(q, [('SUBSTITUI_ID', '987654321', 'l')])
        auto_n = linha_manual(q, r_n)
        lo.ListColumns('SUBSTITUI_ID').DataBodyRange.Cells(r_n, 1).ClearContents()
        # C4: SUBSTITUI_ID sem MOTIVO
        r_y = digitar(q, [('SUBSTITUI_ID', num(y_id), 'l')])
        lo.ListColumns('RESULTADO').DataBodyRange.Cells(r_y, 1).FormulaLocal = txt_valor(valor_y)
        auto_y = linha_manual(q, r_y)
        OUT['C_preenchimento'] = {'com_motivo': auto_x, 'inexistente': auto_n, 'sem_motivo': auto_y}
        erro2, res2, t2 = atualizar(q, 'C5 ATUALIZAR com as substituicoes')
        fin1 = final_por_id(q)
        lx, ly = linha_manual(q, r_x), linha_manual(q, r_y)
        mx, my = lx['ID_REGISTRO'], ly['ID_REGISTRO']
        qa = qa_tab(q)
        p1 = painel(q)
        ini1 = serie(p1['G3']) if isinstance(p1['G3'], dt.datetime) else None
        fim1 = serie(p1['G4']) if isinstance(p1['G4'], dt.datetime) else None
        fin1l = list(fin1.values())
        runs0 = {k: (r['RUN'], r['CORRIDA_NO_DIA']) for k, r in fin0.items()}
        mud = [k for k, r in fin1.items() if k in runs0 and k not in (x_id, y_id)
               and (r['RUN'], r['CORRIDA_NO_DIA']) != runs0[k]]
        OUT['C'] = {'erro1': erro1, 't1': t1, 'erro2': erro2, 'resumo2': res2, 't2': t2, 'valor_novo': valor_novo,
                    'linha_com_motivo': lx, 'linha_sem_motivo': ly,
                    'final': {k: resumo_f(fin1.get(k)) for k in (mx, my, x_id, y_id)},
                    'outros_RUN_mudaram': mud[:10], 'n_outros_RUN_mudaram': len(mud),
                    'qa': [r for r in qa if r.get('ID_REGISTRO') in (mx, my, x_id, y_id)],
                    'qa_codigos': dict(collections.Counter((r.get('SEVERIDADE'), r.get('CODIGO')) for r in qa)),
                    'audit': audit_desde(q, al0)[1][-12:],
                    'painel_depois': p1, 'estat_depois': estat_linhas(q, analito),
                    'ind_depois': {n: stats_ind(pop(fin1l, analito, lote, equip, n, ini1, fim1)) for n in niveis()},
                    'ind_estat_depois': ({n: stats_ind(pop(fin1l, analito, lote, equip, n, None, None, anos)) for n in niveis()}
                                         if anos else None), 'valor_y': valor_y,
                    'dica_ponto': [q.run('mIntegracao.DadosDoId', i) for i in (mx, x_id)],
                    'eng_run': eng_run(q, fin1[mx]['RUN']) if fin1.get(mx) and fin1[mx]['RUN'] else None,
                    'sync_log': sync_ultimas(q, 1)}
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


# ================================================================== sessao D (Hema, ciclo 2026-10-31_ciclo01)
def sessao_d(copia, pasta, saida, saida_a):
    """Isola o achado do V16 da sessao A: EQUIPAMENTO digitado/colado em minusculas ("xn1000"). O RUN e a corrida
    (mEtlCQ.CalcularCorridas) agrupam pelo texto exato do equipamento, mas o motor do Painel/Estatistica da Hema
    (mEstatistica.GarantirIndice) compara em maiusculas. D2: so um N3 "xn1000" numa hora livre do dia de um RUN real;
    confere RUN, Eng_Saida do RUN, engPainel, Estatistica e o recalculo independente (com e sem diferenciar maiusculas).
    D3: o mesmo com o ANALITO em minusculas ("wbc") e o equipamento certo, noutro dia."""
    A = json.load(open(saida_a, encoding='utf-8'))
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        erro1, res1, t1 = atualizar(q, 'D1 DB_SEAC congelado (linha de base)')
        alvo = A['A3_alvo']
        lote, equip, analito = alvo['lote'], alvo['equip'], alvo['analito']
        dia = dt.date.fromisoformat(alvo['dia'])
        set_painel(q, lote=lote, analito=analito, k3=dia.year, k4='Tudo (sem filtro)', ano_estat=dia.year)
        anos = (dia.year, dia.year)

        def foto(rot, fin, run):
            finl = list(fin.values())
            pm = [r for r in finl if r['ANALITO'] == analito and str(r['LOTE']) == lote and r['PARTICIPA_ESTATISTICA'] == 'SIM'
                  and str(r['EQUIPAMENTO']).upper() == equip.upper()]
            OUT[rot] = {
                'painel': painel(q), 'estat': estat_linhas(q, analito),
                'ind_exato': {n: stats_ind(pop(finl, analito, lote, equip, n, None, None)) for n in niveis()},
                'ind_maiusc': {n: stats_ind([float(r['RESULTADO']) for r in pm if int(r['NIVEL'] or 0) == n
                                             and isinstance(r['RESULTADO'], (int, float))]) for n in niveis()},
                'ind_maiusc_ano': {n: stats_ind([float(r['RESULTADO']) for r in pm if int(r['NIVEL'] or 0) == n
                                                 and isinstance(r['RESULTADO'], (int, float))
                                                 and hasattr(r['DATA_HORA'], 'year') and r['DATA_HORA'].year in range(anos[0], anos[1] + 1)])
                                   for n in niveis()},
                'linhas_do_run': [resumo_f(r) for r in pm if r['RUN'] == run] if run else None,
                'eng_run': eng_run(q, run) if run else None}

        fin0 = final_por_id(q)
        # dia anterior ao alvo: o 1o RUN real do dia (alvo dia-1 = 29/10 no ciclo)
        d1 = dia - dt.timedelta(days=1)
        runs_d1 = sorted({int(r['RUN']) for r in fin0.values() if r['ANALITO'] == analito and str(r['LOTE']) == lote
                          and r['EQUIPAMENTO'] == equip and isinstance(r['DATA_HORA'], dt.datetime)
                          and ingenuo(r['DATA_HORA']).date() == d1 and r['RUN']})
        run1 = runs_d1[0] if runs_d1 else None
        OUT['D_runs_dia'] = runs_d1
        foto('D1', fin0, run1)
        v3 = A['A3_valores']['N3']
        r = digitar(q, [('DATA', d1.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('EQUIPAMENTO', equip.lower(), 'l'),
                        ('LOTE', str(lote), 'l'), ('NIVEL', '3', 'l'), ('ANALITO', analito, 'l'),
                        ('RESULTADO', txt_valor(v3), 'l'), ('MOTIVO', MARCA, 'l')])
        erro2, res2, t2 = atualizar(q, 'D2 manual N3 com equipamento "xn1000"')
        fin2 = final_por_id(q)
        lm = linha_manual(q, r)
        set_painel(q, analito=analito)
        foto('D2', fin2, run1)
        OUT['D2']['linha'] = lm
        OUT['D2']['final'] = resumo_f(fin2.get(lm['ID_REGISTRO']))
        OUT['D2']['qa'] = [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm['ID_REGISTRO'] or x.get('CODIGO') == 'A10']
        OUT['D2']['dica'] = q.run('mIntegracao.DadosDoId', lm['ID_REGISTRO'])
        OUT['D2']['erro'], OUT['D2']['t'] = erro2, t2
        # D3: analito em minusculas, equipamento certo, dois dias antes (hora livre presumida, conferida pelo status)
        d2 = dia - dt.timedelta(days=2)
        runs_d2 = sorted({int(x['RUN']) for x in fin2.values() if x['ANALITO'] == analito and str(x['LOTE']) == lote
                          and x['EQUIPAMENTO'] == equip and isinstance(x['DATA_HORA'], dt.datetime)
                          and ingenuo(x['DATA_HORA']).date() == d2 and x['RUN']})
        r3 = digitar(q, [('DATA', d2.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('LOTE', str(lote), 'l'),
                         ('NIVEL', '3', 'l'), ('ANALITO', analito.lower(), 'l'), ('RESULTADO', txt_valor(v3), 'l'),
                         ('MOTIVO', MARCA, 'l')])
        erro3, res3, t3 = atualizar(q, 'D3 manual N3 com analito "wbc"')
        fin3 = final_por_id(q)
        lm3 = linha_manual(q, r3)
        set_painel(q, analito=analito)
        f3 = fin3.get(lm3['ID_REGISTRO'])
        foto('D3', fin3, int(f3['RUN']) if f3 and f3.get('RUN') else None)
        OUT['D3']['runs_dia_antes'] = runs_d2
        OUT['D3']['linha'] = lm3
        OUT['D3']['final'] = resumo_f(f3)
        OUT['D3']['qa'] = [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm3['ID_REGISTRO']]
        OUT['D3']['erro'], OUT['D3']['t'] = erro3, t3
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


# ================================================================== sessao E (Bio, 3o passe do ciclo 2026-10-31_ciclo01)
def sessao_e(copia, pasta, saida, saida_a):
    """Versao Bioquimica da sessao D (pendencia deixada pela OC-HF1-1 da Hema): EQUIPAMENTO "dimension 1" e ANALITO
    "glicose" em minusculas num manual N2 valido, numa hora livre (23:36) de dois dias anteriores ao alvo da sessao A
    que tenham RUN real do analito/lote/equipamento. Confere o RUN que o ETL deu, o Eng_Saida do RUN real, o engPainel,
    a Estatistica e o recalculo independente com o equipamento/analito exatos (o que o ETL usa) e em maiusculas (o que
    o motor da Bio usa: mEstatistica.GarantirIndice/EquipFiltro), sempre no periodo do Painel (G3:G4)."""
    A = json.load(open(saida_a, encoding='utf-8'))
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        erro1, res1, t1 = atualizar(q, 'E1 DB_SEAC congelado (linha de base)')
        alvo = A['A3_alvo']
        lote, equip, analito = alvo['lote'], alvo['equip'], alvo['analito']
        dia = dt.date.fromisoformat(alvo['dia'])
        set_painel(q, lote=lote, equip=equip, analito=analito, k3=dia.year)
        nv = 2
        val = A['A3_valores'][f'N{nv}']

        def popu(finl, ini, fim, exato):
            out = []
            for r in finl:
                if exato:
                    ok = r['ANALITO'] == analito and r['EQUIPAMENTO'] == equip
                else:
                    ok = (str(r['ANALITO']).strip().upper() == analito.upper()
                          and str(r['EQUIPAMENTO']).strip().upper() == equip.upper())
                if not (ok and str(r['LOTE']) == str(lote) and int(r['NIVEL'] or 0) == nv
                        and r['PARTICIPA_ESTATISTICA'] == 'SIM' and isinstance(r['RESULTADO'], (int, float))):
                    continue
                d = r['DATA']
                if hasattr(d, 'year'):
                    s = serie(ingenuo(d))
                    if (ini is None or s >= ini) and (fim is None or s < int(fim) + 1):
                        out.append((r['ID_REGISTRO'], r['RUN'], float(r['RESULTADO'])))
            return out

        def foto(rot, fin, run):
            finl = list(fin.values())
            p = painel(q)
            ini = serie(p['G3']) if isinstance(p['G3'], dt.datetime) else None
            fim = serie(p['G4']) if isinstance(p['G4'], dt.datetime) else None
            ex_, mi_ = popu(finl, ini, fim, True), popu(finl, ini, fim, False)
            OUT[rot] = {'painel': p, 'estat': estat_linhas(q, analito),
                        'ind_exato': stats_ind([x[2] for x in ex_]), 'ind_maiusc': stats_ind([x[2] for x in mi_]),
                        'pop_maiusc': mi_,
                        'linhas_do_run': [resumo_f(r) for r in finl if run and r['RUN'] and int(float(r['RUN'])) == int(run)
                                          and str(r['ANALITO']).upper() == analito.upper()
                                          and str(r['LOTE']) == str(lote)] if run else None,
                        'eng_run': eng_run(q, run) if run else None}

        fin0 = final_por_id(q)
        dias_run = collections.defaultdict(set)
        for r in fin0.values():
            if (r['ANALITO'] == analito and str(r['LOTE']) == lote and r['EQUIPAMENTO'] == equip and r['RUN']
                    and isinstance(r['DATA_HORA'], dt.datetime) and int(r['NIVEL'] or 0) == nv
                    and r['PARTICIPA_ESTATISTICA'] == 'SIM'):
                d = ingenuo(r['DATA_HORA']).date()
                if d < dia:
                    dias_run[d].add(int(r['RUN']))
        dsel = sorted(dias_run, reverse=True)[:2]
        d1, d2 = dsel[0], dsel[1]
        run1 = sorted(dias_run[d1])[0]
        OUT['E_dias'] = {'d1': d1, 'runs_d1': sorted(dias_run[d1]), 'd2': d2, 'runs_d2': sorted(dias_run[d2]),
                         'nivel': nv, 'valor': val}
        foto('E1', fin0, run1)
        al0 = audit_ult(q)
        # E2: equipamento em minusculas
        r = digitar(q, [('DATA', d1.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('EQUIPAMENTO', equip.lower(), 'l'),
                        ('LOTE', str(lote), 'l'), ('NIVEL', str(nv), 'l'), ('ANALITO', analito, 'l'),
                        ('RESULTADO', txt_valor(val), 'l'), ('MOTIVO', MARCA, 'l')])
        OUT['E2_antes_atualizar'] = linha_manual(q, r)
        erro2, res2, t2 = atualizar(q, f'E2 manual N{nv} com equipamento "{equip.lower()}"')
        fin2 = final_por_id(q)
        lm = linha_manual(q, r)
        set_painel(q, analito=analito)
        foto('E2', fin2, run1)
        f2 = fin2.get(lm['ID_REGISTRO'])
        OUT['E2'].update({'linha': lm, 'final': resumo_f(f2), 'erro': erro2, 't': t2, 'resumo': res2,
                          'qa': [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm['ID_REGISTRO']],
                          'qa_codigos': dict(collections.Counter((x.get('SEVERIDADE'), x.get('CODIGO')) for x in qa_tab(q))),
                          'dica': q.run('mIntegracao.DadosDoId', lm['ID_REGISTRO']),
                          'eng_run_manual': eng_run(q, f2['RUN']) if f2 and f2.get('RUN') else None,
                          'audit': audit_desde(q, al0)[1][-8:]})
        # E3: analito em minusculas, equipamento certo (digitado), outro dia
        al1 = audit_ult(q)
        r3 = digitar(q, [('DATA', d2.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('EQUIPAMENTO', equip, 'l'),
                         ('LOTE', str(lote), 'l'), ('NIVEL', str(nv), 'l'), ('ANALITO', analito.lower(), 'l'),
                         ('RESULTADO', txt_valor(val), 'l'), ('MOTIVO', MARCA, 'l')])
        erro3, res3, t3 = atualizar(q, f'E3 manual N{nv} com analito "{analito.lower()}"')
        fin3 = final_por_id(q)
        lm3 = linha_manual(q, r3)
        set_painel(q, analito=analito)
        f3 = fin3.get(lm3['ID_REGISTRO'])
        run3 = int(f3['RUN']) if f3 and f3.get('RUN') else None
        foto('E3', fin3, run3)
        OUT['E3'].update({'linha': lm3, 'final': resumo_f(f3), 'erro': erro3, 't': t3, 'resumo': res3,
                          'run_real_d2': sorted(dias_run[d2])[0],
                          'eng_run_real_d2': eng_run(q, sorted(dias_run[d2])[0]),
                          'qa': [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm3['ID_REGISTRO']],
                          'dica': q.run('mIntegracao.DadosDoId', lm3['ID_REGISTRO']),
                          'audit': audit_desde(q, al1)[1][-8:]})
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


# ================================================================== sessao F (Hema, 2o passe da F1 da Hema, ciclo 2026-10-31_ciclo01)
def _candidatos_f():
    alvo_eq = 'XN1000' if not P['bio'] else 'DIMENSION 1'
    an = P['analito']
    return [
        ('EQUIPAMENTO', alvo_eq), ('EQUIPAMENTO', alvo_eq.lower()), ('EQUIPAMENTO', alvo_eq.capitalize()),
        ('EQUIPAMENTO', ' ' + alvo_eq), ('EQUIPAMENTO', alvo_eq + ' '), ('EQUIPAMENTO', 'XN9999'),
        ('ANALITO', an), ('ANALITO', an.lower()), ('ANALITO', an + ' '), ('ANALITO', 'Glicoze'),
        ('LOTE', None), ('LOTE', '99999'), ('LOTE', 'LOTE_ESPACO'),
        ('DATA', '30/10/2026'), ('DATA', '31/02/2026'), ('DATA', '01/11/2027'), ('DATA', 'abc'), ('DATA', '2026-10-30'),
        ('HORA', '08:05'), ('HORA', '25:99'), ('HORA', 'abc'),
        ('NIVEL', '1'), ('NIVEL', '4'), ('NIVEL', '0'), ('NIVEL', '1,5'), ('NIVEL', 'abc'),
        ('RESULTADO', '3,111'), ('RESULTADO', '90.5'), ('RESULTADO', 'abc'), ('RESULTADO', '-1'),
        ('MOTIVO', MARCA), ('MOTIVO', 'texto livre qualquer'),
    ]


def sessao_f(copia, pasta, saida, saida_a):
    """Fecha o limite 1 do 1o passe: o que a validacao de dados do Excel faz com o que o usuario DIGITA.
    F1: para cada candidato (coluna, texto), escreve o texto numa linha livre da tblResultados_Manuais pelo
        FormulaLocal (o que o Excel pt-BR faz com a digitacao), com os eventos DESLIGADOS (so a celula), le
        Range.Validation.Value (True = a regra de validacao aceita o valor; False = a caixa de erro bloquearia a
        digitacao, se ShowError) e limpa a celula. Nada passa pelo ATUALIZAR nesta parte.
    F2: dois manuais N3 com ESPACO sobrando (variante da OC-HF1-1): ANALITO "WBC " (alvo-3 dias, 23:36) e
        EQUIPAMENTO "XN1000 " (alvo-4 dias, 23:36); ATUALIZAR; confere RUN, STATUS, QA, Painel, Estatistica,
        Eng_Saida do 1o RUN real do dia e o recalculo independente."""
    A = json.load(open(saida_a, encoding='utf-8'))
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        alvo = A['A3_alvo']
        lote, equip, analito = alvo['lote'], alvo['equip'], alvo['analito']
        dia = dt.date.fromisoformat(alvo['dia'])
        # ---------------- F1 validacao de dados x digitacao
        lo = q.lo('tblResultados_Manuais')
        r = linha_livre(q)
        q.ex.xl.EnableEvents = False
        casos = []
        for col, txt in _candidatos_f():
            if txt is None:
                txt = str(lote)
            elif txt == 'LOTE_ESPACO':
                txt = str(lote) + ' '
            cel = lo.ListColumns(col).DataBodyRange.Cells(r, 1)
            d = {'coluna': col, 'digitado': txt}
            try:
                cel.FormulaLocal = txt
                d['value2'] = cel.Value2
                d['tipo'] = type(cel.Value2).__name__
                d['texto'] = cel.Text
                try:
                    d['validacao_aceita'] = bool(cel.Validation.Value)
                    d['show_error'] = bool(cel.Validation.ShowError)
                    d['error_message'] = cel.Validation.ErrorMessage
                except Exception as e:
                    d['validacao_aceita'] = f'sem validacao ({type(e).__name__})'
            except Exception as e:
                d['erro_escrita'] = f'{type(e).__name__}: {e}'
            try:
                cel.ClearContents()
            except Exception as e:
                d['erro_limpar'] = str(e)
            casos.append(d)
        q.ex.xl.EnableEvents = True
        OUT['F1_validacao'] = {'linha': r, 'casos': casos,
                               'linha_depois_vazia': all(x in (None, '') for x in lo.DataBodyRange.Rows(r).Value[0][:12])}
        log('F1 validacao', len(casos), 'casos')
        # ---------------- F2 espacos sobrando
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        erro1, res1, t1 = atualizar(q, 'F0 DB_SEAC congelado (linha de base)')
        set_painel(q, lote=lote, analito=analito, k3=dia.year, k4='Tudo (sem filtro)', ano_estat=dia.year)
        fin0 = final_por_id(q)

        def runs_dia(fin, d):
            return sorted({int(x['RUN']) for x in fin.values() if x['ANALITO'] == analito and str(x['LOTE']) == lote
                           and x['EQUIPAMENTO'] == equip and isinstance(x['DATA_HORA'], dt.datetime)
                           and ingenuo(x['DATA_HORA']).date() == d and x['RUN']})

        def foto(fin):
            finl = list(fin.values())
            pm = [x for x in finl if str(x['ANALITO']).strip().upper() == analito.upper() and str(x['LOTE']).strip() == lote
                  and x['PARTICIPA_ESTATISTICA'] == 'SIM' and str(x['EQUIPAMENTO']).strip().upper() == equip.upper()]
            return {'painel': painel(q), 'estat': estat_linhas(q, analito),
                    'ind_exato': {n: stats_ind(pop(finl, analito, lote, equip, n, None, None)) for n in niveis()},
                    'ind_normalizado': {n: stats_ind([float(x['RESULTADO']) for x in pm if int(x['NIVEL'] or 0) == n
                                                      and isinstance(x['RESULTADO'], (int, float))]) for n in niveis()}}
        d3, d4 = dia - dt.timedelta(days=3), dia - dt.timedelta(days=4)
        OUT['F2_antes'] = foto(fin0)
        OUT['F2_runs'] = {'d3': [str(d3), runs_dia(fin0, d3)], 'd4': [str(d4), runs_dia(fin0, d4)]}
        run_d3 = runs_dia(fin0, d3)[0] if runs_dia(fin0, d3) else None
        run_d4 = runs_dia(fin0, d4)[0] if runs_dia(fin0, d4) else None
        OUT['F2_eng_antes'] = {'d3': eng_run(q, run_d3) if run_d3 else None, 'd4': eng_run(q, run_d4) if run_d4 else None}
        v3 = A['A3_valores']['N3']
        al0 = audit_ult(q)
        ra = digitar(q, [('DATA', d3.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('LOTE', str(lote), 'l'),
                         ('NIVEL', '3', 'l'), ('ANALITO', analito + ' ', 'l'), ('RESULTADO', txt_valor(v3), 'l'),
                         ('MOTIVO', MARCA, 'l')])
        rb = digitar(q, [('DATA', d4.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('EQUIPAMENTO', equip + ' ', 'l'),
                         ('LOTE', str(lote), 'l'), ('NIVEL', '3', 'l'), ('ANALITO', analito, 'l'),
                         ('RESULTADO', txt_valor(v3), 'l'), ('MOTIVO', MARCA, 'l')])
        OUT['F2_digitado'] = {'a': linha_manual(q, ra), 'b': linha_manual(q, rb)}
        erro2, res2, t2 = atualizar(q, 'F2 manuais N3 com espaco sobrando')
        fin2 = final_por_id(q)
        set_painel(q, analito=analito)
        OUT['F2_depois'] = foto(fin2)
        OUT['F2_eng_depois'] = {'d3': eng_run(q, run_d3) if run_d3 else None, 'd4': eng_run(q, run_d4) if run_d4 else None}
        out_l = {}
        for k, rr in (('a', ra), ('b', rb)):
            lm = linha_manual(q, rr)
            f = fin2.get(lm['ID_REGISTRO'])
            out_l[k] = {'linha': lm, 'final': resumo_f(f),
                        'eng_run_proprio': eng_run(q, int(f['RUN'])) if f and f.get('RUN') else None,
                        'qa': [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm['ID_REGISTRO']],
                        'dica': q.run('mIntegracao.DadosDoId', lm['ID_REGISTRO'])}
        OUT['F2_linhas'] = out_l
        OUT['F2_qa_codigos'] = collections.Counter(x.get('CODIGO') for x in qa_tab(q))
        OUT['F2_audit'] = audit_desde(q, al0)[1][-10:]
        OUT['F2_erro'], OUT['F2_t'], OUT['F2_resumo'] = erro2, t2, res2
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


def sessao_g(copia, pasta, saida, saida_a):
    """Resultado implausivel DIGITADO (a validacao do Excel aceita qualquer decimal >= -1E+307, ver sessao F):
    N1 com RESULTADO "-1" (alvo-5 dias, 23:36) e N2 com RESULTADO "0" (alvo-6 dias, 23:36), equipamento, analito e
    lote certos; ATUALIZAR; confere STATUS, RUN, QA, Painel, Estatistica, Eng_Saida e o recalculo independente."""
    A = json.load(open(saida_a, encoding='utf-8'))
    h0 = sha(copia)
    q = QA(P['nome'], copia)
    OUT['excel_pid'] = q.ex.pid
    OUT['produto'] = P['nome']
    try:
        alvo = A['A3_alvo']
        lote, equip, analito = alvo['lote'], alvo['equip'], alvo['analito']
        dia = dt.date.fromisoformat(alvo['dia'])
        congelado = os.path.join(pasta, 'DB_SEAC_congelado.xlsm')
        OUT['dbseac_congelado_sha12'] = sha(congelado)
        q.cfg('CAMINHO_DB_SEAC', congelado)
        atualizar(q, 'G0 DB_SEAC congelado (linha de base)')
        set_painel(q, lote=lote, analito=analito, k3=dia.year, k4='Tudo (sem filtro)', ano_estat=dia.year)

        def foto():
            finl = list(final_por_id(q).values())
            return {'painel': painel(q), 'estat': estat_linhas(q, analito),
                    'ind': {n: stats_ind(pop(finl, analito, lote, equip, n, None, None)) for n in niveis()}}
        OUT['G_antes'] = foto()
        casos = (('N1 "-1"', dia - dt.timedelta(days=5), '1', '-1'), ('N2 "0"', dia - dt.timedelta(days=6), '2', '0'))
        rs = []
        for rot, d, nv, val in casos:
            rs.append((rot, digitar(q, [('DATA', d.strftime('%d/%m/%Y'), 'l'), ('HORA', '23:36', 'l'), ('LOTE', str(lote), 'l'),
                                       ('NIVEL', nv, 'l'), ('ANALITO', analito, 'l'), ('RESULTADO', val, 'l'),
                                       ('MOTIVO', MARCA, 'l')])))
        erro, res, t = atualizar(q, 'G1 manuais com resultado -1 e 0')
        set_painel(q, analito=analito)
        OUT['G_depois'] = foto()
        fin = final_por_id(q)
        OUT['G_linhas'] = {}
        for rot, r in rs:
            lm = linha_manual(q, r)
            f = fin.get(lm['ID_REGISTRO'])
            OUT['G_linhas'][rot] = {'linha': lm, 'final': resumo_f(f),
                                   'eng_run': eng_run(q, int(f['RUN'])) if f and f.get('RUN') else None,
                                   'qa': [x for x in qa_tab(q) if x.get('ID_REGISTRO') == lm['ID_REGISTRO']]}
        OUT['G_qa_codigos'] = collections.Counter(x.get('CODIGO') for x in qa_tab(q))
        OUT['G_erro'], OUT['G_t'], OUT['G_resumo'] = erro, t, res
    except Exception as ex:
        OUT['erros'].append(f'{type(ex).__name__}: {ex}\n{traceback.format_exc()[-2500:]}')
        log('ERRO', ex)
    finally:
        q.fechar(salvar=False)
    OUT['sha_copia_antes_depois'] = [h0, sha(copia)]
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump(j(OUT), f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    if '--produto' in sys.argv:
        k = sys.argv.index('--produto')
        P.update(PRODUTOS[sys.argv[k + 1]])
        del sys.argv[k:k + 2]
    modo, copia, pasta, saida = sys.argv[1], os.path.abspath(sys.argv[2]), os.path.abspath(sys.argv[3]), sys.argv[4]
    if modo == 'A':
        sessao_a(copia, pasta, saida)
    elif modo == 'C':
        sessao_c(copia, pasta, saida, sys.argv[5])
    elif modo == 'D':
        sessao_d(copia, pasta, saida, sys.argv[5])
    elif modo == 'E':
        sessao_e(copia, pasta, saida, sys.argv[5])
    elif modo == 'F':
        sessao_f(copia, pasta, saida, sys.argv[5])
    elif modo == 'G':
        sessao_g(copia, pasta, saida, sys.argv[5])
    else:
        sessao_b(copia, pasta, saida, sys.argv[5])
    log('FIM', modo, 'erros:', len(OUT['erros']))
