# -*- coding: utf-8 -*-
"""reteste05_usuario_f1_hema.py -- Usuario Simulado, ciclo 2026-11-15_reteste05, frente F1, modo RETESTE (Hematologia).

OC-0012 (efeito colateral R4-H01 do reteste 04): na Hema o aviso do corte de 180 corridas (Painel!A39) mandava "veja a
aba Estatistica", que na Hema tem lote e ano proprios. Correcao declarada: ADR-095 par.4 (instalar_adr078.corte_periodo,
entrega 42). Criterio de aceite: na Hema o A39 diz "Periodo inteiro (todos os lotes): N1 .. · N2 .. · N3 .. resultados."
com H3 vazio e "Periodo inteiro do lote: .." com lote, sem citar a Estatistica; contagens = Eng_Saida!AF1 = dados de
origem; indicadores e grafico inalterados. Na Bio o aviso continua apontando para a Estatistica (regressao).

Metodo (caminho do usuario: celulas de filtro do Painel C3/H3/K3/K4/G3/G4 escritas com os eventos ligados), um Excel por
vez, copias NOVAS em %TEMP%/qcw/auditoria/2026-11-15_reteste05/:
  - BASE = copia do backup pre-entrega 42 da Hema (_backup_pre_integracao_2026-11-15_0458, entrega 41, a que o reteste
    04 testou);
  - NOVA = copia da producao atual da Hema (raiz);
  - BIO  = copia da producao atual da Bio (so regressao do texto).
Em cada cenario le: A39 (Value e Formula), Eng_Saida!M1 (corridas cortadas) e AF1 (n do periodo inteiro), engPainel
(n/media/DP/CV do Painel), Painel A6:Q12 (Value2), Calc!A1:AZ190 (Value2, origem do Levey-Jennings) e Eng_Saida!A1:AF190.
BASE x NOVA: celula a celula na precisao total do Value2 (tolerancia zero; so vale para cenarios cujo periodo termina
antes de qualquer resultado novo, por isso ha cenarios com De/Ate fechados em 31/10/2026; a tblCQ_Final das duas
copias e comparada por n de linhas e hash).
Oraculo proprio (Python puro, tblCQ_Final lida da copia): resultados com PARTICIPA_ESTATISTICA = SIM e RESULTADO
numerico, do analito, (do lote, se houver), com data dentro de [De, Ate]. AF1 esperado = "N1 a · N2 b · N3 c" (contagem
por nivel). Corridas = pares (RUN, LOTE) distintos (ADR-080; no modo TODOS um RUN com dois lotes conta duas vezes,
conferido no reteste 04); janela = as 180 mais recentes do periodo; cortadas = total - 180; n da janela por nivel.
Datas do De/Ate escritas como Value2 (serial) para nao herdar o fuso do COM.

Regras: so copias; nada e salvo; fechar() encerra so o PID do QA; nenhuma captura de tela; nenhum dado de teste gravado.
Uso: python -u reteste05_usuario_f1_hema.py -> %TEMP%/qcw/auditoria/2026-11-15_reteste05/Hematologia_F1.json
     python -u reteste05_usuario_f1_hema.py --lote-sintetico -> Hematologia_F1_lote.json
       Nenhum lote real da Hema tem mais de 180 corridas (maximo 89), entao o ramo "Periodo inteiro do lote" nao aparece
       com os dados reais. Neste modo (copia nova da producao, descartada), as linhas de WBC dos lotes 533811, 614011 e
       619611 da tblCQ_Final recebem o LOTE 528211 (dado SINTETICO de teste, so na copia; eventos desligados), o cache
       do motor e invalidado (mEstatistica.InvalidarCache, unico passo fora do caminho do usuario) e o Painel e filtrado
       pelo usuario: WBC + H3 528211 + "Tudo"; + K3 2026 "Ano inteiro"; + H3 vazio (todos os lotes). Oraculo relido
       depois da troca.
"""
import collections
import datetime as dt
import hashlib
import json
import os
import shutil
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import repro_comum as C  # noqa: E402

CICLO = '2026-11-15_reteste05'
PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', CICLO)
C.PASTA = PASTA
BACKUP_HEMA = os.path.join(C.REPO, '_backup_pre_integracao_2026-11-15_0458', 'QC_Hematologia.xlsm')
SAIDA = os.path.join(PASTA, 'Hematologia_F1.json')


def t(v):
    return '' if v is None else str(v).strip()


def up(v):
    return t(v).upper()


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def serial(d):
    return float((d - dt.datetime(1899, 12, 30)).days)


def oraculo(fin, an, lote, de, ate, nlv, nk=180):
    runs = collections.defaultdict(collections.Counter)
    tot = collections.Counter()
    for r in fin:
        if up(r['ANALITO']) != up(an) or t(r['PARTICIPA_ESTATISTICA']) != 'SIM' or not num(r['RESULTADO']):
            continue
        if lote and t(r['LOTE']).lstrip("'") != lote:
            continue
        d = C.limpo(r['DATA_HORA'])
        if not isinstance(d, dt.datetime):
            continue
        if de is not None and d.date() < de.date():
            continue
        if ate is not None and d.date() > ate.date():
            continue
        nv = int(r['NIVEL'] or 0)
        tot[nv] += 1
        runs[(t(r['RUN']), t(r['LOTE']))][nv] += 1
    ks = sorted(runs)
    jan = ks[-nk:]
    nj = collections.Counter()
    for k in jan:
        nj.update(runs[k])
    so_run = len({k[0] for k in ks})
    return {'corridas': len(ks), 'corridas_so_RUN': so_run, 'cortadas': max(0, len(ks) - nk),
            'n_periodo': [tot.get(k, 0) for k in range(1, nlv + 1)],
            'n_janela': [nj.get(k, 0) for k in range(1, nlv + 1)],
            'AF1_esperado': ' · '.join(f'N{k} {tot.get(k, 0)}' for k in range(1, nlv + 1))}


def aplicar(qa, p, passos):
    for c, v in passos:
        if v is None:
            p.Range(c).MergeArea.ClearContents()
            qa.ex.esperar()
            time.sleep(0.5)
        elif isinstance(v, dt.datetime):
            p.Range(c).Value2 = serial(v)
            qa.ex.esperar()
            time.sleep(0.5)
        else:
            C.por(qa, p, c, v, local=False)
    time.sleep(1.5)
    qa.ex.esperar()


def mat(rng):
    v = rng.Value2
    return [list(x) for x in v] if isinstance(v, tuple) else [[v]]


def diff(a, b):
    out = []
    for i, (la, lb) in enumerate(zip(a, b), 1):
        for j, (x, y) in enumerate(zip(la, lb), 1):
            if x != y:
                out.append([i, j, x, y])
    if len(a) != len(b):
        out.append(['linhas', len(a), len(b)])
    return out


def ler(qa, p, fin, nlv):
    eng = qa.ws('Eng_Saida')
    de, ate = qa.nome('filtroDe').Value, qa.nome('filtroAte').Value
    de = C.limpo(de) if isinstance(de, dt.datetime) else None
    ate = C.limpo(ate) if isinstance(ate, dt.datetime) else None
    todos = eng.Range('AD1').Value is True
    an = t(qa.nome('selAnalito').Value)
    lote = None if todos else t(eng.Range('E1').Value).lstrip("'")
    ps = qa.painel_stats()
    o = {'filtro': {'C3': t(p.Range('C3').Value), 'H3': t(p.Range('H3').Value), 'K3': t(p.Range('K3').Value),
                    'K4': t(p.Range('K4').Value), 'G3': str(C.limpo(p.Range('G3').Value)),
                    'G4': str(C.limpo(p.Range('G4').Value)), 'de': str(de), 'ate': str(ate), 'todos': todos,
                    'analito': an, 'lote': lote},
         'A39': t(p.Range('A39').Value), 'A39_formula_tem_Estat': 'Estatística' in t(p.Range('A39').Formula),
         'M1': eng.Range('M1').Value, 'AF1': t(eng.Range('AF1').Value),
         'engPainel': [list(x) for x in ps],
         'J7': [t(p.Range(f'J{7 + k}').Text) for k in range(nlv)],
         'oraculo': oraculo(fin, an, lote, de, ate, nlv)}
    o['_painel'] = mat(p.Range('A6:Q12'))
    o['_calc'] = mat(qa.ws('Calc').Range('A1:AZ190'))
    o['_eng'] = mat(eng.Range('A1:AF190'))
    try:
        a39 = p.Range('A39').MergeArea
        o['A39_area'] = {'end': t(a39.Address), 'wrap': bool(a39.WrapText), 'altura': a39.Height,
                         'largura': a39.Width, 'len': len(o['A39'])}
    except Exception as ex:  # noqa: BLE001
        o['A39_area'] = repr(ex)
    return o


def avaliar(o, produto):
    orc = o['oraculo']
    ns = [x[1] for x in o['engPainel']]
    r = {'AF1_igual_oraculo': o['AF1'] == orc['AF1_esperado'],
         'M1_igual_oraculo': int(o['M1'] or 0) == orc['cortadas'],
         'n_painel_igual_janela': ns == (orc['n_janela'] if orc['cortadas'] else orc['n_periodo'])}
    a = o['A39']
    if orc['cortadas'] > 0:
        if produto == 'Hematologia':
            rot = 'Período inteiro (todos os lotes): ' if o['filtro']['todos'] else 'Período inteiro do lote: '
            r['texto_hema'] = (rot + o['AF1'] + ' resultados.') in a
            r['sem_Estatistica'] = 'Estatística' not in a
            r['cita_indicadores'] = 'fora do gráfico e dos indicadores do Painel' in a
        else:
            r['texto_bio'] = ('Período inteiro: ' + o['AF1'] + ' — veja a aba Estatística.') in a
    else:
        r['sem_aviso'] = 'fora do gráfico' not in a
    r['ok'] = all(r.values())
    return r


def cenarios_hema(lote_wbc, lote_grande):
    D = dt.datetime
    c = [
        ('H1_original_WBC_todos_Ano2026', [('C3', 'WBC'), ('K3', 2026), ('K4', 'Ano inteiro'), ('H3', None)]),
        ('H2_original_WBC_lote_Ano2026', [('H3', "'" + lote_wbc)]),
        ('H3_WBC_lote_Tudo', [('K4', 'Tudo (sem filtro)')]),
        ('H4_WBC_todos_Tudo', [('H3', None), ('K4', 'Tudo (sem filtro)')]),
        ('H5_WBC_todos_DeAte_jan_out2026', [('K3', 2026), ('K4', 'Ano inteiro'), ('G3', D(2026, 1, 1)),
                                            ('G4', D(2026, 10, 31))]),
        ('H6_WBC_todos_Out2026', [('K3', 2026), ('K4', 'Out')]),
        ('H7_HGB_todos_DeAte_jan_out2026', [('C3', 'HGB'), ('H3', None), ('K3', 2026), ('K4', 'Ano inteiro'),
                                            ('G3', D(2026, 1, 1)), ('G4', D(2026, 10, 31))]),
    ]
    if lote_grande:
        an, lt = lote_grande
        c.append(('H8_lote_com_corte_Tudo', [('C3', an), ('H3', "'" + lt), ('K4', 'Tudo (sem filtro)')]))
        c.append(('H9_lote_com_corte_ate_out2026', [('K3', 2026), ('K4', 'Ano inteiro'), ('G3', D(2020, 1, 1)),
                                                   ('G4', D(2026, 10, 31))]))
    return c


def cenarios_bio():
    return [
        ('B1_Glicose_todos_Tudo', [('L4', 'DIMENSION 1'), ('C3', 'Glicose'), ('H3', None), ('K4', 'Tudo (sem filtro)')]),
        ('B2_Glicose_8974_Tudo', [('H3', "'8974")]),
        ('B3_Glicose_8974_Out2026', [('K3', 2026), ('K4', 'Out')]),
    ]


def escolher_lotes(fin):
    cont = collections.Counter(t(r['LOTE']).lstrip("'") for r in fin if up(r['ANALITO']) == 'WBC'
                               and t(r['PARTICIPA_ESTATISTICA']) == 'SIM')
    lote_wbc = cont.most_common(1)[0][0]
    runs = collections.defaultdict(set)
    for r in fin:
        if t(r['PARTICIPA_ESTATISTICA']) == 'SIM' and num(r['RESULTADO']):
            runs[(t(r['ANALITO']), t(r['LOTE']).lstrip("'"))].add(t(r['RUN']))
    top = sorted(((len(v), k) for k, v in runs.items()), reverse=True)[:10]
    grande = top[0][1] if top and top[0][0] > 180 else None
    return lote_wbc, grande, {'lotes_WBC': dict(cont.most_common(6)),
                              'top_analito_lote_por_corridas': [[k[0], k[1], n] for n, k in top]}


def rodar(rotulo, produto, origem, escolha=None):
    destino = os.path.join(PASTA, f'{produto}_F1_{rotulo}.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    nlv = 3 if produto == 'Hematologia' else 2
    R = {'origem': os.path.relpath(origem, C.REPO), 'sha_origem_12': h0, 'copia': os.path.basename(destino), 'cen': {}}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(produto, destino)
    R['excel_pid'] = qa.ex.pid
    print('EXCEL_PID', qa.ex.pid, rotulo, flush=True)
    try:
        p = qa.ws('Painel')
        for w in (p, qa.ws('Estatística')):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        p.Activate()
        qa.ex.esperar()
        lo = qa.lo('tblCQ_Final')
        dv = lo.DataBodyRange.Value2
        R['tblCQ_Final'] = {'linhas': len(dv), 'hash': hashlib.sha256(repr(dv).encode()).hexdigest()[:16]}
        fin = qa.ler('tblCQ_Final')
        dts = [C.limpo(r['DATA_HORA']) for r in fin if isinstance(r['DATA_HORA'], dt.datetime)]
        R['ultimo_resultado'] = str(max(dts)) if dts else None
        R['participam'] = sum(1 for r in fin if t(r['PARTICIPA_ESTATISTICA']) == 'SIM')
        if produto == 'Hematologia':
            if escolha is None:
                lw, lg, info = escolher_lotes(fin)
                escolha = (lw, lg)
                R['escolha'] = info
            R['lote_wbc'], R['lote_grande'] = escolha
            cens = cenarios_hema(*escolha)
        else:
            cens = cenarios_bio()
        R['A39_formula'] = t(p.Range('A39').Formula)
        for nome, passos in cens:
            print('>>', rotulo, nome, time.strftime('%H:%M:%S'), flush=True)
            try:
                aplicar(qa, p, passos)
                o = ler(qa, p, fin, nlv)
                o['avaliacao'] = avaliar(o, produto)
                R['cen'][nome] = o
                print('   A39:', o['A39'][:300], '| aval', o['avaliacao'], flush=True)
            except Exception as ex:  # noqa: BLE001
                R['cen'][nome] = {'erro': repr(ex), 'trace': traceback.format_exc()[-1500:]}
    except Exception as ex:  # noqa: BLE001
        R['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    R['sha_origem_depois_12'] = C.sha(origem)
    try:
        os.remove(destino)
    except Exception:
        pass
    return R, escolha


def main():
    os.makedirs(PASTA, exist_ok=True)
    L = {'inicio': dt.datetime.now().isoformat(timespec='seconds')}
    L['NOVA'], esc = rodar('nova', 'Hematologia', os.path.join(C.REPO, 'QC_Hematologia.xlsm'))
    L['BASE'], _ = rodar('base41', 'Hematologia', BACKUP_HEMA, esc)
    L['BIO'], _ = rodar('bio', 'Bioquimica', os.path.join(C.REPO, 'QC_Bioquimica.xlsm'))
    comp = {}
    for nome, oN in L['NOVA']['cen'].items():
        oB = L['BASE']['cen'].get(nome, {})
        if '_calc' not in oN or '_calc' not in oB:
            comp[nome] = 'faltou leitura'
            continue
        comp[nome] = {k: (lambda d: {'n': len(d), 'amostra': d[:12]})(diff(oB[k], oN[k]))
                      for k in ('_painel', '_calc', '_eng')}
        comp[nome]['engPainel_igual'] = oB['engPainel'] == oN['engPainel']
        comp[nome]['M1_AF1'] = [[oB['M1'], oB['AF1']], [oN['M1'], oN['AF1']]]
        comp[nome]['A39_base'] = oB['A39']
    L['comparacao_base_nova'] = comp
    for k in ('NOVA', 'BASE', 'BIO'):
        for c in L[k]['cen'].values():
            for x in ('_painel', '_calc', '_eng'):
                c.pop(x, None)
    L['fim'] = dt.datetime.now().isoformat(timespec='seconds')
    json.dump(L, open(SAIDA, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
    print('FIM', flush=True)


def diag_ids(qa, fin, o):
    """Diagnostico da janela: IDs (HEM-) que o motor publicou no Eng_Saida x IDs da janela do oraculo."""
    v = qa.ws('Eng_Saida').Range('A1:BJ200').Value2
    ids_motor = {x for linha in v for x in linha if isinstance(x, str) and x.startswith(('HEM-', 'MAN_'))}
    f = o['filtro']
    de = dt.datetime.fromisoformat(f['de']) if f['de'] != 'None' else None
    ate = dt.datetime.fromisoformat(f['ate']) if f['ate'] != 'None' else None
    runs = collections.defaultdict(list)
    for r in fin:
        if up(r['ANALITO']) != up(f['analito']) or t(r['PARTICIPA_ESTATISTICA']) != 'SIM' or not num(r['RESULTADO']):
            continue
        if f['lote'] and t(r['LOTE']).lstrip("'") != f['lote']:
            continue
        d = C.limpo(r['DATA_HORA'])
        if not isinstance(d, dt.datetime) or (de and d.date() < de.date()) or (ate and d.date() > ate.date()):
            continue
        runs[(t(r['RUN']), t(r['LOTE']))].append((t(r['ID_REGISTRO']), int(r['NIVEL'] or 0), str(d)))
    ks = sorted(runs)
    jan = ks[-180:]
    ids_orc = {x[0] for k in jan for x in runs[k]}
    dup = {f'{k[0]}|{k[1]}': [x[1] for x in runs[k]] for k in jan
           if len(runs[k]) != len({x[1] for x in runs[k]})}
    return {'n_ids_motor': len(ids_motor), 'n_ids_oraculo': len(ids_orc),
            'so_oraculo': sorted(ids_orc - ids_motor)[:12], 'so_motor': sorted(ids_motor - ids_orc)[:12],
            'so_oraculo_corridas': sorted({f'{k[0]}|{k[1]}' for k in jan for x in runs[k] if x[0] not in ids_motor})[:6],
            'corridas_com_nivel_repetido_na_janela': dict(list(dup.items())[:6]),
            'borda': [[f'{k[0]}|{k[1]}', runs[k]] for k in ks[-181:-179]]}


def lote_sintetico():
    os.makedirs(PASTA, exist_ok=True)
    origem = os.path.join(C.REPO, 'QC_Hematologia.xlsm')
    destino = os.path.join(PASTA, 'Hematologia_F1_lote.xlsm')
    if os.path.exists(destino):
        os.remove(destino)
    h0 = C.sha(origem)
    shutil.copy2(origem, destino)
    R = {'inicio': dt.datetime.now().isoformat(timespec='seconds'), 'sha_origem_12': h0, 'cen': {}}
    senha = C.senha_planilha()
    qa = C.qa_final.QA('Hematologia', destino)
    R['excel_pid'] = qa.ex.pid
    print('EXCEL_PID', qa.ex.pid, 'lote', flush=True)
    alvo, outros = '528211', ('533811', '614011', '619611')
    try:
        p = qa.ws('Painel')
        for w in (p, qa.ws('Estatística')):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        lo = qa.lo('tblCQ_Final')
        try:
            lo.Parent.Unprotect(senha)
        except Exception:
            pass
        ca = lo.ListColumns('ANALITO').DataBodyRange.Value2
        cl = lo.ListColumns('LOTE').DataBodyRange
        vl = [list(x) for x in cl.Value2]
        tipo = next((type(v[0]).__name__ for v, a in zip(vl, ca) if t(v[0]).lstrip("'") == alvo), None)
        novo = next(v[0] for v in vl if t(v[0]).lstrip("'") == alvo)
        idx = [i for i, (a, v) in enumerate(zip(ca, vl)) if up(a[0]) == 'WBC' and t(v[0]).lstrip("'") in outros]
        fmt = t(cl.Cells(1, 1).NumberFormat)
        escrito = ("'" + str(novo)) if isinstance(novo, str) and fmt != '@' else novo
        R['troca'] = {'linhas_trocadas': len(idx), 'tipo_original': tipo, 'valor_usado': repr(novo),
                      'formato_celula': fmt}
        calc0 = qa.ex.xl.Calculation
        R['troca']['calculo_original'] = calc0
        qa.ex.xl.EnableEvents = False
        qa.ex.xl.Calculation = -4135
        try:
            for i in idx:
                cl.Cells(i + 1, 1).Value = escrito
        finally:
            qa.ex.xl.Calculation = calc0
            qa.ex.xl.EnableEvents = True
        R['troca']['conferencia'] = sorted({repr(cl.Cells(i + 1, 1).Value2) for i in idx[:50]})
        qa.run('mEstatistica.InvalidarCache')
        fin = qa.ler('tblCQ_Final')
        R['troca']['wbc_528211_participa'] = sum(1 for r in fin if up(r['ANALITO']) == 'WBC' and
                                                t(r['LOTE']).lstrip("'") == alvo and
                                                t(r['PARTICIPA_ESTATISTICA']) == 'SIM')
        cens = [('L1_WBC_lote_sint_Tudo', [('C3', 'WBC'), ('H3', "'" + alvo), ('K4', 'Tudo (sem filtro)')]),
                ('L2_WBC_lote_sint_Ano2026', [('K3', 2026), ('K4', 'Ano inteiro')]),
                ('L3_WBC_todos_Ano2026', [('H3', None)]),
                ('L4_WBC_lote_sint_Out2026', [('H3', "'" + alvo), ('K3', 2026), ('K4', 'Out')])]
        for nome, passos in cens:
            print('>>', nome, time.strftime('%H:%M:%S'), flush=True)
            try:
                aplicar(qa, p, passos)
                o = ler(qa, p, fin, 3)
                o['avaliacao'] = avaliar(o, 'Hematologia')
                o['diag'] = diag_ids(qa, fin, o)
                for x in ('_painel', '_calc', '_eng'):
                    o.pop(x, None)
                R['cen'][nome] = o
                print('   A39:', o['A39'][:330], '| aval', o['avaliacao'], flush=True)
            except Exception as ex:  # noqa: BLE001
                R['cen'][nome] = {'erro': repr(ex), 'trace': traceback.format_exc()[-1500:]}
    except Exception as ex:  # noqa: BLE001
        R['ERRO_FATAL'] = {'erro': repr(ex), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:
            pass
    R['sha_origem_depois_12'] = C.sha(origem)
    try:
        os.remove(destino)
    except Exception:
        pass
    json.dump(R, open(os.path.join(PASTA, 'Hematologia_F1_lote.json'), 'w', encoding='utf-8'), ensure_ascii=False,
              indent=1, default=str)
    print('FIM', flush=True)


if __name__ == '__main__':
    t0 = time.time()
    if '--lote-sintetico' in sys.argv:
        lote_sintetico()
    else:
        main()
    print(f'tempo {time.time() - t0:.0f}s', flush=True)
