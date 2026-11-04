# -*- coding: utf-8 -*-
"""valida_adr088.py -- Tester/QA, ciclo 2026-11-04_reteste01, modo VALIDAR (ADR-088, commit aec2c43, entrega 35).

Escrito pelo Tester/QA (quem corrigiu foi o agente programador). Reaproveita repro_comum (ciclo 01) e qa_final.QA.
Uma sessao de Excel por produto, sobre COPIA da producao atual em %TEMP%/qcw/auditoria/2026-11-04_reteste01/.
Nada e salvo; fechar() encerra so o PID que o QA criou; nenhuma captura de tela.

Roteiro (os dois produtos):
  P1  estado da producao entregue (antes de qualquer ATUALIZAR): <= 0 participantes, IMPLAUSIVEL, formulas H7:H9, B103
  P2  OC-0002 ET do Painel sem bias numerico; OC-0006 bias; OC-0013 H3 vazio; OC-0009 B103; OC-0003 troca de lote
  P3  ATUALIZAR (mIntegracao.AtualizarSoEntradas, sem manuais) -> OC-0004 (IMPLAUSIVEL = A11 = varredura independente)
  P4  manuais TESTE-AUDITORIA (minusculas, 31/02, -1, 0, 28/02) + ATUALIZAR -> OC-0001, OC-0005, OC-0004 manual
Uso: python -u valida_adr088.py <Bioquimica|Hematologia>
"""
import datetime as dt
import os
import re
import shutil
import sys
import time
import traceback
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

PROD = sys.argv[1]
BIO = PROD == 'Bioquimica'
NLV = 2 if BIO else 3
C.PASTA = os.path.join(os.environ['TEMP'], 'qcw', 'auditoria', '2026-11-04_reteste01')
RID = re.compile(r'\b(?:BIO|HEM)-\d+|\bMAN_\d+')


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def ids_codigo(qa, cod):
    out = set()
    for r in qa.ler('tblQA_Integracao'):
        if any(str(v).strip() == cod for v in r.values() if v is not None):
            out |= set(RID.findall(' '.join(str(v) for v in r.values() if v is not None)))
    return out


def dup_chaves(fin):
    c = Counter((str(r['EQUIPAMENTO']).strip().upper(), str(r['LOTE']), r['NIVEL'], str(r['ANALITO']).strip().upper(),
                 r['RUN']) for r in fin if r.get('PARTICIPA_ESTATISTICA') == 'SIM')
    return [k for k, n in c.items() if n > 1]


def celula_rotulo(ws, linha, rotulo):
    v = C.bloco(ws, f'A{linha}:AD{linha}')[0]
    for i, x in enumerate(v):
        if str(x).strip() == rotulo:
            return ws.Cells(linha, i + 2).Address.replace('$', '')
    return None


def filtro(qa, p, an, lote, eq=None, ano=2026, mes='Out'):
    C.por(qa, p, 'C3', an)
    if eq:
        C.por(qa, p, 'L4', eq)
    C.por(qa, p, 'H3', lote)
    C.por(qa, p, 'K3', ano, local=False)
    C.por(qa, p, 'K4', mes)


def h_ok(pn):
    """OC-0002: H vazio quando G nao e numero; H = 1,65*CV + |G| quando e (1e-9)."""
    out = []
    for x in pn['niveis']:
        if num(x['bias']):
            out.append(num(x['et']) and C.rel(x['et'], 1.65 * x['cv'] + abs(x['bias'])) < 1e-9)
        else:
            out.append(x['et'] in (None, ''))
    return all(out), out


def main():
    os.makedirs(C.PASTA, exist_ok=True)
    src = os.path.join(C.REPO, f'QC_{PROD}.xlsm')
    copia = os.path.join(C.PASTA, f'{PROD}_validar.xlsm')
    h0 = C.sha(src)
    shutil.copy2(src, copia)
    L = C.Log(os.path.join(C.PASTA, f'valida_adr088_{PROD}.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod': h0, 'copia': copia}
    senha = C.senha_planilha()
    qa = C.qa_final.QA(PROD, copia)
    L['excel_pid'] = qa.ex.pid
    try:
        qa.run('mSeguranca.ReprotectAll')
        p, est = qa.ws('Painel'), qa.ws('Estatística')
        for w in (p, est, qa.ws('Digitar Resultados')):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        # ------------------------------------------------------------------------------------------- P1
        fin0 = C.final(qa)
        le0 = {r['ID_REGISTRO'] for r in fin0 if num(r['RESULTADO']) and r['RESULTADO'] <= 0
               and r['PARTICIPA_ESTATISTICA'] == 'SIM'}
        ativ_le0 = {r['ID_REGISTRO'] for r in fin0 if num(r['RESULTADO']) and r['RESULTADO'] <= 0
                    and r['STATUS_ANALITICO'] in ('ATIVO', 'IMPLAUSIVEL')}
        L['P1'] = {'n_final': len(fin0), 'participantes_le0': len(le0),
                   'lt0': sum(1 for r in fin0 if r['ID_REGISTRO'] in le0 and r['RESULTADO'] < 0),
                   'implausivel': sum(1 for r in fin0 if r['STATUS_ANALITICO'] == 'IMPLAUSIVEL'),
                   'H7': p.Range('H7').Formula, 'H8': p.Range('H8').Formula,
                   'H9': p.Range('H9').Formula if not BIO else None,
                   'B103': (est.Range('B103').Formula, est.Range('B103').Value) if BIO else None}

        # ------------------------------------------------------------------------------------------- P2
        prov = 'M4' if BIO else celula_rotulo(est, 4, 'Provedor')
        provws = p if BIO else est
        r2 = {}
        if BIO:
            casos = (('Hemoglobina glicada', '8590', 'DIMENSION 2'), ('Proteína C-reativa', '8353', 'DIMENSION 2'),
                     ('Proteína total', '8977', 'DIMENSION 2'))
        else:
            casos = (('RET%', '619611', None), ('RDW-SD', '619611', None), ('RBC', '619611', None),
                     ('LYMPH#', '619611', None))
        for an, lote, eq in casos:
            filtro(qa, p, an, lote, eq, mes='Out' if BIO else 'Set')
            pn = C.painel(qa, NLV)
            ok, det = h_ok(pn)
            r2[an + '|CAP'] = {'ok': ok, 'det': det, 'niveis': [(x['n'], x['cv'], x['bias'], x['et']) for x in pn['niveis']]}
            if an in ('Proteína total', 'RBC', 'LYMPH#'):
                C.por(qa, provws, prov, 'Controllab', local=False)
                pn = C.painel(qa, NLV)
                ok, det = h_ok(pn)
                r2[an + '|Controllab'] = {'ok': ok, 'det': det, 'B103': est.Range('B103').Value if BIO else None,
                                          'niveis': [(x['n'], x['cv'], x['bias'], x['et']) for x in pn['niveis']]}
                C.por(qa, provws, prov, 'CAP', local=False)
        L['OC0002'] = r2
        e = C.estat_tabela(qa)
        if BIO:
            L['OC0006'] = {k: (v['G_bias'], v['AC_nEQA'], est.Cells(v['linha'], 30).Value) for k, v in e.items()
                           if 'ilirrubina' in k}
            L['OC0009'] = {'B103': est.Range('B103').Value,
                           'indep_sem_ep_com_analito': sum(1 for v in e.values() if v['G_bias'] == 'SEM EP')}
        else:
            L['OC0006'] = {k: v['G_bias'] for k, v in e.items() if k.split('|')[0] in ('IG%', 'IG#', 'NRBC%', 'NRBC#')}
        # OC-0013: H3 vazio
        filtro(qa, p, 'Glicose' if BIO else 'WBC', '', 'DIMENSION 1' if BIO else None, mes='Out' if BIO else 'Set')
        pn = C.painel(qa, NLV)
        e = C.estat_tabela(qa)
        an = 'Glicose' if BIO else 'WBC'
        L['OC0013'] = {'G6': p.Range('G6').Text, 'I6': p.Range('I6').Text, 'H3': p.Range('H3').Value,
                       'niveis': [(x['n'], x['media'], x['dp'], x['cv'], x['bias'], x['et'], x['sigma']) for x in pn['niveis']],
                       'estat_G_L': [(e[f'{an}|{n}']['G_bias'], e[f'{an}|{n}']['L_sigma']) for n in range(1, NLV + 1)]}
        # OC-0003: troca de lote (Hema 522611 <-> 619611; Bio 8974 <-> 8977) e Ano De/Ate
        t3 = {}
        alvo = ('RBC', ('522611', '619611', '522611')) if not BIO else ('Glicose', ('8974', '8977', '8974'))
        if BIO:
            C.por(qa, p, 'L4', 'DIMENSION 1')
            C.por(qa, p, 'K4', 'Out')      # fora do 'Tudo' (OC-0012, janela de 180 corridas, aberta)
        C.por(qa, p, 'C3', alvo[0])
        for i, lote in enumerate(alvo[1]):
            t0 = time.time()
            C.por(qa, p, 'H3', lote)
            s = round(time.time() - t0, 1)
            e = C.estat_tabela(qa)
            pn = C.painel(qa, NLV)
            ano = (est.Range('B3').Value, est.Range('D3').Value)
            t3[f'{i}_{lote}'] = {'s': s, 'B4': est.Range('B4').Value, 'B3_D3': ano,
                                 'estat': {n: (e[f'{alvo[0]}|{n}']['n'], e[f'{alvo[0]}|{n}']['cv'], e[f'{alvo[0]}|{n}']['H_et'])
                                           for n in range(1, NLV + 1)},
                                 'painel': [(x['n'], x['cv'], x['et']) for x in pn['niveis']]}
            if BIO:
                O = dt.datetime(2026, 10, 1), dt.datetime(2026, 10, 31, 23, 59)
                t3[f'{i}_{lote}']['K3_K4_L4'] = (p.Range('K3').Value, p.Range('K4').Value, p.Range('L4').Value)
                t3[f'{i}_{lote}']['indep_tudo_DIM1'] = {n: C.mdcv([r['RESULTADO'] for r in C.pop(
                    fin0, alvo[0], lote, n, 'DIMENSION 1')]) for n in (1, 2)}
                for eqx in ('DIMENSION 1', None):
                    t3[f'{i}_{lote}'][f'indep_out_{eqx}'] = {n: C.mdcv([r['RESULTADO'] for r in C.pop(
                        fin0, alvo[0], lote, n, eqx, *O)]) for n in (1, 2)}
            if not BIO:
                a0, a1 = int(ano[0]), int(ano[1])
                t3[f'{i}_{lote}']['indep'] = {n: C.mdcv([r['RESULTADO'] for r in C.pop(
                    fin0, alvo[0], lote, n, 'XN1000', dt.datetime(a0, 1, 1), dt.datetime(a1, 12, 31, 23, 59))])
                    for n in range(1, NLV + 1)}
        if not BIO:
            C.por(qa, p, 'H3', '619611')
            for b3, d3 in ((2026, 2026), (2025, 2025)):
                C.por(qa, est, 'B3', b3, local=False)
                C.por(qa, est, 'D3', d3, local=False)
                e = C.estat_tabela(qa)
                t3[f'ano_{b3}_{d3}'] = {'B4': est.Range('B4').Value,
                                        'estat': {n: (e[f'RBC|{n}']['n'], e[f'RBC|{n}']['cv']) for n in (1, 2, 3)},
                                        'indep': {n: C.mdcv([r['RESULTADO'] for r in C.pop(
                                            fin0, 'RBC', '619611', n, 'XN1000', dt.datetime(b3, 1, 1),
                                            dt.datetime(d3, 12, 31, 23, 59))]) for n in (1, 2, 3)}}
        L['OC0003'] = t3

        # ------------------------------------------------------------------------------------------- P3 ATUALIZAR
        pend = [r['ID_REGISTRO'] for r in C.manuais(qa)]
        L['manuais_existentes'] = len(pend)
        e_antes = C.estat_tabela(qa)
        t0 = time.time()
        res = str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))
        L['P3_atualizar'] = {'ret': res[:300], 's': round(time.time() - t0, 1)}
        fin1 = C.final(qa)
        imp = {r['ID_REGISTRO']: r for r in fin1 if r['STATUS_ANALITICO'] == 'IMPLAUSIVEL'}
        a11 = ids_codigo(qa, 'A11')
        p0 = {r['ID_REGISTRO']: r['PARTICIPA_ESTATISTICA'] for r in fin0}
        mud = [r['ID_REGISTRO'] for r in fin1 if r['ID_REGISTRO'] not in imp and p0.get(r['ID_REGISTRO']) not in (None, r['PARTICIPA_ESTATISTICA'])]
        e_dep = C.estat_tabela(qa)
        L['OC0004'] = {
            'implausivel': len(imp), 'igual_varredura_participantes_le0': set(imp) == le0,
            'igual_varredura_ativos_le0': set(imp) == ativ_le0,
            'A11': len(a11), 'A11_igual_implausivel': a11 == set(imp),
            'flags_ok': all(r['PARTICIPA_ESTATISTICA'] != 'SIM' and r['REGISTRAR_RESULTADO_NO_LJ'] != 'SIM' and r['TIPO_PLOTAGEM_LJ'] == 'NAO_PLOTAR'
                            for r in imp.values()),
            'flags_amostra': [(k, r['PARTICIPA_ESTATISTICA'], r['REGISTRAR_RESULTADO_NO_LJ'], r['TIPO_PLOTAGEM_LJ'],
                               str(r.get('MOTIVO_EXCLUSAO_AUTOMATICA'))[:60]) for k, r in list(imp.items())[:2]],
            'participantes_le0_depois': sum(1 for r in fin1 if num(r['RESULTADO']) and r['RESULTADO'] <= 0
                                            and r['PARTICIPA_ESTATISTICA'] == 'SIM'),
            'n_final': len(fin1), 'outros_PARTICIPA_mudaram': mud[:20], 'n_outros_mudaram': len(mud),
            'run_mudou_implausivel': sum(1 for r in fin0 if r['ID_REGISTRO'] in imp and
                                         r['RUN'] != next(x['RUN'] for x in [imp[r['ID_REGISTRO']]])),
            'estat_linhas_mudaram': sorted(k for k in e_antes if any(e_antes[k][c] != e_dep.get(k, {}).get(c)
                                                                     for c in ('n', 'cv', 'G_bias', 'L_sigma'))),
            'analitos_implausivel': sorted(Counter(f"{r['ANALITO']}|{int(r['NIVEL'])}" for r in imp.values()).items()),
            'dup_chaves': len(dup_chaves(fin1))}
        if BIO:
            filtro(qa, p, 'Glicose', '8977', 'DIMENSION 1')
            pn = C.painel(qa, NLV)
            e = C.estat_tabela(qa)
            OUT = dt.datetime(2026, 10, 1), dt.datetime(2026, 10, 31, 23, 59)
            L['OC0004_glicose_8977_dim1_out'] = {
                'painel': [(x['n'], x['media'], x['cv']) for x in pn['niveis']],
                'estat': [(e[f'Glicose|{n}']['n'], e[f'Glicose|{n}']['media'], e[f'Glicose|{n}']['cv']) for n in (1, 2)],
                'indep': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'Glicose', '8977', n, 'DIMENSION 1', *OUT)])
                          for n in (1, 2)},
                'N2_valores': sorted((r['RESULTADO'], r['ID_REGISTRO'], str(r['DATA_HORA']), r['STATUS_ANALITICO'])
                                     for r in fin1 if r['ANALITO'] == 'Glicose' and str(r['LOTE']) == '8977'
                                     and r['EQUIPAMENTO'] == 'DIMENSION 1' and r['NIVEL'] == 2 and isinstance(r['DATA'], dt.datetime)
                                     and OUT[0] <= r['DATA'] <= OUT[1])}
            # inativacao simulada pelo laboratorio (parecer G4, acao 1): BIO-371184 (= 1, positivo, fora da regra <= 0)
            ri = str(qa.run('mIntegracao.RegistrarInativacao', 'BIO-371184', 'Glicose', C.MOTIVO, teto=120))
            res = str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))
            filtro(qa, p, 'Glicose', '8977', 'DIMENSION 1')
            pn = C.painel(qa, NLV)
            fin1 = C.final(qa)
            L['OC0004_apos_inativar_BIO-371184'] = {
                'ret': ri[:200], 'painel': [(x['n'], x['media'], x['cv']) for x in pn['niveis']],
                'indep': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'Glicose', '8977', n, 'DIMENSION 1', *OUT)])
                          for n in (1, 2)}}
            an_m, lote_m, eq_m, eqd, and_ = 'Glicose', '8977', 'DIMENSION 1', 'dimension 1', 'glicose'
            OUT = dt.datetime(2026, 10, 1), dt.datetime(2026, 10, 31, 23, 59)
            nm = 2
        else:
            filtro(qa, p, 'WBC', '619611', None, mes='Tudo (sem filtro)')
            pn = C.painel(qa, NLV)
            L['OC0004_wbc_619611'] = {'painel': [(x['n'], x['cv']) for x in pn['niveis']],
                                      'indep': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'WBC', '619611', n, 'XN1000')])
                                                for n in (1, 2, 3)}}
            an_m, lote_m, eq_m, eqd, and_ = 'WBC', '619611', 'XN1000', 'xn1000', 'wbc'
            OUT = dt.datetime(2026, 9, 1), dt.datetime(2026, 9, 30, 23, 59)
            nm = 3

        # ------------------------------------------------------------------------------------------- P4 manuais
        g = [r for r in fin1 if r['ANALITO'] == an_m and str(r['LOTE']) == lote_m and r['EQUIPAMENTO'] == eq_m and
             r['NIVEL'] == nm and r['PARTICIPA_ESTATISTICA'] == 'SIM' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
             and OUT[0] <= r['DATA'] <= OUT[1]]
        dias = sorted({r['DATA'] for r in g if str(int(r['RUN'] or 0)).endswith('01')}, reverse=True)
        d1, d2, dn, dz = dias[:4]
        reais = {str(d): [(r['ID_REGISTRO'], int(r['RUN']), r['RESULTADO']) for r in g if r['DATA'] == d
                     and str(int(r['RUN'])).endswith('01')] for d in (d1, d2)}
        runs = {reais[str(d1)][0][1], reais[str(d2)][0][1]}
        if BIO:
            filtro(qa, p, an_m, lote_m, eq_m, mes='Out')
        else:
            filtro(qa, p, an_m, lote_m, None, mes='Tudo (sem filtro)')
        eng0 = C.eng_runs(qa, runs, NLV)
        f = lambda d: d.strftime('%d/%m/%Y')  # noqa: E731
        vm = '244,222' if BIO else '17,433'
        base = lambda eqx, d, h, niv, anx, res: [('EQUIPAMENTO', eqx), ('DATA', d), ('HORA', h), ('LOTE', lote_m),  # noqa: E731
                                                  ('NIVEL', str(niv)), ('ANALITO', anx), ('RESULTADO', res), ('MOTIVO', C.MOTIVO)]
        casos = {'G1_equip_minusc': base(eqd, f(d1), '23:36', nm, an_m, vm),
                 'G1_analito_minusc': base(eq_m, f(d2), '23:36', nm, and_, vm),
                 'G2_31_02': base(eq_m, '31/02/2026', '08:00', 1, an_m, '3,111' if not BIO else '77,111'),
                 'G2_28_02_ok': base(eq_m, '28/02/2026', '08:00', 1, an_m, '3,111' if not BIO else '77,111'),
                 'G4_menos1': base(eq_m, f(dn), '23:36', 1, an_m, '-1'),
                 'G4_zero': base(eq_m, f(dz), '23:36', 1, an_m, '0')}
        ids, tela = {}, {}
        for k, cs in casos.items():
            r, i, lida = C.escrever_manual(qa, cs)
            ids[k] = i
            tela[k] = {'EQUIPAMENTO': lida.get('EQUIPAMENTO'), 'ANALITO': lida.get('ANALITO'), 'DATA': lida.get('DATA')}
        L['P4_digitado'] = {'ids': ids, 'na_tabela_ao_digitar': tela, 'reais': reais}
        t0 = time.time()
        res = str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))
        L['P4_atualizar'] = {'ret': res[:300], 's': round(time.time() - t0, 1)}
        fin2 = C.final(qa)
        porid = {r['ID_REGISTRO']: r for r in fin2}
        man = {r['ID_REGISTRO']: r for r in C.manuais(qa)}
        cps = ('EQUIPAMENTO', 'ANALITO', 'DATA', 'RUN', 'NIVEL', 'RESULTADO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA',
               'REGISTRAR_RESULTADO_NO_LJ')
        L['P4_final'] = {k: {c: porid.get(i, {}).get(c) for c in cps} for k, i in ids.items()}
        L['P4_tela'] = {k: {c: man.get(i, {}).get(c) for c in ('STATUS', 'RUN', 'DETALHE')} for k, i in ids.items()}
        L['P4_codigos'] = {c: sorted(set(ids.values()) & ids_codigo(qa, c)) for c in ('A11', 'A12', 'E04')}
        if BIO:
            filtro(qa, p, an_m, lote_m, eq_m, mes='Out')
        else:
            filtro(qa, p, an_m, lote_m, None, mes='Tudo (sem filtro)')
            C.por(qa, est, 'B3', 2026, local=False)     # o lote 619611 so tem dados de 2026
            C.por(qa, est, 'D3', 2026, local=False)
        pn = C.painel(qa, NLV)
        e = C.estat_tabela(qa)
        jan = OUT if BIO else (None, None)
        mx = {n: C.pop(fin2, an_m, lote_m, n, eq_m, *jan, maiusc=True) for n in range(1, NLV + 1)}
        eng1 = C.eng_runs(qa, runs, NLV)
        L['OC0001'] = {
            'painel': [(x['n'], x['media'], x['cv']) for x in pn['niveis']],
            'estat': [(e[f'{an_m}|{n}']['n'], e[f'{an_m}|{n}']['media']) for n in range(1, NLV + 1)],
            'estat_B3_D3_B4': None if BIO else (est.Range('B3').Value, est.Range('D3').Value, est.Range('B4').Value),
            'indep_maiusc': {n: C.mdcv([r['RESULTADO'] for r in mx[n]]) for n in mx},
            'indep_um_por_run': {n: C.mdcv(C.um_por_run(mx[n])) for n in mx},
            'eng_real_antes': {rn: (v.get(f'N{nm}_Valor'), v.get(f'N{nm}_Veredicto')) for rn, v in eng0.items()},
            'eng_real_depois': {rn: (v.get(f'N{nm}_Valor'), v.get(f'N{nm}_Veredicto')) for rn, v in eng1.items()},
            'dup_chaves_toda_final': len(dup_chaves(fin2))}
    except Exception:
        L['ERRO'] = traceback.format_exc()
    finally:
        qa.fechar()
        L['sha_prod_depois'] = C.sha(src)
        L['fim'] = dt.datetime.now().isoformat(timespec='seconds')


if __name__ == '__main__':
    main()
