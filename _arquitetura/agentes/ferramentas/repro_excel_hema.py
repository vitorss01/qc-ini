# -*- coding: utf-8 -*-
"""repro_excel_hema.py -- Tester/QA, ciclo 2026-10-31_ciclo01, modo REPRODUZIR, HEMATOLOGIA (uma sessao de Excel).

Grupos cobertos (ids do tester_reproduzir.md):
  G13 troca de lote no Painel (522611 -> 619611) e Ano De/Ate: a Estatistica fica com B4 novo e numeros do anterior
  G1  minusculas no EQUIPAMENTO ("xn1000") e no ANALITO ("wbc") do manual -> Painel x Estatistica
  G2  manual implausivel (-1 no N1, 0 no N2) entra ATIVO sem alerta
  G3  data invalida colada (31/02/2026)
  G9  ET do Painel com bias SEM EP (RET%; RBC com provedor Controllab)
  G5  correcao do CEQ pela janela (SalvarCEQ mesma chave) e protecao da aba
  G6  correcao na tabela so chega ao bias depois do Consolidar; Z2/Z3
  G7  CEQ com nome fora do mapa / Controllab gravado OK e fora do bias; provedor Controllab zera o bias de todos
  G8  "Lymphs" (sem Absolute) entra no bias do LYMPH%
  G14 ampliacao: alvo (MediaDoLote/DPDoLote) x media medida em todos os lotes

Caminho do usuario como no repro_excel_bio.py. Dados de teste: MOTIVO 'Outro: TESTE-AUDITORIA'; WBC 17,433 / -1 / 0 /
3,111; CEQ rodada 'TESTE-AUDITORIA R1', amostras TA-*. Uso: python -u repro_excel_hema.py
"""
import datetime as dt
import os
import sys
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

NLV = 3


def celula_rotulo(ws, linha, rotulo, ncol=30):
    v = C.bloco(ws, f'A{linha}:{chr(64 + ncol) if ncol <= 26 else "AD"}{linha}')[0]
    for i, x in enumerate(v):
        if str(x).strip() == rotulo:
            return ws.Cells(linha, i + 2).Address.replace('$', '')
    return None


def main():
    src, copia, h0 = C.copiar('Hematologia')
    L = C.Log(os.path.join(C.PASTA, 'repro_excel_hema.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod': h0, 'copia': copia, 'sha_copia_antes': C.sha(copia)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA('Hematologia', copia)
    L['excel_pid'] = qa.ex.pid
    try:
        # estado de uma sessao com login de nao-ADM: protecao com UserInterfaceOnly (mSeguranca.UnlockApp ->
        # ReprotectAll). Sem isto o arquivo aberto por automacao fica protegido SEM UserInterfaceOnly e o botao
        # "Consolidar na EQA_Base" (AtualizarEQABase, sem tratamento de erro) para no erro 1004 -- artefato do metodo.
        qa.run('mSeguranca.ReprotectAll')
        p = qa.ws('Painel')
        est = qa.ws('Estatística')
        for w in (p, est, qa.ws('Digitar Resultados')):
            try:
                w.Unprotect(senha)
            except Exception:
                pass
        fin0 = C.final(qa)
        L['n_final_antes'] = len(fin0)

        # ---------------------------------------------------------------- G13: troca de lote / anos
        ks = ('WBC|1', 'RBC|1', 'RBC|2', 'RBC|3')
        e0 = C.estat_tabela(qa)
        g13 = {'antes': {'B3': est.Range('B3').Value, 'D3': est.Range('D3').Value, 'B4': est.Range('B4').Value,
                         'estat': {k: e0[k] for k in ks}, 'painel': C.painel(qa, NLV)}}
        C.por(qa, p, 'H3', '619611')
        e1 = C.estat_tabela(qa)
        g13['apos_H3_619611'] = {'B3': est.Range('B3').Value, 'D3': est.Range('D3').Value, 'B4': est.Range('B4').Value,
                                 'estat': {k: e1[k] for k in ks}, 'painel': C.painel(qa, NLV),
                                 'celulas_mudaram': sum(1 for k in e0 for c in e0[k] if c != 'linha' and e0[k][c] != e1[k][c])}
        qa.run('mApp.IrEstatistica')
        e2 = C.estat_tabela(qa)
        g13['apos_IrEstatistica'] = {'celulas_mudaram': sum(1 for k in e0 for c in e0[k] if c != 'linha' and e0[k][c] != e2[k][c])}
        C.por(qa, est, 'B3', 2026, local=False)
        C.por(qa, est, 'D3', 2026, local=False)
        e3 = C.estat_tabela(qa)
        g13['apos_B3_D3_2026'] = {'B4': est.Range('B4').Value, 'estat': {k: e3[k] for k in ks},
                                  'celulas_mudaram': sum(1 for k in e0 for c in e0[k] if c != 'linha' and e0[k][c] != e3[k][c])}
        ind = {}
        for k in ks:
            a, n = k.split('|')
            for lote, ano in (('522611', 2025), ('619611', 2026), ('619611', 2025)):
                v = [r['RESULTADO'] for r in C.pop(fin0, a, lote, int(n), 'XN1000',
                                                   dt.datetime(ano, 1, 1), dt.datetime(ano, 12, 31, 23, 59))]
                ind[f'{k}|{lote}|{ano}'] = C.mdcv(v)
        g13['independente'] = ind
        qa.run('mEstatistica.AtualizarEstatisticaAba')
        e4 = C.estat_tabela(qa)
        g13['apos_AtualizarEstatisticaAba'] = {'estat': {k: e4[k] for k in ks}}
        L['G13_troca_lote'] = g13

        # ---------------------------------------------------------------- escolher dias (WBC N3, lote 619611)
        # producao: o WBC N3 de outubro tem 1 resultado (lote 625311); setembro/619611 tem 27 -> setembro
        OUT = dt.datetime(2026, 9, 1), dt.datetime(2026, 9, 30)
        g = [r for r in fin0 if r['ANALITO'] == 'WBC' and str(r['LOTE']) == '619611' and r['EQUIPAMENTO'] == 'XN1000'
             and r['NIVEL'] == 3 and r['PARTICIPA_ESTATISTICA'] == 'SIM' and r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO'
             and OUT[0] <= r['DATA'] <= OUT[1]]
        dias = sorted({r['DATA'] for r in g if str(int(r['RUN'] or 0)).endswith('01')}, reverse=True)
        d1, d2 = dias[0], dias[1]
        dneg, dzero = dias[2], dias[3]
        real1 = [r for r in g if r['DATA'] == d1 and str(int(r['RUN'] or 0)).endswith('01')]
        real2 = [r for r in g if r['DATA'] == d2 and str(int(r['RUN'] or 0)).endswith('01')]
        run1, run2 = int(real1[0]['RUN']), int(real2[0]['RUN'])
        L['dias_escolhidos'] = {'d1': d1, 'real1': [(r['ID_REGISTRO'], r['RUN'], r['RESULTADO']) for r in real1],
                                'd2': d2, 'real2': [(r['ID_REGISTRO'], r['RUN'], r['RESULTADO']) for r in real2],
                                'dneg': dneg, 'dzero': dzero}

        L['validacao'] = {f'{c}={t}': C.validacao(qa, c, t) for c, t in
                          (('EQUIPAMENTO', 'xn1000'), ('EQUIPAMENTO', 'XN1000'), ('ANALITO', 'wbc'),
                           ('ANALITO', 'WBC'), ('DATA', '31/02/2026'), ('RESULTADO', '-1'), ('RESULTADO', '0'))}

        lo, rr = C.linha_livre_manual(qa)
        fv = {}
        for col in ('EQUIPAMENTO', 'ANALITO', 'DATA', 'RESULTADO'):
            try:
                v = lo.ListColumns(col).DataBodyRange.Cells(rr, 1).Validation
                fv[col] = {'Formula1': v.Formula1, 'Type': v.Type, 'IgnoreBlank': v.IgnoreBlank}
            except Exception as e:
                fv[col] = str(e)[:80]
        L['validacao_regras'] = fv

        # Painel WBC 619611 "Tudo"; Estatistica 2026 (ja recalculada acima)
        C.por(qa, p, 'C3', 'WBC')
        C.por(qa, p, 'H3', '619611')
        C.por(qa, p, 'K3', 2026, local=False)
        C.por(qa, p, 'K4', 'Tudo (sem filtro)')
        qa.run('mEstatistica.AtualizarEstatisticaAba')
        e5 = C.estat_tabela(qa)
        wk = ('WBC|1', 'WBC|2', 'WBC|3')
        A26 = dt.datetime(2026, 1, 1), dt.datetime(2026, 12, 31, 23, 59)
        L['G1_G2_antes'] = {'painel': C.painel(qa, NLV), 'estat': {k: e5[k] for k in wk},
                            'ind_exato': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin0, 'WBC', '619611', n, 'XN1000')])
                                          for n in (1, 2, 3)},
                            'eng': C.eng_runs(qa, {run1, run2}, NLV)}

        ids = {}
        f = lambda d: d.strftime('%d/%m/%Y')  # noqa: E731
        base = [('LOTE', '619611'), ('ANALITO', 'WBC'), ('MOTIVO', C.MOTIVO)]
        r, i, _ = C.escrever_manual(qa, [('EQUIPAMENTO', 'xn1000'), ('DATA', f(d1)), ('HORA', '23:36'), ('LOTE', '619611'),
                                         ('NIVEL', '3'), ('ANALITO', 'WBC'), ('RESULTADO', '17,433'), ('MOTIVO', C.MOTIVO)])
        ids['G1_equip_minusc'] = i
        r, i, _ = C.escrever_manual(qa, [('EQUIPAMENTO', 'XN1000'), ('DATA', f(d2)), ('HORA', '23:36'), ('LOTE', '619611'),
                                         ('NIVEL', '3'), ('ANALITO', 'wbc'), ('RESULTADO', '17,433'), ('MOTIVO', C.MOTIVO)])
        ids['G1_analito_minusc'] = i
        r, i, _ = C.escrever_manual(qa, [('EQUIPAMENTO', 'XN1000'), ('DATA', f(dneg)), ('HORA', '23:36'), ('LOTE', '619611'),
                                         ('NIVEL', '1'), ('ANALITO', 'WBC'), ('RESULTADO', '-1'), ('MOTIVO', C.MOTIVO)])
        ids['G2_menos1_N1'] = i
        r, i, _ = C.escrever_manual(qa, [('EQUIPAMENTO', 'XN1000'), ('DATA', f(dzero)), ('HORA', '23:36'), ('LOTE', '619611'),
                                         ('NIVEL', '2'), ('ANALITO', 'WBC'), ('RESULTADO', '0'), ('MOTIVO', C.MOTIVO)])
        ids['G2_zero_N2'] = i
        r, i, lida = C.escrever_manual(qa, [('EQUIPAMENTO', 'XN1000'), ('DATA', '31/02/2026'), ('HORA', '08:00'),
                                            ('LOTE', '619611'), ('NIVEL', '1'), ('ANALITO', 'WBC'), ('RESULTADO', '3,111'),
                                            ('MOTIVO', C.MOTIVO)])
        ids['G3_data_31_02'] = i
        L['G3_digitado'] = {'DATA_na_celula': lida.get('DATA'), 'tipo': type(lida.get('DATA')).__name__}
        L['ids_manuais'] = ids
        t0 = time.time()
        res = str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))
        L['atualizar'] = {'retorno': res[:600], 's': round(time.time() - t0, 1)}

        fin1 = C.final(qa)
        porid = {r['ID_REGISTRO']: r for r in fin1}
        man = {r['ID_REGISTRO']: r for r in C.manuais(qa)}
        campos = ('ID_REGISTRO', 'EQUIPAMENTO', 'ANALITO', 'ANALITO_CADASTRADO', 'DATA_HORA', 'RUN', 'CORRIDA_NO_DIA',
                  'NIVEL', 'LOTE', 'RESULTADO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA')
        L['manuais_final'] = {k: {c: porid.get(i, {}).get(c) for c in campos} for k, i in ids.items()}
        L['manuais_tela'] = {k: {c: man.get(i, {}).get(c) for c in ('STATUS', 'RUN', 'DETALHE')} for k, i in ids.items()}
        L['qa_codigos_dos_ids'] = C.qa_integracao(qa, list(ids.values()))
        e6 = C.estat_tabela(qa)
        mx = {n: C.pop(fin1, 'WBC', '619611', n, 'XN1000', maiusc=True) for n in (1, 2, 3)}
        L['G1_G2_depois'] = {
            'painel': C.painel(qa, NLV), 'estat': {k: e6[k] for k in wk},
            'estat_B3_D3_B4': (est.Range('B3').Value, est.Range('D3').Value, est.Range('B4').Value),
            'ind_exato': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'WBC', '619611', n, 'XN1000')]) for n in (1, 2, 3)},
            'ind_maiusc': {n: C.mdcv([r['RESULTADO'] for r in mx[n]]) for n in (1, 2, 3)},
            'ind_um_por_run': {n: C.mdcv(C.um_por_run(mx[n])) for n in (1, 2, 3)},
            'ind_2026_maiusc': {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'WBC', '619611', n, 'XN1000', *A26, maiusc=True)])
                                for n in (1, 2, 3)},
            'eng': C.eng_runs(qa, {run1, run2}, NLV)}

        # ---------------------------------------------------------------- G9: ET do Painel sem CEQ (RET%)
        C.por(qa, p, 'C3', 'RET%')
        pn = C.painel(qa, NLV)
        L['G9_RET%'] = {'painel': pn, 'H_esperado_1.65xCV': [1.65 * x['cv'] if isinstance(x['cv'], float) else None
                                                             for x in pn['niveis']],
                        'estat': {k: v for k, v in C.estat_tabela(qa).items() if k.startswith('RET%|')}}

        # ---------------------------------------------------------------- CEQ (RBC; LYMPH%)
        C.por(qa, p, 'C3', 'RBC')
        cap = qa.ws('EQA.CAP_Dados')
        k = ('RBC|1', 'RBC|2', 'RBC|3')
        kl = ('LYMPH%|1', 'LYMPH%|2', 'LYMPH%|3', 'LYMPH#|1')
        b0 = C.estat_tabela(qa)
        ceq = {'B0': {x: b0[x] for x in k + kl}}
        t0 = time.time()
        r1 = str(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Red Blood Cell Count',
                                                           'TA-01', '110', unid='x 10E12/L')))
        lin = int(r1.split('|')[1]) if r1.startswith('OK|') else None
        b1 = C.estat_tabela(qa)
        ceq['C1_incluir_110'] = {'ret': r1, 's': round(time.time() - t0, 1), 'estat': {x: b1[x] for x in k},
                                 'eqa_base': C.eqa_teste(qa)}
        r2 = str(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Red Blood Cell Count',
                                                           'TA-01', '105', unid='x 10E12/L')))
        ceq['C2_janela_mesma_chave_105'] = {'ret': r2, 'ProtectContents': cap.ProtectContents,
                                            'A_R_locked': [cap.Cells(lin, c).Locked for c in range(1, 19)] if lin else None}
        try:
            cap.Range(f'F{lin}').Value = 105
            via = 'COM direto (UserInterfaceOnly)'
        except Exception as e:
            cap.Unprotect(senha)
            cap.Range(f'F{lin}').Value = 105
            via = f'ADM desprotegeu ({e.__class__.__name__})'
        qa.ex.esperar()
        b2 = C.estat_tabela(qa)
        ceq['C3_editar_celula_105'] = {'via': via, 'O_bias_tabela': cap.Range(f'O{lin}').Value, 'eqa_base': C.eqa_teste(qa),
                                       'estat_linhas_mudaram': [x for x in b1 if any(b1[x][c] != b2[x][c] for c in b1[x] if c != 'linha')],
                                       'Z2': cap.Range('Z2').Value, 'Z3': cap.Range('Z3').Value}
        t0 = time.time()
        qa.run('mEQA.AtualizarEQABase')
        b3 = C.estat_tabela(qa)
        ceq['C4_consolidar'] = {'s': round(time.time() - t0, 1), 'estat': {x: b3[x] for x in k}, 'eqa_base': C.eqa_teste(qa),
                                'Z2': cap.Range('Z2').Value, 'Z3': cap.Range('Z3').Value,
                                'mudaram': [x for x in b2 if any(b2[x][c] != b3[x][c] for c in b2[x] if c != 'linha')]}
        g7 = {}
        for nome, campos in (
                ('fora_do_mapa', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Red Blood Cell', 'TA-02', '120', unid='x 10E12/L')),
                ('controllab', C.campos_ceq('Controllab', 'TESTE-AUDITORIA CTL1', '2025', 'RBC', 'TA-C1', '130', unid='x 10E12/L')),
                ('G8_lymphs_sem_absolute', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Lymphs', 'TA-05', '110', unid='x 10E9/L'))):
            antes = C.estat_tabela(qa)
            r = str(qa.run('mCEQForm.SalvarCEQ', campos))
            depois = C.estat_tabela(qa)
            g7[nome] = {'ret': r, 'MapearAnalito': qa.run('mEQA.MapearAnalito', campos[0], campos[3]),
                        'estat_linhas_mudaram': [x for x in antes if any(antes[x][c] != depois[x][c]
                                                                         for c in antes[x] if c != 'linha')],
                        'antes': {x: antes[x] for x in kl}, 'depois': {x: depois[x] for x in kl}}
        g7['eqa_base_teste'] = C.eqa_teste(qa)
        ceq['C5_G7_G8'] = g7
        # provedor global -> Controllab
        prov = celula_rotulo(est, 4, 'Provedor')
        antes = C.estat_tabela(qa)
        pn0 = C.painel(qa, NLV)
        C.por(qa, est, prov, 'Controllab', local=False)
        depois = C.estat_tabela(qa)
        pn = C.painel(qa, NLV)
        ceq['C6_provedor_controllab'] = {
            'celula': prov, 'sem_ep_antes': sum(1 for v in antes.values() if v['G_bias'] == 'SEM EP'),
            'sem_ep_depois': sum(1 for v in depois.values() if v['G_bias'] == 'SEM EP'), 'linhas': len(depois),
            'H_vazios_depois': sum(1 for v in depois.values() if v['H_et'] in (None, '')),
            'painel_antes': pn0, 'painel_depois': pn,
            'H_esperado_1.65xCV': [1.65 * x['cv'] for x in pn['niveis']], 'K5': est.Range('K5').Value}
        C.por(qa, est, prov, 'CAP', local=False)
        L['G5_G6_G7_G8_ceq'] = ceq

        # ---------------------------------------------------------------- G14 ampliacao: todos os lotes
        soma = {}
        for r in fin0:
            if r['PARTICIPA_ESTATISTICA'] == 'SIM' and isinstance(r['RESULTADO'], float):
                soma.setdefault((str(r['LOTE']), r['ANALITO'], int(r['NIVEL'])), []).append(r['RESULTADO'])
        pares, sinal = 0, []
        for kk, v in soma.items():
            m, dp = qa.run('mLotes.MediaDoLote', kk[0], kk[1], kk[2]), qa.run('mLotes.DPDoLote', kk[0], kk[1], kk[2])
            if isinstance(m, float) and isinstance(dp, float) and dp > 0 and len(v) >= 2:
                pares += 1
                z = (sum(v) / len(v) - m) / dp
                if abs(z) > 5:
                    sinal.append({'lote': kk[0], 'analito': kk[1], 'nivel': kk[2], 'n': len(v), 'media': sum(v) / len(v),
                                  'alvo': m, 'dp': dp, 'z': z,
                                  'teste_adr076': qa.run('mLotes.AlvoDeTeste', kk[0], kk[1], kk[2])})
        sinal.sort(key=lambda x: (x['lote'], -abs(x['z'])))
        L['G14_alvos_todos_lotes'] = {'pares_avaliados': pares, 'lotes': sorted({kk[0] for kk in soma}),
                                      'z_gt5': sinal}
    except Exception:
        L['ERRO'] = traceback.format_exc()
    finally:
        qa.fechar()
        L['sha_copia_depois'] = C.sha(copia)
        L['sha_prod_depois'] = C.sha(src)
        L['fim'] = dt.datetime.now().isoformat(timespec='seconds')


if __name__ == '__main__':
    main()
