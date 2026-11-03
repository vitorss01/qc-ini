# -*- coding: utf-8 -*-
"""repro_excel_bio.py -- Tester/QA, ciclo 2026-10-31_ciclo01, modo REPRODUZIR, BIOQUIMICA (uma sessao de Excel).

Grupos cobertos (ids do tester_reproduzir.md):
  G1  minusculas no EQUIPAMENTO ("dimension 1") e no ANALITO ("glicose") do manual -> Painel x Estatistica
  G3  data invalida colada (31/02/2026)
  G4  EQUIPAMENTO vazio recebe o equipamento do Painel (L4 = DIMENSION 2)
  G2b validacao da tela aceita RESULTADO -1 (Validation.Value)
  G11 "Tudo (sem filtro)" num lote com mais de 180 corridas (Glicose DIM1 8974)
  G9  ET do Painel com bias SEM EP (HbA1c 8590, PCR 8353; Proteina total com M4 = Controllab)
  G5  correcao do CEQ pela janela (SalvarCEQ mesma chave) e protecao da aba
  G6  correcao na tabela so chega ao bias depois do Consolidar; Z2/Z3
  G7  CEQ com nome fora do mapa / Controllab / analito novo (HbA1c) gravado OK e fora do bias
  G12 ampliacao: alvo (MediaDoLote/DPDoLote) dos lotes que nao sao do analito

Caminho do usuario: Digitar Resultados (celula a celula, FormulaLocal, eventos ligados) -> mIntegracao.AtualizarSoEntradas
(o nucleo do ATUALIZAR sem ler o DB_SEAC: aplica os manuais e roda o motor inteiro; ADR-081). Painel pela propria
celula (aba desprotegida, passo tecnico). CEQ por mCEQForm.SalvarCEQ (nucleo da janela frmCEQ).
Dados de teste: MOTIVO 'Outro: TESTE-AUDITORIA'; valores 77,111 / 244,222; CEQ rodada 'TESTE-AUDITORIA R1', amostras TA-*.

Uso: python -u repro_excel_bio.py   (copia a producao para %TEMP%/qcw/auditoria/2026-10-31_ciclo01/tester/)
"""
import datetime as dt
import os
import sys
import time
import traceback
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repro_comum as C  # noqa: E402

NLV = 2
PARTES = set((os.environ.get('REPRO_PARTES') or 'manual,ceq').split(','))
OUT = dt.datetime(2026, 10, 1), dt.datetime(2026, 10, 31)


def main():
    src, copia, h0 = C.copiar('Bioquimica')
    L = C.Log(os.path.join(C.PASTA, 'repro_excel_bio' + ('' if PARTES == {'manual', 'ceq'} else '_' + '_'.join(sorted(PARTES))) + '.json'))
    L['ambiente'] = {'producao': os.path.basename(src), 'sha_prod': h0, 'copia': copia, 'sha_copia_antes': C.sha(copia)}
    senha = C.senha_planilha()
    qa = C.qa_final.QA('Bioquimica', copia)
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

        if 'manual' in PARTES:
            # ---------------------------------------------------------------- escolher os dias (dados da producao)
            g = [r for r in fin0 if r['ANALITO'] == 'Glicose' and str(r['LOTE']) == '8977' and r['EQUIPAMENTO'] ==
                 'DIMENSION 1' and r['NIVEL'] == 2 and r['PARTICIPA_ESTATISTICA'] == 'SIM' and
                 r['ORIGEM_RESULTADO'] == 'INTERFACEAMENTO' and OUT[0] <= r['DATA'] <= OUT[1]]
            dias = sorted({r['DATA'] for r in g if str(int(r['RUN'] or 0)).endswith('01')}, reverse=True)
            d1, d2 = dias[0], dias[1]
            real1 = [r for r in g if r['DATA'] == d1 and str(int(r['RUN'] or 0)).endswith('01')]
            real2 = [r for r in g if r['DATA'] == d2 and str(int(r['RUN'] or 0)).endswith('01')]
            L['dias_escolhidos'] = {'d1': d1, 'real_run01_d1': [(r['ID_REGISTRO'], r['RUN'], r['RESULTADO']) for r in real1],
                                    'd2': d2, 'real_run01_d2': [(r['ID_REGISTRO'], r['RUN'], r['RESULTADO']) for r in real2]}
            run1, run2 = int(real1[0]['RUN']), int(real2[0]['RUN'])

            # ---------------------------------------------------------------- validacao da tela (o que passa digitado)
            L['validacao'] = {f'{c}={t}': C.validacao(qa, c, t) for c, t in
                              (('EQUIPAMENTO', 'dimension 1'), ('EQUIPAMENTO', 'DIMENSION 1'), ('ANALITO', 'glicose'),
                               ('ANALITO', 'Glicose'), ('DATA', '31/02/2026'), ('RESULTADO', '-1'), ('RESULTADO', '0'))}

            # ---------------------------------------------------------------- Painel: Glicose 8977 DIM1 Out/2026
            C.por(qa, p, 'C3', 'Glicose')
            C.por(qa, p, 'L4', 'DIMENSION 1')
            C.por(qa, p, 'H3', '8977')
            C.por(qa, p, 'K3', 2026, local=False)
            C.por(qa, p, 'K4', 'Out')
            e0 = C.estat_tabela(qa)
            L['G1_antes'] = {'painel': C.painel(qa, NLV), 'estat': {k: e0[k] for k in ('Glicose|1', 'Glicose|2')},
                             'ind_exato_N2': C.mdcv([r['RESULTADO'] for r in C.pop(fin0, 'Glicose', '8977', 2,
                                                                                     'DIMENSION 1', *OUT)]),
                             'eng': C.eng_runs(qa, {run1, run2}, NLV)}

            # ---------------------------------------------------------------- manuais
            ids = {}
            # G4: EQUIPAMENTO vazio, Painel em DIMENSION 2
            C.por(qa, p, 'L4', 'DIMENSION 2')
            r, i, lida = C.escrever_manual(qa, [('DATA', d2.strftime('%d/%m/%Y')), ('HORA', '23:40'), ('LOTE', '8977'),
                                                ('NIVEL', '1'), ('ANALITO', 'Glicose'), ('RESULTADO', '77,111'),
                                                ('MOTIVO', C.MOTIVO)])
            ids['G4_equip_vazio'] = i
            L['G4_digitado'] = {'id': i, 'EQUIPAMENTO_antes_de_atualizar': lida.get('EQUIPAMENTO'),
                                'cfg_EQUIPAMENTO_PADRAO': qa.cfg('EQUIPAMENTO_PADRAO')}
            C.por(qa, p, 'L4', 'DIMENSION 1')
            # G1: "dimension 1" (N2, d1 23:36) e "glicose" (N2, d2 23:36)
            r, i, lida = C.escrever_manual(qa, [('EQUIPAMENTO', 'dimension 1'), ('DATA', d1.strftime('%d/%m/%Y')),
                                                ('HORA', '23:36'), ('LOTE', '8977'), ('NIVEL', '2'),
                                                ('ANALITO', 'Glicose'), ('RESULTADO', '244,222'), ('MOTIVO', C.MOTIVO)])
            ids['G1_equip_minusc'] = i
            r, i, lida = C.escrever_manual(qa, [('EQUIPAMENTO', 'DIMENSION 1'), ('DATA', d2.strftime('%d/%m/%Y')),
                                                ('HORA', '23:36'), ('LOTE', '8977'), ('NIVEL', '2'),
                                                ('ANALITO', 'glicose'), ('RESULTADO', '244,222'), ('MOTIVO', C.MOTIVO)])
            ids['G1_analito_minusc'] = i
            # G3: data invalida colada
            r, i, lida = C.escrever_manual(qa, [('EQUIPAMENTO', 'DIMENSION 1'), ('DATA', '31/02/2026'), ('HORA', '08:00'),
                                                ('LOTE', '8977'), ('NIVEL', '1'), ('ANALITO', 'Glicose'),
                                                ('RESULTADO', '77,111'), ('MOTIVO', C.MOTIVO)])
            ids['G3_data_31_02'] = i
            L['G3_digitado'] = {'id': i, 'DATA_na_celula': lida.get('DATA'), 'tipo': type(lida.get('DATA')).__name__}
            L['ids_manuais'] = ids

            t0 = time.time()
            res = str(qa.run('mIntegracao.AtualizarSoEntradas', '', teto=1500))
            L['atualizar'] = {'retorno': res[:600], 's': round(time.time() - t0, 1)}

            fin1 = C.final(qa)
            man = {r['ID_REGISTRO']: r for r in C.manuais(qa)}
            porid = {r['ID_REGISTRO']: r for r in fin1}
            campos = ('ID_REGISTRO', 'EQUIPAMENTO', 'ANALITO', 'ANALITO_CADASTRADO', 'DATA_HORA', 'RUN', 'CORRIDA_NO_DIA',
                      'NIVEL', 'LOTE', 'RESULTADO', 'STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA')
            L['manuais_final'] = {k: {c: porid.get(i, {}).get(c) for c in campos} for k, i in ids.items()}
            L['manuais_tela'] = {k: {c: man.get(i, {}).get(c) for c in ('STATUS', 'RUN', 'DETALHE')} for k, i in ids.items()}
            L['qa_codigos_dos_ids'] = C.qa_integracao(qa, list(ids.values()))

            e1 = C.estat_tabela(qa)
            pn = C.painel(qa, NLV)
            ex = C.pop(fin1, 'Glicose', '8977', 2, 'DIMENSION 1', *OUT)
            mx = C.pop(fin1, 'Glicose', '8977', 2, 'DIMENSION 1', *OUT, maiusc=True)
            L['G1_depois'] = {
                'painel': pn, 'estat': {k: e1[k] for k in ('Glicose|1', 'Glicose|2')},
                'ind_exato_N2': C.mdcv([r['RESULTADO'] for r in ex]),
                'ind_maiusc_N2': C.mdcv([r['RESULTADO'] for r in mx]),
                'ind_um_por_run_N2': C.mdcv(C.um_por_run(mx)),
                'eng': C.eng_runs(qa, {run1, run2}, NLV),
                'ET_sigma_linha_estat_Glicose2_antes_depois': (e0['Glicose|2']['H_et'], e1['Glicose|2']['H_et'])}
            # G4: em que serie o manual entrou
            i4 = ids['G4_equip_vazio']
            L['G4_resultado'] = {c: porid.get(i4, {}).get(c) for c in campos}

            # ---------------------------------------------------------------- G11: Tudo (sem filtro) no 8974
            C.por(qa, p, 'H3', '8974')
            C.por(qa, p, 'K4', 'Tudo (sem filtro)')
            pn = C.painel(qa, NLV)
            tudo = {n: C.mdcv([r['RESULTADO'] for r in C.pop(fin1, 'Glicose', '8974', n, 'DIMENSION 1')]) for n in (1, 2)}
            runs = {n: len({r['RUN'] for r in C.pop(fin1, 'Glicose', '8974', n, 'DIMENSION 1')}) for n in (1, 2)}
            eng = qa.ws('Eng_Saida')
            L['G11_tudo_8974'] = {'painel': pn, 'independente_todos_os_dados': tudo, 'corridas_distintas': runs,
                                  'eng_n_slots': eng.Range('I1').Value}

            # ---------------------------------------------------------------- G9: ET do Painel sem CEQ
            g9 = {}
            for an, lote, eq in (('Hemoglobina glicada', '8590', 'DIMENSION 2'), ('Proteína C-reativa', '8353', 'DIMENSION 2')):
                C.por(qa, p, 'C3', an)
                C.por(qa, p, 'L4', eq)
                C.por(qa, p, 'H3', lote)
                C.por(qa, p, 'K3', 2026, local=False)
                C.por(qa, p, 'K4', 'Out')
                pn = C.painel(qa, NLV)
                et = C.estat_tabela(qa)
                g9[an] = {'painel': pn, 'estat': {f'{an}|{n}': et.get(f'{an}|{n}') for n in (1, 2)},
                          'H_esperado_1.65xCV': [1.65 * x['cv'] if isinstance(x['cv'], float) else None for x in pn['niveis']],
                          'bloco_ET_C61_C63': C.bloco(p, 'A60:D64')}
            C.por(qa, p, 'C3', 'Proteína total')
            C.por(qa, p, 'L4', 'DIMENSION 2')
            C.por(qa, p, 'H3', '8977')
            C.por(qa, p, 'K4', 'Out')
            antes = C.painel(qa, NLV)
            C.por(qa, p, 'M4', 'Controllab')
            pn = C.painel(qa, NLV)
            g9['Proteína total M4=Controllab'] = {'antes': antes, 'depois': pn,
                                                  'H_esperado_1.65xCV': [1.65 * x['cv'] for x in pn['niveis']]}
            C.por(qa, p, 'M4', 'CAP')
            g9['Proteína total M4=CAP (restaurado)'] = C.painel(qa, NLV)
            L['G9_et_painel'] = g9

        if 'ceq' in PARTES:
            # ---------------------------------------------------------------- G5/G6/G7: CEQ (Proteina total)
            mapa = C.bloco(qa.ws('EQA_Base'), 'W2:Y201')
            L['mapa_proteina'] = [m for m in mapa if m[1] and ('rotein' in str(m[1]) or 'rote' in str(m[2] or '')
                                                             or 'A1c' in str(m[1]) or 'glicada' in str(m[2] or ''))]
            cap = qa.ws('EQA.CAP_Dados')
            k = 'Proteína total|1', 'Proteína total|2'
            b0 = C.estat_tabela(qa)
            ceq = {'B0': {x: b0[x] for x in k}}
            t0 = time.time()
            r1 = str(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Protein, total, serum',
                                                               'TA-01', '110', unid='g/dL')))
            b1 = C.estat_tabela(qa)
            lin = int(r1.split('|')[1]) if r1.startswith('OK|') else None
            ceq['C1_incluir_110'] = {'ret': r1, 's': round(time.time() - t0, 1), 'estat': {x: b1[x] for x in k},
                                     'eqa_base': C.eqa_teste(qa)}
            r2 = str(qa.run('mCEQForm.SalvarCEQ', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Protein, total, serum',
                                                               'TA-01', '105', unid='g/dL')))
            ceq['C2_janela_mesma_chave_105'] = {'ret': r2, 'ProtectContents': cap.ProtectContents,
                                                'UserInterfaceOnly': None,
                                                'F_locked': cap.Range(f'F{lin}').Locked if lin else None,
                                                'A_R_locked': [cap.Cells(lin, c).Locked for c in range(1, 19)] if lin else None}
            try:
                ceq['C2_janela_mesma_chave_105']['UserInterfaceOnly'] = cap.ProtectionMode
            except Exception:
                pass
            # G6: editar a celula (como a mensagem manda) -- escrita por COM passa o UserInterfaceOnly; se nao passar,
            # o ADM desprotege (registrado)
            try:
                cap.Range(f'F{lin}').Value = 105
                via = 'COM direto (UserInterfaceOnly)'
            except Exception as e:
                cap.Unprotect(senha)
                cap.Range(f'F{lin}').Value = 105
                via = f'ADM desprotegeu ({e.__class__.__name__})'
            qa.ex.esperar()
            b2 = C.estat_tabela(qa)
            mud = [x for x in b1 if any(b1[x][c] != b2[x][c] for c in b1[x] if c != 'linha')]
            ceq['C3_editar_celula_105'] = {'via': via, 'O_bias_tabela': cap.Range(f'O{lin}').Value,
                                           'eqa_base': C.eqa_teste(qa), 'estat_linhas_mudaram': mud,
                                           'Z2': cap.Range('Z2').Value, 'Z3': cap.Range('Z3').Value}
            t0 = time.time()
            qa.run('mEQA.AtualizarEQABase')
            b3 = C.estat_tabela(qa)
            ceq['C4_consolidar'] = {'s': round(time.time() - t0, 1), 'estat': {x: b3[x] for x in k},
                                    'eqa_base': C.eqa_teste(qa), 'Z2': cap.Range('Z2').Value, 'Z3': cap.Range('Z3').Value,
                                    'mudaram': [x for x in b2 if any(b2[x][c] != b3[x][c] for c in b2[x] if c != 'linha')]}
            # G7
            g7 = {}
            for nome, campos in (
                    ('fora_do_mapa', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Protein, total', 'TA-02', '120', unid='g/dL')),
                    ('controllab', C.campos_ceq('Controllab', 'TESTE-AUDITORIA CTL1', '2025', 'Proteína total', 'TA-C1', '130', unid='g/dL')),
                    ('hba1c_nome_CAP', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Hemoglobin A1c', 'TA-03', '6,5', alvo='6', unid='%')),
                    ('hba1c_nome_canonico', C.campos_ceq('CAP', 'TESTE-AUDITORIA R1', '2025', 'Hemoglobina glicada', 'TA-04', '6,5', alvo='6', unid='%'))):
                antes = C.estat_tabela(qa)
                r = str(qa.run('mCEQForm.SalvarCEQ', campos))
                depois = C.estat_tabela(qa)
                g7[nome] = {'ret': r, 'MapearAnalito': qa.run('mEQA.MapearAnalito', campos[0], campos[3]),
                            'estat_linhas_mudaram': [x for x in antes if any(antes[x][c] != depois[x][c]
                                                                             for c in antes[x] if c not in ('linha',))],
                            'hba1c_G': (depois['Hemoglobina glicada|1']['G_bias'], depois['Hemoglobina glicada|2']['G_bias'])}
            g7['eqa_base_teste'] = C.eqa_teste(qa)
            ceq['C5_G7'] = g7
            L['G5_G6_G7_ceq'] = ceq

            # ---------------------------------------------------------------- G12 ampliacao: alvo dos lotes estranhos
            g12 = {}
            fing12 = C.final(qa)
            for lote, an in (('5659', 'Glicose'), ('6353', 'Ureia'), ('6354', 'Creatinina'), ('8353', 'Glicose'),
                             ('8353', 'Ureia'), ('8974', 'Hemoglobina glicada'), ('8977', 'Glicose'), ('8974', 'Ureia')):
                g12[f'{lote}|{an}'] = {n: (qa.run('mLotes.MediaDoLote', lote, an, n), qa.run('mLotes.DPDoLote', lote, an, n))
                                       for n in (1, 2)}
            L['G12_alvos'] = g12
            # varredura: para cada lote x analito com >= 1 resultado ativo, alvo cadastrado x media medida
            grupos = Counter((str(r['LOTE']), r['ANALITO'], int(r['NIVEL'])) for r in fing12
                             if r['PARTICIPA_ESTATISTICA'] == 'SIM' and r['ANALITO_CADASTRADO'] == 'SIM')
            soma = {}
            for r in fing12:
                if r['PARTICIPA_ESTATISTICA'] == 'SIM' and r['ANALITO_CADASTRADO'] == 'SIM' and isinstance(r['RESULTADO'], float):
                    kk = (str(r['LOTE']), r['ANALITO'], int(r['NIVEL']))
                    soma.setdefault(kk, []).append(r['RESULTADO'])
            estranhos = []
            for kk, v in soma.items():
                if len(v) < 1:
                    continue
                m, dp = qa.run('mLotes.MediaDoLote', kk[0], kk[1], kk[2]), qa.run('mLotes.DPDoLote', kk[0], kk[1], kk[2])
                if isinstance(m, float) and isinstance(dp, float) and dp > 0:
                    z = (sum(v) / len(v) - m) / dp
                    if abs(z) > 5 or (len(v) <= 2):
                        estranhos.append({'lote': kk[0], 'analito': kk[1], 'nivel': kk[2], 'n': len(v),
                                          'media': sum(v) / len(v), 'alvo': m, 'dp': dp, 'z': z})
            estranhos.sort(key=lambda x: -abs(x['z']))
            L['G12_varredura_lote_analito'] = {'grupos': len(soma), 'z_gt5_ou_n_le2': estranhos[:80],
                                               'total_sinalizados': len(estranhos)}
    except Exception:
        L['ERRO'] = traceback.format_exc()
    finally:
        qa.fechar()
        L['sha_copia_depois'] = C.sha(copia)
        L['sha_prod_depois'] = C.sha(src)
        L['fim'] = dt.datetime.now().isoformat(timespec='seconds')


if __name__ == '__main__':
    main()
