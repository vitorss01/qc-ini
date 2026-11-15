# -*- coding: utf-8 -*-
"""valida_r5_lote_sint.py -- Tester/QA, ciclo 2026-11-15_reteste05: ramo "Período inteiro do lote" do aviso do corte de
180 corridas na Hematologia (OC-0012). Nenhum lote real tem mais de 180 corridas (o maior, 528211, tem 89), entao o
ramo so aparece com dado SINTETICO, so na copia:
  * na tblCQ_Final da copia (eventos desligados), todas as linhas de WBC de dois lotes reais escolhidos pelo numero de
    corridas recebem o LOTE do lote de destino (rotulo sintetico; RESULTADO, RUN, datas e status intactos);
  * mEstatistica.InvalidarCache; Painel C3 = WBC, H3 = lote de destino, K4 "Tudo" e depois 2026 "Ano inteiro"
    (eventos ligados, como o usuario);
  * oraculo proprio (valida_r5_excel.oraculo_painel) sobre a tblCQ_Final JA modificada.
Lotes escolhidos de forma diferente do Usuario Simulado (que usou 533811, 614011 e 619611 -> 528211).

Uso: python -u valida_r5_lote_sint.py
Regras: so COPIA nova em %TEMP%/qcw/auditoria/2026-11-15_reteste05/tester/; nada e salvo; so o PID do QA e encerrado.
"""
import collections
import json
import os
import shutil
import sys
import time
import traceback

sys.argv = [sys.argv[0], 'Hematologia', 'nova']
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import valida_r5_excel as V  # noqa: E402

C = V.C


def main():
    src = os.path.join(C.REPO, 'QC_Hematologia.xlsm')
    h0 = C.sha(src)
    copia = os.path.join(C.PASTA, f'Hematologia_lote_sint_{time.strftime("%H%M%S")}.xlsm')
    shutil.copy2(src, copia)
    L = V.Log(os.path.join(C.PASTA, 'valida_Hematologia_lote_sint.json'))
    L['ambiente'] = {'origem': 'QC_Hematologia.xlsm', 'sha_origem_12': h0, 'copia': os.path.basename(copia)}
    qa = C.qa_final.QA('Hematologia', copia)
    L['excel_pid'] = qa.ex.pid
    try:
        fin = V.Final(qa)
        corr = collections.defaultdict(set)
        for r in fin.rows:
            if V.up(r['ANALITO']) == 'WBC' and V.up(r['PARTICIPA_ESTATISTICA']) == 'SIM':
                corr[V.t(r['LOTE'])].add(V.t(r['RUN']))
        ordem = sorted(corr, key=lambda k: -len(corr[k]))
        destino = '528211' if '528211' in corr else ordem[0]
        fontes = [k for k in ordem if k != destino][:2]
        L['escolha'] = {'destino': [destino, len(corr[destino])], 'fontes': [[k, len(corr[k])] for k in fontes],
                        'corridas_por_lote_top6': [[k, len(corr[k])] for k in ordem[:6]]}
        lo = qa.lo('tblCQ_Final')
        ws = lo.Parent
        try:
            ws.Unprotect(C.senha_planilha())
        except Exception:  # noqa: BLE001
            pass
        qa.ex.xl.EnableEvents = False
        col_l = lo.ListColumns('LOTE').DataBodyRange
        col_a = lo.ListColumns('ANALITO').DataBodyRange.Value2
        vl = [list(x) for x in col_l.Value2]
        n = 0
        for i, (a,) in enumerate(col_a):
            if V.up(a) == 'WBC' and V.t(vl[i][0]) in fontes:
                vl[i][0] = destino
                n += 1
        fmt = col_l.NumberFormat
        col_l.NumberFormat = '@'
        col_l.Value2 = vl
        L['linhas_relabeladas'] = {'n': n, 'formato_coluna_antes': V.t(fmt)}
        qa.ex.xl.EnableEvents = True
        qa.run('mEstatistica.InvalidarCache', teto=120)
        fin2 = V.Final(qa)
        p = qa.ws('Painel')
        p.Activate()
        out = {}
        for nome, passos in [('L1_lote_Tudo', [('C3', 'WBC'), ('H3', "'" + destino), ('K4', 'Tudo (sem filtro)')]),
                             ('L2_lote_2026', [('K3', 2026), ('K4', 'Ano inteiro')]),
                             ('L3_todos_2026', [('H3', None)])]:
            V.aplicar(qa, p, passos)
            c = V.ler_cenario(qa, fin2)
            out[nome] = {k: c[k] for k in ('analito', 'lote', 'todos', 'M1', 'AF1', 'engPainel', 'confere')}
            out[nome]['oraculo'] = {k: c['oraculo'][k] for k in ('cortadas', 'pular_X', 'M1_esperado', 'AF1',
                                                                    'corridas_periodo')}
            a39 = c['A39']
            i = a39.find('⚠')
            out[nome]['A39_aviso'] = a39[i:] if i >= 0 else a39[-200:]
            out[nome]['tem_Estatistica'] = 'Estatística' in a39
        L['cenarios'] = out
    except Exception as e:  # noqa: BLE001
        L['ERRO_FATAL'] = {'erro': repr(e), 'trace': traceback.format_exc()[-3000:]}
    finally:
        try:
            qa.fechar(salvar=False)
        except Exception:  # noqa: BLE001
            pass
    L['origem_intacta'] = {'sha_depois': C.sha(src), 'igual': h0 == C.sha(src)}
    try:
        os.remove(copia)
        L['copia_apagada'] = True
    except Exception as e:  # noqa: BLE001
        L['copia_apagada'] = repr(e)
    print('FIM', flush=True)


if __name__ == '__main__':
    main()
