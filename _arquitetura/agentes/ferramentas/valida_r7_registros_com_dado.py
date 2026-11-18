# -*- coding: utf-8 -*-
"""valida_r7_registros_com_dado.py -- Tester/QA, ciclo 2026-11-16_reteste07, VALIDAR (OC-0082, complemento).

O valida_r7_adr097.py provou a lista vazia ("(nenhum registrado)") e o texto livre recusado. Nas producoes, porem,
nenhum registro da aba Registros tem EQUIPAMENTO nem STATUS, entao o item "com dados a lista e a de antes" nao foi
exercitado para G4/J4. Aqui, numa COPIA nova: grava equipamento e status sinteticos (TESTE-AUDITORIA) nas linhas da
tblRegistros, reconstroi as listas (mCalibracao.AtualizarListasFiltro) e confere:
  - lista de G4/J4 = valores distintos da coluna (oraculo), sem "(nenhum registrado)";
  - texto fora da lista recusado; cada valor da lista aceito;
  - filtrar por um valor mostra exatamente as linhas com ele (AplicarFiltrosRegistros x oraculo);
  - apagar a coluna de novo volta a "(nenhum registrado)".
Uso: python -u valida_r7_registros_com_dado.py <Bioquimica|Hematologia>
Regras: so copia; nada e salvo; fecha so o PID criado pelo qa_final.QA; nenhuma captura de tela.
"""
import json
import os
import shutil
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import valida_r7_adr097 as V  # noqa: E402  (le PROD de sys.argv[1])
import qa_final  # noqa: E402

PROD = sys.argv[1]


def main():
    os.makedirs(V.PASTA, exist_ok=True)
    src = os.path.join(V.C.REPO, f'QC_{PROD}.xlsm')
    dst = os.path.join(V.PASTA, f'{PROD}_reg_{time.strftime("%H%M%S")}.xlsm')
    L = V.Log(os.path.join(V.PASTA, f'valida_r7_registros_{PROD}.json'))
    h0 = V.sha_full(src)
    shutil.copy2(src, dst)
    L['origem'] = {'sha256_antes': h0[:16], 'sha_copia': V.sha_full(dst)[:16]}
    qa = qa_final.QA(PROD, dst)
    L['excel_pid'] = qa.ex.pid
    try:
        xl = qa.ex.xl
        ws = qa.ws('Registros')
        lo = ws.ListObjects('tblRegistros')
        ws.Unprotect(V.C.senha_planilha())
        n = lo.ListRows.Count
        vals0 = lo.DataBodyRange.Value2 if n else []
        # registro = linha com Data, Analito, Observacao ou ID do resultado (o mesmo criterio do ContarRegistros do
        # mCalibracao); a linha de entrada vazia (o Nº e calculado) NAO e registro (1a execucao, 18:36: o oraculo a contava)
        usadas = [i for i in range(n) if any(V.t(vals0[i][c - 1]) for c in (2, 3, 7, 11))]
        L['linhas'] = {'total': n, 'com_dado': len(usadas),
                       'linhas': [[V.t(x) for k, x in enumerate(vals0[i]) if k not in (7, 8)][:13] for i in range(n)]}
        EQS = ['TESTE-AUDITORIA EQ-A', 'TESTE-AUDITORIA EQ-B']
        STS = ['TESTE-AUDITORIA ST-1', 'TESTE-AUDITORIA ST-2']
        xl.EnableEvents = False
        for k, i in enumerate(usadas):
            lo.ListColumns(12).DataBodyRange.Cells(i + 1, 1).Value = EQS[k % 2]
            lo.ListColumns(10).DataBodyRange.Cells(i + 1, 1).Value = STS[0] if k == 0 else STS[1]
        xl.EnableEvents = True
        qa.run('mCalibracao.LimparFiltrosRegistros', teto=120)
        qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
        vals = lo.DataBodyRange.Value2
        out = {}
        for rot, c, ci in (('G4 equipamento', 7, 12), ('J4 status', 10, 10)):
            cel = ws.Cells(4, c)
            f1, it = V.itens_lista(cel)
            orac = sorted({V.t(r[ci - 1]) for r in vals if V.t(r[ci - 1])}, key=lambda x: x.lower())
            r = {'itens': it, 'oraculo': orac, 'igual': it == orac,
                 'fora_recusado': {x: V.aceita(xl, cel, x) for x in ('(nenhum registrado)', 'QA-NAO-EXISTE')},
                 'itens_aceitos': {x: V.aceita(xl, cel, x) for x in it}}
            filt = {}
            for x in it:
                xl.EnableEvents = True
                cel.Value = x
                qa.ex.esperar()
                time.sleep(0.4)
                ret = V.t(qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120))
                o = sum(1 for rr in vals if V.t(rr[ci - 1]) == x and any(V.t(rr[c - 1]) for c in (2, 3, 7, 11)))
                filt[x] = {'retorno': ret, 'oraculo': o, 'K4': V.t(ws.Range('K4').Text),
                           'ok': ret.split('|')[0].isdigit() and int(ret.split('|')[0]) == o}
                xl.EnableEvents = False
                cel.ClearContents()
                xl.EnableEvents = True
            qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120)
            r['filtra'] = filt
            r['ok'] = (r['igual'] and '(nenhum registrado)' not in it and all(v is False for v in r['fora_recusado'].values())
                       and all(v is True for v in r['itens_aceitos'].values()) and all(v['ok'] for v in filt.values())
                       and len(it) >= 1)
            out[rot] = r
        L['com_dado'] = out
        # volta a vazio
        xl.EnableEvents = False
        lo.ListColumns(12).DataBodyRange.ClearContents()
        lo.ListColumns(10).DataBodyRange.ClearContents()
        xl.EnableEvents = True
        qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
        L['vazio_de_novo'] = {rot: V.itens_lista(ws.Cells(4, c))[1] for rot, c in (('G4 equipamento', 7), ('J4 status', 10))}
        L['ok'] = all(v['ok'] for v in out.values()) and all(v == ['(nenhum registrado)'] for v in L['vazio_de_novo'].values())
    finally:
        qa.fechar()
        h1 = V.sha_full(src)
        L['fim'] = {'sha256_depois': h1[:16], 'producao_intacta': h1 == h0}
        try:
            os.remove(dst)
        except Exception:                                            # noqa: BLE001
            pass
    print(json.dumps({'ok': L.get('ok')}))


if __name__ == '__main__':
    main()
