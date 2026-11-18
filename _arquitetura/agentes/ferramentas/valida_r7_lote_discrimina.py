# -*- coding: utf-8 -*-
"""valida_r7_lote_discrimina.py -- Tester/QA, ciclo 2026-11-16_reteste07, VALIDAR (OC-0083, complemento).

No valida_r7_adr097.py o lote digitado (B4) filtrou "2 de 2": todos os registros eram do mesmo lote, entao a prova nao
distinguia "filtra" de "nao filtra". Aqui, numa COPIA nova: troca o lote do 1o registro por outro lote da lista
(sintetico, so na copia), reconstroi as listas e DIGITA (FormulaLocal) cada lote, so numeros, em B4 (celula em texto):
  - a validacao aceita, o valor fica texto, e AplicarFiltrosRegistros mostra exatamente as linhas daquele lote (oraculo);
  - um lote fora da lista continua recusado.
Uso: python -u valida_r7_lote_discrimina.py <Bioquimica|Hematologia>
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
C_LOTE = 13
USO = (2, 3, 7, 11)          # mesmo criterio de "registro" do ContarRegistros (Data, Analito, Observacao, ID)


def main():
    os.makedirs(V.PASTA, exist_ok=True)
    src = os.path.join(V.C.REPO, f'QC_{PROD}.xlsm')
    dst = os.path.join(V.PASTA, f'{PROD}_lote_{time.strftime("%H%M%S")}.xlsm')
    L = V.Log(os.path.join(V.PASTA, f'valida_r7_lote_discrimina_{PROD}.json'))
    h0 = V.sha_full(src)
    shutil.copy2(src, dst)
    L['origem'] = {'sha256_antes': h0[:16], 'copia': os.path.basename(dst), 'sha_copia': V.sha_full(dst)[:16]}
    qa = qa_final.QA(PROD, dst)
    L['excel_pid'] = qa.ex.pid
    try:
        xl = qa.ex.xl
        ws = qa.ws('Registros')
        lo = ws.ListObjects('tblRegistros')
        ws.Unprotect(V.C.senha_planilha())
        n = lo.ListRows.Count
        v0 = lo.DataBodyRange.Value2
        usadas = [i for i in range(n) if any(V.t(v0[i][c - 1]) for c in USO)]
        col = lo.ListColumns(C_LOTE).DataBodyRange
        L['coluna_lote'] = {'tem_formula': bool(col.HasFormula) if col.HasFormula is not None else 'misto',
                            'lotes_antes': [V.t(v0[i][C_LOTE - 1]) for i in usadas]}
        qa.run('mCalibracao.LimparFiltrosRegistros', teto=120)
        qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
        cel = ws.Cells(4, 2)
        _, lista = V.itens_lista(cel)
        lote_b = V.t(v0[usadas[-1]][C_LOTE - 1])
        lote_a = next(x for x in lista if x.isdigit() and x != lote_b)
        if len(usadas) < 2:
            raise RuntimeError('menos de 2 registros na copia')
        xl.EnableEvents = False
        col.Cells(usadas[0] + 1, 1).Value = lote_a          # so na copia
        xl.EnableEvents = True
        qa.run('mCalibracao.AtualizarListasFiltro', teto=120)
        vals = lo.DataBodyRange.Value2
        _, lista2 = V.itens_lista(cel)
        L['lista_B4'] = lista2
        casos = {}
        for x in (lote_a, lote_b):
            r = {}
            xl.EnableEvents = True
            cel.FormulaLocal = x                                  # digitado, so numeros
            qa.ex.esperar()
            time.sleep(0.5)
            r['tipo_valor'] = type(cel.Value).__name__
            r['validacao_ok'] = bool(cel.Validation.Value)
            r['retorno'] = V.t(qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120))
            r['K4'] = V.t(ws.Range('K4').Text)
            orac = sum(1 for rr in vals if V.t(rr[C_LOTE - 1]) == x and any(V.t(rr[c - 1]) for c in USO))
            total = sum(1 for rr in vals if any(V.t(rr[c - 1]) for c in USO))
            vis = r['retorno'].split('|')
            r['oraculo'] = orac
            r['total'] = total
            r['ok'] = (r['tipo_valor'] == 'str' and r['validacao_ok'] and vis[0].isdigit() and int(vis[0]) == orac
                       and orac < total)
            xl.EnableEvents = False
            cel.ClearContents()
            xl.EnableEvents = True
            casos[x] = r
        qa.run('mCalibracao.AplicarFiltrosRegistros', teto=120)
        L['casos'] = casos
        L['fora_da_lista'] = {x: V.aceita(xl, cel, x, 'local') for x in ('99999', '1', lote_a + '0')}
        xl.EnableEvents = False
        cel.ClearContents()
        xl.EnableEvents = True
        L['ok'] = all(v['ok'] for v in casos.values()) and all(v is False for v in L['fora_da_lista'].values())
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
