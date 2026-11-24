# -*- coding: utf-8 -*-
"""diag_e46_liberstore.py -- Tester/QA 24/11/2026: por que a troca de lote ida e volta perde as assinaturas da Liberacao
(repro_e46_liberacao T03). Le o estado (loteAtivo, loteCarregado, regLoteCol, protecao da LiberStore), grava um
marcador em C/D de uma linha da Liberacao, troca o lote (mApp.UsarLote) e olha o bloco da LiberStore. COPIA; nada salvo.

Uso: python diag_e46_liberstore.py <produto> <copia.xlsm>    (DIAG_LOGIN=1: roda mSeguranca.ReprotectAll antes, como o login)
"""
import json
import os
import sys

TESTES = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'integracao_dbseac', 'testes')
sys.path.insert(0, os.path.abspath(TESTES))
from qa_final import QA, sha256_arquivo  # noqa: E402
from qa_fluxo_manual import lib_gravar  # noqa: E402


def t(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def main(produto, caminho):
    h0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    out = {}
    try:
        nm = lambda n: q.wb.Names(n).RefersToRange  # noqa: E731
        cad = [t(c.Value) for c in nm('regLoteCol')]
        stl = q.ws('LiberStore')
        out['inicio'] = {'loteAtivo': t(nm('loteAtivo').Value), 'loteCarregado': t(nm('loteCarregado').Value),
                         'LoteAtivoCore': t(q.run('mDados.LoteAtivoCore')),
                         'regLoteCol_primeiros': cad[:6], 'regLoteCol_endereco': nm('regLoteCol').Address,
                         'libView': nm('libView').Address,
                         'LiberStore_protegida': bool(stl.ProtectContents),
                         'LiberStore_UIOnly': bool(stl.ProtectionMode),
                         'LiberStore_usado': stl.UsedRange.Address}
        lote = out['inicio']['LoteAtivoCore']
        ib = cad.index(out['inicio']['loteCarregado']) + 1 if out['inicio']['loteCarregado'] in cad else 0
        out['bloco_do_carregado'] = ib
        outro = next(c for c in cad if c and c != lote)
        if os.environ.get('DIAG_LOGIN') == '1':
            q.run('mSeguranca.ReprotectAll', teto=300)      # o que o login faz (UserInterfaceOnly na sessao)
        lib_gravar(q, 0, 'QA', 'DIAG marcador linha 4')
        out['antes_ida'] = {'LiberStore_protegida': bool(stl.ProtectContents), 'LiberStore_UIOnly': bool(stl.ProtectionMode)}
        q.run('mApp.UsarLote', outro, teto=1800)
        rb0 = 2 + (ib - 1) * 200 if ib else 2
        blk = [[t(x) for x in r] for r in stl.Range(stl.Cells(rb0, 1), stl.Cells(rb0 + 3, 5)).Value]
        out['depois_ida'] = {'loteCarregado': t(nm('loteCarregado').Value), 'LoteAtivoCore': t(q.run('mDados.LoteAtivoCore')),
                             'bloco_linhas_1a4': blk}
        q.run('mApp.UsarLote', lote, teto=1800)
        lib = q.ws('Liberação')
        out['depois_volta'] = {'loteCarregado': t(nm('loteCarregado').Value),
                               'Liberacao_A4_F5': [[t(x) for x in r] for r in lib.Range('A4:F5').Value]}
    finally:
        q.fechar(salvar=False)
    out['sha_intacto'] = sha256_arquivo(caminho) == h0
    print(json.dumps(out, ensure_ascii=False, indent=1, default=str), flush=True)


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]))
